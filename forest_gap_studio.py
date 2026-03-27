# -*- coding: utf-8 -*-
"""
ForestMetrics Studio v1.0 - 林窗检测集成模块 (完整界面版)
包含：环境修复、Attention U-Net模型、无监督打标、数据集制作、训练、全图推理、精度评价
"""

import os
import sys
import glob
import random
import warnings
import numpy as np
import pandas as pd
import geopandas as gpd
import rasterio
import cv2
import matplotlib
import matplotlib.pyplot as plt
from tqdm import tqdm

# ==============================================================================
# 0. 环境修复 (解决 WinError 127 / fbgemm.dll 依赖冲突)
# ==============================================================================
if os.name == 'nt':
    os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'
    env_base = os.path.dirname(sys.executable)
    dll_paths = [
        os.path.join(env_base, 'Library', 'bin'),
        os.path.join(env_base, 'Lib', 'site-packages', 'torch', 'lib'),
    ]
    for path in dll_paths:
        if os.path.exists(path):
            os.environ['PATH'] = path + os.pathsep + os.environ.get('PATH', '')
            if hasattr(os, 'add_dll_directory'):
                try:
                    os.add_dll_directory(path)
                except:
                    pass

import torch
import torch.nn as nn
from torch.utils.data import Dataset, DataLoader
from torch.optim import Adam
from torch.utils.tensorboard import SummaryWriter
from skimage.filters import threshold_otsu
from skimage.measure import label, regionprops
from skimage.morphology import remove_small_objects, remove_small_holes, binary_opening, binary_closing, disk
from shapely.errors import ShapelyDeprecationWarning
from shapely.ops import unary_union

from PyQt5.QtWidgets import (QApplication, QMainWindow, QDialog, QVBoxLayout, QTabWidget,
                             QWidget, QFormLayout, QLineEdit, QPushButton,
                             QHBoxLayout, QFileDialog, QSpinBox, QDoubleSpinBox,
                             QMessageBox, QLabel)
from PyQt5.QtCore import Qt

matplotlib.use("Qt5Agg")
warnings.filterwarnings("ignore", category=ShapelyDeprecationWarning)
warnings.filterwarnings("ignore", category=FutureWarning)


# ==============================================================================
# 1. 模型架构 (Attention U-Net)
# ==============================================================================
class DoubleConv(nn.Module):
    def __init__(self, in_ch, out_ch):
        super().__init__()
        self.conv = nn.Sequential(
            nn.Conv2d(in_ch, out_ch, kernel_size=3, padding=1),
            nn.BatchNorm2d(out_ch), nn.ReLU(inplace=True),
            nn.Conv2d(out_ch, out_ch, kernel_size=3, padding=1),
            nn.BatchNorm2d(out_ch), nn.ReLU(inplace=True)
        )

    def forward(self, x): return self.conv(x)


class AttentionBlock(nn.Module):
    def __init__(self, F_g, F_l, F_int):
        super().__init__()
        self.W_g = nn.Sequential(nn.Conv2d(F_g, F_int, kernel_size=1), nn.BatchNorm2d(F_int))
        self.W_x = nn.Sequential(nn.Conv2d(F_l, F_int, kernel_size=1), nn.BatchNorm2d(F_int))
        self.psi = nn.Sequential(nn.Conv2d(F_int, 1, kernel_size=1), nn.BatchNorm2d(1), nn.Sigmoid())
        self.relu = nn.ReLU(inplace=True)

    def forward(self, g, x):
        g1, x1 = self.W_g(g), self.W_x(x)
        return x * self.psi(self.relu(g1 + x1))


class AttentionUNet(nn.Module):
    def __init__(self, in_ch=3, out_ch=1):
        super().__init__()
        self.dconv_down1, self.dconv_down2 = DoubleConv(in_ch, 64), DoubleConv(64, 128)
        self.dconv_down3, self.dconv_down4 = DoubleConv(128, 256), DoubleConv(256, 512)
        self.maxpool = nn.MaxPool2d(2)
        self.upsample4 = nn.ConvTranspose2d(512, 512, 2, 2)
        self.att3 = AttentionBlock(512, 256, 256)
        self.dconv_up3 = DoubleConv(768, 256)
        self.upsample3 = nn.ConvTranspose2d(256, 256, 2, 2)
        self.att2 = AttentionBlock(256, 128, 128)
        self.dconv_up2 = DoubleConv(384, 128)
        self.upsample2 = nn.ConvTranspose2d(128, 128, 2, 2)
        self.att1 = AttentionBlock(128, 64, 64)
        self.dconv_up1 = DoubleConv(192, 64)
        self.conv_last = nn.Conv2d(64, out_ch, 1)

    def forward(self, x):
        c1 = self.dconv_down1(x);
        c2 = self.dconv_down2(self.maxpool(c1))
        c3 = self.dconv_down3(self.maxpool(c2));
        c4 = self.dconv_down4(self.maxpool(c3))
        x4 = self.upsample4(c4)
        x3 = self.dconv_up3(torch.cat([x4, self.att3(x4, c3)], dim=1))
        x2_up = self.upsample3(x3)
        x2 = self.dconv_up2(torch.cat([x2_up, self.att2(x2_up, c2)], dim=1))
        x1_up = self.upsample2(x2)
        x1 = self.dconv_up1(torch.cat([x1_up, self.att1(x1_up, c1)], dim=1))
        return torch.sigmoid(self.conv_last(x1))


# ==============================================================================
# 2. 算法核心模块 (整合自 gap_label, ild_unet_dataset, train, infer, metrics)
# ==============================================================================
def robust_normalize(x, mask):
    vals = x[mask]
    if len(vals) == 0: return np.zeros_like(x)
    p2, p98 = np.percentile(vals, [2, 98])
    return np.clip((x - p2) / (p98 - p2 + 1e-6), 0, 1)


def run_gap_extraction(input_tif, output_tif, output_vis, gt_tif=None, min_area=5.0, dark=0.95, open_r=2, close_r=4):
    with rasterio.open(input_tif) as src:
        r, g, b = src.read(1), src.read(2), src.read(3)
        profile, transform = src.profile, src.transform
    valid = (r > 0) & (g > 0) & (b > 0)
    rn, gn, bn = robust_normalize(r, valid), robust_normalize(g, valid), robust_normalize(b, valid)
    gray = 0.299 * rn + 0.587 * gn + 0.114 * bn
    dark_score = 0.5 * gray + 0.5 * (
                cv2.cvtColor((np.dstack([rn, gn, bn] * 255)).astype(np.uint8), cv2.COLOR_RGB2HSV)[:, :, 2] / 255.0)
    th = threshold_otsu(dark_score[valid]) * dark
    gap_init = (dark_score < th) & valid
    min_pix = max(1, int(round(min_area / (abs(transform.a) * abs(transform.e)))))
    mask = remove_small_objects(binary_closing(binary_opening(gap_init, disk(open_r)), disk(close_r)), min_pix)

    profile.update(count=1, dtype=rasterio.uint8, nodata=0)
    os.makedirs(os.path.dirname(output_tif), exist_ok=True)
    with rasterio.open(output_tif, "w", **profile) as dst:
        dst.write(mask.astype(np.uint8), 1)

    vis = (np.dstack([rn, gn, bn]) * 255).astype(np.uint8)
    if gt_tif and os.path.exists(gt_tif):
        with rasterio.open(gt_tif) as g_src:
            gt = g_src.read(1) > 0
        vis[(mask > 0) & gt] = [0, 255, 0]  # TP
        vis[(mask > 0) & ~gt] = [255, 0, 0]  # FP
        vis[~(mask > 0) & gt] = [0, 0, 255]  # FN
    else:
        vis[mask > 0] = [255, 0, 0]
    cv2.imwrite(output_vis, cv2.cvtColor(vis, cv2.COLOR_RGB2BGR))
    return mask.sum() / valid.sum()


def generate_patches(rgb_tif, mask_tif, out_dir, patch_size=256, stride=256):
    with rasterio.open(rgb_tif) as s1, rasterio.open(mask_tif) as s2:
        rgb, mask = np.transpose(s1.read([1, 2, 3]), (1, 2, 0)), s2.read(1)
    os.makedirs(os.path.join(out_dir, "images"), exist_ok=True)
    os.makedirs(os.path.join(out_dir, "masks"), exist_ok=True)
    h, w, _ = rgb.shape;
    count = 0
    for i in range(0, h - patch_size + 1, stride):
        for j in range(0, w - patch_size + 1, stride):
            if np.mean(mask[i:i + patch_size, j:j + patch_size]) < 0.001: continue
            name = f"{count:05d}.png"
            cv2.imwrite(os.path.join(out_dir, "images", name),
                        cv2.cvtColor(rgb[i:i + patch_size, j:j + patch_size], cv2.COLOR_RGB2BGR))
            cv2.imwrite(os.path.join(out_dir, "masks", name), mask[i:i + patch_size, j:j + patch_size] * 255)
            count += 1
    return count


class GapDataset(Dataset):
    def __init__(self, imgs, masks): self.imgs, self.masks = imgs, masks

    def __len__(self): return len(self.imgs)

    def __getitem__(self, idx):
        img = np.transpose(cv2.cvtColor(cv2.imread(self.imgs[idx]), cv2.COLOR_BGR2RGB).astype(np.float32) / 255.0,
                           (2, 0, 1))
        mask = np.expand_dims((cv2.imread(self.masks[idx], 0) > 127).astype(np.float32), 0)
        return torch.tensor(img), torch.tensor(mask)


def main_train(img_dir, mask_dir, output_dir, epochs=60, batch_size=8, lr=2e-4):
    device = "cuda" if torch.cuda.is_available() else "cpu"
    imgs, masks = sorted(glob.glob(os.path.join(img_dir, "*.png"))), sorted(glob.glob(os.path.join(mask_dir, "*.png")))
    loader = DataLoader(GapDataset(imgs, masks), batch_size=batch_size, shuffle=True)
    model = AttentionUNet().to(device);
    opt = Adam(model.parameters(), lr=lr);
    crit = nn.BCELoss()
    for epoch in range(epochs):
        model.train();
        pbar = tqdm(loader, desc=f"Epoch {epoch + 1}/{epochs}")
        for x, y in pbar:
            x, y = x.to(device), y.to(device)
            opt.zero_grad();
            out = model(x);
            loss = crit(out, y)
            loss.backward();
            opt.step();
            pbar.set_postfix(loss=loss.item())
    torch.save(model.state_dict(), os.path.join(output_dir, "best_model.pth"))


def predict_large_tif(model_path, tif_path, out_tif, gt_path=None, tile_size=512, stride=384, threshold=0.35):
    device = "cuda" if torch.cuda.is_available() else "cpu"
    model = AttentionUNet().to(device);
    model.load_state_dict(torch.load(model_path, map_location=device), strict=False);
    model.eval()
    with rasterio.open(tif_path) as src:
        img = np.transpose(src.read()[:3], (1, 2, 0)).astype(np.float32) / 255.0
        profile = src.profile.copy()
    h, w, _ = img.shape;
    res, cnt = np.zeros((h, w)), np.zeros((h, w))
    for y in tqdm(range(0, h - tile_size + 1, stride), desc="Infer"):
        for x in range(0, w - tile_size + 1, stride):
            p = torch.from_numpy(np.transpose(img[y:y + tile_size, x:x + tile_size], (2, 0, 1))).unsqueeze(0).to(device)
            with torch.no_grad(): pred = model(p)[0, 0].cpu().numpy()
            res[y:y + tile_size, x:x + tile_size] += pred;
            cnt[y:y + tile_size, x:x + tile_size] += 1
    bin_res = (res / np.clip(cnt, 1, None) > threshold).astype(np.uint8)
    profile.update(count=1, dtype=rasterio.uint8)
    with rasterio.open(out_tif, "w", **profile) as dst:
        dst.write(bin_res, 1)


def calculate_metrics(gt_shp, pred_shp, out_csv):
    g1, g2 = gpd.read_file(gt_shp), gpd.read_file(pred_shp)
    if g1.crs != g2.crs: g2 = g2.to_crs(g1.crs)
    u1, u2 = unary_union(g1.geometry), unary_union(g2.geometry)
    tp = u1.intersection(u2).area
    iou = tp / (u1.area + u2.area - tp + 1e-8)
    pd.DataFrame({"Metric": ["IoU", "GT_Area", "Pred_Area"], "Value": [iou, u1.area, u2.area]}).to_csv(out_csv,
                                                                                                       index=False)


# ==============================================================================
# 3. 界面逻辑 (还原图一设计)
# ==============================================================================
class GapDetectionModule(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("林窗检测分析模块 (Gap Detection) - ForestMetrics Studio v1.0")
        self.resize(1100, 760)
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout(self)
        self.tabs = QTabWidget()
        self.tabs.addTab(self.create_label_tab(), "1. 标签制作 (Labeling)")
        self.tabs.addTab(self.create_dataset_tab(), "2. 数据集构建 (Dataset)")
        self.tabs.addTab(self.create_train_tab(), "3. 模型训练 (Training)")
        self.tabs.addTab(self.create_infer_tab(), "4. 推理与可视化 (Inference)")
        self.tabs.addTab(self.create_metric_tab(), "5. 精度评价 (Metrics)")
        layout.addWidget(self.tabs)

        btn_layout = QHBoxLayout();
        btn_layout.addStretch(1)
        close_btn = QPushButton("关闭 (Close)");
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(close_btn);
        layout.addLayout(btn_layout)

    def create_file_picker(self, label_text, is_dir=False, is_save=False):
        layout = QHBoxLayout();
        label = QLabel(label_text);
        label.setFixedWidth(170)
        line_edit = QLineEdit();
        btn = QPushButton("浏览...")

        def pick_path():
            if is_dir:
                path = QFileDialog.getExistingDirectory(self, "选择文件夹")
            elif is_save:
                path, _ = QFileDialog.getSaveFileName(self, "保存文件")
            else:
                path, _ = QFileDialog.getOpenFileName(self, "选择文件")
            if path: line_edit.setText(path)

        btn.clicked.connect(pick_path);
        layout.addWidget(label);
        layout.addWidget(line_edit);
        layout.addWidget(btn)
        return layout, line_edit

    # --- 各个标签页布局还原 (核心修改) ---
    def create_label_tab(self):
        widget = QWidget();
        layout = QVBoxLayout(widget)
        self.l1_in_lay, self.l1_in = self.create_file_picker("输入 DOM (.tif):")
        self.l1_out_lay, self.l1_out = self.create_file_picker("输出结果 (.tif):", is_save=True)
        self.l1_gt_lay, self.l1_gt = self.create_file_picker("参考 GT (.tif, 可选):")
        layout.addLayout(self.l1_in_lay);
        layout.addLayout(self.l1_out_lay);
        layout.addLayout(self.l1_gt_lay)

        form = QFormLayout()
        self.l1_area = QDoubleSpinBox();
        self.l1_area.setValue(5.0);
        self.l1_area.setSuffix(" m²")
        self.l1_dark = QDoubleSpinBox();
        self.l1_dark.setValue(0.95);
        self.l1_dark.setSingleStep(0.01)
        self.l1_open = QSpinBox();
        self.l1_open.setValue(2)
        self.l1_close = QSpinBox();
        self.l1_close.setValue(4)
        form.addRow("最小林窗面积:", self.l1_area);
        form.addRow("Otsu暗度阈值系数:", self.l1_dark)
        form.addRow("开运算半径:", self.l1_open);
        form.addRow("闭运算半径:", self.l1_close)
        layout.addLayout(form)

        btn = QPushButton("执行无监督标签提取");
        btn.clicked.connect(self.do_label);
        layout.addWidget(btn);
        layout.addStretch()
        return widget

    def create_dataset_tab(self):
        widget = QWidget();
        layout = QVBoxLayout(widget)
        self.d_rgb_lay, self.d_rgb = self.create_file_picker("RGB 影像 (.tif):")
        self.d_mask_lay, self.d_mask = self.create_file_picker("掩膜 Mask (.tif):")
        self.d_out_lay, self.d_out = self.create_file_picker("输出根目录:", is_dir=True)
        layout.addLayout(self.d_rgb_lay);
        layout.addLayout(self.d_mask_lay);
        layout.addLayout(self.d_out_lay)

        form = QFormLayout()
        self.d_patch = QSpinBox();
        self.d_patch.setMaximum(2048);
        self.d_patch.setValue(256)
        self.d_stride = QSpinBox();
        self.d_stride.setMaximum(2048);
        self.d_stride.setValue(256)
        form.addRow("切片大小 (Patch Size):", self.d_patch);
        form.addRow("滑动步长 (Stride):", self.d_stride)
        layout.addLayout(form)

        btn = QPushButton("生成训练切片");
        btn.clicked.connect(self.do_dataset);
        layout.addWidget(btn);
        layout.addStretch()
        return widget

    def create_train_tab(self):
        widget = QWidget();
        layout = QVBoxLayout(widget)
        self.t_img_lay, self.t_img = self.create_file_picker("训练图像目录:", is_dir=True)
        self.t_mask_lay, self.t_mask = self.create_file_picker("训练掩膜目录:", is_dir=True)
        self.t_out_lay, self.t_out = self.create_file_picker("模型输出目录:", is_dir=True)
        layout.addLayout(self.t_img_lay);
        layout.addLayout(self.t_mask_lay);
        layout.addLayout(self.t_out_lay)

        form = QFormLayout()
        self.t_ep = QSpinBox();
        self.t_ep.setMaximum(1000);
        self.t_ep.setValue(60)
        self.t_bs = QSpinBox();
        self.t_bs.setValue(8)
        self.t_lr = QDoubleSpinBox();
        self.t_lr.setDecimals(5);
        self.t_lr.setValue(0.0002);
        self.t_lr.setSingleStep(0.0001)
        form.addRow("训练轮数 (Epochs):", self.t_ep);
        form.addRow("批次大小 (Batch Size):", self.t_bs);
        form.addRow("学习率 (LR):", self.t_lr)
        layout.addLayout(form)

        btn = QPushButton("启动模型训练");
        btn.clicked.connect(self.do_train);
        layout.addWidget(btn);
        layout.addStretch()
        return widget

    def create_infer_tab(self):
        widget = QWidget();
        layout = QVBoxLayout(widget)
        self.i_mod_lay, self.i_mod = self.create_file_picker("预训练权重 (.pth):")
        self.i_tif_lay, self.i_tif = self.create_file_picker("待预测影像 (.tif):")
        self.i_gt_lay, self.i_gt = self.create_file_picker("参考真值 (可选):")
        self.i_out_lay, self.i_out = self.create_file_picker("预测输出 (.tif):", is_save=True)
        layout.addLayout(self.i_mod_lay);
        layout.addLayout(self.i_tif_lay);
        layout.addLayout(self.i_gt_lay);
        layout.addLayout(self.i_out_lay)

        form = QFormLayout()
        self.i_tile = QSpinBox();
        self.i_tile.setMaximum(2048);
        self.i_tile.setValue(512)
        self.i_stride = QSpinBox();
        self.i_stride.setMaximum(2048);
        self.i_stride.setValue(384)
        self.i_th = QDoubleSpinBox();
        self.i_th.setValue(0.35);
        self.i_th.setSingleStep(0.05)
        form.addRow("推理切片大小:", self.i_tile);
        form.addRow("步长:", self.i_stride);
        form.addRow("阈值:", self.i_th)
        layout.addLayout(form)

        btn = QPushButton("执行全图推理与可视化");
        btn.clicked.connect(self.do_infer);
        layout.addWidget(btn);
        layout.addStretch()
        return widget

    def create_metric_tab(self):
        widget = QWidget();
        layout = QVBoxLayout(widget)
        self.m_gt_lay, self.m_gt = self.create_file_picker("真实值 (.tif):")
        self.m_pd_lay, self.m_pd = self.create_file_picker("预测值 (.tif):")
        self.m_out_lay, self.m_out = self.create_file_picker("输出 CSV 路径:", is_save=True)
        layout.addLayout(self.m_gt_lay);
        layout.addLayout(self.m_pd_lay);
        layout.addLayout(self.m_out_lay)
        btn = QPushButton("计算全局匹配指标");
        btn.clicked.connect(self.do_metric);
        layout.addWidget(btn);
        layout.addStretch()
        return widget

    # --- 执行逻辑 ---
    def do_label(self):
        run_gap_extraction(self.l1_in.text(), self.l1_out.text(), self.l1_out.text().replace(".tif", ".png"),
                           self.l1_gt.text() or None, self.l1_area.value(), self.l1_dark.value(), self.l1_open.value(),
                           self.l1_close.value())
        QMessageBox.information(self, "完成", "标签制作成功！")

    def do_dataset(self):
        count = generate_patches(self.d_img.text(), self.d_mask.text(), self.d_out.text(), self.d_patch.value(),
                                 self.d_stride.value())
        QMessageBox.information(self, "完成", f"已生成 {count} 张有效切片")

    def do_train(self):
        main_train(self.t_img.text(), self.t_mask.text(), self.t_out.text(), self.t_ep.value(), self.t_bs.value(),
                   self.t_lr.value())
        QMessageBox.information(self, "完成", "训练完成，权重已保存。")

    def do_infer(self):
        predict_large_tif(self.i_mod.text(), self.i_tif.text(), self.i_out.text(), self.i_gt.text() or None,
                          self.i_tile.value(), self.i_stride.value(), self.i_th.value())
        QMessageBox.information(self, "完成", "大图推理分析全部完成！")

    def do_metric(self):
        calculate_metrics(self.m_gt.text(), self.m_pd.text(), self.m_out.text())
        QMessageBox.information(self, "完成", "IoU 评价指标计算完成！")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = GapDetectionModule()
    window.show()
    sys.exit(app.exec_())