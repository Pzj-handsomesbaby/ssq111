# -*- coding: utf-8 -*-
"""
激光雷达数据处理模块（带点云去噪、地面点滤波、DEM/DSM/CHM生成及可视化预览）
"""

from PyQt5.QtWidgets import QWidget, QLabel, QLineEdit, QPushButton, QTextEdit, QVBoxLayout, QHBoxLayout, QFileDialog, QSpinBox, QDoubleSpinBox, QCheckBox, QSizePolicy
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
import os
import numpy as np
import laspy
from scipy.interpolate import griddata
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt

class LidarModule(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("激光雷达数据处理模块")
        self.resize(800, 600)

        layout = QVBoxLayout()

        # --- 输入文件 ---
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("输入LAS/LAZ文件夹:"))
        self.input_folder = QLineEdit()
        row1.addWidget(self.input_folder)
        btn_in = QPushButton("浏览")
        row1.addWidget(btn_in)
        btn_in.clicked.connect(self.browse_input)
        layout.addLayout(row1)

        # --- 输出路径 ---
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("输出DEM路径:"))
        self.output_dem = QLineEdit()
        row2.addWidget(self.output_dem)
        btn_out = QPushButton("设置")
        row2.addWidget(btn_out)
        btn_out.clicked.connect(self.set_output_dem)
        layout.addLayout(row2)

        row3 = QHBoxLayout()
        row3.addWidget(QLabel("输出DSM路径:"))
        self.output_dsm = QLineEdit()
        row3.addWidget(self.output_dsm)
        btn_dsm = QPushButton("设置")
        row3.addWidget(btn_dsm)
        btn_dsm.clicked.connect(self.set_output_dsm)
        layout.addLayout(row3)

        row4 = QHBoxLayout()
        row4.addWidget(QLabel("输出CHM路径:"))
        self.output_chm = QLineEdit()
        row4.addWidget(self.output_chm)
        btn_chm = QPushButton("设置")
        row4.addWidget(btn_chm)
        btn_chm.clicked.connect(self.set_output_chm)
        layout.addLayout(row4)

        # --- 参数设置 ---
        param_layout = QHBoxLayout()
        param_layout.addWidget(QLabel("网格分辨率(m):"))
        self.grid_res = QDoubleSpinBox()
        self.grid_res.setRange(0.1, 10.0)
        self.grid_res.setSingleStep(0.5)
        self.grid_res.setValue(1.0)
        param_layout.addWidget(self.grid_res)

        param_layout.addWidget(QLabel("去噪上下百分位(%):"))
        self.percentile = QSpinBox()
        self.percentile.setRange(0, 50)
        self.percentile.setValue(1)
        param_layout.addWidget(self.percentile)

        self.use_class_filter = QCheckBox("使用分类值过滤地面点(分类=2)")
        self.use_class_filter.setChecked(True)
        param_layout.addWidget(self.use_class_filter)

        layout.addLayout(param_layout)

        # --- 日志 ---
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(QLabel("运行日志:"))
        layout.addWidget(self.log)

        # --- 运行按钮 ---
        self.run_btn = QPushButton("开始生成DEM/DSM/CHM")
        layout.addWidget(self.run_btn)
        self.run_btn.clicked.connect(self.run_processing)

        # --- 可视化区域 ---
        self.figure, (self.ax_dem, self.ax_dsm, self.ax_chm) = plt.subplots(1, 3, figsize=(9,3))
        self.canvas = FigureCanvas(self.figure)
        self.canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        layout.addWidget(self.canvas)

        self.setLayout(layout)

    # ---------------- 界面操作 ----------------
    def browse_input(self):
        path = QFileDialog.getExistingDirectory(self, "选择LAS/LAZ文件夹")
        if path:
            self.input_folder.setText(path)

    def set_output_dem(self):
        path, _ = QFileDialog.getSaveFileName(self, "保存DEM文件", "", "TIF文件 (*.tif)")
        if path:
            self.output_dem.setText(path)

    def set_output_dsm(self):
        path, _ = QFileDialog.getSaveFileName(self, "保存DSM文件", "", "TIF文件 (*.tif)")
        if path:
            self.output_dsm.setText(path)

    def set_output_chm(self):
        path, _ = QFileDialog.getSaveFileName(self, "保存CHM文件", "", "TIF文件 (*.tif)")
        if path:
            self.output_chm.setText(path)

    def log_msg(self, msg):
        self.log.append(msg)
        self.log.ensureCursorVisible()

    # ---------------- 核心处理 ----------------
    def run_processing(self):
        self.log.clear()
        self.log_msg("开始激光雷达数据处理...")

        folder = self.input_folder.text()
        if not os.path.exists(folder):
            self.log_msg("输入文件夹不存在")
            return

        all_files = [f for f in os.listdir(folder) if f.lower().endswith(('.las','.laz'))]
        if len(all_files) == 0:
            self.log_msg("未找到LAS/LAZ文件")
            return

        all_points = []

        self.log_msg("读取点云数据...")
        for f in all_files:
            file_path = os.path.join(folder, f)
            self.log_msg(f"读取 {f} ...")
            with laspy.open(file_path) as lasf:
                las = lasf.read()
                points = np.vstack((las.x, las.y, las.z, las.classification)).T
                all_points.append(points)
        points = np.vstack(all_points)
        self.log_msg(f"总点数: {points.shape[0]}")

        # ---------------- 去噪 ----------------
        p = self.percentile.value()
        z = points[:,2]
        mask = (z > np.percentile(z,p)) & (z < np.percentile(z,100-p))
        points = points[mask]
        self.log_msg(f"去噪后点数: {points.shape[0]}")

        # ---------------- 地面点滤波 ----------------
        if self.use_class_filter.isChecked():
            ground_points = points[points[:,3]==2]
            if ground_points.shape[0]==0:
                self.log_msg("警告：未找到地面点，使用所有点生成DEM")
                ground_points = points
        else:
            ground_points = points

        # ---------------- 网格化 ----------------
        res = self.grid_res.value()
        x_min, x_max = np.min(points[:,0]), np.max(points[:,0])
        y_min, y_max = np.min(points[:,1]), np.max(points[:,1])
        xi = np.arange(x_min, x_max, res)
        yi = np.arange(y_min, y_max, res)
        XI, YI = np.meshgrid(xi, yi)

        # DEM
        self.log_msg("生成 DEM ...")
        DEM = griddata((ground_points[:,0], ground_points[:,1]), ground_points[:,2], (XI, YI), method='linear')
        DEM[np.isnan(DEM)] = 0

        # DSM
        self.log_msg("生成 DSM ...")
        DSM = griddata((points[:,0], points[:,1]), points[:,2], (XI, YI), method='max')
        DSM[np.isnan(DSM)] = 0

        # CHM
        self.log_msg("生成 CHM ...")
        CHM = DSM - DEM

        # ---------------- 保存 ----------------
        try:
            from osgeo import gdal
            def save_tif(filename, data, xi, yi):
                driver = gdal.GetDriverByName("GTiff")
                nx, ny = data.shape[1], data.shape[0]
                dataset = driver.Create(filename, nx, ny, 1, gdal.GDT_Float32)
                dataset.SetGeoTransform([xi[0], res, 0, yi[-1], 0, -res])
                dataset.GetRasterBand(1).WriteArray(data)
                dataset.FlushCache()
                del dataset

            if self.output_dem.text(): save_tif(self.output_dem.text(), DEM, xi, yi)
            if self.output_dsm.text(): save_tif(self.output_dsm.text(), DSM, xi, yi)
            if self.output_chm.text(): save_tif(self.output_chm.text(), CHM, xi, yi)
            self.log_msg("DEM/DSM/CHM 已保存")
        except Exception as e:
            self.log_msg(f"TIF 保存失败: {e}")

        # ---------------- 可视化 ----------------
        self.ax_dem.clear()
        self.ax_dem.imshow(DEM, cmap='terrain')
        self.ax_dem.set_title("DEM")
        self.ax_dem.axis('off')

        self.ax_dsm.clear()
        self.ax_dsm.imshow(DSM, cmap='terrain')
        self.ax_dsm.set_title("DSM")
        self.ax_dsm.axis('off')

        self.ax_chm.clear()
        self.ax_chm.imshow(CHM, cmap='Greens')
        self.ax_chm.set_title("CHM")
        self.ax_chm.axis('off')

        self.canvas.draw()
        self.log_msg("处理完成，可视化显示。")
