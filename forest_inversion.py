# -*- coding: utf-8 -*-
"""
Created on Mon Oct 20 11:11:14 2025

@author: AAA
"""

import sys
import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Qt5Agg")   # 🔥 强制使用 Qt 后端（关键）
import matplotlib.pyplot as plt
from matplotlib import rcParams
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QLabel, QFileDialog,
    QLineEdit, QTextEdit, QVBoxLayout, QHBoxLayout, QWidget,
    QSpinBox, QDoubleSpinBox, QComboBox, QCheckBox
)
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from sklearn.model_selection import train_test_split, RandomizedSearchCV
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import r2_score, mean_squared_error
from xgboost import XGBRegressor
from osgeo import gdal

# 支持中文显示
rcParams['font.sans-serif'] = ['SimHei']
rcParams['axes.unicode_minus'] = False

class ForestInversionApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("森林参数反演模块1.0")
        self.resize(1400, 800)

        # 路径设置
        self.excel_path = QLineEdit()
        self.tif_folder = QLineEdit()
        self.reference_tif = QLineEdit()
        self.output_excel = QLineEdit()
        self.output_tif = QLineEdit()

        self.btn_excel = QPushButton("选择Excel")
        self.btn_tif_folder = QPushButton("选择TIF文件夹")
        self.btn_ref_tif = QPushButton("选择参考TIF")
        self.btn_out_excel = QPushButton("设置输出Excel")
        self.btn_out_tif = QPushButton("设置输出TIF")

        # 数据划分
        self.split_mode = QComboBox()
        self.split_mode.addItems(["固定数量", "按比例随机划分"])
        self.n_train = QSpinBox()
        self.n_train.setRange(100, 100000)
        self.n_train.setValue(10958)
        self.train_ratio = QDoubleSpinBox()
        self.train_ratio.setRange(0.1, 0.9)
        self.train_ratio.setSingleStep(0.05)
        self.train_ratio.setValue(0.7)

        # 参数优化
        self.use_hyperopt = QCheckBox("启用超参优化")

        # 自动选择变量
        self.select_vars = QCheckBox("启用变量自动选择")
        self.n_vars = QSpinBox()
        self.n_vars.setRange(0, 50)
        self.n_vars.setValue(0)  # 0表示全部变量

        # 日志窗口
        self.log = QTextEdit()
        self.log.setReadOnly(True)

        # 图像显示：一行两列布局
        self.figure, (self.ax1, self.ax2) = plt.subplots(1,2,figsize=(12,5))
        self.canvas = FigureCanvas(self.figure)

        self.btn_run = QPushButton("开始运行")

        # 布局
        layout = QVBoxLayout()

        # 文件选择布局
        for label, line, btn in [
            ("训练Excel:", self.excel_path, self.btn_excel),
            ("TIF文件夹:", self.tif_folder, self.btn_tif_folder),
            ("参考TIF:", self.reference_tif, self.btn_ref_tif),
            ("输出Excel:", self.output_excel, self.btn_out_excel),
            ("输出TIF:", self.output_tif, self.btn_out_tif)
        ]:
            row = QHBoxLayout()
            row.addWidget(QLabel(label))
            row.addWidget(line)
            row.addWidget(btn)
            layout.addLayout(row)

        # 数据划分布局
        row_split = QHBoxLayout()
        row_split.addWidget(QLabel("数据划分方式:"))
        row_split.addWidget(self.split_mode)
        row_split.addWidget(QLabel("训练数量:"))
        row_split.addWidget(self.n_train)
        row_split.addWidget(QLabel("训练比例:"))
        row_split.addWidget(self.train_ratio)
        layout.addLayout(row_split)

        # 参数优化布局
        row_opt = QHBoxLayout()
        row_opt.addWidget(self.use_hyperopt)
        row_opt.addWidget(self.select_vars)
        row_opt.addWidget(QLabel("使用前 N 个变量 (0=全部):"))
        row_opt.addWidget(self.n_vars)
        layout.addLayout(row_opt)

        layout.addWidget(self.btn_run)
        layout.addWidget(QLabel("运行日志:"))
        layout.addWidget(self.log)
        layout.addWidget(QLabel("预测结果和反演影像:"))
        layout.addWidget(self.canvas)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # 绑定事件
        self.btn_excel.clicked.connect(lambda: self.select_file(self.excel_path, "Excel文件 (*.xlsx)"))
        self.btn_tif_folder.clicked.connect(lambda: self.select_folder(self.tif_folder))
        self.btn_ref_tif.clicked.connect(lambda: self.select_file(self.reference_tif, "TIF文件 (*.tif)"))
        self.btn_out_excel.clicked.connect(lambda: self.save_file(self.output_excel, "Excel文件 (*.xlsx)"))
        self.btn_out_tif.clicked.connect(lambda: self.save_file(self.output_tif, "TIF文件 (*.tif)"))
        self.btn_run.clicked.connect(self.run_model)

    def log_msg(self, msg):
        self.log.append(msg)
        self.log.ensureCursorVisible()
        QApplication.processEvents()

    def select_file(self, line_edit, file_filter):
        path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", file_filter)
        if path:
            line_edit.setText(path)

    def save_file(self, line_edit, file_filter):
        path, _ = QFileDialog.getSaveFileName(self, "保存文件", "", file_filter)
        if path:
            line_edit.setText(path)

    def select_folder(self, line_edit):
        path = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if path:
            line_edit.setText(path)

    def read_tif(self, filename):
        dataset = gdal.Open(filename)
        im_width = dataset.RasterXSize
        im_height = dataset.RasterYSize
        im_geotrans = dataset.GetGeoTransform()
        im_proj = dataset.GetProjection()
        im_data = dataset.ReadAsArray(0, 0, im_width, im_height).astype(np.float32)
        del dataset
        return im_proj, im_geotrans, im_data, im_height, im_width

    def save_tif(self, output_path, data, im_proj, im_geotrans, im_height, im_width):
        driver = gdal.GetDriverByName("GTiff")
        out_tif = driver.Create(output_path, im_width, im_height, 1, gdal.GDT_Float32)
        out_tif.SetGeoTransform(im_geotrans)
        out_tif.SetProjection(im_proj)
        out_tif.GetRasterBand(1).WriteArray(data.reshape(im_height, im_width))
        out_tif.FlushCache()
        del out_tif

    def run_model(self):
        try:
            self.log.clear()
            self.log_msg("读取Excel数据...")
            data = pd.read_excel(self.excel_path.text())
            y = data.iloc[:, 0].values
            X = data.iloc[:, 1:].values
            feature_order = data.columns[1:].tolist()

            # 自动选择变量
            if self.select_vars.isChecked():
                n_var = self.n_vars.value()
                if n_var <= 0 or n_var > X.shape[1]:
                    n_var = X.shape[1]  # 使用全部
                temp_model = XGBRegressor(n_estimators=100, random_state=42)
                temp_model.fit(X, y)
                importances = temp_model.feature_importances_
                sorted_idx = np.argsort(importances)[::-1]
                selected_idx = sorted_idx[:n_var]
                X = X[:, selected_idx]
                feature_order = [feature_order[i] for i in selected_idx]
                self.log_msg(f"启用自动变量选择，使用前 {len(selected_idx)} 个重要性变量")
            else:
                self.log_msg(f"未启用自动变量选择，使用全部 {X.shape[1]} 个变量")

            scaler_X = StandardScaler()
            scaler_y = StandardScaler()
            X_scaled = scaler_X.fit_transform(X)
            y_scaled = scaler_y.fit_transform(y.reshape(-1, 1)).flatten()

            # 数据划分
            if self.split_mode.currentText() == "固定数量":
                n_train = min(self.n_train.value(), len(X))
                X_train, X_test = X_scaled[:n_train], X_scaled[n_train:]
                y_train, y_test = y_scaled[:n_train], y_scaled[n_train:]
                self.log_msg(f"划分方式: 固定数量 {n_train} / {len(X)-n_train}")
            else:
                ratio = self.train_ratio.value()
                X_train, X_test, y_train, y_test = train_test_split(
                    X_scaled, y_scaled, train_size=ratio, random_state=42
                )
                self.log_msg(f"划分方式: 按比例 {ratio*100:.1f}% 训练 / {(1-ratio)*100:.1f}% 测试")

            # XGB 模型
            if self.use_hyperopt.isChecked():
                self.log_msg("执行超参优化...")
                param_grid = {
                    "n_estimators": [200, 300, 400, 500, 600],
                    "max_depth": [3, 4, 5, 6, 8],
                    "learning_rate": [0.01, 0.03, 0.05, 0.1],
                    "subsample": [0.8, 1.0],
                    "colsample_bytree": [0.8, 1.0],
                }
                model = XGBRegressor(random_state=42)
                search = RandomizedSearchCV(model, param_grid, n_iter=20, cv=3, scoring="r2", n_jobs=-1, random_state=42)
                search.fit(X_train, y_train)
                final_model = search.best_estimator_
                self.log_msg(f"最佳参数: {search.best_params_}")
            else:
                final_model = XGBRegressor(
                    n_estimators=400,
                    learning_rate=0.05,
                    max_depth=6,
                    subsample=0.8,
                    colsample_bytree=0.8,
                    random_state=42
                )
                final_model.fit(X_train, y_train)
                self.log_msg("使用默认参数训练")

            # 测试评估
            y_pred_scaled = final_model.predict(X_test)
            y_pred = scaler_y.inverse_transform(y_pred_scaled.reshape(-1, 1)).flatten()
            actual = scaler_y.inverse_transform(y_test.reshape(-1, 1)).flatten()
            r2 = r2_score(actual, y_pred)
            rmse = np.sqrt(mean_squared_error(actual, y_pred))
            self.log_msg(f"Test R²={r2:.3f}, RMSE={rmse:.3f}")

            # 保存预测结果 Excel
            predictions_df = pd.DataFrame({"Actual": actual, "Predicted": y_pred})
            predictions_df.to_excel(self.output_excel.text(), index=False)
            self.log_msg(f"预测结果已保存: {self.output_excel.text()}")

            # 绘制核密度散点图
            self.ax1.clear()
            hb = self.ax1.hexbin(actual, y_pred, gridsize=50, cmap='viridis', mincnt=1)
            self.ax1.plot([min(actual), max(actual)], [min(actual), max(actual)], 'r--')
            self.ax1.set_xlabel("实际值")
            self.ax1.set_ylabel("预测值")
            self.ax1.set_title(f"核密度散点图 $R^2={r2:.3f}$,RMSE={rmse:.3f}")
            cb = self.figure.colorbar(hb, ax=self.ax1)
            cb.set_label("样本数密度")

            # 影像预测
            self.log_msg("对影像进行预测...")
            ref_proj, ref_geotrans, ref_data, im_height, im_width = self.read_tif(self.reference_tif.text())
            feature_matrix = []
            for feat in feature_order:
                path = os.path.join(self.tif_folder.text(), feat + ".tif")
                if not os.path.exists(path):
                    raise FileNotFoundError(f"缺少特征文件 {path}")
                ds = gdal.Open(path)
                arr = gdal.Warp('', ds, format="MEM", dstSRS=ref_proj, xRes=ref_geotrans[1], yRes=-ref_geotrans[5],
                                outputBounds=(ref_geotrans[0], ref_geotrans[3] + im_height*ref_geotrans[5],
                                              ref_geotrans[0] + im_width*ref_geotrans[1], ref_geotrans[3]),
                                width=im_width, height=im_height).ReadAsArray().astype(np.float32)
                feature_matrix.append(arr.flatten())
            feature_matrix = np.array(feature_matrix).T
            feature_matrix = np.nan_to_num(feature_matrix, nan=0.0, posinf=1e6, neginf=-1e6)

            feature_matrix_scaled = scaler_X.transform(feature_matrix)
            tif_pred_scaled = final_model.predict(feature_matrix_scaled)
            tif_pred = scaler_y.inverse_transform(tif_pred_scaled.reshape(-1, 1)).flatten()
            # 掩膜处理：参考影像掩膜阈值
            mask = (ref_data < 0) | (ref_data > 26695) | np.isnan(ref_data) | (ref_data < 7916)
            tif_pred[mask.flatten()] = np.nan

            self.save_tif(self.output_tif.text(), tif_pred, ref_proj, ref_geotrans, im_height, im_width)
            self.log_msg(f"预测影像已保存: {self.output_tif.text()}")

            # 显示反演结果图
            self.ax2.clear()
            img = tif_pred.reshape(im_height, im_width)
            im = self.ax2.imshow(img, cmap='viridis')
            self.ax2.set_title("反演结果图")
            cb2 = self.figure.colorbar(im, ax=self.ax2)
            cb2.set_label("预测值")

            self.canvas.draw()
            self.log_msg("处理完成。")

        except Exception as e:
            self.log_msg(f"错误: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = ForestInversionApp()
    win.show()
    sys.exit(app.exec_())
