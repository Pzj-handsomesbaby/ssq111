# -*- coding: utf-8 -*-
"""林分质量评价（QI）模块。

该文件从 `zhao/eval_menu.py` 彻底迁移到 `modules` 目录，
主程序可直接导入 `StandQualityWindow`。
"""

import sys

import geopandas as gpd
import numpy as np
import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from sklearn.decomposition import PCA
try:
    from modules.ui_style import apply_qt_app_style, apply_qt_window_baseline
except ModuleNotFoundError:
    from ui_style import apply_qt_app_style, apply_qt_window_baseline


class StandQualityWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("林分质量评价")
        apply_qt_window_baseline(self, size=(1200, 780))
        self.df = None
        self.gdf = None
        self._last_selected_cols = []
        self._build_ui()

    def _build_ui(self):
        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)

        left_panel = QVBoxLayout()

        io_group = QGroupBox("数据源管理")
        io_layout = QVBoxLayout()
        self.btn_load = QPushButton("导入数据 (CSV/XLSX/SHP)")
        self.btn_load.setMinimumHeight(40)
        self.btn_load.clicked.connect(self.load_data)
        io_layout.addWidget(self.btn_load)
        self.lbl_status = QLabel("状态: 等待数据...")
        self.lbl_status.setStyleSheet("color: blue;")
        io_layout.addWidget(self.lbl_status)
        io_group.setLayout(io_layout)
        left_panel.addWidget(io_group)

        var_group = QGroupBox("选择评价指标变量")
        var_layout = QVBoxLayout()
        self.list_vars = QListWidget()
        var_layout.addWidget(self.list_vars)
        var_group.setLayout(var_layout)
        left_panel.addWidget(var_group, stretch=3)

        algo_group = QGroupBox("计算设置")
        algo_layout = QVBoxLayout()
        self.combo_method = QComboBox()
        self.combo_method.addItems(["熵权法 (Entropy Weight)", "等权法 (Equal Weight)", "主成分分析 (PCA)"])
        algo_layout.addWidget(self.combo_method)

        self.btn_calc = QPushButton("执行 QI 计算")
        self.btn_calc.setMinimumHeight(50)
        self.btn_calc.setStyleSheet("background-color: #2E7D32; color: white; font-weight: bold; font-size: 14px;")
        self.btn_calc.clicked.connect(self.calculate_qi)
        algo_layout.addWidget(self.btn_calc)
        algo_group.setLayout(algo_layout)
        left_panel.addWidget(algo_group)

        main_layout.addLayout(left_panel, 1)

        right_panel = QVBoxLayout()

        result_group = QGroupBox("QI 计算结果预览 (仅显示已选参数与结果)")
        result_layout = QVBoxLayout()
        self.table_res = QTableWidget()
        result_layout.addWidget(self.table_res)

        export_layout = QHBoxLayout()
        self.btn_export_data = QPushButton("导出表格 (XLSX/CSV)")
        self.btn_export_data.setMinimumHeight(35)
        self.btn_export_data.clicked.connect(self.export_data)

        self.btn_export_shp = QPushButton("导出为 SHP (含空间信息)")
        self.btn_export_shp.setMinimumHeight(35)
        self.btn_export_shp.setStyleSheet("background-color: #1976D2; color: white;")
        self.btn_export_shp.clicked.connect(self.export_shp)

        export_layout.addWidget(self.btn_export_data)
        export_layout.addWidget(self.btn_export_shp)

        result_layout.addLayout(export_layout)
        result_group.setLayout(result_layout)

        right_panel.addWidget(result_group)
        main_layout.addLayout(right_panel, 2)

    def load_data(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "打开文件", "", "Files (*.csv *.xlsx *.shp)")
        if not file_path:
            return

        try:
            self.gdf = None
            if file_path.lower().endswith(".shp"):
                self.gdf = gpd.read_file(file_path)
                self.df = pd.DataFrame(self.gdf.drop(columns="geometry"))
                self.lbl_status.setText(f"已加载SHP: {len(self.df)} 行")
            elif file_path.lower().endswith(".csv"):
                self.df = pd.read_csv(file_path)
                self.lbl_status.setText(f"已加载CSV: {len(self.df)} 行")
            else:
                self.df = pd.read_excel(file_path)
                self.lbl_status.setText(f"已加载Excel: {len(self.df)} 行")

            self.update_variable_list()
            self._append_main_log(f"林分质量评价已加载数据：{file_path}")
        except Exception as exc:
            QMessageBox.critical(self, "错误", f"读取失败: {str(exc)}")

    def update_variable_list(self):
        self.list_vars.clear()
        if self.df is None:
            return

        cols = self.df.select_dtypes(include=[np.number]).columns.tolist()
        for col in cols:
            item = QListWidgetItem(col)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            self.list_vars.addItem(item)

    def calculate_qi(self):
        if self.df is None:
            QMessageBox.warning(self, "提示", "请先导入数据！")
            return

        selected_cols = []
        for i in range(self.list_vars.count()):
            item = self.list_vars.item(i)
            if item.checkState() == Qt.Checked:
                selected_cols.append(item.text())

        if not selected_cols:
            QMessageBox.warning(self, "提示", "请勾选参与计算的指标！")
            return

        data = self.df[selected_cols].copy().fillna(0)
        norm = (data - data.min()) / (data.max() - data.min() + 1e-9)

        method = self.combo_method.currentText()
        try:
            if "熵权" in method:
                p = (norm + 0.0001) / (norm + 0.0001).sum()
                e = -1 / np.log(len(norm)) * (p * np.log(p)).sum()
                w = (1 - e) / (1 - e).sum()
            elif "等权" in method:
                w = np.array([1 / len(selected_cols)] * len(selected_cols))
            else:
                pca = PCA(n_components=1)
                pca.fit(norm)
                w = np.abs(pca.components_[0]) / np.abs(pca.components_[0]).sum()

            self.df["Forest_QI"] = (norm * w).sum(axis=1)
            if self.gdf is not None:
                self.gdf["Forest_QI"] = self.df["Forest_QI"]

            self._last_selected_cols = selected_cols
            self.show_results(selected_cols)
            self._append_main_log(f"林分质量评价计算完成，方法：{method}")
        except Exception as exc:
            QMessageBox.critical(self, "计算错误", f"计算过程中出错: {str(exc)}")

    def show_results(self, selected_cols):
        if self.df is None or "Forest_QI" not in self.df.columns:
            return

        columns_to_show = selected_cols + ["Forest_QI"]
        display_df = self.df[columns_to_show].head(100)

        self.table_res.setRowCount(display_df.shape[0])
        self.table_res.setColumnCount(display_df.shape[1])
        self.table_res.setHorizontalHeaderLabels(display_df.columns)

        header = self.table_res.horizontalHeader()
        header.setStretchLastSection(True)

        for i in range(display_df.shape[0]):
            for j in range(display_df.shape[1]):
                val = display_df.iloc[i, j]
                if isinstance(val, (float, np.floating)):
                    text = f"{float(val):.4f}"
                else:
                    text = str(val)
                self.table_res.setItem(i, j, QTableWidgetItem(text))

    def export_data(self):
        if self.df is None or "Forest_QI" not in self.df.columns:
            QMessageBox.warning(self, "提示", "请先执行计算后再导出！")
            return

        path, _ = QFileDialog.getSaveFileName(self, "保存完整表格", "", "Excel (*.xlsx);;CSV (*.csv)")
        if not path:
            return

        try:
            if path.lower().endswith(".csv"):
                self.df.to_csv(path, index=False, encoding="utf-8-sig")
            else:
                if not path.lower().endswith(".xlsx"):
                    path += ".xlsx"
                self.df.to_excel(path, index=False)
            QMessageBox.information(self, "成功", "表格已保存")
            self._append_main_log(f"林分质量评价表格已导出：{path}")
        except Exception as exc:
            QMessageBox.critical(self, "导出错误", str(exc))

    def export_shp(self):
        if self.gdf is None:
            QMessageBox.warning(self, "提示", "当前没有加载带有空间信息的 SHP 文件！\n请先导入 SHP 数据。")
            return
        if "Forest_QI" not in self.gdf.columns:
            QMessageBox.warning(self, "提示", "请先执行计算后再导出！")
            return

        path, _ = QFileDialog.getSaveFileName(self, "保存空间数据", "", "Shapefile (*.shp)")
        if not path:
            return

        try:
            if not path.lower().endswith(".shp"):
                path += ".shp"
            self.gdf.to_file(path, encoding="utf-8")
            QMessageBox.information(self, "成功", "SHP文件已生成，可在 ArcGIS/QGIS 中打开并进行符号化显示。")
            self._append_main_log(f"林分质量评价SHP已导出：{path}")
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", str(exc))

    def _append_main_log(self, message: str):
        parent = self.parent()
        if parent is not None and hasattr(parent, "log_output"):
            parent.log_output.append(f"[林分质量评价] {message}")


# 向后兼容旧类名
ForestQIApp = StandQualityWindow


if __name__ == "__main__":
    app = QApplication(sys.argv)
    apply_qt_app_style(app)
    window = StandQualityWindow()
    window.show()
    sys.exit(app.exec_())

