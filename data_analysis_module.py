# -*- coding: utf-8 -*-
"""
Created on Mon Nov  3 16:34:07 2025

@author: AAA
"""

from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog, QTextEdit
import geopandas as gpd
import pandas as pd
import matplotlib.pyplot as plt

class DataAnalysisModule(QWidget):
    def __init__(self, mode=None):
        super().__init__()
        self.setWindowTitle(f"数据分析 - {mode or '通用'}")
        self.resize(800, 600)
        layout = QVBoxLayout(self)
        self.text = QTextEdit()
        layout.addWidget(self.text)

        if mode == "excel":
            btn = QPushButton("打开Excel并统计")
            btn.clicked.connect(self.open_excel)
            layout.addWidget(btn)
        elif mode == "shpstat":
            btn = QPushButton("打开Shapefile并查看属性")
            btn.clicked.connect(self.open_shp)
            layout.addWidget(btn)
        elif mode == "plot":
            btn = QPushButton("绘制数据分布图")
            btn.clicked.connect(self.plot_demo)
            layout.addWidget(btn)

    def open_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择Excel", "", "Excel (*.xlsx)")
        if path:
            df = pd.read_excel(path)
            self.text.setText(str(df.describe()))

    def open_shp(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择Shapefile", "", "Shapefile (*.shp)")
        if path:
            gdf = gpd.read_file(path)
            self.text.setText(str(gdf.describe()))

    def plot_demo(self):
        import numpy as np
        data = np.random.randn(1000)
        plt.hist(data, bins=30)
        plt.title("数据分布示例")
        plt.show()
