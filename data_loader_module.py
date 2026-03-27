# -*- coding: utf-8 -*-
"""
Created on Mon Nov  3 17:21:36 2025

@author: AAA
"""

# -*- coding: utf-8 -*-
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel
import geopandas as gpd
import rasterio

class DataLoader(QWidget):
    def __init__(self, callback=None):
        super().__init__()
        self.setWindowTitle("选择数据文件")
        self.resize(400, 200)
        self.layout = QVBoxLayout(self)
        self.setLayout(self.layout)
        self.callback = callback

        # 栅格加载
        self.raster_btn = QPushButton("加载栅格文件 (.tif)")
        self.raster_btn.clicked.connect(self.load_raster)
        self.layout.addWidget(self.raster_btn)

        # Shapefile加载
        self.shp_btn = QPushButton("加载 Shapefile (.shp)")
        self.shp_btn.clicked.connect(self.load_shapefile)
        self.layout.addWidget(self.shp_btn)

        # 状态显示
        self.status_label = QLabel("")
        self.layout.addWidget(self.status_label)

        self.raster_src = None
        self.shapefile_gdf = None

    def load_raster(self):
        fname, _ = QFileDialog.getOpenFileName(self, "选择栅格文件", "", "Raster Files (*.tif *.tiff)")
        if not fname:
            return
        self.raster_src = rasterio.open(fname)
        self.status_label.setText(f"已加载：{fname}")
        if self.callback:
            self.callback(raster=self.raster_src)

    def load_shapefile(self):
        fname, _ = QFileDialog.getOpenFileName(self, "选择 Shapefile 文件", "", "Shapefile (*.shp)")
        if not fname:
            return
        self.shapefile_gdf = gpd.read_file(fname)
        self.status_label.setText(f"已加载：{fname}")
        if self.callback:
            self.callback(shapefile=self.shapefile_gdf)
