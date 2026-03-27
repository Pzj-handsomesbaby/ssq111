# -*- coding: utf-8 -*-
"""
Created on Mon Nov  3 17:21:59 2025

@author: AAA
"""

# -*- coding: utf-8 -*-
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QTreeWidgetItem
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib import cm
from matplotlib.colors import Normalize
import numpy as np

class MapCanvas(QWidget):
    def __init__(self, tree_widget=None):
        super().__init__()
        self.tree_widget = tree_widget
        self.layout = QVBoxLayout(self)
        self.setLayout(self.layout)

        self.fig = Figure(figsize=(5,5))
        self.ax = self.fig.add_subplot(111)
        self.ax.axis("off")
        self.canvas = FigureCanvas(self.fig)
        self.layout.addWidget(self.canvas)

        # 状态标签
        self.status_label = QLabel("")
        self.layout.addWidget(self.status_label)

        # 数据存储
        self.raster_src = None
        self.raster_img = None
        self.shapefile_gdf = None
        self.opened_files = []

        self.zoom_scale = 1.0
        self.canvas.mpl_connect("scroll_event", self.on_scroll)
        self.canvas.mpl_connect("motion_notify_event", self.on_mouse_move)

    def add_raster(self, raster):
        self.raster_src = raster
        self.opened_files.append(raster.name)
        self.zoom_scale = 1.0
        self.update_tree()
        self.update_map()

    def add_shapefile(self, shp):
        self.shapefile_gdf = shp
        self.opened_files.append("Shapefile")
        self.zoom_scale = 1.0
        self.update_tree()
        self.update_map()

    def on_scroll(self, event):
        if self.raster_img is None:
            return
        if event.button == "up":
            self.zoom_scale *= 1.2
        elif event.button == "down":
            self.zoom_scale /= 1.2
        self.update_map()

    def on_mouse_move(self, event):
        if not self.raster_src or not event.inaxes:
            self.status_label.setText("")
            return
        try:
            col, row = int(event.xdata), int(event.ydata)
            val = self.raster_src.read(1)[row, col]
            self.status_label.setText(f"x={event.xdata:.2f}, y={event.ydata:.2f}, value={val:.2f}")
        except Exception:
            self.status_label.setText("")

    def update_tree(self):
        if not self.tree_widget:
            return
        for i in range(self.tree_widget.topLevelItemCount()):
            item = self.tree_widget.topLevelItem(i)
            if item.text(0) == "已加载数据":
                self.tree_widget.takeTopLevelItem(i)
                break
        loaded_root = QTreeWidgetItem(self.tree_widget, ["已加载数据"])
        for f in self.opened_files:
            QTreeWidgetItem(loaded_root, [f])
        self.tree_widget.expandAll()

    def update_map(self):
        self.ax.clear()
        self.ax.axis("off")
        if self.raster_src:
            data = self.raster_src.read(1)
            bounds = self.raster_src.bounds
            norm = Normalize(vmin=np.nanmin(data), vmax=np.nanmax(data))
            self.raster_img = self.ax.imshow(
                data, cmap=cm.viridis, norm=norm,
                extent=[bounds.left, bounds.right, bounds.bottom, bounds.top],
                origin='upper'
            )
            x_center = (bounds.left + bounds.right) / 2
            y_center = (bounds.bottom + bounds.top) / 2
            x_range = (bounds.right - bounds.left) / 2 / self.zoom_scale
            y_range = (bounds.top - bounds.bottom) / 2 / self.zoom_scale
            self.ax.set_xlim(x_center - x_range, x_center + x_range)
            self.ax.set_ylim(y_center - y_range, y_center + y_range)
        if self.shapefile_gdf is not None:
            self.shapefile_gdf.plot(ax=self.ax, facecolor="none", edgecolor="red")
        self.canvas.draw()
