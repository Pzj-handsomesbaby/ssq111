# -*- coding: utf-8 -*-
import os
from typing import Optional

import numpy as np
import pandas as pd
from pandas.api.types import is_bool_dtype, is_numeric_dtype

from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, pyqtSignal
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QTabWidget, QSizePolicy,
    QTreeWidgetItem, QMenu, QLabel, QFileDialog,
    QTableView, QAbstractItemView
)

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure


# =========================
# 1) 表格：Model/View（不会爆栈）
# =========================
class PandasTableModel(QAbstractTableModel):
    def __init__(self, df: pd.DataFrame):
        super().__init__()
        self._df = df

    def rowCount(self, parent=QModelIndex()):
        return 0 if self._df is None else len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return 0 if self._df is None else len(self._df.columns)

    def data(self, index: QModelIndex, role=Qt.DisplayRole):
        if not index.isValid() or self._df is None:
            return None
        if role == Qt.DisplayRole:
            val = self._df.iat[index.row(), index.column()]
            if pd.isna(val):
                return ""
            return str(val)
        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role=Qt.DisplayRole):
        if self._df is None or role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        else:
            return str(section + 1)


class EditablePandasTableModel(QAbstractTableModel):
    dataframe_changed = pyqtSignal()

    def __init__(self, df: Optional[pd.DataFrame] = None):
        super().__init__()
        self._df = df if df is not None else pd.DataFrame()

    @property
    def dataframe(self) -> pd.DataFrame:
        return self._df

    def set_dataframe(self, df: pd.DataFrame):
        self.beginResetModel()
        self._df = df
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df.index)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df.columns)

    def data(self, index: QModelIndex, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        value = self._df.iat[index.row(), index.column()]
        if role in (Qt.DisplayRole, Qt.EditRole):
            if pd.isna(value):
                return ""
            return str(value)
        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section]) if section < len(self._df.columns) else ""
        return str(section + 1)

    def flags(self, index: QModelIndex):
        if not index.isValid():
            return Qt.ItemIsEnabled
        return Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable

    def setData(self, index: QModelIndex, value, role=Qt.EditRole):
        if role != Qt.EditRole or not index.isValid():
            return False
        column_name = self._df.columns[index.column()]
        text = "" if value is None else str(value)
        self._df.iat[index.row(), index.column()] = self._coerce_value(column_name, text)
        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
        self.dataframe_changed.emit()
        return True

    def insert_row(self, values=None):
        row = len(self._df.index)
        self.beginInsertRows(QModelIndex(), row, row)
        row_values = values if values is not None else {col: "" for col in self._df.columns}
        self._df = pd.concat([self._df, pd.DataFrame([row_values])], ignore_index=True)
        self.endInsertRows()
        self.dataframe_changed.emit()

    def remove_row(self, row: int):
        if row < 0 or row >= len(self._df.index):
            return False
        self.beginRemoveRows(QModelIndex(), row, row)
        self._df = self._df.drop(self._df.index[row]).reset_index(drop=True)
        self.endRemoveRows()
        self.dataframe_changed.emit()
        return True

    def add_column(self, name: str):
        column_name = name.strip() or f"字段{len(self._df.columns) + 1}"
        original_name = column_name
        suffix = 1
        while column_name in self._df.columns:
            suffix += 1
            column_name = f"{original_name}_{suffix}"
        self.beginResetModel()
        self._df[column_name] = ""
        self.endResetModel()
        self.dataframe_changed.emit()
        return column_name

    def _coerce_value(self, column_name: str, text: str):
        text = text.strip()
        series = self._df[column_name]
        if text == "":
            if is_numeric_dtype(series.dtype):
                return np.nan
            return ""
        if is_bool_dtype(series.dtype):
            lowered = text.lower()
            if lowered in {"true", "1", "yes", "是"}:
                return True
            if lowered in {"false", "0", "no", "否"}:
                return False
        if is_numeric_dtype(series.dtype):
            try:
                number = pd.to_numeric([text], errors="raise")[0]
                return number.item() if hasattr(number, "item") else number
            except Exception:
                return text
        return text


# =========================
# 2) ��模块
# =========================
class DataFileModule(QWidget):
    table_data_changed = pyqtSignal()
    """
    关键修复点（针对你现在的 0xC0000409）：
    - 栅格只读取“降采样预览图”，不会整幅读入内存
    - 不保存 rasterio dataset 对象（避免句柄/线程/驱动问题），只保存路径+元信息+预览数组
    - hover 不再 raster_src.read(1)，只在预览数组上取值
    - 表格改为 QTableView + QAbstractTableModel（Model/View），避免 QTableWidget 逐格 setItem 崩溃
    """

    def __init__(self, tree_widget=None, parent=None):
        super().__init__()
        self.parent_widget = parent
        self.last_table_error = ""
        self.current_table_path = ""
        self.current_table_model = None

        self.layout = QVBoxLayout(self)
        self.setLayout(self.layout)

        self.tree_widget = tree_widget
        if self.tree_widget:
            self.tree_widget.setContextMenuPolicy(Qt.CustomContextMenu)
            self.tree_widget.customContextMenuRequested.connect(self.right_click_menu)

        # ---------------- Tabs ----------------
        self.tabs = QTabWidget()
        self.layout.addWidget(self.tabs)

        # ---------------- 地图可视化 ----------------
        self.map_tab = QWidget()
        self.map_layout = QVBoxLayout(self.map_tab)

        self.fig_map = Figure()
        self.ax_map = self.fig_map.add_subplot(111)
        self.ax_map.axis("off")
        self.canvas_map = FigureCanvas(self.fig_map)
        self.canvas_map.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.map_layout.addWidget(self.canvas_map)

        self.hover_label_map = QLabel(self.map_tab)
        self.hover_label_map.setStyleSheet(
            "background-color: rgba(0,0,0,150); color: white; padding: 3px; border-radius: 3px;"
        )
        self.hover_label_map.setVisible(False)

        self.tabs.addTab(self.map_tab, "地图可视化")

        # ---------------- 表格可视化（Model/View） ----------------
        self.table_tab = QWidget()
        self.table_layout = QVBoxLayout(self.table_tab)

        self.table_view = QTableView()
        self.table_view.setSortingEnabled(False)
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table_layout.addWidget(self.table_view)

        self.tabs.addTab(self.table_tab, "表格可视化")

        # ---------------- 点云可视化 ----------------
        self.pointcloud_tab = QWidget()
        self.pc_layout = QVBoxLayout(self.pointcloud_tab)

        self.fig_pc = Figure()
        self.ax_pc = self.fig_pc.add_subplot(111, projection="3d")
        self.ax_pc.axis("off")
        self.canvas_pc = FigureCanvas(self.fig_pc)
        self.canvas_pc.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.pc_layout.addWidget(self.canvas_pc)

        self.hover_label_pc = QLabel(self.pointcloud_tab)
        self.hover_label_pc.setStyleSheet(
            "background-color: rgba(0,0,0,150); color: white; padding: 3px; border-radius: 3px;"
        )
        self.hover_label_pc.setVisible(False)

        self.tabs.addTab(self.pointcloud_tab, "点云可视化")

        # ---------------- 数据存储（按类型保存对象，���再存 dataset 句柄） ----------------
        # raster item: dict(path, preview, bounds, crs, transform, nodata, dtype, shape)
        self.rasters = []
        self.lidars = []       # dict(path, data Nx3)
        self.tables = []       # dict(path, df)
        self.shapefiles = []   # dict(path, gdf)

        self.zoom_scale = 1.0
        self.opened_files = []  # 仅用于树显示

        # ---------------- 事件绑定 ----------------
        self.canvas_map.mpl_connect("scroll_event", self.on_scroll_map)
        self.canvas_map.mpl_connect("motion_notify_event", self.on_mouse_move_map)
        self.canvas_pc.mpl_connect("motion_notify_event", self.on_mouse_move_pc)

    # ---------------- 日志输出 ----------------
    def _log(self, msg: str):
        if self.parent_widget is not None and hasattr(self.parent_widget, "log_output"):
            self.parent_widget.log_output.append(msg)

    # ---------------- 滚轮缩放地图 ----------------
    def on_scroll_map(self, event):
        if not self.rasters and not self.shapefiles:
            return
        if event.button == "up":
            self.zoom_scale *= 1.2
        elif event.button == "down":
            self.zoom_scale /= 1.2
        self.update_map()

    # ---------------- 鼠标悬停地图（只基于预览数组取值，不读原始栅格） ----------------
    def on_mouse_move_map(self, event):
        if event.inaxes is None:
            self.hover_label_map.setVisible(False)
            return

        x, y = event.xdata, event.ydata
        val_str = ""

        # 仅显示最后加载的栅格预览值
        if self.rasters:
            r = self.rasters[-1]
            preview = r.get("preview")
            bounds = r.get("bounds")  # (left, bottom, right, top)
            if preview is not None and bounds is not None:
                left, bottom, right, top = bounds
                h, w = preview.shape[:2]

                # 将坐标映射到预览数组行列
                # 注意：imshow extent=[left,right,bottom,top], origin='upper'
                if left <= x <= right and bottom <= y <= top:
                    col = int((x - left) / (right - left) * (w - 1))
                    row = int((top - y) / (top - bottom) * (h - 1))
                    if 0 <= row < h and 0 <= col < w:
                        v = preview[row, col]
                        if np.isnan(v):
                            val_str = "Raster: NaN"
                        else:
                            val_str = f"Raster: {float(v):.4f}"

        elif self.shapefiles:
            val_str = "Shapefile"

        if val_str:
            self.hover_label_map.setText(val_str)
            self.hover_label_map.move(int(event.guiEvent.x()) + 10, int(event.guiEvent.y()) + 10)
            self.hover_label_map.setVisible(True)
        else:
            self.hover_label_map.setVisible(False)

    # ---------------- 鼠标悬停点云 ----------------
    def on_mouse_move_pc(self, event):
        if event.inaxes is None:
            self.hover_label_pc.setVisible(False)
            return
        x, y = event.xdata, event.ydata
        val_str = ""

        if self.lidars:
            lidar_data = self.lidars[-1]["data"]
            if lidar_data is not None and len(lidar_data) > 0:
                # 简单最近点（预览用途）
                distances = np.sqrt((lidar_data[:, 0] - x) ** 2 + (lidar_data[:, 1] - y) ** 2)
                idx = int(np.argmin(distances))
                z = float(lidar_data[idx, 2])
                val_str = f"LiDAR Z: {z:.3f}"

        if val_str:
            self.hover_label_pc.setText(val_str)
            self.hover_label_pc.move(int(event.guiEvent.x()) + 10, int(event.guiEvent.y()) + 10)
            self.hover_label_pc.setVisible(True)
        else:
            self.hover_label_pc.setVisible(False)

    # ---------------- 通用加载文件接口 ----------------
    def load_file_dialog(self, mode):
        if mode == "raster":
            fname, _ = QFileDialog.getOpenFileName(self, "选择栅格文件", "", "Raster (*.tif *.tiff)")
            if fname:
                self.load_raster(fname)
        elif mode == "lidar":
            fname, _ = QFileDialog.getOpenFileName(self, "选择 LiDAR 文件", "", "LiDAR (*.las *.laz *.txt *.csv)")
            if fname:
                self.load_lidar(fname)
        elif mode == "shp":
            fname, _ = QFileDialog.getOpenFileName(self, "选择 Shapefile", "", "Shapefile (*.shp)")
            if fname:
                self.load_shapefile(fname)
        elif mode == "table":
            fname, _ = QFileDialog.getOpenFileName(self, "选择表格文件", "", "Excel/CSV (*.xls *.xlsx *.csv)")
            if fname:
                self.load_table(fname)

    # ---------------- 栅格：安全加载（降采样预览 + 关闭文件句柄） ----------------
    def load_raster(self, fname):
        self._log(f"加载栅格：{fname}")
        try:
            import rasterio
            from rasterio.enums import Resampling

            with rasterio.open(fname) as src:
                # 只取 Band1 预览
                band = 1
                height, width = src.height, src.width

                # 预览最大边控制（防止内存爆）
                max_side = 1200
                scale = max(height / max_side, width / max_side, 1.0)
                out_h = int(height / scale)
                out_w = int(width / scale)

                # 读取降采样预览
                preview = src.read(
                    band,
                    out_shape=(out_h, out_w),
                    resampling=Resampling.nearest
                ).astype(np.float32)

                # nodata 处理
                nodata = src.nodata
                if nodata is not None:
                    preview[preview == nodata] = np.nan

                # 极端值裁剪（防止 imshow 溢出导致 matplotlib 警告/不稳定）
                # 用分位数拉伸到合理范围
                finite = np.isfinite(preview)
                if finite.any():
                    vmin = np.nanpercentile(preview[finite], 2)
                    vmax = np.nanpercentile(preview[finite], 98)
                    if vmax > vmin:
                        preview = np.clip(preview, vmin, vmax)

                bounds = (src.bounds.left, src.bounds.bottom, src.bounds.right, src.bounds.top)

                raster_item = {
                    "path": fname,
                    "preview": preview,
                    "bounds": bounds,
                    "crs": str(src.crs) if src.crs else None,
                    "transform": src.transform,
                    "nodata": nodata,
                    "dtype": str(src.dtypes[0]) if src.dtypes else None,
                    "shape": (height, width),
                }

            self.rasters.append(raster_item)
            self._append_opened_file(fname)
            self.zoom_scale = 1.0
            self.update_tree()
            self.update_map()
            self.tabs.setCurrentWidget(self.map_tab)
            self._log("栅格加载成功（预览模式）。")

        except Exception as e:
            self._log(f"栅格加载失败：{e}")

    # ---------------- LiDAR：安全加载 ----------------
    def load_lidar(self, fname):
        self._log(f"加载 LiDAR：{fname}")
        try:
            import laspy
            ext = os.path.splitext(fname)[1].lower()

            if ext in [".las", ".laz"]:
                las = laspy.read(fname)
                lidar_data = np.vstack([las.x, las.y, las.z]).T.astype(np.float32)
            else:
                # csv/txt: 尝试逗号分隔；失败则空格
                try:
                    lidar_data = np.loadtxt(fname, delimiter=",").astype(np.float32)
                except Exception:
                    lidar_data = np.loadtxt(fname).astype(np.float32)

                if lidar_data.ndim == 1:
                    lidar_data = lidar_data.reshape(-1, 3)

            # 点数太大时做抽样（否则 3D scatter 会卡/崩）
            if lidar_data.shape[0] > 300000:
                idx = np.random.choice(lidar_data.shape[0], 300000, replace=False)
                lidar_data = lidar_data[idx, :]

            self.lidars.append({"path": fname, "data": lidar_data})
            self._append_opened_file(fname)
            self.update_tree()
            self.update_pointcloud()
            self.tabs.setCurrentWidget(self.pointcloud_tab)
            self._log("LiDAR 加载成功。")

        except Exception as e:
            self._log(f"LiDAR 加载失败：{e}")

    # ---------------- Shapefile：延迟 import + 只读一次 ----------------
    def load_shapefile(self, fname):
        self._log(f"加载 Shapefile：{fname}")
        try:
            import geopandas as gpd
            gdf = gpd.read_file(fname)
            self.shapefiles.append({"path": fname, "gdf": gdf})
            self._append_opened_file(fname)
            self.update_tree()
            self.update_map()
            self.tabs.setCurrentWidget(self.map_tab)
            self._log("Shapefile 加载成功。")
        except Exception as e:
            self._log(f"Shapefile 加载失败：{e}")

    # ---------------- 表格：Model/View，不再逐格 setItem ----------------
    def load_table(self, fname):
        self._log(f"加载表格：{fname}")
        self.last_table_error = ""
        try:
            ext = os.path.splitext(fname)[1].lower()

            if ext in (".xls", ".xlsx"):
                try:
                    df = pd.read_excel(fname)
                except ImportError as e:
                    dependency = "openpyxl" if ext == ".xlsx" else "xlrd/openpyxl"
                    msg = (
                        f"表格加载失败：当前 Python 环境缺少 Excel 读取依赖 {dependency}。"
                        f"请先安装后重试，或将文件另存为 CSV 再导入。原始错误：{e}"
                    )
                    self.last_table_error = msg
                    self._log(msg)
                    return False
            else:
                try:
                    df = pd.read_csv(fname, encoding="utf-8")
                except Exception:
                    df = pd.read_csv(fname, encoding="gbk")

            max_rows = 200000
            if len(df) > max_rows:
                df = df.head(max_rows)
                self._log(f"提示：表格行数过大，仅预览前 {max_rows} 行。")

            self.set_table_dataframe(fname, df)
            self._log("表格加载成功（只读可视化）。")
            return True

        except Exception as e:
            self.last_table_error = f"表格加载失败：{e}"
            self._log(self.last_table_error)
            return False

    def set_table_dataframe(self, fname: str, df: pd.DataFrame, switch_to_tab: bool = True):
        self._upsert_table(fname, df)
        self.current_table_path = fname
        self.current_table_model = None
        model = PandasTableModel(df)
        self.table_view.setModel(model)
        self.table_view.resizeColumnsToContents()
        self._append_opened_file(fname)
        self.update_tree()
        if switch_to_tab:
            self.tabs.setCurrentWidget(self.table_tab)
        self.table_data_changed.emit()

    def _upsert_table(self, fname: str, df: pd.DataFrame):
        for idx, table in enumerate(self.tables):
            if table["path"] == fname:
                self.tables.pop(idx)
                break
        self.tables.append({"path": fname, "df": df})

    def _on_current_table_changed(self):
        if not self.current_table_path or self.current_table_model is None:
            return
        self._upsert_table(self.current_table_path, self.current_table_model.dataframe)
        self.table_view.resizeColumnsToContents()
        self.table_data_changed.emit()

    def get_last_table_error(self):
        return self.last_table_error

    def get_loaded_tables(self):
        return [
            {
                "path": table["path"],
                "name": os.path.basename(table["path"]),
                "df": table["df"],
            }
            for table in self.tables
        ]

    def get_latest_table(self):
        if not self.tables:
            return None
        table = self.tables[-1]
        return {
            "path": table["path"],
            "name": os.path.basename(table["path"]),
            "df": table["df"],
        }

    def get_table_by_path(self, path: str):
        for table in self.tables:
            if table["path"] == path:
                return {
                    "path": table["path"],
                    "name": os.path.basename(table["path"]),
                    "df": table["df"],
                }
        return None

    def get_numeric_columns(self, df: pd.DataFrame):
        numeric_columns = []
        for col in df.columns:
            series = pd.to_numeric(df[col], errors="coerce")
            if series.notna().any():
                numeric_columns.append(col)
        return numeric_columns

    # ---------------- 维护 opened_files（避免重复） ----------------
    def _append_opened_file(self, fname: str):
        if fname not in self.opened_files:
            self.opened_files.append(fname)

    # ---------------- 更新目录 ----------------
    def update_tree(self):
        if not self.tree_widget:
            return

        # 移除旧的“已加载数据”节点
        for i in range(self.tree_widget.topLevelItemCount()):
            item = self.tree_widget.topLevelItem(i)
            if item.text(0) == "已加载数据":
                self.tree_widget.takeTopLevelItem(i)
                break

        loaded_root = QTreeWidgetItem(self.tree_widget, ["已加载数据"])
        for f in self.opened_files:
            QTreeWidgetItem(loaded_root, [os.path.basename(f)])
        self.tree_widget.expandAll()

    # ---------------- 右键移除：按 path 精确删除 ----------------
    def right_click_menu(self, pos):
        item = self.tree_widget.itemAt(pos)
        if not item or item.parent() is None:
            return

        menu = QMenu()
        remove_action = menu.addAction("移除")
        action = menu.exec_(self.tree_widget.viewport().mapToGlobal(pos))
        if action != remove_action:
            return

        fname_base = item.text(0)
        had_table = any(os.path.basename(t["path"]) == fname_base for t in self.tables)

        # 从各类型列表中移除匹配 basename 的项
        self.rasters = [r for r in self.rasters if os.path.basename(r["path"]) != fname_base]
        self.lidars = [l for l in self.lidars if os.path.basename(l["path"]) != fname_base]
        self.tables = [t for t in self.tables if os.path.basename(t["path"]) != fname_base]
        self.shapefiles = [s for s in self.shapefiles if os.path.basename(s["path"]) != fname_base]

        # opened_files 移除
        self.opened_files = [f for f in self.opened_files if os.path.basename(f) != fname_base]

        # UI 刷新
        self.update_tree()
        self.update_map()
        self.update_pointcloud()

        # 若移除的是表格，清空 model
        current_model = self.table_view.model()
        if current_model is not None and hasattr(current_model, "_df"):
            # 如果当前显示的表格被删了，直接清空
            #（简单处理：若无表格则清空）
            if len(self.tables) == 0:
                self.table_view.setModel(None)

        if had_table:
            self.table_data_changed.emit()
        self._log(f"已移除：{fname_base}")

    # ---------------- 更新地图：只画预览，不读整幅 ----------------
    def update_map(self):
        self.ax_map.clear()
        self.ax_map.axis("off")

        all_x, all_y = [], []

        # 画栅格预览
        for r in self.rasters:
            preview = r.get("preview")
            bounds = r.get("bounds")
            if preview is None or bounds is None:
                continue

            left, bottom, right, top = bounds
            # 注意：extent 对应地理范围；origin='upper' 与 row 映射一致
            self.ax_map.imshow(
                preview,
                extent=[left, right, bottom, top],
                origin="upper"
            )
            all_x += [left, right]
            all_y += [bottom, top]

        # 画矢量（延迟 import，且只画边界以减少负担）
        for s in self.shapefiles:
            gdf = s.get("gdf")
            if gdf is None or len(gdf) == 0:
                continue

            try:
                # 只画边界，减少渲染压力
                gdf.boundary.plot(ax=self.ax_map, linewidth=1.0)
                tb = gdf.total_bounds  # [minx, miny, maxx, maxy]
                all_x += [tb[0], tb[2]]
                all_y += [tb[1], tb[3]]
            except Exception:
                # 任何矢量绘制错误都不影响主流程
                pass

        # 根据 zoom_scale 设置视野
        if all_x and all_y:
            x_center = (min(all_x) + max(all_x)) / 2.0
            y_center = (min(all_y) + max(all_y)) / 2.0
            x_range = (max(all_x) - min(all_x)) / 2.0 / max(self.zoom_scale, 1e-6)
            y_range = (max(all_y) - min(all_y)) / 2.0 / max(self.zoom_scale, 1e-6)

            if x_range > 0 and y_range > 0:
                self.ax_map.set_xlim(x_center - x_range, x_center + x_range)
                self.ax_map.set_ylim(y_center - y_range, y_center + y_range)

        # draw（避免频繁重绘导致卡顿/崩溃）
        self.canvas_map.draw_idle()
        self.hover_label_map.raise_()

    # ---------------- 更新点云：点数过多时已经抽样 ----------------
    def update_pointcloud(self):
        self.ax_pc.clear()
        self.ax_pc.axis("off")

        for l in self.lidars:
            data = l.get("data")
            if data is None or len(data) == 0:
                continue
            x, y, z = data[:, 0], data[:, 1], data[:, 2]
            # 不指定 cmap（避免部分后端不稳定），只用默认
            self.ax_pc.scatter(x, y, z, s=1)

        self.canvas_pc.draw_idle()
        self.hover_label_pc.raise_()
