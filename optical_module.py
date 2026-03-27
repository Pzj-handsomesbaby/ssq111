# -*- coding: utf-8 -*-
"""
Created on Mon Oct 20 11:12:01 2025

@author: AAA
"""

from PyQt5.QtWidgets import (
    QWidget, QLabel, QLineEdit, QPushButton, QTextEdit,
    QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox, QSplitter, QDialog
)
from PyQt5.QtCore import Qt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure
import numpy as np
from osgeo import gdal, ogr
import os


_ALLOWED_FORMULA_FUNCTIONS = {
    "abs": np.abs,
    "sqrt": np.sqrt,
    "log": np.log,
    "log10": np.log10,
    "exp": np.exp,
    "sin": np.sin,
    "cos": np.cos,
    "tan": np.tan,
    "where": np.where,
    "minimum": np.minimum,
    "maximum": np.maximum,
    "clip": np.clip,
    "power": np.power,
    "nanmean": np.nanmean,
    "nanmax": np.nanmax,
    "nanmin": np.nanmin,
}


class RasterPreviewDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("栅格放大查看")
        self.resize(1200, 860)
        self._initial_xlim = None
        self._initial_ylim = None

        layout = QVBoxLayout(self)
        self.figure = Figure(constrained_layout=True)
        self.canvas = FigureCanvas(self.figure)
        self.toolbar = NavigationToolbar(self.canvas, self)

        button_row = QHBoxLayout()
        self.btn_fit = QPushButton("适应窗口")
        self.btn_reset = QPushButton("重置视图")
        self.btn_close = QPushButton("关闭")
        button_row.addWidget(self.btn_fit)
        button_row.addWidget(self.btn_reset)
        button_row.addStretch(1)
        button_row.addWidget(self.btn_close)

        layout.addWidget(self.toolbar)
        layout.addLayout(button_row)
        layout.addWidget(self.canvas, stretch=1)

        self.btn_fit.clicked.connect(self.fit_to_window)
        self.btn_reset.clicked.connect(self.reset_view)
        self.btn_close.clicked.connect(self.close)

    def set_array(self, array2d, title, cmap):
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        image = ax.imshow(array2d, cmap=cmap)
        ax.set_title(title)
        ax.axis("off")
        self.figure.colorbar(image, ax=ax, fraction=0.046, pad=0.04)
        self._initial_xlim = ax.get_xlim()
        self._initial_ylim = ax.get_ylim()
        self.canvas.draw_idle()

    def fit_to_window(self):
        if self.figure.axes:
            self.figure.axes[0].set_aspect("auto")
            self.canvas.draw_idle()

    def reset_view(self):
        if self.figure.axes and self._initial_xlim is not None and self._initial_ylim is not None:
            ax = self.figure.axes[0]
            ax.set_xlim(self._initial_xlim)
            ax.set_ylim(self._initial_ylim)
            ax.set_aspect("equal")
            self.canvas.draw_idle()


class OpticalModule(QWidget):
    def __init__(self, data_module=None, parent=None):
        super().__init__(parent)
        self.setWindowFlag(Qt.Window, True)
        self.data_module = data_module
        self.dataset_cache = None
        self.input_preview_array = None
        self.result_array = None
        self._input_colorbar = None
        self._result_colorbar = None
        self.zoom_dialog = None
        self.input_view_title = "计算前栅格预览"
        self.result_view_title = "计算后栅格结果"
        self.input_view_cmap = "viridis"
        self.result_view_cmap = "RdYlGn"

        self.setWindowTitle("栅格计算处理")
        self.resize(1480, 940)

        layout = QVBoxLayout(self)
        layout.setSpacing(6)

        row0 = QHBoxLayout()
        row0.addWidget(QLabel("当前已加载栅格:"))
        self.current_raster_label = QLabel("未检测到已加载栅格")
        self.current_raster_label.setStyleSheet("color: #555;")
        row0.addWidget(self.current_raster_label, stretch=1)
        self.btn_use_loaded = QPushButton("载入当前栅格")
        self.btn_refresh_loaded = QPushButton("刷新已加载状态")
        row0.addWidget(self.btn_use_loaded)
        row0.addWidget(self.btn_refresh_loaded)
        layout.addLayout(row0)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("输入TIF路径:"))
        self.input_tif = QLineEdit()
        row1.addWidget(self.input_tif)
        btn_in = QPushButton("浏览")
        row1.addWidget(btn_in)
        btn_in.clicked.connect(self.browse_input)
        layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("输出TIF路径:"))
        self.output_tif = QLineEdit()
        row2.addWidget(self.output_tif)
        btn_out = QPushButton("设置")
        row2.addWidget(btn_out)
        btn_out.clicked.connect(self.set_output)
        layout.addLayout(row2)

        row3 = QHBoxLayout()
        row3.addWidget(QLabel("裁剪Shapefile (可选):"))
        self.shp_path = QLineEdit()
        row3.addWidget(self.shp_path)
        btn_shp = QPushButton("选择")
        row3.addWidget(btn_shp)
        btn_shp.clicked.connect(self.browse_shp)
        layout.addLayout(row3)

        layout.addWidget(QLabel("栅格计算公式:"))
        self.formula_input = QTextEdit()
        self.formula_input.setPlaceholderText(
            "例如:\n(B4 - B3) / (B4 + B3 + 1e-6)\nwhere(B1 > 100, B2 * 0.8, B2)"
        )
        self.formula_input.setMinimumHeight(110)
        self.formula_input.setMaximumHeight(135)
        self.formula_input.setAcceptRichText(False)
        layout.addWidget(self.formula_input)

        formula_toolbar_row1 = QHBoxLayout()
        formula_toolbar_row1.setSpacing(4)
        self.btn_insert_sample = QPushButton("插入示例")
        self.btn_insert_b1 = QPushButton("B1")
        self.btn_insert_b2 = QPushButton("B2")
        self.btn_insert_b3 = QPushButton("B3")
        self.btn_insert_b4 = QPushButton("B4")
        self.btn_insert_where = QPushButton("where()")
        self.btn_insert_clip = QPushButton("clip()")
        self.btn_insert_sqrt = QPushButton("sqrt()")
        self.btn_insert_log = QPushButton("log()")
        self.btn_insert_power = QPushButton("power()")
        for button in [
            self.btn_insert_sample,
            self.btn_insert_b1,
            self.btn_insert_b2,
            self.btn_insert_b3,
            self.btn_insert_b4,
            self.btn_insert_where,
            self.btn_insert_clip,
            self.btn_insert_sqrt,
            self.btn_insert_log,
            self.btn_insert_power,
        ]:
            button.setMaximumHeight(28)
            formula_toolbar_row1.addWidget(button)
        layout.addLayout(formula_toolbar_row1)

        formula_toolbar_row2 = QHBoxLayout()
        formula_toolbar_row2.setSpacing(4)
        self.btn_insert_add = QPushButton("+")
        self.btn_insert_sub = QPushButton("-")
        self.btn_insert_mul = QPushButton("*")
        self.btn_insert_div = QPushButton("/")
        self.btn_insert_lparen = QPushButton("(")
        self.btn_insert_rparen = QPushButton(")")
        self.btn_undo = QPushButton("撤销")
        self.btn_redo = QPushButton("重做")
        self.btn_clear_formula = QPushButton("清空公式")
        for button in [
            self.btn_insert_add,
            self.btn_insert_sub,
            self.btn_insert_mul,
            self.btn_insert_div,
            self.btn_insert_lparen,
            self.btn_insert_rparen,
            self.btn_undo,
            self.btn_redo,
            self.btn_clear_formula,
        ]:
            button.setMaximumHeight(28)
            formula_toolbar_row2.addWidget(button)
        formula_toolbar_row2.addStretch(1)
        layout.addLayout(formula_toolbar_row2)

        self.btn_insert_sample.clicked.connect(self.insert_sample_formula)
        self.btn_insert_b1.clicked.connect(lambda: self.insert_formula_text("B1"))
        self.btn_insert_b2.clicked.connect(lambda: self.insert_formula_text("B2"))
        self.btn_insert_b3.clicked.connect(lambda: self.insert_formula_text("B3"))
        self.btn_insert_b4.clicked.connect(lambda: self.insert_formula_text("B4"))
        self.btn_insert_where.clicked.connect(lambda: self.insert_formula_text("where(B1 > 0, B2, 0)"))
        self.btn_insert_clip.clicked.connect(lambda: self.insert_formula_text("clip(B1, 0, 1)"))
        self.btn_insert_sqrt.clicked.connect(lambda: self.insert_formula_text("sqrt(B1)"))
        self.btn_insert_log.clicked.connect(lambda: self.insert_formula_text("log(B1 + 1e-6)"))
        self.btn_insert_power.clicked.connect(lambda: self.insert_formula_text("power(B1, 2)"))
        self.btn_insert_add.clicked.connect(lambda: self.insert_formula_text(" + "))
        self.btn_insert_sub.clicked.connect(lambda: self.insert_formula_text(" - "))
        self.btn_insert_mul.clicked.connect(lambda: self.insert_formula_text(" * "))
        self.btn_insert_div.clicked.connect(lambda: self.insert_formula_text(" / "))
        self.btn_insert_lparen.clicked.connect(lambda: self.insert_formula_text("("))
        self.btn_insert_rparen.clicked.connect(lambda: self.insert_formula_text(")"))
        self.btn_undo.clicked.connect(self.formula_input.undo)
        self.btn_redo.clicked.connect(self.formula_input.redo)
        self.btn_clear_formula.clicked.connect(self.formula_input.clear)
        self.btn_use_loaded.clicked.connect(self.use_loaded_raster)
        self.btn_refresh_loaded.clicked.connect(self.refresh_loaded_raster_hint)

        self.formula_hint = QLabel(
            "使用 B1、B2、B3... 表示第1/2/3波段；支持 abs、sqrt、log、exp、where、clip、power 等函数。"
        )
        self.formula_hint.setWordWrap(True)
        self.formula_hint.setMaximumHeight(34)
        layout.addWidget(self.formula_hint)

        button_row = QHBoxLayout()
        self.preview_btn = QPushButton("预览输入栅格")
        self.run_btn = QPushButton("开始处理")
        self.export_png_btn = QPushButton("导出结果图 PNG")
        button_row.addWidget(self.preview_btn)
        button_row.addWidget(self.run_btn)
        button_row.addWidget(self.export_png_btn)
        button_row.addStretch(1)
        layout.addLayout(button_row)
        self.preview_btn.clicked.connect(self.preview_input_raster)
        self.run_btn.clicked.connect(self.run_processing)
        self.export_png_btn.clicked.connect(self.export_result_png)

        self.visual_splitter = QSplitter(Qt.Horizontal)
        self.visual_splitter.setChildrenCollapsible(False)
        self.visual_splitter.setMinimumHeight(520)
        layout.addWidget(self.visual_splitter, stretch=3)

        input_panel = QWidget()
        input_layout = QVBoxLayout(input_panel)
        input_header = QHBoxLayout()
        input_header.addWidget(QLabel("计算前栅格预览"))
        self.btn_zoom_input = QPushButton("放大查看")
        self.btn_zoom_input.setMaximumWidth(110)
        input_header.addStretch(1)
        input_header.addWidget(self.btn_zoom_input)
        input_layout.addLayout(input_header)
        self.input_figure = Figure(constrained_layout=True)
        self.input_canvas = FigureCanvas(self.input_figure)
        self.input_toolbar = NavigationToolbar(self.input_canvas, self)
        input_layout.addWidget(self.input_toolbar)
        input_layout.addWidget(self.input_canvas, stretch=1)
        self.visual_splitter.addWidget(input_panel)

        result_panel = QWidget()
        result_layout = QVBoxLayout(result_panel)
        result_header = QHBoxLayout()
        result_header.addWidget(QLabel("计算后栅格结果"))
        self.btn_zoom_result = QPushButton("放大查看")
        self.btn_zoom_result.setMaximumWidth(110)
        result_header.addStretch(1)
        result_header.addWidget(self.btn_zoom_result)
        result_layout.addLayout(result_header)
        self.result_figure = Figure(constrained_layout=True)
        self.result_canvas = FigureCanvas(self.result_figure)
        self.result_toolbar = NavigationToolbar(self.result_canvas, self)
        result_layout.addWidget(self.result_toolbar)
        result_layout.addWidget(self.result_canvas, stretch=1)
        self.visual_splitter.addWidget(result_panel)
        self.visual_splitter.setSizes([760, 760])

        self.btn_zoom_input.clicked.connect(lambda: self.open_zoom_dialog(is_input=True))
        self.btn_zoom_result.clicked.connect(lambda: self.open_zoom_dialog(is_input=False))
        self.input_canvas.mpl_connect("button_press_event", self._on_input_canvas_click)
        self.result_canvas.mpl_connect("button_press_event", self._on_result_canvas_click)

        log_header = QHBoxLayout()
        log_header.addWidget(QLabel("运行日志:"))
        log_header.addStretch(1)
        self.log_tip = QLabel("图像可用工具栏缩放，也可双击或点“放大查看”。")
        self.log_tip.setStyleSheet("color: #666;")
        log_header.addWidget(self.log_tip)
        layout.addLayout(log_header)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(85)
        self.log.setMaximumHeight(100)
        layout.addWidget(self.log)

        self.refresh_loaded_raster_hint()
        self._render_placeholder(self.input_figure, "请选择输入栅格或载入当前栅格", is_input=True)
        self._render_placeholder(self.result_figure, "计算结果将在这里显示", is_input=False)

    # ---------------- 界面操作 ----------------
    def browse_input(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择TIF文件", "", "TIF文件 (*.tif *.tiff)")
        if path:
            self.input_tif.setText(path)
            self.dataset_cache = None
            self.preview_input_raster()

    def set_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "保存TIF文件", "", "TIF文件 (*.tif)")
        if path:
            self.output_tif.setText(path)

    def browse_shp(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择Shapefile", "", "Shapefile (*.shp)")
        if path:
            self.shp_path.setText(path)

    def insert_formula_text(self, text: str):
        cursor = self.formula_input.textCursor()
        cursor.insertText(text)
        self.formula_input.setFocus()

    def insert_sample_formula(self):
        self.formula_input.setPlainText("(B4 - B3) / (B4 + B3 + 1e-6)")
        self.formula_input.setFocus()

    def log_msg(self, msg):
        self.log.append(msg)
        self.log.ensureCursorVisible()

    def _on_input_canvas_click(self, event):
        if getattr(event, "dblclick", False):
            self.open_zoom_dialog(is_input=True)

    def _on_result_canvas_click(self, event):
        if getattr(event, "dblclick", False):
            self.open_zoom_dialog(is_input=False)

    def open_zoom_dialog(self, is_input: bool):
        array = self.input_preview_array if is_input else self.result_array
        if array is None:
            QMessageBox.information(self, "提示", "当前还没有可放大的图像。")
            return
        title = self.input_view_title if is_input else self.result_view_title
        cmap = self.input_view_cmap if is_input else self.result_view_cmap
        if self.zoom_dialog is None:
            self.zoom_dialog = RasterPreviewDialog(self)
        self.zoom_dialog.setWindowTitle(f"放大查看 - {title}")
        self.zoom_dialog.set_array(self._normalize_preview(array), title, cmap)
        self.zoom_dialog.show()
        self.zoom_dialog.raise_()
        self.zoom_dialog.activateWindow()

    def refresh_loaded_raster_hint(self):
        raster = self._get_latest_loaded_raster()
        if raster is None:
            self.current_raster_label.setText("未检测到已加载栅格")
            self.btn_use_loaded.setEnabled(False)
            return
        name = os.path.basename(raster.get("path", ""))
        shape = raster.get("shape")
        shape_text = f"{shape[1]}×{shape[0]}" if shape and len(shape) == 2 else "未知尺寸"
        self.current_raster_label.setText(f"{name}  ({shape_text})")
        self.btn_use_loaded.setEnabled(True)

    def use_loaded_raster(self):
        raster = self._get_latest_loaded_raster()
        if raster is None:
            QMessageBox.information(self, "提示", "当前没有已加载的栅格数据。")
            return
        self.input_tif.setText(raster["path"])
        self.dataset_cache = None
        self.preview_input_raster(prefer_loaded_preview=True)
        self.log_msg(f"已载入主界面当前栅格：{raster['path']}")

    def _get_latest_loaded_raster(self):
        if self.data_module is None:
            return None
        rasters = getattr(self.data_module, "rasters", None)
        if not rasters:
            return None
        return rasters[-1]

    def _open_input_dataset(self):
        tif_path = self.input_tif.text().strip()
        if not tif_path:
            raise ValueError("请输入输入TIF路径。")
        ds = gdal.Open(tif_path)
        if ds is None:
            raise ValueError("无法打开输入TIF。")
        return ds

    def _read_input_array(self):
        ds = self._open_input_dataset()
        arr = ds.ReadAsArray().astype(np.float32)
        if arr.ndim == 2:
            arr = arr[np.newaxis, :, :]
        return ds, arr

    def _normalize_preview(self, array2d):
        preview = np.asarray(array2d, dtype=np.float32)
        if preview.ndim != 2:
            preview = np.squeeze(preview)
        max_side = 1200
        height, width = preview.shape[:2]
        step = max(int(np.ceil(max(height / max_side, width / max_side, 1.0))), 1)
        preview = preview[::step, ::step].copy()
        finite = np.isfinite(preview)
        if finite.any():
            vmin = np.nanpercentile(preview[finite], 2)
            vmax = np.nanpercentile(preview[finite], 98)
            if np.isfinite(vmin) and np.isfinite(vmax) and vmax > vmin:
                preview = np.clip(preview, vmin, vmax)
        return preview

    def _draw_array(self, figure, array2d, title, cmap, is_input):
        figure.clear()
        ax = figure.add_subplot(111)
        preview = self._normalize_preview(array2d)
        image = ax.imshow(preview, cmap=cmap)
        ax.set_title(title)
        ax.axis("off")
        colorbar = figure.colorbar(image, ax=ax, fraction=0.046, pad=0.04)
        if is_input:
            self.input_view_title = title
            self.input_view_cmap = cmap
            self._input_colorbar = colorbar
            self.input_canvas.draw_idle()
        else:
            self.result_view_title = title
            self.result_view_cmap = cmap
            self._result_colorbar = colorbar
            self.result_canvas.draw_idle()

    def _render_placeholder(self, figure, text, is_input):
        figure.clear()
        ax = figure.add_subplot(111)
        ax.axis("off")
        ax.text(0.5, 0.5, text, ha="center", va="center", fontsize=12, color="#666")
        if is_input:
            self.input_view_title = text
            self.input_canvas.draw_idle()
        else:
            self.result_view_title = text
            self.result_canvas.draw_idle()

    def preview_input_raster(self, prefer_loaded_preview=False):
        self.refresh_loaded_raster_hint()
        tif_path = self.input_tif.text().strip()
        if not tif_path:
            self.log_msg("请先选择输入TIF。")
            self._render_placeholder(self.input_figure, "请选择输入栅格或载入当前栅格", is_input=True)
            return

        latest_loaded = self._get_latest_loaded_raster() if prefer_loaded_preview else None
        if latest_loaded and latest_loaded.get("path") == tif_path and latest_loaded.get("preview") is not None:
            self.input_preview_array = latest_loaded["preview"]
            title = f"输入预览：{os.path.basename(tif_path)}（主界面预览）"
            self._draw_array(self.input_figure, self.input_preview_array, title, "viridis", is_input=True)
            return

        try:
            ds, arr = self._read_input_array()
            self.dataset_cache = ds
            self.input_preview_array = arr[0]
            title = f"输入预览：{os.path.basename(tif_path)}（Band 1）"
            self._draw_array(self.input_figure, self.input_preview_array, title, "viridis", is_input=True)
            self.log_msg(f"输入栅格预览完成：{tif_path}")
        except Exception as exc:
            self.log_msg(str(exc))
            self._render_placeholder(self.input_figure, "输入栅格预览失败", is_input=True)

    def export_result_png(self):
        if self.result_array is None:
            QMessageBox.information(self, "提示", "当前还没有计算结果图可导出。")
            return
        path, _ = QFileDialog.getSaveFileName(self, "导出结果图为 PNG", "", "PNG 图片 (*.png)")
        if not path:
            return
        self.result_figure.savefig(path, dpi=200, bbox_inches="tight")
        self.log_msg(f"结果图已导出：{path}")

    def build_formula_context(self, arr):
        context = dict(_ALLOWED_FORMULA_FUNCTIONS)
        context["np"] = np
        for band_index in range(arr.shape[0]):
            band_name = "B{}".format(band_index + 1)
            context[band_name] = arr[band_index, :, :]
            context[band_name.lower()] = arr[band_index, :, :]
        return context

    def evaluate_formula(self, arr, formula: str):
        formula = (formula or "").strip()
        if not formula:
            raise ValueError("请输入栅格计算公式。")

        context = self.build_formula_context(arr)
        try:
            result = eval(formula, {"__builtins__": {}}, context)
        except NameError as exc:
            raise ValueError("公式中包含未定义的波段或函数：{}".format(exc))
        except Exception as exc:
            raise ValueError("公式计算失败：{}".format(exc))

        result = np.asarray(result, dtype=np.float32)
        if result.ndim == 0:
            result = np.full((arr.shape[1], arr.shape[2]), float(result), dtype=np.float32)
        if result.shape != (arr.shape[1], arr.shape[2]):
            raise ValueError(
                "公式结果维度不正确，应返回单波段二维栅格，当前得到：{}".format(result.shape)
            )
        return result

    # ---------------- 核心处理 ----------------
    def run_processing(self):
        self.log.clear()
        self.log_msg("开始栅格计算处理...")

        tif_path = self.input_tif.text().strip()
        out_path = self.output_tif.text().strip()
        shp_path = self.shp_path.text().strip()
        formula = self.formula_input.toPlainText().strip()

        if not tif_path:
            self.log_msg("请输入输入TIF路径。")
            return
        if not formula:
            self.log_msg("请输入栅格计算公式。")
            return

        self.log_msg("输入: {}".format(tif_path))
        self.log_msg("输出: {}".format(out_path or "未设置"))
        if shp_path:
            self.log_msg("启用Shapefile裁剪: {}".format(shp_path))
        self.log_msg("执行公式: {}".format(formula))

        try:
            ds, arr = self._read_input_array()
        except Exception as exc:
            self.log_msg(str(exc))
            return

        self.dataset_cache = ds
        self.input_preview_array = arr[0]
        self._draw_array(
            self.input_figure,
            self.input_preview_array,
            f"计算前：{os.path.basename(tif_path)}（Band 1）",
            "viridis",
            is_input=True,
        )

        if shp_path:
            shp_ds = ogr.Open(shp_path)
            if shp_ds is None:
                self.log_msg("无法打开Shapefile")
                return
            try:
                shp_layer = shp_ds.GetLayer()
                import rasterio
                import rasterio.features
                import shapely.wkt

                mask = np.zeros((arr.shape[1], arr.shape[2]), dtype=np.uint8)
                for feat in shp_layer:
                    geom = feat.GetGeometryRef().ExportToWkt()
                    poly = shapely.wkt.loads(geom)
                    mask |= rasterio.features.rasterize([poly], out_shape=mask.shape, fill=0, all_touched=True)
                arr[:, mask == 0] = np.nan
                self.log_msg("裁剪完成")
            except Exception as exc:
                self.log_msg(f"Shapefile裁剪失败：{exc}")
                return

        try:
            result = self.evaluate_formula(arr, formula)
        except ValueError as exc:
            self.log_msg(str(exc))
            return

        self.result_array = result

        if out_path:
            out_dir = os.path.dirname(out_path)
            if out_dir:
                os.makedirs(out_dir, exist_ok=True)
            driver = gdal.GetDriverByName("GTiff")
            out_ds = driver.Create(out_path, arr.shape[2], arr.shape[1], 1, gdal.GDT_Float32)
            out_ds.SetGeoTransform(ds.GetGeoTransform())
            out_ds.SetProjection(ds.GetProjection())
            out_ds.GetRasterBand(1).WriteArray(result)
            out_ds.FlushCache()
            del out_ds
            self.log_msg("结果已保存: {}".format(out_path))
            if self.data_module is not None and hasattr(self.data_module, "load_raster"):
                try:
                    self.data_module.load_raster(out_path)
                    self.log_msg("结果已同步到主界面地图可视化。")
                except Exception as exc:
                    self.log_msg(f"结果同步到主界面失败：{exc}")

        self._draw_array(
            self.result_figure,
            result,
            f"计算后：{formula}",
            "RdYlGn",
            is_input=False,
        )
        self.log_msg("处理完成，已显示计算前后栅格。")
