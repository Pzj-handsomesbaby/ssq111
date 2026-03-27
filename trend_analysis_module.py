import os
import re
import sys
import traceback

import numpy as np
import pandas as pd
import rasterio
from scipy.stats import kendalltau, linregress, t as t_dist, theilslopes

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure


def safe_makedirs(path: str):
    if path and not os.path.exists(path):
        os.makedirs(path, exist_ok=True)


def extract_year_from_name(filename: str):
    name = os.path.basename(filename)
    match = re.search(r"(19|20)\d{2}", name)
    if match:
        return int(match.group())
    return None


def read_table_file(file_path: str) -> pd.DataFrame:
    ext = os.path.splitext(file_path)[1].lower()
    if ext in [".xlsx", ".xls", ".xlsm", ".xlsb", ".ods"]:
        return pd.read_excel(file_path)
    if ext == ".csv":
        return pd.read_csv(file_path)
    raise ValueError(f"不支持的表格格式: {ext}")


def choose_default_xy_columns(df: pd.DataFrame):
    cols = list(df.columns)
    lower_cols = [str(c).lower() for c in cols]

    x_col = None
    for c, lc in zip(cols, lower_cols):
        if any(k in lc for k in ["year", "time", "date", "年份", "时间", "日期"]):
            x_col = c
            break
    if x_col is None and cols:
        x_col = cols[0]

    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    y_col = None
    for c in numeric_cols:
        if c != x_col:
            y_col = c
            break

    if y_col is None:
        for c in cols:
            if c != x_col:
                y_col = c
                break

    return x_col, y_col


def linear_trend_pixelwise(stack: np.ndarray, years: np.ndarray):
    t_count, height, width = stack.shape
    y_series = stack.reshape(t_count, -1).astype(np.float64)

    valid = np.all(np.isfinite(y_series), axis=0)

    slope = np.full(y_series.shape[1], np.nan, dtype=np.float32)
    intercept = np.full(y_series.shape[1], np.nan, dtype=np.float32)
    r2 = np.full(y_series.shape[1], np.nan, dtype=np.float32)
    pvalue = np.full(y_series.shape[1], np.nan, dtype=np.float32)
    mean_arr = np.full(y_series.shape[1], np.nan, dtype=np.float32)

    if valid.sum() == 0:
        return (
            slope.reshape(height, width),
            intercept.reshape(height, width),
            r2.reshape(height, width),
            pvalue.reshape(height, width),
            mean_arr.reshape(height, width),
        )

    x = years.astype(np.float64)
    x_mean = x.mean()
    x_centered = x - x_mean
    ssx = np.sum(x_centered ** 2)

    y_valid = y_series[:, valid]
    y_mean = np.mean(y_valid, axis=0)
    mean_arr[valid] = y_mean.astype(np.float32)

    cov_xy = np.sum(x_centered[:, None] * (y_valid - y_mean[None, :]), axis=0)
    slope_v = cov_xy / ssx
    intercept_v = y_mean - slope_v * x_mean

    y_hat = intercept_v[None, :] + slope_v[None, :] * x[:, None]

    ss_res = np.sum((y_valid - y_hat) ** 2, axis=0)
    ss_tot = np.sum((y_valid - y_mean[None, :]) ** 2, axis=0)
    with np.errstate(divide="ignore", invalid="ignore"):
        r2_v = 1 - ss_res / ss_tot
    r2_v[~np.isfinite(r2_v)] = np.nan

    ssy = np.sum((y_valid - y_mean[None, :]) ** 2, axis=0)
    with np.errstate(divide="ignore", invalid="ignore"):
        r = cov_xy / np.sqrt(ssx * ssy)
        r = np.clip(r, -1.0, 1.0)
        df = t_count - 2
        t_stat = r * np.sqrt(df / (1 - r ** 2))
        p_v = 2 * (1 - t_dist.cdf(np.abs(t_stat), df))
    p_v[~np.isfinite(p_v)] = np.nan

    slope[valid] = slope_v.astype(np.float32)
    intercept[valid] = intercept_v.astype(np.float32)
    r2[valid] = r2_v.astype(np.float32)
    pvalue[valid] = p_v.astype(np.float32)

    return (
        slope.reshape(height, width),
        intercept.reshape(height, width),
        r2.reshape(height, width),
        pvalue.reshape(height, width),
        mean_arr.reshape(height, width),
    )


def write_tif(output_path, array, profile, nodata=np.nan):
    new_profile = profile.copy()
    new_profile.update(driver="GTiff", count=1, dtype=str(array.dtype), nodata=nodata)
    with rasterio.open(output_path, "w", **new_profile) as dst:
        dst.write(array, 1)


class ExcelTrendTab(QWidget):
    analysis_finished = pyqtSignal(list, str)

    def __init__(self):
        super().__init__()
        self.file_path = ""
        self.df = pd.DataFrame()
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()

        file_box = QGroupBox("Excel / CSV 输入")
        file_layout = QHBoxLayout()
        self.file_edit = QLineEdit()
        self.file_edit.setPlaceholderText("选择 xlsx / xls / csv 文件")
        btn_browse = QPushButton("选择文件")
        btn_browse.clicked.connect(self.browse_file)
        file_layout.addWidget(self.file_edit)
        file_layout.addWidget(btn_browse)
        file_box.setLayout(file_layout)

        col_box = QGroupBox("列设置")
        form = QFormLayout()
        self.x_combo = QComboBox()
        self.y_combo = QComboBox()
        form.addRow("时间列 / X列：", self.x_combo)
        form.addRow("数值列 / Y列：", self.y_combo)
        col_box.setLayout(form)

        out_box = QGroupBox("输出")
        out_layout = QHBoxLayout()
        self.out_edit = QLineEdit()
        self.out_edit.setPlaceholderText("选择输出目录")
        btn_out = QPushButton("输出目录")
        btn_out.clicked.connect(self.browse_output_dir)
        out_layout.addWidget(self.out_edit)
        out_layout.addWidget(btn_out)
        out_box.setLayout(out_layout)

        btn_layout = QHBoxLayout()
        run_btn = QPushButton("开始 Excel 趋势分析")
        run_btn.clicked.connect(self.run_analysis)
        btn_layout.addStretch()
        btn_layout.addWidget(run_btn)

        self.log = QTextEdit()
        self.log.setReadOnly(True)

        layout.addWidget(file_box)
        layout.addWidget(col_box)
        layout.addWidget(out_box)
        layout.addLayout(btn_layout)
        layout.addWidget(QLabel("运行日志 / 结果"))
        layout.addWidget(self.log)
        self.setLayout(layout)

    def append_log(self, text: str):
        self.log.append(text)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择表格文件",
            "",
            "Table Files (*.xlsx *.xls *.xlsm *.xlsb *.ods *.csv)",
        )
        if not file_path:
            return
        self.file_path = file_path
        self.file_edit.setText(file_path)

        try:
            self.df = read_table_file(file_path)
            cols = [str(c) for c in self.df.columns]
            self.x_combo.clear()
            self.y_combo.clear()
            self.x_combo.addItems(cols)
            self.y_combo.addItems(cols)

            x_col, y_col = choose_default_xy_columns(self.df)
            if x_col is not None:
                self.x_combo.setCurrentText(str(x_col))
            if y_col is not None:
                self.y_combo.setCurrentText(str(y_col))

            self.append_log(f"已加载表格: {file_path}")
            self.append_log(f"字段: {cols}")
        except Exception as exc:
            QMessageBox.critical(self, "读取失败", str(exc))
            self.append_log("读取失败：\n" + traceback.format_exc())

    def browse_output_dir(self):
        out_dir = QFileDialog.getExistingDirectory(self, "选择输出目录", "")
        if out_dir:
            self.out_edit.setText(out_dir)

    def run_analysis(self):
        try:
            if self.df.empty:
                raise ValueError("请先加载 Excel/CSV 文件。")

            x_col = self.x_combo.currentText()
            y_col = self.y_combo.currentText()
            out_dir = self.out_edit.text().strip()
            if not out_dir:
                raise ValueError("请选择输出目录。")

            safe_makedirs(out_dir)
            data = self.df[[x_col, y_col]].copy().dropna()
            data[x_col] = pd.to_numeric(data[x_col], errors="coerce")
            data[y_col] = pd.to_numeric(data[y_col], errors="coerce")
            data = data.dropna()

            if len(data) < 3:
                raise ValueError("有效数据少于 3 行，无法进行趋势分析。")

            x = data[x_col].values.astype(float)
            y = data[y_col].values.astype(float)
            order = np.argsort(x)
            x = x[order]
            y = y[order]

            reg = linregress(x, y)
            tau, tau_p = kendalltau(x, y)
            sen = theilslopes(y, x)

            result_df = pd.DataFrame(
                {
                    "metric": [
                        "n",
                        "linear_slope",
                        "linear_intercept",
                        "r2",
                        "linear_pvalue",
                        "kendall_tau",
                        "kendall_pvalue",
                        "sen_slope",
                        "sen_intercept",
                    ],
                    "value": [
                        len(x),
                        reg.slope,
                        reg.intercept,
                        reg.rvalue ** 2,
                        reg.pvalue,
                        tau,
                        tau_p,
                        sen.slope,
                        sen.intercept,
                    ],
                }
            )

            result_excel = os.path.join(out_dir, "excel_trend_result.xlsx")
            result_csv = os.path.join(out_dir, "excel_trend_result.csv")
            png_path = os.path.join(out_dir, "excel_trend_plot.png")
            result_df.to_excel(result_excel, index=False)
            result_df.to_csv(result_csv, index=False, encoding="utf-8-sig")

            plt.figure(figsize=(9, 6))
            plt.scatter(x, y, s=50, label="Observed")
            plt.plot(x, reg.intercept + reg.slope * x, linewidth=2, label="Linear fit")
            plt.plot(x, sen.intercept + sen.slope * x, linestyle="--", linewidth=2, label="Sen slope")
            plt.xlabel(x_col)
            plt.ylabel(y_col)
            plt.title("Trend Analysis")
            plt.legend()
            plt.tight_layout()
            plt.savefig(png_path, dpi=300)
            plt.close()

            self.append_log("Excel 趋势分析完成。")
            self.append_log(f"结果表: {result_excel}")
            self.append_log(f"结果图: {png_path}")

            result_items = [
                ("Excel 趋势图(PNG)", png_path),
                ("Excel 结果表(XLSX)", result_excel),
                ("Excel 结果表(CSV)", result_csv),
            ]
            self.analysis_finished.emit(result_items, png_path)
            QMessageBox.information(self, "完成", "Excel 趋势分析完成。")
        except Exception as exc:
            QMessageBox.critical(self, "运行失败", str(exc))
            self.append_log("运行失败：\n" + traceback.format_exc())


class RasterTrendTab(QWidget):
    analysis_finished = pyqtSignal(list, str)

    def __init__(self):
        super().__init__()
        self.file_list = []
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()
        input_box = QGroupBox("栅格时序输入")
        input_layout = QVBoxLayout()

        btn_row = QHBoxLayout()
        btn_add = QPushButton("添加 TIF")
        btn_add.clicked.connect(self.add_files)
        btn_remove = QPushButton("删除选中")
        btn_remove.clicked.connect(self.remove_selected)
        btn_clear = QPushButton("清空")
        btn_clear.clicked.connect(self.clear_files)
        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_remove)
        btn_row.addWidget(btn_clear)
        btn_row.addStretch()

        self.list_widget = QListWidget()
        input_layout.addLayout(btn_row)
        input_layout.addWidget(self.list_widget)
        input_box.setLayout(input_layout)

        out_box = QGroupBox("输出")
        out_layout = QHBoxLayout()
        self.out_edit = QLineEdit()
        self.out_edit.setPlaceholderText("选择输出目录")
        btn_out = QPushButton("输出目录")
        btn_out.clicked.connect(self.browse_output_dir)
        out_layout.addWidget(self.out_edit)
        out_layout.addWidget(btn_out)
        out_box.setLayout(out_layout)

        run_layout = QHBoxLayout()
        run_btn = QPushButton("开始栅格趋势分析")
        run_btn.clicked.connect(self.run_analysis)
        run_layout.addStretch()
        run_layout.addWidget(run_btn)

        self.log = QTextEdit()
        self.log.setReadOnly(True)

        layout.addWidget(input_box)
        layout.addWidget(out_box)
        layout.addLayout(run_layout)
        layout.addWidget(QLabel("运行日志 / 结果"))
        layout.addWidget(self.log)
        self.setLayout(layout)

    def append_log(self, text: str):
        self.log.append(text)

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择多个 TIF 文件", "", "GeoTIFF (*.tif *.tiff)")
        for f in files:
            if f not in self.file_list:
                self.file_list.append(f)
                self.list_widget.addItem(QListWidgetItem(f))

    def remove_selected(self):
        for item in self.list_widget.selectedItems():
            path = item.text()
            if path in self.file_list:
                self.file_list.remove(path)
            self.list_widget.takeItem(self.list_widget.row(item))

    def clear_files(self):
        self.file_list = []
        self.list_widget.clear()

    def browse_output_dir(self):
        out_dir = QFileDialog.getExistingDirectory(self, "选择输出目录", "")
        if out_dir:
            self.out_edit.setText(out_dir)

    def run_analysis(self):
        try:
            if len(self.file_list) < 2:
                raise ValueError("至少需要 2 个 TIF 文件进行趋势分析。")
            out_dir = self.out_edit.text().strip()
            if not out_dir:
                raise ValueError("请选择输出目录。")
            safe_makedirs(out_dir)

            years = [extract_year_from_name(f) for f in self.file_list]
            if all(y is not None for y in years):
                sorted_pairs = sorted(zip(self.file_list, years), key=lambda x: x[1])
                file_list = [p[0] for p in sorted_pairs]
                years = np.array([p[1] for p in sorted_pairs], dtype=np.float64)
            else:
                file_list = list(self.file_list)
                years = np.arange(1, len(file_list) + 1, dtype=np.float64)

            with rasterio.open(file_list[0]) as src0:
                profile = src0.profile.copy()
                base_shape = (src0.height, src0.width)
                base_crs = src0.crs
                base_transform = src0.transform

            stack_list = []
            for tif in file_list:
                with rasterio.open(tif) as src:
                    if (src.height, src.width) != base_shape:
                        raise ValueError(f"栅格尺寸不一致: {tif}")
                    if src.crs != base_crs:
                        raise ValueError(f"投影不一致: {tif}")
                    if src.transform != base_transform:
                        raise ValueError(f"空间范围/像元大小不一致: {tif}")
                    arr = src.read(1).astype(np.float32)
                    nodata = src.nodata
                    if nodata is not None:
                        arr = np.where(arr == nodata, np.nan, arr)
                    stack_list.append(arr)

            stack = np.stack(stack_list, axis=0)
            slope, intercept, r2_arr, pvalue, mean_arr = linear_trend_pixelwise(stack, years)
            direction = np.full_like(slope, np.nan, dtype=np.float32)
            direction[slope > 0] = 1
            direction[slope < 0] = -1
            direction[slope == 0] = 0

            out_profile = profile.copy()
            out_profile.update(dtype="float32", count=1)
            slope_tif = os.path.join(out_dir, "slope.tif")
            intercept_tif = os.path.join(out_dir, "intercept.tif")
            r2_tif = os.path.join(out_dir, "r2.tif")
            pvalue_tif = os.path.join(out_dir, "pvalue.tif")
            mean_tif = os.path.join(out_dir, "mean.tif")
            direction_tif = os.path.join(out_dir, "direction.tif")

            write_tif(slope_tif, slope.astype(np.float32), out_profile)
            write_tif(intercept_tif, intercept.astype(np.float32), out_profile)
            write_tif(r2_tif, r2_arr.astype(np.float32), out_profile)
            write_tif(pvalue_tif, pvalue.astype(np.float32), out_profile)
            write_tif(mean_tif, mean_arr.astype(np.float32), out_profile)
            write_tif(direction_tif, direction.astype(np.float32), out_profile)

            # 额外导出 PNG：栅格渲染图与斜率直方图，便于直接查看与汇报。
            slope_png = os.path.join(out_dir, "slope_render.png")
            hist_png = os.path.join(out_dir, "slope_histogram.png")

            plt.figure(figsize=(9, 6))
            plt.imshow(slope, cmap="RdYlGn")
            plt.title("Slope Raster")
            plt.colorbar(shrink=0.8)
            plt.axis("off")
            plt.tight_layout()
            plt.savefig(slope_png, dpi=300)
            plt.close()

            valid = slope[np.isfinite(slope)]
            if valid.size > 0:
                plt.figure(figsize=(9, 6))
                plt.hist(valid, bins=50)
                plt.title("Slope Histogram")
                plt.xlabel("Slope")
                plt.ylabel("Frequency")
                plt.tight_layout()
                plt.savefig(hist_png, dpi=300)
                plt.close()

            result_items = [
                ("Slope(TIF)", slope_tif),
                ("Intercept(TIF)", intercept_tif),
                ("R2(TIF)", r2_tif),
                ("PValue(TIF)", pvalue_tif),
                ("Mean(TIF)", mean_tif),
                ("Direction(TIF)", direction_tif),
                ("Slope 渲染图(PNG)", slope_png),
            ]
            if os.path.exists(hist_png):
                result_items.append(("Slope 直方图(PNG)", hist_png))

            self.analysis_finished.emit(result_items, slope_tif)

            self.append_log("栅格趋势分析完成。")
            QMessageBox.information(self, "完成", "栅格趋势分析完成。")
        except Exception as exc:
            QMessageBox.critical(self, "运行失败", str(exc))
            self.append_log("运行失败：\n" + traceback.format_exc())


class VisualizationTab(QWidget):
    def __init__(self):
        super().__init__()
        self.current_array = None
        self.current_title = ""
        self.current_source_path = ""
        self.result_path_set = set()
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()
        file_box = QGroupBox("结果可视化")
        file_layout = QHBoxLayout()

        self.result_combo = QComboBox()
        self.result_combo.setMinimumWidth(260)
        self.result_combo.currentIndexChanged.connect(self._display_selected_result)

        btn_show = QPushButton("显示所选结果")
        btn_show.clicked.connect(self._display_selected_result)
        btn_hist = QPushButton("显示直方图")
        btn_hist.clicked.connect(self.show_histogram)
        btn_save = QPushButton("保存当前视图PNG")
        btn_save.clicked.connect(self.save_current_png)

        file_layout.addWidget(self.result_combo)
        file_layout.addWidget(btn_show)
        file_layout.addWidget(btn_hist)
        file_layout.addWidget(btn_save)
        file_box.setLayout(file_layout)

        self.figure = Figure(figsize=(8, 6))
        self.canvas = FigureCanvas(self.figure)
        self.info_box = QTextEdit()
        self.info_box.setReadOnly(True)

        layout.addWidget(file_box)
        layout.addWidget(self.canvas)
        layout.addWidget(QLabel("结果信息"))
        layout.addWidget(self.info_box)
        self.setLayout(layout)

    def append_info(self, text: str):
        self.info_box.append(text)

    def add_result_options(self, items, preferred_path=""):
        for label, path in items:
            if not path or path in self.result_path_set:
                continue
            self.result_path_set.add(path)
            self.result_combo.addItem(label, path)

        if self.result_combo.count() == 0:
            return

        if preferred_path:
            for idx in range(self.result_combo.count()):
                if self.result_combo.itemData(idx) == preferred_path:
                    self.result_combo.setCurrentIndex(idx)
                    self._display_selected_result()
                    return

        if self.result_combo.currentIndex() < 0:
            self.result_combo.setCurrentIndex(0)
        self._display_selected_result()

    def _display_selected_result(self, *_):
        idx = self.result_combo.currentIndex()
        if idx < 0:
            return
        file_path = self.result_combo.itemData(idx)
        if not file_path:
            return
        self.display_file(file_path)

    def display_file(self, file_path):
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "文件不存在", f"未找到结果文件：\n{file_path}")
            return

        ext = os.path.splitext(file_path)[1].lower()
        self.current_source_path = file_path
        self.current_title = os.path.basename(file_path)

        try:
            if ext == ".png":
                img = plt.imread(file_path)
                self.figure.clear()
                ax = self.figure.add_subplot(111)
                ax.imshow(img)
                ax.axis("off")
                ax.set_title(os.path.basename(file_path))
                self.canvas.draw()
                self.current_array = None
                self.info_box.clear()
                self.append_info(f"已显示 PNG：{file_path}")
            elif ext in [".tif", ".tiff"]:
                with rasterio.open(file_path) as src:
                    arr = src.read(1).astype(np.float32)
                    nodata = src.nodata
                    if nodata is not None:
                        arr = np.where(arr == nodata, np.nan, arr)
                self.figure.clear()
                ax = self.figure.add_subplot(111)
                im = ax.imshow(arr, cmap="RdYlGn")
                ax.axis("off")
                ax.set_title(os.path.basename(file_path))
                self.figure.colorbar(im, ax=ax, shrink=0.8)
                self.canvas.draw()
                self.current_array = arr
                self.info_box.clear()
                valid = arr[np.isfinite(arr)]
                self.append_info(f"已显示 TIF：{file_path}")
                self.append_info(f"有效像元数：{valid.size}")
            else:
                df = pd.read_csv(file_path) if ext == ".csv" else pd.read_excel(file_path)
                self.figure.clear()
                ax = self.figure.add_subplot(111)
                ax.axis("off")
                preview = df.head(10)
                table = ax.table(cellText=preview.values, colLabels=preview.columns, loc="center")
                table.auto_set_font_size(False)
                table.set_fontsize(9)
                table.scale(1, 1.5)
                ax.set_title(os.path.basename(file_path))
                self.canvas.draw()
                self.current_array = None
                self.info_box.clear()
                self.append_info(f"已显示表格：{file_path}")
                self.append_info(f"行数：{df.shape[0]}，列数：{df.shape[1]}")
        except Exception as exc:
            QMessageBox.critical(self, "打开失败", str(exc))

    def show_histogram(self):
        if self.current_array is None:
            QMessageBox.information(self, "提示", "当前不是栅格数据，无法显示直方图。")
            return
        valid = self.current_array[np.isfinite(self.current_array)]
        if valid.size == 0:
            QMessageBox.warning(self, "提示", "当前栅格没有有效值。")
            return

        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.hist(valid, bins=50)
        ax.set_title("栅格值分布直方图")
        ax.set_xlabel("Value")
        ax.set_ylabel("Frequency")
        self.canvas.draw()
        self.info_box.clear()
        self.append_info(f"已生成直方图：{self.current_title}")

    def save_current_png(self):
        if not self.current_title:
            QMessageBox.information(self, "提示", "当前没有可保存的图形。")
            return
        default_name = os.path.splitext(self.current_title)[0] + "_view.png"
        save_path, _ = QFileDialog.getSaveFileName(self, "保存当前视图", default_name, "PNG Files (*.png)")
        if not save_path:
            return
        try:
            self.figure.savefig(save_path, dpi=300, bbox_inches="tight")
            self.append_info(f"已保存PNG：{save_path}")
            QMessageBox.information(self, "完成", f"已保存：\n{save_path}")
        except Exception as exc:
            QMessageBox.critical(self, "保存失败", str(exc))


class TrendAnalysisApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("趋势预测-时序变化建模")
        self.resize(1000, 760)
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()

        title = QLabel("趋势预测-时序变化建模")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 22px; font-weight: bold; padding: 8px;")

        subtitle = QLabel("支持 Excel/CSV 与时序 GeoTIFF 的趋势分析、结果导出与结果可视化")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: #555; padding-bottom: 4px;")

        splitter = QSplitter(Qt.Horizontal)

        # 左侧：数据导入与分析
        self.analysis_tabs = QTabWidget()
        self.excel_tab = ExcelTrendTab()
        self.raster_tab = RasterTrendTab()
        self.analysis_tabs.addTab(self.excel_tab, "Excel 趋势分析")
        self.analysis_tabs.addTab(self.raster_tab, "栅格趋势分析")

        # 右侧：结果展示（不再单独作为 Tab）
        self.visual_panel = VisualizationTab()

        self.excel_tab.analysis_finished.connect(self._on_analysis_finished)
        self.raster_tab.analysis_finished.connect(self._on_analysis_finished)

        splitter.addWidget(self.analysis_tabs)
        splitter.addWidget(self.visual_panel)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)

        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addWidget(splitter)
        self.setLayout(layout)

    def _on_analysis_finished(self, items, preferred_path):
        self.visual_panel.add_result_options(items, preferred_path)


def main():
    app = QApplication(sys.argv)
    win = TrendAnalysisApp()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

