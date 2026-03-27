# -*- coding: utf-8 -*-

import math

import numpy as np
import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)
from matplotlib import rcParams
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure

rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "DejaVu Sans"]
rcParams["axes.unicode_minus"] = False


class DataPlotWindow(QMainWindow):
    INDEX_KEY = "__row_index__"
    INDEX_LABEL = "行号（索引）"
    MAX_HIST_POINTS = 50000

    def __init__(self, data_module, parent=None):
        super().__init__(parent)
        self.data_module = data_module
        self.setWindowTitle("数据绘图")
        self.resize(1100, 760)

        self._build_ui()
        self._bind_signals()
        self.refresh_tables(auto_plot=False)

        if hasattr(self.data_module, "table_data_changed"):
            self.data_module.table_data_changed.connect(self.on_table_data_changed)

    def _build_ui(self):
        central = QWidget(self)
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        control_row = QHBoxLayout()
        layout.addLayout(control_row)

        control_panel = QWidget(central)
        form_layout = QFormLayout(control_panel)
        control_row.addWidget(control_panel, stretch=1)

        self.table_combo = QComboBox(control_panel)
        self.chart_combo = QComboBox(control_panel)
        self.chart_combo.addItems(["散点图", "折线图", "直方图"])
        self.x_combo = QComboBox(control_panel)
        self.y_combo = QComboBox(control_panel)

        self.bin_label = QLabel("直方图分箱", control_panel)
        self.bin_spin = QSpinBox(control_panel)
        self.bin_spin.setRange(5, 100)
        self.bin_spin.setValue(20)

        self.sample_label = QLabel("最大采样点数", control_panel)
        self.sample_spin = QSpinBox(control_panel)
        self.sample_spin.setRange(100, 100000)
        self.sample_spin.setSingleStep(500)
        self.sample_spin.setValue(8000)

        form_layout.addRow("数据表", self.table_combo)
        form_layout.addRow("图形类型", self.chart_combo)
        self.x_row_label = QLabel("X 轴字段", control_panel)
        self.y_row_label = QLabel("Y 轴字段", control_panel)
        form_layout.addRow(self.x_row_label, self.x_combo)
        form_layout.addRow(self.y_row_label, self.y_combo)
        form_layout.addRow(self.bin_label, self.bin_spin)
        form_layout.addRow(self.sample_label, self.sample_spin)

        button_panel = QWidget(central)
        button_layout = QVBoxLayout(button_panel)
        control_row.addWidget(button_panel)

        self.add_table_button = QPushButton("添加表格/CSV", button_panel)
        self.refresh_button = QPushButton("刷新数据", button_panel)
        self.plot_button = QPushButton("绘制图形", button_panel)
        self.save_button = QPushButton("导出图片", button_panel)

        button_layout.addWidget(self.add_table_button)
        button_layout.addWidget(self.refresh_button)
        button_layout.addWidget(self.plot_button)
        button_layout.addWidget(self.save_button)
        button_layout.addStretch(1)

        self.summary_label = QLabel("等待加载表格数据")
        self.status_label = QLabel("提示：请先在“表格数据”中加载 CSV 或 Excel，也可以直接点击右侧“添加表格/CSV”。")
        self.import_hint_label = QLabel("支持 *.csv / *.xlsx / *.xls；如果当前环境缺少 openpyxl，可优先导入 CSV。")
        self.summary_label.setStyleSheet("font-weight: bold;")
        self.status_label.setStyleSheet("color: #555;")
        self.import_hint_label.setStyleSheet("color: #777;")
        layout.addWidget(self.summary_label)
        layout.addWidget(self.status_label)
        layout.addWidget(self.import_hint_label)

        self.figure = Figure(constrained_layout=True)
        self.canvas = FigureCanvas(self.figure)
        self.toolbar = NavigationToolbar(self.canvas, self)
        layout.addWidget(self.toolbar)
        layout.addWidget(self.canvas, stretch=1)

        self._render_placeholder("请先加载表格数据，然后选择图形类型。")
        self._sync_control_visibility()

    def _bind_signals(self):
        self.add_table_button.clicked.connect(self.add_table_file)
        self.refresh_button.clicked.connect(lambda: self.refresh_tables(auto_plot=False))
        self.plot_button.clicked.connect(self.plot_current)
        self.save_button.clicked.connect(self.save_figure)
        self.table_combo.currentIndexChanged.connect(self.on_table_changed)
        self.chart_combo.currentTextChanged.connect(self.on_chart_changed)

    def add_table_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择表格或 CSV 文件",
            "",
            "表格文件 (*.csv *.xlsx *.xls)"
        )
        if not path:
            self.status_label.setText("已取消导入。")
            return

        if not hasattr(self.data_module, "load_table"):
            self.status_label.setText("当前数据模块不支持表格导入。")
            return

        ok = self.data_module.load_table(path)
        if ok:
            self.refresh_tables(auto_plot=False)
            table_name = self.data_module.get_latest_table()["name"] if hasattr(self.data_module, "get_latest_table") and self.data_module.get_latest_table() else path
            self.status_label.setText(f"已导入表格：{table_name}，现在可以选择字段绘图。")
        else:
            error_msg = self.data_module.get_last_table_error() if hasattr(self.data_module, "get_last_table_error") else "表格加载失败。"
            self.status_label.setText(error_msg)
            self._render_placeholder("导入失败，请检查文件格式或依赖后重试")

    def on_table_data_changed(self):
        self.refresh_tables(auto_plot=False)

    def refresh_tables(self, auto_plot=False):
        tables = self.data_module.get_loaded_tables() if hasattr(self.data_module, "get_loaded_tables") else []
        current_path = self.table_combo.currentData()

        self.table_combo.blockSignals(True)
        self.table_combo.clear()

        for table in tables:
            self.table_combo.addItem(table["name"], table["path"])

        self.table_combo.blockSignals(False)

        if not tables:
            self._set_controls_enabled(False)
            self.summary_label.setText("当前没有已加载的表格数据")
            self.status_label.setText("提示：请先从菜单“文件模块 -> 表格数据”加载 CSV 或 Excel，或直接点击“添加表格/CSV”。")
            self._render_placeholder("暂无表格数据可供绘图")
            return

        index = self.table_combo.findData(current_path, flags=Qt.MatchExactly)
        if index < 0:
            latest = self.data_module.get_latest_table() if hasattr(self.data_module, "get_latest_table") else None
            latest_path = latest["path"] if latest else tables[-1]["path"]
            index = self.table_combo.findData(latest_path, flags=Qt.MatchExactly)
            if index < 0:
                index = len(tables) - 1

        self.table_combo.setCurrentIndex(index)
        self._set_controls_enabled(True)
        self.update_columns(auto_plot=auto_plot)

    def on_table_changed(self):
        self.update_columns(auto_plot=False)

    def on_chart_changed(self):
        self._sync_control_visibility()
        self._apply_default_axis_selection()

    def update_columns(self, auto_plot=False):
        table = self.get_selected_table()
        if table is None:
            self._set_controls_enabled(False)
            self.summary_label.setText("未选择可用数据表")
            self.status_label.setText("提示：没有匹配的数据表。")
            self._render_placeholder("请选择有效的数据表")
            return

        df = table["df"]
        numeric_columns = self.data_module.get_numeric_columns(df) if hasattr(self.data_module, "get_numeric_columns") else []
        previous_x = self.x_combo.currentData()
        previous_y = self.y_combo.currentData()

        self.x_combo.blockSignals(True)
        self.y_combo.blockSignals(True)
        self.x_combo.clear()
        self.y_combo.clear()

        self.x_combo.addItem(self.INDEX_LABEL, self.INDEX_KEY)
        for column in numeric_columns:
            text = str(column)
            self.x_combo.addItem(text, column)
            self.y_combo.addItem(text, column)

        self.x_combo.blockSignals(False)
        self.y_combo.blockSignals(False)

        row_count, col_count = df.shape
        numeric_count = len(numeric_columns)
        self.summary_label.setText(
            f"当前表：{table['name']}  |  行数：{row_count}  |  列数：{col_count}  |  可绘图数值列：{numeric_count}"
        )

        if numeric_count == 0:
            self._set_controls_enabled(False)
            self.table_combo.setEnabled(True)
            self.chart_combo.setEnabled(True)
            self.refresh_button.setEnabled(True)
            self.status_label.setText("当前表格没有可用的数值列，无法绘制散点图、折线图或直方图。")
            self._render_placeholder("当前表格没有数值列")
            return

        self._set_controls_enabled(True)
        self._restore_or_default_selection(previous_x, previous_y, numeric_columns)
        self._sync_control_visibility()
        self.status_label.setText("已准备就绪：选择字段后点击“绘制图形”，或直接导出当前图。")

        if auto_plot:
            self.plot_current()

    def _restore_or_default_selection(self, previous_x, previous_y, numeric_columns):
        x_index = self.x_combo.findData(previous_x, flags=Qt.MatchExactly)
        if x_index >= 0:
            self.x_combo.setCurrentIndex(x_index)
        else:
            self.x_combo.setCurrentIndex(0 if self.chart_combo.currentText() == "折线图" else 1)

        y_index = self.y_combo.findData(previous_y, flags=Qt.MatchExactly)
        if y_index >= 0:
            self.y_combo.setCurrentIndex(y_index)
        else:
            self.y_combo.setCurrentIndex(1 if len(numeric_columns) > 1 else 0)

        self._apply_default_axis_selection()

    def _apply_default_axis_selection(self):
        chart_type = self.chart_combo.currentText()
        if self.x_combo.count() == 0:
            return

        if chart_type == "直方图":
            if self.x_combo.currentData() == self.INDEX_KEY and self.x_combo.count() > 1:
                self.x_combo.setCurrentIndex(1)
            return

        if chart_type == "散点图":
            if self.x_combo.currentData() == self.INDEX_KEY and self.x_combo.count() > 2:
                self.x_combo.setCurrentIndex(1)
            if self.y_combo.count() > 1 and self.y_combo.currentIndex() == self.x_combo.currentIndex() - 1:
                self.y_combo.setCurrentIndex(min(self.y_combo.count() - 1, max(0, self.x_combo.currentIndex())))

    def _sync_control_visibility(self):
        is_hist = self.chart_combo.currentText() == "直方图"
        self.y_row_label.setVisible(not is_hist)
        self.y_combo.setVisible(not is_hist)
        self.bin_label.setVisible(is_hist)
        self.bin_spin.setVisible(is_hist)
        self.x_row_label.setText("字段" if is_hist else "X 轴字段")

    def _set_controls_enabled(self, enabled: bool):
        self.x_combo.setEnabled(enabled)
        self.y_combo.setEnabled(enabled)
        self.bin_spin.setEnabled(enabled)
        self.sample_spin.setEnabled(enabled)
        self.plot_button.setEnabled(enabled)
        self.save_button.setEnabled(enabled)

    def get_selected_table(self):
        path = self.table_combo.currentData()
        if not path:
            return None
        if hasattr(self.data_module, "get_table_by_path"):
            return self.data_module.get_table_by_path(path)
        return None

    def plot_current(self):
        table = self.get_selected_table()
        if table is None:
            self.status_label.setText("提示：没有可用的数据表。")
            self._render_placeholder("请选择有效的数据表")
            return

        df = table["df"]
        chart_type = self.chart_combo.currentText()
        sample_limit = self.sample_spin.value()

        if chart_type == "直方图":
            self._plot_histogram(df, sample_limit)
        elif chart_type == "散点图":
            self._plot_scatter(df, sample_limit)
        else:
            self._plot_line(df, sample_limit)

    def _plot_histogram(self, df: pd.DataFrame, sample_limit: int):
        column_key = self.x_combo.currentData()
        series = self._extract_numeric_series(df, column_key)
        if series is None:
            self.status_label.setText("请选择一个数值列来绘制直方图。")
            self._render_placeholder("直方图需要一个数值列")
            return

        data = series.dropna()
        if data.empty:
            self.status_label.setText("所选列没有可用的数值数据。")
            self._render_placeholder("没有可绘制的数据")
            return

        max_points = min(max(sample_limit, 1000), self.MAX_HIST_POINTS)
        data, sampled = self._downsample_series(data, max_points)

        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.hist(data, bins=self.bin_spin.value(), color="#4C78A8", edgecolor="white", alpha=0.88)
        ax.set_title(f"直方图 - {self.x_combo.currentText()}")
        ax.set_xlabel(self.x_combo.currentText())
        ax.set_ylabel("频数")
        ax.grid(True, linestyle="--", alpha=0.3)
        self.canvas.draw_idle()

        sample_tip = "（已抽样显示）" if sampled else ""
        self.status_label.setText(
            f"直方图绘制完成：{self.x_combo.currentText()}，样本数 {len(data)} {sample_tip}".strip()
        )

    def _plot_scatter(self, df: pd.DataFrame, sample_limit: int):
        points, x_label, y_label, sampled = self._prepare_xy_data(df, sample_limit)
        if points is None:
            return

        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.scatter(points["x"], points["y"], s=20, alpha=0.72, color="#DD8452", edgecolors="none")
        ax.set_title(f"散点图 - {y_label} vs {x_label}")
        ax.set_xlabel(x_label)
        ax.set_ylabel(y_label)
        ax.grid(True, linestyle="--", alpha=0.3)
        self.canvas.draw_idle()

        sample_tip = "，已抽样显示" if sampled else ""
        self.status_label.setText(f"散点图绘制完成：有效点数 {len(points)}{sample_tip}")

    def _plot_line(self, df: pd.DataFrame, sample_limit: int):
        points, x_label, y_label, sampled = self._prepare_xy_data(df, sample_limit)
        if points is None:
            return

        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.plot(points["x"], points["y"], color="#55A868", linewidth=1.8, marker="o", markersize=3)
        ax.set_title(f"折线图 - {y_label} vs {x_label}")
        ax.set_xlabel(x_label)
        ax.set_ylabel(y_label)
        ax.grid(True, linestyle="--", alpha=0.3)
        self.canvas.draw_idle()

        sample_tip = "，已抽样显示" if sampled else ""
        self.status_label.setText(f"折线图绘制完成：有效点数 {len(points)}{sample_tip}")

    def _prepare_xy_data(self, df: pd.DataFrame, sample_limit: int):
        x_key = self.x_combo.currentData()
        y_key = self.y_combo.currentData()
        y_series = self._extract_numeric_series(df, y_key)

        if y_series is None:
            self.status_label.setText("请选择一个数值列作为 Y 轴。")
            self._render_placeholder("Y 轴需要数值列")
            return None, None, None, False

        if x_key == self.INDEX_KEY:
            x_series = pd.Series(np.arange(len(df)), index=df.index, dtype=float)
            x_label = self.INDEX_LABEL
        else:
            x_series = self._extract_numeric_series(df, x_key)
            if x_series is None:
                self.status_label.setText("请选择一个数值列作为 X 轴。")
                self._render_placeholder("X 轴需要数值列")
                return None, None, None, False
            x_label = self.x_combo.currentText()

        points = pd.DataFrame({"x": x_series, "y": y_series}).dropna()
        if points.empty:
            self.status_label.setText("所选字段没有可配对的有效数值数据。")
            self._render_placeholder("没有可绘制的有效点")
            return None, None, None, False

        points, sampled = self._downsample_frame(points, sample_limit)
        return points, x_label, self.y_combo.currentText(), sampled

    def _extract_numeric_series(self, df: pd.DataFrame, column_key):
        if column_key in (None, ""):
            return None
        if column_key == self.INDEX_KEY:
            return pd.Series(np.arange(len(df)), index=df.index, dtype=float)
        if column_key not in df.columns:
            return None
        return pd.to_numeric(df[column_key], errors="coerce")

    def _downsample_frame(self, frame: pd.DataFrame, limit: int):
        if limit <= 0 or len(frame) <= limit:
            return frame, False
        step = max(1, int(math.ceil(len(frame) / limit)))
        return frame.iloc[::step].head(limit), True

    def _downsample_series(self, series: pd.Series, limit: int):
        if limit <= 0 or len(series) <= limit:
            return series, False
        step = max(1, int(math.ceil(len(series) / limit)))
        return series.iloc[::step].head(limit), True

    def save_figure(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "导出绘图",
            "plot.png",
            "PNG 图片 (*.png);;JPEG 图片 (*.jpg *.jpeg);;SVG 图片 (*.svg)"
        )
        if not path:
            return
        self.figure.savefig(path, dpi=200, bbox_inches="tight")
        self.status_label.setText(f"图片已导出：{path}")

    def _render_placeholder(self, message: str):
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.axis("off")
        ax.text(0.5, 0.5, message, ha="center", va="center", fontsize=12, color="#666")
        self.canvas.draw_idle()

