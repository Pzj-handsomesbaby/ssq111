# -*- coding: utf-8 -*-
"""
Created on Mon Oct 20 11:17:18 2025

@author: AAA
"""

from __future__ import annotations

import math
import os

import numpy as np
import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QSplitter,
    QTableView,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from matplotlib import rcParams
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure

from modules.data_file_module import PandasTableModel
from modules.stand_quality_core import run_stand_quality_assessment, weight_method_display_name
from modules.structure_metrics_core import RunOptions, execute_pipeline
try:
    from modules.ui_style import apply_qt_window_baseline
except ModuleNotFoundError:
    from ui_style import apply_qt_window_baseline

rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "DejaVu Sans"]
rcParams["axes.unicode_minus"] = False


class QualityEvaluationWindow(QMainWindow):
    def __init__(self, mode_title: str, parent=None):
        super().__init__(parent)
        self.mode_title = mode_title
        self.setWindowTitle(mode_title)
        apply_qt_window_baseline(self, size=(1480, 900))
        self._build_ui()
        self._init_placeholder_content()

    def _build_ui(self):
        central = QWidget(self)
        self.setCentralWidget(central)
        root_layout = QHBoxLayout(central)

        splitter = QSplitter(Qt.Horizontal, central)
        root_layout.addWidget(splitter)

        left_panel = QWidget(splitter)
        left_layout = QVBoxLayout(left_panel)

        input_group = QGroupBox("输入 / 参数 / 导出", left_panel)
        input_form = QFormLayout(input_group)
        self.input_path_edit = QLineEdit(input_group)
        self.input_path_edit.setPlaceholderText("后续在这里接入输入数据或文件路径…")
        self.output_path_edit = QLineEdit(input_group)
        self.output_path_edit.setPlaceholderText("后续在这里接入输出目录或结果文件…")
        self.method_combo = QComboBox(input_group)
        self.method_combo.addItems(["默认模板", "预留方法 A", "预留方法 B"])
        input_form.addRow("输入数据", self.input_path_edit)
        input_form.addRow("输出位置", self.output_path_edit)
        input_form.addRow("评价方法", self.method_combo)

        option_group = QGroupBox("评价选项", left_panel)
        option_layout = QVBoxLayout(option_group)
        self.ck_option_a = QCheckBox("启用基础评价流程", option_group)
        self.ck_option_b = QCheckBox("启用结果汇总", option_group)
        self.ck_option_c = QCheckBox("启用图表输出", option_group)
        for checkbox in [self.ck_option_a, self.ck_option_b, self.ck_option_c]:
            checkbox.setChecked(True)
            option_layout.addWidget(checkbox)

        export_group = QGroupBox("导出设置", left_panel)
        export_layout = QGridLayout(export_group)
        self.ck_export_main = QCheckBox("导出主结果", export_group)
        self.ck_export_report = QCheckBox("导出评价报告", export_group)
        self.ck_export_chart = QCheckBox("导出图表", export_group)
        self.ck_export_log = QCheckBox("导出日志", export_group)
        for checkbox in [self.ck_export_main, self.ck_export_report, self.ck_export_chart, self.ck_export_log]:
            checkbox.setChecked(True)
        export_layout.addWidget(self.ck_export_main, 0, 0)
        export_layout.addWidget(self.ck_export_report, 0, 1)
        export_layout.addWidget(self.ck_export_chart, 1, 0)
        export_layout.addWidget(self.ck_export_log, 1, 1)

        action_group = QGroupBox("执行状态", left_panel)
        action_layout = QVBoxLayout(action_group)
        action_row = QWidget(action_group)
        action_row_layout = QHBoxLayout(action_row)
        action_row_layout.setContentsMargins(0, 0, 0, 0)
        self.run_button = QPushButton("开始评价", action_group)
        self.export_button = QPushButton("导出当前结果", action_group)
        action_row_layout.addWidget(self.run_button, stretch=1)
        action_row_layout.addWidget(self.export_button, stretch=1)
        self.progress_bar = QProgressBar(action_group)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.summary_label = QLabel("这是从“结构参数计算”模板衍生出的质量评价占位窗口，具体业务内容后续再补。", action_group)
        self.summary_label.setWordWrap(True)
        self.log_output = QTextEdit(action_group)
        self.log_output.setReadOnly(True)
        self.log_output.setMinimumHeight(260)
        action_layout.addWidget(action_row)
        action_layout.addWidget(self.progress_bar)
        action_layout.addWidget(self.summary_label)
        action_layout.addWidget(self.log_output, stretch=1)

        left_layout.addWidget(input_group)
        left_layout.addWidget(option_group)
        left_layout.addWidget(export_group)
        left_layout.addWidget(action_group, stretch=1)
        left_panel.setMinimumWidth(400)
        splitter.addWidget(left_panel)

        right_panel = QWidget(splitter)
        right_layout = QVBoxLayout(right_panel)
        self.tabs = QTabWidget(right_panel)
        right_layout.addWidget(self.tabs)

        self.preview_tab = self._build_placeholder_tab(
            title="结果预览",
            description=f"这里将放置 {self.mode_title} 的结果表、评分卡或关键指标。",
        )
        self.chart_tab = self._build_placeholder_tab(
            title="图表分析",
            description=f"这里将放置 {self.mode_title} 的图表、空间分布或趋势分析。",
        )
        self.report_tab = self._build_placeholder_tab(
            title="评价报告",
            description=f"这里将放置 {self.mode_title} 的报告摘要、说明和导出内容。",
        )
        self.log_tab = self._build_placeholder_tab(
            title="接口预留",
            description="这里预留后续接入模型参数、评价规则、接口结果和扩展配置。",
        )

        self.tabs.addTab(self.preview_tab, "结果预览")
        self.tabs.addTab(self.chart_tab, "图表分析")
        self.tabs.addTab(self.report_tab, "评价报告")
        self.tabs.addTab(self.log_tab, "接口预留")
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(1, 1)

    def _build_placeholder_tab(self, title: str, description: str) -> QWidget:
        tab = QWidget(self)
        layout = QVBoxLayout(tab)
        header = QLabel(title, tab)
        header.setStyleSheet("font-size: 18px; font-weight: 600;")
        desc = QLabel(description, tab)
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #666; line-height: 1.6;")
        placeholder = QTextEdit(tab)
        placeholder.setReadOnly(True)
        placeholder.setPlainText(
            "模板说明：\n"
            "1. 当前窗口仅复制结构参数计算的界面骨架。\n"
            "2. 后续你可以继续把真实的输入项、评价逻辑、图表和导出功能逐步填进来。\n"
            f"3. 当前占位窗口类型：{self.mode_title}。"
        )
        layout.addWidget(header)
        layout.addWidget(desc)
        layout.addWidget(placeholder, stretch=1)
        return tab

    def _init_placeholder_content(self):
        self.log_output.append(f"{self.mode_title}窗口已创建。")
        self.log_output.append("当前阶段仅提供模板框架，等待后续补充具体评价内容。")


class SingleTreeQualityWindow(QMainWindow):
    def __init__(self, data_module=None, parent=None):
        super().__init__(parent)
        self.data_module = data_module
        self.current_result = None
        self._last_metric_list: list[str] = []

        self.setWindowTitle("单木质量评价")
        apply_qt_window_baseline(self, size=(1480, 900))
        self._build_ui()
        self._bind_signals()
        self.refresh_tables()

        if self.data_module is not None and hasattr(self.data_module, "table_data_changed"):
            self.data_module.table_data_changed.connect(self.on_table_data_changed)

    def _build_ui(self):
        central = QWidget(self)
        self.setCentralWidget(central)
        root_layout = QHBoxLayout(central)

        splitter = QSplitter(Qt.Horizontal, central)
        root_layout.addWidget(splitter)

        left_panel = QWidget(splitter)
        left_layout = QVBoxLayout(left_panel)

        input_group = QGroupBox("输入 / 参数 / 导出", left_panel)
        input_form = QFormLayout(input_group)

        table_row = QWidget(input_group)
        table_row_layout = QHBoxLayout(table_row)
        table_row_layout.setContentsMargins(0, 0, 0, 0)
        self.table_combo = QComboBox(table_row)
        self.add_table_button = QPushButton("添加表格/CSV", table_row)
        table_row_layout.addWidget(self.table_combo, stretch=1)
        table_row_layout.addWidget(self.add_table_button)
        input_form.addRow("已加载表格", table_row)

        self.input_path_edit = QLineEdit(input_group)
        self.input_path_edit.setReadOnly(True)
        input_form.addRow("当前数据源", self.input_path_edit)

        out_dir_row = QWidget(input_group)
        out_dir_layout = QHBoxLayout(out_dir_row)
        out_dir_layout.setContentsMargins(0, 0, 0, 0)
        self.out_dir_edit = QLineEdit(out_dir_row)
        self.pick_out_dir_button = QPushButton("选择", out_dir_row)
        out_dir_layout.addWidget(self.out_dir_edit, stretch=1)
        out_dir_layout.addWidget(self.pick_out_dir_button)
        input_form.addRow("输出文件夹", out_dir_row)

        self.weight_method_combo = QComboBox(input_group)
        self.weight_method_combo.addItem("等权", "equal")
        self.weight_method_combo.addItem("熵权", "entropy")
        self.weight_method_combo.setCurrentIndex(1)
        input_form.addRow("权重算法", self.weight_method_combo)

        self.buffer_edit = QLineEdit(input_group)
        self.buffer_edit.setText("2")
        input_form.addRow("BufferSize（米）", self.buffer_edit)

        self.infl_k_edit = QLineEdit(input_group)
        self.infl_k_edit.setText("1.5")
        input_form.addRow("F 影响圈系数", self.infl_k_edit)

        self.format_combo = QComboBox(input_group)
        self.format_combo.addItems(["TXT（制表符）", "CSV（逗号）"])
        input_form.addRow("导出格式", self.format_combo)

        modules_group = QGroupBox("评价指标", left_panel)
        modules_layout = QVBoxLayout(modules_group)
        self.ck_wm = QCheckBox("W / M 指标", modules_group)
        self.ck_layer = QCheckBox("层级 / CI / Openness / PathOrder", modules_group)
        self.ck_uuci = QCheckBox("U + UCI(主=UCI_D) / UCI_H", modules_group)
        self.ck_cf = QCheckBox("C + F 指标", modules_group)
        self.ck_q = QCheckBox("单木综合质量指数 Q", modules_group)
        for checkbox in [self.ck_wm, self.ck_layer, self.ck_uuci, self.ck_cf, self.ck_q]:
            checkbox.setChecked(True)
            modules_layout.addWidget(checkbox)

        export_group = QGroupBox("导出文件", left_panel)
        export_layout = QGridLayout(export_group)
        self.ck_exp_main = QCheckBox("主结果", export_group)
        self.ck_exp_means = QCheckBox("综合均值", export_group)
        self.ck_exp_weights = QCheckBox("权重表", export_group)
        self.ck_exp_q = QCheckBox("综合指数表", export_group)
        for checkbox in [self.ck_exp_main, self.ck_exp_means, self.ck_exp_weights, self.ck_exp_q]:
            checkbox.setChecked(True)
        export_layout.addWidget(self.ck_exp_main, 0, 0)
        export_layout.addWidget(self.ck_exp_means, 0, 1)
        export_layout.addWidget(self.ck_exp_weights, 1, 0)
        export_layout.addWidget(self.ck_exp_q, 1, 1)

        action_group = QGroupBox("执行状态", left_panel)
        action_layout = QVBoxLayout(action_group)
        action_row = QWidget(action_group)
        action_row_layout = QHBoxLayout(action_row)
        action_row_layout.setContentsMargins(0, 0, 0, 0)
        self.run_button = QPushButton("开始评价", action_group)
        self.export_png_button = QPushButton("导出当前图为 PNG", action_group)
        action_row_layout.addWidget(self.run_button, stretch=1)
        action_row_layout.addWidget(self.export_png_button, stretch=1)
        self.progress_bar = QProgressBar(action_group)
        self.progress_bar.setRange(0, 100)
        self.summary_label = QLabel("请先选择已加载表格，或点击“添加表格/CSV”。", action_group)
        self.summary_label.setWordWrap(True)
        self.log_output = QTextEdit(action_group)
        self.log_output.setReadOnly(True)
        self.log_output.setMinimumHeight(260)
        action_layout.addWidget(action_row)
        action_layout.addWidget(self.progress_bar)
        action_layout.addWidget(self.summary_label)
        action_layout.addWidget(self.log_output, stretch=1)

        left_layout.addWidget(input_group)
        left_layout.addWidget(modules_group)
        left_layout.addWidget(export_group)
        left_layout.addWidget(action_group, stretch=1)
        left_panel.setMinimumWidth(400)
        splitter.addWidget(left_panel)

        right_panel = QWidget(splitter)
        right_layout = QVBoxLayout(right_panel)
        self.tabs = QTabWidget(right_panel)
        right_layout.addWidget(self.tabs)

        self.preview_table = QTableView(self.tabs)
        self.means_table = QTableView(self.tabs)
        self.weights_table = QTableView(self.tabs)
        self.q_table_view = QTableView(self.tabs)
        for view in [self.preview_table, self.means_table, self.weights_table, self.q_table_view]:
            view.setSortingEnabled(True)
            view.setAlternatingRowColors(True)

        self.spatial_tab = QWidget(self.tabs)
        spatial_layout = QVBoxLayout(self.spatial_tab)
        spatial_controls = QWidget(self.spatial_tab)
        spatial_controls_layout = QHBoxLayout(spatial_controls)
        spatial_controls_layout.setContentsMargins(0, 0, 0, 0)
        self.spatial_metric_combo = QComboBox(spatial_controls)
        self.spatial_group_combo = QComboBox(spatial_controls)
        self.spatial_group_combo.addItems(["ALL", "CORE", "BUFFER"])
        self.refresh_spatial_button = QPushButton("刷新空间图", spatial_controls)
        spatial_controls_layout.addWidget(QLabel("指标", spatial_controls))
        spatial_controls_layout.addWidget(self.spatial_metric_combo, stretch=1)
        spatial_controls_layout.addWidget(QLabel("分组", spatial_controls))
        spatial_controls_layout.addWidget(self.spatial_group_combo)
        spatial_controls_layout.addWidget(self.refresh_spatial_button)
        spatial_layout.addWidget(spatial_controls)
        self.spatial_figure = Figure(constrained_layout=True)
        self.spatial_canvas = FigureCanvas(self.spatial_figure)
        self.spatial_toolbar = NavigationToolbar(self.spatial_canvas, self.spatial_tab)
        spatial_layout.addWidget(self.spatial_toolbar)
        spatial_layout.addWidget(self.spatial_canvas, stretch=1)

        self.dist_tab = QWidget(self.tabs)
        dist_layout = QVBoxLayout(self.dist_tab)
        dist_controls = QWidget(self.dist_tab)
        dist_controls_layout = QHBoxLayout(dist_controls)
        dist_controls_layout.setContentsMargins(0, 0, 0, 0)
        self.dist_metric_combo = QComboBox(dist_controls)
        self.dist_group_combo = QComboBox(dist_controls)
        self.dist_group_combo.addItems(["ALL", "CORE", "BUFFER"])
        self.refresh_dist_button = QPushButton("刷新分布图", dist_controls)
        dist_controls_layout.addWidget(QLabel("指标", dist_controls))
        dist_controls_layout.addWidget(self.dist_metric_combo, stretch=1)
        dist_controls_layout.addWidget(QLabel("分组", dist_controls))
        dist_controls_layout.addWidget(self.dist_group_combo)
        dist_controls_layout.addWidget(self.refresh_dist_button)
        dist_layout.addWidget(dist_controls)
        self.dist_figure = Figure(constrained_layout=True)
        self.dist_canvas = FigureCanvas(self.dist_figure)
        self.dist_toolbar = NavigationToolbar(self.dist_canvas, self.dist_tab)
        dist_layout.addWidget(self.dist_toolbar)
        dist_layout.addWidget(self.dist_canvas, stretch=1)

        self.tabs.addTab(self.preview_table, "主结果预览")
        self.tabs.addTab(self.spatial_tab, "空间分布")
        self.tabs.addTab(self.dist_tab, "指标分布")
        self.tabs.addTab(self.means_table, "综合均值")
        self.tabs.addTab(self.weights_table, "权重表")
        self.tabs.addTab(self.q_table_view, "综合指数表")
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(1, 1)

        self._render_placeholder(self.spatial_figure, self.spatial_canvas, "完成评价后将在此显示空间分布图。")
        self._render_placeholder(self.dist_figure, self.dist_canvas, "完成评价后将在此显示指标分布图。")

    def _bind_signals(self):
        self.add_table_button.clicked.connect(self.add_table_file)
        self.pick_out_dir_button.clicked.connect(self.pick_out_dir)
        self.table_combo.currentIndexChanged.connect(self.on_table_changed)
        self.run_button.clicked.connect(self.run_quality_evaluation)
        self.export_png_button.clicked.connect(self.export_current_plot_png)
        self.refresh_spatial_button.clicked.connect(self.refresh_spatial_plot)
        self.refresh_dist_button.clicked.connect(self.refresh_distribution_plot)
        self.spatial_metric_combo.currentTextChanged.connect(lambda _: self.refresh_spatial_plot())
        self.spatial_group_combo.currentTextChanged.connect(lambda _: self.refresh_spatial_plot())
        self.dist_metric_combo.currentTextChanged.connect(lambda _: self.refresh_distribution_plot())
        self.dist_group_combo.currentTextChanged.connect(lambda _: self.refresh_distribution_plot())

    def on_table_data_changed(self):
        self.refresh_tables(auto_select_latest=True)

    def add_table_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择表格或 CSV 文件", "", "表格文件 (*.csv *.xlsx *.xls)")
        if not path:
            self.append_log("已取消导入表格。")
            return
        if self.data_module is None or not hasattr(self.data_module, "load_table"):
            QMessageBox.warning(self, "不可用", "当前主程序未提供表格加载模块。")
            return
        ok = self.data_module.load_table(path)
        if ok:
            self.refresh_tables(auto_select_latest=True)
            self.append_log(f"已导入表格：{os.path.basename(path)}")
        else:
            error_msg = self.data_module.get_last_table_error() if hasattr(self.data_module, "get_last_table_error") else "表格导入失败。"
            QMessageBox.warning(self, "导入失败", error_msg)
            self.append_log(error_msg)

    def pick_out_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出文件夹", "")
        if path:
            self.out_dir_edit.setText(path)
            self.append_log(f"输出目录：{path}")

    def refresh_tables(self, auto_select_latest: bool = False):
        tables = self.data_module.get_loaded_tables() if self.data_module is not None and hasattr(self.data_module, "get_loaded_tables") else []
        previous_path = self.table_combo.currentData()

        self.table_combo.blockSignals(True)
        self.table_combo.clear()
        for table in tables:
            self.table_combo.addItem(table["name"], table["path"])
        self.table_combo.blockSignals(False)

        if not tables:
            self.input_path_edit.clear()
            self.summary_label.setText("当前没有已加载的表格数据，请先导入 CSV / Excel。")
            return

        target_path = None
        if auto_select_latest and hasattr(self.data_module, "get_latest_table"):
            latest = self.data_module.get_latest_table()
            target_path = latest["path"] if latest else None
        target_path = target_path or previous_path

        index = self.table_combo.findData(target_path, flags=Qt.MatchExactly) if target_path else -1
        if index < 0:
            index = max(0, self.table_combo.count() - 1)
        self.table_combo.setCurrentIndex(index)
        self.on_table_changed()

    def on_table_changed(self):
        table = self.get_selected_table()
        if table is None:
            self.input_path_edit.clear()
            self.summary_label.setText("请选择可用数据表。")
            return
        self.input_path_edit.setText(table["path"])
        rows, cols = table["df"].shape
        self.summary_label.setText(f"当前表：{table['name']} | 行数：{rows} | 列数：{cols} | 运行后将在右侧显示单木质量评价结果。")

    def get_selected_table(self):
        path = self.table_combo.currentData()
        if not path:
            return None
        if self.data_module is not None and hasattr(self.data_module, "get_table_by_path"):
            return self.data_module.get_table_by_path(path)
        return None

    def run_quality_evaluation(self):
        table = self.get_selected_table()
        if table is None:
            QMessageBox.warning(self, "缺少数据", "请先选择已加载表格，或点击“添加表格/CSV”。")
            return

        weight_method = self.weight_method_combo.currentData()
        options = RunOptions(
            use_wm=self.ck_wm.isChecked(),
            use_layer=self.ck_layer.isChecked(),
            use_uuci=self.ck_uuci.isChecked(),
            use_cf=self.ck_cf.isChecked(),
            use_q=self.ck_q.isChecked(),
            export_main=self.ck_exp_main.isChecked(),
            export_means=self.ck_exp_means.isChecked(),
            export_entropy=self.ck_exp_weights.isChecked(),
            export_q_table=self.ck_exp_q.isChecked(),
            buffer_size=self._parse_float(self.buffer_edit.text(), 2.0),
            infl_k=self._parse_float(self.infl_k_edit.text(), 1.5),
            weight_method=weight_method,
            export_weight_label="权重",
        )
        export_format = "CSV" if self.format_combo.currentIndex() == 1 else "TXT"
        weight_label = weight_method_display_name(weight_method)

        self.run_button.setEnabled(False)
        self.progress_bar.setValue(0)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.append_log(f"---- 单木质量评价开始（{weight_label}） ----")
        try:
            self.current_result = execute_pipeline(
                source_df=table["df"],
                input_path=table["path"],
                out_dir=self.out_dir_edit.text().strip(),
                export_format=export_format,
                options=options,
                progress_callback=self.on_progress,
            )
            self.set_table_dataframe(self.preview_table, self.current_result.main_preview, max_rows=600)
            self.set_table_dataframe(self.means_table, self.current_result.means, max_rows=600)
            self.set_table_dataframe(self.weights_table, self.current_result.entropy, max_rows=200)
            self.set_table_dataframe(self.q_table_view, self.current_result.q_table, max_rows=800)
            self.populate_metric_controls(self.current_result.metric_list)
            self.refresh_spatial_plot()
            self.refresh_distribution_plot()
            self.log_export_paths(self.current_result.export_paths)
            self.append_log(f"结果记录数：{len(self.current_result.computed_data)}")
            self.append_log(f"---- 单木质量评价完成（{weight_label}） ----")
            self.summary_label.setText(
                f"评价完成：标准记录 {len(self.current_result.standard_data)} 条，权重算法：{weight_label}。"
            )
        except Exception as exc:
            self.append_log(f"评价失败：{exc}")
            QMessageBox.critical(self, "单木质量评价失败", str(exc))
        finally:
            self.run_button.setEnabled(True)
            QApplication.restoreOverrideCursor()

    def on_progress(self, value: int, message: str):
        self.progress_bar.setValue(int(value))
        if message:
            self.append_log(message)
        QApplication.processEvents()

    def populate_metric_controls(self, metric_list: list[str]):
        if metric_list == self._last_metric_list:
            return
        self._last_metric_list = list(metric_list)
        for combo in [self.spatial_metric_combo, self.dist_metric_combo]:
            current = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(metric_list)
            combo.blockSignals(False)
            if current:
                index = combo.findText(current, flags=Qt.MatchExactly)
                if index >= 0:
                    combo.setCurrentIndex(index)
        if self.spatial_metric_combo.count() > 0 and self.spatial_metric_combo.currentIndex() < 0:
            self.spatial_metric_combo.setCurrentIndex(0)
        if self.dist_metric_combo.count() > 0 and self.dist_metric_combo.currentIndex() < 0:
            self.dist_metric_combo.setCurrentIndex(0)

    def set_table_dataframe(self, view: QTableView, df: pd.DataFrame, max_rows: int):
        clipped = df.head(max_rows).copy() if df is not None else pd.DataFrame()
        if not clipped.empty:
            for column in clipped.columns:
                if pd.api.types.is_numeric_dtype(clipped[column]):
                    clipped[column] = clipped[column].map(self.format_cell)
        model = PandasTableModel(clipped)
        view.setModel(model)
        view.resizeColumnsToContents()

    def refresh_spatial_plot(self):
        if self.current_result is None:
            return
        metric = self.spatial_metric_combo.currentText()
        filtered = self.filter_by_group(self.current_result.computed_data, self.spatial_group_combo.currentText())
        if filtered.empty or metric not in filtered.columns:
            self._render_placeholder(self.spatial_figure, self.spatial_canvas, "当前分组下没有可绘制的空间分布数据。")
            return
        plot_df = filtered[["X", "Y", metric]].copy()
        plot_df[metric] = pd.to_numeric(plot_df[metric], errors="coerce")
        plot_df = plot_df[np.isfinite(plot_df[metric])]
        if plot_df.empty:
            self._render_placeholder(self.spatial_figure, self.spatial_canvas, f"指标 {metric} 没有可用数值。")
            return
        self.spatial_figure.clear()
        ax = self.spatial_figure.add_subplot(111)
        sc = ax.scatter(plot_df["X"], plot_df["Y"], c=plot_df[metric], cmap="viridis", s=28, edgecolors="none")
        ax.set_title(f"{metric} 空间分布 ({self.spatial_group_combo.currentText()})")
        ax.set_xlabel("X")
        ax.set_ylabel("Y")
        ax.grid(True, linestyle="--", linewidth=0.5, alpha=0.5)
        self.spatial_figure.colorbar(sc, ax=ax, shrink=0.9, label=metric)
        self.spatial_canvas.draw_idle()

    def refresh_distribution_plot(self):
        if self.current_result is None:
            return
        metric = self.dist_metric_combo.currentText()
        filtered = self.filter_by_group(self.current_result.computed_data, self.dist_group_combo.currentText())
        if filtered.empty or metric not in filtered.columns:
            self._render_placeholder(self.dist_figure, self.dist_canvas, "当前分组下没有可绘制的分布数据。")
            return
        values = pd.to_numeric(filtered[metric], errors="coerce")
        values = values[np.isfinite(values)]
        if values.empty:
            self._render_placeholder(self.dist_figure, self.dist_canvas, f"指标 {metric} 没有可用数值。")
            return
        self.dist_figure.clear()
        ax = self.dist_figure.add_subplot(111)
        arr = values.to_numpy(dtype=float)
        is_quarter = np.all((arr >= -1e-9) & (arr <= 1.000001)) and np.allclose(arr * 4.0, np.round(arr * 4.0), atol=1e-8)
        if is_quarter:
            bins = np.array([0.0, 0.25, 0.5, 0.75, 1.0])
            counts = [int(np.sum(np.isclose(np.round(arr * 4.0) / 4.0, val))) for val in bins]
            ax.bar([str(v) for v in bins], counts, color="#5B8FF9", edgecolor="#27408b")
        else:
            bin_count = min(20, max(6, int(math.sqrt(len(arr)))))
            ax.hist(arr, bins=bin_count, color="#5B8FF9", edgecolor="white", alpha=0.9)
        ax.set_title(f"{metric} 指标分布 ({self.dist_group_combo.currentText()}) | n={len(arr)}")
        ax.set_xlabel(metric)
        ax.set_ylabel("频数")
        ax.grid(True, axis="y", linestyle="--", linewidth=0.5, alpha=0.5)
        self.dist_canvas.draw_idle()

    def filter_by_group(self, df: pd.DataFrame, group: str) -> pd.DataFrame:
        if group == "CORE" and "IsCore" in df.columns:
            return df[df["IsCore"].astype(bool)].copy()
        if group == "BUFFER" and "IsCore" in df.columns:
            return df[~df["IsCore"].astype(bool)].copy()
        return df.copy()

    def export_current_plot_png(self):
        tab_name, figure = self.get_current_plot_context()
        if figure is None:
            QMessageBox.information(self, "无法导出", "当前页没有可导出的图，请切换到“空间分布”或“指标分布”页。")
            return
        if self.current_result is None:
            QMessageBox.information(self, "暂无图形", "请先完成一次单木质量评价，再导出当前图。")
            return
        default_dir = self._suggest_export_dir()
        default_name = f"{tab_name}_{self._current_metric_name(tab_name)}.png"
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "导出当前图为 PNG",
            os.path.join(default_dir, default_name),
            "PNG 图片 (*.png)",
        )
        if not save_path:
            self.append_log("已取消导出 PNG。")
            return
        try:
            save_path = self.save_current_plot_png_to_path(save_path)
            self.append_log(f"当前图已导出为 PNG：{save_path}")
            self.summary_label.setText(f"PNG 导出完成：{os.path.basename(save_path)}")
        except Exception as exc:
            self.append_log(f"导出 PNG 失败：{exc}")
            QMessageBox.critical(self, "导出失败", str(exc))

    def save_current_plot_png_to_path(self, save_path: str) -> str:
        _, figure = self.get_current_plot_context()
        if figure is None:
            raise ValueError("当前页没有可导出的图。")
        if self.current_result is None:
            raise ValueError("请先完成一次单木质量评价，再导出当前图。")
        if not save_path.lower().endswith(".png"):
            save_path += ".png"
        figure.savefig(save_path, dpi=200, bbox_inches="tight", facecolor="white")
        return save_path

    def get_current_plot_context(self):
        current_widget = self.tabs.currentWidget()
        if current_widget is self.spatial_tab:
            return "空间分布", self.spatial_figure
        if current_widget is self.dist_tab:
            return "指标分布", self.dist_figure
        return "", None

    def _suggest_export_dir(self) -> str:
        if self.out_dir_edit.text().strip():
            return self.out_dir_edit.text().strip()
        if self.input_path_edit.text().strip():
            return os.path.dirname(self.input_path_edit.text().strip()) or os.getcwd()
        return os.getcwd()

    def _current_metric_name(self, tab_name: str) -> str:
        if tab_name == "空间分布":
            metric_name = self.spatial_metric_combo.currentText().strip()
        elif tab_name == "指标分布":
            metric_name = self.dist_metric_combo.currentText().strip()
        else:
            metric_name = "plot"
        return metric_name or "plot"

    def log_export_paths(self, export_paths):
        for label, path in [
            ("主结果", export_paths.main),
            ("综合均值", export_paths.means),
            ("权重表", export_paths.entropy),
            ("综合指数表", export_paths.q_tree),
        ]:
            if path:
                self.append_log(f"已导出{label}：{path}")

    def append_log(self, message: str):
        self.log_output.append(message)
        parent = self.parent()
        if parent is not None and hasattr(parent, "log_output"):
            parent.log_output.append(f"[单木质量评价] {message}")

    def _render_placeholder(self, figure: Figure, canvas: FigureCanvas, message: str):
        figure.clear()
        ax = figure.add_subplot(111)
        ax.text(0.5, 0.5, message, ha="center", va="center", fontsize=12, color="#666")
        ax.set_axis_off()
        canvas.draw_idle()

    @staticmethod
    def _parse_float(text: str, fallback: float) -> float:
        text = str(text or "").strip()
        for candidate in (text, text.replace(",", ""), text.replace(",", ".")):
            try:
                return float(candidate)
            except ValueError:
                continue
        return fallback

    @staticmethod
    def format_cell(value):
        if value is None or (isinstance(value, float) and not math.isfinite(value)):
            return ""
        if isinstance(value, (float, np.floating)):
            return f"{float(value):.6f}".rstrip("0").rstrip(".")
        return value


class EvaluationModule(QualityEvaluationWindow):
    def __init__(self, mode="单木质量评价", parent=None):
        super().__init__(mode_title=mode, parent=parent)


class StandQualityWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_result = None
        self.setWindowTitle("林分质量评价")
        apply_qt_window_baseline(self, size=(1420, 860))
        self._build_ui()
        self._bind_signals()

    def _build_ui(self):
        central = QWidget(self)
        self.setCentralWidget(central)
        root_layout = QHBoxLayout(central)

        splitter = QSplitter(Qt.Horizontal, central)
        root_layout.addWidget(splitter)

        left_panel = QWidget(splitter)
        left_layout = QVBoxLayout(left_panel)

        input_group = QGroupBox("输入 / 参数 / 导出", left_panel)
        input_form = QFormLayout(input_group)

        input_dir_row = QWidget(input_group)
        input_dir_layout = QHBoxLayout(input_dir_row)
        input_dir_layout.setContentsMargins(0, 0, 0, 0)
        self.input_dir_edit = QLineEdit(input_dir_row)
        self.pick_input_dir_button = QPushButton("选择目录", input_dir_row)
        input_dir_layout.addWidget(self.input_dir_edit, stretch=1)
        input_dir_layout.addWidget(self.pick_input_dir_button)
        input_form.addRow("输入目录", input_dir_row)

        output_dir_row = QWidget(input_group)
        output_dir_layout = QHBoxLayout(output_dir_row)
        output_dir_layout.setContentsMargins(0, 0, 0, 0)
        self.output_dir_edit = QLineEdit(output_dir_row)
        self.pick_output_dir_button = QPushButton("选择目录", output_dir_row)
        output_dir_layout.addWidget(self.output_dir_edit, stretch=1)
        output_dir_layout.addWidget(self.pick_output_dir_button)
        input_form.addRow("输出目录", output_dir_row)

        self.weight_method_combo = QComboBox(input_group)
        self.weight_method_combo.addItem("等权", "equal")
        self.weight_method_combo.addItem("熵权", "entropy")
        self.weight_method_combo.setCurrentIndex(1)
        input_form.addRow("权重算法", self.weight_method_combo)

        self.format_hint_label = QLabel("输入目录需要包含：树高 tif、胸径 tif、年龄 tif，以及 1 个边界 shp。\n系统会按文件名自动识别参数角色，例如 height / dbh / age / boundary / 树高 / 胸径 / 林龄 / 样地。", input_group)
        self.format_hint_label.setWordWrap(True)
        self.format_hint_label.setStyleSheet("color: #666;")
        input_form.addRow("命名说明", self.format_hint_label)

        option_group = QGroupBox("评价设置", left_panel)
        option_layout = QVBoxLayout(option_group)
        self.ck_export_weights = QCheckBox("导出权重表 CSV", option_group)
        self.ck_export_weights.setChecked(True)
        self.ck_growth_hint = QCheckBox("启用生长率参与评价（树高 / 林龄）", option_group)
        self.ck_growth_hint.setChecked(True)
        self.ck_png_hint = QCheckBox("允许导出当前图为 PNG", option_group)
        self.ck_png_hint.setChecked(True)
        for checkbox in [self.ck_export_weights, self.ck_growth_hint, self.ck_png_hint]:
            checkbox.setEnabled(False if checkbox is self.ck_growth_hint else True)
            option_layout.addWidget(checkbox)

        action_group = QGroupBox("执行状态", left_panel)
        action_layout = QVBoxLayout(action_group)
        action_row = QWidget(action_group)
        action_row_layout = QHBoxLayout(action_row)
        action_row_layout.setContentsMargins(0, 0, 0, 0)
        self.run_button = QPushButton("开始评价", action_group)
        self.export_png_button = QPushButton("导出当前图为 PNG", action_group)
        action_row_layout.addWidget(self.run_button, stretch=1)
        action_row_layout.addWidget(self.export_png_button, stretch=1)
        self.progress_bar = QProgressBar(action_group)
        self.progress_bar.setRange(0, 100)
        self.summary_label = QLabel("请先选择包含 tif / shp 的输入目录。", action_group)
        self.summary_label.setWordWrap(True)
        self.log_output = QTextEdit(action_group)
        self.log_output.setReadOnly(True)
        self.log_output.setMinimumHeight(280)
        action_layout.addWidget(action_row)
        action_layout.addWidget(self.progress_bar)
        action_layout.addWidget(self.summary_label)
        action_layout.addWidget(self.log_output, stretch=1)

        left_layout.addWidget(input_group)
        left_layout.addWidget(option_group)
        left_layout.addWidget(action_group, stretch=1)
        left_panel.setMinimumWidth(420)
        splitter.addWidget(left_panel)

        right_panel = QWidget(splitter)
        right_layout = QVBoxLayout(right_panel)
        self.tabs = QTabWidget(right_panel)
        right_layout.addWidget(self.tabs)

        self.raster_tab = QWidget(self.tabs)
        raster_layout = QVBoxLayout(self.raster_tab)
        self.raster_figure = Figure(constrained_layout=True)
        self.raster_canvas = FigureCanvas(self.raster_figure)
        self.raster_toolbar = NavigationToolbar(self.raster_canvas, self.raster_tab)
        raster_layout.addWidget(self.raster_toolbar)
        raster_layout.addWidget(self.raster_canvas, stretch=1)

        self.hist_tab = QWidget(self.tabs)
        hist_layout = QVBoxLayout(self.hist_tab)
        self.hist_figure = Figure(constrained_layout=True)
        self.hist_canvas = FigureCanvas(self.hist_figure)
        self.hist_toolbar = NavigationToolbar(self.hist_canvas, self.hist_tab)
        hist_layout.addWidget(self.hist_toolbar)
        hist_layout.addWidget(self.hist_canvas, stretch=1)

        self.tabs.addTab(self.raster_tab, "结果栅格")
        self.tabs.addTab(self.hist_tab, "直方图")
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(1, 1)

        self._render_placeholder(self.raster_figure, self.raster_canvas, "完成林分质量评价后将在此显示质量栅格。")
        self._render_placeholder(self.hist_figure, self.hist_canvas, "完成林分质量评价后将在此显示质量指数直方图。")

    def _bind_signals(self):
        self.pick_input_dir_button.clicked.connect(self.pick_input_dir)
        self.pick_output_dir_button.clicked.connect(self.pick_output_dir)
        self.run_button.clicked.connect(self.run_quality_evaluation)
        self.export_png_button.clicked.connect(self.export_current_plot_png)

    def pick_input_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择输入目录", "")
        if path:
            self.input_dir_edit.setText(path)
            self.append_log(f"输入目录：{path}")

    def pick_output_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出目录", "")
        if path:
            self.output_dir_edit.setText(path)
            self.append_log(f"输出目录：{path}")

    def run_quality_evaluation(self):
        input_dir = self.input_dir_edit.text().strip()
        if not input_dir:
            QMessageBox.warning(self, "缺少输入", "请先选择包含 tif / shp 的输入目录。")
            return

        output_dir = self.output_dir_edit.text().strip() or input_dir
        weight_method = self.weight_method_combo.currentData()
        weight_label = weight_method_display_name(weight_method)

        self.run_button.setEnabled(False)
        self.progress_bar.setValue(0)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.append_log(f"---- 林分质量评价开始（{weight_label}） ----")
        try:
            self.current_result = run_stand_quality_assessment(
                input_dir=input_dir,
                out_dir=output_dir,
                weight_method=weight_method,
                export_weights=self.ck_export_weights.isChecked(),
                progress_callback=self.on_progress,
            )
            self.render_raster_result()
            self.render_histogram_result()
            self.append_log(f"树高栅格：{self.current_result.input_paths.height_path}")
            self.append_log(f"胸径栅格：{self.current_result.input_paths.dbh_path}")
            self.append_log(f"年龄栅格：{self.current_result.input_paths.age_path}")
            self.append_log(f"边界文件：{self.current_result.input_paths.boundary_path}")
            if self.current_result.input_paths.crown_path:
                self.append_log(f"冠幅栅格：{self.current_result.input_paths.crown_path}")
            self.append_log(f"结果栅格已输出：{self.current_result.output_raster_path}")
            if self.current_result.output_weights_path:
                self.append_log(f"权重表已输出：{self.current_result.output_weights_path}")
            self.append_log(f"平均生长率：{self.current_result.growth_rate_mean:.4f}")
            self.append_log(f"有效像元数：{self.current_result.valid_pixel_count}")
            self.append_log(f"---- 林分质量评价完成（{weight_label}） ----")
            self.summary_label.setText(
                f"林分质量评价完成：算法={weight_label}，结果栅格={os.path.basename(self.current_result.output_raster_path)}。"
            )
        except Exception as exc:
            self.append_log(f"林分质量评价失败：{exc}")
            QMessageBox.critical(self, "林分质量评价失败", str(exc))
        finally:
            self.run_button.setEnabled(True)
            QApplication.restoreOverrideCursor()

    def on_progress(self, value: int, message: str):
        self.progress_bar.setValue(int(value))
        if message:
            self.append_log(message)
        QApplication.processEvents()

    def render_raster_result(self):
        if self.current_result is None:
            return
        self.raster_figure.clear()
        ax = self.raster_figure.add_subplot(111)
        image = ax.imshow(self.current_result.preview_array, cmap="YlGn", interpolation="nearest")
        ax.set_title(f"林分质量栅格（{weight_method_display_name(self.current_result.weight_method)}）")
        ax.set_axis_off()
        self.raster_figure.colorbar(image, ax=ax, shrink=0.85, label="质量指数")
        self.raster_canvas.draw_idle()

    def render_histogram_result(self):
        if self.current_result is None:
            return
        values = self.current_result.histogram_values
        self.hist_figure.clear()
        ax = self.hist_figure.add_subplot(111)
        bin_count = min(30, max(8, int(math.sqrt(len(values)))))
        ax.hist(values, bins=bin_count, color="#5B8FF9", edgecolor="white", alpha=0.9)
        ax.set_title(f"林分质量指数直方图（n={len(values)}）")
        ax.set_xlabel("质量指数")
        ax.set_ylabel("频数")
        ax.grid(True, axis="y", linestyle="--", linewidth=0.5, alpha=0.5)
        self.hist_canvas.draw_idle()

    def export_current_plot_png(self):
        tab_name, figure = self.get_current_plot_context()
        if figure is None:
            QMessageBox.information(self, "无法导出", "当前页没有可导出的图。")
            return
        if self.current_result is None:
            QMessageBox.information(self, "暂无图形", "请先完成一次林分质量评价，再导出当前图。")
            return
        default_dir = self.output_dir_edit.text().strip() or self.input_dir_edit.text().strip() or os.getcwd()
        default_name = f"{tab_name}_{weight_method_display_name(self.current_result.weight_method)}.png"
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "导出当前图为 PNG",
            os.path.join(default_dir, default_name),
            "PNG 图片 (*.png)",
        )
        if not save_path:
            self.append_log("已取消导出 PNG。")
            return
        try:
            save_path = self.save_current_plot_png_to_path(save_path)
            self.append_log(f"当前图已导出为 PNG：{save_path}")
        except Exception as exc:
            self.append_log(f"导出 PNG 失败：{exc}")
            QMessageBox.critical(self, "导出失败", str(exc))

    def save_current_plot_png_to_path(self, save_path: str) -> str:
        _, figure = self.get_current_plot_context()
        if figure is None:
            raise ValueError("当前页没有可导出的图。")
        if self.current_result is None:
            raise ValueError("请先完成一次林分质量评价，再导出当前图。")
        if not save_path.lower().endswith(".png"):
            save_path += ".png"
        figure.savefig(save_path, dpi=200, bbox_inches="tight", facecolor="white")
        return save_path

    def get_current_plot_context(self):
        current_widget = self.tabs.currentWidget()
        if current_widget is self.raster_tab:
            return "结果栅格", self.raster_figure
        if current_widget is self.hist_tab:
            return "直方图", self.hist_figure
        return "", None

    def append_log(self, message: str):
        self.log_output.append(message)
        parent = self.parent()
        if parent is not None and hasattr(parent, "log_output"):
            parent.log_output.append(f"[林分质量评价] {message}")

    def _render_placeholder(self, figure: Figure, canvas: FigureCanvas, message: str):
        figure.clear()
        ax = figure.add_subplot(111)
        ax.text(0.5, 0.5, message, ha="center", va="center", fontsize=12, color="#666")
        ax.set_axis_off()
        canvas.draw_idle()

