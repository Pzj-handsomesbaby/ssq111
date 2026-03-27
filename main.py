# -*- coding: utf-8 -*-

import os
import sys

# ==============================================================================
# 🔥 深度环境修复：针对 WinError 127 强化路径搜索
# ==============================================================================
if os.name == 'nt':
    env_base = os.path.dirname(sys.executable)
    # 显式定义需要检查的三个核心 DLL 路径
    dll_paths = [
        os.path.join(env_base, 'Library', 'bin'),
        os.path.join(env_base, 'Scripts'),
        os.path.join(env_base, 'Lib', 'site-packages', 'torch', 'lib'),
    ]

    for path in dll_paths:
        if os.path.exists(path):
            # 将路径加入系统 PATH 环境变量
            os.environ['PATH'] = path + os.pathsep + os.environ.get('PATH', '')
            # 针对 Python 3.8+ 强制添加 DLL 搜索目录
            if hasattr(os, 'add_dll_directory'):
                try:
                    os.add_dll_directory(path)
                except Exception:
                    pass
# ==============================================================================
import matplotlib

matplotlib.use("Qt5Agg")

import subprocess
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTreeWidget, QTreeWidgetItem, QAction, QTextEdit, QSplitter, QMessageBox, QLabel
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap

from modules.data_file_module import DataFileModule
from modules.data_plot_window import DataPlotWindow
from modules.optical_module import OpticalModule
from modules.forest_inversion import ForestInversionApp
from modules.structure_module import StructureModule
# from modules.evaluation_module import QualityEvaluationWindow, SingleTreeQualityWindow
from modules.stand_quality_qi_module import StandQualityWindow
# from modules.structure_metrics_module import StructureMetricsWindow
from modules.ui_style import apply_qt_app_style, apply_qt_window_baseline, QT_MAIN_SIZE

try:
    from modules.individual_tree_quality_module import StructureMetricsWindow as IndividualTreeQualityWindow
except Exception:
    IndividualTreeQualityWindow = None

# 兼容可能从 main 导入 TreeQualityApp 的旧调用。
TreeQualityApp = IndividualTreeQualityWindow


class _StderrWarningFilter:
    """Drop known non-fatal libpng metadata warnings from console output."""

    def __init__(self, wrapped):
        self._wrapped = wrapped

    def write(self, text):
        if "libpng warning: tRNS: invalid with alpha channel" in text:
            return
        self._wrapped.write(text)

    def flush(self):
        self._wrapped.flush()


def _install_console_warning_filter():
    if os.environ.get("FOREST_KEEP_LIBPNG_WARNING", "0") == "1":
        return
    sys.stderr = _StderrWarningFilter(sys.stderr)


# ---------------- 轻量模块管理器（让调度更清晰） ----------------
class ModuleManager:
    def __init__(self, host_window: "ForestMain"):
        self.host = host_window

    def show(self, widget: QWidget, log: str = ""):
        self.host.show_module(widget)
        if log:
            self.host.log_output.append(log)


class ForestMain(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("森林质量评价与智能决策平台 V1.0")
        apply_qt_window_baseline(self, size=QT_MAIN_SIZE)
        self.plot_window = None
        self.structure_metrics_window = None
        self.inversion_window = None
        self.optical_window = None
        self.functional_window = None
        self.gap_detection_window = None
        self.single_tree_structure_window = None
        self.descriptive_stats_process = None
        self.drive_structure_dynamics_process = None
        self.data_statistics_process = None
        self.quality_windows = {}
        self.trend_analysis_window = None

        # ---------------- 中心主 Widget ----------------
        central = QWidget()
        self.setCentralWidget(central)
        main_layout_vertical = QVBoxLayout(central)

        # 统一模块显示管理（轻量）
        self.module_manager = ModuleManager(self)

        # ---------------- 顶部菜单栏 ----------------
        menubar = self.menuBar()

        # 文件模块
        file_menu = menubar.addMenu("文件模块")
        self.act_table = QAction("表格数据", self)
        self.act_shp = QAction("Shapefile数据", self)
        self.act_raster = QAction("栅格数据", self)
        self.act_lidar = QAction("LiDAR数据", self)
        file_menu.addActions([self.act_table, self.act_shp, self.act_raster, self.act_lidar])

        # 数据处理模块
        data_menu = menubar.addMenu("数据处理模块")
        self.act_data_lidar_process = QAction("LiDAR数据处理", self)
        self.act_data_raster_calc = QAction("栅格计算器", self)
        self.act_data_plot = QAction("数据绘图", self)
        data_menu.addActions([
            self.act_data_lidar_process,
            self.act_data_raster_calc,
            self.act_data_plot,
        ])

        # 数据分析
        data_analysis_menu = menubar.addMenu("数据分析模块")
        self.act_data_corr = QAction("Pearson 相关性分析", self)
        self.act_data_significance = QAction("Spearman 相关性分析", self)
        self.act_data_importance = QAction("RF 特征重要性排序", self)
        self.act_data_explainability = QAction("可解释性分析 (SHAP)", self)
        self.act_data_descriptive_stats = QAction("描述统计", self)
        data_analysis_menu.addActions([
            self.act_data_corr,
            self.act_data_significance,
            self.act_data_importance,
            self.act_data_explainability,
            self.act_data_descriptive_stats
        ])

        # 特征提取模块
        analysis_menu = menubar.addMenu("特征提取模块")
        self.act_feature_single_tree = QAction("单木参数提取", self)
        self.act_feature_stand = QAction("林分参数提取", self)
        self.act_feature_gap = QAction("林窗识别", self)
        analysis_menu.addActions([self.act_feature_single_tree, self.act_feature_stand, self.act_feature_gap])

        # 结构分析模块
        structure_menu = menubar.addMenu("结构分析模块")
        self.act_structure_single_tree = QAction("单木结构分析", self)
        self.act_structure_stand_metrics = QAction("林分结构参数分析", self)
        self.act_structure_forest_inversion = QAction("区域森林参数反演", self)
        structure_menu.addActions([
            self.act_structure_single_tree,
            self.act_structure_stand_metrics,
            self.act_structure_forest_inversion,
        ])

        # 功能分析模块
        functional_menu = menubar.addMenu("功能分析模块")
        self.act_function_ecology = QAction("生态功能", self)
        self.act_function_production = QAction("生产力功能", self)
        self.act_function_social = QAction("社会服务功能", self)
        functional_menu.addActions([
            self.act_function_ecology,
            self.act_function_production,
            self.act_function_social,
        ])

        # 质量评价模块
        eval_menu = menubar.addMenu("质量评价模块")
        self.act_quality_single_tree = QAction("单木质量评价", self)
        self.act_quality_stand = QAction("林分质量评价", self)
        self.act_quality_region = QAction("区域质量评价", self)
        eval_menu.addActions([
            self.act_quality_single_tree,
            self.act_quality_stand,
            self.act_quality_region,
        ])

        # 驱动分析模块
        qudong_menu = menubar.addMenu("驱动分析模块")
        self.act_drive_structure_dynamics = QAction("结构动态与演替分析", self)
        qudong_menu.addActions([self.act_drive_structure_dynamics])

        # 趋势分析模块
        trend_menu = menubar.addMenu("趋势分析模块")
        self.act_trend_timeseries = QAction("趋势预测-时序变化建模", self)
        trend_menu.addActions([self.act_trend_timeseries])

        # ---------------- 主布局（左中右） ----------------
        main_layout = QHBoxLayout()

        # 左侧：数据目录
        self.tree_widget = QTreeWidget()
        self.tree_widget.setHeaderLabel("数据目录")
        root = QTreeWidgetItem(self.tree_widget, ["项目"])
        QTreeWidgetItem(root, ["光学数据"])
        QTreeWidgetItem(root, ["LiDAR数据"])
        QTreeWidgetItem(root, ["Shapefile"])
        QTreeWidgetItem(root, ["栅格数据"])
        self.tree_widget.expandAll()
        self.tree_widget.setFixedWidth(250)
        main_layout.addWidget(self.tree_widget)

        # 中间：Splitter（数据模块 + 动态模块区）
        self.center_splitter = QSplitter(Qt.Vertical)
        main_layout.addWidget(self.center_splitter, stretch=1)

        # 数据可视化模块
        self.data_module = DataFileModule(tree_widget=self.tree_widget, parent=self)
        self.center_splitter.addWidget(self.data_module)

        # 功能模块显示区（动态加载）
        self.module_container = QWidget()
        self.module_layout = QVBoxLayout(self.module_container)
        self.module_container.setLayout(self.module_layout)
        self.center_splitter.addWidget(self.module_container)
        self.center_splitter.setStretchFactor(0, 3)
        self.center_splitter.setStretchFactor(1, 2)

        # 右侧：工具箱
        self.toolbox_tree = QTreeWidget()
        self.toolbox_tree.setHeaderLabel("工具箱")
        toolbox_items = [
            "加载数据",
            "栅格计算处理",
            "森林参数反演",
            "结构参数计算",
            "数据绘图",
        ]
        for mod in toolbox_items:
            QTreeWidgetItem(self.toolbox_tree, [mod])
        self.toolbox_tree.expandAll()
        self.toolbox_tree.setFixedWidth(250)
        main_layout.addWidget(self.toolbox_tree)

        # 底部日志输出
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setFixedHeight(120)

        main_layout_vertical.addLayout(main_layout)
        main_layout_vertical.addWidget(self.log_output)
        self.action_registry = {
            "加载数据": self.act_table,
            "栅格计算处理": self.act_data_raster_calc,
            "森林参数反演": self.act_structure_forest_inversion,
            "结构参数计算": self.act_structure_stand_metrics,
            "数据绘图": self.act_data_plot,
        }

        # ---------------- 统一：构建映射表（注意：必须在 data_module 创建之后） ----------------
        self._build_action_maps()

        if hasattr(self.data_module, "tabs"):
            self.data_module.tabs.currentChanged.connect(self._on_data_tab_changed)

        # ---------------- 统一：信号绑定（用映射调度，不用 if/elif） ----------------
        self.tree_widget.itemDoubleClicked.connect(lambda item, col: self._dispatch_tree(item.text(col)))
        self.toolbox_tree.itemDoubleClicked.connect(lambda item, col: self._dispatch_toolbox(item.text(col)))

        self._bind_menu_actions()

        # 状态栏版权与徽标
        self._init_status_bar()

        # 初始日志
        self.log_output.append("系统初始化完成。")

    def _init_status_bar(self):
        bar = self.statusBar()
        bar.setSizeGripEnabled(False)
        bar.showMessage("系统就绪")

        show_status_logos = False
        if show_status_logos:
            base_dir = os.path.dirname(__file__)
            school_logo_path = os.path.join(base_dir, "R3.png")
            team_logo_path = os.path.join(base_dir, "R4.jpg")

            for logo_path in (school_logo_path, team_logo_path):
                if os.path.exists(logo_path):
                    pix = QPixmap(logo_path)
                    if not pix.isNull():
                        logo = QLabel()
                        logo.setPixmap(pix.scaled(18, 18, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                        logo.setToolTip(os.path.basename(logo_path))
                        bar.addPermanentWidget(logo)

        copyright_label = QLabel("© 2025 西南林业大学林智创新课题组  保留所有权利")
        copyright_label.setStyleSheet("padding-left: 6px; padding-right: 4px;")
        bar.addPermanentWidget(copyright_label)

    # ---------------- 构建映射表 ----------------
    def _build_action_maps(self):
        # 右侧工具箱映射
        self.toolbox_actions = {
            "加载数据": lambda: self.data_module.load_file_dialog("table"),
            "栅格计算处理": self.open_optical,
            "森林参数反演": self.open_inversion,
            "结构参数计算": self.open_structure_metrics,
            "数据绘图": self.open_plot_window,
        }

        # 左侧数据目录映射
        self.left_tree_actions = {
            "光学数据": self.open_optical,
            "LiDAR数据": lambda: self.data_module.load_file_dialog("lidar"),
            "Shapefile": lambda: self.data_module.load_file_dialog("shp"),
            "栅格数据": lambda: self.data_module.load_file_dialog("raster"),
        }

    # ---------------- 菜单绑定（统一在这里管理） ----------------
    def _bind_menu_actions(self):
        self.act_table.triggered.connect(lambda: self.data_module.load_file_dialog("table"))
        self.act_shp.triggered.connect(lambda: self.data_module.load_file_dialog("shp"))
        self.act_raster.triggered.connect(lambda: self.data_module.load_file_dialog("raster"))
        self.act_lidar.triggered.connect(lambda: self.data_module.load_file_dialog("lidar"))

        self.act_data_raster_calc.triggered.connect(self.open_optical)
        self.act_data_plot.triggered.connect(self.open_plot_window)
        self.act_structure_stand_metrics.triggered.connect(self.open_structure_metrics)
        self.act_data_descriptive_stats.triggered.connect(self.open_descriptive_stats)

        self.act_structure_forest_inversion.triggered.connect(self.open_inversion)

        self.act_quality_single_tree.triggered.connect(lambda: self.open_quality_window("single", "单木质量评价"))
        self.act_quality_stand.triggered.connect(lambda: self.open_quality_window("stand", "林分质量评价"))
        self.act_quality_region.triggered.connect(lambda: self.open_quality_window("region", "区域质量评价"))

        self.act_data_lidar_process.triggered.connect(lambda: self._log_placeholder("LiDAR数据处理"))
        self.act_data_corr.triggered.connect(lambda: self.open_data_statistics_analysis("Pearson 相关性分析"))
        self.act_data_significance.triggered.connect(lambda: self.open_data_statistics_analysis("Spearman 相关性分析"))
        self.act_data_importance.triggered.connect(lambda: self.open_data_statistics_analysis("RF 特征重要性排序"))
        self.act_data_explainability.triggered.connect(
            lambda: self.open_data_statistics_analysis("可解释性分析 (SHAP)"))

        self.act_feature_single_tree.triggered.connect(lambda: self._log_placeholder("单木参数提取"))
        self.act_feature_stand.triggered.connect(lambda: self._log_placeholder("林分参数提取"))
        self.act_feature_gap.triggered.connect(self.open_gap_detection)

        self.act_structure_single_tree.triggered.connect(self.open_single_tree_structure)

        self.act_function_ecology.triggered.connect(lambda: self.open_functional_analysis("生态功能", "生态功能"))
        self.act_function_production.triggered.connect(
            lambda: self.open_functional_analysis("生产力功能", "生产力功能"))
        self.act_function_social.triggered.connect(
            lambda: self.open_functional_analysis("社会服务功能", "社会服务功能"))

        self.act_drive_structure_dynamics.triggered.connect(self.open_drive_structure_dynamics)
        self.act_trend_timeseries.triggered.connect(self.open_trend_analysis)

    # ---------------- 左侧目录点击调度 ----------------
    def _dispatch_tree(self, name: str):
        action = self.left_tree_actions.get(name)
        if action:
            action()
            self.log_output.append(f"数据目录执行：{name}")
        else:
            self.log_output.append(f"提示：数据目录节点无动作：{name}")

    # ---------------- 右侧工具箱点击调度 ----------------
    def _dispatch_toolbox(self, name: str):
        action = self.toolbox_actions.get(name)
        if action:
            action()
            self.log_output.append(f"工具箱执行：{name}")
        else:
            self.log_output.append(f"⚠ 未注册的工具：{name}")

    def _show_and_focus_window(self, window: QWidget, log_message: str = "", splitter_ratio=None):
        window.show()
        window.raise_()
        window.activateWindow()
        if splitter_ratio is not None:
            self._apply_center_splitter_sizes(splitter_ratio[0], splitter_ratio[1])
        if log_message:
            self.log_output.append(log_message)

    def _open_external_process(
            self,
            process_attr: str,
            script_rel_path: str,
            opened_message: str,
            running_message: str,
            missing_message: str,
            error_prefix: str,
            script_args=None,
    ):
        process = getattr(self, process_attr)
        if process is not None and process.poll() is None:
            self.log_output.append(running_message)
            return

        script_path = os.path.join(os.path.dirname(__file__), script_rel_path)
        if not os.path.exists(script_path):
            self.log_output.append(missing_message)
            QMessageBox.critical(self, "错误", missing_message)
            return

        cmd = [sys.executable, script_path]
        if script_args:
            cmd.extend(script_args)

        try:
            setattr(
                self,
                process_attr,
                subprocess.Popen(cmd, cwd=os.path.dirname(__file__)),
            )
            self.log_output.append(opened_message)
        except Exception as exc:
            self.log_output.append(f"{error_prefix}：{exc}")
            QMessageBox.critical(self, "错误", f"{error_prefix}：\n{exc}")

    # ---------------- 模块打开函数 ----------------
    def open_optical(self):
        if self.optical_window is None:
            self.optical_window = OpticalModule(self.data_module)
        self._show_and_focus_window(
            self.optical_window,
            log_message="栅格计算器窗口已打开",
            splitter_ratio=(0.58, 0.42),
        )

    def open_inversion(self):
        if self.inversion_window is None:
            self.inversion_window = ForestInversionApp()
        self._show_and_focus_window(self.inversion_window, log_message="区域森林参数反演窗口已打开")

    def open_plot_window(self):
        if self.plot_window is None:
            self.plot_window = DataPlotWindow(self.data_module, self)
        self._show_and_focus_window(
            self.plot_window,
            log_message="数据绘图窗口已打开",
            splitter_ratio=(0.58, 0.42),
        )

    def open_descriptive_stats(self):
        self._open_external_process(
            process_attr="descriptive_stats_process",
            script_rel_path=os.path.join("modules", "descriptive_stats_gui.py"),
            opened_message="描述统计窗口已打开",
            running_message="描述统计窗口已在运行",
            missing_message="未找到描述分析模块文件：modules/descriptive_stats_gui.py",
            error_prefix="描述统计模块打开失败",
        )

    def open_data_statistics_analysis(self, analysis_name: str):
        self._open_external_process(
            process_attr="data_statistics_process",
            script_rel_path=os.path.join("modules", "data_statistics_analysis_module.py"),
            opened_message=f"{analysis_name}窗口已打开",
            running_message=f"{analysis_name}窗口已在运行。",
            missing_message=f"未找到{analysis_name}模块文件：modules/data_statistics_analysis_module.py",
            error_prefix=f"{analysis_name}模块打开失败",
            script_args=["--analysis", analysis_name, "--window-title", analysis_name],
        )

    def open_drive_structure_dynamics(self):
        self._open_external_process(
            process_attr="drive_structure_dynamics_process",
            script_rel_path=os.path.join("modules", "drive_structure_dynamics_module.py"),
            opened_message="结构动态与演替分析窗口已打开",
            running_message="结构动态与演替分析窗口已在运行",
            missing_message="未找到结构动态与演替分析模块文件：modules/drive_structure_dynamics_module.py",
            error_prefix="结构动态与演替分析模块打开失败",
        )

    def open_structure(self):
        self.module_manager.show(StructureModule(), "林分结构估计模块已打开")

    def open_gap_detection(self):
        if self.gap_detection_window is None:
            try:
                from modules.forest_gap_studio import GapDetectionModule
                self.gap_detection_window = GapDetectionModule(self)
            except Exception as exc:
                self.log_output.append(f"林窗识别模块打开失败：{exc}")
                QMessageBox.critical(self, "错误",
                                     f"林窗识别模块打开失败，请检查 modules/forest_gap_studio.py 是否存在：\n{exc}")
                return

        self.gap_detection_window.show()
        self.gap_detection_window.raise_()
        self.gap_detection_window.activateWindow()
        self.log_output.append("林窗识别窗口已打开")

    def open_single_tree_structure(self):
        if self.single_tree_structure_window is None:
            try:
                from modules.single_tree_structure_module import SingleTreeStructureWindow
                self.single_tree_structure_window = SingleTreeStructureWindow()
            except Exception as exc:
                self.log_output.append(f"单木结构分析模块打开失败：{exc}")
                QMessageBox.critical(self, "错误", f"单木结构分析模块打开失败：\n{exc}")
                return

        self.single_tree_structure_window.setWindowTitle("单木结构分析")
        self.single_tree_structure_window.show()
        self.single_tree_structure_window.raise_()
        self.single_tree_structure_window.activateWindow()
        self.log_output.append("单木结构分析窗口已打开")

    def open_functional_analysis(self, top_module_name: str, menu_name: str):
        if self.functional_window is None:
            try:
                from modules.functional_analysis_module import FunctionalAnalysisWindow
                self.functional_window = FunctionalAnalysisWindow()
            except Exception as exc:
                self.log_output.append(f"{menu_name}模块打开失败：{exc}")
                return

        if hasattr(self.functional_window, "select_top_module"):
            self.functional_window.select_top_module(top_module_name)
        elif hasattr(self.functional_window, "top_combo"):
            idx = self.functional_window.top_combo.findText(top_module_name)
            if idx >= 0:
                self.functional_window.top_combo.setCurrentIndex(idx)

        self.functional_window.setWindowTitle(menu_name)
        self.functional_window.show()
        self.functional_window.raise_()
        self.functional_window.activateWindow()
        self.log_output.append(f"{menu_name}窗口已打开")

    def open_trend_analysis(self):
        if self.trend_analysis_window is None:
            try:
                from modules.trend_analysis_module import TrendAnalysisApp
                self.trend_analysis_window = TrendAnalysisApp()
            except Exception as exc:
                self.log_output.append(f"趋势分析模块打开失败：{exc}")
                QMessageBox.critical(self, "错误", f"趋势分析模块打开失败：\n{exc}")
                return

        self.trend_analysis_window.setWindowTitle("趋势预测-时序变化建模")
        self.trend_analysis_window.show()
        self.trend_analysis_window.raise_()
        self.trend_analysis_window.activateWindow()
        self.log_output.append("趋势预测-时序变化建模窗口已打开")

    def open_structure_metrics(self):
        if self.structure_metrics_window is None:
            self.structure_metrics_window = StructureMetricsWindow(self.data_module, self)
        self.structure_metrics_window.show()
        self.structure_metrics_window.raise_()
        self.structure_metrics_window.activateWindow()
        self.log_output.append("林分结构参数分析窗口已打开")

    def open_quality_window(self, key: str, title: str):
        window = self.quality_windows.get(key)
        if window is None:
            if key == "single":
                window = self._create_single_tree_window(title)
            elif key == "stand":
                window = StandQualityWindow(self)
            else:
                window = QualityEvaluationWindow(title, self)
            self.quality_windows[key] = window
        window.show()
        window.raise_()
        window.activateWindow()
        self.log_output.append(f"{title}窗口已打开")

    def open_pca(self, mode="单木"):
        self.open_quality_window(mode, f"{mode}质量评价")

    def _create_single_tree_window(self, title: str):
        if IndividualTreeQualityWindow is not None:
            window = IndividualTreeQualityWindow()
            window.setWindowTitle(title)
            return window
        self.log_output.append("提示：未能加载 modules/individual_tree_quality_module.py，已回退到内置单木质量评价窗口。")
        return SingleTreeQualityWindow(self.data_module, self)

    def _log_placeholder(self, module_name: str):
        self.log_output.append(f"{module_name}模块已就绪，功能建设中。")

    def _apply_center_splitter_sizes(self, top_ratio: float, bottom_ratio: float):
        total = max(self.center_splitter.height(), 1)
        ratio_sum = max(top_ratio + bottom_ratio, 1e-6)
        top_size = int(total * (top_ratio / ratio_sum))
        bottom_size = max(total - top_size, 1)
        self.center_splitter.setSizes([max(top_size, 1), bottom_size])

    def _on_data_tab_changed(self, index: int):
        table_tab = getattr(self.data_module, "table_tab", None)
        tabs = getattr(self.data_module, "tabs", None)
        if table_tab is not None and tabs is not None and tabs.indexOf(table_tab) == index:
            self._apply_center_splitter_sizes(0.78, 0.22)
        else:
            self._apply_center_splitter_sizes(0.62, 0.38)

    # ---------------- 通用显示模块 ----------------
    def show_module(self, widget: QWidget):
        for i in reversed(range(self.module_layout.count())):
            old_widget = self.module_layout.itemAt(i).widget()
            if old_widget is not None:
                old_widget.setParent(None)

        self.module_layout.addWidget(widget)


if __name__ == "__main__":
    _install_console_warning_filter()
    app = QApplication(sys.argv)
    apply_qt_app_style(app)
    win = ForestMain()
    win.show()
    sys.exit(app.exec_())