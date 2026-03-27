# -*- coding: utf-8 -*-
"""数据统计分析模块（彻底迁移版）。"""

import argparse
import os
import sys

import numpy as np
import pandas as pd

import matplotlib

matplotlib.use("Qt5Agg")
import matplotlib.pyplot as plt

from PyQt5.QtCore import QThread, Qt, pyqtSignal
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from scipy.stats import pearsonr, spearmanr
from sklearn.ensemble import RandomForestRegressor
try:
    from modules.ui_style import apply_qt_app_style, apply_qt_window_baseline
except ModuleNotFoundError:
    from ui_style import apply_qt_app_style, apply_qt_window_baseline

plt.rcParams["font.sans-serif"] = ["SimSun", "Songti SC", "STSong", "SimHei", "Microsoft YaHei"]
plt.rcParams["font.family"] = "sans-serif"
plt.rcParams["axes.unicode_minus"] = False

ANALYSIS_ITEMS = [
    "Pearson 相关性分析",
    "Spearman 相关性分析",
    "RF 特征重要性排序",
    "可解释性分析 (SHAP)",
]


class AnalysisCore:
    @staticmethod
    def get_correlation(df: pd.DataFrame, method: str = "pearson"):
        corr_matrix = df.corr(method=method)
        p_matrix = pd.DataFrame(np.zeros_like(corr_matrix), columns=corr_matrix.columns, index=corr_matrix.index)

        for col1 in df.columns:
            for col2 in df.columns:
                if col1 == col2:
                    p_matrix.loc[col1, col2] = 0.0
                elif method == "pearson":
                    _, p = pearsonr(df[col1], df[col2])
                    p_matrix.loc[col1, col2] = p
                else:
                    _, p = spearmanr(df[col1], df[col2])
                    p_matrix.loc[col1, col2] = p
        return corr_matrix, p_matrix

    @staticmethod
    def get_rf_model(x_data: pd.DataFrame, y_data: pd.Series, seed: int = 42):
        model = RandomForestRegressor(n_estimators=100, max_depth=20, random_state=seed, n_jobs=1)
        model.fit(x_data, y_data)
        return model


class MLWorker(QThread):
    finished = pyqtSignal(object, object, object, str)
    error = pyqtSignal(str)

    def __init__(self, x_data: pd.DataFrame, y_data: pd.Series, seed: int, mode: str):
        super().__init__()
        self.x_data = x_data
        self.y_data = y_data
        self.seed = seed
        self.mode = mode

    def run(self):
        try:
            max_train_samples = 20000
            if self.x_data.shape[0] > max_train_samples:
                np.random.seed(self.seed)
                train_idx = np.random.choice(self.x_data.shape[0], max_train_samples, replace=False)
                train_x = self.x_data.iloc[train_idx]
                train_y = self.y_data.iloc[train_idx]
            else:
                train_x = self.x_data
                train_y = self.y_data

            model = AnalysisCore.get_rf_model(train_x, train_y, self.seed)

            if self.mode == "rf":
                importance = pd.DataFrame(
                    {"Importance (Gini)": model.feature_importances_}, index=self.x_data.columns
                ).sort_values(by="Importance (Gini)", ascending=False)
                self.finished.emit(importance, None, None, self.mode)
                return

            if self.mode == "shap":
                import shap

                max_plot_samples = 500
                if self.x_data.shape[0] > max_plot_samples:
                    np.random.seed(self.seed)
                    sample_idx = np.random.choice(self.x_data.shape[0], max_plot_samples, replace=False)
                    plot_x = self.x_data.iloc[sample_idx]
                else:
                    plot_x = self.x_data

                explainer = shap.TreeExplainer(model)
                shap_values = explainer.shap_values(plot_x, check_additivity=False, approximate=True)
                mean_abs_shap = np.abs(shap_values).mean(axis=0)
                shap_table = pd.DataFrame(
                    {"Mean |SHAP| (Impact Magnitude)": mean_abs_shap}, index=plot_x.columns
                ).sort_values(by="Mean |SHAP| (Impact Magnitude)", ascending=False)
                self.finished.emit(shap_table, shap_values, plot_x, self.mode)
        except Exception as exc:
            self.error.emit(str(exc))


class CopyableTableWidget(QTableWidget):
    def keyPressEvent(self, event):
        if event.matches(QKeySequence.Copy):
            self.copy_selection()
            return
        super().keyPressEvent(event)

    def copy_selection(self):
        selection = self.selectedIndexes()
        if not selection:
            return
        rows = sorted({index.row() for index in selection})
        cols = sorted({index.column() for index in selection})
        chunks = []
        for row in rows:
            row_data = []
            for col in cols:
                item = self.item(row, col)
                row_data.append(item.text() if item else "")
            chunks.append("\t".join(row_data))
        QApplication.clipboard().setText("\n".join(chunks))


class DataStatisticsAnalysisWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("数据统计分析工具 (UI精简版)")
        apply_qt_window_baseline(self, size=(1320, 860))
        self.data = None
        self.ml_worker = None
        self.current_analysis_type = None
        self.current_table_data = None
        self.current_target_name = ""
        self._init_ui()

    def _init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)

        left_panel = QVBoxLayout()
        self.btn_load = QPushButton("📂 1. 导入数据 (CSV/Excel)")
        self.btn_load.setMinimumHeight(40)
        self.btn_load.setStyleSheet("font-weight: bold; font-size: 11pt;")
        self.btn_load.clicked.connect(self.load_data)
        left_panel.addWidget(self.btn_load)
        left_panel.addSpacing(10)

        left_panel.addWidget(QLabel("<b>2. 选择分析模块:</b>"))
        self.analysis_combo = QComboBox()
        self.analysis_combo.addItems(ANALYSIS_ITEMS)
        self.analysis_combo.setMinimumHeight(30)
        left_panel.addWidget(self.analysis_combo)
        left_panel.addSpacing(10)

        column_group = QGroupBox("3. 选择参与的列")
        column_layout = QVBoxLayout()
        column_layout.addWidget(QLabel("<b>目标变量 (Y):</b>"))
        self.target_combo = QComboBox()
        self.target_combo.currentTextChanged.connect(self.update_feature_list)
        column_layout.addWidget(self.target_combo)
        column_layout.addSpacing(10)

        column_layout.addWidget(QLabel("<b>参与计算的特征 (X):</b>"))
        self.feature_list = QListWidget()
        column_layout.addWidget(self.feature_list)

        btn_layout = QHBoxLayout()
        self.btn_select_all = QPushButton("全选")
        self.btn_select_all.clicked.connect(lambda: self.set_all_features(Qt.Checked))
        self.btn_deselect_all = QPushButton("清空")
        self.btn_deselect_all.clicked.connect(lambda: self.set_all_features(Qt.Unchecked))
        btn_layout.addWidget(self.btn_select_all)
        btn_layout.addWidget(self.btn_deselect_all)
        column_layout.addLayout(btn_layout)

        column_group.setLayout(column_layout)
        left_panel.addWidget(column_group)
        left_panel.addSpacing(10)

        left_panel.addWidget(QLabel("<b>设置出图的特征数量 (Top N):</b>"))
        self.top_n_spinbox = QSpinBox()
        self.top_n_spinbox.setRange(1, 999)
        self.top_n_spinbox.setValue(20)
        self.top_n_spinbox.setMinimumHeight(30)
        left_panel.addWidget(self.top_n_spinbox)
        left_panel.addSpacing(10)

        left_panel.addWidget(QLabel("<b>4. 输出保存路径:</b>"))
        path_layout = QHBoxLayout()
        self.path_input = QLineEdit(os.getcwd())
        self.path_input.setReadOnly(True)
        self.path_input.setMinimumHeight(30)
        self.btn_path = QPushButton("选择路径")
        self.btn_path.setMinimumHeight(30)
        self.btn_path.clicked.connect(self.select_output_dir)
        path_layout.addWidget(self.path_input)
        path_layout.addWidget(self.btn_path)
        left_panel.addLayout(path_layout)
        left_panel.addSpacing(10)

        left_panel.addWidget(QLabel("<b>设置随机种子 (Random Seed):</b>"))
        self.seed_spinbox = QSpinBox()
        self.seed_spinbox.setRange(0, 999999)
        self.seed_spinbox.setValue(42)
        self.seed_spinbox.setMinimumHeight(30)
        left_panel.addWidget(self.seed_spinbox)
        left_panel.addSpacing(15)

        self.btn_run = QPushButton("🚀 开始运行数据分析")
        self.btn_run.setMinimumHeight(45)
        self.btn_run.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; font-size: 11pt; border-radius: 5px;")
        self.btn_run.clicked.connect(self.execute_analysis)
        left_panel.addWidget(self.btn_run)
        left_panel.addSpacing(10)

        self.btn_save_all = QPushButton("💾 一键保存结果")
        self.btn_save_all.setMinimumHeight(45)
        self.btn_save_all.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; font-size: 11pt; border-radius: 5px;")
        self.btn_save_all.clicked.connect(self.save_current_results)
        left_panel.addWidget(self.btn_save_all)
        left_panel.addStretch()

        right_panel = QVBoxLayout()
        right_panel.addWidget(QLabel("<b>📊 分析数据表</b>"))

        self.table = CopyableTableWidget()
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setStyleSheet(
            """
            QTableWidget {
                alternate-background-color: #f9f9f9;
                gridline-color: #e0e0e0;
                border: 1px solid #d0d0d0;
                font-family: "Times New Roman", "SimSun", "Microsoft YaHei";
                font-size: 10pt;
            }
            QHeaderView::section {
                background-color: #5c7c99;
                color: white;
                font-family: "Times New Roman", "SimSun", "Microsoft YaHei";
                font-weight: bold;
                font-size: 10pt;
                padding: 4px;
                border: 1px solid #d0d0d0;
            }
            """
        )
        right_panel.addWidget(self.table, 1)
        right_panel.addWidget(QLabel("<b>📈 交互式图表</b>"))

        self.figure = plt.figure(figsize=(9, 7))
        self.canvas = FigureCanvas(self.figure)
        self.toolbar = NavigationToolbar(self.canvas, self)
        right_panel.addWidget(self.toolbar)
        right_panel.addWidget(self.canvas, 2)

        main_layout.addLayout(left_panel, 1)
        main_layout.addLayout(right_panel, 3)

    def load_data(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "Data Files (*.csv *.xlsx)")
        if not path:
            return
        basename = os.path.splitext(os.path.basename(path))[0]
        if basename in ["1", "2", "3", "L1"]:
            QMessageBox.warning(self, "拦截提示", f"已触发内部限制条件：文件名“{basename}”不参与运算。")
            return
        try:
            if path.endswith(".csv"):
                try:
                    self.data = pd.read_csv(path, encoding="utf-8")
                except UnicodeDecodeError:
                    self.data = pd.read_csv(path, encoding="gbk")
            else:
                self.data = pd.read_excel(path)

            if "CZhfp" in self.data.columns:
                self.data.rename(columns={"CZhfp": "HFP"}, inplace=True)

            self.target_combo.blockSignals(True)
            self.target_combo.clear()
            self.target_combo.addItems(self.data.columns)
            self.target_combo.blockSignals(False)
            self.update_feature_list()
            QMessageBox.information(self, "成功", f"成功加载数据，共 {self.data.shape[0]} 行。")
        except Exception as exc:
            QMessageBox.critical(self, "错误", f"读取文件失败: {str(exc)}")

    def select_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择保存路径", self.path_input.text())
        if dir_path:
            self.path_input.setText(dir_path)

    def update_feature_list(self):
        if self.data is None:
            return
        self.feature_list.clear()
        target = self.target_combo.currentText()
        numeric_cols = self.data.select_dtypes(include=[np.number]).columns
        for col in numeric_cols:
            if col == target:
                continue
            item = QListWidgetItem(col)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            self.feature_list.addItem(item)

    def set_all_features(self, state):
        for i in range(self.feature_list.count()):
            self.feature_list.item(i).setCheckState(state)

    def get_selected_features(self):
        return [
            self.feature_list.item(i).text()
            for i in range(self.feature_list.count())
            if self.feature_list.item(i).checkState() == Qt.Checked
        ]

    def render_dataframe_to_table(self, df: pd.DataFrame, index_name: str = "Features"):
        self.table.clear()
        df_display = df.copy()
        df_display.insert(0, index_name, df_display.index)
        self.table.setRowCount(df_display.shape[0])
        self.table.setColumnCount(df_display.shape[1])
        self.table.setHorizontalHeaderLabels(df_display.columns.astype(str))
        for i in range(df_display.shape[0]):
            for j in range(df_display.shape[1]):
                val = df_display.iloc[i, j]
                item = QTableWidgetItem(f"{val:.4f}" if isinstance(val, float) else str(val))
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(i, j, item)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

    def execute_analysis(self):
        if self.data is None:
            QMessageBox.warning(self, "提醒", "请先导入数据！")
            return
        mode = self.analysis_combo.currentText()
        if "相关性分析" in mode:
            method = "pearson" if "Pearson" in mode else "spearman"
            self.run_corr(method)
        elif "RF" in mode:
            self.run_ml_analysis(mode="rf")
        elif "SHAP" in mode:
            self.run_ml_analysis(mode="shap")

    def run_corr(self, method: str):
        selected_features = self.get_selected_features()
        target = self.target_combo.currentText()
        if target in self.data.select_dtypes(include=[np.number]).columns and target not in selected_features:
            selected_features.append(target)
        if len(selected_features) < 2:
            QMessageBox.warning(self, "警告", "请至少选择两列用于计算相关性！")
            return
        df_selected = self.data[selected_features].apply(pd.to_numeric, errors="coerce").dropna()
        if df_selected.empty:
            QMessageBox.critical(self, "错误", "剔除缺失值后数据为空！")
            return

        corr_matrix, p_matrix = AnalysisCore.get_correlation(df_selected, method)
        formatted_corr = pd.DataFrame(index=corr_matrix.index, columns=corr_matrix.columns)
        for i in corr_matrix.index:
            for j in corr_matrix.columns:
                r = corr_matrix.loc[i, j]
                p = p_matrix.loc[i, j]
                if i == j:
                    formatted_corr.loc[i, j] = f"{r:.3f}"
                else:
                    stars = "**" if p < 0.01 else ("*" if p < 0.05 else "")
                    formatted_corr.loc[i, j] = f"{r:.3f}{stars}"

        self.current_analysis_type = "Correlation"
        self.current_table_data = formatted_corr
        self.current_target_name = method.capitalize()
        self.render_dataframe_to_table(formatted_corr, index_name="Variables")

        self.figure.clear()
        ax = self.figure.add_subplot(111)
        n_vars = corr_matrix.shape[0]
        show_annot = n_vars <= 12

        try:
            import seaborn as sns
        except ImportError:
            QMessageBox.critical(self, "依赖缺失", "缺少 seaborn，请安装后再进行相关性热力图绘制。")
            return

        sns.heatmap(corr_matrix, annot=show_annot, cmap="RdBu_r", fmt=".2f", ax=ax, annot_kws={"family": "Times New Roman"})
        font_size = max(5, 11 - int(n_vars / 15))
        plt.setp(ax.get_xticklabels(), fontfamily="SimSun", fontsize=font_size, rotation=45, ha="right")
        plt.setp(ax.get_yticklabels(), fontfamily="SimSun", fontsize=font_size)
        ax.set_title("QI Correlation Analysis", fontfamily="Times New Roman", fontsize=12, pad=15)
        ax.set_xlabel("样地QI", fontfamily="SimSun", fontsize=10)
        ax.set_ylabel("LiDAR QI", fontfamily="SimSun", fontsize=10)
        self.figure.tight_layout()
        self.canvas.draw()

    def run_ml_analysis(self, mode: str):
        target = self.target_combo.currentText()
        selected_features = self.get_selected_features()
        if not selected_features:
            QMessageBox.warning(self, "警告", "请至少勾选一个特征！")
            return
        df_ml = (
            self.data[selected_features + [target]]
            .replace([np.inf, -np.inf], np.nan)
            .apply(pd.to_numeric, errors="coerce")
            .dropna()
        )
        if df_ml.empty:
            QMessageBox.critical(self, "错误", "数据全部是空值或非数值，无法计算！")
            return

        x_data = df_ml[selected_features]
        y_data = df_ml[target]
        current_seed = self.seed_spinbox.value()

        self.btn_run.setEnabled(False)
        self.btn_run.setText("⏳ 飞速计算中...")
        self.btn_run.setStyleSheet("background-color: #FFA500; color: white; font-weight: bold; font-size: 11pt; border-radius: 5px;")
        QApplication.processEvents()

        self.ml_worker = MLWorker(x_data, y_data, current_seed, mode)
        self.ml_worker.finished.connect(self._on_ml_finished)
        self.ml_worker.error.connect(self._on_ml_error)
        self.ml_worker.start()

    def _on_ml_finished(self, table_data, plot_shap, plot_x, mode):
        self.btn_run.setEnabled(True)
        self.btn_run.setText("🚀 开始运行数据分析")
        self.btn_run.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; font-size: 11pt; border-radius: 5px;")

        self.current_table_data = table_data
        self.render_dataframe_to_table(table_data, index_name="Feature")
        self.figure.clear()
        top_n = self.top_n_spinbox.value()

        if mode == "rf":
            self.current_analysis_type = "RF_Importance"
            ax = self.figure.add_subplot(111)
            importance_subset = table_data.head(top_n).iloc[::-1]
            colors = plt.cm.tab20(np.linspace(0, 1, len(importance_subset)))
            ax.barh(importance_subset.index, importance_subset["Importance (Gini)"], color=colors, edgecolor="none")
            n_vars = len(importance_subset)
            font_size = max(8, 12 - int(n_vars / 10))
            plt.setp(ax.get_xticklabels(), fontfamily="Times New Roman", fontsize=10)
            plt.setp(ax.get_yticklabels(), fontfamily="Times New Roman", fontweight="bold", fontsize=font_size)
            ax.set_title(f"Random Forest Feature Importance ({self.current_target_name})", fontfamily="Times New Roman", fontsize=12, pad=15)
            ax.set_xlabel("Importance (Gini Impurity)", fontfamily="Times New Roman", fontweight="bold", fontsize=10)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
        elif mode == "shap":
            self.current_analysis_type = "SHAP_Analysis"
            import shap

            plt.figure(self.figure.number)
            shap.summary_plot(plot_shap, plot_x, show=False, max_display=top_n)
            ax = plt.gca()
            n_vars = min(top_n, len(plot_x.columns))
            font_size = max(8, 12 - int(n_vars / 10))
            plt.setp(ax.get_xticklabels(), fontfamily="Times New Roman", fontsize=10)
            plt.setp(ax.get_yticklabels(), fontfamily="SimSun", fontweight="bold", fontsize=font_size)
            ax.set_title(f"SHAP Value Impact on Output ({self.current_target_name})", fontfamily="Times New Roman", fontsize=12, pad=15)

        self.figure.tight_layout()
        self.canvas.draw()

    def _on_ml_error(self, err_msg: str):
        self.btn_run.setEnabled(True)
        self.btn_run.setText("🚀 开始运行数据分析")
        self.btn_run.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; font-size: 11pt; border-radius: 5px;")
        QMessageBox.critical(self, "错误", f"后台分析失败: {err_msg}")

    def save_current_results(self):
        if self.current_analysis_type is None:
            QMessageBox.warning(self, "提醒", "当前没有任何分析结果，请先运行分析！")
            return
        output_dir = self.path_input.text()
        if not os.path.exists(output_dir):
            QMessageBox.critical(self, "错误", "输出路径不存在，请重新选择！")
            return

        try:
            safe_target_name = str(self.current_target_name).replace("/", "_").replace("\\", "_")
            excel_name = f"{self.current_analysis_type}_Table_{safe_target_name}.xlsx"
            excel_path = os.path.join(output_dir, excel_name)

            df_export = self.current_table_data.copy()
            index_label = "Variables" if self.current_analysis_type == "Correlation" else "Feature"
            df_export.insert(0, index_label, df_export.index)
            with pd.ExcelWriter(excel_path) as writer:
                df_export.to_excel(writer, sheet_name="Results", index=False)

            img_name = f"{self.current_analysis_type}_Plot_{safe_target_name}.png"
            img_path = os.path.join(output_dir, img_name)
            self.figure.savefig(img_path, dpi=600, bbox_inches="tight")
            QMessageBox.information(self, "导出成功", f"结果已成功无损导出！\n\n表格: {excel_name}\n图片: {img_name} (600dpi)")
        except Exception as exc:
            QMessageBox.critical(self, "错误", f"导出失败: {str(exc)}")


# 兼容旧类名
MainWindow = DataStatisticsAnalysisWindow


def launch(default_analysis: str = "", window_title: str = "") -> int:
    app = QApplication(sys.argv)
    apply_qt_app_style(app)
    window = DataStatisticsAnalysisWindow()
    if window_title:
        window.setWindowTitle(window_title)
    if default_analysis:
        idx = window.analysis_combo.findText(default_analysis)
        if idx >= 0:
            window.analysis_combo.setCurrentIndex(idx)
    window.show()
    return app.exec_()


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="启动数据统计分析模块")
    parser.add_argument("--analysis", default="", choices=ANALYSIS_ITEMS, help="默认分析项名称")
    parser.add_argument("--window-title", default="", help="窗口标题")
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    raise SystemExit(launch(args.analysis, args.window_title))

