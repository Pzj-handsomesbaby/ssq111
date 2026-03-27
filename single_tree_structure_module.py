# -*- coding: utf-8 -*-
"""单木结构分析模块（彻底迁移版）。

该文件已完整迁移 `lxh/yisushengz.py` 的功能实现，
主程序可直接从 `modules.single_tree_structure_module` 导入窗口类。
"""

import sys
import pandas as pd
import numpy as np
from scipy.optimize import curve_fit
from sklearn.metrics import r2_score
from sklearn.model_selection import KFold, LeaveOneOut
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QLineEdit,
    QFileDialog,
    QComboBox,
    QTextEdit,
    QMessageBox,
    QGroupBox,
    QGridLayout,
    QListWidget,
    QListWidgetItem,
    QSpinBox,
)
from PyQt5.QtCore import Qt

import matplotlib

matplotlib.use("Qt5Agg")
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
try:
    from modules.ui_style import apply_qt_app_style, apply_qt_window_baseline
except ModuleNotFoundError:
    from ui_style import apply_qt_app_style, apply_qt_window_baseline

plt.rcParams["font.sans-serif"] = ["SimHei", "Arial Unicode MS"]
plt.rcParams["axes.unicode_minus"] = False


class SingleTreeStructureWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.df = None
        self.plot_data = []
        self.current_plot_index = 0
        self.initUI()

    def initUI(self):
        self.setWindowTitle("专业级多模型拟合工具 (多模型对比 & 交叉验证)")
        apply_qt_window_baseline(self, size=(1220, 820))

        main_layout = QHBoxLayout()
        left_panel = QVBoxLayout()

        file_group = QGroupBox("1. 加载与结果保存 (.csv, .xlsx)")
        file_layout = QGridLayout()

        self.input_path = QLineEdit()
        self.input_btn = QPushButton("浏览并读取")
        self.input_btn.clicked.connect(self.load_data)

        self.output_path = QLineEdit()
        self.output_btn = QPushButton("输出保存为...")
        self.output_btn.clicked.connect(self.select_output_file)

        file_layout.addWidget(QLabel("输入数据:"), 0, 0)
        file_layout.addWidget(self.input_path, 0, 1)
        file_layout.addWidget(self.input_btn, 0, 2)
        file_layout.addWidget(QLabel("结果输出:"), 1, 0)
        file_layout.addWidget(self.output_path, 1, 1)
        file_layout.addWidget(self.output_btn, 1, 2)
        file_group.setLayout(file_layout)
        left_panel.addWidget(file_group)

        var_group = QGroupBox("2. 变量与模型设置 (勾选需要的选项)")
        var_layout = QGridLayout()

        self.target_combo = QComboBox()
        self.feature_list = QListWidget()

        self.model_list = QListWidget()
        models = ["幂函数", "对数线性", "多元线性", "指数"]
        for m in models:
            item = QListWidgetItem(m)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            self.model_list.addItem(item)
        self.model_list.setMaximumHeight(100)

        var_layout.addWidget(QLabel("因变量 (Y):"), 0, 0)
        var_layout.addWidget(self.target_combo, 0, 1)
        var_layout.addWidget(QLabel("自变量 (X):"), 1, 0, Qt.AlignTop)
        var_layout.addWidget(self.feature_list, 1, 1)
        var_layout.addWidget(QLabel("选择模型:"), 2, 0, Qt.AlignTop)
        var_layout.addWidget(self.model_list, 2, 1)
        var_group.setLayout(var_layout)
        left_panel.addWidget(var_group)

        cv_group = QGroupBox("3. 交叉验证设置")
        cv_layout = QHBoxLayout()

        self.cv_combo = QComboBox()
        self.cv_combo.addItems(["无 (全部数据拟合)", "K折交叉验证 (K-Fold)", "留一交叉验证 (LOO)"])
        self.cv_combo.currentIndexChanged.connect(self.update_cv_state)

        self.k_label = QLabel("K值:")
        self.k_spin = QSpinBox()
        self.k_spin.setRange(2, 20)
        self.k_spin.setValue(5)
        self.k_spin.setEnabled(False)

        cv_layout.addWidget(QLabel("验证方法:"))
        cv_layout.addWidget(self.cv_combo)
        cv_layout.addWidget(self.k_label)
        cv_layout.addWidget(self.k_spin)
        cv_group.setLayout(cv_layout)
        left_panel.addWidget(cv_group)

        log_group = QGroupBox("4. 日志与执行")
        log_layout = QVBoxLayout()
        self.run_btn = QPushButton("开始模型拟合与验证")
        self.run_btn.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; padding: 10px;")
        self.run_btn.clicked.connect(self.run_fitting)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)

        log_layout.addWidget(self.run_btn)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        left_panel.addWidget(log_group)

        plot_group = QGroupBox("拟合结果与可视化")
        plot_layout = QVBoxLayout()

        self.result_banner = QTextEdit("等待加载数据并拟合...")
        self.result_banner.setReadOnly(True)
        self.result_banner.setMaximumHeight(120)
        self.result_banner.setStyleSheet(
            """
            QTextEdit {
                font-size: 14px;
                color: #2c3e50;
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 4px;
            }
            """
        )

        nav_layout = QHBoxLayout()
        self.btn_prev = QPushButton("◀ 上一个模型")
        self.btn_prev.clicked.connect(self.show_prev_plot)
        self.btn_prev.setEnabled(False)

        self.lbl_plot_info = QLabel("图表: 0 / 0")
        self.lbl_plot_info.setAlignment(Qt.AlignCenter)
        self.lbl_plot_info.setStyleSheet("font-weight: bold;")

        self.btn_next = QPushButton("下一个模型 ▶")
        self.btn_next.clicked.connect(self.show_next_plot)
        self.btn_next.setEnabled(False)

        self.btn_save_fig = QPushButton("保存当前图表 (PNG)")
        self.btn_save_fig.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        self.btn_save_fig.clicked.connect(self.save_current_plot)
        self.btn_save_fig.setEnabled(False)

        nav_layout.addWidget(self.btn_prev)
        nav_layout.addWidget(self.lbl_plot_info)
        nav_layout.addWidget(self.btn_next)
        nav_layout.addWidget(self.btn_save_fig)

        self.figure = Figure()
        self.canvas = FigureCanvas(self.figure)

        plot_layout.addWidget(self.result_banner)
        plot_layout.addLayout(nav_layout)
        plot_layout.addWidget(self.canvas)
        plot_group.setLayout(plot_layout)

        main_layout.addLayout(left_panel, 1)
        main_layout.addWidget(plot_group, 2)
        self.setLayout(main_layout)

    def update_cv_state(self):
        idx = self.cv_combo.currentIndex()
        self.k_spin.setEnabled(idx == 1)

    def load_data(self):
        file_filter = "Data Files (*.csv *.xlsx *.xls);;CSV Files (*.csv);;Excel Files (*.xlsx *.xls)"
        filename, _ = QFileDialog.getOpenFileName(self, "选择输入文件", "", file_filter)
        if not filename:
            return

        self.input_path.setText(filename)
        try:
            if filename.lower().endswith(".csv"):
                self.df = pd.read_csv(filename)
            else:
                self.df = pd.read_excel(filename)

            columns = self.df.columns.tolist()
            if not columns:
                raise ValueError("文件为空或格式不正确。")

            self.target_combo.clear()
            self.feature_list.clear()
            self.target_combo.addItems(columns)

            for col in columns:
                item = QListWidgetItem(col)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked)
                self.feature_list.addItem(item)

            if "AGB" in columns:
                self.target_combo.setCurrentText("AGB")

            self.log(f"成功加载数据，共 {len(self.df)} 行记录。请在列表中勾选变量和模型。")
            self.result_banner.setText("数据已加载，请在左侧配置变量和多模型后点击【开始模型拟合与验证】。")

        except Exception as e:
            QMessageBox.critical(self, "读取错误", f"无法读取文件:\n{str(e)}")

    def select_output_file(self):
        file_filter = "CSV Files (*.csv);;Excel Files (*.xlsx)"
        filename, _ = QFileDialog.getSaveFileName(self, "选择保存路径", "fitting_results.csv", file_filter)
        if filename:
            self.output_path.setText(filename)

    def save_current_plot(self):
        if not self.plot_data:
            return

        data = self.plot_data[self.current_plot_index]
        default_name = f"{data['name']}拟合图.png"
        filename, _ = QFileDialog.getSaveFileName(self, "保存高清图表", default_name, "PNG Images (*.png)")

        if filename:
            try:
                self.figure.savefig(filename, dpi=300, bbox_inches="tight")
                self.log(f"成功导出图表至: {filename}")
                QMessageBox.information(self, "保存成功", "图表已成功保存为高清 PNG！")
            except Exception as e:
                self.log(f"导出图表失败: {str(e)}")
                QMessageBox.critical(self, "保存失败", f"无法保存图片:\n{str(e)}")

    def log(self, message):
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())

    @staticmethod
    def power_model(X, *params):
        a = params[0]
        b = np.array(params[1:])
        return a * np.prod(X ** b[:, np.newaxis], axis=0)

    def fit_model_core(self, X_train, y_train, model_idx):
        if model_idx == 0:
            log_y = np.log(y_train)
            log_x = np.log(X_train)
            X_design = np.vstack([np.ones_like(log_y), log_x]).T
            coeffs, _, _, _ = np.linalg.lstsq(X_design, log_y, rcond=None)
            p0_init = np.concatenate(([np.exp(coeffs[0])], coeffs[1:]))
            try:
                params, _ = curve_fit(self.power_model, X_train, y_train, p0=p0_init, maxfev=10000)
            except Exception:
                params = p0_init
            return params, lambda X: self.power_model(X, *params)

        if model_idx == 1:
            log_y = np.log(y_train)
            log_x = np.log(X_train)
            X_design = np.vstack([np.ones_like(log_y), log_x]).T
            coeffs, _, _, _ = np.linalg.lstsq(X_design, log_y, rcond=None)
            params = np.concatenate(([np.exp(coeffs[0])], coeffs[1:]))
            return params, lambda X: np.exp(
                np.vstack([np.ones(X.shape[1]), np.log(X)]).T.dot(np.concatenate(([np.log(params[0])], params[1:])))
            )

        if model_idx == 2:
            X_design = np.vstack([np.ones(X_train.shape[1]), X_train]).T
            coeffs, _, _, _ = np.linalg.lstsq(X_design, y_train, rcond=None)
            return coeffs, lambda X: np.vstack([np.ones(X.shape[1]), X]).T.dot(coeffs)

        if model_idx == 3:
            log_y = np.log(y_train)
            X_design = np.vstack([np.ones(X_train.shape[1]), X_train]).T
            coeffs, _, _, _ = np.linalg.lstsq(X_design, log_y, rcond=None)
            params = np.concatenate(([np.exp(coeffs[0])], coeffs[1:]))
            return params, lambda X: params[0] * np.exp(np.vstack([np.zeros(X.shape[1]), X]).T[:, 1:].dot(params[1:]))

        raise ValueError(f"不支持的模型索引: {model_idx}")

    def run_fitting(self):
        if self.df is None:
            QMessageBox.warning(self, "警告", "请先加载数据！")
            return

        target_col = self.target_combo.currentText()
        selected_features = []
        for i in range(self.feature_list.count()):
            item = self.feature_list.item(i)
            if item.checkState() == Qt.Checked:
                selected_features.append(item.text())

        selected_models = []
        for i in range(self.model_list.count()):
            item = self.model_list.item(i)
            if item.checkState() == Qt.Checked:
                selected_models.append(i)

        if not target_col or not selected_features:
            QMessageBox.warning(self, "警告", "请至少勾选一个自变量 (X)！")
            return

        if not selected_models:
            QMessageBox.warning(self, "警告", "请至少勾选一个拟合模型！")
            return

        if target_col in selected_features:
            QMessageBox.warning(self, "警告", "因变量不能同时勾选为自变量！")
            return

        output_file = self.output_path.text()
        cv_idx = self.cv_combo.currentIndex()
        model_names = ["幂函数", "对数线性", "多元线性", "指数"]

        try:
            self.log("-" * 40)
            df_clean = self.df[[target_col] + selected_features].copy()

            for col in df_clean.columns:
                df_clean[col] = pd.to_numeric(df_clean[col], errors="coerce")

            df_clean = df_clean.dropna()

            if any(m in [0, 1, 3] for m in selected_models):
                mask = (df_clean > 0).all(axis=1)
                df_clean = df_clean[mask]

            y_data = df_clean[target_col].values
            x_data = df_clean[selected_features].values.T
            n_samples = len(y_data)

            if n_samples == 0:
                self.log("【错误】清洗后数据量为0。请检查数据（非线性/对数模型要求数值必须大于0）。")
                QMessageBox.critical(self, "错误", "可用数据不足，无法拟合。")
                return

            self.log(f"用于拟合的数据点: {n_samples} | 自变量: {selected_features}")

            self.plot_data = []
            banner_equations = []
            result_df = pd.DataFrame({"Actual": y_data})
            for i, col in enumerate(selected_features):
                result_df[col] = x_data[i]

            for model_idx in selected_models:
                m_name = model_names[model_idx]
                y_pred_final = np.zeros_like(y_data)

                if cv_idx == 0:
                    params, predict_func = self.fit_model_core(x_data, y_data, model_idx)
                    y_pred_final = predict_func(x_data)
                else:
                    k = self.k_spin.value() if cv_idx == 1 else n_samples
                    if k > n_samples:
                        k = n_samples
                    cv = KFold(n_splits=k, shuffle=True, random_state=42) if cv_idx == 1 else LeaveOneOut()

                    x_t = x_data.T
                    for train_idx, test_idx in cv.split(x_t):
                        x_train, x_test = x_t[train_idx].T, x_t[test_idx].T
                        y_train = y_data[train_idx]

                        _, predict_func = self.fit_model_core(x_train, y_train, model_idx)
                        y_pred_final[test_idx] = predict_func(x_test)

                    params, _ = self.fit_model_core(x_data, y_data, model_idx)

                r_squared = r2_score(y_data, y_pred_final)
                rmse = np.sqrt(np.mean((y_data - y_pred_final) ** 2))

                if model_idx == 0:
                    eq = f"{params[0]:.4f}"
                    for i, v in enumerate(selected_features):
                        eq += f" * {v}^{params[i + 1]:.4f}"
                elif model_idx == 1:
                    eq = f"ln({params[0]:.4f})"
                    for i, v in enumerate(selected_features):
                        eq += f" + {params[i + 1]:.4f}*ln({v})"
                elif model_idx == 2:
                    eq = f"{params[0]:.4f}"
                    for i, v in enumerate(selected_features):
                        eq += f" + {params[i + 1]:.4f}*{v}"
                else:
                    eq = f"{params[0]:.4f} * e^("
                    terms = [f"{params[i + 1]:.4f}*{v}" for i, v in enumerate(selected_features)]
                    eq += " + ".join(terms) + ")"

                lhs = f"ln({target_col})" if model_idx == 1 else f"{target_col}"
                full_eq = f"{lhs} = {eq}"

                banner_equations.append(f"<b>{m_name}:</b>&nbsp;&nbsp;&nbsp;{full_eq}")

                self.plot_data.append(
                    {
                        "name": m_name,
                        "y_actual": y_data,
                        "y_pred": y_pred_final,
                        "r2": r_squared,
                        "rmse": rmse,
                        "target_col": target_col,
                    }
                )

                result_df[f"Predicted_{m_name}"] = y_pred_final
                result_df[f"Residuals_{m_name}"] = y_data - y_pred_final

                self.log(f"[{m_name}] R^2: {r_squared:.4f} | RMSE: {rmse:.4f}")

            html_text = "<div style='line-height: 1.6; font-family: Consolas, monospace;'>" + "<br>".join(
                banner_equations
            ) + "</div>"
            self.result_banner.setHtml(html_text)

            self.current_plot_index = 0
            self.update_plot_ui()
            self.btn_save_fig.setEnabled(True)

            if output_file:
                if output_file.lower().endswith(".csv"):
                    result_df.to_csv(output_file, index=False, encoding="utf-8-sig")
                else:
                    result_df.to_excel(output_file, index=False)
                self.log(f"汇总结果已保存至: {output_file}")
                QMessageBox.information(self, "完成", "拟合完毕！数据已保存，您可以滑动图表或单独保存 PNG。")
            else:
                QMessageBox.information(self, "完成", "所有选定模型拟合完毕！您可以左右滑动图表查看对比。")

        except Exception as e:
            self.log(f"发生异常: {str(e)}")
            QMessageBox.critical(self, "系统错误", f"计算异常:\n{str(e)}")

    def show_prev_plot(self):
        if self.current_plot_index > 0:
            self.current_plot_index -= 1
            self.update_plot_ui()

    def show_next_plot(self):
        if self.current_plot_index < len(self.plot_data) - 1:
            self.current_plot_index += 1
            self.update_plot_ui()

    def update_plot_ui(self):
        n_plots = len(self.plot_data)
        if n_plots == 0:
            return

        self.btn_prev.setEnabled(self.current_plot_index > 0)
        self.btn_next.setEnabled(self.current_plot_index < n_plots - 1)
        self.lbl_plot_info.setText(f"图表: {self.current_plot_index + 1} / {n_plots}")

        data = self.plot_data[self.current_plot_index]
        self.render_plot(data)

    def render_plot(self, data):
        self.figure.clear()
        ax = self.figure.add_subplot(111)

        y_actual = data["y_actual"]
        y_pred = data["y_pred"]

        ax.scatter(y_actual, y_pred, alpha=0.7, edgecolors="k", c="#2196F3", s=50)

        min_val = min(np.min(y_actual), np.min(y_pred))
        max_val = max(np.max(y_actual), np.max(y_pred))
        ax.plot([min_val, max_val], [min_val, max_val], "r--", lw=2)

        text_str = f"$R^2$ = {data['r2']:.4f}\nRMSE = {data['rmse']:.4f}"
        props = dict(boxstyle="round", facecolor="white", alpha=0.8, edgecolor="gray")
        ax.text(
            0.05,
            0.95,
            text_str,
            transform=ax.transAxes,
            fontsize=11,
            verticalalignment="top",
            bbox=props,
            fontweight="bold",
            color="#333333",
        )

        ax.set_title(data["name"], fontsize=14, fontweight="bold")
        ax.set_xlabel(f"实际值 {data['target_col']}")
        ax.set_ylabel(f"预测值 {data['target_col']}")

        ax.grid(True, linestyle=":", alpha=0.7)
        self.figure.tight_layout()
        self.canvas.draw()


# 向后兼容旧类名
AllometricFitterApp = SingleTreeStructureWindow
AllometricFitterAppWindow = SingleTreeStructureWindow


if __name__ == "__main__":
    app = QApplication(sys.argv)
    apply_qt_app_style(app)
    ex = SingleTreeStructureWindow()
    ex.show()
    sys.exit(app.exec_())


