# -*- coding: utf-8 -*-
"""
Created on Sun Mar 15 19:37:03 2026

@author: Pengzongjian
"""

# -*- coding: utf-8 -*-
"""
第8部分：界面化模块（增强版）
功能：
1. 左侧：输入参数 + 结果勾选 + 运行日志
2. 右侧：结果预览（表格 / 图片）
3. 支持按需生成结果，而不是默认全部生成
4. 支持 Excel / CSV 输入
5. 支持将已勾选的表格结果导出为 Excel

依赖：
- part1_data_loader.py
- part2_numeric_stats.py
- part3_categorical_stats.py
- part4_missing_stats.py
- part5_plot_generator.py
- part6_excel_writer.py

适用环境：Spyder / Python 3.x
界面库：tkinter
"""

import os
import glob
import queue
import threading
import traceback
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from contextlib import redirect_stdout, redirect_stderr

import pandas as pd

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

from descriptive_stats_data_loader import (
    create_output_dirs,
    read_data_file,
    clean_data,
    identify_variable_types,
    summarize_data_info
)
from descriptive_stats_numeric_stats import generate_numeric_stats
from descriptive_stats_categorical_stats import generate_categorical_stats
from descriptive_stats_missing_stats import generate_missing_stats, generate_overall_missing_summary
from descriptive_stats_plot_generator import generate_histograms, generate_boxplots, generate_bar_charts
from descriptive_stats_excel_writer import auto_adjust_column_width, apply_worksheet_style
try:
    from modules.ui_style import apply_tk_window_baseline
except ModuleNotFoundError:
    from ui_style import apply_tk_window_baseline


# ===== 图像预览：优先使用 Pillow =====
PIL_AVAILABLE = False
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False


class QueueWriter:
    """
    将 print 输出重定向到队列，供 tkinter 日志窗口实时显示
    """
    def __init__(self, log_queue):
        self.log_queue = log_queue

    def write(self, message):
        if message:
            self.log_queue.put(("log", message))

    def flush(self):
        pass


def clear_png_files(folder_path):
    """
    清空指定目录下已有的 PNG 文件，避免重复预览旧图
    """
    if not os.path.exists(folder_path):
        return

    for file_path in glob.glob(os.path.join(folder_path, "*.png")):
        try:
            os.remove(file_path)
        except Exception:
            pass


def export_selected_results_to_excel(output_excel_path, sheet_data_dict):
    """
    仅将用户勾选并生成的表格结果导出到 Excel
    参数：
        output_excel_path: 输出 Excel 路径
        sheet_data_dict: dict，键为工作表名，值为 DataFrame
    """
    if not sheet_data_dict:
        return

    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        for sheet_name, df in sheet_data_dict.items():
            if df is None:
                continue

            if df.empty:
                df = pd.DataFrame({"提示信息": [f"{sheet_name} 当前无可导出内容。"]})

            df.to_excel(writer, sheet_name=sheet_name, index=False)
            auto_adjust_column_width(writer, sheet_name, df)

        for sheet_name in writer.book.sheetnames:
            ws = writer.book[sheet_name]
            apply_worksheet_style(ws)


class DescriptiveStatsGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("描述统计分析")
        apply_tk_window_baseline(self.root)

        self.log_queue = queue.Queue()
        self.is_running = False

        self.result_tables = {}
        self.result_images = []
        self.current_preview_image = None

        self._build_widgets()
        self._poll_log_queue()

    def _build_widgets(self):
        # ===== 主容器 =====
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)

        title_label = ttk.Label(main_frame, text="描述统计分析工具", font=("Arial", 18, "bold"))
        title_label.pack(anchor="center", pady=(0, 8))

        # ===== 左右分栏 =====
        paned = ttk.PanedWindow(main_frame, orient="horizontal")
        paned.pack(fill="both", expand=True)

        left_frame = ttk.Frame(paned, padding=6)
        right_frame = ttk.Frame(paned, padding=6)

        paned.add(left_frame, weight=1)
        paned.add(right_frame, weight=1)

        # =========================================================
        # 左侧：参数设置 + 勾选项 + 按钮 + 日志
        # =========================================================
        form_frame = ttk.LabelFrame(left_frame, text="参数设置", padding=10)
        form_frame.pack(fill="x", pady=(0, 8))

        # 输入文件
        ttk.Label(form_frame, text="输入文件：").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        self.file_path_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.file_path_var, width=52).grid(
            row=0, column=1, columnspan=3, sticky="we", padx=6, pady=6
        )
        ttk.Button(form_frame, text="浏览", command=self.select_input_file).grid(
            row=0, column=4, sticky="w", padx=6, pady=6
        )

        # 输出目录
        ttk.Label(form_frame, text="输出目录：").grid(row=1, column=0, sticky="e", padx=6, pady=6)
        self.output_dir_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.output_dir_var, width=52).grid(
            row=1, column=1, columnspan=3, sticky="we", padx=6, pady=6
        )
        ttk.Button(form_frame, text="浏览", command=self.select_output_dir).grid(
            row=1, column=4, sticky="w", padx=6, pady=6
        )

        # 工作表
        ttk.Label(form_frame, text="工作表名称：").grid(row=2, column=0, sticky="e", padx=6, pady=6)
        self.sheet_name_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.sheet_name_var, width=20).grid(
            row=2, column=1, sticky="w", padx=6, pady=6
        )
        ttk.Label(
            form_frame,
            text="Excel 可填写工作表名；CSV 留空即可",
            foreground="gray"
        ).grid(row=2, column=2, columnspan=3, sticky="w", padx=6, pady=6)

        # 参数
        ttk.Label(form_frame, text="小数位数：").grid(row=3, column=0, sticky="e", padx=6, pady=6)
        self.round_digits_var = tk.StringVar(value="4")
        ttk.Entry(form_frame, textvariable=self.round_digits_var, width=10).grid(
            row=3, column=1, sticky="w", padx=6, pady=6
        )

        ttk.Label(form_frame, text="直方图分箱数：").grid(row=3, column=2, sticky="e", padx=6, pady=6)
        self.bins_var = tk.StringVar(value="10")
        ttk.Entry(form_frame, textvariable=self.bins_var, width=10).grid(
            row=3, column=3, sticky="w", padx=6, pady=6
        )

        ttk.Label(form_frame, text="图片分辨率 dpi：").grid(row=3, column=4, sticky="e", padx=6, pady=6)
        self.dpi_var = tk.StringVar(value="300")
        ttk.Entry(form_frame, textvariable=self.dpi_var, width=10).grid(
            row=3, column=5, sticky="w", padx=6, pady=6
        )

        # 列宽
        form_frame.columnconfigure(1, weight=1)
        form_frame.columnconfigure(3, weight=0)
        form_frame.columnconfigure(5, weight=0)

        # ===== 结果勾选区 =====
        option_frame = ttk.LabelFrame(left_frame, text="生成内容选择", padding=10)
        option_frame.pack(fill="x", pady=(0, 8))

        # 表格类
        ttk.Label(option_frame, text="表格结果：").grid(row=0, column=0, sticky="w", padx=6, pady=4)

        self.opt_numeric_var = tk.BooleanVar(value=True)
        self.opt_categorical_var = tk.BooleanVar(value=True)
        self.opt_missing_var = tk.BooleanVar(value=True)
        self.opt_overall_missing_var = tk.BooleanVar(value=True)

        ttk.Checkbutton(option_frame, text="数值型描述统计", variable=self.opt_numeric_var).grid(
            row=1, column=0, sticky="w", padx=12, pady=3
        )
        ttk.Checkbutton(option_frame, text="分类型频数统计", variable=self.opt_categorical_var).grid(
            row=1, column=1, sticky="w", padx=12, pady=3
        )
        ttk.Checkbutton(option_frame, text="缺失值统计", variable=self.opt_missing_var).grid(
            row=2, column=0, sticky="w", padx=12, pady=3
        )
        ttk.Checkbutton(option_frame, text="整体缺失汇总", variable=self.opt_overall_missing_var).grid(
            row=2, column=1, sticky="w", padx=12, pady=3
        )

        # 图片类
        ttk.Label(option_frame, text="图片结果：").grid(row=3, column=0, sticky="w", padx=6, pady=(10, 4))

        self.opt_hist_var = tk.BooleanVar(value=True)
        self.opt_box_var = tk.BooleanVar(value=True)
        self.opt_bar_var = tk.BooleanVar(value=True)

        ttk.Checkbutton(option_frame, text="直方图", variable=self.opt_hist_var).grid(
            row=4, column=0, sticky="w", padx=12, pady=3
        )
        ttk.Checkbutton(option_frame, text="箱线图", variable=self.opt_box_var).grid(
            row=4, column=1, sticky="w", padx=12, pady=3
        )
        ttk.Checkbutton(option_frame, text="柱状图", variable=self.opt_bar_var).grid(
            row=5, column=0, sticky="w", padx=12, pady=3
        )

        # 导出控制
        ttk.Label(option_frame, text="导出设置：").grid(row=6, column=0, sticky="w", padx=6, pady=(10, 4))
        self.opt_export_excel_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(option_frame, text="导出已勾选表格结果到 Excel", variable=self.opt_export_excel_var).grid(
            row=7, column=0, columnspan=2, sticky="w", padx=12, pady=3
        )

        # 快捷按钮
        quick_btn_frame = ttk.Frame(option_frame)
        quick_btn_frame.grid(row=8, column=0, columnspan=3, sticky="w", pady=(8, 0))

        ttk.Button(quick_btn_frame, text="全部勾选", command=self.select_all_options).pack(side="left", padx=4)
        ttk.Button(quick_btn_frame, text="全部取消", command=self.clear_all_options).pack(side="left", padx=4)
        ttk.Button(quick_btn_frame, text="仅表格结果", command=self.select_table_options).pack(side="left", padx=4)
        ttk.Button(quick_btn_frame, text="仅图片结果", command=self.select_image_options).pack(side="left", padx=4)

        # ===== 按钮区 =====
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill="x", pady=(0, 8))

        self.run_button = ttk.Button(btn_frame, text="开始运行", command=self.start_pipeline)
        self.run_button.pack(side="left", padx=4)

        ttk.Button(btn_frame, text="清空日志", command=self.clear_log).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="填入当前示例路径", command=self.fill_demo_paths).pack(side="left", padx=4)

        # ===== 日志区 =====
        log_frame = ttk.LabelFrame(left_frame, text="运行日志", padding=8)
        log_frame.pack(fill="both", expand=True)

        self.log_text = ScrolledText(log_frame, wrap="word", font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True)
        self.log_text.insert("end", "程序已启动，请设置参数并勾选要生成的结果。\n")
        self.log_text.configure(state="disabled")

        # =========================================================
        # 右侧：结果预览
        # =========================================================
        result_frame = ttk.LabelFrame(right_frame, text="结果预览", padding=8)
        result_frame.pack(fill="both", expand=True)

        self.result_notebook = ttk.Notebook(result_frame)
        self.result_notebook.pack(fill="both", expand=True)

        # ===== 表格预览页 =====
        table_tab = ttk.Frame(self.result_notebook)
        self.result_notebook.add(table_tab, text="表格结果")

        table_top_frame = ttk.Frame(table_tab)
        table_top_frame.pack(fill="x", pady=(0, 6))

        ttk.Label(table_top_frame, text="结果表：").pack(side="left", padx=4)

        self.table_selector_var = tk.StringVar()
        self.table_selector = ttk.Combobox(
            table_top_frame,
            textvariable=self.table_selector_var,
            state="readonly",
            width=28
        )
        self.table_selector.pack(side="left", padx=4)
        self.table_selector.bind("<<ComboboxSelected>>", self.on_table_selected)

        # Treeview 表格
        table_tree_frame = ttk.Frame(table_tab)
        table_tree_frame.pack(fill="both", expand=True)

        self.table_tree = ttk.Treeview(table_tree_frame, show="headings")
        self.table_tree.pack(side="left", fill="both", expand=True)

        table_y_scroll = ttk.Scrollbar(table_tree_frame, orient="vertical", command=self.table_tree.yview)
        table_y_scroll.pack(side="right", fill="y")
        self.table_tree.configure(yscrollcommand=table_y_scroll.set)

        # ===== 图片预览页 =====
        image_tab = ttk.Frame(self.result_notebook)
        self.result_notebook.add(image_tab, text="图片结果")

        image_paned = ttk.PanedWindow(image_tab, orient="horizontal")
        image_paned.pack(fill="both", expand=True)

        image_list_frame = ttk.Frame(image_paned, padding=4)
        image_preview_frame = ttk.Frame(image_paned, padding=4)

        image_paned.add(image_list_frame, weight=1)
        image_paned.add(image_preview_frame, weight=3)

        ttk.Label(image_list_frame, text="已生成图片：").pack(anchor="w", padx=4, pady=(0, 4))

        self.image_listbox = tk.Listbox(image_list_frame, height=20)
        self.image_listbox.pack(fill="both", expand=True, padx=4, pady=4)
        self.image_listbox.bind("<<ListboxSelect>>", self.on_image_selected)

        ttk.Label(image_preview_frame, text="图片预览：").pack(anchor="w", padx=4, pady=(0, 4))

        self.image_preview_label = ttk.Label(
            image_preview_frame,
            text="暂无图片可预览",
            anchor="center",
            relief="solid"
        )
        self.image_preview_label.pack(fill="both", expand=True, padx=4, pady=4)

        # ===== 底部状态栏 =====
        self.status_var = tk.StringVar(value="状态：待运行")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(fill="x", pady=(8, 0))

    # ==============================
    # 左侧功能
    # ==============================
    def select_input_file(self):
        file_path = filedialog.askopenfilename(
            title="选择输入数据文件",
            filetypes=[
                ("数据文件", "*.xlsx *.xls *.csv"),
                ("Excel 文件", "*.xlsx *.xls"),
                ("CSV 文件", "*.csv"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self.file_path_var.set(file_path)

    def select_output_dir(self):
        folder_path = filedialog.askdirectory(title="选择输出目录")
        if folder_path:
            self.output_dir_var.set(folder_path)

    def fill_demo_paths(self):
        self.file_path_var.set(r"C:\Users\彭宗健\Desktop\软件设置\试验数据.xlsx")
        self.output_dir_var.set(r"C:\Users\彭宗健\Desktop\软件设置\输出结果\界面化输出结果")
        self.sheet_name_var.set("")
        self.round_digits_var.set("4")
        self.bins_var.set("10")
        self.dpi_var.set("300")

    def select_all_options(self):
        self.opt_numeric_var.set(True)
        self.opt_categorical_var.set(True)
        self.opt_missing_var.set(True)
        self.opt_overall_missing_var.set(True)
        self.opt_hist_var.set(True)
        self.opt_box_var.set(True)
        self.opt_bar_var.set(True)
        self.opt_export_excel_var.set(True)

    def clear_all_options(self):
        self.opt_numeric_var.set(False)
        self.opt_categorical_var.set(False)
        self.opt_missing_var.set(False)
        self.opt_overall_missing_var.set(False)
        self.opt_hist_var.set(False)
        self.opt_box_var.set(False)
        self.opt_bar_var.set(False)
        self.opt_export_excel_var.set(False)

    def select_table_options(self):
        self.opt_numeric_var.set(True)
        self.opt_categorical_var.set(True)
        self.opt_missing_var.set(True)
        self.opt_overall_missing_var.set(True)
        self.opt_hist_var.set(False)
        self.opt_box_var.set(False)
        self.opt_bar_var.set(False)
        self.opt_export_excel_var.set(True)

    def select_image_options(self):
        self.opt_numeric_var.set(False)
        self.opt_categorical_var.set(False)
        self.opt_missing_var.set(False)
        self.opt_overall_missing_var.set(False)
        self.opt_hist_var.set(True)
        self.opt_box_var.set(True)
        self.opt_bar_var.set(True)
        self.opt_export_excel_var.set(False)

    def clear_log(self):
        if self.is_running:
            messagebox.showwarning("提示", "程序运行中，暂不建议清空日志。")
            return
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.insert("end", "日志已清空。\n")
        self.log_text.configure(state="disabled")

    def append_log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def validate_inputs(self):
        file_path = self.file_path_var.get().strip()
        output_dir = self.output_dir_var.get().strip()
        sheet_name = self.sheet_name_var.get().strip()

        if not file_path:
            messagebox.showerror("错误", "请输入或选择输入文件。")
            return None

        if not os.path.exists(file_path):
            messagebox.showerror("错误", f"输入文件不存在：\n{file_path}")
            return None

        ext = os.path.splitext(file_path)[1].lower()
        if ext not in [".xlsx", ".xls", ".csv"]:
            messagebox.showerror("错误", "仅支持 .xlsx、.xls、.csv 文件。")
            return None

        if not output_dir:
            messagebox.showerror("错误", "请输入或选择输出目录。")
            return None

        try:
            round_digits = int(self.round_digits_var.get().strip())
            bins = int(self.bins_var.get().strip())
            dpi = int(self.dpi_var.get().strip())
        except ValueError:
            messagebox.showerror("错误", "小数位数、分箱数、dpi 必须为整数。")
            return None

        if round_digits < 0:
            messagebox.showerror("错误", "小数位数不能小于 0。")
            return None
        if bins <= 0:
            messagebox.showerror("错误", "直方图分箱数必须大于 0。")
            return None
        if dpi <= 0:
            messagebox.showerror("错误", "图片分辨率 dpi 必须大于 0。")
            return None

        # 检查是否至少选了一个结果
        output_selected = any([
            self.opt_numeric_var.get(),
            self.opt_categorical_var.get(),
            self.opt_missing_var.get(),
            self.opt_overall_missing_var.get(),
            self.opt_hist_var.get(),
            self.opt_box_var.get(),
            self.opt_bar_var.get()
        ])
        if not output_selected:
            messagebox.showerror("错误", "请至少勾选一个要生成的结果。")
            return None

        if ext == ".csv":
            sheet_name = None
        else:
            if sheet_name == "":
                sheet_name = None

        return {
            "file_path": file_path,
            "output_dir": output_dir,
            "sheet_name": sheet_name,
            "round_digits": round_digits,
            "bins": bins,
            "dpi": dpi,
            "opt_numeric": self.opt_numeric_var.get(),
            "opt_categorical": self.opt_categorical_var.get(),
            "opt_missing": self.opt_missing_var.get(),
            "opt_overall_missing": self.opt_overall_missing_var.get(),
            "opt_hist": self.opt_hist_var.get(),
            "opt_box": self.opt_box_var.get(),
            "opt_bar": self.opt_bar_var.get(),
            "opt_export_excel": self.opt_export_excel_var.get()
        }

    def start_pipeline(self):
        if self.is_running:
            messagebox.showinfo("提示", "程序正在运行，请稍候。")
            return

        params = self.validate_inputs()
        if params is None:
            return

        self.clear_preview_area()

        self.is_running = True
        self.run_button.config(state="disabled")
        self.status_var.set("状态：运行中...")

        self.append_log("\n" + "=" * 70 + "\n")
        self.append_log("开始执行描述统计流程...\n")

        worker = threading.Thread(
            target=self.run_pipeline_in_thread,
            args=(params,),
            daemon=True
        )
        worker.start()

    # ==============================
    # 右侧结果预览
    # ==============================
    def clear_preview_area(self):
        self.result_tables = {}
        self.result_images = []
        self.current_preview_image = None

        self.table_selector["values"] = []
        self.table_selector_var.set("")

        self.table_tree.delete(*self.table_tree.get_children())
        self.table_tree["columns"] = ()

        self.image_listbox.delete(0, "end")
        self.image_preview_label.configure(text="暂无图片可预览", image="")

    def update_table_preview(self, table_dict):
        self.result_tables = table_dict if table_dict else {}

        sheet_names = list(self.result_tables.keys())
        self.table_selector["values"] = sheet_names

        if sheet_names:
            self.table_selector_var.set(sheet_names[0])
            self.display_dataframe(self.result_tables[sheet_names[0]])
        else:
            self.table_selector_var.set("")
            self.table_tree.delete(*self.table_tree.get_children())
            self.table_tree["columns"] = ()

    def update_image_preview_list(self, image_paths):
        self.result_images = image_paths if image_paths else []
        self.image_listbox.delete(0, "end")

        for img_path in self.result_images:
            self.image_listbox.insert("end", os.path.basename(img_path))

        if not self.result_images:
            self.image_preview_label.configure(text="暂无图片可预览", image="")
        else:
            self.image_listbox.selection_set(0)
            self.show_image_preview(self.result_images[0])

    def on_table_selected(self, event=None):
        sheet_name = self.table_selector_var.get()
        if sheet_name in self.result_tables:
            self.display_dataframe(self.result_tables[sheet_name])

    def display_dataframe(self, df):
        # 清空旧表
        self.table_tree.delete(*self.table_tree.get_children())

        if df is None or df.empty:
            self.table_tree["columns"] = ("提示信息",)
            self.table_tree.heading("提示信息", text="提示信息")
            self.table_tree.column("提示信息", width=400, anchor="center")
            self.table_tree.insert("", "end", values=("当前结果为空",))
            return

        columns = list(df.columns)
        self.table_tree["columns"] = columns

        for col in columns:
            self.table_tree.heading(col, text=col)

            # 简单估算列宽
            max_len = max(
                [len(str(col))] + [len(str(v)) for v in df[col].fillna("").astype(str).tolist()]
            )
            width = min(max(80, max_len * 12), 220)
            self.table_tree.column(col, width=width, anchor="center")

        for _, row in df.iterrows():
            values = ["" if pd.isna(v) else str(v) for v in row.tolist()]
            self.table_tree.insert("", "end", values=values)

    def on_image_selected(self, event=None):
        selection = self.image_listbox.curselection()
        if not selection:
            return

        index = selection[0]
        if 0 <= index < len(self.result_images):
            self.show_image_preview(self.result_images[index])

    def show_image_preview(self, image_path):
        if not os.path.exists(image_path):
            self.image_preview_label.configure(text="图片不存在，无法预览", image="")
            return

        try:
            if PIL_AVAILABLE:
                img = Image.open(image_path)
                img.thumbnail((620, 620))
                tk_img = ImageTk.PhotoImage(img)
                self.current_preview_image = tk_img
                self.image_preview_label.configure(image=tk_img, text="")
            else:
                tk_img = tk.PhotoImage(file=image_path)
                self.current_preview_image = tk_img
                self.image_preview_label.configure(image=tk_img, text="")
        except Exception as e:
            self.image_preview_label.configure(text=f"图片预览失败：{e}", image="")

    # ==============================
    # 后台执行
    # ==============================
    def run_pipeline_in_thread(self, params):
        writer = QueueWriter(self.log_queue)

        try:
            with redirect_stdout(writer), redirect_stderr(writer):
                # 1. 创建输出目录
                create_output_dirs(params["output_dir"])

                # 若勾选图片结果，先清空对应目录中的旧 PNG
                if params["opt_hist"]:
                    clear_png_files(os.path.join(params["output_dir"], "直方图"))
                if params["opt_box"]:
                    clear_png_files(os.path.join(params["output_dir"], "箱线图"))
                if params["opt_bar"]:
                    clear_png_files(os.path.join(params["output_dir"], "柱状图"))

                # 2. 读取数据
                df, file_name = read_data_file(params["file_path"], sheet_name=params["sheet_name"])

                # 3. 数据预处理
                df = clean_data(df)

                # 4. 识别变量类型
                numeric_cols, categorical_cols = identify_variable_types(df)

                # 5. 输出基本信息
                summarize_data_info(df, numeric_cols, categorical_cols)

                # 6. 按需生成结果
                result_tables = {}
                generated_image_paths = []
                output_excel_path = None

                # ===== 表格类 =====
                if params["opt_numeric"]:
                    print("正在生成数值型描述统计结果...")
                    numeric_stats_df = generate_numeric_stats(df, numeric_cols, round_digits=params["round_digits"])
                    result_tables["数值型描述统计"] = numeric_stats_df

                if params["opt_categorical"]:
                    print("正在生成分类型频数统计结果...")
                    categorical_stats_df = generate_categorical_stats(df, categorical_cols, round_digits=params["round_digits"])
                    result_tables["分类型频数统计"] = categorical_stats_df

                if params["opt_missing"]:
                    print("正在生成缺失值统计结果...")
                    missing_stats_df = generate_missing_stats(
                        df,
                        numeric_cols=numeric_cols,
                        categorical_cols=categorical_cols,
                        round_digits=params["round_digits"]
                    )
                    result_tables["缺失值统计"] = missing_stats_df

                if params["opt_overall_missing"]:
                    print("正在生成整体缺失汇总结果...")
                    overall_missing_summary_df = generate_overall_missing_summary(df, round_digits=params["round_digits"])
                    result_tables["整体缺失汇总"] = overall_missing_summary_df

                # ===== 图片类 =====
                if params["opt_hist"]:
                    print("正在生成直方图...")
                    generate_histograms(df, numeric_cols, params["output_dir"], bins=params["bins"], dpi=params["dpi"])
                    hist_files = sorted(glob.glob(os.path.join(params["output_dir"], "直方图", "*.png")))
                    generated_image_paths.extend(hist_files)

                if params["opt_box"]:
                    print("正在生成箱线图...")
                    generate_boxplots(df, numeric_cols, params["output_dir"], dpi=params["dpi"])
                    box_files = sorted(glob.glob(os.path.join(params["output_dir"], "箱线图", "*.png")))
                    generated_image_paths.extend(box_files)

                if params["opt_bar"]:
                    print("正在生成柱状图...")
                    generate_bar_charts(df, categorical_cols, params["output_dir"], dpi=params["dpi"])
                    bar_files = sorted(glob.glob(os.path.join(params["output_dir"], "柱状图", "*.png")))
                    generated_image_paths.extend(bar_files)

                # ===== 按需导出 Excel =====
                if params["opt_export_excel"]:
                    if result_tables:
                        print("正在导出 Excel 结果文件...")
                        excel_output_dir = os.path.join(params["output_dir"], "统计结果Excel")
                        output_excel_path = os.path.join(excel_output_dir, f"{file_name}_描述统计结果.xlsx")
                        export_selected_results_to_excel(output_excel_path, result_tables)
                        print(f"Excel 结果文件已导出：{output_excel_path}")
                    else:
                        print("当前未勾选任何表格结果，已跳过 Excel 导出。")

                # 7. 结果回传到界面
                result_payload = {
                    "file_name": file_name,
                    "result_tables": result_tables,
                    "generated_image_paths": generated_image_paths,
                    "output_excel_path": output_excel_path,
                    "numeric_cols": numeric_cols,
                    "categorical_cols": categorical_cols
                }
                self.log_queue.put(("result", result_payload))
                self.log_queue.put(("log", "\n界面任务执行成功。\n"))
                self.log_queue.put(("log", f"输出目录：{params['output_dir']}\n"))
                self.log_queue.put(("done", None))

        except Exception:
            error_msg = traceback.format_exc()
            self.log_queue.put(("log", "\n程序运行失败，错误信息如下：\n"))
            self.log_queue.put(("log", error_msg + "\n"))
            self.log_queue.put(("failed", None))

    def _poll_log_queue(self):
        try:
            while True:
                item_type, item_value = self.log_queue.get_nowait()

                if item_type == "log":
                    self.append_log(item_value)

                elif item_type == "result":
                    self.update_table_preview(item_value.get("result_tables", {}))
                    self.update_image_preview_list(item_value.get("generated_image_paths", []))

                elif item_type == "done":
                    self.is_running = False
                    self.run_button.config(state="normal")
                    self.status_var.set("状态：运行完成")

                elif item_type == "failed":
                    self.is_running = False
                    self.run_button.config(state="normal")
                    self.status_var.set("状态：运行失败")

        except queue.Empty:
            pass

        self.root.after(100, self._poll_log_queue)


def main():
    root = tk.Tk()

    style = ttk.Style()

    # 优先使用 Windows 原生主题，使勾选框显示为“√”而不是“×”
    preferred_themes = ["vista", "xpnative", "default", "alt", "clam"]
    available_themes = style.theme_names()

    for theme in preferred_themes:
        if theme in available_themes:
            try:
                style.theme_use(theme)
                print(f"当前界面主题：{theme}")
                break
            except Exception:
                continue

    app = DescriptiveStatsGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()