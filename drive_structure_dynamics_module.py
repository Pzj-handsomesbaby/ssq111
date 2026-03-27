# -*- coding: utf-8 -*-
"""驱动分析模块：结构动态与演替分析（PLS-SEM）。

该文件由 `wy/SEM.py` 彻底迁移到 `modules` 目录，
可被主程序通过子进程直接启动。
"""

import os
import math
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
try:
    from modules.ui_style import apply_tk_window_baseline
except ModuleNotFoundError:
    from ui_style import apply_tk_window_baseline

PLSPM_IMPORT_ERROR = None
try:
    import plspm.config as c
    from plspm.plspm import Plspm
    from plspm.scheme import Scheme
    from plspm.mode import Mode
except Exception as exc:
    PLSPM_IMPORT_ERROR = exc
    c = None
    Plspm = None
    Scheme = None
    Mode = None


class DriveStructureDynamicsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("结构动态与演替分析")
        apply_tk_window_baseline(self.root)

        self.csv_path = tk.StringVar()
        self.output_dir = tk.StringVar()

        self.seed_var = tk.StringVar(value="123")
        self.target_var = tk.StringVar()
        self.scheme_var = tk.StringVar(value="PATH")

        self.drop_na_var = tk.BooleanVar(value=True)
        self.remove_outlier_var = tk.BooleanVar(value=True)
        self.scale_predictor_var = tk.BooleanVar(value=True)

        self.boot_val_var = tk.BooleanVar(value=True)
        self.boot_iter_var = tk.StringVar(value="1000")

        self.all_columns = []
        self.blocks = {}
        self.paths = []

        self.selected_block_name = tk.StringVar()

        self.build_ui()

    def build_ui(self):
        self.build_top_panel()
        self.build_middle_panel()
        self.build_bottom_panel()

    def build_top_panel(self):
        frame = tk.LabelFrame(self.root, text="文件与参数设置", padx=10, pady=10)
        frame.pack(fill="x", padx=10, pady=8)

        tk.Label(frame, text="输入 CSV").grid(row=0, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.csv_path, width=86).grid(row=0, column=1, padx=5, sticky="we")
        tk.Button(frame, text="选择文件", width=12, command=self.choose_csv).grid(row=0, column=2, padx=5)
        tk.Button(frame, text="读取列名", width=12, command=self.load_columns).grid(row=0, column=3, padx=5)

        tk.Label(frame, text="输出目录").grid(row=1, column=0, sticky="w", pady=(8, 0))
        tk.Entry(frame, textvariable=self.output_dir, width=86).grid(row=1, column=1, padx=5, pady=(8, 0), sticky="we")
        tk.Button(frame, text="选择目录", width=12, command=self.choose_output_dir).grid(
            row=1, column=2, padx=5, pady=(8, 0)
        )
        tk.Button(frame, text="打开目录", width=12, command=self.open_output_dir).grid(
            row=1, column=3, padx=5, pady=(8, 0)
        )

        tk.Label(frame, text="随机种子").grid(row=2, column=0, sticky="w", pady=(8, 0))
        tk.Entry(frame, textvariable=self.seed_var, width=12).grid(row=2, column=1, sticky="w", padx=5, pady=(8, 0))

        tk.Label(frame, text="目标变量").grid(row=2, column=2, sticky="e", padx=(10, 2), pady=(8, 0))
        self.target_combo = ttk.Combobox(frame, textvariable=self.target_var, state="readonly", width=28)
        self.target_combo.grid(row=2, column=3, sticky="w", pady=(8, 0))

        tk.Label(frame, text="Scheme").grid(row=3, column=0, sticky="w", pady=(8, 0))
        self.scheme_combo = ttk.Combobox(
            frame,
            textvariable=self.scheme_var,
            state="readonly",
            values=["PATH", "CENTROID", "FACTORIAL"],
            width=12,
        )
        self.scheme_combo.grid(row=3, column=1, sticky="w", padx=5, pady=(8, 0))

        tk.Checkbutton(frame, text="删除缺失值/非有限值", variable=self.drop_na_var).grid(
            row=3, column=2, sticky="w", padx=5, pady=(8, 0)
        )
        tk.Checkbutton(frame, text="3×SD 剔除异常值", variable=self.remove_outlier_var).grid(
            row=3, column=3, sticky="w", padx=5, pady=(8, 0)
        )

        tk.Checkbutton(frame, text="开启 Bootstrap", variable=self.boot_val_var).grid(
            row=4, column=0, sticky="w", pady=(8, 0)
        )
        boot_frame = tk.Frame(frame)
        boot_frame.grid(row=4, column=1, sticky="w", padx=5, pady=(8, 0))
        tk.Label(boot_frame, text="Bootstrap 次数:").pack(side="left")
        tk.Entry(boot_frame, textvariable=self.boot_iter_var, width=10).pack(side="left", padx=5)

        tk.Checkbutton(frame, text="自变量标准化", variable=self.scale_predictor_var).grid(
            row=4, column=3, sticky="w", padx=5, pady=(8, 0)
        )

        frame.columnconfigure(1, weight=1)

    def build_middle_panel(self):
        middle = tk.Frame(self.root)
        middle.pack(fill="both", expand=True, padx=10, pady=8)

        left = tk.LabelFrame(middle, text="变量选择", padx=8, pady=8)
        left.pack(side="left", fill="both", expand=True, padx=(0, 6))

        tk.Label(left, text="数据列").pack(anchor="w")
        self.var_listbox = tk.Listbox(left, selectmode=tk.MULTIPLE, width=35, height=24, exportselection=False)
        self.var_listbox.pack(fill="both", expand=True, pady=5)

        var_btn = tk.Frame(left)
        var_btn.pack(fill="x", pady=5)
        tk.Button(var_btn, text="全选", width=10, command=self.select_all_vars).pack(side="left", padx=3)
        tk.Button(var_btn, text="清空选择", width=10, command=self.clear_var_selection).pack(side="left", padx=3)

        center = tk.LabelFrame(middle, text="潜变量配置", padx=8, pady=8)
        center.pack(side="left", fill="both", expand=True, padx=6)

        top_row = tk.Frame(center)
        top_row.pack(fill="x", pady=4)
        tk.Label(top_row, text="潜变量名称").pack(side="left")
        self.latent_name_entry = tk.Entry(top_row, width=24)
        self.latent_name_entry.pack(side="left", padx=5)
        tk.Button(top_row, text="新增潜变量", width=12, command=self.add_latent).pack(side="left", padx=5)

        row2 = tk.Frame(center)
        row2.pack(fill="x", pady=4)
        tk.Label(row2, text="当前潜变量").pack(side="left")
        self.block_combo = ttk.Combobox(row2, textvariable=self.selected_block_name, state="readonly", width=26)
        self.block_combo.pack(side="left", padx=5)
        tk.Button(row2, text="加入选中变量", width=14, command=self.add_selected_vars_to_block).pack(side="left", padx=5)
        tk.Button(row2, text="删除潜变量", width=12, command=self.delete_block).pack(side="left", padx=5)

        tk.Label(center, text="潜变量及其指标").pack(anchor="w", pady=(8, 2))
        self.block_listbox = tk.Listbox(center, width=50, height=20, exportselection=False)
        self.block_listbox.pack(fill="both", expand=True, pady=5)

        block_btn = tk.Frame(center)
        block_btn.pack(fill="x", pady=4)
        tk.Button(block_btn, text="移除所选指标", width=14, command=self.remove_selected_indicator).pack(side="left", padx=3)
        tk.Button(block_btn, text="刷新显示", width=10, command=self.refresh_block_display).pack(side="left", padx=3)

        right = tk.LabelFrame(middle, text="路径关系配置", padx=8, pady=8)
        right.pack(side="left", fill="both", expand=True, padx=(6, 0))

        path_row = tk.Frame(right)
        path_row.pack(fill="x", pady=4)
        tk.Label(path_row, text="起点").pack(side="left")
        self.path_from_combo = ttk.Combobox(path_row, state="readonly", width=18)
        self.path_from_combo.pack(side="left", padx=5)
        tk.Label(path_row, text="终点").pack(side="left")
        self.path_to_combo = ttk.Combobox(path_row, state="readonly", width=18)
        self.path_to_combo.pack(side="left", padx=5)
        tk.Button(path_row, text="添加路径", width=10, command=self.add_path).pack(side="left", padx=5)

        tk.Label(right, text="当前路径").pack(anchor="w", pady=(8, 2))
        self.path_listbox = tk.Listbox(right, width=45, height=20, exportselection=False)
        self.path_listbox.pack(fill="both", expand=True, pady=5)

        path_btn = tk.Frame(right)
        path_btn.pack(fill="x", pady=4)
        tk.Button(path_btn, text="删除所选路径", width=14, command=self.remove_selected_path).pack(side="left", padx=3)
        tk.Button(path_btn, text="刷新显示", width=10, command=self.refresh_path_display).pack(side="left", padx=3)

    def build_bottom_panel(self):
        frame = tk.LabelFrame(self.root, text="运行与日志", padx=8, pady=8)
        frame.pack(fill="both", expand=False, padx=10, pady=8)

        btn_frame = tk.Frame(frame)
        btn_frame.pack(fill="x", pady=4)

        tk.Button(btn_frame, text="运行模型", width=14, height=2, bg="#2d6cdf", fg="white", command=self.run_model).pack(
            side="left", padx=4
        )

        tk.Button(btn_frame, text="生成默认示例", width=14, height=2, command=self.load_default_sem).pack(side="left", padx=4)

        tk.Button(btn_frame, text="清空日志", width=10, height=2, command=self.clear_log).pack(side="left", padx=4)

        self.log_text = tk.Text(frame, height=12, bg="black", fg="#00ff66")
        self.log_text.pack(fill="both", expand=True, pady=5)

    def log(self, msg):
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.root.update_idletasks()

    def clear_log(self):
        self.log_text.delete("1.0", "end")

    def choose_csv(self):
        path = filedialog.askopenfilename(title="选择 CSV 文件", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if path:
            self.csv_path.set(path)
            if not self.output_dir.get().strip():
                self.output_dir.set(os.path.join(os.path.dirname(path), "PLS_PM_results"))

    def choose_output_dir(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir.set(path)

    def open_output_dir(self):
        path = self.output_dir.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择输出目录。")
            return
        if not os.path.exists(path):
            messagebox.showwarning("提示", "输出目录不存在。")
            return
        os.startfile(path)

    def load_columns(self):
        csv_path = self.csv_path.get().strip()
        if not csv_path:
            messagebox.showwarning("提示", "请先选择 CSV 文件。")
            return
        if not os.path.exists(csv_path):
            messagebox.showerror("错误", "CSV 文件不存在。")
            return
        try:
            try:
                df = pd.read_csv(csv_path, nrows=5, encoding="utf-8")
            except UnicodeDecodeError:
                df = pd.read_csv(csv_path, nrows=5, encoding="gbk")

            self.all_columns = list(df.columns)
            self.refresh_var_list()
            self.target_combo["values"] = self.all_columns
            if "CZ" in self.all_columns:
                self.target_var.set("CZ")
            elif self.all_columns:
                self.target_var.set(self.all_columns[0])
            self.log(f"已读取列名，共 {len(self.all_columns)} 个变量。")
        except Exception as exc:
            messagebox.showerror("错误", f"读取列名失败：{exc}")

    def refresh_var_list(self):
        self.var_listbox.delete(0, "end")
        for col in self.all_columns:
            self.var_listbox.insert("end", col)

    def select_all_vars(self):
        self.var_listbox.select_set(0, "end")

    def clear_var_selection(self):
        self.var_listbox.selection_clear(0, "end")

    def add_latent(self):
        name = self.latent_name_entry.get().strip()
        if not name:
            messagebox.showwarning("提示", "请输入潜变量名称。")
            return
        if name in self.blocks:
            messagebox.showwarning("提示", "该潜变量已存在。")
            return
        self.blocks[name] = []
        self.latent_name_entry.delete(0, "end")
        self.refresh_block_controls()
        self.log(f"已新增潜变量：{name}")

    def refresh_block_controls(self):
        names = list(self.blocks.keys())
        self.block_combo["values"] = names
        self.path_from_combo["values"] = names
        self.path_to_combo["values"] = names
        if names and not self.selected_block_name.get():
            self.selected_block_name.set(names[0])
        self.refresh_block_display()
        self.refresh_path_display()

    def add_selected_vars_to_block(self):
        block_name = self.selected_block_name.get().strip()
        if not block_name:
            messagebox.showwarning("提示", "请先选择潜变量。")
            return
        idxs = self.var_listbox.curselection()
        if not idxs:
            messagebox.showwarning("提示", "请先在左侧选择变量。")
            return

        selected_vars = [self.var_listbox.get(i) for i in idxs]
        old_vars = self.blocks.get(block_name, [])

        for val in selected_vars:
            if val not in old_vars:
                old_vars.append(val)

        self.blocks[block_name] = old_vars
        self.refresh_block_display()
        self.log(f"已将变量加入潜变量 {block_name}: {', '.join(selected_vars)}")

    def refresh_block_display(self):
        self.block_listbox.delete(0, "end")
        for block_name, vars_ in self.blocks.items():
            text = f"{block_name} = {', '.join(vars_)}" if vars_ else f"{block_name} = "
            self.block_listbox.insert("end", text)

    def delete_block(self):
        block_name = self.selected_block_name.get().strip()
        if not block_name:
            messagebox.showwarning("提示", "请先选择潜变量。")
            return
        if block_name not in self.blocks:
            return

        del self.blocks[block_name]
        self.paths = [(a, b) for a, b in self.paths if a != block_name and b != block_name]
        names = list(self.blocks.keys())
        self.selected_block_name.set(names[0] if names else "")
        self.refresh_block_controls()
        self.log(f"已删除潜变量：{block_name}")

    def remove_selected_indicator(self):
        sel = self.block_listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请先在中间列表中选择一行。")
            return

        line = self.block_listbox.get(sel[0])
        if "=" not in line:
            return

        block_name = line.split("=", 1)[0].strip()
        if block_name not in self.blocks or not self.blocks[block_name]:
            return

        win = tk.Toplevel(self.root)
        win.title(f"移除指标 - {block_name}")
        win.geometry("360x320")

        lb = tk.Listbox(win, selectmode=tk.MULTIPLE, exportselection=False)
        lb.pack(fill="both", expand=True, padx=10, pady=10)
        for val in self.blocks[block_name]:
            lb.insert("end", val)

        def confirm_remove():
            idxs = lb.curselection()
            if not idxs:
                return
            remove_vars = [lb.get(i) for i in idxs]
            self.blocks[block_name] = [v for v in self.blocks[block_name] if v not in remove_vars]
            self.refresh_block_display()
            self.log(f"已从 {block_name} 中移除：{', '.join(remove_vars)}")
            win.destroy()

        tk.Button(win, text="确认移除", command=confirm_remove).pack(pady=8)

    def add_path(self):
        from_name = self.path_from_combo.get().strip()
        to_name = self.path_to_combo.get().strip()

        if not from_name or not to_name:
            messagebox.showwarning("提示", "请选择起点和终点。")
            return
        if from_name == to_name:
            messagebox.showwarning("提示", "起点和终点不能相同。")
            return
        if (from_name, to_name) in self.paths:
            messagebox.showwarning("提示", "该路径已存在。")
            return

        self.paths.append((from_name, to_name))
        self.refresh_path_display()
        self.log(f"已添加路径：{from_name} -> {to_name}")

    def refresh_path_display(self):
        self.path_listbox.delete(0, "end")
        for a, b in self.paths:
            self.path_listbox.insert("end", f"{a} -> {b}")

    def remove_selected_path(self):
        sel = self.path_listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请先选择路径。")
            return
        line = self.path_listbox.get(sel[0])
        if "->" not in line:
            return
        a, b = [x.strip() for x in line.split("->", 1)]
        if (a, b) in self.paths:
            self.paths.remove((a, b))
            self.refresh_path_display()
            self.log(f"已删除路径：{a} -> {b}")

    def load_default_sem(self):
        needed = [
            "CZ", "bio_4", "bio_3", "bio_12", "bio_13", "bio_15", "bio_16", "bio_18", "bio_2",
            "HFP", "t_clay", "t_sand", "t_gravel", "bio_17", "bio_11", "bio_7", "bio_1",
            "bio_19", "bio_9", "bio_10", "DEM",
        ]

        existing = [cname for cname in needed if (not self.all_columns or cname in self.all_columns)]

        self.blocks = {
            "topography": [v for v in ["DEM"] if v in existing],
            "climate": [
                v for v in [
                    "bio_4", "bio_3", "bio_12", "bio_13", "bio_15", "bio_16", "bio_18", "bio_2",
                    "bio_17", "bio_11", "bio_7", "bio_1", "bio_19", "bio_9", "bio_10",
                ] if v in existing
            ],
            "HFP_latent": [v for v in ["HFP"] if v in existing],
            "soil": [v for v in ["t_clay", "t_sand", "t_gravel"] if v in existing],
            "CZ_latent": [v for v in ["CZ"] if v in existing],
        }

        self.paths = [
            ("topography", "climate"),
            ("topography", "soil"),
            ("climate", "soil"),
            ("HFP_latent", "soil"),
            ("topography", "CZ_latent"),
            ("climate", "CZ_latent"),
            ("HFP_latent", "CZ_latent"),
            ("soil", "CZ_latent"),
        ]

        self.refresh_block_controls()
        self.log("已加载默认示例模型。")

    @staticmethod
    def remove_outliers_3sd(df, vars_):
        df_clean = df.copy()
        for var_name in vars_:
            x = df_clean[var_name]
            m = x.mean()
            s = x.std()
            if pd.isna(s) or s == 0:
                continue
            df_clean = df_clean[(df_clean[var_name] >= m - 3 * s) & (df_clean[var_name] <= m + 3 * s)]
        return df_clean

    def preprocess_data(self, df, selected_columns, target_var):
        data = df[selected_columns].copy()

        for col in data.columns:
            data[col] = pd.to_numeric(data[col], errors="coerce")

        if self.drop_na_var.get():
            data = data.dropna()
            data = data[np.isfinite(data).all(axis=1)]

        predictor_cols = [cname for cname in data.columns if cname != target_var]

        if self.remove_outlier_var.get():
            data = self.remove_outliers_3sd(data, predictor_cols)

        if self.drop_na_var.get():
            data = data.dropna()
            data = data[np.isfinite(data).all(axis=1)]

        if self.scale_predictor_var.get() and predictor_cols:
            for cname in predictor_cols:
                s = data[cname].std()
                if pd.notna(s) and s != 0:
                    data[cname] = (data[cname] - data[cname].mean()) / s

        return data

    def build_python_model(self, data):
        if PLSPM_IMPORT_ERROR is not None:
            raise ImportError(
                "未检测到 plspm 依赖，请先安装：pip install plspm"
            ) from PLSPM_IMPORT_ERROR

        structure = c.Structure()
        by_from = {}
        for src, dst in self.paths:
            by_from.setdefault(src, []).append(dst)
        for src, dsts in by_from.items():
            structure.add_path([src], dsts)

        config = c.Config(structure.path(), scaled=False)

        for lv_name, mvs in self.blocks.items():
            if not mvs:
                raise ValueError(f"潜变量 {lv_name} 尚未分配指标。")
            mv_objs = [c.MV(v) for v in mvs]
            config.add_lv(lv_name, Mode.A, *mv_objs)

        scheme_map = {"PATH": Scheme.PATH, "CENTROID": Scheme.CENTROID, "FACTORIAL": Scheme.FACTORIAL}
        scheme = scheme_map[self.scheme_var.get().strip().upper()]

        boot_val = self.boot_val_var.get()
        try:
            br = int(self.boot_iter_var.get().strip() or "1000")
        except ValueError:
            br = 1000

        try:
            model = Plspm(data, config, scheme, bootstrap=boot_val, bootstrap_iterations=br)
        except TypeError:
            try:
                model = Plspm(data, config, scheme, boot_val=boot_val, br=br)
            except TypeError:
                self.log("警告: 当前 plspm 版本不支持 Bootstrap 参数，将以基础模式运行。")
                model = Plspm(data, config, scheme)

        return model

    @staticmethod
    def draw_simple_path_figure(paths, out_png):
        if not paths:
            return

        nodes = []
        for a, b in paths:
            if a not in nodes:
                nodes.append(a)
            if b not in nodes:
                nodes.append(b)

        n = len(nodes)
        angle_step = 2 * math.pi / max(n, 1)
        positions = {}

        for i, node in enumerate(nodes):
            angle = i * angle_step
            positions[node] = (math.cos(angle), math.sin(angle))

        plt.figure(figsize=(10, 8))
        ax = plt.gca()
        ax.set_aspect("equal")

        for node, (x, y) in positions.items():
            circle = plt.Circle((x, y), 0.16, fill=False, linewidth=2)
            ax.add_patch(circle)
            ax.text(x, y, node, ha="center", va="center", fontsize=10)

        for a, b in paths:
            x1, y1 = positions[a]
            x2, y2 = positions[b]
            dx, dy = x2 - x1, y2 - y1
            length = math.sqrt(dx * dx + dy * dy)
            if length == 0:
                continue
            shrink = 0.18
            sx = x1 + dx * shrink / length
            sy = y1 + dy * shrink / length
            ex = x2 - dx * shrink / length
            ey = y2 - dy * shrink / length
            ax.annotate("", xy=(ex, ey), xytext=(sx, sy), arrowprops=dict(arrowstyle="->", lw=1.8))

        ax.set_xlim(-1.4, 1.4)
        ax.set_ylim(-1.4, 1.4)
        ax.axis("off")
        plt.tight_layout()
        plt.savefig(out_png, dpi=200, bbox_inches="tight")
        plt.close()

    def safe_get_attr(self, obj, attr_name, default_val):
        if hasattr(obj, attr_name):
            val = getattr(obj, attr_name)
            if callable(val):
                try:
                    return val()
                except Exception:
                    pass
            return val
        return default_val

    def run_model(self):
        try:
            csv_path = self.csv_path.get().strip()
            output_dir = self.output_dir.get().strip()
            target_var = self.target_var.get().strip()

            if not csv_path:
                raise ValueError("请先选择输入 CSV 文件。")
            if not os.path.exists(csv_path):
                raise ValueError("输入 CSV 文件不存在。")
            if not output_dir:
                raise ValueError("请选择输出目录。")
            if not self.blocks:
                raise ValueError("请至少配置一个潜变量。")
            if not target_var:
                raise ValueError("请选择目标变量。")

            os.makedirs(output_dir, exist_ok=True)

            try:
                seed = int(self.seed_var.get().strip() or "123")
            except ValueError:
                seed = 123
            np.random.seed(seed)

            used_vars = set()
            for lv_name, vars_ in self.blocks.items():
                if not vars_:
                    raise ValueError(f"潜变量 {lv_name} 尚未分配指标。")
                used_vars.update(vars_)

            selected_columns = list(used_vars)
            if target_var not in selected_columns:
                selected_columns.append(target_var)

            self.log("开始读取数据...")

            try:
                raw = pd.read_csv(csv_path, encoding="utf-8")
            except UnicodeDecodeError:
                raw = pd.read_csv(csv_path, encoding="gbk")

            missing_cols = [v for v in selected_columns if v not in raw.columns]
            if missing_cols:
                raise ValueError("以下列在数据中不存在：\n" + ", ".join(missing_cols))

            self.log(f"原始样本量：{len(raw)}")
            data = self.preprocess_data(raw, selected_columns, target_var)
            self.log(f"预处理后样本量：{len(data)}")

            if len(data) < 10:
                raise ValueError("有效样本量不足，至少建议保留 10 行以上数据。")

            self.log("开始构建 PLS-PM 模型...")
            self.log("正在执行算法及 Bootstrap（可能需要一些时间，请耐心等待）...")

            model = self.build_python_model(data)

            self.log("开始提取详细结果...")

            inner_summary = model.inner_summary()
            path_coefs = model.path_coefficients()
            inner_model = model.inner_model()
            outer_model = model.outer_model()
            effects = model.effects()
            scores = model.scores()

            unidim = self.safe_get_attr(model, "unidimensionality", pd.DataFrame())
            crossloadings = self.safe_get_attr(model, "crossloadings", pd.DataFrame())
            gof = self.safe_get_attr(model, "goodness_of_fit", "Not available in this version")

            lv_corrs = scores.corr()

            ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")

            inner_summary.to_csv(os.path.join(output_dir, f"{target_var}_inner_summary_{ts}.csv"), encoding="utf-8-sig")
            path_coefs.to_csv(os.path.join(output_dir, f"{target_var}_path_coefs_{ts}.csv"), encoding="utf-8-sig")
            inner_model.to_csv(os.path.join(output_dir, f"{target_var}_inner_model_{ts}.csv"), encoding="utf-8-sig")
            outer_model.to_csv(os.path.join(output_dir, f"{target_var}_outer_model_{ts}.csv"), encoding="utf-8-sig")
            effects.to_csv(os.path.join(output_dir, f"{target_var}_effects_{ts}.csv"), encoding="utf-8-sig", index=False)
            scores.to_csv(os.path.join(output_dir, f"{target_var}_latent_scores_{ts}.csv"), encoding="utf-8-sig")

            self.log("正在生成报告 txt...")
            summary_txt = os.path.join(output_dir, f"{target_var}_summary_{ts}.txt")
            with open(summary_txt, "w", encoding="utf-8") as f:
                f.write("PARTIAL LEAST SQUARES PATH MODELING (PLS-PM)\n\n")

                f.write("----------------------------------------------------------\n")
                f.write("MODEL SPECIFICATION\n")
                f.write(f"1   Number of Cases      {len(data)}\n")
                f.write(f"2   Latent Variables     {len(self.blocks)}\n")
                mv_count = len(used_vars)
                f.write(f"3   Manifest Variables   {mv_count}\n")
                f.write(f"4   Scale of Data        {'Standardized' if self.scale_predictor_var.get() else 'Raw Data'}\n")
                f.write("5   Non-Metric PLS       FALSE\n")
                f.write(f"6   Weighting Scheme     {self.scheme_var.get().lower()}\n")
                f.write("7   Tolerance Crit       1e-06\n")
                f.write("8   Max Num Iters        100\n")
                f.write("9   Convergence Iters    N/A\n")
                f.write(f"10  Bootstrapping        {self.boot_val_var.get()}\n")
                f.write(f"11  Bootstrap samples    {self.boot_iter_var.get() if self.boot_val_var.get() else '0'}\n\n")

                f.write("----------------------------------------------------------\n")
                f.write("BLOCKS DEFINITION\n")
                f.write(f"{'':<11} {'Block':<15} {'Type':<12} {'Size':<6} {'Mode':<6}\n")
                for i, (lv, mvs) in enumerate(self.blocks.items(), 1):
                    is_endo = any(dst == lv for src, dst in self.paths)
                    blk_type = "Endogenous" if is_endo else "Exogenous"
                    f.write(f"{i:<3} {lv:<15} {blk_type:<12} {len(mvs):<6} A\n")
                f.write("\n")

                f.write("----------------------------------------------------------\n")
                f.write("BLOCKS UNIDIMENSIONALITY\n")
                if isinstance(unidim, pd.DataFrame) and not unidim.empty:
                    f.write(unidim.to_string())
                else:
                    f.write("Unidimensionality metric not available in the current plspm package version.")
                f.write("\n\n")

                f.write("----------------------------------------------------------\n")
                f.write("OUTER MODEL\n")
                f.write(outer_model.to_string())
                f.write("\n\n")

                f.write("----------------------------------------------------------\n")
                f.write("CROSSLOADINGS\n")
                if isinstance(crossloadings, pd.DataFrame) and not crossloadings.empty:
                    f.write(crossloadings.to_string())
                else:
                    f.write("Crossloadings metric not available in the current plspm package version.")
                f.write("\n\n")

                f.write("----------------------------------------------------------\n")
                f.write("INNER MODEL\n")
                f.write(inner_model.to_string())
                f.write("\n\n")

                f.write("----------------------------------------------------------\n")
                f.write("CORRELATIONS BETWEEN LVs\n")
                f.write(lv_corrs.to_string())
                f.write("\n\n")

                f.write("----------------------------------------------------------\n")
                f.write("SUMMARY INNER MODEL\n")
                f.write(inner_summary.to_string())
                f.write("\n\n")

                f.write("----------------------------------------------------------\n")
                f.write("GOODNESS-OF-FIT\n")
                f.write(f"[1]  {gof}")
                f.write("\n\n")

                f.write("----------------------------------------------------------\n")
                f.write("TOTAL EFFECTS\n")
                f.write(effects.to_string(index=False))
                f.write("\n\n")

                if self.boot_val_var.get():
                    f.write("----------------------------------------------------------\n")
                    f.write("BOOTSTRAP VALIDATION\n")

                    try:
                        boot_obj = self.safe_get_attr(model, "bootstrap", None)
                        if boot_obj is not None:
                            if isinstance(boot_obj, dict):
                                for key, df_boot in boot_obj.items():
                                    f.write(f"\n{key}\n")
                                    f.write(df_boot.to_string())
                                    f.write("\n")
                            else:
                                try:
                                    f.write("\nweights\n")
                                    f.write(boot_obj.weights().to_string())
                                    f.write("\n")
                                except Exception:
                                    pass

                                try:
                                    f.write("\nloadings\n")
                                    f.write(boot_obj.loading().to_string())
                                    f.write("\n")
                                except Exception:
                                    pass

                                try:
                                    f.write("\npaths\n")
                                    f.write(boot_obj.paths().to_string())
                                    f.write("\n")
                                except Exception:
                                    pass

                                try:
                                    f.write("\nrsq\n")
                                    f.write(boot_obj.r_squared().to_string())
                                    f.write("\n")
                                except Exception:
                                    pass

                                try:
                                    f.write("\ntotal.efs\n")
                                    f.write(boot_obj.total_effects().to_string())
                                    f.write("\n")
                                except Exception:
                                    pass
                        else:
                            f.write("\nBootstrap validation returned None. It might have failed silently.\n")
                    except Exception as exc:
                        f.write(f"\nError trying to extract Bootstrap details: {str(exc)}\n")

            fig_png = os.path.join(output_dir, f"{target_var}_path_diagram_{ts}.png")
            self.draw_simple_path_figure(self.paths, fig_png)

            self.log("模型运行完成。")
            self.log(f"结果目录：{output_dir}")
            messagebox.showinfo("完成", f"模型运行完成。\n结果已输出到：\n{output_dir}")

        except Exception as exc:
            self.log("运行失败：")
            self.log(str(exc))
            self.log(traceback.format_exc())
            messagebox.showerror("错误", str(exc))


# 向后兼容旧类名
SEMApp = DriveStructureDynamicsApp


if __name__ == "__main__":
    if PLSPM_IMPORT_ERROR is not None:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("依赖缺失", f"无法启动结构动态与演替分析模块：\n{PLSPM_IMPORT_ERROR}\n\n请先安装：pip install plspm")
        root.destroy()
        raise SystemExit(1)

    root = tk.Tk()
    app = DriveStructureDynamicsApp(root)
    root.mainloop()

