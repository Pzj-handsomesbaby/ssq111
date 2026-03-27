# -*- coding: utf-8 -*-
"""
单木质量评价 PyQt5 小程序（最终双模式权重版）

适配的 Excel 表头：
- Tree_ID
- Tree_speice
- Y
- X
- DBH (cm)
- H (m)
- Stem quality
- Branch height (m)
- CW_S (m)
- CW_N (m)
- CW_E (m)
- CW_W (m)
- Viability
"""

import os
import sys
import traceback
import numpy as np
import pandas as pd
from scipy.spatial import distance_matrix

from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog,
    QTextEdit, QVBoxLayout, QHBoxLayout, QGridLayout, QMessageBox,
    QComboBox, QDoubleSpinBox, QSpinBox, QGroupBox
)


# =========================================================
# 一、基础工具函数
# =========================================================

def check_file_exists(file_path):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"未找到输入文件：{file_path}")


def read_excel_data(file_path):
    return pd.read_excel(file_path)


def rename_columns(df, mapping):
    missing_cols = [v for v in mapping.values() if v not in df.columns]
    if missing_cols:
        raise ValueError(f"Excel 中缺少这些列：{missing_cols}")

    reverse_map = {v: k for k, v in mapping.items()}
    return df.rename(columns=reverse_map).copy()


def convert_stem_quality(series):
    numeric_try = pd.to_numeric(series, errors="coerce")
    if numeric_try.notna().all():
        return numeric_try

    text_map = {
        "excellent": 5,
        "very good": 4.5,
        "good": 4,
        "medium": 3,
        "fair": 2.5,
        "poor": 2,
        "very poor": 1,
        "优": 5,
        "良": 4,
        "中": 3,
        "差": 2,
        "很差": 1,
    }

    s = series.astype(str).str.strip().str.lower()
    return s.map(text_map)


def convert_viability(series):
    if series is None:
        return None

    numeric_try = pd.to_numeric(series, errors="coerce")
    if numeric_try.notna().all():
        return numeric_try

    text_map = {
        "strong": 5,
        "good": 4,
        "medium": 3,
        "weak": 2,
        "poor": 1,
        "dead": 0,
        "死亡": 0,
        "died": 0,
        "die": 0,
        "强": 5,
        "中": 3,
        "弱": 1,
    }

    s = series.astype(str).str.strip().str.lower()
    return s.map(text_map)


def normalize_positive(series):
    s_min = series.min()
    s_max = series.max()
    if np.isclose(s_max, s_min):
        return pd.Series(np.full(len(series), 0.5), index=series.index)
    return (series - s_min) / (s_max - s_min)


def normalize_negative(series):
    s_min = series.min()
    s_max = series.max()
    if np.isclose(s_max, s_min):
        return pd.Series(np.full(len(series), 0.5), index=series.index)
    return (s_max - series) / (s_max - s_min)


def normalize_weights(weight_dict):
    total = sum(weight_dict.values())
    if total <= 0:
        raise ValueError("权重总和必须大于 0")
    return {k: v / total for k, v in weight_dict.items()}


# =========================================================
# 二、熵权法
# =========================================================

def calculate_entropy_weights(df, std_cols):
    data = df[std_cols].copy().astype(float)

    for col in std_cols:
        if np.isclose(data[col].sum(), 0):
            data[col] = 1e-12

    P = data.div(data.sum(axis=0), axis=1)
    P = P.replace(0, 1e-12)

    n = len(data)
    if n <= 1:
        return {col: 1.0 / len(std_cols) for col in std_cols}

    k = 1.0 / np.log(n)
    e = -k * (P * np.log(P)).sum(axis=0)
    d = 1 - e

    if np.isclose(d.sum(), 0):
        w = pd.Series(np.ones(len(std_cols)) / len(std_cols), index=std_cols)
    else:
        w = d / d.sum()

    return w.to_dict()


# =========================================================
# 三、死亡木识别与数据检查
# =========================================================

def filter_dead_trees(df, logger=print):
    dead_mask = pd.Series(False, index=df.index)

    if "viability" in df.columns:
        v = df["viability"].astype(str).str.strip().str.lower()
        dead_keywords = ["dead", "死亡", "die", "died"]
        dead_mask = dead_mask | v.isin(dead_keywords)

    if "height" in df.columns:
        h = pd.to_numeric(df["height"], errors="coerce")
        dead_mask = dead_mask | (h <= 0)

    dead_df = df.loc[dead_mask].copy()
    dead_count = int(dead_mask.sum())

    if dead_count > 0:
        logger(f"检测到死亡木或无效木 {dead_count} 株，已自动剔除，不参与单木质量评价。")
        df = df.loc[~dead_mask].copy()
    else:
        logger("未检测到死亡木，全部记录参与评价。")

    if df.empty:
        raise ValueError("剔除死亡木后，数据为空，无法继续计算。")

    return df, dead_df


def filter_invalid_crown_trees(df, logger=print):
    for col in ["cw_s", "cw_n", "cw_e", "cw_w"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["crown_width"] = (df["cw_s"] + df["cw_n"] + df["cw_e"] + df["cw_w"]) / 4.0

    invalid_mask = df["crown_width"] <= 0
    invalid_df = df.loc[invalid_mask].copy()
    invalid_count = int(invalid_mask.sum())

    if invalid_count > 0:
        bad_ids = df.loc[invalid_mask, "tree_id"].tolist()[:10]
        logger(f"检测到平均冠幅<=0 的无效木 {invalid_count} 株，已自动剔除。问题树编号示例：{bad_ids}")
        df = df.loc[~invalid_mask].copy()
    else:
        logger("平均冠幅均有效。")

    if df.empty:
        raise ValueError("剔除平均冠幅无效木后，数据为空，无法继续计算。")

    return df, invalid_df


def validate_data(df):
    required_cols = [
        "tree_id", "x", "y", "dbh", "height",
        "branch_height", "cw_s", "cw_n", "cw_e", "cw_w", "stem_quality"
    ]

    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"缺少必要字段：{col}")

    numeric_cols = ["x", "y", "dbh", "height", "branch_height", "cw_s", "cw_n", "cw_e", "cw_w"]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["stem_quality"] = convert_stem_quality(df["stem_quality"])

    if "viability" in df.columns:
        df["viability_score"] = convert_viability(df["viability"])
    else:
        df["viability_score"] = np.nan

    df["crown_width"] = (df["cw_s"] + df["cw_n"] + df["cw_e"] + df["cw_w"]) / 4.0

    check_cols = required_cols + ["crown_width"]
    if df[check_cols].isnull().any().any():
        bad_cols = df[check_cols].columns[df[check_cols].isnull().any()].tolist()
        raise ValueError(f"以下字段存在空值或无法转换：{bad_cols}")

    bad_dbh = df[df["dbh"] <= 0]
    if not bad_dbh.empty:
        raise ValueError(
            f"胸径存在 <= 0 的值，请检查。问题记录示例："
            f"{bad_dbh[['tree_id', 'dbh']].head(10).to_dict(orient='records')}"
        )

    bad_height = df[df["height"] <= 0]
    if not bad_height.empty:
        raise ValueError(
            f"树高存在 <= 0 的值，请检查。问题记录示例："
            f"{bad_height[['tree_id', 'height']].head(10).to_dict(orient='records')}"
        )

    bad_branch_height = df[df["branch_height"] < 0]
    if not bad_branch_height.empty:
        raise ValueError(
            f"枝下高存在 < 0 的值，请检查。问题记录示例："
            f"{bad_branch_height[['tree_id', 'branch_height']].head(10).to_dict(orient='records')}"
        )

    bad_branch_vs_height = df[df["branch_height"] >= df["height"]]
    if not bad_branch_vs_height.empty:
        raise ValueError(
            f"存在枝下高 >= 树高的记录，请检查。问题记录示例："
            f"{bad_branch_vs_height[['tree_id', 'branch_height', 'height']].head(10).to_dict(orient='records')}"
        )

    bad_crown = df[df["crown_width"] <= 0]
    if not bad_crown.empty:
        raise ValueError(
            f"平均冠幅存在 <= 0 的值，请检查。问题记录示例："
            f"{bad_crown[['tree_id', 'crown_width']].head(10).to_dict(orient='records')}"
        )

    if df["tree_id"].duplicated().any():
        raise ValueError("Tree_ID 存在重复，请保证每株树唯一")

    return df


def compute_basic_indices(df):
    df["hdr"] = df["height"] / df["dbh"]
    df["crown_length"] = df["height"] - df["branch_height"]
    df["crown_ratio"] = df["crown_length"] / df["height"]
    return df


# =========================================================
# 四、空间结构指标
# =========================================================

def compute_distance_matrix(df):
    coords = df[["x", "y"]].values
    dist_mat = distance_matrix(coords, coords)
    np.fill_diagonal(dist_mat, np.inf)
    return dist_mat


def get_neighbors(dist_mat, method="k", k=4, radius_val=3.0):
    neighbors_list = []

    if method == "k":
        sorted_idx = np.argsort(dist_mat, axis=1)
        for i in range(dist_mat.shape[0]):
            neighbors_list.append(sorted_idx[i, :k])

    elif method == "radius":
        for i in range(dist_mat.shape[0]):
            idx = np.where(dist_mat[i] <= radius_val)[0]
            neighbors_list.append(idx)
    else:
        raise ValueError("neighbor_method 只能是 'k' 或 'radius'")

    return neighbors_list


def compute_spatial_indices(df, dist_mat, neighbors_list):
    competition_index = []
    nearest_distance = []
    size_diff = []
    neighbor_count = []

    for i in range(len(df)):
        dbh_i = df.loc[i, "dbh"]
        neigh_idx = neighbors_list[i]

        if len(neigh_idx) == 0:
            competition_index.append(0.0)
            nearest_distance.append(np.nan)
            size_diff.append(np.nan)
            neighbor_count.append(0)
            continue

        neigh_dist = dist_mat[i, neigh_idx]
        neigh_dbh = df.loc[neigh_idx, "dbh"].values

        valid_mask = np.isfinite(neigh_dist) & (neigh_dist > 0)
        neigh_dist = neigh_dist[valid_mask]
        neigh_dbh = neigh_dbh[valid_mask]

        if len(neigh_dist) == 0:
            competition_index.append(0.0)
            nearest_distance.append(np.nan)
            size_diff.append(np.nan)
            neighbor_count.append(0)
            continue

        ci = np.sum((neigh_dbh / dbh_i) / neigh_dist)
        nnd = np.min(neigh_dist)
        sd = dbh_i / np.mean(neigh_dbh)

        competition_index.append(ci)
        nearest_distance.append(nnd)
        size_diff.append(sd)
        neighbor_count.append(len(neigh_dist))

    df["competition_index"] = competition_index
    df["nearest_distance"] = nearest_distance
    df["size_diff"] = size_diff
    df["neighbor_count"] = neighbor_count

    for col in ["nearest_distance", "size_diff"]:
        df[col] = df[col].fillna(df[col].median())

    return df


# =========================================================
# 五、标准化与双模式评分
# =========================================================

def standardize_all_indices(df, use_viability=False):
    positive_cols = [
        "dbh", "height", "crown_width",
        "crown_ratio", "stem_quality",
        "nearest_distance", "size_diff"
    ]
    negative_cols = ["hdr", "competition_index"]

    if use_viability and "viability_score" in df.columns and df["viability_score"].notna().any():
        positive_cols.append("viability_score")

    for col in positive_cols:
        df[f"{col}_std"] = normalize_positive(df[col])

    for col in negative_cols:
        df[f"{col}_std"] = normalize_negative(df[col])

    return df


def compute_quality_score(df, weight_mode="entropy", manual_weights=None, use_viability=False):
    std_cols = [
        "dbh_std",
        "height_std",
        "crown_width_std",
        "crown_ratio_std",
        "hdr_std",
        "stem_quality_std",
        "competition_index_std",
        "nearest_distance_std",
        "size_diff_std",
    ]

    if use_viability and "viability_score_std" in df.columns:
        std_cols.append("viability_score_std")

    missing_cols = [col for col in std_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"缺少标准化字段：{missing_cols}")

    if weight_mode == "entropy":
        weight_dict = calculate_entropy_weights(df, std_cols)
    elif weight_mode == "manual":
        if manual_weights is None:
            raise ValueError("手动权重模式下，manual_weights 不能为空")

        manual_weights = {k: v for k, v in manual_weights.items() if k in std_cols}
        missing_manual = [col for col in std_cols if col not in manual_weights]
        if missing_manual:
            raise ValueError(f"手动权重缺少这些字段：{missing_manual}")

        weight_dict = normalize_weights(manual_weights)
    else:
        raise ValueError("weight_mode 只能是 'entropy' 或 'manual'")

    df["quality_score"] = 0.0
    for col, w in weight_dict.items():
        df["quality_score"] += df[col] * w

    df.attrs["weight_dict"] = weight_dict
    return df


# =========================================================
# 六、分类
# =========================================================

def classify_target_trees(df, q=0.80):
    threshold = df["quality_score"].quantile(q)
    df["tree_class"] = df["quality_score"].apply(lambda x: "目标树" if x >= threshold else "一般木")
    return df


def identify_interference_trees(df, dist_mat, ci_q=0.70, max_dist=3.0):
    target_idx = df.index[df["tree_class"] == "目标树"].tolist()

    if len(target_idx) == 0:
        df["is_interference"] = False
        return df

    ci_threshold = df["competition_index"].quantile(ci_q)

    is_interference = []
    for i in range(len(df)):
        if i in target_idx:
            is_interference.append(False)
            continue

        min_dist_to_target = np.min(dist_mat[i, target_idx])

        if (df.loc[i, "competition_index"] >= ci_threshold) and (min_dist_to_target <= max_dist):
            is_interference.append(True)
        else:
            is_interference.append(False)

    df["is_interference"] = is_interference
    df.loc[df["is_interference"], "tree_class"] = "干扰树"

    return df


# =========================================================
# 七、结果输出
# =========================================================

def build_abnormal_sheet(dead_df, invalid_crown_df, df_result):
    abnormal_list = []

    if dead_df is not None and not dead_df.empty:
        temp = dead_df.copy()
        temp["异常类型"] = "死亡木/无效木"
        abnormal_list.append(temp)

    if invalid_crown_df is not None and not invalid_crown_df.empty:
        temp = invalid_crown_df.copy()
        temp["异常类型"] = "平均冠幅<=0"
        abnormal_list.append(temp)

    if df_result is not None and not df_result.empty:
        suspicious = df_result[df_result["nearest_distance"] < 0.05].copy()
        if not suspicious.empty:
            suspicious["异常类型"] = "最近邻距离过小(<0.05m)"
            abnormal_list.append(suspicious)

    if abnormal_list:
        abnormal_df = pd.concat(abnormal_list, ignore_index=True, sort=False)
    else:
        abnormal_df = pd.DataFrame(columns=["异常类型"])

    return abnormal_df


def generate_summary(df):
    summary = pd.DataFrame({
        "总株数": [len(df)],
        "目标树数量": [(df["tree_class"] == "目标树").sum()],
        "干扰树数量": [(df["tree_class"] == "干扰树").sum()],
        "一般木数量": [(df["tree_class"] == "一般木").sum()],
        "平均综合得分": [df["quality_score"].mean()],
        "平均竞争指数": [df["competition_index"].mean()],
        "平均最近邻距离": [df["nearest_distance"].mean()],
        "平均高径比": [df["hdr"].mean()],
        "平均冠长率": [df["crown_ratio"].mean()],
    })

    weight_dict = df.attrs.get("weight_dict", {})
    weight_df = pd.DataFrame({
        "指标": list(weight_dict.keys()),
        "权重": list(weight_dict.values())
    })

    return summary, weight_df


def save_to_excel(df_result, df_summary, df_weights, df_abnormal, file_path):
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        df_result.to_excel(writer, sheet_name="结果表", index=False)
        df_summary.to_excel(writer, sheet_name="汇总表", index=False)
        df_weights.to_excel(writer, sheet_name="权重表", index=False)
        df_abnormal.to_excel(writer, sheet_name="异常记录表", index=False)


# =========================================================
# 八、主计算流程
# =========================================================

def run_tree_quality_program(
    input_file,
    output_file,
    col_map,
    neighbor_method,
    k_neighbors,
    radius,
    weight_mode,
    manual_weights,
    target_quantile,
    competition_quantile,
    interference_radius,
    use_viability,
    logger=print
):
    logger("开始执行单木质量评价...")

    check_file_exists(input_file)

    logger("1/11 读取 Excel 数据...")
    df = read_excel_data(input_file)

    logger("2/11 统一字段名...")
    df = rename_columns(df, col_map)

    logger("3/11 剔除死亡木 / 无效木...")
    df, dead_df = filter_dead_trees(df, logger=logger)

    logger("4/11 剔除平均冠幅无效木...")
    df, invalid_crown_df = filter_invalid_crown_trees(df, logger=logger)

    logger("5/11 检查数据完整性...")
    df = validate_data(df)

    logger("6/11 计算基础指标（平均冠幅、高径比、冠长、冠长率）...")
    df = compute_basic_indices(df)

    logger("7/11 构建邻域...")
    dist_mat = compute_distance_matrix(df)
    neighbors_list = get_neighbors(
        dist_mat,
        method=neighbor_method,
        k=k_neighbors,
        radius_val=radius
    )

    logger("8/11 计算空间结构指标...")
    df = compute_spatial_indices(df, dist_mat, neighbors_list)

    logger("9/11 标准化指标...")
    df = standardize_all_indices(df, use_viability=use_viability)

    if weight_mode == "entropy":
        logger("10/11 使用熵权法自动计算权重并计算综合得分...")
    else:
        logger("10/11 使用手动权重计算综合得分...")

    df = compute_quality_score(
        df,
        weight_mode=weight_mode,
        manual_weights=manual_weights,
        use_viability=use_viability
    )

    logger("11/11 识别目标树和干扰树...")
    df = classify_target_trees(df, q=target_quantile)
    df = identify_interference_trees(
        df,
        dist_mat,
        ci_q=competition_quantile,
        max_dist=interference_radius
    )

    logger("正在输出结果...")
    df_result = df.copy()

    rename_result = {
        "tree_id": "木编号",
        "tree_species": "树种",
        "x": "X坐标",
        "y": "Y坐标",
        "dbh": "胸径_cm",
        "height": "树高_m",
        "stem_quality": "干形质量",
        "branch_height": "枝下高_m",
        "cw_s": "南向冠幅_m",
        "cw_n": "北向冠幅_m",
        "cw_e": "东向冠幅_m",
        "cw_w": "西向冠幅_m",
        "crown_width": "平均冠幅_m",
        "crown_length": "冠长_m",
        "crown_ratio": "冠长率",
        "viability": "活力等级",
        "viability_score": "活力得分",
        "hdr": "高径比",
        "competition_index": "竞争指数",
        "nearest_distance": "最近邻距离_m",
        "size_diff": "相对大小指标",
        "neighbor_count": "邻木数量",
        "quality_score": "综合质量得分",
        "tree_class": "树木类别",
        "is_interference": "是否干扰树",
    }
    df_result = df_result.rename(columns=rename_result)

    std_cols = [c for c in df_result.columns if c.endswith("_std")]
    if std_cols:
        df_result = df_result.drop(columns=std_cols)

    df_summary, df_weights = generate_summary(df)
    df_abnormal = build_abnormal_sheet(dead_df, invalid_crown_df, df)

    save_to_excel(df_result, df_summary, df_weights, df_abnormal, output_file)

    logger(f"完成！结果已保存到：{output_file}")

    return df.attrs.get("weight_dict", {})


# =========================================================
# 九、PyQt5 图形界面
# =========================================================

class StructureMetricsWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("单木质量评价小程序（最终双模式权重版）")
        self.resize(980, 700)
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()

        # 文件设置
        file_group = QGroupBox("文件设置")
        file_layout = QGridLayout()

        self.input_edit = QLineEdit()
        self.output_edit = QLineEdit()

        btn_input = QPushButton("选择输入 Excel")
        btn_output = QPushButton("选择输出位置")

        btn_input.clicked.connect(self.select_input_file)
        btn_output.clicked.connect(self.select_output_file)

        file_layout.addWidget(QLabel("输入文件："), 0, 0)
        file_layout.addWidget(self.input_edit, 0, 1)
        file_layout.addWidget(btn_input, 0, 2)

        file_layout.addWidget(QLabel("输出文件："), 1, 0)
        file_layout.addWidget(self.output_edit, 1, 1)
        file_layout.addWidget(btn_output, 1, 2)

        file_group.setLayout(file_layout)

        # 参数设置
        param_group = QGroupBox("参数设置")
        param_layout = QGridLayout()

        self.neighbor_method_combo = QComboBox()
        self.neighbor_method_combo.addItems(["k", "radius"])

        self.k_spin = QSpinBox()
        self.k_spin.setRange(1, 20)
        self.k_spin.setValue(4)

        self.radius_spin = QDoubleSpinBox()
        self.radius_spin.setRange(0.1, 100.0)
        self.radius_spin.setDecimals(2)
        self.radius_spin.setValue(3.0)

        self.target_q_spin = QDoubleSpinBox()
        self.target_q_spin.setRange(0.01, 0.99)
        self.target_q_spin.setDecimals(2)
        self.target_q_spin.setValue(0.80)

        self.comp_q_spin = QDoubleSpinBox()
        self.comp_q_spin.setRange(0.01, 0.99)
        self.comp_q_spin.setDecimals(2)
        self.comp_q_spin.setValue(0.70)

        self.interference_radius_spin = QDoubleSpinBox()
        self.interference_radius_spin.setRange(0.1, 100.0)
        self.interference_radius_spin.setDecimals(2)
        self.interference_radius_spin.setValue(3.0)

        self.use_viability_combo = QComboBox()
        self.use_viability_combo.addItems(["否", "是"])

        self.weight_mode_combo = QComboBox()
        self.weight_mode_combo.addItems(["自动-熵权法", "手动权重"])
        self.weight_mode_combo.currentTextChanged.connect(self.on_weight_mode_changed)

        param_layout.addWidget(QLabel("邻域方式："), 0, 0)
        param_layout.addWidget(self.neighbor_method_combo, 0, 1)

        param_layout.addWidget(QLabel("最近邻株数 k："), 0, 2)
        param_layout.addWidget(self.k_spin, 0, 3)

        param_layout.addWidget(QLabel("半径 r（m）："), 0, 4)
        param_layout.addWidget(self.radius_spin, 0, 5)

        param_layout.addWidget(QLabel("目标树分位数："), 1, 0)
        param_layout.addWidget(self.target_q_spin, 1, 1)

        param_layout.addWidget(QLabel("竞争木分位数："), 1, 2)
        param_layout.addWidget(self.comp_q_spin, 1, 3)

        param_layout.addWidget(QLabel("干扰半径（m）："), 1, 4)
        param_layout.addWidget(self.interference_radius_spin, 1, 5)

        param_layout.addWidget(QLabel("是否纳入活力："), 2, 0)
        param_layout.addWidget(self.use_viability_combo, 2, 1)

        param_layout.addWidget(QLabel("权重模式："), 2, 2)
        param_layout.addWidget(self.weight_mode_combo, 2, 3)

        param_group.setLayout(param_layout)

        # 权重设置
        weight_group = QGroupBox("权重设置")
        weight_layout = QGridLayout()

        self.weight_dbh = self.create_weight_spin(0.14)
        self.weight_height = self.create_weight_spin(0.08)
        self.weight_crown = self.create_weight_spin(0.10)
        self.weight_crown_ratio = self.create_weight_spin(0.10)
        self.weight_hdr = self.create_weight_spin(0.10)
        self.weight_stem = self.create_weight_spin(0.14)
        self.weight_ci = self.create_weight_spin(0.18)
        self.weight_nnd = self.create_weight_spin(0.08)
        self.weight_size = self.create_weight_spin(0.08)
        self.weight_viability = self.create_weight_spin(0.00)

        weight_layout.addWidget(QLabel("胸径权重"), 0, 0)
        weight_layout.addWidget(self.weight_dbh, 0, 1)

        weight_layout.addWidget(QLabel("树高权重"), 0, 2)
        weight_layout.addWidget(self.weight_height, 0, 3)

        weight_layout.addWidget(QLabel("平均冠幅权重"), 0, 4)
        weight_layout.addWidget(self.weight_crown, 0, 5)

        weight_layout.addWidget(QLabel("冠长率权重"), 1, 0)
        weight_layout.addWidget(self.weight_crown_ratio, 1, 1)

        weight_layout.addWidget(QLabel("高径比权重"), 1, 2)
        weight_layout.addWidget(self.weight_hdr, 1, 3)

        weight_layout.addWidget(QLabel("干形质量权重"), 1, 4)
        weight_layout.addWidget(self.weight_stem, 1, 5)

        weight_layout.addWidget(QLabel("竞争指数权重"), 2, 0)
        weight_layout.addWidget(self.weight_ci, 2, 1)

        weight_layout.addWidget(QLabel("最近邻距离权重"), 2, 2)
        weight_layout.addWidget(self.weight_nnd, 2, 3)

        weight_layout.addWidget(QLabel("相对大小权重"), 2, 4)
        weight_layout.addWidget(self.weight_size, 2, 5)

        weight_layout.addWidget(QLabel("活力权重"), 3, 0)
        weight_layout.addWidget(self.weight_viability, 3, 1)

        weight_group.setLayout(weight_layout)

        # 说明
        note_group = QGroupBox("说明")
        note_layout = QVBoxLayout()
        self.note_text = QTextEdit()
        self.note_text.setReadOnly(True)
        self.note_text.setPlainText(
            "权重模式说明：\n"
            "1. 自动-熵权法：权重框显示但不可编辑，运行后自动回填计算结果\n"
            "2. 手动权重：权重框可编辑，程序按输入权重参与计算\n"
            "程序还会自动：\n"
            "1. 剔除死亡木/无效木\n"
            "2. 剔除平均冠幅<=0的无效木\n"
            "3. 导出结果表、汇总表、权重表、异常记录表"
        )
        note_layout.addWidget(self.note_text)
        note_group.setLayout(note_layout)

        # 按钮区
        btn_layout = QHBoxLayout()
        self.run_btn = QPushButton("开始运行")
        self.clear_btn = QPushButton("清空日志")

        self.run_btn.clicked.connect(self.run_program)
        self.clear_btn.clicked.connect(self.clear_log)

        btn_layout.addStretch()
        btn_layout.addWidget(self.run_btn)
        btn_layout.addWidget(self.clear_btn)

        # 日志区
        log_group = QGroupBox("运行日志")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)

        main_layout.addWidget(file_group)
        main_layout.addWidget(param_group)
        main_layout.addWidget(weight_group)
        main_layout.addWidget(note_group)
        main_layout.addLayout(btn_layout)
        main_layout.addWidget(log_group)

        self.setLayout(main_layout)

        self.on_weight_mode_changed()

    def create_weight_spin(self, value):
        spin = QDoubleSpinBox()
        spin.setRange(0.0, 10.0)
        spin.setDecimals(3)
        spin.setValue(value)
        return spin

    def on_weight_mode_changed(self):
        """
        根据权重模式控制权重输入框是否可编辑
        """
        mode = self.weight_mode_combo.currentText().strip()
        is_manual = (mode == "手动权重")

        weight_widgets = [
            self.weight_dbh,
            self.weight_height,
            self.weight_crown,
            self.weight_crown_ratio,
            self.weight_hdr,
            self.weight_stem,
            self.weight_ci,
            self.weight_nnd,
            self.weight_size,
            self.weight_viability,
        ]

        for widget in weight_widgets:
            widget.setEnabled(is_manual)

        # 说明文字同步更新
        if is_manual:
            self.note_text.setPlainText(
                "当前为【手动权重】模式。\n"
                "1. 权重框可编辑\n"
                "2. 程序按输入权重参与计算\n"
                "3. 程序会自动将手动权重归一化\n"
                "4. 仍会自动剔除死亡木、冠幅无效木，并导出异常记录"
            )
        else:
            self.note_text.setPlainText(
                "当前为【自动-熵权法】模式。\n"
                "1. 权重框显示但不可编辑\n"
                "2. 运行后程序会自动回填熵权结果\n"
                "3. 仍会自动剔除死亡木、冠幅无效木，并导出异常记录"
            )

    def update_weight_display(self, weight_dict):
        mapping = {
            "dbh_std": self.weight_dbh,
            "height_std": self.weight_height,
            "crown_width_std": self.weight_crown,
            "crown_ratio_std": self.weight_crown_ratio,
            "hdr_std": self.weight_hdr,
            "stem_quality_std": self.weight_stem,
            "competition_index_std": self.weight_ci,
            "nearest_distance_std": self.weight_nnd,
            "size_diff_std": self.weight_size,
            "viability_score_std": self.weight_viability,
        }

        for key, widget in mapping.items():
            if key in weight_dict:
                widget.setValue(float(weight_dict[key]))

    def log(self, msg):
        self.log_text.append(msg)
        self.log_text.ensureCursorVisible()

    def clear_log(self):
        self.log_text.clear()

    def select_input_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择输入 Excel 文件", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.input_edit.setText(file_path)
            if not self.output_edit.text().strip():
                base = os.path.splitext(file_path)[0]
                self.output_edit.setText(base + "_result.xlsx")

    def select_output_file(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "选择输出 Excel 文件", "", "Excel Files (*.xlsx)"
        )
        if file_path:
            if not file_path.lower().endswith(".xlsx"):
                file_path += ".xlsx"
            self.output_edit.setText(file_path)

    def run_program(self):
        input_file = self.input_edit.text().strip()
        output_file = self.output_edit.text().strip()

        if not input_file:
            QMessageBox.warning(self, "提示", "请先选择输入 Excel 文件")
            return

        if not output_file:
            QMessageBox.warning(self, "提示", "请先设置输出 Excel 文件")
            return

        col_map = {
            "tree_id": "Tree_ID",
            "tree_species": "Tree_speice",
            "x": "X",
            "y": "Y",
            "dbh": "DBH (cm)",
            "height": "H (m)",
            "stem_quality": "Stem quality",
            "branch_height": "Branch height (m)",
            "cw_s": "CW_S (m)",
            "cw_n": "CW_N (m)",
            "cw_e": "CW_E (m)",
            "cw_w": "CW_W (m)",
            "viability": "Viability",
        }

        use_viability = self.use_viability_combo.currentText() == "是"
        weight_mode = "entropy" if self.weight_mode_combo.currentText().strip() == "自动-熵权法" else "manual"

        manual_weights = {
            "dbh_std": self.weight_dbh.value(),
            "height_std": self.weight_height.value(),
            "crown_width_std": self.weight_crown.value(),
            "crown_ratio_std": self.weight_crown_ratio.value(),
            "hdr_std": self.weight_hdr.value(),
            "stem_quality_std": self.weight_stem.value(),
            "competition_index_std": self.weight_ci.value(),
            "nearest_distance_std": self.weight_nnd.value(),
            "size_diff_std": self.weight_size.value(),
        }

        if use_viability:
            manual_weights["viability_score_std"] = self.weight_viability.value()

        try:
            self.log("=" * 60)
            self.log("开始运行...")

            weight_dict = run_tree_quality_program(
                input_file=input_file,
                output_file=output_file,
                col_map=col_map,
                neighbor_method=self.neighbor_method_combo.currentText(),
                k_neighbors=self.k_spin.value(),
                radius=self.radius_spin.value(),
                weight_mode=weight_mode,
                manual_weights=manual_weights,
                target_quantile=self.target_q_spin.value(),
                competition_quantile=self.comp_q_spin.value(),
                interference_radius=self.interference_radius_spin.value(),
                use_viability=use_viability,
                logger=self.log
            )

            if weight_mode == "entropy":
                self.update_weight_display(weight_dict)

            QMessageBox.information(self, "完成", f"程序运行完成！\n结果已保存到：\n{output_file}")

        except Exception as e:
            self.log("程序运行失败！")
            self.log(str(e))
            self.log(traceback.format_exc())
            QMessageBox.critical(self, "错误", f"程序运行失败：\n{e}")


# =========================================================
# 十、程序入口
# =========================================================

def main():
    app = QApplication(sys.argv)
    window = TreeQualityApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()