# -*- coding: utf-8 -*-
from __future__ import annotations

import math
import os
import re
from dataclasses import dataclass, field
from typing import Callable, Optional, List, Tuple

import numpy as np
import pandas as pd
from scipy.spatial import KDTree

# 补全丢失的导出路径类
class ExportPaths:
    def __init__(self):
        self.main = None
        self.means = None
        self.entropy = None
        self.q_tree = None

# 补全丢失的分类编码器
class CategoryEncoder:
    def __init__(self):
        self.mapping = {}
        self.counter = 0
    def encode(self, label):
        if label not in self.mapping:
            self.mapping[label] = float(self.counter)
            self.counter += 1
        return self.mapping[label]

# 补全丢失的 KDTree 包装器
class KDTree2D:
    def __init__(self, points):
        self.points = list(points)
        self.tree = KDTree(self.points)
    def k_nearest(self, i, k):
        distances, indices = self.tree.query(self.points[i], k=k)
        # 返回 (index, distance_squared) 兼容后续逻辑
        return [(idx, d**2) for d, idx in zip(distances, indices)]

# --- 以下是根据你原有残余代码修复的函数开始 ---

def mark_core_buffer(out, buffer_size):
    # 修复原本被 XML 截断的逻辑
    out["IsCore"] = (out["DistToBoundary"] >= buffer_size).astype(bool)
    out["Zone"] = np.where(out["IsCore"], "CORE", "BUFFER")
    return out


def compute_all(standard_df: pd.DataFrame, infl_k: float) -> pd.DataFrame:
    if len(standard_df) < 5:
        raise ValueError("至少需要 5 株树（n>=5）才能计算四近邻指标。")

    n = len(standard_df)
    out = standard_df.copy().reset_index(drop=True)

    tree = out["Tree"].to_numpy(dtype=float)
    x = out["X"].to_numpy(dtype=float)
    y = out["Y"].to_numpy(dtype=float)
    dbh = out["DBH"].to_numpy(dtype=float)
    species = out["Species"].to_numpy(dtype=float)
    height = out["Height"].to_numpy(dtype=float)
    crown_d = out["CrownRadius"].to_numpy(dtype=float)

    kd = KDTree2D(zip(x, y))
    nn = np.zeros((n, 4), dtype=int)
    dd = np.zeros((n, 4), dtype=float)

    for i in range(n):
        neighbors = [(idx, math.sqrt(dist2)) for idx, dist2 in kd.k_nearest(i, 6) if idx != i][:4]
        if len(neighbors) < 4:
            raise ValueError("近邻不足（可能存在大量重复坐标或极端分布）。")
        for k, (idx, dist) in enumerate(neighbors):
            nn[i, k] = idx
            dd[i, k] = dist

    for idx in range(4):
        out[f"Tree{idx + 1}"] = tree[nn[:, idx]]
        out[f"Dist{idx + 1}"] = dd[:, idx]

    top_h = np.sort(height)[-min(10, n) :]
    avg_h = float(np.mean(top_h)) if len(top_h) else 0.0

    def layer_cat(h_value: float) -> float:
        if h_value > avg_h * (2.0 / 3.0):
            return 1.0
        if h_value > avg_h * (1.0 / 3.0):
            return 0.0
        return -1.0

    rows: list[dict[str, object]] = []
    for i in range(n):
        angles = []
        for k in range(4):
            dx = x[nn[i, k]] - x[i]
            dy = y[nn[i, k]] - y[i]
            ang = math.atan2(dy, dx)
            if ang < 0:
                ang += 2 * math.pi
            angles.append(ang)
        angles.sort()
        gaps = [
            angles[1] - angles[0],
            angles[2] - angles[1],
            angles[3] - angles[2],
            2 * math.pi - (angles[3] - angles[0]),
        ]
        threshold = 72 * math.pi / 180.0
        w_value = sum(gap < threshold for gap in gaps) / 4.0

        m_value = sum(species[nn[i, k]] != species[i] for k in range(4)) / 4.0

        lc = layer_cat(height[i])
        lc_values = [lc] + [layer_cat(height[nn[i, k]]) for k in range(4)]
        layer_count = float(len(set(lc_values)))
        discreteness = sum(v != lc for v in lc_values[1:])
        layer_index = (layer_count / 3.0) * (discreteness / 4.0)

        openness = sum(dd[i, k] / max(height[nn[i, k]], np.finfo(float).eps) for k in range(4)) / 4.0
        ci_parts = []
        ci_total = 0.0
        dbh_i = max(dbh[i], np.finfo(float).eps)
        for k in range(4):
            dij = max(float(dd[i, k]), 0.1)
            cij = dbh[nn[i, k]] / (dbh_i * dij)
            ci_parts.append(cij)
            ci_total += cij

        path_order = 4.0 * math.floor((dbh[i] * 100.0 - 2.0) / 4.0) + 4.0

        bigger_d = sum(dbh[nn[i, k]] > dbh[i] for k in range(4))
        u_value = bigger_d / 4.0

        di = max(dbh[i], np.finfo(float).eps)
        hi = max(height[i], np.finfo(float).eps)

        u_d = sum(dbh[nn[i, k]] > di for k in range(4)) / 4.0
        u_h = sum(height[nn[i, k]] > hi for k in range(4)) / 4.0

        sum_term_d = 0.0
        sum_term_h = 0.0
        for k in range(4):
            j = nn[i, k]
            dij = max(float(dd[i, k]), float(np.finfo(float).eps))

            dj = max(dbh[j], np.finfo(float).eps)
            hj = max(height[j], np.finfo(float).eps)

            c_h = 1.0 if hj > hi else 0.0
            a1_h = math.atan(min(hi, hj) / dij)
            a2_h = max(math.atan((hj - hi) / dij), 0.0)
            sum_term_h += (a1_h + a2_h * c_h) / math.pi

            c_d = 1.0 if dj > di else 0.0
            a1_d = math.atan(min(di, dj) / dij)
            a2_d = max(math.atan((dj - di) / dij), 0.0)
            sum_term_d += (a1_d + a2_d * c_d) / math.pi

        uci_h = clamp01((sum_term_h / 4.0) * u_h)
        uci_d = clamp01((sum_term_d / 4.0) * u_d)

        c_value = sum(dd[i, k] < (crown_d[i] + crown_d[nn[i, k]]) for k in range(4)) / 4.0

        r_infl = infl_k * crown_d[i]
        f_value = sum(dd[i, k] > r_infl for k in range(4)) / 4.0

        rows.append(
            {
                "W": w_value,
                "M": m_value,
                "DominantHeight": avg_h,
                "LayerCategory": lc,
                "LayerCategory1": lc_values[1],
                "LayerCategory2": lc_values[2],
                "LayerCategory3": lc_values[3],
                "LayerCategory4": lc_values[4],
                "LayerCount": layer_count,
                "LayerIndex": layer_index,
                "Discreteness": float(discreteness),
                "Openness": openness,
                "K": openness,
                "CI1": ci_parts[0],
                "CI2": ci_parts[1],
                "CI3": ci_parts[2],
                "CI4": ci_parts[3],
                "CI": ci_total,
                "PathOrder": path_order,
                "U": u_value,
                "U_class": u_class(u_value),
                "UCI_D": uci_d,
                "UCI_H": uci_h,
                "UCI": uci_d,
                "UCI_class": uci_class(uci_d),
                "C": c_value,
                "F": f_value,
            }
        )

    metrics_df = pd.DataFrame(rows)
    for column in metrics_df.columns:
        out[column] = metrics_df[column]
    return out


def _entropy_weight_impl(x: np.ndarray, kinds: list[str]) -> np.ndarray:
    n, m = x.shape
    y = x.astype(float).copy()

    for col in range(m):
        col_min = np.nanmin(y[:, col])
        col_max = np.nanmax(y[:, col])
        rng = col_max - col_min
        if not np.isfinite(rng) or rng == 0:
            y[:, col] = 0.0
        else:
            y[:, col] = (y[:, col] - col_min) / rng
        if kinds[col] == "neg":
            y[:, col] = 1.0 - y[:, col]

    col_sum = np.nansum(y, axis=0)
    col_sum[col_sum == 0] = np.finfo(float).eps
    p = y / col_sum
    k = 1.0 / math.log(max(n, 2))
    e = -k * np.nansum(p * np.log(p + np.finfo(float).eps), axis=0)
    d = 1.0 - e
    sum_d = float(np.nansum(d))
    if sum_d == 0 or not math.isfinite(sum_d):
        d = np.ones_like(d)
        sum_d = float(np.nansum(d))
    return d / sum_d


def entropy_weight(x: np.ndarray, kinds: Optional[list[str]] = None) -> np.ndarray:
    values = np.asarray(x, dtype=float)
    if values.ndim != 2 or values.shape[1] == 0:
        raise ValueError("entropy_weight 需要二维指标矩阵。")
    resolved_kinds = kinds or ["pos"] * values.shape[1]
    if len(resolved_kinds) != values.shape[1]:
        raise ValueError("entropy_weight 的指标方向数量与列数不一致。")
    return _entropy_weight_impl(values, resolved_kinds)


def compute_weighted_q(df: pd.DataFrame, weight_method: str = "entropy") -> tuple[pd.DataFrame, pd.DataFrame, float, float, float]:
    required = ["W", "M", "LayerIndex", "K", "F", "C", "U", "UCI", "Tree", "Zone", "IsCore"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"EntropyQ 缺少字段: {', '.join(missing)}")

    normalized_weight_method = normalize_weight_method(weight_method)
    core_mask = df["IsCore"].astype(bool).to_numpy()
    if core_mask.sum() < 5:
        core_mask = np.ones(len(df), dtype=bool)

    x_all = np.column_stack(
        [
            np.abs(df["W"].to_numpy(dtype=float) - 0.5),
            df["C"].to_numpy(dtype=float),
            df["U"].to_numpy(dtype=float),
            df["UCI"].to_numpy(dtype=float),
            df["M"].to_numpy(dtype=float),
            df["LayerIndex"].to_numpy(dtype=float),
            df["K"].to_numpy(dtype=float),
            df["F"].to_numpy(dtype=float),
        ]
    )
    x_core = x_all[core_mask]
    if normalized_weight_method == "equal":
        weights = equal_weight(x_core.shape[1])
    else:
        weights = _entropy_weight_impl(x_core, ["neg", "neg", "neg", "neg", "pos", "pos", "pos", "pos"])

    ew, ec, eu, euci, em, es, ek, ef = weights
    numerator = ((1 + df["K"]) * ek) * ((1 + df["M"]) * em) * ((1 + df["LayerIndex"]) * es) * ((1 + df["F"]) * ef)
    denominator = ((1 + np.abs(df["W"] - 0.5)) * ew) * ((1 + df["C"]) * ec) * ((1 + df["U"]) * eu) * ((1 + df["UCI"]) * euci)
    q_series = numerator / denominator.replace(0, np.nan)

    out = df.copy()
    out["Q"] = q_series.astype(float)

    q_all = finite_mean(out["Q"])
    q_core = finite_mean(out.loc[out["IsCore"].astype(bool), "Q"])
    q_buf = finite_mean(out.loc[~out["IsCore"].astype(bool), "Q"])

    metric_names = ["|W-0.5|", "C", "U", "UCI", "M", "S", "K", "F"]
    directions = ["负向", "负向", "负向", "负向", "正向", "正向", "正向", "正向"]
    method_name = weight_method_display_name(normalized_weight_method)
    entropy_df = pd.DataFrame(
        {
            "Metric": metric_names,
            "Weight": weights,
            "Method": [method_name] * len(metric_names),
            "Direction": directions,
        }
    )

    q_table = pd.DataFrame(
        {
            "Tree": out["Tree"],
            "Zone": out["Zone"],
            "IsCore": out["IsCore"],
            "W": out["W"],
            "M": out["M"],
            "S": out["LayerIndex"],
            "K": out["K"],
            "C": out["C"],
            "U": out["U"],
            "UCI": out["UCI"],
            "F": out["F"],
            "Ew": ew,
            "Em": em,
            "Es": es,
            "Ek": ek,
            "Ec": ec,
            "Eu": eu,
            "Euci": euci,
            "Ef": ef,
            "WeightMethod": method_name,
            "Q": out["Q"],
        }
    )

    df["Q"] = out["Q"]
    return entropy_df, q_table, q_all, q_core, q_buf


def compute_entropy_and_q(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, float, float, float]:
    return compute_weighted_q(df, "entropy")


def equal_weight(metric_count: int) -> np.ndarray:
    metric_count = max(1, int(metric_count))
    return np.full(metric_count, 1.0 / metric_count, dtype=float)


def normalize_weight_method(weight_method: str) -> str:
    key = str(weight_method or "entropy").strip().lower()
    if key in {"equal", "equal_weight", "均权", "等权", "ew"}:
        return "equal"
    return "entropy"


def weight_method_display_name(weight_method: str) -> str:
    return "等权" if normalize_weight_method(weight_method) == "equal" else "熵权"


def build_stand_means(df: pd.DataFrame, q_all: float, q_core: float, q_buf: float) -> pd.DataFrame:
    metrics = [
        "W",
        "M",
        "U",
        "UCI",
        "C",
        "F",
        "Openness",
        "K",
        "CI",
        "LayerIndex",
        "UCI_D",
        "UCI_H",
        "DistToBoundary",
        "Q",
    ]
    present = [metric for metric in metrics if metric in df.columns]
    rows: list[dict[str, object]] = []

    groups = {
        "ALL": np.ones(len(df), dtype=bool),
        "CORE": df["IsCore"].astype(bool).to_numpy() if "IsCore" in df.columns else np.zeros(len(df), dtype=bool),
        "BUFFER": ~df["IsCore"].astype(bool).to_numpy() if "IsCore" in df.columns else np.zeros(len(df), dtype=bool),
    }

    for group_name, mask in groups.items():
        for metric in present:
            mean_value = finite_mean(df.loc[mask, metric])
            if math.isfinite(mean_value):
                rows.append({"Group": group_name, "Metric": metric, "Mean": mean_value})

    rows.extend(
        [
            {"Group": "ALL", "Metric": "Qbar", "Mean": q_all},
            {"Group": "CORE", "Metric": "Qbar", "Mean": q_core},
            {"Group": "BUFFER", "Metric": "Qbar", "Mean": q_buf},
        ]
    )
    return pd.DataFrame(rows, columns=["Group", "Metric", "Mean"])


def build_main_for_preview(df: pd.DataFrame, use_wm: bool, use_layer: bool, use_uuci: bool, use_cf: bool, use_q: bool) -> pd.DataFrame:
    keep: list[str] = []

    def add(*cols: str):
        for col in cols:
            if col in df.columns and col not in keep:
                keep.append(col)

    add("Tree", "X", "Y", "DBH", "Species", "Height", "CrownRadius")
    add("BufferSize_m", "DistToBoundary", "IsCore", "Zone")
    add("Tree1", "Tree2", "Tree3", "Tree4", "Dist1", "Dist2", "Dist3", "Dist4")
    if use_wm:
        add("W", "M")
    if use_layer:
        add(
            "DominantHeight",
            "LayerCategory",
            "LayerCount",
            "LayerIndex",
            "Discreteness",
            "Openness",
            "K",
            "CI1",
            "CI2",
            "CI3",
            "CI4",
            "CI",
            "PathOrder",
        )
    if use_uuci:
        add("U", "U_class", "UCI", "UCI_class", "UCI_D", "UCI_H")
    if use_cf:
        add("C", "F")
    if use_q:
        add("Q")
    return df.loc[:, keep].copy()


def build_metric_list(use_wm: bool, use_layer: bool, use_uuci: bool, use_cf: bool, use_q: bool) -> list[str]:
    metrics: list[str] = []
    if use_wm:
        metrics.extend(["W", "M"])
    if use_uuci:
        metrics.extend(["U", "UCI", "UCI_D", "UCI_H"])
    if use_cf:
        metrics.extend(["C", "F"])
    if use_layer:
        metrics.extend(["Openness", "K", "LayerIndex", "CI", "PathOrder"])
    if use_q:
        metrics.append("Q")
    metrics.extend(["DBH", "Height", "DistToBoundary"])
    return metrics


def write_all(
    main_df: pd.DataFrame,
    means_df: pd.DataFrame,
    entropy_df: pd.DataFrame,
    q_table_df: pd.DataFrame,
    input_path: str,
    out_dir: str,
    export_format: str,
    write_main: bool,
    write_means: bool,
    write_entropy: bool,
    write_q: bool,
    weight_label: str = "熵权",
) -> ExportPaths:
    if not any([write_main, write_means, write_entropy, write_q]):
        return ExportPaths()

    input_path = input_path or "structure_metrics"
    in_folder = os.path.dirname(input_path) or "."
    base_name = os.path.splitext(os.path.basename(input_path))[0] or "structure_metrics"
    ext = os.path.splitext(os.path.basename(input_path))[1]

    target_dir = out_dir.strip() or in_folder
    os.makedirs(target_dir, exist_ok=True)

    csv_mode = str(export_format).upper() == "CSV"
    suffix = ".csv" if csv_mode else ".txt"
    sep = "," if csv_mode else "\t"

    paths = ExportPaths()
    if write_main:
        paths.main = os.path.join(target_dir, f"{base_name}{ext}-输出{suffix}")
        main_df.to_csv(paths.main, sep=sep, index=False, encoding="utf-8-sig")
    if write_means:
        paths.means = os.path.join(target_dir, f"{base_name}{ext}-输出-林分均值{suffix}")
        means_df.to_csv(paths.means, sep=sep, index=False, encoding="utf-8-sig")
    if write_entropy:
        paths.entropy = os.path.join(target_dir, f"{base_name}{ext}-输出-{weight_label}{suffix}")
        entropy_df.to_csv(paths.entropy, sep=sep, index=False, encoding="utf-8-sig")
    if write_q:
        paths.q_tree = os.path.join(target_dir, f"{base_name}{ext}-输出-综合指数{suffix}")
        q_table_df.to_csv(paths.q_tree, sep=sep, index=False, encoding="utf-8-sig")
    return paths


def normalize_name(value: str) -> str:
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "", (value or "").strip().lower())


def post_process_dbh_cm_to_m(df: pd.DataFrame):
    dbhs = pd.to_numeric(df["DBH"], errors="coerce")
    dbhs = dbhs[np.isfinite(dbhs)]
    if dbhs.empty:
        return
    median_value = float(np.median(dbhs))
    if 2 < median_value < 200:
        df["DBH"] = pd.to_numeric(df["DBH"], errors="coerce") / 100.0


def to_float(value) -> float:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return float("nan")
    if pd.isna(value):
        return float("nan")
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)
    text = str(value).strip()
    if not text:
        return float("nan")
    for candidate in (text, text.replace(",", ""), text.replace(",", ".")):
        try:
            return float(candidate)
        except ValueError:
            continue
    return float("nan")


def to_float_or_category(value, encoder: CategoryEncoder) -> float:
    number = to_float(value)
    if math.isfinite(number):
        return number
    return float(encoder.encode(str(value or "")))


def clamp01(value: float) -> float:
    if not math.isfinite(value) or value < 0:
        return 0.0
    if value >= 1:
        return 0.999999
    return float(value)


def u_class(value: float) -> str:
    if value == 0:
        return "优势"
    if value <= 0.25:
        return "亚优势"
    if value <= 0.50:
        return "中庸"
    if value <= 0.75:
        return "劣势"
    return "绝对劣势"


def uci_class(value: float) -> str:
    if value == 0:
        return "无竞争"
    if value <= 0.25:
        return "较小"
    if value <= 0.50:
        return "中等"
    if value <= 0.75:
        return "较大"
    return "极大"


def finite_mean(series) -> float:
    values = pd.to_numeric(pd.Series(series), errors="coerce")
    values = values[np.isfinite(values)]
    if values.empty:
        return float("nan")
    return float(values.mean())


# ---------------------------------------------------------------------------
# Column mapping: raw DataFrame → standard columns
# ---------------------------------------------------------------------------

_COLUMN_ALIASES: dict[str, list[str]] = {
    "tree":    ["tree", "tag", "id", "树号", "编号", "treeid", "no"],
    "species": ["species", "sp", "树种", "speciesname", "spname", "种"],
    "dbh":     ["dbh", "d", "胸径", "径", "dbhcm", "d1.3", "d13", "d130"],
    "height":  ["height", "h", "树高", "高", "ht"],
    "crownd":  ["crowndiameter", "crdiameter", "cd", "冠径", "冠幅", "crownd", "crownwidth"],
    "crownr":  ["crownradius", "cr", "冠半径", "冠幅半径", "crownr"],
    "x":       ["x", "coordx", "lon", "xcoord", "east", "coordinatex", "横坐标"],
    "y":       ["y", "coordy", "lat", "ycoord", "north", "coordinatey", "纵坐标"],
}


def _normalize_col(name) -> str:
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "", str(name or "").strip().lower())


def map_to_standard(df: pd.DataFrame) -> pd.DataFrame:
    """Map raw DataFrame columns to the standard schema:
    Tree, X, Y, DBH, Species, Height, CrownRadius.

    Supports named-column lookup (aliased) and positional fallback
    (Plot/Tag/SP/D/H/CR/X/Y in columns 0-7).
    """
    norm_cols = [_normalize_col(c) for c in df.columns]

    def find_col(key: str) -> int:
        for alias in _COLUMN_ALIASES[key]:
            for i, norm in enumerate(norm_cols):
                if norm == alias:
                    return i
        return -1

    i_tree = find_col("tree")
    i_x    = find_col("x")
    i_y    = find_col("y")
    i_dbh  = find_col("dbh")
    i_h    = find_col("height")
    i_sp   = find_col("species")
    i_cr   = find_col("crownr")
    i_cd   = find_col("crownd")

    has_required = i_tree >= 0 and i_x >= 0 and i_y >= 0 and i_dbh >= 0 and i_h >= 0

    tree_enc = CategoryEncoder()
    sp_enc   = CategoryEncoder()
    result_rows: list[dict] = []

    if has_required:
        for _, row in df.iterrows():
            tree_val = to_float(row.iloc[i_tree])
            if not math.isfinite(tree_val):
                tree_val = float(tree_enc.encode(str(row.iloc[i_tree] or "")))

            x_val   = to_float(row.iloc[i_x])
            y_val   = to_float(row.iloc[i_y])
            dbh_val = to_float(row.iloc[i_dbh])
            h_val   = to_float(row.iloc[i_h])

            if not all(math.isfinite(v) for v in [x_val, y_val, dbh_val, h_val]):
                continue

            sp_val = to_float_or_category(row.iloc[i_sp], sp_enc) if i_sp >= 0 else 0.0

            if i_cr >= 0:
                crown_val = to_float(row.iloc[i_cr])
            elif i_cd >= 0:
                crown_val = to_float(row.iloc[i_cd])
            else:
                crown_val = float("nan")

            if not math.isfinite(crown_val) or crown_val <= 0:
                # Empirical fallback: ~30 % of tree height is a typical crown radius
                crown_val = 0.3 * h_val

            result_rows.append(
                {"Tree": tree_val, "X": x_val, "Y": y_val, "DBH": dbh_val,
                 "Species": sp_val, "Height": h_val, "CrownRadius": crown_val}
            )

    elif len(df.columns) >= 8:
        # Positional fallback: Plot(0) Tag(1) SP(2) D(3) H(4) CR(5) X(6) Y(7)
        for _, row in df.iterrows():
            tree_val = to_float(row.iloc[1])
            if not math.isfinite(tree_val):
                tree_val = float(tree_enc.encode(str(row.iloc[1] or "")))

            sp_val  = to_float_or_category(row.iloc[2], sp_enc)
            dbh_val = to_float(row.iloc[3])
            h_val   = to_float(row.iloc[4])
            crown_val = to_float(row.iloc[5])
            x_val   = to_float(row.iloc[6])
            y_val   = to_float(row.iloc[7])

            if not all(math.isfinite(v) for v in [x_val, y_val, dbh_val, h_val]):
                continue

            if not math.isfinite(crown_val) or crown_val <= 0:
                # Empirical fallback: ~30 % of tree height is a typical crown radius
                crown_val = 0.3 * h_val

            result_rows.append(
                {"Tree": tree_val, "X": x_val, "Y": y_val, "DBH": dbh_val,
                 "Species": sp_val, "Height": h_val, "CrownRadius": crown_val}
            )

    else:
        raise ValueError(
            "无法识别列名/列顺序。支持: A) Tree/X/Y/DBH/Species/Height/[CrownRadius|CrownDiameter] "
            "或 B) Plot/Tag/SP/D/H/CR/X/Y"
        )

    if not result_rows:
        raise ValueError("标准化处理后没有有效数据行，请检查输入数据格式。")

    result_df = pd.DataFrame(result_rows)
    post_process_dbh_cm_to_m(result_df)
    return result_df


# ---------------------------------------------------------------------------
# Buffer zone annotation
# ---------------------------------------------------------------------------

def add_bbox_zone_columns(df: pd.DataFrame, buffer_size: float) -> pd.DataFrame:
    """Annotate each tree with distance-to-bounding-box-edge and CORE/BUFFER zone.

    Parameters
    ----------
    df : DataFrame with at least X and Y columns.
    buffer_size : edge width in metres (0 → all trees are CORE).
    """
    out = df.copy()
    x = out["X"].to_numpy(dtype=float)
    y = out["Y"].to_numpy(dtype=float)

    xmin, xmax = float(np.min(x)), float(np.max(x))
    ymin, ymax = float(np.min(y)), float(np.max(y))

    dist = np.minimum(
        np.minimum(x - xmin, xmax - x),
        np.minimum(y - ymin, ymax - y),
    )

    is_core = np.ones(len(out), dtype=bool) if buffer_size <= 0 else (dist >= buffer_size)
    # buffer_size <= 0 treats zero and negative values as "no buffer" → all trees are CORE

    out["BufferSize_m"]   = buffer_size
    out["DistToBoundary"] = dist
    out["IsCore"]         = is_core
    out["Zone"]           = np.where(is_core, "CORE", "BUFFER")
    return out


# ---------------------------------------------------------------------------
# Pipeline orchestration
# ---------------------------------------------------------------------------

@dataclass
class RunOptions:
    """Options for the stand-metrics computation pipeline."""
    use_wm: bool = True
    use_layer: bool = True
    use_uuci: bool = True
    use_cf: bool = True
    use_q: bool = True
    export_main: bool = True
    export_means: bool = True
    export_entropy: bool = True
    export_q_table: bool = True
    buffer_size: float = 2.0
    infl_k: float = 1.5


@dataclass
class PipelineResult:
    """Return value of :func:`execute_pipeline`."""
    standard_data: pd.DataFrame = field(default_factory=pd.DataFrame)
    computed_data: pd.DataFrame = field(default_factory=pd.DataFrame)
    main_preview: pd.DataFrame  = field(default_factory=pd.DataFrame)
    means: pd.DataFrame         = field(default_factory=pd.DataFrame)
    entropy: pd.DataFrame       = field(default_factory=pd.DataFrame)
    q_table: pd.DataFrame       = field(default_factory=pd.DataFrame)
    metric_list: list            = field(default_factory=list)
    export_paths: ExportPaths   = field(default_factory=ExportPaths)


def execute_pipeline(
    source_df: pd.DataFrame,
    input_path: str,
    out_dir: str,
    export_format: str,
    options: RunOptions,
    progress_callback: Optional[Callable[[int, str], None]] = None,
) -> PipelineResult:
    """Run the full stand-metrics pipeline and return a :class:`PipelineResult`.

    Parameters
    ----------
    source_df : raw input DataFrame (e.g. loaded from CSV / Excel).
    input_path : original file path (used for naming export files).
    out_dir : output directory; empty string → same folder as *input_path*.
    export_format : ``"CSV"`` or ``"TXT"``.
    options : computation and export flags.
    progress_callback : optional ``(pct: int, msg: str) -> None`` callable.
    """

    def _progress(pct: int, msg: str = "") -> None:
        if progress_callback is not None:
            progress_callback(pct, msg)

    _progress(10, "正在映射列名...")
    standard = map_to_standard(source_df)

    _progress(25, "正在标注缓冲区...")
    standard = add_bbox_zone_columns(standard, options.buffer_size)

    _progress(50, "正在计算结构参数...")
    computed = compute_all(standard, options.infl_k)

    entropy_df = pd.DataFrame({"Metric": [], "Weight": [], "Method": [], "Direction": []})
    q_table_df = pd.DataFrame({"Tree": [], "Q": []})
    q_all = float("nan")
    q_core = float("nan")
    q_buf = float("nan")

    if options.use_q:
        _progress(70, "正在计算熵权与综合质量指数...")
        entropy_df, q_table_df, q_all, q_core, q_buf = compute_entropy_and_q(computed)

    _progress(78, "正在计算林分均值...")
    means_df = build_stand_means(computed, q_all, q_core, q_buf)

    _progress(85, "正在生成主表预览...")
    main_preview = build_main_for_preview(
        computed,
        use_wm=options.use_wm,
        use_layer=options.use_layer,
        use_uuci=options.use_uuci,
        use_cf=options.use_cf,
        use_q=options.use_q,
    )

    metric_list = build_metric_list(
        use_wm=options.use_wm,
        use_layer=options.use_layer,
        use_uuci=options.use_uuci,
        use_cf=options.use_cf,
        use_q=options.use_q,
    )

    _progress(90, "正在导出结果文件...")
    export_paths = write_all(
        main_df=main_preview,
        means_df=means_df,
        entropy_df=entropy_df,
        q_table_df=q_table_df,
        input_path=input_path,
        out_dir=out_dir,
        export_format=export_format,
        write_main=options.export_main,
        write_means=options.export_means,
        write_entropy=options.export_entropy,
        write_q=options.export_q_table,
    )

    _progress(100, "完成。")
    return PipelineResult(
        standard_data=standard,
        computed_data=computed,
        main_preview=main_preview,
        means=means_df,
        entropy=entropy_df,
        q_table=q_table_df,
        metric_list=metric_list,
        export_paths=export_paths,
    )

