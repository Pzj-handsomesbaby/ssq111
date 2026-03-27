# -*- coding: utf-8 -*-

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Callable, Optional

import geopandas as gpd
import numpy as np
import pandas as pd
import rasterio
from rasterio.enums import Resampling
from rasterio.features import geometry_mask
from rasterio.warp import reproject

from modules.structure_metrics_core import (
    equal_weight,
    entropy_weight,
    normalize_weight_method,
    weight_method_display_name,
)

ProgressCallback = Optional[Callable[[int, str], None]]


ROLE_ALIASES = {
    "height": ["treeheight", "height", "tree_h", "树高", "gao"],
    "dbh": ["dbh", "diameter", "breastheightdiameter", "xiongjing", "胸径"],
    "age": ["standage", "age", "linling", "林龄", "龄"],
    "crown": ["crownwidth", "crowndiameter", "crownd", "crown", "guanfu", "冠幅", "冠径"],
}

BOUNDARY_ALIASES = ["boundary", "mask", "plot", "stand", "sample", "yangdi", "边界", "样地", "林分"]


@dataclass
class StandInputPaths:
    input_dir: str
    boundary_path: str
    height_path: str
    dbh_path: str
    age_path: str
    crown_path: Optional[str] = None


@dataclass
class StandQualityResult:
    input_paths: StandInputPaths
    output_raster_path: str
    output_weights_path: Optional[str]
    quality_array: np.ndarray
    preview_array: np.ndarray
    histogram_values: np.ndarray
    weights_table: pd.DataFrame
    summary_table: pd.DataFrame
    weight_method: str
    growth_rate_mean: float
    valid_pixel_count: int


def run_stand_quality_assessment(
    input_dir: str,
    out_dir: str,
    weight_method: str = "entropy",
    export_weights: bool = True,
    progress_callback: ProgressCallback = None,
) -> StandQualityResult:
    input_paths = scan_input_directory(input_dir)
    emit_progress(progress_callback, 10, "已识别输入文件，开始读取栅格…")

    normalized_weight_method = normalize_weight_method(weight_method)
    with rasterio.open(input_paths.height_path) as ref_src:
        ref_profile = ref_src.profile.copy()
        ref_transform = ref_src.transform
        ref_crs = ref_src.crs
        ref_shape = (ref_src.height, ref_src.width)
        height_array = _read_aligned_raster(input_paths.height_path, ref_shape, ref_transform, ref_crs)
        dbh_array = _read_aligned_raster(input_paths.dbh_path, ref_shape, ref_transform, ref_crs)
        age_array = _read_aligned_raster(input_paths.age_path, ref_shape, ref_transform, ref_crs)
        crown_array = (
            _read_aligned_raster(input_paths.crown_path, ref_shape, ref_transform, ref_crs)
            if input_paths.crown_path
            else None
        )

    emit_progress(progress_callback, 30, "正在生成林分边界掩膜…")
    boundary_mask = _build_boundary_mask(input_paths.boundary_path, ref_shape, ref_transform, ref_crs)

    valid_mask = boundary_mask & np.isfinite(height_array) & np.isfinite(dbh_array) & np.isfinite(age_array) & (age_array > 0)
    if crown_array is not None:
        valid_mask &= np.isfinite(crown_array)
    if not np.any(valid_mask):
        raise ValueError("边界范围内没有足够的有效像元，无法进行林分质量评价。")

    emit_progress(progress_callback, 50, "正在计算生长率与指标归一化…")
    growth_rate = np.full(ref_shape, np.nan, dtype=np.float32)
    growth_rate[valid_mask] = (height_array[valid_mask] / age_array[valid_mask]).astype(np.float32)

    factors = {
        "树高": _normalize_positive(height_array, valid_mask),
        "胸径": _normalize_positive(dbh_array, valid_mask),
        "生长率": _normalize_positive(growth_rate, valid_mask),
    }
    if crown_array is not None:
        factors["冠幅"] = _normalize_positive(crown_array, valid_mask)

    factor_matrix = np.column_stack([factor[valid_mask] for factor in factors.values()])
    if normalized_weight_method == "equal":
        weights = equal_weight(factor_matrix.shape[1])
    else:
        weights = entropy_weight(factor_matrix, ["pos"] * factor_matrix.shape[1])

    emit_progress(progress_callback, 70, f"正在计算{weight_method_display_name(normalized_weight_method)}林分质量栅格…")
    quality_array = np.full(ref_shape, np.nan, dtype=np.float32)
    weighted = np.zeros(valid_mask.sum(), dtype=np.float32)
    for idx, factor_name in enumerate(factors.keys()):
        weighted += factor_matrix[:, idx].astype(np.float32) * float(weights[idx])
    quality_array[valid_mask] = weighted * 100.0

    hist_values = quality_array[valid_mask].astype(float)
    preview_array = np.where(valid_mask, quality_array, np.nan).astype(np.float32)

    weights_table = pd.DataFrame(
        {
            "Metric": list(factors.keys()),
            "Weight": weights,
            "Method": [weight_method_display_name(normalized_weight_method)] * len(factors),
            "Direction": ["正向"] * len(factors),
        }
    )

    summary_table = pd.DataFrame(
        [
            {"Metric": "有效像元数", "Value": int(valid_mask.sum())},
            {"Metric": "平均生长率", "Value": float(np.nanmean(growth_rate[valid_mask]))},
            {"Metric": "平均质量指数", "Value": float(np.nanmean(hist_values))},
            {"Metric": "最大质量指数", "Value": float(np.nanmax(hist_values))},
            {"Metric": "最小质量指数", "Value": float(np.nanmin(hist_values))},
        ]
    )

    emit_progress(progress_callback, 85, "正在导出林分质量结果…")
    os.makedirs(out_dir or input_dir, exist_ok=True)
    output_dir = out_dir or input_dir
    base_name = os.path.basename(os.path.normpath(input_dir)) or "stand_quality"
    output_raster_path = os.path.join(output_dir, f"{base_name}_stand_quality_{normalized_weight_method}.tif")
    weights_path = os.path.join(output_dir, f"{base_name}_stand_quality_{normalized_weight_method}_weights.csv") if export_weights else None

    profile = ref_profile.copy()
    profile.update(dtype=rasterio.float32, count=1, nodata=np.nan, compress="lzw")
    with rasterio.open(output_raster_path, "w", **profile) as dst:
        dst.write(quality_array.astype(np.float32), 1)

    if weights_path:
        weights_table.to_csv(weights_path, index=False, encoding="utf-8-sig")

    emit_progress(progress_callback, 100, "林分质量评价完成")
    return StandQualityResult(
        input_paths=input_paths,
        output_raster_path=output_raster_path,
        output_weights_path=weights_path,
        quality_array=quality_array,
        preview_array=preview_array,
        histogram_values=hist_values,
        weights_table=weights_table,
        summary_table=summary_table,
        weight_method=normalized_weight_method,
        growth_rate_mean=float(np.nanmean(growth_rate[valid_mask])),
        valid_pixel_count=int(valid_mask.sum()),
    )


def scan_input_directory(input_dir: str) -> StandInputPaths:
    input_dir = os.path.abspath(input_dir)
    if not os.path.isdir(input_dir):
        raise ValueError(f"输入目录不存在：{input_dir}")

    tif_files: list[str] = []
    shp_files: list[str] = []
    for root, _, files in os.walk(input_dir):
        for file_name in files:
            full_path = os.path.join(root, file_name)
            lower = file_name.lower()
            if lower.endswith((".tif", ".tiff")):
                tif_files.append(full_path)
            elif lower.endswith(".shp"):
                shp_files.append(full_path)

    if not tif_files:
        raise ValueError("输入目录中未找到 tif/tiff 栅格文件。")
    if not shp_files:
        raise ValueError("输入目录中未找到 shp 边界文件。")

    role_matches: dict[str, str] = {}
    for tif_path in tif_files:
        role = _match_role(os.path.basename(tif_path), ROLE_ALIASES)
        if not role:
            continue
        if role in role_matches:
            raise ValueError(f"角色“{role}”匹配到多个栅格文件，请检查命名：{role_matches[role]} / {tif_path}")
        role_matches[role] = tif_path

    boundary_path = _pick_boundary_path(shp_files)
    missing = [role for role in ["height", "dbh", "age"] if role not in role_matches]
    if missing:
        missing_labels = {
            "height": "树高栅格",
            "dbh": "胸径栅格",
            "age": "年龄/林龄栅格",
        }
        readable = "、".join(missing_labels[item] for item in missing)
        raise ValueError(f"输入目录缺少必要文件：{readable}。请按参数命名 tif 文件。")

    return StandInputPaths(
        input_dir=input_dir,
        boundary_path=boundary_path,
        height_path=role_matches["height"],
        dbh_path=role_matches["dbh"],
        age_path=role_matches["age"],
        crown_path=role_matches.get("crown"),
    )


def emit_progress(callback: ProgressCallback, value: int, message: str):
    if callback is not None:
        callback(int(value), message)


def _match_role(file_name: str, aliases: dict[str, list[str]]) -> Optional[str]:
    stem = normalize_name(os.path.splitext(file_name)[0])
    candidates: list[tuple[int, str]] = []
    for role, role_aliases in aliases.items():
        for alias in sorted(role_aliases, key=len, reverse=True):
            normalized_alias = normalize_name(alias)
            if normalized_alias and normalized_alias in stem:
                candidates.append((len(normalized_alias), role))
                break
    if not candidates:
        return None
    candidates.sort(reverse=True)
    return candidates[0][1]


def _pick_boundary_path(shp_files: list[str]) -> str:
    if len(shp_files) == 1:
        return shp_files[0]
    matches = []
    for shp_path in shp_files:
        stem = normalize_name(os.path.splitext(os.path.basename(shp_path))[0])
        if any(normalize_name(alias) in stem for alias in BOUNDARY_ALIASES):
            matches.append(shp_path)
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise ValueError(f"检测到多个候选边界 shp，请只保留一个或调整命名：{matches}")
    raise ValueError(f"检测到多个 shp 文件且无法自动识别边界文件：{shp_files}")


def _read_aligned_raster(path: str, ref_shape, ref_transform, ref_crs) -> np.ndarray:
    with rasterio.open(path) as src:
        if src.count < 1:
            raise ValueError(f"栅格没有可用波段：{path}")
        array = src.read(1).astype(np.float32)
        nodata = src.nodata
        if nodata is not None:
            array[array == nodata] = np.nan
        if src.shape == ref_shape and src.transform == ref_transform and src.crs == ref_crs:
            return array

        dest = np.full(ref_shape, np.nan, dtype=np.float32)
        reproject(
            source=array,
            destination=dest,
            src_transform=src.transform,
            src_crs=src.crs,
            dst_transform=ref_transform,
            dst_crs=ref_crs,
            src_nodata=np.nan,
            dst_nodata=np.nan,
            resampling=Resampling.bilinear,
        )
        return dest


def _build_boundary_mask(boundary_path: str, ref_shape, ref_transform, ref_crs) -> np.ndarray:
    gdf = gpd.read_file(boundary_path)
    if gdf.empty:
        raise ValueError(f"边界 shp 为空：{boundary_path}")
    if gdf.crs is None:
        raise ValueError(f"边界 shp 缺少坐标系：{boundary_path}")
    if ref_crs is not None and gdf.crs != ref_crs:
        gdf = gdf.to_crs(ref_crs)
    mask = geometry_mask(gdf.geometry, out_shape=ref_shape, transform=ref_transform, invert=True)
    return mask.astype(bool)


def _normalize_positive(array: np.ndarray, valid_mask: np.ndarray) -> np.ndarray:
    normalized = np.full(array.shape, np.nan, dtype=np.float32)
    if not np.any(valid_mask):
        return normalized
    valid_values = array[valid_mask].astype(float)
    min_value = float(np.nanmin(valid_values))
    max_value = float(np.nanmax(valid_values))
    if not np.isfinite(min_value) or not np.isfinite(max_value):
        return normalized
    if max_value == min_value:
        normalized[valid_mask] = 1.0
    else:
        normalized[valid_mask] = ((valid_values - min_value) / (max_value - min_value)).astype(np.float32)
    return normalized


def normalize_name(value: str) -> str:
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "", str(value or "").strip().lower())

