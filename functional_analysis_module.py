import os
import sys
import traceback
import numpy as np
import pandas as pd

# ==================== 环境变量 ====================
if hasattr(sys, "prefix"):
    base = sys.prefix
    gdal_data = os.path.join(base, "Library", "share", "gdal")
    proj_lib = os.path.join(base, "Library", "share", "proj")
    qt_plugin = os.path.join(base, "Library", "plugins", "platforms")

    if os.path.exists(gdal_data):
        os.environ["GDAL_DATA"] = gdal_data
    if os.path.exists(proj_lib):
        os.environ["PROJ_LIB"] = proj_lib
    if os.path.exists(qt_plugin):
        os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = qt_plugin

import rasterio
from rasterio.enums import Resampling
from rasterio.transform import from_origin
from rasterio.io import MemoryFile
from rasterio.warp import reproject

import geopandas as gpd
from shapely.geometry import Point
try:
    from pykrige.ok import OrdinaryKriging
except Exception:
    OrdinaryKriging = None
from scipy.interpolate import Rbf

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QFileDialog, QMessageBox, QCheckBox, QComboBox,
    QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem,
    QSplitter, QGroupBox, QFormLayout, QStackedWidget, QHeaderView,
    QProgressBar, QGraphicsView, QGraphicsScene, QGraphicsPixmapItem,
    QDialog, QAbstractItemView, QScrollArea
)
from PyQt5.QtCore import Qt, QRectF
from PyQt5.QtGui import QPixmap

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


# ==================== 模块配置 ====================
MODULES = {
    "生态功能": {
        "水源涵养": {
            "森林调节水循环能力": {
                "result_name": "WATER",
                "desc": "森林调节水循环能力计算",
                "calc_type": "water_cycle",
                "need_formula": False,
                "need_age": False,
                "need_second_raster": True,
                "need_third_raster": True,
                "need_years": False,
                "need_rho_k": False,
                "need_cf": False,
                "need_comprehensive": False,
                "input1_name": "降水量 P 栅格",
                "input2_name": "蒸散量 E 栅格",
                "input3_name": "径流量 R 栅格",
                "formula_help": (
                    "公式说明：W = P - E - R\n"
                    "W：水源涵养量；P：降水量；E：蒸散量；R：径流量。"
                ),
                "route_help": (
                    "流程说明：输入降水量 P、蒸散量 E、径流量 R 三个栅格 → "
                    "可选重采样与对齐 → 按公式 W = P - E - R 计算 → "
                    "提值到点 → 可选插值（支持单独设置插值像元大小）→ "
                    "导出水源涵养结果栅格、点文件、属性表与 PNG。"
                )
            }
        },
        "水土保持": {
            "土壤侵蚀量": {
                "result_name": "SOIL_EROSION",
                "desc": "土壤侵蚀量计算",
                "calc_type": "soil_erosion",
                "need_formula": False,
                "need_age": False,
                "need_second_raster": False,
                "need_third_raster": False,
                "need_years": False,
                "need_rho_k": False,
                "need_cf": False,
                "need_comprehensive": False,
                "input1_name": "参考栅格",
                "input2_name": "第二输入栅格",
                "input3_name": "第三输入栅格",
                "formula_help": (
                    "公式说明：A = R × K × LS × C × P\n"
                    "A：土壤侵蚀量；R：降雨侵蚀力；K：土壤可蚀性；LS：坡度坡长；C：植被覆盖；P：水土保持措施。"
                ),
                "route_help": (
                    "流程说明：选择参考栅格 → 分别设置 R/K/LS/C/P 为栅格或常数 → "
                    "自动按参考栅格对齐全部因子 → 计算土壤侵蚀量 A = R × K × LS × C × P → "
                    "导出结果栅格、点文件、属性表和 PNG，并可继续插值。"
                )
            },
            "土壤保持量": {
                "result_name": "SOIL_RETENTION",
                "desc": "土壤保持量计算",
                "calc_type": "soil_retention",
                "need_formula": False,
                "need_age": False,
                "need_second_raster": True,
                "need_third_raster": False,
                "need_years": False,
                "need_rho_k": False,
                "need_cf": False,
                "need_comprehensive": False,
                "input1_name": "潜在侵蚀量栅格",
                "input2_name": "实际侵蚀量栅格",
                "input3_name": "第三输入栅格",
                "formula_help": (
                    "公式说明：SR = A_potential - A_actual\n"
                    "A_potential：潜在侵蚀量；A_actual：实际侵蚀量；SR：土壤保持量。"
                ),
                "route_help": (
                    "流程说明：输入潜在侵蚀量栅格与实际侵蚀量栅格 → "
                    "自动将实际侵蚀量对齐到潜在侵蚀量栅格 → 计算 SR = A_potential - A_actual → "
                    "导出结果栅格、点文件、属性表和 PNG，并可继续插值。"
                )
            }
        },
        "固碳功能": {
            "森林碳储量": {
                "result_name": "CARBON",
                "desc": "森林碳储量计算",
                "calc_type": "carbon_storage",
                "need_formula": False,
                "need_age": False,
                "need_second_raster": False,
                "need_third_raster": False,
                "need_years": False,
                "need_rho_k": False,
                "need_cf": True,
                "need_comprehensive": False,
                "input1_name": "AGB 生物量栅格",
                "input2_name": "后一期栅格",
                "input3_name": "第三输入栅格",
                "formula_help": (
                    "公式说明：C = AGB × CF\n"
                    "C：碳储量；AGB：生物量；CF：碳含量系数，常用 0.45–0.50。"
                ),
                "route_help": (
                    "流程说明：输入 AGB 生物量栅格 → 可选重采样 → 输入 CF（碳含量系数）→ "
                    "按 C = AGB × CF 计算碳储量 → 提值到点 → 可选插值生成碳储量栅格 → "
                    "导出结果栅格、点文件、属性表与 PNG。"
                )
            }
        }
    },
    "生产力功能": {
        "由AGB计算ANPP（生物量）": {
            "result_name": "ANPP",
            "default_formula": "",
            "desc": "由双期AGB增量计算ANPP",
            "calc_type": "agb_dual",
            "need_formula": False,
            "need_age": False,
            "need_second_raster": True,
            "need_third_raster": False,
            "need_years": True,
            "need_rho_k": False,
            "need_cf": False,
            "need_comprehensive": False,
            "input1_name": "前一期AGB栅格",
            "input2_name": "后一期AGB栅格",
            "input3_name": "第三输入栅格",
            "formula_help": "",
            "route_help": (
                "流程说明：输入前一期与后一期AGB栅格 → 可选重采样与对齐 → 输入两个时期年份 → "
                "计算 growth = AGB_t2 - AGB_t1 → 按 ANPP = growth / 年限差 计算 → "
                "提值到点 → 可选插值（支持单独设置插值像元大小）→ 生成属性表。"
            )
        },
        "由碳储量计算ANPP": {
            "result_name": "ANPP",
            "default_formula": "",
            "desc": "由双期碳储量增量计算ANPP",
            "calc_type": "carbon_dual",
            "need_formula": False,
            "need_age": False,
            "need_second_raster": True,
            "need_third_raster": False,
            "need_years": True,
            "need_rho_k": False,
            "need_cf": False,
            "need_comprehensive": False,
            "input1_name": "前一期碳储量栅格",
            "input2_name": "后一期碳储量栅格",
            "input3_name": "第三输入栅格",
            "formula_help": "",
            "route_help": (
                "流程说明：输入前一期与后一期碳储量栅格 → 可选重采样与对齐 → 输入两个时期年份 → "
                "计算 carbon_increment = C_t2 - C_t1 → 按 ANPP = (C_t2 - C_t1) / 年限差 计算。\n"
                "提示：当输入为碳储量时，ANPP 按碳储量增量计算，不乘含碳系数 CF。"
            )
        },
        "由蓄积量计算ANPP": {
            "result_name": "ANPP",
            "default_formula": "",
            "desc": "由双期蓄积增量计算ANPP",
            "calc_type": "volume_dual",
            "need_formula": False,
            "need_age": False,
            "need_second_raster": True,
            "need_third_raster": False,
            "need_years": True,
            "need_rho_k": True,
            "need_cf": False,
            "need_comprehensive": False,
            "input1_name": "前一期蓄积量栅格",
            "input2_name": "后一期蓄积量栅格",
            "input3_name": "第三输入栅格",
            "formula_help": "",
            "route_help": (
                "流程说明：输入前一期与后一期蓄积量栅格 → 可选重采样与对齐 → 输入两个时期年份 → "
                "手动输入ρ（木材密度）与k（树干到地上总生产扩展系数）→ "
                "计算 volume_increment = V_t2 - V_t1 → "
                "按 ANPP = (V_t2 - V_t1) / 年限差 × ρ × k 计算 → "
                "提值到点 → 可选插值（支持单独设置插值像元大小）→ 生成属性表。"
            )
        },
        "综合生产力计算": {
            "result_name": "CPI",
            "default_formula": "",
            "desc": "综合生产力计算（等权法 / 熵权法）",
            "calc_type": "comprehensive",
            "need_formula": False,
            "need_age": False,
            "need_second_raster": False,
            "need_third_raster": False,
            "need_years": False,
            "need_rho_k": False,
            "need_cf": False,
            "need_comprehensive": True,
            "input1_name": "输入栅格",
            "input2_name": "后一期栅格",
            "input3_name": "第三输入栅格",
            "formula_help": "",
            "route_help": (
                "流程说明：点击“设置综合生产力因子”弹窗 → 输入多个因子栅格（如碳储量、生物量、NDVI、LAI、AGB、CHM、降水、土壤湿度、坡度等）→ "
                "可选重采样 → 以第一个栅格为基准自动对齐其他栅格 → "
                "按正向/负向进行标准化 → 采用等权法或熵权法计算综合生产力指数 → "
                "提值到点 → 可选插值（支持单独设置插值像元大小）→ 生成属性表与权重表。"
            )
        }
    },
    "社会服务功能": {
        "森林游憩": {
            "旅游价值": {
                "result_name": "TOURISM_VALUE",
                "desc": "森林游憩-旅游价值计算",
                "route_help": (
                    "流程说明：导入 Excel 文件（字段应包含“序号”“旅游人数”“平均消费水平（元）”）→ "
                    "逐行计算 旅游人数 × 平均消费水平（元）→ 汇总求和 → "
                    "直接在界面中显示旅游价值金额，并可导出明细表。"
                ),
                "social_type": "tourism_excel"
            }
        },
        "文化教育": {
            "文化教育价值": {
                "result_name": "EDU_VALUE",
                "desc": "文化教育价值计算",
                "route_help": (
                    "流程说明：手动输入 WTP（人均支付意愿）与 N（愿意支付人数）→ "
                    "按公式 V = WTP × N 计算文化教育价值 → 直接在界面中显示结果。"
                ),
                "social_type": "culture_manual"
            }
        },
        "景观价值": {
            "景观与文化服务": {
                "result_name": "LANDSCAPE_VALUE",
                "desc": "景观与文化服务价值计算",
                "route_help": (
                    "流程说明：手动输入 A（森林景观面积）与 Pl（单位面积景观价值系数）→ "
                    "按公式 V = A × Pl 计算景观与文化服务价值 → 直接在界面中显示结果。"
                ),
                "social_type": "landscape_manual"
            }
        },
        "综合服务评价": {
            "等权法": {
                "result_name": "SOCIAL_EVAL_EQUAL",
                "desc": "综合服务评价（等权法）",
                "route_help": (
                    "流程说明：导入 Excel 文件（包含多个样本的旅游价值、文化教育价值、景观价值）→ "
                    "按正向指标进行极差标准化 → 采用等权法计算综合服务评价指数。"
                ),
                "social_type": "social_eval_equal"
            },
            "熵权法": {
                "result_name": "SOCIAL_EVAL_ENTROPY",
                "desc": "综合服务评价（熵权法）",
                "route_help": (
                    "流程说明：导入 Excel 文件（包含多个样本的旅游价值、文化教育价值、景观价值）→ "
                    "按正向指标进行极差标准化 → 基于多样本数据计算熵权 → 计算综合服务评价指数。"
                ),
                "social_type": "social_eval_entropy"
            }
        }
    }
}


# ==================== 预览控件 ====================
class ZoomableGraphicsView(QGraphicsView):
    def __init__(self):
        super().__init__()
        self.setScene(QGraphicsScene(self))
        self.pixmap_item = QGraphicsPixmapItem()
        self.scene().addItem(self.pixmap_item)

        self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorViewCenter)
        self.setBackgroundBrush(Qt.lightGray)
        self.setStyleSheet("background:#fafafa; border:1px solid #ccc;")
        self._has_image = False

    def set_pixmap(self, pixmap):
        self.scene().clear()
        self.pixmap_item = QGraphicsPixmapItem(pixmap)
        self.scene().addItem(self.pixmap_item)
        self.setSceneRect(QRectF(pixmap.rect()))
        self._has_image = not pixmap.isNull()
        self.reset_view()

    def clear_pixmap(self):
        self.scene().clear()
        self._has_image = False
        self.setSceneRect(QRectF())
        self.resetTransform()

    def reset_view(self):
        self.resetTransform()
        if self._has_image:
            self.fitInView(self.pixmap_item.boundingRect(), Qt.KeepAspectRatio)

    def wheelEvent(self, event):
        if not self._has_image:
            super().wheelEvent(event)
            return
        factor = 1.15 if event.angleDelta().y() > 0 else 1 / 1.15
        self.scale(factor, factor)

    def mouseDoubleClickEvent(self, event):
        if self._has_image:
            self.reset_view()
        super().mouseDoubleClickEvent(event)


# ==================== 工具函数 ====================
def parse_invalid_values(text):
    text = (text or "").strip()
    if not text:
        return []
    vals = []
    for s in text.split(","):
        s = s.strip()
        if not s:
            continue
        try:
            vals.append(float(s))
        except Exception:
            pass
    return vals


def evaluate_formula(x_array, formula_text):
    allowed = {
        "x": x_array,
        "np": np,
        "abs": np.abs,
        "sqrt": np.sqrt,
        "log": np.log,
        "log10": np.log10,
        "exp": np.exp,
        "sin": np.sin,
        "cos": np.cos,
        "tan": np.tan,
        "where": np.where,
        "minimum": np.minimum,
        "maximum": np.maximum,
        "clip": np.clip,
        "round": np.round
    }
    return np.array(eval(formula_text, {"__builtins__": {}}, allowed), dtype=np.float32)


def idw_interpolation(x, y, z, xi, yi, power=2):
    xi_grid, yi_grid = np.meshgrid(xi, yi)
    dist = np.sqrt((xi_grid[..., None] - x) ** 2 + (yi_grid[..., None] - y) ** 2)
    dist[dist == 0] = 1e-10
    weights = 1.0 / (dist ** power)
    weights /= weights.sum(axis=2, keepdims=True)
    return np.sum(weights * z, axis=2)


def rbf_interpolation(x, y, z, xi, yi):
    rbf = Rbf(x, y, z, function="linear")
    xi_grid, yi_grid = np.meshgrid(xi, yi)
    return rbf(xi_grid, yi_grid)


def sanitize_for_preview(arr, nodata=None, invalid_values=None):
    arr = np.array(arr, dtype=np.float32, copy=True)
    mask = np.isnan(arr)

    if nodata is not None:
        mask |= np.isclose(arr, nodata, equal_nan=False)

    if invalid_values:
        for v in invalid_values:
            mask |= np.isclose(arr, v, equal_nan=False)

    arr[mask] = np.nan
    return arr


def save_array_png(arr, png_path, title=None):
    arr = np.array(arr, dtype=np.float32)
    arr_masked = np.ma.masked_invalid(arr)

    if arr_masked.count() == 0:
        return None

    cmap = plt.cm.viridis.copy()
    cmap.set_bad(color=(1, 1, 1, 0))

    fig, ax = plt.subplots(figsize=(6, 5))
    ax.imshow(arr_masked, cmap=cmap, origin="upper")

    if title:
        ax.set_title(title)

    ax.set_xticks([])
    ax.set_yticks([])
    ax.set_frame_on(False)

    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    plt.savefig(
        png_path,
        dpi=150,
        bbox_inches="tight",
        pad_inches=0,
        transparent=True
    )
    plt.close(fig)
    return png_path


def load_table_to_widget(table_widget, df, max_rows=500):
    table_widget.clear()
    if df is None or df.empty:
        table_widget.setRowCount(0)
        table_widget.setColumnCount(0)
        return

    show_df = df.head(max_rows).copy()
    table_widget.setRowCount(show_df.shape[0])
    table_widget.setColumnCount(show_df.shape[1])
    table_widget.setHorizontalHeaderLabels([str(c) for c in show_df.columns])

    for i in range(show_df.shape[0]):
        for j in range(show_df.shape[1]):
            table_widget.setItem(i, j, QTableWidgetItem(str(show_df.iloc[i, j])))

    table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)


def safe_remove_file(path):
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


def safe_remove_shapefile(shp_path):
    if not shp_path:
        return
    base, _ = os.path.splitext(shp_path)
    for ext in [".shp", ".shx", ".dbf", ".prj", ".cpg", ".qix", ".fix", ".sbn", ".sbx", ".shp.xml"]:
        p = base + ext
        if os.path.exists(p):
            try:
                os.remove(p)
            except Exception:
                pass


def clear_output_targets(file_paths, shp_paths=None):
    shp_paths = shp_paths or []
    for p in file_paths:
        safe_remove_file(p)
    for shp in shp_paths:
        safe_remove_shapefile(shp)


def resample_clean_raster(clean_data, src_meta, new_h, new_w, out_nodata):
    temp_meta = src_meta.copy()
    temp_meta.update({
        "driver": "GTiff",
        "count": 1,
        "dtype": "float32",
        "nodata": out_nodata
    })

    temp_arr = np.where(np.isnan(clean_data), out_nodata, clean_data).astype(np.float32)

    with MemoryFile() as memfile:
        with memfile.open(**temp_meta) as mem:
            mem.write(temp_arr, 1)
            resampled = mem.read(
                1,
                masked=True,
                out_shape=(new_h, new_w),
                resampling=Resampling.bilinear
            )

    return np.array(resampled.filled(np.nan), dtype=np.float32)


def read_and_process_raster(path, do_resample, target_res, invalid_values, out_nodata):
    with rasterio.open(path) as src:
        raw_data = src.read(1).astype(np.float32)
        transform = src.transform
        crs = src.crs
        nodata = src.nodata
        res = src.res
        width = src.width
        height = src.height
        bounds = src.bounds
        meta = src.meta.copy()

    raw_clean = sanitize_for_preview(raw_data, nodata=nodata, invalid_values=invalid_values)

    work_transform = transform
    work_meta = meta.copy()

    if do_resample:
        scale_x = res[0] / target_res
        scale_y = abs(res[1]) / target_res
        new_w = max(1, int(round(width * scale_x)))
        new_h = max(1, int(round(height * scale_y)))

        work_data = resample_clean_raster(raw_clean, meta, new_h, new_w, out_nodata)
        work_transform = from_origin(bounds.left, bounds.top, target_res, target_res)
        work_meta.update({
            "height": new_h,
            "width": new_w,
            "transform": work_transform
        })
        clean_data = sanitize_for_preview(work_data, nodata=out_nodata, invalid_values=invalid_values)
    else:
        work_data = raw_clean.copy()
        clean_data = raw_clean.copy()

    return {
        "raw_data": raw_data,
        "raw_clean": raw_clean,
        "work_data": work_data,
        "clean_data": clean_data,
        "transform": work_transform,
        "meta": work_meta,
        "crs": crs,
        "nodata": nodata
    }


def align_raster_to_reference(path, ref_meta, ref_transform, ref_crs, invalid_values, out_nodata):
    with rasterio.open(path) as src:
        src_arr = src.read(1).astype(np.float32)
        src_nodata = src.nodata
        src_clean = sanitize_for_preview(src_arr, nodata=src_nodata, invalid_values=invalid_values)
        src_fill = np.where(np.isnan(src_clean), out_nodata, src_clean).astype(np.float32)

        dst = np.full((ref_meta["height"], ref_meta["width"]), out_nodata, dtype=np.float32)
        reproject(
            source=src_fill,
            destination=dst,
            src_transform=src.transform,
            src_crs=src.crs,
            dst_transform=ref_transform,
            dst_crs=ref_crs,
            src_nodata=out_nodata,
            dst_nodata=out_nodata,
            resampling=Resampling.bilinear
        )

    return sanitize_for_preview(dst, nodata=out_nodata, invalid_values=invalid_values)


def build_interp_grid(work_transform, work_meta, interp_res=None):
    left = work_transform.c
    top = work_transform.f
    base_xres = float(work_transform.a)
    base_yres = abs(float(work_transform.e))

    width = int(work_meta["width"])
    height = int(work_meta["height"])

    right = left + width * base_xres
    bottom = top - height * base_yres

    if interp_res is None or interp_res <= 0:
        interp_res = base_xres

    grid_width = max(1, int(np.ceil((right - left) / interp_res)))
    grid_height = max(1, int(np.ceil((top - bottom) / interp_res)))

    gridx = left + np.arange(grid_width) * interp_res + interp_res / 2.0
    gridy = top - np.arange(grid_height) * interp_res - interp_res / 2.0

    interp_transform = from_origin(left, top, interp_res, interp_res)

    return {
        "gridx": gridx,
        "gridy": gridy,
        "width": grid_width,
        "height": grid_height,
        "transform": interp_transform,
        "res": interp_res,
        "left": left,
        "top": top,
        "right": right,
        "bottom": bottom
    }


def build_valid_mask_for_interp(result_data, work_transform, interp_grid):
    orig_h, orig_w = result_data.shape
    a = float(work_transform.a)
    e = float(work_transform.e)
    c = float(work_transform.c)
    f = float(work_transform.f)

    xi, yi = np.meshgrid(interp_grid["gridx"], interp_grid["gridy"])

    col = np.floor((xi - c) / a).astype(int)
    row = np.floor((yi - f) / e).astype(int)

    in_range = (row >= 0) & (row < orig_h) & (col >= 0) & (col < orig_w)

    mask = np.zeros((interp_grid["height"], interp_grid["width"]), dtype=bool)
    valid_orig = ~np.isnan(result_data)

    rr = row[in_range]
    cc = col[in_range]
    mask[in_range] = valid_orig[rr, cc]

    return mask


# ==================== 综合生产力辅助函数 ====================
def is_negative_factor_name(name):
    n = (name or "").strip().lower()
    keywords = ["坡度", "slope", "干旱", "drought", "温度", "temperature", "地表温度", "lst"]
    return any(k in n for k in keywords)


def normalize_single_array(arr, direction):
    arr = np.array(arr, dtype=np.float32, copy=True)
    valid = np.isfinite(arr)
    out = np.full(arr.shape, np.nan, dtype=np.float32)

    if not np.any(valid):
        return out

    vmin = np.nanmin(arr)
    vmax = np.nanmax(arr)

    if np.isclose(vmax, vmin):
        out[valid] = 0.0
        return out

    if direction == "负向":
        out[valid] = (vmax - arr[valid]) / (vmax - vmin)
    else:
        out[valid] = (arr[valid] - vmin) / (vmax - vmin)

    return out.astype(np.float32)


def entropy_weights(norm_stack):
    n_factors = norm_stack.shape[0]
    flat = norm_stack.reshape(n_factors, -1).T
    valid_mask = np.all(np.isfinite(flat), axis=1)
    X = flat[valid_mask]

    if X.shape[0] == 0:
        raise ValueError("所有像元均为无效值，无法计算熵权。")

    eps = 1e-12
    X = np.clip(X, eps, None)

    col_sum = np.sum(X, axis=0, keepdims=True)
    col_sum[col_sum == 0] = eps
    P = X / col_sum

    m = X.shape[0]
    if m <= 1:
        return np.ones(n_factors, dtype=np.float32) / n_factors

    k = 1.0 / np.log(m)
    e = -k * np.sum(P * np.log(P), axis=0)
    d = 1.0 - e

    if np.allclose(np.sum(d), 0):
        weights = np.ones(n_factors, dtype=np.float32) / n_factors
    else:
        weights = d / np.sum(d)

    return weights.astype(np.float32)


def save_weights_csv(weights_path, factor_names, directions, weights):
    df = pd.DataFrame({
        "factor_name": factor_names,
        "direction": directions,
        "weight": weights
    })
    try:
        xlsx_path = weights_path.replace(".csv", ".xlsx")
        df.to_excel(xlsx_path, index=False)
        return xlsx_path, df
    except Exception:
        df.to_csv(weights_path, index=False, encoding="utf-8-sig")
        return weights_path, df


# ==================== 社会服务辅助函数 ====================
def parse_float(text, name="参数"):
    try:
        return float(str(text).strip())
    except Exception:
        raise ValueError(f"{name}输入不正确")


def minmax_normalize_series(series):
    s = pd.to_numeric(series, errors="coerce")
    vmin = s.min()
    vmax = s.max()
    if pd.isna(vmin) or pd.isna(vmax):
        return pd.Series(np.nan, index=s.index)
    if np.isclose(vmax, vmin):
        return pd.Series(0.0, index=s.index)
    return (s - vmin) / (vmax - vmin)


def entropy_weight_from_dataframe(norm_df):
    X = norm_df.to_numpy(dtype=np.float64)
    eps = 1e-12
    X = np.clip(X, eps, None)

    col_sum = X.sum(axis=0, keepdims=True)
    col_sum[col_sum == 0] = eps
    P = X / col_sum

    m = X.shape[0]
    if m <= 1:
        raise ValueError("熵权法至少需要 2 个以上样本。")

    k = 1.0 / np.log(m)
    e = -k * np.sum(P * np.log(P), axis=0)
    d = 1.0 - e

    if np.allclose(np.sum(d), 0):
        w = np.ones(X.shape[1], dtype=np.float64) / X.shape[1]
    else:
        w = d / np.sum(d)

    return w


# ==================== 综合生产力弹窗 ====================
class ComprehensiveFactorDialog(QDialog):
    def __init__(self, factors=None, method="等权法", result_name="CPI", parent=None):
        super().__init__(parent)
        self.setWindowTitle("综合生产力因子设置")
        self.resize(920, 580)

        self.factors = factors[:] if factors else []
        self.method = method
        self.result_name = result_name

        self.init_ui()
        self.load_existing()

    def init_ui(self):
        layout = QVBoxLayout(self)

        top_row = QHBoxLayout()
        top_row.addWidget(QLabel("计算方法"))
        self.method_combo = QComboBox()
        self.method_combo.addItems(["等权法", "熵权法"])
        self.method_combo.setCurrentText(self.method)
        top_row.addWidget(self.method_combo)

        top_row.addSpacing(20)
        top_row.addWidget(QLabel("结果名称"))
        self.result_name_edit = QLineEdit(self.result_name)
        top_row.addWidget(self.result_name_edit)
        top_row.addStretch()
        layout.addLayout(top_row)

        self.tip_label = QLabel(
            "提示：可输入不同因子如碳储量、生物量、NDVI、LAI、AGB、CHM、降水、土壤湿度、坡度等；"
            "若变量之间强相关，建议仅保留一种相似变量参与计算，以避免信息重复。"
            "建议输入空间分辨率相近的因子栅格；程序将以第一个栅格作为基准进行对齐。"
        )
        self.tip_label.setWordWrap(True)
        self.tip_label.setStyleSheet("""
            QLabel{
                color:red;
                font-size:13px;
                font-weight:bold;
                padding:4px;
            }
        """)
        layout.addWidget(self.tip_label)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["因子名称", "栅格路径", "指标方向"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.table)

        btn_row = QHBoxLayout()
        self.btn_add = QPushButton("添加因子栅格")
        self.btn_add.clicked.connect(self.add_rasters)
        btn_row.addWidget(self.btn_add)

        self.btn_remove = QPushButton("删除选中行")
        self.btn_remove.clicked.connect(self.remove_rows)
        btn_row.addWidget(self.btn_remove)

        btn_row.addStretch()
        layout.addLayout(btn_row)

        bottom_row = QHBoxLayout()
        bottom_row.addStretch()

        self.btn_ok = QPushButton("确定")
        self.btn_ok.clicked.connect(self.accept_data)
        bottom_row.addWidget(self.btn_ok)

        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.clicked.connect(self.reject)
        bottom_row.addWidget(self.btn_cancel)

        layout.addLayout(bottom_row)

    def load_existing(self):
        for fac in self.factors:
            self.append_row(fac["name"], fac["path"], fac["direction"])

    def append_row(self, name, path, direction="正向"):
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QTableWidgetItem(name))
        self.table.setItem(row, 1, QTableWidgetItem(path))

        combo = QComboBox()
        combo.addItems(["正向", "负向"])
        combo.setCurrentText(direction if direction in ["正向", "负向"] else "正向")
        self.table.setCellWidget(row, 2, combo)

    def add_rasters(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择因子栅格", "", "Raster Files (*.tif *.tiff *.img *.bil *.asc)"
        )
        if not files:
            return

        for fp in files:
            factor_name = os.path.splitext(os.path.basename(fp))[0]
            direction = "负向" if is_negative_factor_name(factor_name) else "正向"
            self.append_row(factor_name, fp, direction)

    def remove_rows(self):
        rows = sorted(set(index.row() for index in self.table.selectedIndexes()), reverse=True)
        for row in rows:
            self.table.removeRow(row)

    def accept_data(self):
        factors = []
        for r in range(self.table.rowCount()):
            name_item = self.table.item(r, 0)
            path_item = self.table.item(r, 1)
            dir_widget = self.table.cellWidget(r, 2)

            name = name_item.text().strip() if name_item else ""
            path = path_item.text().strip() if path_item else ""
            direction = dir_widget.currentText().strip() if dir_widget else "正向"

            if not name or not path:
                continue

            factors.append({
                "name": name,
                "path": path,
                "direction": direction
            })

        if len(factors) < 2:
            QMessageBox.warning(self, "提示", "至少需要添加 2 个因子栅格。")
            return

        self.factors = factors
        self.method = self.method_combo.currentText().strip()
        self.result_name = self.result_name_edit.text().strip() or "CPI"
        self.accept()

    def get_data(self):
        return self.factors, self.method, self.result_name


# ==================== 主窗口 ====================
class SoilFactorInput(QGroupBox):
    def __init__(self, title, default_value="1.0", parent=None):
        super().__init__(title, parent)
        layout = QFormLayout(self)

        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["栅格输入", "常数输入"])
        self.mode_combo.currentIndexChanged.connect(self.update_mode)

        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("选择栅格文件")
        self.browse_btn = QPushButton("选择")
        self.browse_btn.clicked.connect(self.choose_file)

        row = QHBoxLayout()
        row.addWidget(self.path_edit)
        row.addWidget(self.browse_btn)
        row_widget = QWidget()
        row_widget.setLayout(row)

        self.const_edit = QLineEdit(str(default_value))

        layout.addRow("输入方式", self.mode_combo)
        layout.addRow("栅格文件", row_widget)
        layout.addRow("常数值", self.const_edit)
        self.update_mode()

    def choose_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择栅格", "", "Raster Files (*.tif *.img *.asc);;All Files (*)")
        if path:
            self.path_edit.setText(path)

    def update_mode(self):
        raster_mode = self.mode_combo.currentText() == "栅格输入"
        self.path_edit.setEnabled(raster_mode)
        self.browse_btn.setEnabled(raster_mode)
        self.const_edit.setEnabled(not raster_mode)

    def use_raster(self):
        return self.mode_combo.currentText() == "栅格输入"

    def raster_path(self):
        return self.path_edit.text().strip()

    def constant_value(self):
        return float(self.const_edit.text().strip())


def read_raster_to_reference(input_path, ref_path, invalid_values=None, resampling=Resampling.bilinear):
    invalid_values = invalid_values or []
    with rasterio.open(ref_path) as ref:
        dst = np.full((ref.height, ref.width), np.nan, dtype=np.float32)
        with rasterio.open(input_path) as src:
            arr = src.read(1).astype(np.float32)
            if src.nodata is not None:
                arr[np.isclose(arr, src.nodata)] = np.nan
            for v in invalid_values:
                arr[np.isclose(arr, v)] = np.nan

            reproject(
                source=arr,
                destination=dst,
                src_transform=src.transform,
                src_crs=src.crs,
                dst_transform=ref.transform,
                dst_crs=ref.crs,
                src_nodata=np.nan,
                dst_nodata=np.nan,
                resampling=resampling
            )
    return dst


def create_constant_like_reference(ref_path, value):
    with rasterio.open(ref_path) as ref:
        return np.full((ref.height, ref.width), float(value), dtype=np.float32)


def choose_resampling_by_name(name):
    name = (name or "").strip().lower()
    categorical_keywords = [
        "landuse", "lulc", "type", "class", "p", "措施",
        "土地", "利用", "分类", "区划"
    ]
    for k in categorical_keywords:
        if k in name:
            return Resampling.nearest
    return Resampling.bilinear


def check_raster_validity(arr, name):
    if np.all(np.isnan(arr)):
        raise ValueError(f"{name} 全部为无效值，无法参与计算。")
    valid = arr[np.isfinite(arr)]
    if valid.size > 0 and np.isclose(np.nanmax(valid), np.nanmin(valid)):
        print(f"警告：{name} 的有效值几乎没有变化。")


def load_soil_factor_array(widget, ref_path, invalid_values=None, resampling_method=Resampling.bilinear):
    if widget.use_raster():
        p = widget.raster_path()
        if not p:
            raise ValueError(f"请为 {widget.title()} 选择栅格文件")
        if not os.path.exists(p):
            raise FileNotFoundError(f"文件不存在：{p}")
        return read_raster_to_reference(p, ref_path, invalid_values=invalid_values, resampling=resampling_method)
    return create_constant_like_reference(ref_path, widget.constant_value())


class FunctionalAnalysisWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("生态-生产力-社会服务功能计算平台")
        self.resize(1640, 1040)

        self.current_cfg = None
        self.current_top = None
        self.current_sub = None
        self.current_third = None
        self.result_records = []

        self.comp_factors = []
        self.comp_method = "等权法"
        self.comp_result_name = "CPI"

        self.init_ui()
        self.init_module_boxes()

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        top_bar = QGroupBox("功能模块")
        top_layout = QHBoxLayout(top_bar)

        self.top_combo = QComboBox()
        self.sub_combo = QComboBox()
        self.third_combo = QComboBox()
        self.third_combo.setMinimumWidth(320)

        self.top_combo.currentIndexChanged.connect(self.on_top_changed)
        self.sub_combo.currentIndexChanged.connect(self.on_sub_changed)
        self.third_combo.currentIndexChanged.connect(self.on_third_changed)

        top_layout.addWidget(QLabel("一级功能"))
        top_layout.addWidget(self.top_combo)
        top_layout.addSpacing(10)
        top_layout.addWidget(QLabel("二级功能"))
        top_layout.addWidget(self.sub_combo)
        top_layout.addSpacing(10)
        self.third_label = QLabel("三级功能")
        top_layout.addWidget(self.third_label)
        top_layout.addWidget(self.third_combo)
        top_layout.addSpacing(20)

        self.module_desc = QLabel("说明：")
        self.module_desc.setStyleSheet("color:#555;")
        top_layout.addWidget(self.module_desc, 1)

        root.addWidget(top_bar)

        splitter = QSplitter(Qt.Horizontal)
        root.addWidget(splitter, 1)

        # ==================== 左侧 ====================
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)

        io_group = QGroupBox("数据输入输出")
        io_form = QFormLayout(io_group)

        self.input1_label = QLabel("输入栅格")
        self.input_edit = QLineEdit()
        btn_input = QPushButton("选择")
        btn_input.clicked.connect(self.pick_input_raster)
        row1 = QHBoxLayout()
        row1.addWidget(self.input_edit)
        row1.addWidget(btn_input)
        self.input1_container = self.wrap_layout(row1)
        io_form.addRow(self.input1_label, self.input1_container)

        self.input2_label = QLabel("第二输入栅格")
        self.input2_edit = QLineEdit()
        btn_input2 = QPushButton("选择")
        btn_input2.clicked.connect(self.pick_input_raster2)
        row2 = QHBoxLayout()
        row2.addWidget(self.input2_edit)
        row2.addWidget(btn_input2)
        self.input2_container = self.wrap_layout(row2)
        io_form.addRow(self.input2_label, self.input2_container)

        self.input3_label = QLabel("第三输入栅格")
        self.input3_edit = QLineEdit()
        btn_input3 = QPushButton("选择")
        btn_input3.clicked.connect(self.pick_input_raster3)
        row3 = QHBoxLayout()
        row3.addWidget(self.input3_edit)
        row3.addWidget(btn_input3)
        self.input3_container = self.wrap_layout(row3)
        io_form.addRow(self.input3_label, self.input3_container)

        self.output_label = QLabel("结果保存目录")
        self.output_edit = QLineEdit()
        btn_output = QPushButton("选择")
        btn_output.clicked.connect(self.pick_output_dir)
        row4 = QHBoxLayout()
        row4.addWidget(self.output_edit)
        row4.addWidget(btn_output)
        self.output_container = self.wrap_layout(row4)
        io_form.addRow(self.output_label, self.output_container)

        self.res_row_label = QLabel("输入1分辨率")
        self.res_label = QLabel("未加载")
        io_form.addRow(self.res_row_label, self.res_label)

        self.res2_row_label = QLabel("输入2分辨率")
        self.res2_label = QLabel("未加载")
        io_form.addRow(self.res2_row_label, self.res2_label)

        self.res3_row_label = QLabel("输入3分辨率")
        self.res3_label = QLabel("未加载")
        io_form.addRow(self.res3_row_label, self.res3_label)

        left_layout.addWidget(io_group)

        self.soil_group = QGroupBox("水土保持因子设置")
        soil_form = QFormLayout(self.soil_group)

        self.soil_ref_label = QLabel("参考栅格")
        self.soil_ref_edit = QLineEdit()
        self.soil_ref_edit.setPlaceholderText("用于统一范围、分辨率和投影")
        self.soil_ref_btn = QPushButton("选择")
        self.soil_ref_btn.clicked.connect(self.pick_soil_ref_raster)
        soil_ref_row = QHBoxLayout()
        soil_ref_row.addWidget(self.soil_ref_edit)
        soil_ref_row.addWidget(self.soil_ref_btn)
        self.soil_ref_container = self.wrap_layout(soil_ref_row)
        soil_form.addRow(self.soil_ref_label, self.soil_ref_container)

        self.soil_r_widget = SoilFactorInput("R（降雨侵蚀力）", "1.0")
        self.soil_k_widget = SoilFactorInput("K（土壤可蚀性）", "1.0")
        self.soil_ls_widget = SoilFactorInput("LS（坡度坡长）", "1.0")
        self.soil_c_widget = SoilFactorInput("C（植被覆盖）", "1.0")
        self.soil_p_widget = SoilFactorInput("P（水土保持措施）", "1.0")

        soil_form.addRow(self.soil_r_widget)
        soil_form.addRow(self.soil_k_widget)
        soil_form.addRow(self.soil_ls_widget)
        soil_form.addRow(self.soil_c_widget)
        soil_form.addRow(self.soil_p_widget)
        left_layout.addWidget(self.soil_group)

        param_group = QGroupBox("条件设置")
        self.param_form = QFormLayout(param_group)

        self.chk_resample = QCheckBox("启用重采样")
        self.chk_resample.stateChanged.connect(self.on_resample_toggled)
        self.param_form.addRow(self.chk_resample)

        self.target_res_label = QLabel("目标分辨率（米）")
        self.target_res_edit = QLineEdit("3000")
        self.param_form.addRow(self.target_res_label, self.target_res_edit)

        self.formula_label = QLabel("计算公式")
        self.formula_edit = QLineEdit()
        self.param_form.addRow(self.formula_label, self.formula_edit)

        self.formula_help_title = QLabel("公式说明")
        self.formula_help_label = QLabel("")
        self.formula_help_label.setWordWrap(True)
        self.formula_help_label.setStyleSheet("color:#666;")
        self.param_form.addRow(self.formula_help_title, self.formula_help_label)

        self.age_label = QLabel("林分平均年龄")
        self.age_edit = QLineEdit("20")
        self.param_form.addRow(self.age_label, self.age_edit)

        self.year1_label = QLabel("前一期年份")
        self.year1_edit = QLineEdit("2018")
        self.param_form.addRow(self.year1_label, self.year1_edit)

        self.year2_label = QLabel("后一期年份")
        self.year2_edit = QLineEdit("2023")
        self.param_form.addRow(self.year2_label, self.year2_edit)

        self.rho_label = QLabel("ρ（木材密度）")
        self.rho_edit = QLineEdit("0.50")
        self.rho_edit.setToolTip("ρ 为木材密度。请按你的数据单位体系统一设置。常见取值大致约 0.30–0.90。")
        self.rho_edit.setPlaceholderText("例如 0.50")
        self.param_form.addRow(self.rho_label, self.rho_edit)

        self.k_label = QLabel("k（扩展系数）")
        self.k_edit = QLineEdit("1.30")
        self.k_edit.setToolTip("k 为树干到地上总生产的扩展系数，用于将树干生产扩展到地上总生产。常用范围约 1.1–1.6。")
        self.k_edit.setPlaceholderText("例如 1.30")
        self.param_form.addRow(self.k_label, self.k_edit)

        self.rhok_help_label = QLabel(
            "参数提示：ρ 为木材密度；k 为树干到地上总生产的扩展系数，常用范围一般约为 1.1–1.6。"
        )
        self.rhok_help_label.setWordWrap(True)
        self.rhok_help_label.setStyleSheet("color:#666;")
        self.param_form.addRow("参数说明", self.rhok_help_label)

        self.cf_label = QLabel("CF（碳含量系数）")
        self.cf_edit = QLineEdit("0.50")
        self.cf_edit.setPlaceholderText("常用 0.45–0.50")
        self.cf_edit.setToolTip("CF 为碳含量系数，常用范围通常约为 0.45–0.50。")
        self.param_form.addRow(self.cf_label, self.cf_edit)

        self.cf_help_label = QLabel(
            "参数提示：CF 为碳含量系数，森林碳储量常按 C = AGB × CF 计算，通常可取 0.45–0.50。"
        )
        self.cf_help_label.setWordWrap(True)
        self.cf_help_label.setStyleSheet("color:#666;")
        self.param_form.addRow("参数说明", self.cf_help_label)

        self.invalid_label = QLabel("无效值")
        self.invalid_edit = QLineEdit("-9999")
        self.param_form.addRow(self.invalid_label, self.invalid_edit)

        self.chk_interp = QCheckBox("启用插值")
        self.chk_interp.stateChanged.connect(self.on_interp_toggled)
        self.param_form.addRow(self.chk_interp)

        self.interp_method_label = QLabel("插值方式")
        self.interp_combo = QComboBox()
        self.interp_combo.addItems(["普通克里金", "反距离权重（IDW）", "径向基函数（RBF）"])
        self.param_form.addRow(self.interp_method_label, self.interp_combo)

        self.interp_res_label = QLabel("插值像元大小（米）")
        self.interp_res_edit = QLineEdit("")
        self.interp_res_edit.setPlaceholderText("留空则沿用当前结果栅格分辨率")
        self.param_form.addRow(self.interp_res_label, self.interp_res_edit)

        self.max_interp_points_label = QLabel("插值采样点数")
        self.max_interp_points_edit = QLineEdit("30")
        self.param_form.addRow(self.max_interp_points_label, self.max_interp_points_edit)

        left_layout.addWidget(param_group)

        # ==================== 社会服务功能设置 ====================
        self.social_group = QGroupBox("社会服务功能设置")
        social_layout = QFormLayout(self.social_group)

        self.social_excel_label = QLabel("导入Excel")
        self.social_excel_edit = QLineEdit()
        self.social_excel_btn = QPushButton("选择")
        self.social_excel_btn.clicked.connect(self.pick_social_excel)
        self.social_excel_help_btn = QPushButton("字段格式说明")
        self.social_excel_help_btn.clicked.connect(self.show_social_excel_format_help)

        social_excel_row = QHBoxLayout()
        social_excel_row.addWidget(self.social_excel_edit)
        social_excel_row.addWidget(self.social_excel_btn)
        social_excel_row.addWidget(self.social_excel_help_btn)
        self.social_excel_container = self.wrap_layout(social_excel_row)
        social_layout.addRow(self.social_excel_label, self.social_excel_container)

        self.culture_wtp_label = QLabel("WTP")
        self.culture_wtp_edit = QLineEdit()
        self.culture_wtp_edit.setPlaceholderText("人均支付意愿")
        social_layout.addRow(self.culture_wtp_label, self.culture_wtp_edit)

        self.culture_n_label = QLabel("N")
        self.culture_n_edit = QLineEdit()
        self.culture_n_edit.setPlaceholderText("愿意支付的人数")
        social_layout.addRow(self.culture_n_label, self.culture_n_edit)

        self.culture_help_btn = QPushButton("公式说明")
        self.culture_help_btn.clicked.connect(self.show_culture_help)
        social_layout.addRow("", self.culture_help_btn)

        self.landscape_a_label = QLabel("A")
        self.landscape_a_edit = QLineEdit()
        self.landscape_a_edit.setPlaceholderText("森林景观面积（ha 或 km²）")
        social_layout.addRow(self.landscape_a_label, self.landscape_a_edit)

        self.landscape_pl_label = QLabel("Pl")
        self.landscape_pl_edit = QLineEdit()
        self.landscape_pl_edit.setPlaceholderText("单位面积景观价值系数（元·ha⁻¹·年⁻¹）")
        social_layout.addRow(self.landscape_pl_label, self.landscape_pl_edit)

        self.landscape_help_btn = QPushButton("公式说明")
        self.landscape_help_btn.clicked.connect(self.show_landscape_help)
        social_layout.addRow("", self.landscape_help_btn)

        self.social_result_label = QLabel("结果显示")
        self.social_result_value = QLabel("未计算")
        self.social_result_value.setStyleSheet("""
            QLabel{
                color:#c0392b;
                font-size:18px;
                font-weight:bold;
                padding:6px;
            }
        """)
        social_layout.addRow(self.social_result_label, self.social_result_value)

        left_layout.addWidget(self.social_group)

        # ==================== 综合生产力设置 ====================
        self.comp_group = QGroupBox("综合生产力设置")
        comp_layout = QVBoxLayout(self.comp_group)

        self.comp_tip_summary = QLabel("尚未设置综合生产力因子")
        self.comp_tip_summary.setWordWrap(True)
        self.comp_tip_summary.setStyleSheet("color:#444;")
        comp_layout.addWidget(self.comp_tip_summary)

        self.btn_open_comp_dialog = QPushButton("设置综合生产力因子")
        self.btn_open_comp_dialog.clicked.connect(self.open_comp_dialog)
        comp_layout.addWidget(self.btn_open_comp_dialog)

        left_layout.addWidget(self.comp_group)

        export_group = QGroupBox("输出选项")
        export_layout = QVBoxLayout(export_group)

        self.chk_export_calc = QCheckBox("导出计算结果栅格")
        self.chk_export_calc.setChecked(True)

        self.chk_export_points = QCheckBox("导出提值点结果（shp）")
        self.chk_export_points.setChecked(True)

        self.chk_export_excel = QCheckBox("导出属性表 Excel/CSV")
        self.chk_export_excel.setChecked(True)

        self.chk_export_interp = QCheckBox("导出插值结果栅格")
        self.chk_export_interp.setChecked(True)

        self.chk_export_png = QCheckBox("导出栅格 PNG 预览图")
        self.chk_export_png.setChecked(True)

        export_layout.addWidget(self.chk_export_calc)
        export_layout.addWidget(self.chk_export_points)
        export_layout.addWidget(self.chk_export_excel)
        export_layout.addWidget(self.chk_export_interp)
        export_layout.addWidget(self.chk_export_png)

        left_layout.addWidget(export_group)

        self.run_btn = QPushButton("开始运行")
        self.run_btn.setStyleSheet("""
            QPushButton{
                background:#27ae60;
                color:white;
                font-weight:bold;
                height:42px;
                border-radius:6px;
            }
            QPushButton:hover{
                background:#219150;
            }
        """)
        self.run_btn.clicked.connect(self.run_workflow)
        left_layout.addWidget(self.run_btn)

        progress_group = QGroupBox("运行进度")
        progress_layout = QVBoxLayout(progress_group)

        self.status_label = QLabel("状态：等待运行")
        self.status_label.setStyleSheet("color:#444;")

        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")

        progress_layout.addWidget(self.status_label)
        progress_layout.addWidget(self.progress_bar)
        left_layout.addWidget(progress_group)

        self.route_group = QGroupBox("流程说明")
        route_layout = QVBoxLayout(self.route_group)
        self.route_help_label = QLabel("")
        self.route_help_label.setWordWrap(True)
        self.route_help_label.setStyleSheet("color:#444;")
        route_layout.addWidget(self.route_help_label)
        left_layout.addWidget(self.route_group)

        left_layout.addStretch()
        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setWidget(left_panel)
        splitter.addWidget(left_scroll)

        # ==================== 右侧 ====================
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)

        result_group = QGroupBox("本次输出结果")
        result_layout = QVBoxLayout(result_group)

        self.result_list = QListWidget()
        self.result_list.itemClicked.connect(self.on_result_item_clicked)
        result_layout.addWidget(self.result_list)

        right_layout.addWidget(result_group, 2)

        preview_group = QGroupBox("结果预览")
        preview_layout = QVBoxLayout(preview_group)

        self.preview_stack = QStackedWidget()

        self.image_view = ZoomableGraphicsView()
        self.preview_stack.addWidget(self.image_view)

        self.table_widget = QTableWidget()
        self.preview_stack.addWidget(self.table_widget)

        preview_layout.addWidget(self.preview_stack)
        right_layout.addWidget(preview_group, 4)

        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)
        right_scroll.setWidget(right_panel)
        splitter.addWidget(right_scroll)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)

    def wrap_layout(self, layout):
        w = QWidget()
        w.setLayout(layout)
        return w

    def update_progress(self, value, text):
        self.progress_bar.setValue(value)
        self.status_label.setText(f"状态：{text}")
        QApplication.processEvents()

    def init_module_boxes(self):
        self.top_combo.blockSignals(True)
        self.top_combo.clear()
        self.top_combo.addItems(list(MODULES.keys()))
        self.top_combo.blockSignals(False)
        self.on_top_changed()

    def is_leaf_cfg(self, obj):
        return isinstance(obj, dict) and ("desc" in obj or "social_type" in obj)

    def on_top_changed(self):
        top_name = self.top_combo.currentText()
        if not top_name:
            return

        self.current_top = top_name

        self.sub_combo.blockSignals(True)
        self.sub_combo.clear()
        self.third_combo.blockSignals(True)
        self.third_combo.clear()

        self.sub_combo.addItems(list(MODULES[top_name].keys()))

        self.sub_combo.blockSignals(False)
        self.third_combo.blockSignals(False)
        self.on_sub_changed()

    def on_sub_changed(self):
        top_name = self.top_combo.currentText()
        sub_name = self.sub_combo.currentText()
        if not top_name or not sub_name:
            return

        self.current_sub = sub_name
        node = MODULES[top_name][sub_name]

        if self.is_leaf_cfg(node):
            self.third_label.setVisible(False)
            self.third_combo.setVisible(False)
            self.current_third = None
            self.current_cfg = node
            self.module_desc.setText(f"说明：{self.current_cfg['desc']}")
            if self.current_cfg.get("need_formula", False):
                self.formula_edit.setText(self.current_cfg.get("default_formula", ""))
            else:
                self.formula_edit.setText("")
            if self.current_cfg.get("need_comprehensive", False):
                self.comp_result_name = self.current_cfg.get("result_name", "CPI")
            self.update_dynamic_fields()
        else:
            self.third_label.setVisible(True)
            self.third_combo.setVisible(True)
            self.third_combo.blockSignals(True)
            self.third_combo.clear()
            self.third_combo.addItems(list(node.keys()))
            self.third_combo.blockSignals(False)
            self.on_third_changed()

    def on_third_changed(self):
        top_name = self.top_combo.currentText()
        sub_name = self.sub_combo.currentText()
        third_name = self.third_combo.currentText()
        if not top_name or not sub_name or not third_name:
            return

        self.current_third = third_name
        self.current_cfg = MODULES[top_name][sub_name][third_name]
        self.module_desc.setText(f"说明：{self.current_cfg['desc']}")
        if self.current_cfg.get("need_formula", False):
            self.formula_edit.setText(self.current_cfg.get("default_formula", ""))
        else:
            self.formula_edit.setText("")
        self.update_dynamic_fields()

    def calc_type(self):
        if not self.current_cfg:
            return ""
        return self.current_cfg.get("calc_type", "")

    def is_agb_dual(self):
        return self.calc_type() == "agb_dual"

    def is_carbon_dual(self):
        return self.calc_type() == "carbon_dual"

    def is_volume_dual(self):
        return self.calc_type() == "volume_dual"

    def is_comprehensive(self):
        return bool(self.current_cfg and self.current_cfg.get("need_comprehensive", False))

    def is_social_service(self):
        return self.current_top == "社会服务功能"

    def is_water_cycle(self):
        return self.calc_type() == "water_cycle"

    def is_carbon_storage(self):
        return self.calc_type() == "carbon_storage"

    def is_soil_erosion(self):
        return self.calc_type() == "soil_erosion"

    def is_soil_retention(self):
        return self.calc_type() == "soil_retention"

    def is_soil_mode(self):
        return self.calc_type() == "soil_erosion"

    def is_soil_double_raster_mode(self):
        return self.calc_type() == "soil_retention"

    def update_dynamic_fields(self):
        cfg = self.current_cfg
        if cfg is None:
            return

        is_social = self.is_social_service()

        if not is_social:
            self.input1_label.setText(cfg.get("input1_name", "输入栅格"))
            self.input2_label.setText(cfg.get("input2_name", "第二输入栅格"))
            self.input3_label.setText(cfg.get("input3_name", "第三输入栅格"))

            need_formula = cfg.get("need_formula", False)
            need_age = cfg.get("need_age", False)
            need_second_raster = cfg.get("need_second_raster", False)
            need_third_raster = cfg.get("need_third_raster", False)
            need_years = cfg.get("need_years", False)
            need_rho_k = cfg.get("need_rho_k", False)
            need_cf = cfg.get("need_cf", False)
            need_comprehensive = cfg.get("need_comprehensive", False)
            soil_mode = self.is_soil_mode()
            soil_double_mode = self.is_soil_double_raster_mode()

            self.input1_label.setVisible((not need_comprehensive) and (not soil_mode))
            self.input1_container.setVisible((not need_comprehensive) and (not soil_mode))
            self.res_row_label.setVisible((not need_comprehensive) and (not soil_mode))
            self.res_label.setVisible((not need_comprehensive) and (not soil_mode))

            self.input2_label.setVisible(need_second_raster and not need_comprehensive and not soil_mode)
            self.input2_container.setVisible(need_second_raster and not need_comprehensive and not soil_mode)
            self.res2_row_label.setVisible(need_second_raster and not need_comprehensive and not soil_mode)
            self.res2_label.setVisible(need_second_raster and not need_comprehensive and not soil_mode)

            self.input3_label.setVisible(need_third_raster and not need_comprehensive and not soil_mode)
            self.input3_container.setVisible(need_third_raster and not need_comprehensive and not soil_mode)
            self.res3_row_label.setVisible(need_third_raster and not need_comprehensive and not soil_mode)
            self.res3_label.setVisible(need_third_raster and not need_comprehensive and not soil_mode)

            self.soil_group.setVisible(soil_mode)
            self.comp_group.setVisible(need_comprehensive)

            self.formula_label.setVisible(need_formula and not need_comprehensive and not soil_mode)
            self.formula_edit.setVisible(need_formula and not need_comprehensive and not soil_mode)
            self.formula_help_title.setVisible((need_formula or self.is_water_cycle() or self.is_carbon_storage() or soil_mode or soil_double_mode) and not need_comprehensive)
            self.formula_help_label.setVisible((need_formula or self.is_water_cycle() or self.is_carbon_storage() or soil_mode or soil_double_mode) and not need_comprehensive)
            self.formula_help_label.setText(cfg.get("formula_help", ""))

            self.age_label.setVisible(need_age and not need_comprehensive)
            self.age_edit.setVisible(need_age and not need_comprehensive)

            self.year1_label.setVisible(need_years and not need_comprehensive)
            self.year1_edit.setVisible(need_years and not need_comprehensive)
            self.year2_label.setVisible(need_years and not need_comprehensive)
            self.year2_edit.setVisible(need_years and not need_comprehensive)

            self.rho_label.setVisible(need_rho_k and not need_comprehensive)
            self.rho_edit.setVisible(need_rho_k and not need_comprehensive)
            self.k_label.setVisible(need_rho_k and not need_comprehensive)
            self.k_edit.setVisible(need_rho_k and not need_comprehensive)
            self.rhok_help_label.setVisible(need_rho_k and not need_comprehensive)

            self.cf_label.setVisible(need_cf and not need_comprehensive)
            self.cf_edit.setVisible(need_cf and not need_comprehensive)
            self.cf_help_label.setVisible(need_cf and not need_comprehensive)

            self.chk_resample.setVisible(not soil_mode)
            self.chk_interp.setVisible(True)
            self.invalid_label.setVisible(True)
            self.invalid_edit.setVisible(True)
            if soil_mode:
                self.target_res_label.setVisible(False)
                self.target_res_edit.setVisible(False)
            self.chk_export_calc.setVisible(True)
            self.chk_export_points.setVisible(True)
            self.chk_export_interp.setVisible(True)
            self.chk_export_png.setVisible(True)

            self.social_group.setVisible(False)

            self.route_help_label.setText(cfg.get("route_help", ""))
            self.on_resample_toggled()
            self.on_interp_toggled()

        else:
            self.input1_label.setVisible(False)
            self.input1_container.setVisible(False)
            self.res_row_label.setVisible(False)
            self.res_label.setVisible(False)

            self.input2_label.setVisible(False)
            self.input2_container.setVisible(False)
            self.res2_row_label.setVisible(False)
            self.res2_label.setVisible(False)

            self.input3_label.setVisible(False)
            self.input3_container.setVisible(False)
            self.res3_row_label.setVisible(False)
            self.res3_label.setVisible(False)

            self.comp_group.setVisible(False)
            self.soil_group.setVisible(False)

            self.formula_label.setVisible(False)
            self.formula_edit.setVisible(False)
            self.formula_help_title.setVisible(False)
            self.formula_help_label.setVisible(False)

            self.age_label.setVisible(False)
            self.age_edit.setVisible(False)
            self.year1_label.setVisible(False)
            self.year1_edit.setVisible(False)
            self.year2_label.setVisible(False)
            self.year2_edit.setVisible(False)
            self.rho_label.setVisible(False)
            self.rho_edit.setVisible(False)
            self.k_label.setVisible(False)
            self.k_edit.setVisible(False)
            self.rhok_help_label.setVisible(False)
            self.cf_label.setVisible(False)
            self.cf_edit.setVisible(False)
            self.cf_help_label.setVisible(False)

            self.chk_resample.setVisible(False)
            self.target_res_label.setVisible(False)
            self.target_res_edit.setVisible(False)
            self.chk_interp.setVisible(False)
            self.interp_method_label.setVisible(False)
            self.interp_combo.setVisible(False)
            self.interp_res_label.setVisible(False)
            self.interp_res_edit.setVisible(False)
            self.max_interp_points_label.setVisible(False)
            self.max_interp_points_edit.setVisible(False)
            self.invalid_label.setVisible(False)
            self.invalid_edit.setVisible(False)

            self.chk_export_calc.setVisible(False)
            self.chk_export_points.setVisible(False)
            self.chk_export_interp.setVisible(False)
            self.chk_export_png.setVisible(False)
            self.chk_export_excel.setVisible(True)

            social_type = self.current_cfg.get("social_type", "")

            self.social_group.setVisible(True)

            need_excel = social_type in ["tourism_excel", "social_eval_equal", "social_eval_entropy"]
            self.social_excel_label.setVisible(need_excel)
            self.social_excel_container.setVisible(need_excel)

            self.culture_wtp_label.setVisible(social_type == "culture_manual")
            self.culture_wtp_edit.setVisible(social_type == "culture_manual")
            self.culture_n_label.setVisible(social_type == "culture_manual")
            self.culture_n_edit.setVisible(social_type == "culture_manual")
            self.culture_help_btn.setVisible(social_type == "culture_manual")

            self.landscape_a_label.setVisible(social_type == "landscape_manual")
            self.landscape_a_edit.setVisible(social_type == "landscape_manual")
            self.landscape_pl_label.setVisible(social_type == "landscape_manual")
            self.landscape_pl_edit.setVisible(social_type == "landscape_manual")
            self.landscape_help_btn.setVisible(social_type == "landscape_manual")

            self.social_result_label.setVisible(True)
            self.social_result_value.setVisible(True)
            self.social_result_value.setText("未计算")

            self.route_help_label.setText(self.current_cfg.get("route_help", ""))

    def on_resample_toggled(self):
        show_resample = self.chk_resample.isChecked()
        self.target_res_label.setVisible(show_resample)
        self.target_res_edit.setVisible(show_resample)

    def on_interp_toggled(self):
        show_interp = self.chk_interp.isChecked()
        self.interp_method_label.setVisible(show_interp)
        self.interp_combo.setVisible(show_interp)
        self.interp_res_label.setVisible(show_interp)
        self.interp_res_edit.setVisible(show_interp)
        self.max_interp_points_label.setVisible(show_interp)
        self.max_interp_points_edit.setVisible(show_interp)

    def pick_input_raster(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择输入栅格", "", "Raster (*.tif *.tiff *.img)")
        if not path:
            return
        self.input_edit.setText(path)
        if not self.output_edit.text().strip():
            self.output_edit.setText(os.path.dirname(path))
        try:
            with rasterio.open(path) as src:
                self.res_label.setText(f"{src.res[0]} × {src.res[1]}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取栅格失败：\n{e}")

    def pick_input_raster2(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择第二输入栅格", "", "Raster (*.tif *.tiff *.img)")
        if not path:
            return
        self.input2_edit.setText(path)
        try:
            with rasterio.open(path) as src:
                self.res2_label.setText(f"{src.res[0]} × {src.res[1]}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取栅格失败：\n{e}")

    def pick_input_raster3(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择第三输入栅格", "", "Raster (*.tif *.tiff *.img)")
        if not path:
            return
        self.input3_edit.setText(path)
        try:
            with rasterio.open(path) as src:
                self.res3_label.setText(f"{src.res[0]} × {src.res[1]}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取栅格失败：\n{e}")

    def pick_output_dir(self):
        folder = QFileDialog.getExistingDirectory(self, "选择结果保存目录")
        if folder:
            self.output_edit.setText(folder)

    def pick_soil_ref_raster(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择水土保持参考栅格",
            "",
            "Raster (*.tif *.tiff *.img *.asc);;All Files (*)"
        )
        if not path:
            return
        self.soil_ref_edit.setText(path)
        if not self.output_edit.text().strip():
            self.output_edit.setText(os.path.dirname(path))

    def pick_social_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel",
            "",
            "Excel Files (*.xlsx *.xls);;CSV Files (*.csv)"
        )
        if path:
            self.social_excel_edit.setText(path)
            if not self.output_edit.text().strip():
                self.output_edit.setText(os.path.dirname(path))

    def show_social_excel_format_help(self):
        if not self.is_social_service() or not self.current_cfg:
            return

        social_type = self.current_cfg.get("social_type", "")

        if social_type == "tourism_excel":
            msg = (
                "旅游价值 Excel 字段格式说明：\n\n"
                "必需字段：\n"
                "1. 序号\n"
                "2. 旅游人数\n"
                "3. 平均消费水平（元）\n\n"
                "示例：\n"
                "序号 | 旅游人数 | 平均消费水平（元）\n"
                "1    | 1200    | 350\n"
                "2    | 900     | 420\n\n"
                "程序将自动计算：\n"
                "旅游价值（元） = 旅游人数 × 平均消费水平（元）"
            )
        else:
            msg = (
                "综合服务评价 Excel 字段格式说明：\n\n"
                "必需字段：\n"
                "1. 样本名称（或样本编号）\n"
                "2. 旅游价值\n"
                "3. 文化教育价值\n"
                "4. 景观价值\n\n"
                "字段名要求：\n"
                "样本名称、旅游价值、文化教育价值、景观价值\n\n"
                "示例：\n"
                "样本名称 | 旅游价值 | 文化教育价值 | 景观价值\n"
                "A区      | 120000  | 23000        | 56000\n"
                "B区      | 150000  | 28000        | 49000\n"
                "C区      | 98000   | 17000        | 61000\n\n"
                "程序将对三项指标先标准化，再按等权法或熵权法进行综合服务评价。"
            )

        QMessageBox.information(self, "字段格式说明", msg)

    def show_culture_help(self):
        QMessageBox.information(
            self,
            "文化教育公式说明",
            "文化教育价值公式：\n\nV = WTP × N\n\n"
            "其中：\n"
            "WTP：人均支付意愿\n"
            "N：愿意支付的人数"
        )

    def show_landscape_help(self):
        QMessageBox.information(
            self,
            "景观与文化服务公式说明",
            "景观与文化服务价值公式：\n\nV = A × Pl\n\n"
            "其中：\n"
            "A：森林景观面积（ha 或 km²）\n"
            "Pl：单位面积景观价值系数（元·ha⁻¹·年⁻¹）"
        )

    def clear_results(self):
        self.result_records = []
        self.result_list.clear()
        self.image_view.clear_pixmap()
        self.table_widget.clear()
        self.table_widget.setRowCount(0)
        self.table_widget.setColumnCount(0)

    def add_result_record(self, label, path=None, kind="file", data=None):
        self.result_records.append({"label": label, "path": path, "kind": kind, "data": data})
        self.result_list.addItem(QListWidgetItem(label))

    def on_result_item_clicked(self, item):
        idx = self.result_list.row(item)
        if idx < 0 or idx >= len(self.result_records):
            return

        record = self.result_records[idx]
        kind = record["kind"]
        path = record["path"]
        data = record["data"]

        if kind == "png" and path and os.path.exists(path):
            pix = QPixmap(path)
            if not pix.isNull():
                self.image_view.set_pixmap(pix)
                self.preview_stack.setCurrentIndex(0)
                return

        if kind == "table":
            load_table_to_widget(self.table_widget, data, max_rows=500)
            self.preview_stack.setCurrentIndex(1)
            return

        QMessageBox.information(self, "提示", "该结果当前不支持直接预览。")

    def open_comp_dialog(self):
        dlg = ComprehensiveFactorDialog(
            factors=self.comp_factors,
            method=self.comp_method,
            result_name=self.comp_result_name,
            parent=self
        )
        if dlg.exec_():
            self.comp_factors, self.comp_method, self.comp_result_name = dlg.get_data()

            factor_names = [f["name"] for f in self.comp_factors]
            preview_names = "、".join(factor_names[:4])
            if len(factor_names) > 4:
                preview_names += f" 等{len(factor_names)}个因子"

            self.comp_tip_summary.setText(
                f"已设置：方法={self.comp_method}；结果名={self.comp_result_name}；因子={preview_names}"
            )

            if self.comp_factors and not self.output_edit.text().strip():
                first_path = self.comp_factors[0]["path"]
                self.output_edit.setText(os.path.dirname(first_path))

    def load_comprehensive_factor_stack(self, factors, do_resample, target_res, invalid_values, out_nodata):
        ref_info = read_and_process_raster(factors[0]["path"], do_resample, target_res, invalid_values, out_nodata)
        ref_arr = ref_info["clean_data"].copy()
        ref_meta = ref_info["meta"].copy()
        ref_transform = ref_info["transform"]
        ref_crs = ref_info["crs"]

        arrays = [ref_arr]
        factor_names = [factors[0]["name"]]
        directions = [factors[0]["direction"]]
        aligned_dict = {factors[0]["name"]: ref_arr.copy()}

        for fac in factors[1:]:
            dst_clean = align_raster_to_reference(
                fac["path"], ref_meta, ref_transform, ref_crs, invalid_values, out_nodata
            )
            arrays.append(dst_clean)
            factor_names.append(fac["name"])
            directions.append(fac["direction"])
            aligned_dict[fac["name"]] = dst_clean.copy()

        stack = np.stack(arrays, axis=0).astype(np.float32)
        valid_mask = np.all(np.isfinite(stack), axis=0)
        stack[:, ~valid_mask] = np.nan

        ref_meta.update({
            "driver": "GTiff",
            "count": 1,
            "dtype": "float32",
            "nodata": out_nodata
        })

        return {
            "stack": stack,
            "meta": ref_meta,
            "transform": ref_transform,
            "crs": ref_crs,
            "factor_names": factor_names,
            "directions": directions,
            "aligned_dict": aligned_dict
        }

    # ==================== 社会服务功能运行 ====================
    def run_social_service_workflow(self):
        if self.current_cfg is None:
            QMessageBox.warning(self, "提示", "请先选择功能模块")
            return

        out_dir = self.output_edit.text().strip()
        if not out_dir:
            QMessageBox.warning(self, "提示", "请选择结果保存目录")
            return

        os.makedirs(out_dir, exist_ok=True)

        self.run_btn.setEnabled(False)
        self.clear_results()
        self.progress_bar.setValue(0)
        self.status_label.setText("状态：开始运行")
        QApplication.processEvents()

        try:
            cfg = self.current_cfg
            sub_name = self.current_sub
            third_name = self.current_third
            result_name = cfg.get("result_name", "RESULT")
            social_type = cfg.get("social_type", "")

            self.update_progress(10, "开始社会服务功能计算")

            if social_type == "tourism_excel":
                excel_path = self.social_excel_edit.text().strip()
                if not excel_path:
                    QMessageBox.warning(self, "提示", "请选择旅游价值Excel文件")
                    return

                self.update_progress(30, "读取旅游价值Excel")

                if excel_path.lower().endswith(".csv"):
                    df = pd.read_csv(excel_path)
                else:
                    df = pd.read_excel(excel_path)

                needed_cols = ["序号", "旅游人数", "平均消费水平（元）"]
                missing = [c for c in needed_cols if c not in df.columns]
                if missing:
                    QMessageBox.warning(self, "提示", f"Excel缺少必要字段：{', '.join(missing)}")
                    return

                df = df.copy()
                df["旅游人数"] = pd.to_numeric(df["旅游人数"], errors="coerce")
                df["平均消费水平（元）"] = pd.to_numeric(df["平均消费水平（元）"], errors="coerce")
                df["旅游价值（元）"] = df["旅游人数"] * df["平均消费水平（元）"]

                total_value = float(df["旅游价值（元）"].fillna(0).sum())
                self.social_result_value.setText(f"{total_value:,.2f} 元")

                out_xlsx = os.path.join(out_dir, f"{result_name}_旅游价值明细.xlsx")
                out_csv = os.path.join(out_dir, f"{result_name}_旅游价值明细.csv")

                try:
                    df.to_excel(out_xlsx, index=False)
                except Exception:
                    df.to_csv(out_csv, index=False, encoding="utf-8-sig")

                self.add_result_record(f"{sub_name}-{third_name} -> 旅游价值明细表", kind="table", data=df)

            elif social_type == "culture_manual":
                self.update_progress(40, "计算文化教育价值")

                wtp = parse_float(self.culture_wtp_edit.text(), "WTP")
                n = parse_float(self.culture_n_edit.text(), "N")

                if wtp < 0 or n < 0:
                    QMessageBox.warning(self, "提示", "WTP和N不能为负数")
                    return

                value = wtp * n
                self.social_result_value.setText(f"{value:,.2f} 元")

                df = pd.DataFrame([{
                    "WTP": wtp,
                    "N": n,
                    "文化教育价值（元）": value
                }])

                out_xlsx = os.path.join(out_dir, f"{result_name}_文化教育价值.xlsx")
                out_csv = os.path.join(out_dir, f"{result_name}_文化教育价值.csv")
                try:
                    df.to_excel(out_xlsx, index=False)
                except Exception:
                    df.to_csv(out_csv, index=False, encoding="utf-8-sig")

                self.add_result_record(f"{sub_name}-{third_name} -> 文化教育价值结果表", kind="table", data=df)

            elif social_type == "landscape_manual":
                self.update_progress(40, "计算景观与文化服务价值")

                a = parse_float(self.landscape_a_edit.text(), "A")
                pl = parse_float(self.landscape_pl_edit.text(), "Pl")

                if a < 0 or pl < 0:
                    QMessageBox.warning(self, "提示", "A和Pl不能为负数")
                    return

                value = a * pl
                self.social_result_value.setText(f"{value:,.2f} 元")

                df = pd.DataFrame([{
                    "A": a,
                    "Pl": pl,
                    "景观与文化服务价值（元）": value
                }])

                out_xlsx = os.path.join(out_dir, f"{result_name}_景观与文化服务价值.xlsx")
                out_csv = os.path.join(out_dir, f"{result_name}_景观与文化服务价值.csv")
                try:
                    df.to_excel(out_xlsx, index=False)
                except Exception:
                    df.to_csv(out_csv, index=False, encoding="utf-8-sig")

                self.add_result_record(f"{sub_name}-{third_name} -> 景观与文化服务价值结果表", kind="table", data=df)

            elif social_type in ["social_eval_equal", "social_eval_entropy"]:
                excel_path = self.social_excel_edit.text().strip()
                if not excel_path:
                    QMessageBox.warning(self, "提示", "请选择综合服务评价Excel文件")
                    return

                self.update_progress(30, "读取综合服务评价Excel")

                if excel_path.lower().endswith(".csv"):
                    df = pd.read_csv(excel_path)
                else:
                    df = pd.read_excel(excel_path)

                needed_cols = ["样本名称", "旅游价值", "文化教育价值", "景观价值"]
                missing = [c for c in needed_cols if c not in df.columns]
                if missing:
                    QMessageBox.warning(self, "提示", f"Excel缺少必要字段：{', '.join(missing)}")
                    return

                work_df = df.copy()
                for col in ["旅游价值", "文化教育价值", "景观价值"]:
                    work_df[col] = pd.to_numeric(work_df[col], errors="coerce")

                if work_df[["旅游价值", "文化教育价值", "景观价值"]].isna().any().any():
                    QMessageBox.warning(self, "提示", "综合服务评价 Excel 中存在无法识别的数值，请检查三项指标字段。")
                    return

                self.update_progress(55, "执行标准化")
                norm_df = pd.DataFrame({
                    "样本名称": work_df["样本名称"],
                    "旅游价值_标准化": minmax_normalize_series(work_df["旅游价值"]),
                    "文化教育价值_标准化": minmax_normalize_series(work_df["文化教育价值"]),
                    "景观价值_标准化": minmax_normalize_series(work_df["景观价值"])
                })

                eval_cols = ["旅游价值_标准化", "文化教育价值_标准化", "景观价值_标准化"]

                if social_type == "social_eval_equal":
                    weights = np.array([1 / 3, 1 / 3, 1 / 3], dtype=np.float64)
                    method_name = "等权法"
                else:
                    if len(norm_df) < 2:
                        QMessageBox.warning(self, "提示", "熵权法至少需要 2 个以上样本。")
                        return
                    weights = entropy_weight_from_dataframe(norm_df[eval_cols])
                    method_name = "熵权法"

                self.update_progress(75, f"按{method_name}计算综合服务评价")
                norm_df["综合服务评价指数"] = (
                    norm_df["旅游价值_标准化"] * weights[0] +
                    norm_df["文化教育价值_标准化"] * weights[1] +
                    norm_df["景观价值_标准化"] * weights[2]
                )

                result_df = work_df.copy()
                result_df["旅游价值_标准化"] = norm_df["旅游价值_标准化"]
                result_df["文化教育价值_标准化"] = norm_df["文化教育价值_标准化"]
                result_df["景观价值_标准化"] = norm_df["景观价值_标准化"]
                result_df["旅游价值权重"] = weights[0]
                result_df["文化教育价值权重"] = weights[1]
                result_df["景观价值权重"] = weights[2]
                result_df["综合服务评价指数"] = norm_df["综合服务评价指数"]

                score_mean = float(result_df["综合服务评价指数"].mean())
                self.social_result_value.setText(f"平均评价指数：{score_mean:.4f}")

                weight_df = pd.DataFrame({
                    "指标": ["旅游价值", "文化教育价值", "景观价值"],
                    "权重": weights
                })

                out_xlsx = os.path.join(out_dir, f"{result_name}_{method_name}_综合服务评价结果.xlsx")
                out_csv = os.path.join(out_dir, f"{result_name}_{method_name}_综合服务评价结果.csv")
                out_weight_xlsx = os.path.join(out_dir, f"{result_name}_{method_name}_权重表.xlsx")
                out_weight_csv = os.path.join(out_dir, f"{result_name}_{method_name}_权重表.csv")

                try:
                    result_df.to_excel(out_xlsx, index=False)
                except Exception:
                    result_df.to_csv(out_csv, index=False, encoding="utf-8-sig")

                try:
                    weight_df.to_excel(out_weight_xlsx, index=False)
                except Exception:
                    weight_df.to_csv(out_weight_csv, index=False, encoding="utf-8-sig")

                self.add_result_record(f"{sub_name}-{third_name} -> 综合服务评价结果表", kind="table", data=result_df)
                self.add_result_record(f"{sub_name}-{third_name} -> 权重表", kind="table", data=weight_df)

            self.update_progress(100, "运行完成")

            if self.result_list.count() > 0:
                self.result_list.setCurrentRow(0)
                self.on_result_item_clicked(self.result_list.item(0))

            QMessageBox.information(self, "完成", "社会服务功能计算完成，结果已显示在界面中。")

        except Exception as e:
            self.progress_bar.setValue(0)
            self.status_label.setText("状态：运行失败")
            QMessageBox.critical(
                self,
                "错误",
                f"运行失败：\n{str(e)}\n\n{traceback.format_exc()}"
            )
        finally:
            self.run_btn.setEnabled(True)

    # ==================== 栅格主流程 ====================
    def run_workflow(self):
        if self.current_cfg is None:
            QMessageBox.warning(self, "提示", "请先选择功能模块")
            return

        if self.is_social_service():
            self.run_social_service_workflow()
            return

        out_dir = self.output_edit.text().strip()
        cfg = self.current_cfg

        need_formula = cfg.get("need_formula", False)
        need_second_raster = cfg.get("need_second_raster", False)
        need_third_raster = cfg.get("need_third_raster", False)
        need_years = cfg.get("need_years", False)
        need_age = cfg.get("need_age", False)
        need_rho_k = cfg.get("need_rho_k", False)
        need_cf = cfg.get("need_cf", False)
        need_comprehensive = cfg.get("need_comprehensive", False)

        in_path = self.input_edit.text().strip()
        in_path2 = self.input2_edit.text().strip()
        in_path3 = self.input3_edit.text().strip()
        formula = self.formula_edit.text().strip() if need_formula else None

        if not out_dir:
            QMessageBox.warning(self, "提示", "请选择结果保存目录")
            return

        if self.is_soil_mode():
            soil_ref = self.soil_ref_edit.text().strip()
            if not soil_ref:
                QMessageBox.warning(self, "提示", "请选择水土保持计算的参考栅格")
                return
            if not os.path.exists(soil_ref):
                QMessageBox.warning(self, "提示", "参考栅格不存在")
                return
            for factor_widget in [self.soil_r_widget, self.soil_k_widget, self.soil_ls_widget, self.soil_c_widget, self.soil_p_widget]:
                if factor_widget.use_raster():
                    fp = factor_widget.raster_path()
                    if not fp:
                        QMessageBox.warning(self, "提示", f"请为 {factor_widget.title()} 选择栅格")
                        return
                    if not os.path.exists(fp):
                        QMessageBox.warning(self, "提示", f"文件不存在：{fp}")
                        return
                else:
                    try:
                        float(factor_widget.const_edit.text().strip())
                    except Exception:
                        QMessageBox.warning(self, "提示", f"{factor_widget.title()} 的常数值输入不正确")
                        return
        elif self.is_soil_double_raster_mode():
            if not in_path:
                QMessageBox.warning(self, "提示", "请选择潜在侵蚀量栅格")
                return
            if not os.path.exists(in_path):
                QMessageBox.warning(self, "提示", f"文件不存在：{in_path}")
                return
            if not in_path2:
                QMessageBox.warning(self, "提示", "请选择实际侵蚀量栅格")
                return
            if not os.path.exists(in_path2):
                QMessageBox.warning(self, "提示", f"文件不存在：{in_path2}")
                return
        elif not need_comprehensive:
            if not in_path:
                QMessageBox.warning(self, "提示", f"请选择{cfg.get('input1_name', '输入栅格')}")
                return
            if need_second_raster and not in_path2:
                QMessageBox.warning(self, "提示", f"请选择{cfg.get('input2_name', '第二输入栅格')}")
                return
            if need_third_raster and not in_path3:
                QMessageBox.warning(self, "提示", f"请选择{cfg.get('input3_name', '第三输入栅格')}")
                return
            if need_formula and not formula:
                QMessageBox.warning(self, "提示", "请输入公式")
                return
        else:
            factors = self.comp_factors
            if len(factors) < 2:
                QMessageBox.warning(self, "提示", "请先点击“设置综合生产力因子”并添加至少 2 个栅格。")
                return

        age = None
        year1 = None
        year2 = None
        year_diff = None
        rho = None
        k = None
        cf = None

        if need_age:
            try:
                age = float(self.age_edit.text().strip())
                if age <= 0:
                    QMessageBox.warning(self, "提示", "林分平均年龄必须大于 0")
                    return
            except Exception:
                QMessageBox.warning(self, "提示", "林分平均年龄输入不正确")
                return

        if need_years:
            try:
                year1 = float(self.year1_edit.text().strip())
                year2 = float(self.year2_edit.text().strip())
                year_diff = year2 - year1
                if year_diff <= 0:
                    QMessageBox.warning(self, "提示", "后一期年份必须大于前一期年份")
                    return
            except Exception:
                QMessageBox.warning(self, "提示", "年份输入不正确")
                return

        if need_rho_k:
            try:
                rho = float(self.rho_edit.text().strip())
                if rho <= 0:
                    QMessageBox.warning(self, "提示", "ρ（木材密度）必须大于 0")
                    return
            except Exception:
                QMessageBox.warning(self, "提示", "ρ（木材密度）输入不正确")
                return

            try:
                k = float(self.k_edit.text().strip())
                if k <= 0:
                    QMessageBox.warning(self, "提示", "k（扩展系数）必须大于 0")
                    return
            except Exception:
                QMessageBox.warning(self, "提示", "k（扩展系数）输入不正确")
                return

        if need_cf:
            try:
                cf = float(self.cf_edit.text().strip())
                if cf <= 0:
                    QMessageBox.warning(self, "提示", "CF（碳含量系数）必须大于 0")
                    return
            except Exception:
                QMessageBox.warning(self, "提示", "CF（碳含量系数）输入不正确")
                return

        try:
            self.run_btn.setEnabled(False)
            self.progress_bar.setValue(0)
            self.update_progress(1, "开始运行")
            self.clear_results()

            invalid_values = parse_invalid_values(self.invalid_edit.text())

            try:
                max_interp_points = int(float(self.max_interp_points_edit.text().strip()))
                if max_interp_points <= 0:
                    max_interp_points = 30
            except Exception:
                max_interp_points = 30

            interp_res = None
            if self.chk_interp.isChecked():
                interp_res_text = self.interp_res_edit.text().strip()
                if interp_res_text:
                    try:
                        interp_res = float(interp_res_text)
                        if interp_res <= 0:
                            QMessageBox.warning(self, "提示", "插值像元大小必须大于 0")
                            return
                    except Exception:
                        QMessageBox.warning(self, "提示", "插值像元大小输入不正确")
                        return

            do_resample = self.chk_resample.isChecked() and (not self.is_soil_mode())
            target_res = None
            if do_resample:
                try:
                    target_res = float(self.target_res_edit.text().strip())
                    if target_res <= 0:
                        QMessageBox.warning(self, "提示", "目标分辨率必须大于 0")
                        return
                except Exception:
                    QMessageBox.warning(self, "提示", "目标分辨率输入不正确")
                    return

            use_interp = self.chk_interp.isChecked()
            interp_method = self.interp_combo.currentText() if use_interp else None
            if interp_method == "普通克里金":
                interp_tag = "kriging"
            elif interp_method == "反距离权重（IDW）":
                interp_tag = "idw"
            elif interp_method == "径向基函数（RBF）":
                interp_tag = "rbf"
            else:
                interp_tag = "interp"

            os.makedirs(out_dir, exist_ok=True)
            out_nodata = -9999.0

            rst1 = None
            rst2 = None
            rst3 = None
            factor_names = []
            directions = []
            weights = None
            weights_df = None
            aligned_factor_dict = {}
            result_data = None

            if self.is_soil_mode():
                soil_ref = self.soil_ref_edit.text().strip()
                input_base = os.path.splitext(os.path.basename(soil_ref))[0] + "_soil"
                result_name = cfg["result_name"]
            elif self.is_soil_double_raster_mode():
                input_base = f"{os.path.splitext(os.path.basename(in_path))[0]}_to_{os.path.splitext(os.path.basename(in_path2))[0]}"
                result_name = cfg["result_name"]
            elif not need_comprehensive:
                input_base = os.path.splitext(os.path.basename(in_path))[0]
                if self.is_water_cycle():
                    input_base = f"{os.path.splitext(os.path.basename(in_path))[0]}_P_E_R"
                elif need_second_raster:
                    input_base2 = os.path.splitext(os.path.basename(in_path2))[0]
                    input_base = f"{input_base}_to_{input_base2}"
                result_name = cfg["result_name"]
            else:
                factors = self.comp_factors
                input_base = "comprehensive_productivity"
                result_name = self.comp_result_name.strip() or cfg["result_name"]

            resample_raster_path = os.path.join(out_dir, f"{input_base}_resample.tif")
            resample_raster2_path = os.path.join(out_dir, f"{input_base}_resample_t2.tif")
            resample_raster3_path = os.path.join(out_dir, f"{input_base}_resample_t3.tif")
            resample_png_path = os.path.join(out_dir, f"{input_base}_resample.png")
            resample_png2_path = os.path.join(out_dir, f"{input_base}_resample_t2.png")
            resample_png3_path = os.path.join(out_dir, f"{input_base}_resample_t3.png")
            result_raster_path = os.path.join(out_dir, f"{input_base}_{result_name}_result.tif")
            result_png_path = os.path.join(out_dir, f"{input_base}_{result_name}_result.png")
            points_path = os.path.join(out_dir, f"{input_base}_{result_name}_points.shp")
            excel_path = os.path.join(out_dir, f"{input_base}_{result_name}_attributes.xlsx")
            csv_path = os.path.join(out_dir, f"{input_base}_{result_name}_attributes.csv")
            interp_path = os.path.join(out_dir, f"{input_base}_{result_name}_{interp_tag}.tif")
            interp_png_path = os.path.join(out_dir, f"{input_base}_{result_name}_{interp_tag}.png")
            weights_path = os.path.join(out_dir, f"{input_base}_{result_name}_weights.csv")

            clear_output_targets(
                file_paths=[
                    resample_raster_path, resample_raster2_path, resample_raster3_path,
                    resample_png_path, resample_png2_path, resample_png3_path,
                    result_raster_path, result_png_path,
                    excel_path, csv_path,
                    interp_path, interp_png_path,
                    weights_path, weights_path.replace(".csv", ".xlsx")
                ],
                shp_paths=[points_path]
            )

            self.update_progress(10, "读取输入数据")

            if self.is_soil_mode():
                soil_ref = self.soil_ref_edit.text().strip()
                with rasterio.open(soil_ref) as ref_src:
                    work_transform = ref_src.transform
                    work_meta = ref_src.meta.copy()
                    crs = ref_src.crs

                soil_r = load_soil_factor_array(
                    self.soil_r_widget, soil_ref, invalid_values=invalid_values,
                    resampling_method=choose_resampling_by_name("R")
                )
                soil_k = load_soil_factor_array(
                    self.soil_k_widget, soil_ref, invalid_values=invalid_values,
                    resampling_method=choose_resampling_by_name("K")
                )
                soil_ls = load_soil_factor_array(
                    self.soil_ls_widget, soil_ref, invalid_values=invalid_values,
                    resampling_method=choose_resampling_by_name("LS")
                )
                soil_c = load_soil_factor_array(
                    self.soil_c_widget, soil_ref, invalid_values=invalid_values,
                    resampling_method=choose_resampling_by_name("C")
                )
                soil_p = load_soil_factor_array(
                    self.soil_p_widget, soil_ref, invalid_values=invalid_values,
                    resampling_method=choose_resampling_by_name("P")
                )

                check_raster_validity(soil_r, "R")
                check_raster_validity(soil_k, "K")
                check_raster_validity(soil_ls, "LS")
                check_raster_validity(soil_c, "C")
                check_raster_validity(soil_p, "P")

            elif self.is_soil_double_raster_mode():
                rst1 = read_and_process_raster(in_path, do_resample, target_res, invalid_values, out_nodata)
                work_transform = rst1["transform"]
                work_meta = rst1["meta"].copy()
                crs = rst1["crs"]

                self.update_progress(18, "读取并对齐实际侵蚀量栅格")
                rst2_data = align_raster_to_reference(
                    in_path2, work_meta, work_transform, crs, invalid_values, out_nodata
                )
                rst2 = {
                    "clean_data": rst2_data,
                    "meta": work_meta.copy(),
                    "transform": work_transform,
                    "crs": crs
                }

                potential_erosion = rst1["clean_data"].copy()
                actual_erosion = rst2["clean_data"].copy()
                check_raster_validity(potential_erosion, "潜在侵蚀量")
                check_raster_validity(actual_erosion, "实际侵蚀量")

                if do_resample:
                    self.update_progress(30, "导出重采样/对齐结果")
                    meta1 = work_meta.copy()
                    meta1.update({
                        "driver": "GTiff",
                        "count": 1,
                        "dtype": "float32",
                        "nodata": out_nodata,
                        "crs": crs,
                        "transform": work_transform
                    })

                    save_arr1 = np.where(np.isnan(rst1["clean_data"]), out_nodata, rst1["clean_data"]).astype(np.float32)
                    with rasterio.open(resample_raster_path, "w", **meta1) as dst:
                        dst.write(save_arr1, 1)
                    if self.chk_export_png.isChecked():
                        png_ok = save_array_png(rst1["clean_data"], resample_png_path, title=None)
                        if png_ok:
                            self.add_result_record(f"{input_base} -> 潜在侵蚀量重采样结果PNG", resample_png_path, kind="png")

                    save_arr2 = np.where(np.isnan(rst2["clean_data"]), out_nodata, rst2["clean_data"]).astype(np.float32)
                    with rasterio.open(resample_raster2_path, "w", **meta1) as dst:
                        dst.write(save_arr2, 1)
                    if self.chk_export_png.isChecked():
                        png_ok = save_array_png(rst2["clean_data"], resample_png2_path, title=None)
                        if png_ok:
                            self.add_result_record(f"{input_base} -> 实际侵蚀量重采样结果PNG", resample_png2_path, kind="png")

            elif not need_comprehensive:
                rst1 = read_and_process_raster(in_path, do_resample, target_res, invalid_values, out_nodata)
                work_transform = rst1["transform"]
                work_meta = rst1["meta"].copy()
                crs = rst1["crs"]

                if need_second_raster:
                    self.update_progress(18, "读取并对齐第二输入栅格")
                    rst2_data = align_raster_to_reference(
                        in_path2, work_meta, work_transform, crs, invalid_values, out_nodata
                    )
                    rst2 = {
                        "clean_data": rst2_data,
                        "meta": work_meta.copy(),
                        "transform": work_transform,
                        "crs": crs
                    }

                if need_third_raster:
                    self.update_progress(24, "读取并对齐第三输入栅格")
                    rst3_data = align_raster_to_reference(
                        in_path3, work_meta, work_transform, crs, invalid_values, out_nodata
                    )
                    rst3 = {
                        "clean_data": rst3_data,
                        "meta": work_meta.copy(),
                        "transform": work_transform,
                        "crs": crs
                    }

                if do_resample:
                    self.update_progress(30, "导出重采样/对齐结果")
                    meta1 = work_meta.copy()
                    meta1.update({
                        "driver": "GTiff",
                        "count": 1,
                        "dtype": "float32",
                        "nodata": out_nodata,
                        "crs": crs,
                        "transform": work_transform
                    })

                    save_arr1 = np.where(np.isnan(rst1["clean_data"]), out_nodata, rst1["clean_data"]).astype(np.float32)
                    with rasterio.open(resample_raster_path, "w", **meta1) as dst:
                        dst.write(save_arr1, 1)
                    if self.chk_export_png.isChecked():
                        png_ok = save_array_png(rst1["clean_data"], resample_png_path, title=None)
                        if png_ok:
                            self.add_result_record(f"{input_base} -> 输入1重采样结果PNG", resample_png_path, kind="png")

                    if rst2 is not None:
                        save_arr2 = np.where(np.isnan(rst2["clean_data"]), out_nodata, rst2["clean_data"]).astype(np.float32)
                        with rasterio.open(resample_raster2_path, "w", **meta1) as dst:
                            dst.write(save_arr2, 1)
                        if self.chk_export_png.isChecked():
                            png_ok = save_array_png(rst2["clean_data"], resample_png2_path, title=None)
                            if png_ok:
                                self.add_result_record(f"{input_base} -> 输入2重采样结果PNG", resample_png2_path, kind="png")

                    if rst3 is not None:
                        save_arr3 = np.where(np.isnan(rst3["clean_data"]), out_nodata, rst3["clean_data"]).astype(np.float32)
                        with rasterio.open(resample_raster3_path, "w", **meta1) as dst:
                            dst.write(save_arr3, 1)
                        if self.chk_export_png.isChecked():
                            png_ok = save_array_png(rst3["clean_data"], resample_png3_path, title=None)
                            if png_ok:
                                self.add_result_record(f"{input_base} -> 输入3重采样结果PNG", resample_png3_path, kind="png")

            else:
                self.update_progress(22, "读取并对齐多因子栅格")
                comp_info = self.load_comprehensive_factor_stack(
                    factors=factors,
                    do_resample=do_resample,
                    target_res=target_res,
                    invalid_values=invalid_values,
                    out_nodata=out_nodata
                )
                work_transform = comp_info["transform"]
                work_meta = comp_info["meta"]
                crs = comp_info["crs"]
                factor_names = comp_info["factor_names"]
                directions = comp_info["directions"]
                aligned_factor_dict = comp_info["aligned_dict"]
                factor_stack = comp_info["stack"]

            self.update_progress(48, "执行栅格计算")

            if self.is_soil_erosion():
                valid_mask = np.isfinite(soil_r) & np.isfinite(soil_k) & np.isfinite(soil_ls) & np.isfinite(soil_c) & np.isfinite(soil_p)
                result_data = np.full_like(soil_r, np.nan, dtype=np.float32)
                result_data[valid_mask] = soil_r[valid_mask] * soil_k[valid_mask] * soil_ls[valid_mask] * soil_c[valid_mask] * soil_p[valid_mask]

            elif self.is_soil_retention():
                valid_mask = np.isfinite(potential_erosion) & np.isfinite(actual_erosion)
                result_data = np.full_like(potential_erosion, np.nan, dtype=np.float32)
                result_data[valid_mask] = potential_erosion[valid_mask] - actual_erosion[valid_mask]

            elif self.is_water_cycle():
                p = rst1["clean_data"]
                e = rst2["clean_data"]
                r = rst3["clean_data"]
                result_data = np.full_like(p, np.nan, dtype=np.float32)
                valid_mask = np.isfinite(p) & np.isfinite(e) & np.isfinite(r)
                result_data[valid_mask] = p[valid_mask] - e[valid_mask] - r[valid_mask]

            elif self.is_carbon_storage():
                agb = rst1["clean_data"]
                result_data = np.full_like(agb, np.nan, dtype=np.float32)
                valid_mask = np.isfinite(agb)
                result_data[valid_mask] = agb[valid_mask] * cf

            elif self.is_agb_dual():
                clean1 = rst1["clean_data"]
                clean2 = rst2["clean_data"]
                growth = clean2 - clean1
                anpp = growth / year_diff

                result_data = np.full_like(clean1, np.nan, dtype=np.float32)
                valid_mask = np.isfinite(clean1) & np.isfinite(clean2) & np.isfinite(anpp)
                result_data[valid_mask] = anpp[valid_mask]

            elif self.is_carbon_dual():
                clean1 = rst1["clean_data"]
                clean2 = rst2["clean_data"]
                carbon_increment = clean2 - clean1
                anpp = carbon_increment / year_diff

                result_data = np.full_like(clean1, np.nan, dtype=np.float32)
                valid_mask = np.isfinite(clean1) & np.isfinite(clean2) & np.isfinite(anpp)
                result_data[valid_mask] = anpp[valid_mask]

            elif self.is_volume_dual():
                clean1 = rst1["clean_data"]
                clean2 = rst2["clean_data"]
                volume_increment = clean2 - clean1
                anpp = (volume_increment / year_diff) * rho * k

                result_data = np.full_like(clean1, np.nan, dtype=np.float32)
                valid_mask = np.isfinite(clean1) & np.isfinite(clean2) & np.isfinite(anpp)
                result_data[valid_mask] = anpp[valid_mask]

            elif self.is_comprehensive():
                self.update_progress(52, "标准化多因子栅格")
                norm_list = []
                for i in range(factor_stack.shape[0]):
                    norm_list.append(normalize_single_array(factor_stack[i], directions[i]))
                norm_stack = np.stack(norm_list, axis=0).astype(np.float32)

                if self.comp_method == "等权法":
                    weights = np.ones(len(factor_names), dtype=np.float32) / len(factor_names)
                else:
                    weights = entropy_weights(norm_stack)

                weighted_stack = norm_stack * weights[:, None, None]
                result_data = np.full(norm_stack.shape[1:], np.nan, dtype=np.float32)
                valid_mask = np.all(np.isfinite(norm_stack), axis=0)
                result_data[valid_mask] = np.nansum(weighted_stack[:, valid_mask], axis=0)

                _, weights_df = save_weights_csv(weights_path, factor_names, directions, weights)
                self.add_result_record(f"{input_base} -> {result_name} 权重表", kind="table", data=weights_df)

            else:
                clean_data = rst1["clean_data"]
                calc_result = evaluate_formula(clean_data, formula)
                result_data = np.full_like(clean_data, np.nan, dtype=np.float32)
                valid_mask = np.isfinite(clean_data) & np.isfinite(calc_result)
                result_data[valid_mask] = calc_result[valid_mask]

            self.update_progress(62, "导出计算结果")
            if self.chk_export_calc.isChecked():
                save_arr = np.where(np.isnan(result_data), out_nodata, result_data).astype(np.float32)
                out_meta = work_meta.copy()
                out_meta.update({
                    "driver": "GTiff",
                    "count": 1,
                    "dtype": "float32",
                    "nodata": out_nodata,
                    "crs": crs,
                    "transform": work_transform
                })
                with rasterio.open(result_raster_path, "w", **out_meta) as dst:
                    dst.write(save_arr, 1)
                    try:
                        dst.set_band_description(1, result_name)
                    except Exception:
                        pass

                if self.is_soil_retention():
                    potential_path = os.path.join(out_dir, f"{input_base}_SOIL_POTENTIAL_result.tif")
                    actual_path = os.path.join(out_dir, f"{input_base}_SOIL_ACTUAL_result.tif")
                    with rasterio.open(potential_path, "w", **out_meta) as dst:
                        dst.write(np.where(np.isnan(potential_erosion), out_nodata, potential_erosion).astype(np.float32), 1)
                    with rasterio.open(actual_path, "w", **out_meta) as dst:
                        dst.write(np.where(np.isnan(actual_erosion), out_nodata, actual_erosion).astype(np.float32), 1)

            self.update_progress(72, "生成计算结果预览")
            if self.chk_export_png.isChecked():
                png_ok = save_array_png(result_data, result_png_path, title=None)
                if png_ok:
                    self.add_result_record(f"{input_base} -> {result_name} 计算结果PNG", result_png_path, kind="png")

            self.update_progress(82, "生成提值点结果与属性表")
            need_points = (
                self.chk_export_points.isChecked()
                or self.chk_export_excel.isChecked()
                or (use_interp and self.chk_export_interp.isChecked())
            )

            df = None
            xs = ys = result_vals = None

            if need_points:
                rows, cols = np.where(~np.isnan(result_data))
                if len(rows) == 0:
                    QMessageBox.warning(self, "提示", "没有有效像元可导出")
                    return

                a = work_transform.a
                b = work_transform.b
                c = work_transform.c
                d = work_transform.d
                e = work_transform.e
                f = work_transform.f

                xs = c + cols * a + rows * b + a / 2.0
                ys = f + cols * d + rows * e + e / 2.0

                xs = xs.astype(np.float64)
                ys = ys.astype(np.float64)
                result_vals = result_data[rows, cols].astype(np.float64)

                if self.is_soil_erosion():
                    df = pd.DataFrame({
                        "x": xs,
                        "y": ys,
                        "R": soil_r[rows, cols].astype(np.float64),
                        "K": soil_k[rows, cols].astype(np.float64),
                        "LS": soil_ls[rows, cols].astype(np.float64),
                        "C": soil_c[rows, cols].astype(np.float64),
                        "P": soil_p[rows, cols].astype(np.float64),
                        "SOIL_EROSION": result_vals
                    })

                elif self.is_soil_retention():
                    df = pd.DataFrame({
                        "x": xs,
                        "y": ys,
                        "potential_erosion": potential_erosion[rows, cols].astype(np.float64),
                        "actual_erosion": actual_erosion[rows, cols].astype(np.float64),
                        "SOIL_RETENTION": result_vals
                    })

                elif self.is_water_cycle():
                    p_vals = rst1["clean_data"][rows, cols].astype(np.float64)
                    e_vals = rst2["clean_data"][rows, cols].astype(np.float64)
                    r_vals = rst3["clean_data"][rows, cols].astype(np.float64)
                    df = pd.DataFrame({
                        "x": xs,
                        "y": ys,
                        "P": p_vals,
                        "E": e_vals,
                        "R": r_vals,
                        "WATER": result_vals
                    })

                elif self.is_carbon_storage():
                    agb_vals = rst1["clean_data"][rows, cols].astype(np.float64)
                    df = pd.DataFrame({
                        "x": xs,
                        "y": ys,
                        "AGB": agb_vals,
                        "CF": cf,
                        "CARBON": result_vals
                    })

                elif self.is_agb_dual():
                    agb_t1 = rst1["clean_data"][rows, cols].astype(np.float64)
                    agb_t2 = rst2["clean_data"][rows, cols].astype(np.float64)
                    growth_vals = (agb_t2 - agb_t1).astype(np.float64)

                    df = pd.DataFrame({
                        "x": xs,
                        "y": ys,
                        "agb_t1": agb_t1,
                        "agb_t2": agb_t2,
                        "growth": growth_vals,
                        "ANPP": result_vals
                    })

                elif self.is_carbon_dual():
                    c_t1 = rst1["clean_data"][rows, cols].astype(np.float64)
                    c_t2 = rst2["clean_data"][rows, cols].astype(np.float64)
                    carbon_increment = (c_t2 - c_t1).astype(np.float64)

                    df = pd.DataFrame({
                        "x": xs,
                        "y": ys,
                        "carbon_t1": c_t1,
                        "carbon_t2": c_t2,
                        "carbon_increment": carbon_increment,
                        "ANPP": result_vals
                    })

                elif self.is_volume_dual():
                    v_t1 = rst1["clean_data"][rows, cols].astype(np.float64)
                    v_t2 = rst2["clean_data"][rows, cols].astype(np.float64)
                    volume_increment = (v_t2 - v_t1).astype(np.float64)

                    df = pd.DataFrame({
                        "x": xs,
                        "y": ys,
                        "v_t1": v_t1,
                        "v_t2": v_t2,
                        "volume_increment": volume_increment,
                        "rho": rho,
                        "k": k,
                        "ANPP": result_vals
                    })

                elif self.is_comprehensive():
                    data_dict = {
                        "x": xs,
                        "y": ys
                    }
                    for name in factor_names:
                        data_dict[name] = aligned_factor_dict[name][rows, cols].astype(np.float64)

                    if weights is not None:
                        weight_map = {name: float(w) for name, w in zip(factor_names, weights)}
                        for name in factor_names:
                            data_dict[f"{name}_w"] = np.full(len(xs), weight_map[name], dtype=np.float64)

                    data_dict[result_name] = result_vals
                    df = pd.DataFrame(data_dict)

                else:
                    raw_vals = rst1["clean_data"][rows, cols].astype(np.float64)
                    df = pd.DataFrame({
                        "x": xs,
                        "y": ys,
                        "raw_val": raw_vals,
                        result_name: result_vals
                    })

                self.add_result_record(f"{input_base} -> {result_name} 提值点属性表", kind="table", data=df)

                if self.chk_export_points.isChecked():
                    safe_remove_shapefile(points_path)
                    gdf = gpd.GeoDataFrame(
                        df.copy(),
                        geometry=[Point(x, y) for x, y in zip(xs, ys)],
                        crs=crs
                    )
                    gdf.to_file(points_path, driver="ESRI Shapefile", encoding="utf-8")

                if self.chk_export_excel.isChecked():
                    safe_remove_file(excel_path)
                    safe_remove_file(csv_path)
                    try:
                        df.to_excel(excel_path, index=False)
                    except Exception:
                        df.to_csv(csv_path, index=False, encoding="utf-8-sig")

            if use_interp and self.chk_export_interp.isChecked():
                self.update_progress(92, f"执行{interp_method}")

                if df is not None and len(df) >= 3:
                    if len(result_vals) > max_interp_points:
                        idx = np.random.choice(len(result_vals), max_interp_points, replace=False)
                        interp_x = xs[idx]
                        interp_y = ys[idx]
                        interp_z = result_vals[idx]
                    else:
                        interp_x = xs
                        interp_y = ys
                        interp_z = result_vals

                    interp_grid = build_interp_grid(work_transform, work_meta, interp_res=interp_res)

                    gridx = interp_grid["gridx"]
                    gridy = interp_grid["gridy"]
                    grid_width = interp_grid["width"]
                    grid_height = interp_grid["height"]
                    interp_transform = interp_grid["transform"]

                    z = None
                    if interp_method == "普通克里金":
                        if OrdinaryKriging is None:
                            raise ImportError("未安装 pykrige，无法使用普通克里金插值。请先安装 pykrige。")
                        ok = OrdinaryKriging(
                            interp_x, interp_y, interp_z,
                            variogram_model="linear",
                            verbose=False,
                            enable_plotting=False
                        )
                        z, _ = ok.execute("grid", gridx, gridy)
                        z = np.array(z, dtype=np.float32)

                    elif interp_method == "反距离权重（IDW）":
                        z = idw_interpolation(interp_x, interp_y, interp_z, gridx, gridy).astype(np.float32)

                    elif interp_method == "径向基函数（RBF）":
                        z = rbf_interpolation(interp_x, interp_y, interp_z, gridx, gridy).astype(np.float32)

                    if z is not None:
                        valid_area_mask = build_valid_mask_for_interp(result_data, work_transform, interp_grid)
                        z = np.array(z, dtype=np.float32)
                        z[~valid_area_mask] = np.nan

                        with rasterio.open(
                            interp_path,
                            "w",
                            driver="GTiff",
                            height=grid_height,
                            width=grid_width,
                            count=1,
                            dtype="float32",
                            crs=crs,
                            transform=interp_transform,
                            nodata=out_nodata
                        ) as dst:
                            dst.write(np.where(np.isnan(z), out_nodata, z).astype(np.float32), 1)

                        if self.chk_export_png.isChecked():
                            png_ok = save_array_png(z, interp_png_path, title=None)
                            if png_ok:
                                self.add_result_record(
                                    f"{input_base} -> {result_name} 插值结果PNG",
                                    interp_png_path,
                                    kind="png"
                                )

            self.update_progress(100, "运行完成")

            if self.result_list.count() > 0:
                self.result_list.setCurrentRow(0)
                self.on_result_item_clicked(self.result_list.item(0))

            QMessageBox.information(self, "完成", "流程运行完成，本次输出结果已显示在右侧。")

        except Exception as e:
            self.progress_bar.setValue(0)
            self.status_label.setText("状态：运行失败")
            QMessageBox.critical(
                self,
                "错误",
                f"运行失败：\n{str(e)}\n\n{traceback.format_exc()}"
            )
        finally:
            self.run_btn.setEnabled(True)

