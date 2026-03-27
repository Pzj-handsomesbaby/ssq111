# -*- coding: utf-8 -*-
"""
第5部分：绘图输出模块（字体优化版）
依赖：part1_data_loader.py
适用环境：Spyder / Python 3.x

功能：
1. 为数值型变量生成直方图
2. 为数值型变量生成箱线图
3. 为分类型变量生成柱状图
4. 自动保存为 PNG 文件
5. 中文优先宋体，英文和拉丁字符优先 Times New Roman
"""

import os
import re
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
from matplotlib import font_manager as fm

from descriptive_stats_data_loader import (
    create_output_dirs,
    read_data_file,
    clean_data,
    identify_variable_types,
    summarize_data_info
)


# =========================
# 1. 字体与全局样式设置
# =========================

def pick_first_available_font(candidates):
    """
    从候选字体中选择当前系统已安装的第一个字体
    """
    available_fonts = {f.name for f in fm.fontManager.ttflist}
    for font_name in candidates:
        if font_name in available_fonts:
            return font_name
    return None


def configure_matplotlib_fonts():
    """
    配置 matplotlib 字体：
    - 英文/数字优先 Times New Roman
    - 中文优先宋体（SimSun）
    """
    english_font = pick_first_available_font([
        "Times New Roman",
        "Times New Roman PS MT",
        "DejaVu Serif"
    ])

    chinese_font = pick_first_available_font([
        "SimSun",        # 宋体（Windows 常见）
        "NSimSun",       # 新宋体
        "Songti SC",     # macOS 常见宋体
        "STSong",
        "SimHei",        # 备选
        "Microsoft YaHei",
        "DejaVu Sans"
    ])

    if english_font is None:
        english_font = "DejaVu Serif"

    if chinese_font is None:
        chinese_font = "DejaVu Sans"

    # 关键：将字体族设置为“英文优先 + 中文回退”
    mpl.rcParams["font.family"] = [english_font, chinese_font]

    # 负号正常显示
    mpl.rcParams["axes.unicode_minus"] = False

    # 白底输出
    mpl.rcParams["figure.facecolor"] = "white"
    mpl.rcParams["axes.facecolor"] = "white"
    mpl.rcParams["savefig.facecolor"] = "white"

    # 字号
    mpl.rcParams["axes.titlesize"] = 16
    mpl.rcParams["axes.labelsize"] = 13
    mpl.rcParams["xtick.labelsize"] = 11
    mpl.rcParams["ytick.labelsize"] = 11

    # print(f"当前英文优先字体：{english_font}")
    # print(f"当前中文回退字体：{chinese_font}")

    return english_font, chinese_font


ENGLISH_FONT, CHINESE_FONT = configure_matplotlib_fonts()


def apply_axis_fonts(ax):
    """
    对标题、坐标轴标题、刻度统一应用字体回退列表
    """
    font_family_list = [ENGLISH_FONT, CHINESE_FONT]

    ax.title.set_fontfamily(font_family_list)
    ax.xaxis.label.set_fontfamily(font_family_list)
    ax.yaxis.label.set_fontfamily(font_family_list)

    for label in ax.get_xticklabels():
        label.set_fontfamily(font_family_list)

    for label in ax.get_yticklabels():
        label.set_fontfamily(font_family_list)


# =========================
# 2. 通用辅助函数
# =========================

def sanitize_filename(filename):
    """
    清理文件名中的非法字符，避免 Windows 保存报错
    """
    filename = str(filename).strip()
    filename = re.sub(r'[\\/:*?"<>|]', "_", filename)
    return filename


# =========================
# 3. 绘图函数
# =========================

def generate_histograms(df, numeric_cols, output_dir, bins=10, dpi=300):
    """
    为每个数值型变量生成直方图并保存
    """
    hist_dir = os.path.join(output_dir, "直方图")

    if not numeric_cols:
        print("未识别到数值型变量，无法生成直方图。")
        return

    for col in numeric_cols:
        series = pd.to_numeric(df[col], errors="coerce").dropna()

        if series.empty:
            print(f"变量【{col}】无有效数据，跳过直方图生成。")
            continue

        safe_col = sanitize_filename(col)
        save_path = os.path.join(hist_dir, f"{safe_col}_直方图.png")

        fig, ax = plt.subplots(figsize=(8, 6))

        ax.hist(series, bins=bins, edgecolor="black")
        ax.set_title(f"{col} 直方图")
        ax.set_xlabel(col)
        ax.set_ylabel("频数")

        apply_axis_fonts(ax)

        fig.tight_layout()
        fig.savefig(save_path, dpi=dpi, bbox_inches="tight")
        plt.close(fig)

        print(f"直方图已保存：{save_path}")


def generate_boxplots(df, numeric_cols, output_dir, dpi=300):
    """
    为每个数值型变量生成箱线图并保存
    """
    boxplot_dir = os.path.join(output_dir, "箱线图")

    if not numeric_cols:
        print("未识别到数值型变量，无法生成箱线图。")
        return

    for col in numeric_cols:
        series = pd.to_numeric(df[col], errors="coerce").dropna()

        if series.empty:
            print(f"变量【{col}】无有效数据，跳过箱线图生成。")
            continue

        safe_col = sanitize_filename(col)
        save_path = os.path.join(boxplot_dir, f"{safe_col}_箱线图.png")

        fig, ax = plt.subplots(figsize=(6, 6))

        ax.boxplot(series, vert=True)
        ax.set_title(f"{col} 箱线图")
        ax.set_ylabel(col)

        apply_axis_fonts(ax)

        fig.tight_layout()
        fig.savefig(save_path, dpi=dpi, bbox_inches="tight")
        plt.close(fig)

        print(f"箱线图已保存：{save_path}")


def generate_bar_charts(df, categorical_cols, output_dir, dpi=300):
    """
    为每个分类型变量生成柱状图并保存
    """
    bar_dir = os.path.join(output_dir, "柱状图")

    if not categorical_cols:
        print("未识别到分类型变量，无法生成柱状图。")
        return

    for col in categorical_cols:
        series = df[col].dropna()

        if series.empty:
            print(f"变量【{col}】无有效数据，跳过柱状图生成。")
            continue

        freq_series = series.astype(str).value_counts(sort=False)

        safe_col = sanitize_filename(col)
        save_path = os.path.join(bar_dir, f"{safe_col}_柱状图.png")

        fig, ax = plt.subplots(figsize=(8, 6))

        bars = ax.bar(freq_series.index, freq_series.values, edgecolor="black")
        ax.set_title(f"{col} 柱状图")
        ax.set_xlabel(col)
        ax.set_ylabel("频数")

        # 柱顶数字
        for bar in bars:
            height = bar.get_height()
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                height,
                f"{int(height)}",
                ha="center",
                va="bottom",
                fontfamily=[ENGLISH_FONT, CHINESE_FONT],
                fontsize=11
            )

        apply_axis_fonts(ax)

        fig.tight_layout()
        fig.savefig(save_path, dpi=dpi, bbox_inches="tight")
        plt.close(fig)

        print(f"柱状图已保存：{save_path}")


# =========================
# 4. 主函数
# =========================

def main():
    """
    独立测试本模块
    """
    # ====== 这里改成你自己的路径 ======
    file_path = r"C:\Users\彭宗健\Desktop\软件设置\试验数据.xlsx"
    output_dir = r"C:\Users\彭宗健\Desktop\软件设置\输出结果\第五步输出结果"
    sheet_name = None

    try:
        # 1. 创建输出目录
        create_output_dirs(output_dir)

        # 2. 读取数据
        df, file_name = read_data_file(file_path, sheet_name=sheet_name)

        # 3. 数据预处理
        df = clean_data(df)

        # 4. 识别变量类型
        numeric_cols, categorical_cols = identify_variable_types(df)

        # 5. 输出数据基本信息
        summarize_data_info(df, numeric_cols, categorical_cols)

        # 6. 生成图形
        generate_histograms(df, numeric_cols, output_dir, bins=10, dpi=300)
        generate_boxplots(df, numeric_cols, output_dir, dpi=300)
        generate_bar_charts(df, categorical_cols, output_dir, dpi=300)

        print("\n第5部分（字体优化版）运行完成。")

    except Exception as e:
        print(f"程序运行出错：{e}")


if __name__ == "__main__":
    main()