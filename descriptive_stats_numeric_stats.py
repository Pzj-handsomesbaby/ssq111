# -*- coding: utf-8 -*-
"""
Created on Sun Mar 15 10:51:11 2026

@author: Pengzongjian
"""

# -*- coding: utf-8 -*-
"""
第2部分：数值型变量描述统计
依赖：part1_data_loader.py
适用环境：Spyder / Python 3.x
"""

import pandas as pd
import numpy as np

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)
from descriptive_stats_data_loader import (
    create_output_dirs,
    read_data_file,
    clean_data,
    identify_variable_types,
    summarize_data_info
)


def calculate_numeric_stats_for_column(series, col_name, round_digits=4):
    """
    对单个数值型变量计算描述统计指标

    参数：
        series: pandas Series
        col_name: 变量名
        round_digits: 小数保留位数

    返回：
        stats_dict: 单个变量的统计结果字典
    """
    total_n = len(series)
    valid_series = pd.to_numeric(series, errors="coerce").dropna()
    valid_n = len(valid_series)
    missing_n = total_n - valid_n
    missing_rate = missing_n / total_n if total_n > 0 else np.nan

    # 若该列全为空，直接返回空统计
    if valid_n == 0:
        return {
            "变量名": col_name,
            "总样本数": total_n,
            "有效样本数": 0,
            "缺失值个数": missing_n,
            "缺失值比例": round(missing_rate, round_digits) if pd.notna(missing_rate) else np.nan,
            "均值": np.nan,
            "中位数": np.nan,
            "众数": np.nan,
            "最小值": np.nan,
            "最大值": np.nan,
            "极差": np.nan,
            "方差": np.nan,
            "标准差": np.nan,
            "变异系数(%)": np.nan,
            "Q1": np.nan,
            "Q2": np.nan,
            "Q3": np.nan,
            "IQR": np.nan,
            "偏度": np.nan,
            "峰度": np.nan
        }

    # 众数：若有多个众数，只取第一个
    mode_values = valid_series.mode()
    mode_value = mode_values.iloc[0] if not mode_values.empty else np.nan

    mean_value = valid_series.mean()
    median_value = valid_series.median()
    min_value = valid_series.min()
    max_value = valid_series.max()
    range_value = max_value - min_value

    # 样本方差/标准差（ddof=1）
    variance_value = valid_series.var(ddof=1) if valid_n > 1 else np.nan
    std_value = valid_series.std(ddof=1) if valid_n > 1 else np.nan

    # 变异系数：均值不为0时计算
    if pd.notna(mean_value) and mean_value != 0 and pd.notna(std_value):
        cv_value = (std_value / mean_value) * 100
    else:
        cv_value = np.nan

    q1_value = valid_series.quantile(0.25)
    q2_value = valid_series.quantile(0.50)
    q3_value = valid_series.quantile(0.75)
    iqr_value = q3_value - q1_value

    skew_value = valid_series.skew() if valid_n >= 3 else np.nan
    kurt_value = valid_series.kurt() if valid_n >= 4 else np.nan

    stats_dict = {
        "变量名": col_name,
        "总样本数": total_n,
        "有效样本数": valid_n,
        "缺失值个数": missing_n,
        "缺失值比例": round(missing_rate, round_digits),
        "均值": round(mean_value, round_digits),
        "中位数": round(median_value, round_digits),
        "众数": round(mode_value, round_digits) if pd.notna(mode_value) else np.nan,
        "最小值": round(min_value, round_digits),
        "最大值": round(max_value, round_digits),
        "极差": round(range_value, round_digits),
        "方差": round(variance_value, round_digits) if pd.notna(variance_value) else np.nan,
        "标准差": round(std_value, round_digits) if pd.notna(std_value) else np.nan,
        "变异系数(%)": round(cv_value, round_digits) if pd.notna(cv_value) else np.nan,
        "Q1": round(q1_value, round_digits),
        "Q2": round(q2_value, round_digits),
        "Q3": round(q3_value, round_digits),
        "IQR": round(iqr_value, round_digits),
        "偏度": round(skew_value, round_digits) if pd.notna(skew_value) else np.nan,
        "峰度": round(kurt_value, round_digits) if pd.notna(kurt_value) else np.nan
    }

    return stats_dict


def generate_numeric_stats(df, numeric_cols, round_digits=4):
    """
    对所有数值型变量批量生成描述统计表

    参数：
        df: 原始数据表
        numeric_cols: 数值型变量列表
        round_digits: 小数保留位数

    返回：
        result_df: 数值型描述统计结果表
    """
    if not numeric_cols:
        print("未识别到数值型变量，无法生成数值型描述统计表。")
        return pd.DataFrame()

    results = []

    for col in numeric_cols:
        stats_dict = calculate_numeric_stats_for_column(df[col], col, round_digits=round_digits)
        results.append(stats_dict)

    result_df = pd.DataFrame(results)

    return result_df


def main():
    """
    独立测试本模块
    """
    # ====== 这里改成你自己的路径 ======
    file_path = r"C:\Users\彭宗健\Desktop\软件设置\试验数据.xlsx"
    output_dir = r"C:\Users\彭宗健\Desktop\软件设置\输出结果\第二步输出结果"
    sheet_name = None

    try:
        # 1. 输出目录准备
        create_output_dirs(output_dir)

        # 2. 读取数据
        df, file_name = read_data_file(file_path, sheet_name=sheet_name)

        # 3. 数据预处理
        df = clean_data(df)

        # 4. 识别变量类型
        numeric_cols, categorical_cols = identify_variable_types(df)

        # 5. 输出基本信息
        summarize_data_info(df, numeric_cols, categorical_cols)

        # 6. 生成数值型描述统计表
        numeric_stats_df = generate_numeric_stats(df, numeric_cols, round_digits=4)

        print("数值型描述统计结果如下：")
        print(numeric_stats_df)

        print("\n第2部分运行完成。")

    except Exception as e:
        print(f"程序运行出错：{e}")


if __name__ == "__main__":
    main()