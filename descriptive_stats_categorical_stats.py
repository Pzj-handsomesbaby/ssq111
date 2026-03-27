# -*- coding: utf-8 -*-
"""
Created on Sun Mar 15 11:13:05 2026

@author: Pengzongjian
"""

# -*- coding: utf-8 -*-
"""
第3部分：分类型变量频数统计
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


def calculate_categorical_stats_for_column(series, col_name, round_digits=4):
    """
    对单个分类型变量计算频数统计

    参数：
        series: pandas Series
        col_name: 变量名
        round_digits: 小数保留位数

    返回：
        result_df: 单个分类型变量的频数统计结果表
    """
    total_n = len(series)
    missing_n = series.isna().sum()
    valid_series = series.dropna()
    valid_n = len(valid_series)
    missing_rate = missing_n / total_n if total_n > 0 else np.nan

    # 如果该列全为空，则返回空表
    if valid_n == 0:
        return pd.DataFrame({
            "变量名": [col_name],
            "类别": [np.nan],
            "频数": [0],
            "频率": [np.nan],
            "百分比(%)": [np.nan],
            "累计频数": [np.nan],
            "累计百分比(%)": [np.nan],
            "总样本数": [total_n],
            "有效样本数": [0],
            "缺失值个数": [missing_n],
            "缺失值比例": [round(missing_rate, round_digits) if pd.notna(missing_rate) else np.nan]
        })

    # 保留原始出现顺序，不强制排序
    freq_series = valid_series.astype(str).value_counts(dropna=False, sort=False)

    result_df = freq_series.reset_index()
    result_df.columns = ["类别", "频数"]

    result_df["频率"] = result_df["频数"] / valid_n
    result_df["百分比(%)"] = result_df["频率"] * 100
    result_df["累计频数"] = result_df["频数"].cumsum()
    result_df["累计百分比(%)"] = result_df["百分比(%)"].cumsum()

    # 加上变量层面的信息
    result_df.insert(0, "变量名", col_name)
    result_df["总样本数"] = total_n
    result_df["有效样本数"] = valid_n
    result_df["缺失值个数"] = missing_n
    result_df["缺失值比例"] = round(missing_rate, round_digits)

    # 数值列保留小数位
    for col in ["频率", "百分比(%)", "累计百分比(%)"]:
        result_df[col] = result_df[col].round(round_digits)

    return result_df


def generate_categorical_stats(df, categorical_cols, round_digits=4):
    """
    对所有分类型变量批量生成频数统计表

    参数：
        df: 原始数据表
        categorical_cols: 分类型变量列表
        round_digits: 小数保留位数

    返回：
        final_df: 所有分类型变量合并后的频数统计结果表
    """
    if not categorical_cols:
        print("未识别到分类型变量，无法生成分类型频数统计表。")
        return pd.DataFrame()

    all_results = []

    for col in categorical_cols:
        col_result_df = calculate_categorical_stats_for_column(df[col], col, round_digits=round_digits)
        all_results.append(col_result_df)

    final_df = pd.concat(all_results, ignore_index=True)

    return final_df


def main():
    """
    独立测试本模块
    """
    # ====== 这里改成你自己的路径 ======
    file_path = r"C:\Users\彭宗健\Desktop\软件设置\试验数据.xlsx"
    output_dir = r"C:\Users\彭宗健\Desktop\软件设置\输出结果\第三步输出结果"
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

        # 6. 生成分类型频数统计表
        categorical_stats_df = generate_categorical_stats(df, categorical_cols, round_digits=4)

        print("分类型变量频数统计结果如下：")
        print(categorical_stats_df)

        print("\n第3部分运行完成。")

    except Exception as e:
        print(f"程序运行出错：{e}")


if __name__ == "__main__":
    main()