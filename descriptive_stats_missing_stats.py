# -*- coding: utf-8 -*-
"""
Created on Sun Mar 15 11:23:26 2026

@author: Pengzongjian
"""

# -*- coding: utf-8 -*-
"""
第4部分：缺失值统计
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


def generate_missing_stats(df, numeric_cols=None, categorical_cols=None, round_digits=4):
    """
    生成每一列的缺失值统计表

    参数：
        df: 原始数据表
        numeric_cols: 数值型变量列表
        categorical_cols: 分类型变量列表
        round_digits: 小数保留位数

    返回：
        missing_stats_df: 每列缺失值统计表
    """
    results = []

    for col in df.columns:
        total_n = len(df)
        missing_n = df[col].isna().sum()
        valid_n = total_n - missing_n
        missing_rate = missing_n / total_n if total_n > 0 else np.nan
        non_missing_rate = valid_n / total_n if total_n > 0 else np.nan

        # 判断变量类型
        if numeric_cols is not None and col in numeric_cols:
            var_type = "数值型"
        elif categorical_cols is not None and col in categorical_cols:
            var_type = "分类型"
        else:
            var_type = "未识别"

        results.append({
            "变量名": col,
            "变量类型": var_type,
            "总样本数": total_n,
            "非缺失值个数": valid_n,
            "缺失值个数": missing_n,
            "非缺失值比例": round(non_missing_rate, round_digits) if pd.notna(non_missing_rate) else np.nan,
            "缺失值比例": round(missing_rate, round_digits) if pd.notna(missing_rate) else np.nan
        })

    missing_stats_df = pd.DataFrame(results)
    return missing_stats_df


def generate_overall_missing_summary(df, round_digits=4):
    """
    生成整个数据表的总体缺失情况汇总表

    参数：
        df: 原始数据表
        round_digits: 小数保留位数

    返回：
        summary_df: 总体缺失情况汇总表
    """
    total_rows = df.shape[0]
    total_cols = df.shape[1]
    total_cells = total_rows * total_cols

    missing_cells = df.isna().sum().sum()
    non_missing_cells = total_cells - missing_cells

    missing_rate = missing_cells / total_cells if total_cells > 0 else np.nan
    non_missing_rate = non_missing_cells / total_cells if total_cells > 0 else np.nan

    summary_df = pd.DataFrame([{
        "总行数": total_rows,
        "总列数": total_cols,
        "总单元格数": total_cells,
        "非缺失单元格数": non_missing_cells,
        "缺失单元格数": missing_cells,
        "非缺失单元格比例": round(non_missing_rate, round_digits) if pd.notna(non_missing_rate) else np.nan,
        "缺失单元格比例": round(missing_rate, round_digits) if pd.notna(missing_rate) else np.nan
    }])

    return summary_df


def main():
    """
    独立测试本模块
    """
    # ====== 这里改成你自己的路径 ======
    file_path = r"C:\Users\彭宗健\Desktop\软件设置\试验数据.xlsx"
    output_dir = r"C:\Users\彭宗健\Desktop\软件设置\输出结果\第四步输出结果"
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

        # 6. 每列缺失值统计
        missing_stats_df = generate_missing_stats(
            df,
            numeric_cols=numeric_cols,
            categorical_cols=categorical_cols,
            round_digits=4
        )

        print("每列缺失值统计结果如下：")
        print(missing_stats_df)

        # 7. 整体缺失情况汇总
        overall_missing_summary_df = generate_overall_missing_summary(df, round_digits=4)

        print("\n整体缺失情况汇总如下：")
        print(overall_missing_summary_df)

        print("\n第4部分运行完成。")

    except Exception as e:
        print(f"程序运行出错：{e}")


if __name__ == "__main__":
    main()