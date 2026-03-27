# -*- coding: utf-8 -*-
"""
Created on Sun Mar 15 10:31:37 2026

@author: Pengzongjian
"""

# -*- coding: utf-8 -*-
"""
第1部分：文件读取 + 输出目录创建 + 数据预处理 + 变量类型识别
适用环境：Spyder / Python 3.x
"""

import os
import pandas as pd


def create_output_dirs(output_dir):
    """
    创建输出目录及后续要用到的子目录
    """
    subdirs = [
        output_dir,
        os.path.join(output_dir, "统计结果Excel"),
        os.path.join(output_dir, "直方图"),
        os.path.join(output_dir, "箱线图"),
        os.path.join(output_dir, "柱状图")
    ]

    for folder in subdirs:
        if not os.path.exists(folder):
            os.makedirs(folder)

    print(f"输出目录已准备完成：{output_dir}")


def read_data_file(file_path, sheet_name=None):
    """
    读取 Excel 或 CSV 文件
    参数：
        file_path: 输入文件路径
        sheet_name: Excel工作表名或序号；如果为None，则默认读取第一个工作表
    返回：
        df: 读取后的DataFrame
        file_name: 不带后缀的文件名
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在：{file_path}")

    file_name = os.path.splitext(os.path.basename(file_path))[0]
    ext = os.path.splitext(file_path)[1].lower()

    if ext in [".xlsx", ".xls"]:
        if sheet_name is None:
            df = pd.read_excel(file_path, sheet_name=0)
            print("未指定工作表，默认读取第一个工作表。")
        else:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"已读取工作表：{sheet_name}")

    elif ext == ".csv":
        # 优先尝试 utf-8，失败再尝试 gbk
        try:
            df = pd.read_csv(file_path, encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(file_path, encoding="gbk")
        print("已读取 CSV 文件。")

    else:
        raise ValueError("仅支持 .xlsx、.xls、.csv 文件格式。")

    print(f"文件读取成功：{file_path}")
    return df, file_name


def clean_data(df):
    """
    数据预处理：
    1. 删除全空行
    2. 删除全空列
    3. 清理列名首尾空格
    """
    original_shape = df.shape

    # 删除全空行和全空列
    df = df.dropna(axis=0, how="all")
    df = df.dropna(axis=1, how="all")

    # 清理列名
    df.columns = [str(col).strip() for col in df.columns]

    new_shape = df.shape

    print("数据预处理完成。")
    print(f"原始数据维度：{original_shape}")
    print(f"清理后数据维度：{new_shape}")

    return df


def identify_variable_types(df, unique_ratio_threshold=0.05, unique_count_threshold=10):
    """
    自动识别变量类型
    识别规则（当前为较稳妥的基础版）：
    1. pandas识别为数值型的列，默认归为数值型
    2. 但如果某列唯一值个数很少，且更像类别编号，可归为分类型
    3. 其余非数值型列归为分类型

    参数：
        unique_ratio_threshold: 唯一值占比阈值
        unique_count_threshold: 唯一值个数阈值

    返回：
        numeric_cols: 数值型变量列表
        categorical_cols: 分类型变量列表
    """
    numeric_cols = []
    categorical_cols = []

    n_rows = len(df)

    for col in df.columns:
        series = df[col]

        # 先判断是否为数值型
        if pd.api.types.is_numeric_dtype(series):
            unique_count = series.nunique(dropna=True)
            unique_ratio = unique_count / n_rows if n_rows > 0 else 0

            # 如果唯一值太少，可能是类别编码，先归到分类型
            if unique_count <= unique_count_threshold and unique_ratio <= unique_ratio_threshold:
                categorical_cols.append(col)
            else:
                numeric_cols.append(col)
        else:
            categorical_cols.append(col)

    return numeric_cols, categorical_cols


def summarize_data_info(df, numeric_cols, categorical_cols):
    """
    输出数据基本信息
    """
    print("\n" + "=" * 50)
    print("数据基本信息")
    print("=" * 50)
    print(f"总行数：{df.shape[0]}")
    print(f"总列数：{df.shape[1]}")
    print(f"字段名称：{list(df.columns)}")
    print(f"数值型变量数：{len(numeric_cols)}")
    print(f"分类型变量数：{len(categorical_cols)}")
    print(f"数值型变量：{numeric_cols}")
    print(f"分类型变量：{categorical_cols}")
    print("=" * 50 + "\n")


def main():
    """
    主函数：先测试第1部分功能
    """

    # ====== 这里改成你自己的文件路径 ======
    file_path = r"C:\Users\彭宗健\Desktop\软件设置\试验数据.xlsx"
    output_dir = r"C:\Users\彭宗健\Desktop\软件设置\输出结果\第一步输出结果"

    # 如果读取Excel且要指定工作表，可填写工作表名称或序号
    # 例如：sheet_name = "Sheet1"
    # 如果不指定，填 None
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

        # 5. 输出基本信息
        summarize_data_info(df, numeric_cols, categorical_cols)

        print("第1部分运行完成。")

    except Exception as e:
        print(f"程序运行出错：{e}")


if __name__ == "__main__":
    main()