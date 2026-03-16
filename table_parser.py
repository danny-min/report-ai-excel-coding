"""
表格解析模块
支持 Excel (.xlsx, .xls) 和 CSV 文件
"""

import pandas as pd
from pathlib import Path
from typing import Dict, List, Any, Tuple


def parse_table(file_path: str) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    解析表格文件，返回 DataFrame 和元信息
    
    Args:
        file_path: 文件路径
    
    Returns:
        (DataFrame, 元信息字典)
    """
    path = Path(file_path)
    
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")
    
    # 根据扩展名选择读取方式
    suffix = path.suffix.lower()
    
    if suffix == ".csv":
        df = pd.read_csv(file_path, encoding="utf-8-sig")
    elif suffix == ".xlsx":
        df = pd.read_excel(file_path, engine="openpyxl")
    elif suffix == ".xls":
        df = pd.read_excel(file_path, engine="xlrd")
    else:
        raise ValueError(f"不支持的文件格式: {suffix}")
    
    # 生成元信息
    meta = {
        "file_name": path.name,
        "table_name": path.stem,
        "columns": df.columns.tolist(),
        "dtypes": {col: str(dtype) for col, dtype in df.dtypes.items()},
        "row_count": len(df),
        "sample": df.head(3).to_dict(orient="records")
    }
    
    return df, meta


def get_column_type_display(dtype_str: str) -> str:
    """
    将 pandas dtype 转换为用户友好的显示
    """
    if "int" in dtype_str:
        return "整数"
    elif "float" in dtype_str:
        return "小数"
    elif "datetime" in dtype_str:
        return "日期时间"
    elif "object" in dtype_str:
        return "文本"
    elif "bool" in dtype_str:
        return "布尔"
    else:
        return dtype_str


def display_table_info(meta: Dict[str, Any]) -> str:
    """
    格式化显示表格信息
    """
    lines = [
        f"📊 表名: {meta['table_name']}",
        f"📝 行数: {meta['row_count']}",
        f"📋 列信息:",
    ]
    
    for col in meta["columns"]:
        dtype = meta["dtypes"].get(col, "unknown")
        type_display = get_column_type_display(dtype)
        lines.append(f"   - {col} ({type_display})")
    
    return "\n".join(lines)


def create_sample_data() -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    创建示例数据用于测试
    """
    # 销售数据
    sales_data = {
        "日期": pd.date_range("2024-01-01", periods=20, freq="D"),
        "销售额": [1000, 1500, 1200, 1800, 2000, 1600, 1400, 1900, 2100, 1700,
                  1100, 1300, 1600, 1900, 2200, 1800, 1500, 2000, 2300, 1900],
        "数量": [10, 15, 12, 18, 20, 16, 14, 19, 21, 17,
                11, 13, 16, 19, 22, 18, 15, 20, 23, 19],
        "区域": ["华东", "华北", "华南", "华东", "华北"] * 4,
        "产品": ["A", "B", "C", "A", "B"] * 4,
    }
    df_sales = pd.DataFrame(sales_data)
    
    # 产品信息
    product_data = {
        "产品": ["A", "B", "C"],
        "产品名称": ["产品A-标准版", "产品B-高级版", "产品C-旗舰版"],
        "类别": ["基础", "进阶", "高端"],
        "单价": [100, 150, 200],
    }
    df_product = pd.DataFrame(product_data)
    
    return df_sales, df_product


if __name__ == "__main__":
    # 测试示例数据
    df_sales, df_product = create_sample_data()
    
    print("=== 销售表 ===")
    print(df_sales.head())
    print()
    
    print("=== 产品表 ===")
    print(df_product)


















































