"""
代码执行模块
安全执行 AI 生成的 pandas 代码
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, Tuple
import traceback


# 允许使用的模块和函数
ALLOWED_BUILTINS = {
    'len', 'range', 'enumerate', 'zip', 'map', 'filter',
    'sum', 'min', 'max', 'abs', 'round',
    'list', 'dict', 'set', 'tuple', 'str', 'int', 'float', 'bool',
    'sorted', 'reversed',
    'True', 'False', 'None',
}


def create_safe_globals(dataframes: Dict[str, pd.DataFrame]) -> Dict[str, Any]:
    """
    创建安全的执行环境
    """
    safe_globals = {
        '__builtins__': {k: __builtins__[k] if isinstance(__builtins__, dict) else getattr(__builtins__, k) 
                        for k in ALLOWED_BUILTINS if hasattr(__builtins__, k) or (isinstance(__builtins__, dict) and k in __builtins__)},
        'pd': pd,
        'np': np,
    }
    
    # 添加 DataFrame 变量
    safe_globals.update(dataframes)
    
    return safe_globals


def execute_code(
    code: str,
    dataframes: Dict[str, pd.DataFrame],
    timeout: int = 30
) -> Tuple[bool, Any, str]:
    """
    执行生成的代码
    
    Args:
        code: 要执行的 Python 代码
        dataframes: 可用的 DataFrame 字典 {变量名: DataFrame}
        timeout: 超时时间（秒）
    
    Returns:
        (是否成功, 结果, 错误信息)
    """
    
    # 基本安全检查
    dangerous_keywords = [
        'import os', 'import sys', 'import subprocess',
        '__import__', 'eval(', 'exec(',
        'open(', 'file(',
        'os.', 'sys.', 'subprocess.',
    ]
    
    for keyword in dangerous_keywords:
        if keyword in code:
            return False, None, f"安全检查失败: 代码包含不允许的操作 '{keyword}'"
    
    # 创建执行环境
    safe_globals = create_safe_globals(dataframes)
    local_vars = {}
    
    try:
        # 执行代码
        exec(code, safe_globals, local_vars)
        
        # 获取结果
        if 'result' in local_vars:
            result = local_vars['result']
            return True, result, ""
        else:
            return False, None, "代码执行完成，但没有找到 'result' 变量"
    
    except Exception as e:
        error_msg = f"{type(e).__name__}: {str(e)}\n{traceback.format_exc()}"
        return True, None, error_msg


def format_result(result: Any) -> str:
    """
    格式化结果用于显示
    """
    if isinstance(result, pd.DataFrame):
        from tabulate import tabulate
        return tabulate(result, headers='keys', tablefmt='pretty', showindex=True)
    elif isinstance(result, pd.Series):
        return result.to_string()
    else:
        return str(result)


if __name__ == "__main__":
    # 测试代码执行
    from table_parser import create_sample_data
    
    df_sales, df_product = create_sample_data()
    
    test_code = """
result = df_sales.groupby('区域')['销售额'].sum().reset_index()
result.columns = ['区域', '总销售额']
"""
    
    success, result, error = execute_code(
        test_code,
        {"df_sales": df_sales, "df_product": df_product}
    )
    
    if success and result is not None:
        print("执行成功！")
        print(format_result(result))
    else:
        print(f"执行失败: {error}")


















































