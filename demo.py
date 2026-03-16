"""
核心链路演示
展示完整的: 表格导入 → 字段选择 → 需求描述 → 代码生成 → 执行预览
"""

import sys
import io

# 解决 Windows 终端编码问题
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import pandas as pd
from tabulate import tabulate

from table_parser import create_sample_data, display_table_info
from api_client import generate_pandas_code
from code_executor import execute_code, format_result


def run_demo():
    """
    运行完整的演示流程
    """
    print("=" * 60)
    print("🚀 报表生成 AI Demo - 核心链路验证")
    print("=" * 60)
    print()
    
    # ========== 1. 加载示例数据 ==========
    print("📂 Step 1: 加载示例数据...")
    df_sales, df_product = create_sample_data()
    
    # 准备表格元信息
    tables_info = {
        "销售表": {
            "var_name": "df_sales",
            "columns": df_sales.columns.tolist(),
            "dtypes": {col: str(dtype) for col, dtype in df_sales.dtypes.items()},
            "sample": df_sales.head(3).to_dict(orient="records"),
        },
        "产品表": {
            "var_name": "df_product",
            "columns": df_product.columns.tolist(),
            "dtypes": {col: str(dtype) for col, dtype in df_product.dtypes.items()},
            "sample": df_product.to_dict(orient="records"),
        }
    }
    
    print("\n📊 销售表预览:")
    print(tabulate(df_sales.head(5), headers='keys', tablefmt='pretty', showindex=False))
    print(f"共 {len(df_sales)} 行")
    
    print("\n📊 产品表预览:")
    print(tabulate(df_product, headers='keys', tablefmt='pretty', showindex=False))
    print()
    
    # ========== 2. 模拟用户的拖拽选择 ==========
    print("🎯 Step 2: 用户字段选择 (模拟拖拽操作)")
    
    selected_fields = [
        {"table": "销售表", "column": "区域", "role": "分组"},
        {"table": "销售表", "column": "产品", "role": "分组"},
        {"table": "销售表", "column": "销售额", "role": "聚合(求和)"},
        {"table": "销售表", "column": "数量", "role": "聚合(求和)"},
    ]
    
    print("用户选择的字段:")
    for field in selected_fields:
        print(f"  📌 {field['table']}.{field['column']} → {field['role']}")
    print()
    
    # ========== 3. 用户输入需求描述 ==========
    print("💬 Step 3: 用户需求描述")
    user_description = "按区域和产品统计销售额和数量，计算平均单价，结果按销售额降序排列"
    print(f"  \"{user_description}\"")
    print()
    
    # ========== 4. 调用大模型生成代码 ==========
    print("🤖 Step 4: 调用 Qwen/QwQ-32B 生成代码...")
    print("  (正在调用硅基流动 API...)")
    
    code = generate_pandas_code(
        table_info=tables_info,
        selected_fields=selected_fields,
        user_description=user_description
    )
    
    print("\n📝 生成的代码:")
    print("-" * 40)
    print(code)
    print("-" * 40)
    print()
    
    # ========== 5. 执行代码 ==========
    print("⚙️ Step 5: 执行生成的代码...")
    
    success, result, error = execute_code(
        code,
        {"df_sales": df_sales, "df_product": df_product}
    )
    
    if success and result is not None:
        print("\n✅ 执行成功！")
        print("\n📋 结果预览:")
        print(format_result(result))
    elif error:
        print(f"\n❌ 执行出错: {error}")
        print("\n🔄 可以尝试重新生成或手动修改代码")
    else:
        print("\n⚠️ 代码执行完成，但没有返回结果")
    
    print()
    print("=" * 60)
    print("✨ Demo 完成！")
    print("=" * 60)


def interactive_demo():
    """
    交互式演示 - 用户可以输入自己的需求
    """
    print("=" * 60)
    print("🚀 报表生成 AI - 交互模式")
    print("=" * 60)
    print()
    
    # 加载数据
    df_sales, df_product = create_sample_data()
    
    tables_info = {
        "销售表": {
            "var_name": "df_sales",
            "columns": df_sales.columns.tolist(),
            "dtypes": {col: str(dtype) for col, dtype in df_sales.dtypes.items()},
            "sample": df_sales.head(3).to_dict(orient="records"),
        },
        "产品表": {
            "var_name": "df_product",
            "columns": df_product.columns.tolist(),
            "dtypes": {col: str(dtype) for col, dtype in df_product.dtypes.items()},
            "sample": df_product.to_dict(orient="records"),
        }
    }
    
    print("📊 可用的数据表:")
    print("\n销售表列: ", df_sales.columns.tolist())
    print("产品表列: ", df_product.columns.tolist())
    print()
    
    while True:
        print("-" * 40)
        user_input = input("💬 请输入你的分析需求 (输入 'q' 退出): ").strip()
        
        if user_input.lower() == 'q':
            print("👋 再见！")
            break
        
        if not user_input:
            continue
        
        # 简化版：不做字段选择，直接用自然语言
        selected_fields = [
            {"table": "销售表", "column": col, "role": "可用"}
            for col in df_sales.columns
        ]
        
        print("\n🤖 正在生成代码...")
        code = generate_pandas_code(
            table_info=tables_info,
            selected_fields=selected_fields,
            user_description=user_input
        )
        
        print("\n📝 生成的代码:")
        print(code)
        
        print("\n⚙️ 执行中...")
        success, result, error = execute_code(
            code,
            {"df_sales": df_sales, "df_product": df_product}
        )
        
        if success and result is not None:
            print("\n✅ 结果:")
            print(format_result(result))
        else:
            print(f"\n❌ 错误: {error}")
        
        print()


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == "-i":
        interactive_demo()
    else:
        run_demo()

