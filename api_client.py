"""
硅基流动 API 客户端
调用 Qwen/QwQ-32B 模型生成 pandas 代码
"""

import requests
import json
from config import SILICONFLOW_API_URL, SILICONFLOW_API_KEY, MODEL_NAME, DEFAULT_PARAMS


def call_llm(prompt: str, system_prompt: str = None) -> str:
    """
    调用硅基流动的大模型 API
    
    Args:
        prompt: 用户提示词
        system_prompt: 系统提示词（可选）
    
    Returns:
        模型生成的文本
    """
    messages = []
    
    if system_prompt:
        messages.append({
            "role": "system",
            "content": system_prompt
        })
    
    messages.append({
        "role": "user",
        "content": prompt
    })
    
    payload = {
        "model": MODEL_NAME,
        "messages": messages,
        "stream": False,
        **DEFAULT_PARAMS
    }
    
    headers = {
        "Authorization": f"Bearer {SILICONFLOW_API_KEY}",
        "Content-Type": "application/json"
    }
    
    try:
        response = requests.post(
            SILICONFLOW_API_URL,
            headers=headers,
            json=payload,
            timeout=300  # 增加到5分钟
        )
        response.raise_for_status()
        
        result = response.json()
        return result["choices"][0]["message"]["content"]
    
    except requests.exceptions.RequestException as e:
        return f"API 调用失败: {str(e)}"
    except (KeyError, IndexError) as e:
        return f"解析响应失败: {str(e)}"


def generate_pandas_code(
    table_info: dict,
    selected_fields: list,
    user_description: str
) -> str:
    """
    根据表格信息和用户需求生成 pandas 代码
    
    Args:
        table_info: 表格元信息 {表名: {columns: [...], sample: [...], dtypes: {...}}}
        selected_fields: 用户选择的字段 [{"table": "表名", "column": "列名", "role": "分组/聚合/筛选"}, ...]
        user_description: 用户的自然语言描述
    
    Returns:
        生成的 Python 代码
    """
    
    # 构建系统提示词
    system_prompt = """你是一个专业的 pandas 数据分析代码生成器。

规则：
1. 只输出可直接执行的 Python 代码，不要任何解释文字
2. 代码用 ```python 和 ``` 包裹
3. 数据已经加载为 DataFrame 变量（变量名见下方）
4. 最终结果必须存入名为 `result` 的变量
5. 不要使用 print 语句
6. 不要读写文件
7. 确保代码简洁高效"""

    # 构建表格信息描述
    tables_desc = []
    for table_name, info in table_info.items():
        var_name = info.get("var_name", f"df_{table_name}")
        cols_desc = []
        for col in info["columns"]:
            dtype = info["dtypes"].get(col, "unknown")
            cols_desc.append(f"  - {col} ({dtype})")
        
        sample_str = ""
        if info.get("sample"):
            import pandas as pd
            sample_df = pd.DataFrame(info["sample"])
            sample_str = f"\n示例数据:\n{sample_df.to_string(index=False)}"
        
        tables_desc.append(f"""### {var_name} ({table_name})
列信息:
{chr(10).join(cols_desc)}
{sample_str}""")
    
    # 构建字段选择描述
    fields_desc = []
    for field in selected_fields:
        fields_desc.append(f"- {field['table']}.{field['column']} → {field['role']}")
    
    # 构建完整提示词
    prompt = f"""## 可用数据表

{chr(10).join(tables_desc)}

## 用户选择的字段
{chr(10).join(fields_desc)}

## 用户需求描述
{user_description}

请生成 pandas 代码来完成上述需求。"""

    # 调用模型
    response = call_llm(prompt, system_prompt)
    
    # 提取代码
    code = extract_code(response)
    return code


def extract_code(response: str) -> str:
    """
    从模型响应中提取 Python 代码
    """
    import re
    
    # 尝试提取 ```python ... ``` 包裹的代码
    pattern = r"```python\s*(.*?)\s*```"
    matches = re.findall(pattern, response, re.DOTALL)
    
    if matches:
        return matches[0].strip()
    
    # 尝试提取 ``` ... ``` 包裹的代码
    pattern = r"```\s*(.*?)\s*```"
    matches = re.findall(pattern, response, re.DOTALL)
    
    if matches:
        return matches[0].strip()
    
    # 如果没有代码块，返回整个响应
    return response.strip()


if __name__ == "__main__":
    # 简单测试
    print("测试 API 连接...")
    result = call_llm("你好，请用一句话介绍自己")
    print(f"模型响应: {result}")

