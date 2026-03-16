"""
测试 API 连接 - 生成 pandas 代码
"""
import sys
import io

if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import requests
import json

API_URL = "https://api.siliconflow.cn/v1/chat/completions"
API_KEY = "sk-nrryzlftodhwachgnqqmnvwiqnkuhtskjjuezthoofgxprrz"

# 测试生成 pandas 代码
system_prompt = """你是一个 pandas 代码生成器。只输出可执行的 Python 代码，用 ```python 包裹。结果存入 result 变量。"""

user_prompt = """
数据表 df_sales 包含列: 日期, 销售额, 数量, 区域, 产品
需求: 按区域统计销售额总和，按销售额降序排列
"""

payload = {
    "model": "Qwen/QwQ-32B",
    "messages": [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt}
    ],
    "stream": False,
    "max_tokens": 1024,
    "temperature": 0.7,
    "top_p": 0.7,
}

headers = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json"
}

print("发送代码生成请求...")

try:
    response = requests.post(API_URL, headers=headers, json=payload, timeout=180)
    print(f"状态码: {response.status_code}")
    
    if response.status_code == 200:
        data = response.json()
        content = data['choices'][0]['message']['content']
        print("\n✅ 生成成功!")
        print("=" * 50)
        print(content)
        print("=" * 50)
    else:
        print(f"错误: {response.text}")
except Exception as e:
    print(f"错误: {e}")

