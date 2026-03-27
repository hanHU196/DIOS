import os
from dotenv import load_dotenv

# 加载 .env 文件中的环境变量
load_dotenv()

# 从环境变量中获取 API Key
DEEPSEEK_API_KEY = "sk-4cbb2ea6e387462383eaeefdbcaa3314"

# 如果没有获取到，抛出错误
if not DEEPSEEK_API_KEY:
    raise ValueError("请在 .env 文件中设置 API_KEY")