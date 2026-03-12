import os
from dotenv import load_dotenv

# 加载 .env 文件中的环境变量
load_dotenv()

# 从环境变量中获取 API Key
ZHIPU_API_KEY = os.getenv("ZHIPU_API_KEY")

# 如果没有获取到，抛出错误
if not ZHIPU_API_KEY:
    raise ValueError("请在 .env 文件中设置 ZHIPU_API_KEY")