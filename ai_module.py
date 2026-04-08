import json
import logging
import re
import time
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ==================== 计时器装饰器 ====================
def timer(func):
    """计时器装饰器，打印函数执行时间"""
    def wrapper(*args, **kwargs):
        start = time.time()
        result = func(*args, **kwargs)
        end = time.time()
        logger.info(f"⏱️ {func.__name__} 耗时: {end - start:.2f} 秒")
        return result
    return wrapper

def log_time(message):
    """打印时间日志"""
    logger.info(f"⏱️ {message}")
# =====================================================

# ==================== 模型配置 ====================
# 可选模型：
# - "ollama"            (本地 Ollama，推荐)
# - "deepseek-chat"     (DeepSeek API)
# - "glm-4-flash"       (智谱 AI)
# - "glm-4"             (智谱 AI 标准版)
# - "qwen-turbo"        (通义千问 API)
USE_MODEL = "qwen-turbo"  # ← 改成 ollama 使用本地模型

# Ollama 配置
OLLAMA_MODEL = "phi3.5:3.8b" # 可选: qwen2.5:0.5b, qwen2.5:1.5b, qwen2.5:3b, qwen2.5:7b,phi3.5:3.8b
OLLAMA_URL = "http://localhost:11434"
# =====================================================

# ==================== API 密钥配置 ====================
DEEPSEEK_API_KEY = "sk-4cbb2ea6e387462383eaeefdbcaa3314"
ZHIPU_API_KEY = "fb5f5b006b53493690d18d756963810a.f9mfzGzSWmNvJfys"
QWEN_API_KEY = "sk-63c82a8d5e9f4281bf19d1405cb7cc58"
# =====================================================

# 全局客户端
client = None
zhipu_client = None
ollama_available = False


def init_client():
    """初始化客户端"""
    global client, zhipu_client, ollama_available
    
    if USE_MODEL == "ollama":
        # 检查 Ollama 服务
        try:
            response = requests.get(f"{OLLAMA_URL}/api/tags", timeout=5)
            if response.status_code == 200:
                models = response.json().get("models", [])
                model_names = [m["name"] for m in models]
                logger.info(f"✅ Ollama 服务已连接")
                logger.info(f"📦 可用模型: {model_names}")
                
                # 检查模型是否存在
                model_base = OLLAMA_MODEL.split(":")[0]
                if not any(model_base in m for m in model_names):
                    logger.warning(f"⚠️ 模型 {OLLAMA_MODEL} 未找到")
                    logger.info(f"   请运行: ollama pull {OLLAMA_MODEL}")
                else:
                    ollama_available = True
                    logger.info(f"✅ 使用本地 Ollama 模型: {OLLAMA_MODEL}")
            else:
                logger.error("❌ Ollama 服务异常")
        except Exception as e:
            logger.error(f"❌ Ollama 服务未启动，请运行: ollama serve")
            logger.error(f"   错误: {e}")
            logger.info("将回退到 API 模式")
    
    elif USE_MODEL == "deepseek-chat":
        from openai import OpenAI
        client = OpenAI(
            api_key=DEEPSEEK_API_KEY,
            base_url="https://api.deepseek.com/v1"
        )
        logger.info("✅ 使用 DeepSeek 模型")
    
    elif USE_MODEL in ["glm-4-flash", "glm-4"]:
        from zhipuai import ZhipuAI
        zhipu_client = ZhipuAI(api_key=ZHIPU_API_KEY)
        logger.info(f"✅ 使用智谱 AI 模型: {USE_MODEL}")
    
    elif USE_MODEL == "qwen-turbo":
        import dashscope
        dashscope.api_key = QWEN_API_KEY
        logger.info(f"✅ 使用通义千问模型: {USE_MODEL}")


def call_ollama(prompt: str, max_tokens: int = 2000, temperature: float = 0.1) -> str:
    """调用本地 Ollama 模型"""
    response = requests.post(
        f"{OLLAMA_URL}/api/generate",
        json={
            "model": OLLAMA_MODEL,
            "prompt": prompt,
            "stream": False,
            "options": {
                "temperature": temperature,
                "num_predict": max_tokens
            }
        },
        timeout=120
    )
    
    if response.status_code == 200:
        return response.json()["response"]
    else:
        raise Exception(f"Ollama 调用失败: {response.text}")


@timer
def call_model(prompt: str, max_tokens: int = 2000) -> str:
    """统一调用接口"""
    try:
        start = time.time()
        
        if USE_MODEL == "ollama" and ollama_available:
            result = call_ollama(prompt, max_tokens)
            logger.info(f"   Ollama 本地调用: {time.time() - start:.2f}s, 输出长度: {len(result)}")
            return result
        
        elif USE_MODEL == "deepseek-chat":
            response = client.chat.completions.create(
                model=USE_MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.001,
                max_tokens=max_tokens,
            )
            result = response.choices[0].message.content.strip()
            logger.info(f"   DeepSeek API 调用: {time.time() - start:.2f}s, 输出长度: {len(result)}")
            return result
        
        elif USE_MODEL in ["glm-4-flash", "glm-4"]:
            response = zhipu_client.chat.completions.create(
                model=USE_MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.001,
                max_tokens=max_tokens,
            )
            result = response.choices[0].message.content.strip()
            logger.info(f"   智谱 AI API 调用: {time.time() - start:.2f}s, 输出长度: {len(result)}")
            return result
        
        elif USE_MODEL == "qwen-turbo":
            from dashscope import Generation
            response = Generation.call(
                model=USE_MODEL,
                messages=[{"role": "user", "content": prompt}],
                result_format='message',
                temperature=0.001,
                max_tokens=max_tokens,
            )
            result = response.output.choices[0].message.content.strip()
            logger.info(f"   通义千问 API 调用: {time.time() - start:.2f}s, 输出长度: {len(result)}")
            return result
        
        else:
            raise ValueError(f"不支持的模型: {USE_MODEL}")
            
    except Exception as e:
        logger.error(f"模型调用失败: {e}")
        return ""


# ==================== 信息提取函数 ====================

@timer
def extract_entities(text: str, targets: list) -> list:
    """
    通用提取：根据模板字段自动匹配文档内容
    """
    if not targets:
        return []
        
    has_target = any(t in text for t in targets)
    has_table_marker = "|" in text or "：" in text or ":" in text
    has_digit = any(c.isdigit() for c in text)
    
    if not (has_target or has_table_marker or has_digit):
        return []
    
    fields_str = "、".join(targets)
    
    prompt = f"""你是一个专业的信息提取助手。

任务：从以下文本中提取这些字段的值：{fields_str}。

要求：
1. **地域层级强制降维（最重要）**：如果字段名是“国家/地区”或类似宏观概念，但文本当前段落具体描述的是“省份”、“州”或“城市”（例如“湖北省”、“广东省”、“纽约州”），请【务必提取具体的省/市/州名称填入该字段】，绝对不要全局统一填国家名（如“中国”或“美国”）！同时，提取的语言必须与字段表头语言一致。
2. **数值与单位绝对标准化（最重要）**：
   - 观察字段名是否有单位（如“GDP（亿元）”），如果有，必须将原文数值换算为该单位后的【纯数字】。
   - 例如：原文为“56708.71亿元”，字段为“GDP总量（亿元）”，请输出 56708.71；
   - 若字段名【无单位】（如“人均GDP”、“人口”），请一律换算为【最基础单位】（金额换为“元”，人口换为“人”）。
   - 无论哪种情况，最终【只允许输出纯数字】！例如：原文“7.3万元”，必须输出 73000；原文“5775万人”，必须输出 57750000；“无”输出 0。绝对不准在 JSON 中返回带有“万”、“亿”、“元”等汉字的字符串，必须转为对应位数的纯数字！
3. **极速输出**：返回 JSON 数组，必须是单行无格式字符串！如果该段文本【完全没有】我们要的数据，请直接输出严格的 []，不要写任何多余的话！
4. **严格锁死键名**：键名必须与任务字段一字不差！
5.**语言绝对自适应（最重要）**：提取结果的语言【必须严格跟随字段表头的语言】！如果字段名是中文（如“大洲”、“国家”），请输出中文（如“亚洲”、“中国”）；如果字段名是英文（如“Continent”、“Country”），请输出英文（如“Asia”、“China”）。

文本：
{text[:4000]}

只返回 JSON 数组。"""

    max_retries = 3
    for attempt in range(max_retries):
        try:
            result_text = call_model(prompt, max_tokens=8000)
            clean_text = result_text.replace('```json', '').replace('```', '').strip()
            
            # ================= 🚀 提速核心：空数据极速放行 =================
            # 如果 AI 明确表示没有数据（输出 []），直接返回，绝对不浪费 6 秒钟去重试！
            if clean_text == '[]' or clean_text == '':
                return []
            # ==============================================================
            
            match_list = re.search(r'\[.*\]', clean_text, re.DOTALL)
            if match_list:
                result = json.loads(match_list.group())
                return result if isinstance(result, list) else [result]
                
            match_dict = re.search(r'\{.*\}', clean_text, re.DOTALL)
            if match_dict:
                return [json.loads(match_dict.group())]
                
            salvaged_data = []
            dict_matches = re.findall(r'\{.*?\}', clean_text, re.DOTALL)
            for d_str in dict_matches:
                try:
                    item = json.loads(d_str)
                    if isinstance(item, dict):
                        salvaged_data.append(item)
                except Exception:
                    continue
                    
            if salvaged_data:
                logger.warning(f"⚠️ 触发抢救模式，成功挽回 {len(salvaged_data)} 条记录！")
                return salvaged_data
                
            raise ValueError("未能解析出任何有效的 JSON 对象")
            
        except Exception as e:
            logger.warning(f"⚠️ 第 {attempt + 1} 次提取失败: {str(e)[:50]}... 准备重试")
            if attempt == max_retries - 1:
                return []
            import time
            time.sleep(2)  # 仅在真正报错时才停顿重试
    return []


@timer
def extract_entities_safe(text: str, targets: list) -> list:
    """安全提取：分段处理大文件"""
    CHUNK_SIZE = 2500
    
    if len(text) <= CHUNK_SIZE:
        return extract_entities(text, targets)
    
    logger.info(f"文本过长 ({len(text)} 字符)，将分段处理")
    
    # 分段
    chunks = []
    i = 0
    while i < len(text):
        chunk = text[i:i + CHUNK_SIZE]
        if i + CHUNK_SIZE < len(text):
            boundary = max(
                chunk.rfind('。'),
                chunk.rfind('\n'),
                chunk.rfind('，'),
                chunk.rfind('、'),
                chunk.rfind(' ')
            )
            if boundary > CHUNK_SIZE // 2:
                chunk = chunk[:boundary + 1]
                chunks.append(chunk)
                i += (boundary + 1)  # 关键：根据实际截断位置前进
                continue
        chunks.append(chunk)
        i += CHUNK_SIZE
    
    logger.info(f"共分为 {len(chunks)} 段")
    
    all_results = []
    for idx, chunk in enumerate(chunks, 1):
        logger.info(f"处理第 {idx}/{len(chunks)} 段")
        result = extract_entities(chunk, targets)
        if isinstance(result, list):
            all_results.extend(result)
        elif isinstance(result, dict):
            all_results.append(result)
    
    logger.info(f"分段提取完成，共获取 {len(all_results)} 条数据")
    return all_results


@timer
def extract_entities_safe_parallel(text: str, targets: list, max_workers=4) -> list:
    """并行分段处理大文件（已修复多线程抢占与乱序问题）"""
    # ================= 修复：将大块改回小块，防止大模型超限截断 =================
    CHUNK_SIZE = 2500
    # =======================================================================
    
    if len(text) <= CHUNK_SIZE:
        return extract_entities(text, targets)
    
    logger.info(f"文本过长 ({len(text)} 字符)，将分段并行处理 (workers={max_workers})")
    
    # 1. 分段逻辑
    chunks = []
    i = 0
    while i < len(text):
        chunk = text[i:i + CHUNK_SIZE]
        if i + CHUNK_SIZE < len(text):
            boundary = -1
            # 优先级：优先从换行符(表格行尾)切断，其次是句号，最后才是逗号和空格
            for sep in ['\n', '。', '；', '，', '、', ' ']:
                pos = chunk.rfind(sep)
                if pos > CHUNK_SIZE // 2:
                    boundary = pos
                    break  # 只要找到了高级别的安全边界，立刻停止，绝不使用低级别的空格
            
            if boundary != -1:
                chunk = chunk[:boundary + 1]
                chunks.append(chunk)
                i += (boundary + 1)
                continue
        chunks.append(chunk)
        i += CHUNK_SIZE
    
    logger.info(f"共分为 {len(chunks)} 段")
    
    # 2. 提前占位
    ordered_results = [[] for _ in range(len(chunks))]
    
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(extract_entities, chunk, targets): idx 
                   for idx, chunk in enumerate(chunks)}
        
        for future in as_completed(futures):
            idx = futures[future]  
            try:
                result = future.result()
                if isinstance(result, list):
                    ordered_results[idx] = result
                elif isinstance(result, dict):
                    ordered_results[idx] = [result]
                logger.info(f"  ✅ 第 {idx+1}/{len(chunks)} 段完成")
            except Exception as e:
                logger.error(f"  ❌ 第 {idx+1} 段处理失败: {e}")
    
    # 3. 按顺序拼接
    all_results = []
    for res in ordered_results:
        all_results.extend(res)
        
    logger.info(f"并行提取完成，共获取 {len(all_results)} 条数据")
    return all_results
# ==================== 指令解析函数 ====================

def parse_instruction(instruction: str) -> dict:
    """解析用户指令，返回包含意图和操作参数的字典"""
    start = time.time()
    
    instruction_lower = instruction.lower()
    
    result = {
        "intent": "unknown",
        "action": None,
        "target": None,
        "position": None,
        "row": None,
        "col": None,
        "size": None,
        "width": None,
        "rows": None,
        "cols": None
    }

    # 提取位置信息
    para_match = re.search(r"第(\d+)段", instruction)
    if para_match:
        result["position"] = int(para_match.group(1))
        result["target"] = "paragraph"
    
    row_match = re.search(r"第(\d+)行", instruction)
    if row_match:
        result["row"] = int(row_match.group(1))
        result["target"] = "row"
    
    col_match = re.search(r"第(\d+)列", instruction)
    if col_match:
        result["col"] = int(col_match.group(1))
        result["target"] = "column"
    
    if row_match and col_match:
        result["target"] = "cell"

    # 判断意图
    is_operate = any(k in instruction_lower for k in [
        "加粗", "斜体", "下划线", "居中", "左对齐", "右对齐",
        "设置字体", "字体大小", "插入表格", "加粗表头", "求和", "列宽"
    ])
    
    if is_operate:
        result["intent"] = "operate"
    elif any(k in instruction_lower for k in ["填表", "填写表格", "生成表格"]):
        result["intent"] = "fill_form"
    elif any(k in instruction_lower for k in ["搜索", "查找", "找到"]):
        result["intent"] = "search"
    elif any(k in instruction_lower for k in ["问", "是什么", "为什么"]):
        result["intent"] = "ask"

    # 识别操作类型
    if "加粗" in instruction_lower:
        result["action"] = "bold"
    elif "斜体" in instruction_lower:
        result["action"] = "italic"
    elif "下划线" in instruction_lower:
        result["action"] = "underline"
    elif "居中" in instruction_lower:
        result["action"] = "center"
    elif "左对齐" in instruction_lower:
        result["action"] = "left"
    elif "右对齐" in instruction_lower:
        result["action"] = "right"
    elif "字体大小" in instruction_lower or "设置字体" in instruction_lower:
        result["action"] = "font_size"
        size_match = re.search(r"(\d+)", instruction)
        if size_match:
            result["size"] = int(size_match.group(1))
        else:
            result["size"] = 12
    elif "插入表格" in instruction_lower:
        result["action"] = "insert_table"
        dims = re.findall(r"(\d+)", instruction)
        if len(dims) >= 2:
            result["rows"] = int(dims[0])
            result["cols"] = int(dims[1])
        else:
            result["rows"] = 3
            result["cols"] = 3
    elif "加粗表头" in instruction_lower:
        result["action"] = "excel_bold"
        result["target"] = "header"
    elif "求和" in instruction_lower:
        result["action"] = "excel_sum"
    elif "列宽" in instruction_lower:
        result["action"] = "excel_width"
        width_match = re.search(r"(\d+)", instruction)
        if width_match:
            result["width"] = int(width_match.group(1))
    elif "提取" in instruction_lower:
        result["action"] = "extract"
        match = re.search(r"提取(.*?)(?=[，。、]|$)", instruction)
        if match:
            result["target"] = match.group(1).strip()

    if result["target"] is None and result["intent"] == "operate":
        result["target"] = "all"
    
    logger.info(f"⏱️ parse_instruction 耗时: {time.time() - start:.3f} 秒")
    return result


def parse_filter_conditions(instruction: str) -> dict:
    """解析筛选条件"""
    prompt = f"""
从指令中提取筛选条件，返回 JSON。

指令：{instruction}

支持的筛选类型：
- date_range: 日期范围，需要 start 和 end
- numeric: 数值比较，需要 operator(>/</>=/<=/==) 和 value

只返回 JSON，不要其他文字。
"""
    
    try:
        result_text = call_model(prompt, max_tokens=500)
        if not result_text:
            return {"filters": []}
        
        json_match = re.search(r'\{.*\}', result_text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group())
        return {"filters": []}
    except Exception as e:
        logger.error(f"解析筛选条件失败: {e}")
        return {"filters": []}


# ==================== 初始化 ====================
init_client()


# ==================== 测试部分 ====================
if __name__ == "__main__":
    print(f"当前使用模型: {USE_MODEL}")
    
    if USE_MODEL == "ollama":
        print(f"Ollama 模型: {OLLAMA_MODEL}")
        print(f"Ollama 地址: {OLLAMA_URL}")
    
    print("正在测试模型连接...")
    
    if USE_MODEL == "ollama" and ollama_available:
        test_response = call_ollama("你好，请回复'OK'", max_tokens=10)
        if test_response:
            print(f"✅ 模型连接成功！回复: {test_response}")
        else:
            print("❌ 模型连接失败！")
            exit(1)
    else:
        test_response = call_model("你好，请回复'OK'", max_tokens=10)
        if test_response:
            print(f"✅ 模型连接成功！回复: {test_response}")
        else:
            print("❌ 模型连接失败！")
            exit(1)
    
    print("\n" + "="*50)
    
    sample_text = "张三向李四借款1000元，于2023年5月1日归还。"
    print("默认抽取（人名、金额）：", extract_entities(sample_text, ["人名", "金额"]))
    print("只抽取人名：", extract_entities(sample_text, ["人名"]))
    
    print("\n========== 指令解析测试 ==========")
    test_instructions = [
        "帮我填一下这个表格",
        "提取所有日期",
        "把第三行加粗",
        "设置第3段字体为16",
        "今天天气怎么样"
    ]
    for inst in test_instructions:
        result = parse_instruction(inst)
        print(f"指令 '{inst}' -> {result}")