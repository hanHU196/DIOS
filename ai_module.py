import json
import logging
import re
from openai import OpenAI
from concurrent.futures import ThreadPoolExecutor, as_completed
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ==================== 配置 ====================
# 你的 DeepSeek API Key
DEEPSEEK_API_KEY = "sk-4cbb2ea6e387462383eaeefdbcaa3314"

# 初始化 DeepSeek 客户端
client = OpenAI(
    api_key=DEEPSEEK_API_KEY,
    base_url="https://api.deepseek.com/v1"
)

DEEPSEEK_MODEL = "deepseek-chat"  # 使用 chat 模型


# ==================== 信息提取函数 ====================

def call_deepseek(prompt: str, max_tokens: int = 2000) -> str:
    """调用 DeepSeek API"""
    try:
        response = client.chat.completions.create(
            model=DEEPSEEK_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
            max_tokens=max_tokens,
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        logger.error(f"调用 DeepSeek API 出错: {e}")
        return ""


def extract_entities(text: str, targets: list) -> list:
    """
    从文本中提取指定的字段值，返回列表
    """
    MAX_LEN = 4000
    
    # 如果文本太长，自动切换到安全模式
    if len(text) > MAX_LEN:
        logger.info(f"文本过长 ({len(text)} 字符)，自动切换到安全模式")
        return extract_entities_safe(text, targets)
    
    logger.info(f"提取文本长度: {len(text)}，前100字符: {text[:100]}")
    
    if not isinstance(targets, list) or len(targets) == 0:
        logger.error("targets 必须为非空列表")
        return []
    
    target_str = "、".join(targets)
    
    if len(targets) >= 2:
        example = f'{{"{targets[0]}": "值1", "{targets[1]}": "值2"}}'
    else:
        example = f'{{"{targets[0]}": "值1"}}'
    
    prompt = f"""
    你是一个信息提取助手。请从以下文本中提取以下字段的值：{target_str}。

    文本中包含多个条目，请提取**所有条目**。
    
    要求：
    1. 必须以 JSON **数组**形式返回，每个数组元素是一个对象，包含所有字段。
    2. 如果某个字段在文本中不存在，则值为 null。
    
    例如：
    [
      {example},
      {example}
    ]
    
    文本内容：
    {text}
    
    只返回 JSON 数组，不要有任何其他文字、注释或标记。
    """
    
    try:
        logger.info(f"调用 DeepSeek 提取字段（目标：{target_str}）...")
        
        result_text = call_deepseek(prompt)
        
        if not result_text:
            logger.error("DeepSeek 返回为空")
            return []
        
        logger.info(f"AI返回原始内容: {result_text[:200]}...")
        
        try:
            result = json.loads(result_text)
        except json.JSONDecodeError:
            match = re.search(r'(\[.*\]|\{.*\})', result_text, re.DOTALL)
            if match:
                try:
                    result = json.loads(match.group())
                except json.JSONDecodeError:
                    return []
            else:
                return []
        
        if isinstance(result, list):
            logger.info(f"成功提取 {len(result)} 条数据")
            return result
        elif isinstance(result, dict):
            return [result]
        else:
            return []
            
    except Exception as e:
        logger.error(f"出错: {e}")
        return []


def extract_entities_safe(text: str, targets: list) -> list:
    """
    安全提取：按固定长度分段处理大文件
    """
    CHUNK_SIZE = 3000
    
    if len(text) <= CHUNK_SIZE:
        return extract_entities(text, targets)
    
    logger.info(f"文本过长 ({len(text)} 字符)，将分段处理")
    
    chunks = []
    for i in range(0, len(text), CHUNK_SIZE):
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
    
    logger.info(f"共分为 {len(chunks)} 段")
    
    all_results = []
    for idx, chunk in enumerate(chunks, 1):
        logger.info(f"处理第 {idx}/{len(chunks)} 段，长度 {len(chunk)}")
        result = extract_entities(chunk, targets)
        if isinstance(result, list):
            all_results.extend(result)
        elif isinstance(result, dict):
            all_results.append(result)
    
    logger.info(f"分段提取完成，共获取 {len(all_results)} 条数据")
    return all_results

def extract_entities_safe_parallel(text: str, targets: list, max_workers=4) -> list:
    """
    并行分段处理大文件
    max_workers: 同时处理的段数（建议 4-8）
    """
    CHUNK_SIZE = 3000  # 每段字符数
    
    if len(text) <= CHUNK_SIZE:
        return extract_entities(text, targets)
    
    logger.info(f"文本过长 ({len(text)} 字符)，将分段并行处理 (workers={max_workers})")
    
    # 1. 分段
    chunks = []
    for i in range(0, len(text), CHUNK_SIZE):
        chunk = text[i:i + CHUNK_SIZE]
        # 尽量在句子边界断开
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
    
    logger.info(f"共分为 {len(chunks)} 段")
    
    # 2. 并行处理
    all_results = []
    
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # 提交所有任务
        futures = {executor.submit(extract_entities, chunk, targets): idx 
                   for idx, chunk in enumerate(chunks)}
        
        # 收集结果
        for future in as_completed(futures):
            idx = futures[future]
            try:
                result = future.result()
                if isinstance(result, list):
                    all_results.extend(result)
                elif isinstance(result, dict):
                    all_results.append(result)
                logger.info(f"  ✅ 第 {idx+1}/{len(chunks)} 段完成，当前共 {len(all_results)} 条")
            except Exception as e:
                logger.error(f"  ❌ 第 {idx+1} 段处理失败: {e}")
    
    logger.info(f"并行分段提取完成，共获取 {len(all_results)} 条数据")
    return all_results
# ==================== 指令解析函数 ====================
def parse_instruction(instruction: str) -> dict:
    """
    解析用户指令，返回包含意图和操作参数的字典。
    支持位置定位。
    """
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

    # ========== 1. 提取位置信息 ==========
    # 匹配 "第2段"
    para_match = re.search(r"第(\d+)段", instruction)
    if para_match:
        result["position"] = int(para_match.group(1))
        result["target"] = "paragraph"
    
    # 匹配 "第3行"
    row_match = re.search(r"第(\d+)行", instruction)
    if row_match:
        result["row"] = int(row_match.group(1))
        result["target"] = "row"
    
    # 匹配 "第2列"
    col_match = re.search(r"第(\d+)列", instruction)
    if col_match:
        result["col"] = int(col_match.group(1))
        result["target"] = "column"
    
    # 组合匹配 "第2行第3列"
    if row_match and col_match:
        result["target"] = "cell"

    # ========== 2. 判断意图 ==========
    # 先检查是否是操作类指令
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

    # ========== 3. 识别操作类型 ==========
    # Word 操作
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
    # 字体大小（支持多种说法）
    elif "字体大小" in instruction_lower or "设置字体" in instruction_lower:
        result["action"] = "font_size"
        size_match = re.search(r"(\d+)", instruction)
        if size_match:
            result["size"] = int(size_match.group(1))
        else:
            result["size"] = 12  # 默认
    # 插入表格
    elif "插入表格" in instruction_lower:
        result["action"] = "insert_table"
        dims = re.findall(r"(\d+)", instruction)
        if len(dims) >= 2:
            result["rows"] = int(dims[0])
            result["cols"] = int(dims[1])
        else:
            result["rows"] = 3
            result["cols"] = 3
    # Excel 操作
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
    # 提取操作
    elif "提取" in instruction_lower:
        result["action"] = "extract"
        match = re.search(r"提取(.*?)(?=[，。、]|$)", instruction)
        if match:
            result["target"] = match.group(1).strip()

    # ========== 4. 如果没有提取到 target，设置默认值 ==========
    if result["target"] is None and result["intent"] == "operate":
        result["target"] = "all"

    return result

#==================== 筛选条件解析函数 ====================
def parse_filter_conditions(instruction: str) -> dict:
    """
    用 AI 解析用户指令中的筛选条件
    返回结构化筛选条件
    """
    prompt = f"""
    你是一个筛选条件解析助手。用户会输入指令，你需要从中提取筛选条件。

    支持的筛选类型：
    1. 日期范围：如 "2020/7/1到2020/8/31"、"2020-07-01至2020-08-31"
    2. 数值比较：如 "金额大于1000"、"GDP超过5000亿"、"人口少于100万"
    3. 文本匹配：如 "城市包含'上海'"、"项目名称等于'GDP'"

    请以 JSON 格式返回筛选条件，格式如下：

    示例1："提取2020/7/1到2020/8/31的数据"
    返回：{{
        "filters": [
            {{
                "type": "date_range",
                "column": "日期",
                "start": "2020-07-01",
                "end": "2020-08-31"
            }}
        ]
    }}

    示例2："筛选金额大于1000的记录"
    返回：{{
        "filters": [
            {{
                "type": "numeric",
                "column": "金额",
                "operator": ">",
                "value": 1000
            }}
        ]
    }}

    示例3："提取GDP超过5000亿且人口少于100万的城市"
    返回：{{
        "filters": [
            {{"type": "numeric", "column": "GDP", "operator": ">", "value": 5000}},
            {{"type": "numeric", "column": "人口", "operator": "<", "value": 100}}
        ]
    }}

    用户指令：{instruction}

    只返回 JSON 对象，不要有其他文字。
    """
    
    try:
        result_text = call_deepseek(prompt, max_tokens=500)
        if not result_text:
            return {"filters": []}
        
        # 提取 JSON
        import re
        json_match = re.search(r'\{.*\}', result_text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group())
        return {"filters": []}
    except Exception as e:
        logger.error(f"解析筛选条件失败: {e}")
        return {"filters": []}




# ==================== 测试部分 ====================
if __name__ == "__main__":
    # 先测试 DeepSeek 连接
    print("正在测试 DeepSeek 连接...")
    test_response = call_deepseek("你好，请回复'OK'", max_tokens=10)
    if test_response:
        print(f"DeepSeek 连接成功！回复: {test_response}")
    else:
        print("DeepSeek 连接失败！请检查 API Key 是否正确")
        print("获取 API Key: https://platform.deepseek.com/")
        exit(1)
    
    print("\n" + "="*50)
    
    sample_text = "张三向李四借款1000元，于2023年5月1日归还。"
    print("默认抽取（人名、金额）：", extract_entities(sample_text, ["人名", "金额"]))
    print("只抽取人名：", extract_entities(sample_text, ["人名"]))
    print("抽取人名、日期：", extract_entities(sample_text, ["人名", "日期"]))
    print("抽取自定义类型：", extract_entities(sample_text, ["借款人", "借款金额"]))
    
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