import json
import logging
import re
from zhipuai import ZhipuAI
from config import ZHIPU_API_KEY

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

client = ZhipuAI(api_key=ZHIPU_API_KEY)


def extract_entities(text: str, targets: list) -> list:
    """
    从文本中提取指定的字段值，返回列表
    自动选择普通或安全模式
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
        logger.info(f"调用AI接口提取字段（目标：{target_str}）...")
        response = client.chat.completions.create(
            model="glm-4-flash",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
            max_tokens=4000,
        )
        result_text = response.choices[0].message.content.strip()
        
        try:
            result = json.loads(result_text)
        except json.JSONDecodeError:
            match = re.search(r'(\[.*\]|\{.*\})', result_text, re.DOTALL)
            if match:
                result = json.loads(match.group())
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
    安全提取：按固定长度分段处理大文件（不丢失数据）
    """
    CHUNK_SIZE = 3000  # 每段字符数
    
    if len(text) <= CHUNK_SIZE:
        return extract_entities(text, targets)
    
    logger.info(f"文本过长 ({len(text)} 字符)，将分段处理")
    
    # 按固定长度强制分段
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


def parse_instruction(instruction: str) -> dict:
    """
    解析用户指令，返回包含意图和操作参数的字典。
    """
    prompt = f"""
    你是一个智能助手，需要理解用户的指令并返回结构化的解析结果。

    可能的意图类型：
    - fill_form: 用户希望根据文档内容填写表格或生成汇总表格
    - search: 用户想搜索文档中的信息
    - ask: 用户对文档内容提问
    - operate: 用户希望对文档进行编辑、排版、格式调整、内容提取等操作
    - unknown: 无法确定意图

    如果是 operate 意图，还需要识别具体的操作类型和参数。操作类型包括：

    【Word 操作】
    - bold: 加粗（用户可能说：加粗、粗体、bold）
    - italic: 斜体（斜体、italic）
    - underline: 下划线（下划线、underline）
    - center: 居中对齐（居中、居中对齐、center）
    - left: 左对齐（左对齐、left）
    - right: 右对齐（右对齐、right）
    - font_size: 设置字体大小，需要提取 size 参数（数字），例如“字体大小设为12”
    - insert_table: 插入表格，需要提取 rows 和 cols 参数（数字），例如“插入3行4列表格”
    - insert_image: 插入图片（例如“插入图片”）

    【Excel 操作】
    - excel_bold: 加粗表头或标题（例如“加粗表头”、“加粗标题”）
    - excel_center: 居中（Excel 单元格居中，例如“单元格居中”）
    - excel_width: 设置列宽，需要提取 width 参数（数字），例如“列宽设为15”
    - excel_sum: 求和或合计（例如“求和”、“合计”）
    - excel_chart: 插入图表（例如“插入图表”）

    【通用操作】
    - extract: 提取内容，需要提取 target 参数（要提取什么），例如“提取所有姓名”、“提取日期”
    - fill: 填表或填写（例如“填表”、“填写”）

    请以 JSON 格式返回，包含 intent 字段，如果是 operate 意图，还需要包含 action 和必要的参数字段。

    示例：
    指令："帮我填一下这个表格" -> {{"intent": "fill_form"}}
    指令："查找关于张三的信息" -> {{"intent": "search"}}
    指令："文档里说了什么" -> {{"intent": "ask"}}
    指令："把标题加粗" -> {{"intent": "operate", "action": "bold", "target": "标题"}}
    指令："设置字体大小为12" -> {{"intent": "operate", "action": "font_size", "size": 12}}
    指令："插入3行4列表格" -> {{"intent": "operate", "action": "insert_table", "rows": 3, "cols": 4}}
    指令："加粗表头" -> {{"intent": "operate", "action": "excel_bold"}}
    指令："列宽设为15" -> {{"intent": "operate", "action": "excel_width", "width": 15}}
    指令："求和" -> {{"intent": "operate", "action": "excel_sum"}}
    指令："提取所有姓名" -> {{"intent": "operate", "action": "extract", "target": "姓名"}}
    指令："今天天气怎么样" -> {{"intent": "unknown"}}

    用户指令："{instruction}"

    只返回 JSON 对象，不要有其他文字。
    """
    try:
        response = client.chat.completions.create(
            model="glm-4-flash",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
            max_tokens=300,
        )
        result_text = response.choices[0].message.content.strip()
        logger.info(f"解析指令原始返回: {result_text}")
        
        try:
            result = json.loads(result_text)
        except json.JSONDecodeError:
            match = re.search(r'(\{.*\})', result_text, re.DOTALL)
            if match:
                result = json.loads(match.group())
            else:
                raise
        return result
    except Exception as e:
        logger.error(f"解析指令出错: {e}")
        return {"intent": "unknown"}


if __name__ == "__main__":
    sample_text = "张三向李四借款1000元，于2023年5月1日归还。"
    
    print("默认抽取（人名、金额）：", extract_entities(sample_text, ["人名", "金额"]))
    print("只抽取人名：", extract_entities(sample_text, ["人名"]))
    print("抽取人名、日期：", extract_entities(sample_text, ["人名", "日期"]))
    print("抽取自定义类型：", extract_entities(sample_text, ["借款人", "借款金额"]))
    
    print("传入空列表：", extract_entities(sample_text, []))
    
    # 测试分段提取
    long_text = "这是一段很长的文本。" * 1000
    print("长文本分段测试：", len(extract_entities_safe(long_text, ["测试"])))
    
    test_instructions = [
        "帮我填一下这个表格",
        "查找关于张三的信息",
        "文档里说了什么",
        "把标题加粗",
        "设置字体大小为14",
        "插入2行3列表格",
        "加粗表头",
        "列宽设为20",
        "求和",
        "提取所有日期",
        "今天天气怎么样"
    ]
    for inst in test_instructions:
        print(f"指令 '{inst}' -> 解析结果: {parse_instruction(inst)}")