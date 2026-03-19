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
    从文本中提取指定的字段值，返回列表，每个元素是一个字典（对应一条记录）。
    :param text: 输入文本
    :param targets: 要提取的字段列表，例如 ["甲方", "乙方", "金额"]。如果为空或无效，返回空列表。
    :return: 列表，如 [{"甲方":"张三","金额":"1000元"}, ...]
    """
    logger.info(f"提取文本长度: {len(text)}，前100字符: {text[:100]}")
    if not isinstance(targets, list) or len(targets) == 0:
        logger.error("targets 必须为非空列表")
        return []
    
    target_str = "、".join(targets)
    
    # 构造示例，避免索引越界
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
            model="glm-4",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
            max_tokens=1000,
        )
        result_text = response.choices[0].message.content.strip()
        logger.info(f"AI返回原始内容: {result_text}")
        
        result = None
        try:
            result = json.loads(result_text)
        except json.JSONDecodeError:
            match = re.search(r'(\[.*\]|\{.*\})', result_text, re.DOTALL)
            if match:
                try:
                    result = json.loads(match.group())
                except json.JSONDecodeError:
                    pass
        
        # 确保返回列表
        if isinstance(result, list):
            return result
        elif isinstance(result, dict):
            # 如果是单个对象，包装成列表
            return [result]
        else:
            return []
    except Exception as e:
        logger.error(f"出错: {e}")
        return []
def extract_entities_safe(text: str, targets: list) -> list:
    """
    安全提取：自动处理大文件
    """
    MAX_LENGTH = 10000  # 每段最大字符数

    if len(text) > MAX_LENGTH:
        paragraphs = text.split('\n')
        all_results = []
        current_chunk = ""
        for para in paragraphs:
            if len(current_chunk) + len(para) < MAX_LENGTH:
                current_chunk += para + '\n'
            else:
                if current_chunk:
                    result = extract_entities(current_chunk, targets)
                    all_results.extend(result)  # 直接 extend
                current_chunk = para + '\n'
        if current_chunk:
            result = extract_entities(current_chunk, targets)
            all_results.extend(result)
        return all_results
    else:
        return extract_entities(text, targets)

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
            model="glm-4",
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