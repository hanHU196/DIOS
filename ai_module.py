import json
import logging
import re
from zhipuai import ZhipuAI
from config import ZHIPU_API_KEY

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

client = ZhipuAI(api_key=ZHIPU_API_KEY)

def extract_entities(text: str, targets: list) -> dict:
    """
    从文本中提取指定的字段值。
    :param text: 输入文本
    :param targets: 要提取的字段列表，例如 ["甲方", "乙方", "金额"]。如果为空或无效，返回空字典。
    :return: 字典，键为字段名，值为提取的值。
    """
    # 防御性检查：如果 targets 不是列表或为空，返回空字典并记录错误
    if not isinstance(targets, list) or len(targets) == 0:
        logger.error("targets 必须为非空列表")
        return {}
    
    target_str = "、".join(targets)
    
    prompt = f"""
    你是一个信息提取助手。请从以下文本中提取以下字段的值：{target_str}。
    必须以**严格的JSON对象**形式返回，键是字段名，值是提取的内容。**绝对不要返回数组（列表）**。
    如果某个字段在文本中不存在，则值为null。
    
    例如：{{"{targets[0]}": "张三"}}
    
    文本内容：
    {text}
    
    只返回JSON对象，不要有任何其他文字、注释或标记。
    """
    
    try:
        logger.info(f"调用AI接口提取字段（目标：{target_str}）...")
        response = client.chat.completions.create(
            model="glm-4",  # 如果不可用，请改为 "glm-3-turbo"
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
            max_tokens=1000,
        )
        result_text = response.choices[0].message.content.strip()
        logger.info(f"AI返回原始内容: {result_text}")
        
        # 步骤1：尝试直接解析JSON
        result = None
        try:
            result = json.loads(result_text)
        except json.JSONDecodeError:
            # 步骤2：用正则提取可能的JSON部分（{} 或 []）
            match = re.search(r'(\{.*\}|\[.*\])', result_text, re.DOTALL)
            if match:
                try:
                    result = json.loads(match.group())
                except json.JSONDecodeError:
                    pass
        
        # 步骤3：如果result是列表，转换成字典
        if isinstance(result, list):
            dict_result = {}
            for item in result:
                if isinstance(item, dict) and "type" in item and "value" in item:
                    dict_result[item["type"]] = item["value"]
                elif isinstance(item, dict) and len(item) == 1:
                    for k, v in item.items():
                        dict_result[k] = v
            return dict_result
        
        # 步骤4：如果result是字典，直接返回
        elif isinstance(result, dict):
            return result
        
        # 步骤5：如果result是其他类型，包装成字典
        elif result is not None:
            return {"value": result}
        
        else:
            return {}
    
    except Exception as e:
        logger.error(f"出错: {e}")
        return {}

def parse_instruction(instruction: str) -> str:
    """
    解析用户指令，返回意图名称。
    可能的意图: "fill_form" (填表), "search" (搜索), "ask" (问答), "unknown" (未知)
    """
    prompt = f"""
    你是一个智能助手，需要理解用户的指令并分类到以下类别之一：
    - fill_form: 用户希望根据文档内容填写表格或生成汇总表格。
    - search: 用户想搜索文档中的信息。
    - ask: 用户对文档内容提问。
    - unknown: 无法确定意图。
    
    用户指令："{instruction}"
    
    请只返回一个单词作为类别，不要有其他解释文字。
    """
    try:
        logger.info("调用AI接口解析指令...")
        response = client.chat.completions.create(
            model="glm-4",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
            max_tokens=10,
        )
        intent = response.choices[0].message.content.strip().lower()
        if intent in ["fill_form", "search", "ask"]:
            return intent
        else:
            return "unknown"
    except Exception as e:
        logger.error(f"解析指令出错: {e}")
        return "unknown"

if __name__ == "__main__":
    # 测试实体抽取（必须传入 targets，但代码有防御检查）
    sample_text = "张三向李四借款1000元，于2023年5月1日归还。"
    
    print("抽取人名、金额：", extract_entities(sample_text, ["人名", "金额"]))
    print("只抽取人名：", extract_entities(sample_text, ["人名"]))
    print("抽取人名、日期：", extract_entities(sample_text, ["人名", "日期"]))
    print("抽取自定义类型：", extract_entities(sample_text, ["借款人", "借款金额"]))
    
    # 测试传入空列表的情况（不会报错，返回空字典）
    print("传入空列表：", extract_entities(sample_text, []))
    
    # 测试指令解析
    test_instructions = [
        "帮我填一下这个表格",
        "查找关于张三的信息",
        "文档里说了什么",
        "今天天气怎么样"
    ]
    for inst in test_instructions:
        print(f"指令 '{inst}' -> 意图: {parse_instruction(inst)}")