import json
import logging
import re
from zhipuai import ZhipuAI
from config import ZHIPU_API_KEY

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 初始化客户端
client = ZhipuAI(api_key=ZHIPU_API_KEY)

def extract_entities(text: str, targets: list = None) -> list:
    """
    从文本中提取指定的字段值。
    返回列表，每个元素是一个字典
    """
    if targets is None or len(targets) == 0:
        targets = ["人名", "金额"]
    
    target_str = "、".join(targets)
    
    prompt = f"""
    请从以下文本中提取以下字段的值：{target_str}。
    
    文本中包含多个条目，请提取**所有条目**。
    
    要求：
    1. 必须以JSON数组形式返回，每个数组元素是一个对象，包含所有字段
    2. 如果某个字段在条目中不存在，填null
    3. 不要有其他文字，只返回JSON数组
    
    例如：
    [
      {{"{targets[0]}": "值1", "{targets[1]}": "值2"}},
      {{"{targets[0]}": "值3", "{targets[1]}": "值4"}}
    ]
    
    文本内容：
    {text}
    
    只返回JSON数组，不要有其他文字。
    """
    try:
        logger.info(f"调用AI接口提取字段（目标：{target_str}）...")
        response = client.chat.completions.create(
            model="glm-4-flash",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
            max_tokens=16000,  # 增加token数
        )
        result_text = response.choices[0].message.content.strip()
        logger.info(f"AI返回原始内容长度: {len(result_text)}")
        
        # 尝试解析JSON
        try:
            result = json.loads(result_text)
            if isinstance(result, list):
                return result
            elif isinstance(result, dict):
                return [result]
            else:
                logger.error(f"返回类型错误: {type(result)}")
                return []
        except json.JSONDecodeError as e:
            logger.error(f"JSON解析失败: {e}")
            
            # 尝试用正则提取数组
            match = re.search(r'(\[.*\])', result_text, re.DOTALL)
            if match:
                try:
                    result = json.loads(match.group())
                    if isinstance(result, list):
                        logger.info(f"正则提取成功，获取到 {len(result)} 条数据")
                        return result
                except:
                    pass
            
            # 如果还是失败，尝试修复不完整的JSON
            logger.warning("尝试修复不完整的JSON...")
            # 去掉末尾不完整的内容
            if result_text.rstrip().endswith(','):
                result_text = result_text.rstrip()[:-1] + "]"
            elif not result_text.rstrip().endswith(']'):
                # 找到最后一个完整的对象
                last_brace = result_text.rfind('}')
                if last_brace > 0:
                    result_text = result_text[:last_brace+1] + "]"
            
            try:
                result = json.loads(result_text)
                if isinstance(result, list):
                    logger.info(f"修复成功，获取到 {len(result)} 条数据")
                    return result
            except:
                pass
            
            logger.error("无法修复JSON")
            return []
            
    except Exception as e:
        logger.error(f"出错: {e}")
        return []

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
    # 测试实体抽取
    sample_text = "张三向李四借款1000元，于2023年5月1日归还。"
    
    print("默认抽取（人名、金额）：", extract_entities(sample_text))
    print("只抽取人名：", extract_entities(sample_text, ["人名"]))
    print("抽取人名、日期：", extract_entities(sample_text, ["人名", "日期"]))
    print("抽取自定义类型：", extract_entities(sample_text, ["借款人", "借款金额"]))
    
    # 测试指令解析
    test_instructions = [
        "帮我填一下这个表格",
        "查找关于张三的信息",
        "文档里说了什么",
        "今天天气怎么样"
    ]
    for inst in test_instructions:
        print(f"指令 '{inst}' -> 意图: {parse_instruction(inst)}")
    