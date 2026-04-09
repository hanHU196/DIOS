# deepseek_parser.py
import requests
import json
import logging
import re

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class DeepSeekParser:
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.api_url = "https://api.deepseek.com/v1/chat/completions"
    
    def parse_instruction(self, instruction: str) -> dict:
        """
        将自然语言指令解析为结构化操作
        支持多个 JSON 对象（每行一个）的合并
        """
        prompt = f"""
你是一个文档操作指令解析器。用户会输入对 Word 的操作指令，你需要解析出具体的操作。

支持的 Word 操作：
- bold（加粗）
- italic（斜体）
- underline（下划线）
- center（居中）
- left（左对齐）
- right（右对齐）
- font_name（字体，如：宋体、黑体、楷体）
- font_size（字号，如：12、14、16、18）
- font_color（字体颜色：red、blue、green）

位置表示：
- "第2段" → "target": "paragraph", "position": 2
- "全文" → "target": "all"

请将以下指令解析为 JSON 对象，如果有多个操作，请合并成一个 JSON 对象。

指令：{instruction}

输出格式示例：
{{"action": "italic", "target": "paragraph", "position": 1, "font": "楷体", "size": 18, "color": "red"}}

注意：只输出一个 JSON 对象，不要输出多个，不要换行输出多个对象。
"""
        try:
            response = requests.post(
                self.api_url,
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type": "application/json"
                },
                json={
                    "model": "deepseek-chat",
                    "messages": [{"role": "user", "content": prompt}],
                    "temperature": 0.1,
                    "max_tokens": 300
                },
                timeout=30
            )
            
            if response.status_code != 200:
                logger.error(f"DeepSeek API 错误: {response.status_code}")
                return {"action": "unknown", "error": "API调用失败"}
            
            result = response.json()
            content = result["choices"][0]["message"]["content"]
            logger.info(f"DeepSeek 返回: {content}")
            
            # ========== 处理多种返回格式 ==========
            
            # 1. 尝试匹配单个 JSON 对象
            json_match = re.search(r'\{[^{}]*\}', content, re.DOTALL)
            if json_match:
                try:
                    parsed = json.loads(json_match.group())
                    logger.info(f"解析为单个对象: {parsed}")
                    return parsed
                except:
                    pass
            
            # 2. 尝试匹配多个 JSON 对象（每行一个），合并成一个
            lines = content.strip().split('\n')
            merged = {}
            
            for line in lines:
                line = line.strip()
                if line.startswith('{') and line.endswith('}'):
                    try:
                        obj = json.loads(line)
                        for key, value in obj.items():
                            if value is not None:
                                merged[key] = value
                    except:
                        continue
            
            if merged:
                # 确保 action 存在
                if 'action' not in merged:
                    # 从其他字段推断 action
                    if 'font_name' in merged:
                        merged['action'] = 'font_name'
                    elif 'font_size' in merged:
                        merged['action'] = 'font_size'
                    elif 'font_color' in merged:
                        merged['action'] = 'font_color'
                    elif 'bold' in str(merged.values()):
                        merged['action'] = 'bold'
                    elif 'italic' in str(merged.values()):
                        merged['action'] = 'italic'
                
                logger.info(f"合并多个对象结果: {merged}")
                return merged
            
            # 3. 尝试匹配 JSON 数组 [ {...}, {...} ]
            array_match = re.search(r'\[[\s\S]*\]', content, re.DOTALL)
            if array_match:
                try:
                    parsed_list = json.loads(array_match.group())
                    merged = {}
                    for obj in parsed_list:
                        for key, value in obj.items():
                            if value is not None:
                                merged[key] = value
                    logger.info(f"数组解析结果: {merged}")
                    return merged
                except:
                    pass
            
            logger.error(f"无法解析返回内容: {content}")
            return {"action": "unknown"}
                
        except Exception as e:
            logger.error(f"DeepSeek 解析失败: {e}")
            return {"action": "unknown", "error": str(e)}