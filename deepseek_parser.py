# deepseek_parser.py
import requests
import json
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class DeepSeekParser:
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.api_url = "https://api.deepseek.com/v1/chat/completions"
    
    def parse_instruction(self, instruction: str) -> dict:
        """
        将自然语言指令解析为结构化操作
        """
        prompt = f"""
你是一个文档操作指令解析器。用户会输入对 Word 或 Excel 的操作指令，你需要解析出具体的操作。

支持的 Word 操作：
- bold（加粗）
- italic（斜体）
- underline（下划线）
- center（居中）
- left（左对齐）
- right（右对齐）
- font_size（字体大小）
- insert_table（插入表格）

支持的 Excel 操作：
- excel_bold（加粗）
- excel_center（居中）
- excel_width（设置列宽）
- excel_sum（求和）

位置表示：
- "第2段" → target: "paragraph", position: 2
- "第3行" → target: "row", row: 3
- "第2列" → target: "column", col: 2
- "第2行第3列" → target: "cell", row: 2, col: 3

请将以下指令解析为 JSON 格式，只返回 JSON，不要有其他文字：

指令：{instruction}

输出格式示例：
{{"action": "bold", "target": "paragraph", "position": 2}}
{{"action": "center", "target": "row", "row": 3}}
{{"action": "font_size", "target": "all", "size": 16}}
{{"action": "excel_bold", "target": "cell", "row": 2, "col": 3}}
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
                logger.error(f"响应内容: {response.text}")
                return {"action": "unknown", "error": "API调用失败"}
            
            result = response.json()
            content = result["choices"][0]["message"]["content"]
            
            # 提取JSON
            import re
            json_match = re.search(r'\{.*\}', content, re.DOTALL)
            if json_match:
                parsed = json.loads(json_match.group())
                logger.info(f"解析结果: {parsed}")
                return parsed
            else:
                logger.error(f"无法解析返回内容: {content}")
                return {"action": "unknown"}
                
        except Exception as e:
            logger.error(f"DeepSeek 解析失败: {e}")
            return {"action": "unknown", "error": str(e)}