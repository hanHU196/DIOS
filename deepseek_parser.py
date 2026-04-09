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
    
    def parse_instruction(self, instruction: str, file_type: str = "word") -> dict:
        """
        将自然语言指令解析为结构化操作
        file_type: "word" 或 "excel"
        """
        if file_type == "word":
            prompt = self._get_word_prompt(instruction)
        else:
            prompt = self._get_excel_prompt(instruction)
        
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
            
            # 提取JSON
            json_match = re.search(r'\[[\s\S]*\]|\{[\s\S]*\}', content, re.DOTALL)
            if not json_match:
                return {"action": "unknown"}
            
            json_str = json_match.group()
            
            if json_str.startswith('['):
                parsed_list = json.loads(json_str)
                merged = {"action": "unknown"}
                for item in parsed_list:
                    if isinstance(item, dict):
                        for key, value in item.items():
                            if value is not None:
                                merged[key] = value
                        if merged.get('action') == 'unknown' and item.get('action'):
                            merged['action'] = item['action']
                return merged
            else:
                return json.loads(json_str)
                
        except Exception as e:
            logger.error(f"DeepSeek 解析失败: {e}")
            return {"action": "unknown", "error": str(e)}
    
    def _get_word_prompt(self, instruction):
        """Word 操作的提示词"""
        return f"""
你是一个 Word 文档操作指令解析器。请将以下指令解析为 JSON 格式。

支持的 Word 操作：
- bold（加粗）
- italic（斜体）
- underline（下划线）
- center（居中）
- left（左对齐）
- right（右对齐）
- font_name（字体）
- font_size（字号）
- font_color（字体颜色）

位置表示：
- "第2段" → "target": "paragraph", "position": 2
- "全文" → "target": "all"

指令：{instruction}

输出格式示例（只返回 JSON）：
{{"action": "bold", "target": "paragraph", "position": 2}}
{{"action": "center", "target": "all"}}
{{"action": "font_name", "target": "paragraph", "position": 2, "font": "宋体"}}
{{"action": "font_size", "target": "all", "size": 14}}
"""
    
    def _get_excel_prompt(self, instruction):
        """Excel 操作的提示词"""
        return f"""
你是一个 Excel 文档操作指令解析器。请将以下指令解析为 JSON 格式。

支持的 Excel 操作：
- excel_bold（加粗）
- excel_italic（斜体）
- excel_center（居中）
- excel_left（左对齐）
- excel_right（右对齐）
- excel_font_name（字体）
- excel_font_size（字号）
- excel_font_color（字体颜色）
- excel_width（设置列宽）
- excel_height（设置行高）
- excel_merge（合并单元格）

位置表示：
- "第2行" → "row": 2
- "第3列" → "col": 3
- "第2行第3列" → "row": 2, "col": 3
- "A1单元格" → "cell": "A1"
- "第2行到第5行" → "row_start": 2, "row_end": 5
- "第2-4行" → "row_start": 2, "row_end": 4

指令：{instruction}

输出格式示例（只返回 JSON）：
{{"action": "excel_bold", "row": 2, "col": 3}}
{{"action": "excel_center", "row_start": 2, "row_end": 5}}
{{"action": "excel_width", "col": 2, "width": 20}}
{{"action": "excel_merge", "row_start": 1, "col_start": 1, "row_end": 1, "col_end": 3}}
"""