# ollama_client.py
import requests
import json
import re
import logging

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class OllamaClient:
    """Ollama 千问模型客户端"""
    
    def __init__(self, model="qwen2.5:1.5b", base_url="http://localhost:11434"):
        self.model = model
        self.base_url = base_url
        self._check_connection()
    
    def _check_connection(self):
        """检查 Ollama 服务是否运行"""
        try:
            response = requests.get(f"{self.base_url}/api/tags", timeout=5)
            if response.status_code == 200:
                models = response.json().get("models", [])
                model_names = [m["name"] for m in models]
                logger.info(f"✅ Ollama 服务已连接")
                logger.info(f"📦 可用模型: {model_names}")
                if not any(self.model in m for m in model_names):
                    logger.warning(f"⚠️ 警告: 模型 {self.model} 未找到，请运行: ollama pull {self.model}")
            else:
                logger.warning("⚠️ Ollama 服务异常")
        except Exception as e:
            logger.error(f"❌ Ollama 服务未启动，请运行: ollama serve")
            logger.error(f"   错误: {e}")
            raise
    
    def generate(self, prompt: str, max_tokens: int = 500, temperature: float = 0.1) -> str:
        """调用模型生成文本"""
        response = requests.post(
            f"{self.base_url}/api/generate",
            json={
                "model": self.model,
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
    
    def extract_entities(self, text: str, fields: list) -> list:
        """
        按字段提取信息
        返回 JSON 数组
        """
        if not fields:
            return []
        
        fields_str = "、".join(fields)
        
        prompt = f"""你是一个信息提取助手。请从以下文本中提取这些字段的值：{fields_str}。
返回 JSON 数组，每个元素是一个对象。如果字段不存在填 null。

文本：{text[:3000]}

只返回 JSON 数组，不要有其他文字。"""
        
        try:
            response = self.generate(prompt, max_tokens=1000)
            
            # 提取 JSON 数组
            match = re.search(r'(\[.*\])', response, re.DOTALL)
            if match:
                result = json.loads(match.group())
                if isinstance(result, list):
                    return result
                elif isinstance(result, dict):
                    return [result]
            
            # 尝试提取单个 JSON 对象
            match = re.search(r'(\{.*\})', response, re.DOTALL)
            if match:
                result = json.loads(match.group())
                return [result]
            
            logger.warning(f"⚠️ 无法解析返回内容: {response[:200]}")
            return []
            
        except Exception as e:
            logger.error(f"❌ 提取失败: {e}")
            return []
    
    def extract_entities_safe(self, text: str, fields: list, chunk_size: int = 2000) -> list:
        """
        安全提取：分段处理大文件
        """
        if len(text) <= chunk_size:
            return self.extract_entities(text, fields)
        
        logger.info(f"文本过长 ({len(text)} 字符)，将分段处理")
        
        # 分段
        chunks = []
        for i in range(0, len(text), chunk_size):
            chunk = text[i:i + chunk_size]
            # 在句子边界断开
            if i + chunk_size < len(text):
                boundary = max(chunk.rfind('。'), chunk.rfind('\n'), chunk.rfind('，'))
                if boundary > chunk_size // 2:
                    chunk = chunk[:boundary + 1]
            chunks.append(chunk)
        
        logger.info(f"共分为 {len(chunks)} 段")
        
        all_results = []
        for idx, chunk in enumerate(chunks, 1):
            logger.info(f"处理第 {idx}/{len(chunks)} 段，长度 {len(chunk)}")
            result = self.extract_entities(chunk, fields)
            if isinstance(result, list):
                all_results.extend(result)
            elif isinstance(result, dict):
                all_results.append(result)
        
        logger.info(f"分段提取完成，共获取 {len(all_results)} 条数据")
        return all_results


# 测试
if __name__ == "__main__":
    client = OllamaClient()
    
    test_text = "采购合同，甲方：XX科技有限公司，乙方：YY大学，合同金额：10000元，签订日期：2024年3月1日"
    fields = ["甲方", "乙方", "金额", "签订日期"]
    
    print(f"测试文本: {test_text}")
    print(f"提取字段: {fields}")
    print("-" * 40)
    
    result = client.extract_entities(test_text, fields)
    print(f"提取结果: {json.dumps(result, ensure_ascii=False, indent=2)}")