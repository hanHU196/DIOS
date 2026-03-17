import os
import re
import json
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ---------- 尝试导入队友模块，失败则使用模拟函数 ----------
# 文档读取模块(已整合)
try:
    from document_reader import DocumentReader
    logger.info("成功导入 document_reader.DocumentReader")
    
    #类的实例（全局使用）
    reader = DocumentReader()
    
    def read_documents(file_paths):
        """调用甲的类读取文档"""
        all_text = ""
        for path in file_paths:
            logger.info(f"📖 读取文件：{path}")
            try:
                # 调用read 方法
                text = reader.read(path)
                # 只取前1000字符作为预览，避免日志太长
                logger.info(f"  读取成功，长度：{len(text)} 字符")
                all_text += text + "\n"
            except FileNotFoundError:
                logger.error(f"文件不存在：{path}")
                all_text += f"[文件不存在：{path}]\n"
            except ValueError as e:
                logger.error(f"格式不支持：{e}")
                all_text += f"[格式不支持：{path}]\n"
            except Exception as e:
                logger.error(f"读取失败：{e}")
                all_text += f"[读取失败：{path}]\n"
        
        logger.info(f"所有文件读取完成，总字符数：{len(all_text)}")
        return all_text
        
except ImportError as e:
    logger.error(f"导入 document_reader 失败：{e}")
    logger.warning("使用模拟读取函数")
    
    def read_documents(file_paths):
        """模拟读取文档"""
        return "模拟文档内容：这是一个测试文档。\n甲方：XX公司\n乙方：YY大学\n金额：10000元"
# -------------------------------------------------------

# 表格生成模块
try:
    from excel_handler import fill_template as real_fill_template
    logger.info("成功导入 excel_handler.fill_template")
except ImportError:
    logger.warning("导入 excel_handler 失败，使用模拟函数")
    def real_fill_template(data, output_filename='filled_form.xlsx'):
        """模拟生成Excel"""
        df = pd.DataFrame([data])
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        df.to_excel(output_path, index=False)
        return output_path

# AI模块
try:
    from ai_module import parse_instruction, extract_entities
    logger.info("成功导入 ai_module")
except ImportError:
    logger.warning("导入 ai_module 失败，使用模拟函数")
    def parse_instruction(instruction):
        if '填表' in instruction or '表格' in instruction:
            return 'fill_form'
        elif '搜索' in instruction:
            return 'search'
        elif '问答' in instruction:
            return 'ask'
        else:
            return 'unknown'
    
    def extract_entities(text, targets):
        result = {}
        # 简单正则模拟提取
        if '姓名' in targets:
            names = re.findall(r'([\u4e00-\u9fa5]{2,3})', text)
            result['姓名'] = names[0] if names else '张三'
        if '金额' in targets:
            money = re.findall(r'(\d+\.?\d*)\s*元', text)
            result['金额'] = money[0] if money else '1000'
        if '日期' in targets:
            dates = re.findall(r'(\d{4}年\d{1,2}月\d{1,2}日)', text)
            result['日期'] = dates[0] if dates else '2023年1月1日'
        for t in targets:
            if t not in result:
                result[t] = f'示例{t}'
        return result
# ---------------------------------------------------------

app = Flask(__name__)

# 配置
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'txt', 'docx', 'md', 'xlsx'}  # 支持的文件类型
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 限制16MB

# 确保上传文件夹存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def parse_fields_from_instruction(instruction):
    """
    从自然语言指令中解析出要提取的字段列表。
    例如："提取甲方、乙方、金额" -> ["甲方", "乙方", "金额"]
    如果无法解析，返回空列表。
    """
    # 简单规则：查找“提取”后面的内容，按中文顿号、逗号分割
    match = re.search(r'提取[：:]\s*(.+)', instruction)
    if not match:
        match = re.search(r'提取\s*(.+)', instruction)
    if match:
        fields_str = match.group(1)
        # 按中文顿号、逗号、空格分割
        fields = re.split(r'[、，,\s]+', fields_str)
        # 过滤空字符串
        fields = [f.strip() for f in fields if f.strip()]
        return fields
    # 如果没有“提取”关键词，尝试将整个指令按标点分割作为字段（不准确，但作为备选）
    words = re.split(r'[、，,\s]+', instruction)
    # 只保留可能的中文词（长度≥2且不含数字字母）
    fields = [w for w in words if len(w) >= 2 and not re.search(r'[a-zA-Z0-9]', w)]
    return fields

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    # 获取上传的文件和指令
    file = request.files.get('file')
    instruction = request.form.get('command', '').strip()

    if not file or not instruction:
        return "请上传文件并输入指令", 400

    if not allowed_file(file.filename):
        return "不支持的文件格式", 400

    # 保存文件
    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    # ---------- 1. 读取文档内容（调用甲的模块） ----------
    try:
        # 假设 document_reader 有一个 read_document 函数，接受文件路径，返回文本字符串
        doc_text = document_reader.read_document(filepath)
    except Exception as e:
        return f"读取文档失败：{str(e)}", 500

    # ---------- 2. 解析指令意图（调用丁的 parse_instruction） ----------
    try:
        intent = ai_module.parse_instruction(instruction)
    except Exception as e:
        return f"解析指令失败：{str(e)}", 500

    # 如果不是填表意图，返回提示（基础版只支持填表）
    if intent != 'fill_form':
        return f"当前仅支持填表操作，您的指令意图为：{intent}，请尝试输入类似“提取甲方、乙方、金额”的指令。"

    # ---------- 3. 从指令中解析要提取的字段列表 ----------
    fields = parse_fields_from_instruction(instruction)
    if not fields:
        # 如果未能解析出字段，返回提示
        return "无法从指令中识别出要提取的字段，请明确指定，例如“提取甲方、乙方、金额”。"

    # ---------- 4. 调用 AI 提取信息（丁的 extract_entities） ----------
    try:
        extracted_data = ai_module.extract_entities(doc_text, fields)
    except Exception as e:
        return f"信息提取失败：{str(e)}", 500

    # ---------- 5. 生成 Excel 表格（调用乙的模块） ----------
    try:
        # 假设 excel_handler 有一个 fill_template 函数，接受数据字典和模板路径（可选）
        # 这里我们不需要模板，直接生成一个新表格（或使用默认模板）
        # 乙同学需要提供此函数，返回生成的 Excel 文件路径
        output_excel = excel_handler.fill_template(extracted_data, fields)
    except Exception as e:
        return f"生成表格失败：{str(e)}", 500

    # ---------- 6. 返回生成的 Excel 文件供下载 ----------
    return send_file(output_excel, as_attachment=True, download_name='result.xlsx')

if __name__ == '__main__':
    app.run(debug=True)