import os
import re
import json
import logging
import pandas as pd
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename


# 导入队友模块（需要确保文件存在）
import document_reader  # 甲负责：读取文档
import excel_handler    # 乙负责：处理表格
import ai_module        # 丁负责：AI 接口


app = Flask(__name__)

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
    import excel_handler  # ✅ 直接导入整个模块
    logger.info("✅ 成功导入 excel_handler")
    
    # 检查关键函数是否存在
    if not hasattr(excel_handler, 'fill_excel_with_data'):
        logger.error("excel_handler 缺少 fill_excel_with_data 函数")
        raise ImportError("缺少必要函数")
        
except ImportError as e:
    logger.error(f"❌ 导入 excel_handler 失败：{e}")
    logger.warning("使用模拟填表函数")
    
    # 创建模拟模块（因为app已经定义，这里可以用app.config）
    class MockExcelHandler:
        @staticmethod
        def fill_excel_with_data(template_path, data_list, output_path):
            df = pd.DataFrame(data_list)
            df.to_excel(output_path, index=False)
            return output_path
        
        @staticmethod
        def MongoDBHandler():
            class MockMongo:
                def insert_data(self, *args, **kwargs):
                    logger.info("模拟MongoDB插入")
                    return None
            return MockMongo()
    
    excel_handler = MockExcelHandler()  # 用模拟对象替换

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



# 配置
UPLOAD_FOLDER = 'uploads'
ALLOWED_DOC_EXTENSIONS = {'txt', 'docx', 'md', 'xlsx'}  # 支持的文件类型
ALLOWED_TEMPLATE_EXTENSIONS =   {'xlsx','xls'}  # 模板文件类型
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 限制16MB

# 确保上传文件夹存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename, allowed_set=ALLOWED_DOC_EXTENSIONS):
    return '.' in filename and  filename.rsplit('.', 1)[1].lower() in allowed_set

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
    doc_file = request.files.get('document')
    template_file = request.files.get('template')
    instruction = request.form.get('command', '').strip()

    if not doc_file or not template_file or not instruction:
        return "请上传文档和模板文件，并输入指令", 400

    # 检查文件格式
    if not allowed_file(doc_file.filename, ALLOWED_DOC_EXTENSIONS):
        return "文档格式不支持，请上传 txt/docx/md/xlsx 文件", 400
    if not allowed_file(template_file.filename, ALLOWED_TEMPLATE_EXTENSIONS):
        return "模板格式不支持，请上传 Excel 文件（.xlsx 或 .xls）", 400

    # 保存文档
    doc_filename = secure_filename(doc_file.filename)
    doc_path = os.path.join(app.config['UPLOAD_FOLDER'], doc_filename)
    doc_file.save(doc_path)
    # 保存模板（重命名避免覆盖，确保有.xlsx后缀）
    template_filename = secure_filename(template_file.filename)
    # 确保文件名有 .xlsx 后缀
    if not template_filename.endswith('.xlsx'):
        template_filename = template_filename + '.xlsx'
    template_path = os.path.join(app.config['UPLOAD_FOLDER'], 'template_' + template_filename)
    template_file.save(template_path)
    logger.info(f"✅ 模板保存为：{template_path}")

    # ---------- 以下为原有流程（调用甲、丁） ----------
    try:
        doc_text = document_reader.read_document(doc_path)
    except Exception as e:
        return f"读取文档失败：{str(e)}", 500

    try:
        intent = ai_module.parse_instruction(instruction)
    except Exception as e:
        return f"解析指令失败：{str(e)}", 500

    if intent != 'fill_form':
        return f"当前仅支持填表操作，您的指令意图为：{intent}"
    # ---------- 正确的流程：先用乙解析模板 ----------
    try:
        # 1. 先用乙解析模板，得到真正的字段
        template_fields = excel_handler.parse_excel_template(template_path)
        logger.info(f"📋 模板字段：{template_fields}")
    except Exception as e:
        return f"解析模板失败：{str(e)}", 500

        # 2. 用模板字段去调用AI
    try:
        extracted_data = ai_module.extract_entities(doc_text, template_fields)
        logger.info(f"✅ AI提取结果：{extracted_data}")
        logger.info(f"提取到 {len(extracted_data)} 行数据")
    except Exception as e:
        return f"信息提取失败：{str(e)}", 500

    # 3. 准备数据（extracted_data 现在已经是列表了）
    if isinstance(extracted_data, list):
        data_list = extracted_data  # 直接使用
    else:
        # 兼容旧版本，如果是字典就包装成列表
        data_list = [extracted_data]

    # 创建输出目录
    output_dir = "outputs"
    os.makedirs(output_dir, exist_ok=True)

    # 生成唯一输出文件名（避免多人同时使用覆盖）
    from datetime import datetime
    output_filename = f"result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    output_path = os.path.join(output_dir, output_filename)

    # 先填表（必须成功）
    # ===== 调试信息 =====
    print("\n" + "="*60)
    print("🔍 表格生成调试信息")
    print("="*60)
    print(f"1. 模板路径：{template_path}")
    print(f"2. 模板是否存在：{os.path.exists(template_path)}")
    if os.path.exists(template_path):
        print(f"3. 文件大小：{os.path.getsize(template_path)} 字节")
        print(f"4. 文件权限：可读？{os.access(template_path, os.R_OK)}")
        
        # 尝试用 openpyxl 直接读取
        try:
            from openpyxl import load_workbook
            wb = load_workbook(template_path)
            print(f"✅ openpyxl 能正常读取")
            print(f"工作表：{wb.sheetnames}")
            ws = wb.active
            print(f"表头：{[cell.value for cell in ws[1]]}")
        except Exception as e:
            print(f"❌ openpyxl 读取失败：{e}")
            
            # 用 pandas 再试试
            try:
                import pandas as pd
                df = pd.read_excel(template_path)
                print(f"✅ pandas 能读取")
                print(f"列名：{df.columns.tolist()}")
            except Exception as e2:
                print(f"❌ pandas 也读取失败：{e2}")
    print(f"5. 数据列表：{data_list}")
    print(f"6. 输出路径：{output_path}")
    print(f"7. 输出目录是否存在：{os.path.exists(os.path.dirname(output_path))}")
    print("="*60 + "\n")
    # ===== 调试信息结束 =====
    
    try:
        # 在调用 fill_excel_with_data 之前
        print("\n" + "="*60)
        print("🔍 传递给乙的数据检查")
        print("="*60)
        print(f"data_list 类型：{type(data_list)}")
        print(f"data_list 长度：{len(data_list)}")
        if len(data_list) > 0:
            print(f"第一行数据：{data_list[0]}")
            if len(data_list) > 1:
                print(f"第二行数据：{data_list[1]}")
        print("="*60 + "\n")
        excel_handler.fill_excel_with_data(template_path, data_list, output_path)
        logger.info(f"✅ Excel生成成功：{output_path}")
    except Exception as e:
        import traceback
        error_msg = str(e)
        error_detail = traceback.format_exc()
        logger.error(f"生成表格失败：{error_msg}")
        logger.error(f"详细错误：{error_detail}")
        return f"生成表格失败：{error_msg}", 500
    # 返回文件下载
    return send_file(output_path, as_attachment=True, download_name='result.xlsx')

if __name__ == '__main__':
    app.run(debug=True)