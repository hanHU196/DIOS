"""
A23 智能填表系统 - Flask 主程序
适配现有项目结构：document_reader.py, excel_handler.py, ai_module.py, config.py
"""
import os
import time
import re
import logging
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ---------- 尝试导入队友模块，失败则使用模拟函数 ----------
# 文档读取模块
try:
    from document_reader import read_documents as real_read_documents
    logger.info("成功导入 document_reader.read_documents")
except ImportError:
    logger.warning("导入 document_reader 失败，使用模拟函数")
    def real_read_documents(file_paths):
        """模拟读取文档，仅支持 .txt"""
        all_text = ""
        for path in file_paths:
            ext = os.path.splitext(path)[1].lower()
            if ext == '.txt':
                try:
                    with open(path, 'r', encoding='utf-8') as f:
                        all_text += f.read() + "\n"
                except Exception as e:
                    logger.error(f"读取文件 {path} 失败: {e}")
            else:
                all_text += f"[暂不支持读取 {ext} 文件，演示继续]\n"
        return all_text

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

# 配置上传和输出文件夹
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    instruction = request.form.get('instruction', '').strip()
    fields_input = request.form.get('fields', '').strip()
    uploaded_files = request.files.getlist('files')

    if not uploaded_files or uploaded_files[0].filename == '':
        return jsonify({'error': '请至少上传一个文件'}), 400

    # 保存文件
    saved_paths = []
    for file in uploaded_files:
        if file.filename == '':
            continue
        filename = secure_filename(file.filename)
        path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(path)
        saved_paths.append(path)

    # 1. 读取文档（调用真实或模拟函数）
    all_text = real_read_documents(saved_paths)
    if not all_text.strip():
        return jsonify({'error': '无法读取文档内容，请确保上传了支持的格式'}), 400

    # 2. 解析指令意图
    intent = parse_instruction(instruction)
    logger.info(f"解析意图: {intent}")

    if intent != 'fill_form':
        return jsonify({'message': f'当前只演示填表功能，您的意图是：{intent}'}), 200

    # 3. 确定要提取的字段
    targets = []
    if fields_input:
        targets = [f.strip() for f in fields_input.split(',') if f.strip()]
    else:
        match = re.search(r'提取(.*)', instruction)
        if match:
            fields_str = match.group(1)
            targets = re.split(r'[和、,，]', fields_str)
            targets = [t.strip() for t in targets if t.strip()]
        else:
            targets = ['姓名', '金额', '日期']

    if not targets:
        return jsonify({'error': '未指定要提取的字段'}), 400

    # 4. 调用AI提取
    logger.info(f"开始提取字段: {targets}")
    extracted = extract_entities(all_text, targets)
    logger.info(f"提取结果: {extracted}")

    # 5. 生成Excel（调用真实或模拟函数）
    output_file = real_fill_template(extracted, output_filename='filled_form.xlsx')

    # 6. 返回文件下载
    return send_file(output_file, as_attachment=True, download_name='filled_form.xlsx')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)