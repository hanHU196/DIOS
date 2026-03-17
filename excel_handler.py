import os
import sys

import pandas as pd
from docx import Document
from flask import Flask, request, send_file
from openpyxl import load_workbook
from pymongo import MongoClient


# Excel 工具函数
# 解析 Excel 模板，返回字段列表
def parse_excel_template(template_path):
    df = pd.read_excel(template_path, nrows=0)
    return df.columns.tolist()

# 填充 Excel 模板
def fill_excel_with_data(template_path, data, output_path):
    """填充 Excel 模板（支持多行数据）"""
    from openpyxl import load_workbook
    
    wb = load_workbook(template_path)
    ws = wb.active
    
    # 获取表头（第一行）
    headers = [cell.value for cell in ws[1]]
    print(f"📋 表头：{headers}")
    print(f"📊 数据行数：{len(data)}")
    
    # 如果只有表头，从第二行开始添加数据
    for row_idx, record in enumerate(data, start=2):
        print(f"  写入第 {row_idx} 行：{record}")
        for col_idx, header in enumerate(headers, start=1):
            value = record.get(header, '')
            ws.cell(row=row_idx, column=col_idx, value=value)
    
    wb.save(output_path)
    print(f"✅ Excel 文件已生成：{output_path}，共 {len(data)} 行数据")
# Word 工具函数
# 替换段落中的占位符
def replace_placeholders_in_paragraph(paragraph, data):
    """遍历 runs 替换占位符，保留样式"""
    for run in paragraph.runs:
        text = run.text
        for key, value in data.items():
            placeholder = '{{' + key + '}}'
            if placeholder in text:
                text = text.replace(placeholder, str(value))
        run.text = text

# 替换表格中的占位符
def replace_placeholders_in_table(table, data):
    """替换表格单元格中的占位符"""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                replace_placeholders_in_paragraph(paragraph, data)

# 填充 Word 模板
def fill_word_with_data(template_path, data_records, output_dir, filename_prefix="output"):
    os.makedirs(output_dir, exist_ok=True)
    for idx, record in enumerate(data_records):
        doc = Document(template_path)
        # 替换段落
        for paragraph in doc.paragraphs:
            replace_placeholders_in_paragraph(paragraph, record)
        # 替换表格
        for table in doc.tables:
            replace_placeholders_in_table(table, record)
        # 替换页眉页脚
        for section in doc.sections:
            for paragraph in section.header.paragraphs:
                replace_placeholders_in_paragraph(paragraph, record)
            for paragraph in section.footer.paragraphs:
                replace_placeholders_in_paragraph(paragraph, record)
        # 保存文件
        filename = f"{filename_prefix}_{idx+1}.docx"
        output_path = os.path.join(output_dir, filename)
        doc.save(output_path)
        print(f"Word 文件已生成：{output_path}")

# MongoDB 操作类
class MongoDBHandler:
    def __init__(self, uri="mongodb://localhost:27017/", db_name="testdb"):
        self.client = MongoClient(uri)
        self.db = self.client[db_name]

    def insert_data(self, collection_name, data):
        coll = self.db[collection_name]
        if isinstance(data, list):
            return coll.insert_many(data).inserted_ids
        else:
            return coll.insert_one(data).inserted_id

    def query_data(self, collection_name, query=None):
        query = query or {}
        cursor = self.db[collection_name].find(query)
        results = []
        for doc in cursor:
            doc.pop('_id', None)
            results.append(doc)
        return results

    def clear_collection(self, collection_name):
        self.db[collection_name].delete_many({})

#  测试
def run_test():
    """运行测试：Excel 和 Word 填充"""
    # 确保模板文件夹存在
    if not os.path.exists("templates"):
        os.makedirs("templates1")
        print("请将 Excel 模板放入 templates/template1.xlsx，Word 模板放入 templates/template2.docx")
        return

    # 测试数据
    test_data = [
        {"区域": "A", "村民委员会": "B", "总户数（户）": 12345, "总人数（人）": 56789,"粮食面积（亩）":159348,"粮食总产量（吨）":954625},
        {"区域": "C", "村民委员会": "D", "总户数（户）": 55555, "总人数（人）": 44444,"粮食面积（亩）":333333,"粮食总产量（吨）":111111},
    ]

    # 创建输出目录
    os.makedirs("output", exist_ok=True)

    excel_template = "templates/template1.xlsx"
    if os.path.exists(excel_template):
        fields = parse_excel_template(excel_template)
        print("Excel 模板字段：", fields)
        fill_excel_with_data(excel_template, test_data, "output/test_fill.xlsx")
    else:
        print("跳过 Excel 测试：模板文件不存在")

    word_template = "templates/template2.docx"
    if os.path.exists(word_template):
        fill_word_with_data(word_template, test_data, "output/word_output", filename_prefix="合同")
    else:
        print("跳过 Word 测试：模板文件不存在")

    #这里是MongDB测试
    db = MongoDBHandler()
    db.clear_collection("contracts")
    sample = [
        {"区域": "A", "村民委员会": "B", "总户数（户）": 12345, "总人数（人）": 56789, "粮食面积（亩）": 159348,
          "粮食总产量（吨）": 954625},
        {"区域": "C", "村民委员会": "D", "总户数（户）": 55555, "总人数（人）": 44444, "粮食面积（亩）": 333333,
         "粮食总产量（吨）": 111111},
    ]
    db.insert_data("contracts", sample)
    data_from_db = db.query_data("contracts", {})
    if os.path.exists(excel_template):
        fill_excel_with_data(excel_template, data_from_db, "output/db_fill.xlsx")
    if os.path.exists(word_template):
        fill_word_with_data(word_template, data_from_db, "output/word_db_output", filename_prefix="db_contract")
    print("测试完成！")

# Flask API
def create_app():
    app = Flask(__name__)

    @app.route('/fill/excel', methods=['POST'])
    def fill_excel():
        """Excel 填充 API"""
        f = request.files['template']
        data = request.get_json()
        temp_path = "temp_template.xlsx"
        f.save(temp_path)
        out_path = "output_excel.xlsx"
        fill_excel_with_data(temp_path, data, out_path)
        return send_file(out_path, as_attachment=True)

    @app.route('/fill/word', methods=['POST'])
    def fill_word():
        f = request.files['template']
        data = request.get_json()
        temp_path = "temp_template.docx"
        f.save(temp_path)
        # 由于可能生成多个文件，简单起见只处理第一条记录
        if data:
            first_record = [data[0]] if isinstance(data, list) else [data]
            fill_word_with_data(temp_path, first_record, "temp_word_output", "api_output")
            out_file = os.path.join("temp_word_output", "api_output_1.docx")
            return send_file(out_file, as_attachment=True)
        else:
            return "No data provided", 400

    return app

# 命令行入口
if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--api":
        app = create_app()
        print("启动 API 服务，访问 http://127.0.0.1:5000/fill/excel 或 /fill/word")
        app.run(debug=True)
    else:
        run_test()