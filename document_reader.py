# -*- coding: utf-8 -*-
import sys
import io
import os
import pdfplumber
from docx import Document
import pandas as pd
import markdown
import re
from pathlib import Path

# 强制标准输出使用UTF-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 定义类 
class DocumentReader:
    """
    万能文档读取器
    用的时候：reader = DocumentReader()  # 创建工具箱
            content = reader.read("文件")  # 用工具箱里的read工具
    """
    
    def __init__(self):
        """
        初始化函数：创建工具箱时自动执行
        """
        # 创建一个字典：文件后缀 -> 对应的读取函数
        self.supported_formats = {
            '.txt': self._read_txt,      # txt文件用_read_txt函数读
            '.md': self._read_md,        # md文件用_read_md函数读
            '.xlsx': self._read_excel,    # Excel文件用_read_excel读
            '.xls': self._read_excel,     # 老版Excel也一样
            '.docx': self._read_docx,     # Word文件用_read_docx读
            '.pdf': self._read_pdf,       # PDF文件用_read_pdf读
            '.csv': self._read_csv        # CSV文件用_read_csv读
        }
        
        # 打印初始化成功信息
        print("✅ 文档读取器初始化成功！支持格式：", list(self.supported_formats.keys()))
    
    def read_document(self, file_path):
        """兼容旧代码的别名方法"""
        return self.read(file_path)
    # 统一入口函数
    def read(self, file_path):
        """
        对外统一接口
        参数：file_path 文件路径（字符串）
        返回：文件内容（字符串）
        异常：如果文件不存在或读取失败，抛出异常
        """
        # 1. 检查文件是否存在
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在：{file_path}")
        
        # 2. 获取文件后缀
        ext = os.path.splitext(file_path)[1].lower()
        file_name = os.path.basename(file_path)
        
        print(f"📖 正在读取：{file_name}")
        
        # 3. 根据后缀选择读取方法
        if ext in self.supported_formats:
            try:
                content = self.supported_formats[ext](file_path)
                print(f"✅ 读取成功！共 {len(content)} 字符")
                return content
            except Exception as e:
                raise Exception(f"读取失败：{str(e)}")
        else:
            raise ValueError(f"不支持的文件格式：{ext}，支持格式：{list(self.supported_formats.keys())}")
    
    # 各种格式的具体读取函数
    def _read_txt(self, file_path):
        """读取TXT文件 - 完整内容"""
        encodings = ['utf-8', 'gbk', 'gb2312', 'utf-16', 'ansi']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    content = f.read()
                print(f"  使用编码：{encoding}")
                return f"【TXT文件】{os.path.basename(file_path)}\n{'='*60}\n{content}"
            except UnicodeDecodeError:
                continue
            except Exception:
                continue
        
        # 如果所有编码都失败，用二进制模式读取
        with open(file_path, 'rb') as f:
            content = f.read().decode('utf-8', errors='ignore')
        print("  使用二进制模式")
        return f"【TXT文件】{os.path.basename(file_path)}\n{'='*60}\n{content}"
    
    def _read_md(self, file_path):
        """读取Markdown文件 - 完整内容"""
        content = self._read_txt(file_path)
        
        try:
            html = markdown.markdown(content)
            return f"【Markdown文档】{os.path.basename(file_path)}\n{'='*60}\n【原始内容】\n{content}\n\n{'='*60}\n【HTML转换】\n{html}"
        except:
            return f"【Markdown文档】{os.path.basename(file_path)}\n{'='*60}\n{content}"
    
    def _read_excel(self, file_path):
        """读取Excel文件 - 完整内容"""
        try:
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            
            all_content = []
            all_content.append(f"【Excel文件】{os.path.basename(file_path)}")
            all_content.append(f"工作表数量：{len(sheet_names)}")
            all_content.append("="*60)
            
            for sheet_idx, sheet_name in enumerate(sheet_names):
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                rows, cols = df.shape
                
                all_content.append(f"\n📊 工作表 {sheet_idx+1}: {sheet_name}")
                all_content.append(f"行数: {rows}, 列数: {cols}")
                all_content.append("-"*40)
                
                # 设置pandas显示选项以确保完整输出
                with pd.option_context('display.max_rows', None, 'display.max_columns', None, 'display.width', None, 'display.max_colwidth', None):
                    # 将DataFrame转换为字符串
                    df_string = df.to_string(index=True, header=True)
                    all_content.append(df_string)
                
                all_content.append("-"*40)
            
            return '\n'.join(all_content)
            
        except Exception as e:
            return f"❌ Excel读取失败：{str(e)}"
    
    def _read_csv(self, file_path):
        """读取CSV文件 - 完整内容"""
        try:
            df = pd.read_csv(file_path)
            rows, cols = df.shape
            
            result = []
            result.append(f"【CSV文件】{os.path.basename(file_path)}")
            result.append(f"总行数：{rows}，总列数：{cols}")
            result.append("="*60)
            result.append("列名：")
            result.append(", ".join([str(col) for col in df.columns]))
            result.append("-"*40)
            result.append("【完整数据】")
            
            # 设置pandas显示选项以确保完整输出
            with pd.option_context('display.max_rows', None, 'display.max_columns', None, 'display.width', None, 'display.max_colwidth', None):
                result.append(df.to_string(index=True, header=True))
            
            return '\n'.join(result)
            
        except Exception as e:
            return f"❌ CSV读取失败：{str(e)}"
    
    def _read_docx(self, file_path):
        """读取Word文档 - 使用PaddleOCR识别图片"""
        try:
            from docx import Document
            import zipfile
            import os
            import tempfile
            from PIL import Image
            import io
            
            # 导入PaddleOCR
            try:
                from paddleocr import PaddleOCR
                # 初始化OCR（只在第一次调用时下载模型）
                print("  🔍 正在初始化PaddleOCR（首次运行会下载模型，请稍候）...")
                ocr = PaddleOCR(
                    use_angle_cls=True,      # 启用方向分类
                    lang='ch',                # 中文识别
                    show_log=False,           # 不显示详细日志
                    use_gpu=False             # 如果您有GPU可设为True
                )
                ocr_available = True
                print("  ✅ PaddleOCR初始化成功")
            except ImportError:
                print("  ⚠️ PaddleOCR未安装，请运行: pip install paddleocr")
                ocr_available = False
            except Exception as e:
                print(f"  ⚠️ PaddleOCR初始化失败: {e}")
                ocr_available = False
            
            # 打开Word文档
            doc = Document(file_path)
            
            result = []
            result.append(f"【Word文档】{os.path.basename(file_path)}")
            result.append("="*60)
            
            # 读取所有段落
            paragraphs = []
            for para in doc.paragraphs:
                if para.text.strip():
                    paragraphs.append(para.text)
            
            result.append(f"\n【段落内容】（共{len(paragraphs)}段）")
            result.append("-"*40)
            for i, para in enumerate(paragraphs, 1):
                # 如果段落太长，只显示前200字符作为预览（但完整内容会保存）
                if len(para) > 200:
                    result.append(f"{i}. {para[:200]}...（共{len(para)}字符）")
                else:
                    result.append(f"{i}. {para}")
            
            # 读取所有表格
            if doc.tables:
                result.append("\n【表格内容】")
                for table_idx, table in enumerate(doc.tables, 1):
                    result.append(f"\n表格 {table_idx}:")
                    result.append("-"*20)
                    for row in table.rows:
                        row_cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                        if row_cells:
                            result.append(" | ".join(row_cells))
            
            # 提取并识别图片
            result.append("\n【图片内容】")
            result.append("-"*40)
            
            # 方法1：通过rels提取图片
            image_count = 0
            recognized_texts = []
            
            for rel in doc.part.rels.values():
                if "image" in rel.reltype:
                    try:
                        image_count += 1
                        image_data = rel.target_part.blob
                        
                        # 使用PIL打开图片
                        image = Image.open(io.BytesIO(image_data))
                        
                        # 保存为临时文件（PaddleOCR需要文件路径）
                        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                            image.save(tmp_file, format='PNG')
                            tmp_path = tmp_file.name
                        
                        result.append(f"\n📷 图片 {image_count}:")
                        result.append(f"  图片大小: {image.size}")
                        
                        # OCR识别
                        if ocr_available:
                            print(f"    正在识别图片{image_count}...")
                            
                            # 调用PaddleOCR识别
                            ocr_result = ocr.ocr(tmp_path, cls=True)
                            
                            if ocr_result and ocr_result[0]:
                                result.append("  识别文字:")
                                # 遍历识别结果
                                for line in ocr_result[0]:
                                    # line格式: [ [[坐标], (文本, 置信度)] ]
                                    text = line[1][0]  # 识别出的文本
                                    confidence = line[1][1]  # 置信度
                                    if text.strip():
                                        result.append(f"    {text} (置信度: {confidence:.2f})")
                                        recognized_texts.append(text)
                            else:
                                result.append("  未识别到文字")
                        
                        # 清理临时文件
                        os.unlink(tmp_path)
                        
                    except Exception as e:
                        result.append(f"  图片处理失败: {str(e)}")
            
            # 方法2：通过解压方式提取所有图片（备用）
            if image_count == 0:
                result.append("\n尝试备用方法提取图片...")
                try:
                    with zipfile.ZipFile(file_path, 'r') as docx_zip:
                        for file_name in docx_zip.namelist():
                            if file_name.startswith('word/media/'):
                                image_count += 1
                                image_data = docx_zip.read(file_name)
                                
                                # 保存临时文件
                                ext = os.path.splitext(file_name)[1] or '.png'
                                with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp_file:
                                    tmp_file.write(image_data)
                                    tmp_path = tmp_file.name
                                
                                result.append(f"\n📷 图片 {image_count} ({os.path.basename(file_name)}):")
                                
                                if ocr_available:
                                    print(f"    正在识别图片{image_count}...")
                                    ocr_result = ocr.ocr(tmp_path, cls=True)
                                    
                                    if ocr_result and ocr_result[0]:
                                        result.append("  识别文字:")
                                        for line in ocr_result[0]:
                                            text = line[1][0]
                                            confidence = line[1][1]
                                            if text.strip():
                                                result.append(f"    {text} (置信度: {confidence:.2f})")
                                                recognized_texts.append(text)
                                    else:
                                        result.append("  未识别到文字")
                                
                                os.unlink(tmp_path)
                except Exception as e:
                    result.append(f"  解压提取失败: {str(e)}")
            
            if image_count == 0:
                result.append("  未找到图片")
            else:
                result.append(f"\n共识别到 {len(recognized_texts)} 条文字数据")
            
            return '\n'.join(result)
            
        except Exception as e:
            import traceback
            return f"❌ Word读取失败：{str(e)}\n{traceback.format_exc()}"
           
        
    def _read_pdf(self, file_path):
        """读取PDF文件 - 完整内容"""
        try:
            with pdfplumber.open(file_path) as pdf:
                total_pages = len(pdf.pages)
                result = []
                result.append(f"【PDF文件】{os.path.basename(file_path)}")
                result.append(f"总页数：{total_pages}")
                result.append("="*60)
                
                for page_num, page in enumerate(pdf.pages, 1):
                    result.append(f"\n📄 第 {page_num} 页")
                    result.append("-"*40)
                    
                    # 提取页面文字
                    page_text = page.extract_text() or ""
                    if page_text.strip():
                        result.append(page_text)
                    else:
                        result.append("（此页无文字内容）")
                    
                    # 提取表格
                    tables = page.extract_tables()
                    if tables:
                        result.append("\n表格：")
                        for table_idx, table in enumerate(tables, 1):
                            if table:
                                result.append(f"表格 {table_idx}:")
                                for row in table:
                                    row_text = ' | '.join([str(cell) if cell else '' for cell in row])
                                    if row_text.strip():
                                        result.append(f"  {row_text}")
                
                return '\n'.join(result)
            
        except Exception as e:
            return f"❌ PDF读取失败：{str(e)}"
    
    def save_to_file(self, content, output_path):
        """保存读取的内容到文件"""
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"✅ 内容已保存到：{output_path}")
            return True
        except Exception as e:
            print(f"❌ 保存失败：{e}")
            return False


# 在文件末尾添加这个模块级别的函数
def read_document(file_path):
    """模块级别的函数，供 app.py 直接调用"""
    reader = DocumentReader()
    return reader.read(file_path)

# ===== 交互式测试代码 =====
if __name__ == "__main__":
    print("="*60)
    print("📚 文档读取工具 - 完整输出版本")
    print("="*60)
    
    # 创建读取器
    reader = DocumentReader()
    
    while True:
        print("\n" + "-"*60)
        print("请选择操作：")
        print("1. 读取单个文件（输入路径）")
        print("2. 批量读取 test_files 文件夹")
        print("3. 退出")
        print("-"*60)
        
        choice = input("请输入选项 (1/2/3): ").strip()
        
        if choice == '1':
            # 读取单个文件
            file_path = input("\n请输入文件路径：").strip()
            # 自动去除可能存在的引号
            file_path = file_path.strip('"').strip("'")
            
            if not os.path.exists(file_path):
                print(f"❌ 文件不存在：{file_path}")
                continue
            
            try:
                # 读取文件
                content = reader.read(file_path)
                
                # 询问是否保存
                save = input("\n是否保存到文件？(y/n): ").strip().lower()
                if save == 'y':
                    output_path = f"output_{os.path.basename(file_path)}.txt"
                    reader.save_to_file(content, output_path)
                else:
                    # 如果用户不保存，则显示全部内容
                    print("\n" + "="*60)
                    print("📄 读取结果（完整内容）：")
                    print("="*60)
                    print(content)
                    
            except Exception as e:
                print(f"❌ 读取失败：{e}")
        
        elif choice == '2':
            # 批量读取 test_files 文件夹
            test_dir = "test_files"
            
            if not os.path.exists(test_dir):
                print(f"❌ 文件夹不存在：{test_dir}")
                os.makedirs(test_dir)
                print(f"已创建文件夹：{test_dir}，请把文件放进去再试")
                continue
            
            files = os.listdir(test_dir)
            if not files:
                print(f"❌ 文件夹为空：{test_dir}")
                continue
            
            print(f"\n📊 找到 {len(files)} 个文件：")
            for i, f in enumerate(files, 1):
                ext = os.path.splitext(f)[1].lower()
                status = "✅" if ext in reader.supported_formats else "❌"
                print(f"  {i}. {status} {f}")
            
            # 创建结果文件夹
            output_dir = "output_results"
            os.makedirs(output_dir, exist_ok=True)
            print("\n开始批量读取...")
            
            for filename in files:
                ext = os.path.splitext(filename)[1].lower()
                if ext not in reader.supported_formats:
                    print(f"⏭️ 跳过不支持格式：{filename}")
                    continue
                    
                file_path = os.path.join(test_dir, filename)
                print(f"\n📄 处理：{filename}")
                
                try:
                    content = reader.read(file_path)
                    # 保存结果
                    output_path = os.path.join(output_dir, f"{filename}.txt")
                    reader.save_to_file(content, output_path)
                except Exception as e:
                    print(f"❌ 处理失败：{e}")
            
            print(f"\n✅ 批量处理完成！结果保存在 {output_dir} 文件夹")
        
        elif choice == '3':
            print("\n👋 再见！")
            break
        
        else:
            print("❌ 无效选项，请重新输入")