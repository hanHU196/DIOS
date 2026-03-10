#文档读取

import os          # 操作系统相关功能，比如检查文件是否存在、获取文件后缀
import pdfplumber  # 读取PDF的库
from docx import Document  # 读取Word的库
import pandas as pd        # 处理Excel的库（最强大的数据处理工具）
import markdown    # 处理Markdown的库
import re          # 正则表达式，用来处理文字（暂时用不到，但留着备用）
from pathlib import Path  # 处理文件路径的现代方式

# ===== 第二部分：定义类 =====
class DocumentReader:
    """
    万能文档读取器
    把类的功能想象成一个"工具箱"，里面有很多工具（函数）
    用的时候：reader = DocumentReader()  # 创建工具箱
            content = reader.read("文件")  # 用工具箱里的read工具
    """
    
    def __init__(self):
        """
        初始化函数：创建工具箱时自动执行
        类似于"开机自检"，检查支持哪些格式
        """
        # 创建一个字典：文件后缀 -> 对应的读取函数
        # 这样以后加新格式很方便，只要在这里加一行
        self.supported_formats = {
            '.txt': self._read_txt,      # txt文件用_read_txt函数读
            '.md': self._read_md,        # md文件用_read_md函数读
            '.xlsx': self._read_excel,    # Excel文件用_read_excel读
            '.xls': self._read_excel,     # 老版Excel也一样
            '.docx': self._read_docx,     # Word文件用_read_docx读
            '.pdf': self._read_pdf,       # PDF文件用_read_pdf读
            '.csv': self._read_csv        # CSV文件用_read_csv读
        }
        
        # 打印初始化成功信息（调试用）
        print("✅ 文档读取器初始化成功！支持格式：", list(self.supported_formats.keys()))
    
    # ===== 第三部分：统一入口函数 =====
    def read(self, file_path):
        """
        这是最重要的函数！对外统一接口
        别人调用时只需要：reader.read("文件名")
        不用管是什么格式，内部自动处理
        
        参数：file_path 文件路径（字符串）
        返回：文件内容（字符串）
        """
        # 1. 检查文件是否存在
        # os.path.exists 是Python内置函数，判断文件/文件夹是否存在
        if not os.path.exists(file_path):
            return f"❌ 文件不存在：{file_path}"
        
        # 2. 获取文件后缀
        # os.path.splitext 把文件名拆成(名字, 后缀)
        # 例如 "合同.pdf" -> ("合同", ".pdf")
        ext = os.path.splitext(file_path)[1].lower()  # .lower()转成小写
        
        # 获取文件名（不含路径），用于显示
        file_name = os.path.basename(file_path)
        
        print(f"📖 正在读取：{file_name}")
        
        # 3. 根据后缀选择读取方法
        # 如果后缀在支持的格式字典里
        if ext in self.supported_formats:
            try:
                # 从字典里取出对应的函数，然后调用它
                # 比如 .pdf 对应 self._read_pdf，所以这里就是调用 self._read_pdf(file_path)
                content = self.supported_formats[ext](file_path)
                print(f"✅ 读取成功！共 {len(content)} 字符")
                return content
            except Exception as e:
                # 如果出错，捕获异常并返回错误信息
                error_msg = f"❌ 读取失败：{str(e)}"
                print(error_msg)
                return error_msg
        else:
            # 不支持的文件格式
            error_msg = f"❌ 不支持的文件格式：{ext}，支持格式：{list(self.supported_formats.keys())}"
            print(error_msg)
            return error_msg
    
    # ===== 第四部分：各种格式的具体读取函数 =====
    
    def _read_txt(self, file_path):
        """
        读取TXT文件
        函数名前加_表示"内部使用"，不建议外部直接调用
        
        难点：中文编码问题
        不同电脑保存txt用的编码可能不同：utf-8, gbk, gb2312...
        """
        # 准备一个列表，放着可能用到的编码
        encodings = ['utf-8', 'gbk', 'gb2312', 'utf-16', 'ansi']
        
        # 逐个尝试每种编码
        for encoding in encodings:
            try:
                # 尝试用这种编码打开文件
                with open(file_path, 'r', encoding=encoding) as f:
                    content = f.read()  # 读取全部内容
                print(f"  使用编码：{encoding}")
                return content  # 成功就返回
            except UnicodeDecodeError:
                # 如果解码失败，继续尝试下一种编码
                continue
            except Exception as e:
                # 其他错误也继续尝试
                continue
        
        # 如果所有编码都失败，用二进制模式读取
        # 'rb' 表示二进制模式，然后用ignore忽略无法解码的字符
        with open(file_path, 'rb') as f:
            content = f.read().decode('utf-8', errors='ignore')
        print("  使用二进制模式")
        return content
    
    def _read_md(self, file_path):
        """
        读取Markdown文件
        MD其实就是带特殊格式的txt，所以先按txt读
        """
        # 先按txt读
        content = self._read_txt(file_path)
        
        # 可选：用markdown库转成HTML，保留结构
        try:
            # markdown.markdown 把MD转成HTML
            html = markdown.markdown(content)
            # 返回原始内容和部分HTML（方便调试）
            return f"【Markdown文档】\n{content}\n\n【HTML转换】\n{html[:500]}..."
        except:
            # 如果转换失败，至少返回原始内容
            return f"【Markdown文档】\n{content}"
    
    def _read_excel(self, file_path):
        """
        读取Excel文件
        用pandas库，它是最强大的数据处理工具
        """
        try:
            # 1. 用ExcelFile读取整个Excel文件（可以获取所有sheet）
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names  # 获取所有sheet的名字
            
            result = []  # 用来存放输出的每一行
            result.append(f"【Excel文件】共 {len(sheet_names)} 个工作表")
            
            # 遍历每个sheet
            for sheet_name in sheet_names:
                # 读取这个sheet的数据，存到DataFrame（pandas的核心数据结构，类似表格）
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                rows, cols = df.shape  # df.shape 返回(行数, 列数)
                
                result.append(f"\n📊 工作表：{sheet_name}")
                result.append(f"  行数：{rows}，列数：{cols}")
                
                # 显示列名（只显示前10个，防止太长）
                # df.columns 获取列名
                result.append(f"  列名：{', '.join(str(c) for c in df.columns[:10])}")
                
                # 显示前3行数据
                if rows > 0:
                    result.append("  数据预览（前3行）：")
                    for i in range(min(3, rows)):
                        # df.iloc[i] 获取第i行，.to_dict()转成字典
                        row_data = df.iloc[i].to_dict()
                        row_str = f"    第{i+1}行："
                        # 只显示前5列，防止太长
                        for k, v in list(row_data.items())[:5]:
                            row_str += f"{k}={v} "
                        result.append(row_str)
            
            # 用换行符把所有内容连起来
            return '\n'.join(result)
            
        except Exception as e:
            return f"Excel读取失败：{str(e)}"
    
    def _read_csv(self, file_path):
        """
        读取CSV文件
        CSV就是逗号分隔的文本表格
        """
        try:
            # pandas直接读csv
            df = pd.read_csv(file_path)
            rows, cols = df.shape
            
            result = []
            result.append(f"【CSV文件】{rows}行×{cols}列")
            result.append(f"列名：{', '.join(str(c) for c in df.columns[:10])}")
            
            # 显示前3行
            result.append("数据预览（前3行）：")
            # df.head(3) 取前3行，to_string()转成字符串
            result.append(df.head(3).to_string())
            
            return '\n'.join(result)
            
        except Exception as e:
            return f"CSV读取失败：{str(e)}"
    
    def _read_docx(self, file_path):
        """
        读取Word文档
        Word的结构：段落(paragraphs) + 表格(tables) + 可能还有图片
        """
        try:
            # 打开Word文档
            doc = Document(file_path)
            
            # 1. 读取所有段落
            paragraphs = []
            for para in doc.paragraphs:
                if para.text.strip():  # strip()去掉空格，如果是空行就跳过
                    paragraphs.append(para.text)
            
            # 2. 读取所有表格
            tables = []
            for table in doc.tables:
                table_data = []
                for row in table.rows:
                    # 把一行中每个单元格的内容用 | 连接起来
                    row_text = ' | '.join([cell.text for cell in row.cells])
                    if row_text.strip():
                        table_data.append(row_text)
                if table_data:
                    tables.append('\n'.join(table_data))
            
            # 3. 检测是否有图片
            # Word文档里，图片是作为一种"关系"存在的
            has_images = False
            for rel in doc.part.rels.values():
                if "image" in rel.reltype:  # 如果关系类型包含image
                    has_images = True
                    break
            
            # 4. 组装结果
            result = []
            result.append(f"【Word文档】{file_path}")
            result.append(f"段落数：{len(paragraphs)}")
            result.append(f"表格数：{len(tables)}")
            result.append(f"包含图片：{'是' if has_images else '否'}")
            
            if has_images:
                result.append("⚠️ 文档包含图片，当前只能读取文字")
            
            # 显示部分段落（最多20段）
            if paragraphs:
                result.append("\n【文字内容】")
                for i, para in enumerate(paragraphs[:20]):
                    # 只显示前100字，后面加...
                    preview = para[:100] + ('...' if len(para) > 100 else '')
                    result.append(f"{i+1}. {preview}")
                if len(paragraphs) > 20:
                    result.append(f"... 还有{len(paragraphs)-20}段")
            
            # 显示部分表格（最多3个）
            if tables:
                result.append("\n【表格内容】")
                for i, table in enumerate(tables[:3]):
                    result.append(f"表格{i+1}：\n{table[:500]}")
            
            return '\n'.join(result)
            
        except Exception as e:
            return f"Word读取失败：{str(e)}"
    
    def _read_pdf(self, file_path):
        """
        读取PDF文件
        PDF的每页可能包含文字和表格
        """
        try:
            text_content = []
            
            # 打开PDF
            with pdfplumber.open(file_path) as pdf:
                total_pages = len(pdf.pages)  # 总页数
                result = [f"【PDF文件】共 {total_pages} 页"]
                
                # 遍历每一页
                for i, page in enumerate(pdf.pages):
                    # 提取文字
                    page_text = page.extract_text() or ""  # 如果没文字就返回空字符串
                    
                    # 提取表格
                    tables = page.extract_tables()
                    
                    result.append(f"\n--- 第 {i+1} 页 ---")
                    
                    # 显示文字预览
                    if page_text.strip():
                        preview = page_text[:200] + ('...' if len(page_text) > 200 else '')
                        result.append(f"文字：{preview}")
                    else:
                        result.append("文字：无")
                    
                    # 显示表格信息
                    if tables:
                        result.append(f"表格：{len(tables)}个")
                        
                        # 显示第一个表格预览
                        if tables[0]:
                            result.append("表格预览：")
                            for row in tables[0][:3]:  # 只显示前3行
                                # 把每个单元格转成字符串，用 | 连接
                                row_str = ' | '.join([str(cell) if cell else '' for cell in row])
                                result.append(f"  {row_str}")
                    
                    # 只处理前5页，避免输出太长
                    if i >= 4:
                        result.append(f"\n... 还有 {total_pages-5} 页未显示")
                        break
            
            return '\n'.join(result)
            
        except Exception as e:
            return f"PDF读取失败：{str(e)}"
    
    def save_to_file(self, content, output_path):
        """
        保存读取的内容到文件（调试用）
        可以把读出来的内容存成txt，方便查看
        """
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"✅ 内容已保存到：{output_path}")
            return True
        except Exception as e:
            print(f"❌ 保存失败：{e}")
            return False


# ===== 第五部分：测试代码 =====
# 当直接运行这个文件时（python document_reader.py），会执行下面的代码
if __name__ == "__main__":
    print("="*50)
    print("文档读取模块测试")
    print("="*50)
    
    # 创建读取器
    reader = DocumentReader()
    
    # 创建测试文件夹
    test_dir = "test_files"
    os.makedirs(test_dir, exist_ok=True)  # 如果不存在就创建
    
    # 1. 测试TXT
    print("\n1. 测试TXT文件")
    txt_path = os.path.join(test_dir, "test.txt")
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write("""这是一个测试文本文件。
甲方：XX科技有限公司
乙方：YY大学
金额：10000元
签订日期：2024年3月1日
备注：这是一份采购合同。""")
    content = reader.read(txt_path)
    print(content[:200])
    
    # 2. 测试MD
    print("\n2. 测试MD文件")
    md_path = os.path.join(test_dir, "test.md")
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write("""# 项目文档
## 第一章
这是一个测试段落。
- 列表项1
- 列表项2
## 第二章
结束。""")
    content = reader.read(md_path)
    print(content[:200])
    
    # 3. 测试Excel
    print("\n3. 测试Excel文件")
    excel_path = os.path.join(test_dir, "test.xlsx")
    df = pd.DataFrame({
        '姓名': ['张三', '李四', '王五'],
        '年龄': [20, 21, 22],
        '部门': ['技术部', '市场部', '财务部'],
        '工资': [8000, 9000, 8500]
    })
    df.to_excel(excel_path, index=False, sheet_name='员工信息')
    content = reader.read(excel_path)
    print(content[:300])
    
    # 4. 测试Word（需要有word文件，这里跳过）
    print("\n4. Word测试（需要有word文件）")
    word_path = os.path.join(test_dir, "test.docx")
    if os.path.exists(word_path):
        content = reader.read(word_path)
        print(content[:300])
    else:
        print(f"请手动放一个Word文件到：{word_path}")
    
    # 5. 测试PDF（需要有pdf文件）
    print("\n5. PDF测试（需要有pdf文件）")
    pdf_path = os.path.join(test_dir, "test.pdf")
    if os.path.exists(pdf_path):
        content = reader.read(pdf_path)
        print(content[:300])
    else:
        print(f"请手动放一个PDF文件到：{pdf_path}")
    
    print("\n" + "="*50)
    print("测试完成！")
    print("="*50)