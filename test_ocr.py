from paddleocr import PaddleOCR
import os

print("="*60)
print("PaddleOCR 3.x 测试脚本")
print("="*60)

# 跳过模型源检查
os.environ['PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK'] = 'True'

try:
    # 3.x 版本的正确初始化方式（最简参数）
    print("\n1. 初始化OCR...")
    ocr = PaddleOCR(
        lang='ch',  # 中文识别
        use_textline_orientation=True  # 新版的方向检测参数
    )
    print("✅ OCR初始化成功！")

    # 查看可用参数
    print("\n2. 查看可用参数:")
    import inspect
    sig = inspect.signature(ocr.__class__.__init__)
    for param in sig.parameters:
        print(f"   - {param}")

    print("\n3. 测试识别（需要一张图片）...")
    # 如果有测试图片，可以取消下面的注释
    # result = ocr.ocr('test.jpg')
    # if result and result[0]:
    #     for line in result[0]:
    #         print(f"   识别文字: {line[1][0]}")

except Exception as e:
    print(f"❌ 错误: {e}")
    import traceback
    traceback.print_exc()