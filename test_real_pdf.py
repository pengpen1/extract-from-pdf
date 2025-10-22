"""测试真实PDF文件提取"""

from pathlib import Path
from main import PDFExtractor, InfoExtractor

# 创建提取器
pdf_extractor = PDFExtractor()
info_extractor = InfoExtractor()

# 测试文件
test_file = Path("datas/饶璐的简历.pdf")

if test_file.exists():
    print("=" * 60)
    print(f"测试文件: {test_file.name}")
    print("=" * 60)
    
    try:
        # 提取PDF文本
        text = pdf_extractor.extract_text(test_file)
        print(f"\n✓ 文本提取成功，长度: {len(text)} 字符")
        
        # 显示前500字符
        print(f"\n--- 文本预览（前500字符）---")
        print(text[:500])
        print("--- 文本预览结束 ---\n")
        
        # 从文件名中解析信息
        filename_info = info_extractor.parse_filename(test_file.name)
        print("\n从文件名解析的信息:")
        print(f"  姓名: {filename_info.get('name')}")
        print(f"  岗位: {filename_info.get('position')}")
        print(f"  地区: {filename_info.get('location')}")
        print(f"  工资: {filename_info.get('salary')}")
        
        # 从PDF内容提取信息
        print("\n从PDF内容提取的信息:")
        name = info_extractor.extract_name(text)
        gender = info_extractor.extract_gender(text)
        age = info_extractor.extract_age(text)
        date = info_extractor.extract_date(text)
        phone = info_extractor.extract_phone(text)
        position = info_extractor.extract_position(text)
        location = info_extractor.extract_location(text)
        salary = info_extractor.extract_salary(text)
        email = info_extractor.extract_email(text)
        
        print(f"  姓名: {name}")
        print(f"  性别: {gender}")
        print(f"  年龄: {age}")
        print(f"  时间: {date}")
        print(f"  电话: {phone}")
        print(f"  岗位: {position}")
        print(f"  地区: {location}")
        print(f"  工资: {salary}")
        print(f"  邮箱: {email}")
        
        # 合并结果（优先PDF内容，文件名作为补充）
        print("\n最终结果（合并）:")
        final_name = name or filename_info.get('name') or ""
        final_position = position or filename_info.get('position') or ""
        final_location = location or filename_info.get('location') or ""
        final_salary = salary or filename_info.get('salary') or ""
        
        print(f"  姓名: {final_name}")
        print(f"  性别: {gender or ''}")
        print(f"  年龄: {age or ''}")
        print(f"  时间: {date or ''}")
        print(f"  电话: {phone or ''}")
        print(f"  岗位: {final_position}")
        print(f"  地区: {final_location}")
        print(f"  工资: {final_salary}")
        print(f"  邮箱: {email or ''}")
        
    except Exception as e:
        print(f"✗ 提取失败: {e}")
else:
    print(f"测试文件不存在: {test_file}")
