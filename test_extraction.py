"""测试信息提取功能"""

from main import InfoExtractor

# 创建提取器
extractor = InfoExtractor()

# 测试文件名解析
print("=" * 60)
print("测试文件名解析")
print("=" * 60)

test_filenames = [
    "【前端工程师_成都 9-14K】刘存麟 6年.pdf",
    "【前端开发工程师_成都 8-12K】李志华 5年.pdf",
    "【前端开发工程师_成都 8-12K】陈雄 5年.pdf",
]

for filename in test_filenames:
    info = extractor.parse_filename(filename)
    print(f"\n文件名: {filename}")
    print(f"  姓名: {info.get('name')}")
    print(f"  岗位: {info.get('position')}")
    print(f"  地区: {info.get('location')}")
    print(f"  工资: {info.get('salary')}")

# 测试电话提取
print("\n" + "=" * 60)
print("测试电话提取")
print("=" * 60)

test_phones = [
    "电话：13812345678",
    "(+86) 159-2842-3292",
    "手机：138 1234 5678",
    "联系方式：15912345678",
]

for text in test_phones:
    phone = extractor.extract_phone(text)
    print(f"\n文本: {text}")
    print(f"  提取结果: {phone}")

# 测试邮箱提取
print("\n" + "=" * 60)
print("测试邮箱提取")
print("=" * 60)

test_emails = [
    "邮箱：gsm2832954437@163.com",
    "Email: test123@qq.com",
    "联系邮箱：5@163.com gsm2832954437@163.com",
]

for text in test_emails:
    email = extractor.extract_email(text)
    print(f"\n文本: {text}")
    print(f"  提取结果: {email}")

# 测试地区提取
print("\n" + "=" * 60)
print("测试地区提取")
print("=" * 60)

test_locations = [
    "期望城市：成都",
    "工作地点：四川 成都 籍贯：四川 德阳 中江县",
    "期望城市：成都 c 0",
]

for text in test_locations:
    location = extractor.extract_location(text)
    print(f"\n文本: {text}")
    print(f"  提取结果: {location}")

# 测试工资提取
print("\n" + "=" * 60)
print("测试工资提取")
print("=" * 60)

test_salaries = [
    "期望薪资：9-14K | 期望城市：成都 c 0",
    "薪资要求：12-915k | 到岗时间：周内到岗",
    "期望工资：9k~12k 5 c",
    "期望薪资：8-12K",
]

for text in test_salaries:
    salary = extractor.extract_salary(text)
    print(f"\n文本: {text}")
    print(f"  提取结果: {salary}")

print("\n" + "=" * 60)
print("测试完成")
print("=" * 60)
