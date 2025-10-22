"""
简历信息提取工具
从PDF简历中自动提取姓名、电话、邮箱等关键信息，并导出为Excel文件
"""

from dataclasses import dataclass
from typing import List, Optional
from pathlib import Path


# ==================== 数据模型 ====================


@dataclass
class ResumeInfo:
    """简历信息数据类"""

    index: int  # 序号
    name: str  # 姓名
    gender: str  # 性别
    age: str  # 年龄
    date: str  # 时间
    phone: str  # 电话
    position: str  # 岗位
    location: str  # 地区
    salary: str  # 工资
    email: str  # 邮箱
    filename: str  # 文件名


@dataclass
class ProcessingResult:
    """处理结果统计"""

    total: int  # 总文件数
    success: int  # 成功处理数
    failed: int  # 失败数
    failed_files: List[str]  # 失败文件列表


# ==================== 自定义异常 ====================


class PDFExtractionError(Exception):
    """PDF提取异常"""

    pass


class ExcelExportError(Exception):
    """Excel导出异常"""

    pass


# ==================== 文件扫描模块 ====================


class FileScanner:
    """文件扫描器，负责递归扫描指定目录中的所有PDF文件"""

    def __init__(self, data_folder: str):
        """初始化文件扫描器

        Args:
            data_folder: 数据文件夹路径
        """
        self.data_folder = Path(data_folder)

        # 如果文件夹不存在，自动创建
        if not self.data_folder.exists():
            self.data_folder.mkdir(parents=True, exist_ok=True)
            print(f"已自动创建数据文件夹: {self.data_folder.absolute()}")

    def scan_pdf_files(self) -> List[Path]:
        """递归扫描所有PDF文件

        使用pathlib.Path.rglob递归查找所有.pdf文件

        Returns:
            PDF文件路径列表
        """
        # 使用rglob递归查找所有PDF文件（不区分大小写）
        pdf_files = []

        # 查找.pdf文件
        pdf_files.extend(self.data_folder.rglob("*.pdf"))

        # 查找.PDF文件（大写扩展名）
        pdf_files.extend(self.data_folder.rglob("*.PDF"))

        # 去重并排序
        pdf_files = sorted(set(pdf_files))

        return pdf_files


# ==================== PDF文本提取模块 ====================


class PDFExtractor:
    """PDF文本提取器，负责从PDF文件中提取文本内容"""

    def extract_text(self, pdf_path: Path) -> str:
        """提取PDF文本内容

        优先使用pdfplumber提取（更准确），失败时回退到PyPDF2。
        只提取前3页内容以提高处理速度。

        Args:
            pdf_path: PDF文件路径

        Returns:
            提取的文本内容

        Raises:
            PDFExtractionError: 当所有提取方法都失败时抛出
        """
        text = ""

        # 方法1: 尝试使用pdfplumber提取（主要方法）
        try:
            text = self._extract_with_pdfplumber(pdf_path)
            if text.strip():  # 如果成功提取到非空文本
                return text
        except Exception as e:
            # pdfplumber失败，记录错误但继续尝试备用方法
            print(f"  pdfplumber提取失败: {str(e)}")

        # 方法2: 回退到PyPDF2（备用方法）
        try:
            text = self._extract_with_pypdf2(pdf_path)
            if text.strip():  # 如果成功提取到非空文本
                return text
        except Exception as e:
            # PyPDF2也失败，抛出异常
            raise PDFExtractionError(f"无法提取PDF文本: {str(e)}")

        # 如果两种方法都没有提取到有效文本
        if not text.strip():
            raise PDFExtractionError("PDF文件为空或无法提取文本内容")

        return text

    def _extract_with_pdfplumber(self, pdf_path: Path) -> str:
        """使用pdfplumber提取PDF文本

        Args:
            pdf_path: PDF文件路径

        Returns:
            提取的文本内容
        """
        import pdfplumber

        text_parts = []

        try:
            with pdfplumber.open(pdf_path) as pdf:
                # 只提取前3页
                max_pages = min(3, len(pdf.pages))

                for i in range(max_pages):
                    page = pdf.pages[i]
                    page_text = page.extract_text()

                    if page_text:
                        text_parts.append(page_text)
        except Exception as e:
            raise PDFExtractionError(f"pdfplumber提取失败: {str(e)}")

        return "\n".join(text_parts)

    def _extract_with_pypdf2(self, pdf_path: Path) -> str:
        """使用PyPDF2提取PDF文本（备用方法）

        Args:
            pdf_path: PDF文件路径

        Returns:
            提取的文本内容
        """
        import PyPDF2

        text_parts = []

        try:
            with open(pdf_path, "rb") as file:
                pdf_reader = PyPDF2.PdfReader(file)

                # 只提取前3页
                max_pages = min(3, len(pdf_reader.pages))

                for i in range(max_pages):
                    page = pdf_reader.pages[i]
                    page_text = page.extract_text()

                    if page_text:
                        text_parts.append(page_text)
        except Exception as e:
            raise PDFExtractionError(f"PyPDF2提取失败: {str(e)}")

        return "\n".join(text_parts)


# ==================== 信息提取模块 ====================


class InfoExtractor:
    """信息提取器，负责从文本中识别和提取姓名、电话、邮箱"""

    def __init__(self):
        """初始化信息提取器"""
        import re

        self.re = re

        # 常见标题词，用于过滤非姓名文本
        self.common_title_words = {
            "个人简历",
            "求职简历",
            "简历",
            "个人信息",
            "基本信息",
            "求职意向",
            "工作经历",
            "教育经历",
            "项目经验",
            "自我评价",
            "技能特长",
            "联系方式",
            "应聘岗位",
            "期望职位",
            "个人资料",
        }

    def extract_phone(self, text: str) -> Optional[str]:
        """提取手机号码

        使用正则表达式匹配中国大陆11位手机号码格式（1开头的11位数字）。
        当存在多个手机号时，返回第一个匹配的结果。

        Args:
            text: 简历文本

        Returns:
            手机号码字符串，未找到时返回None
        """
        if not text:
            return None

        # 正则表达式: 匹配1开头，第二位是3-9，后面9位数字
        # 使用\b确保是完整的11位数字，不是更长数字的一部分
        pattern = r"\b1[3-9]\d{9}\b"

        match = self.re.search(pattern, text)

        if match:
            return match.group(0)

        return None

    def extract_email(self, text: str) -> Optional[str]:
        """提取邮箱地址

        使用正则表达式匹配标准邮箱地址格式。
        当存在多个邮箱时，返回第一个匹配的结果。

        Args:
            text: 简历文本

        Returns:
            邮箱地址字符串，未找到时返回None
        """
        if not text:
            return None

        # 正则表达式: 匹配标准邮箱格式
        # 用户名部分: 字母、数字、点、下划线、百分号、加号、减号
        # @符号
        # 域名部分: 字母、数字、点、减号
        # 顶级域名: 至少2个字母
        pattern = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"

        match = self.re.search(pattern, text)

        if match:
            return match.group(0)

        return None

    def extract_name(self, text: str) -> Optional[str]:
        """提取姓名

        使用多策略组合方法提取姓名：
        1. 关键词模式匹配（优先）
        2. 位置启发式策略（回退）

        Args:
            text: 简历文本

        Returns:
            姓名字符串，未找到时返回None
        """
        if not text:
            return None

        # 策略1: 关键词模式匹配
        name = self._extract_name_by_keyword(text)
        if name:
            return name

        # 策略2: 位置启发式策略
        name = self._extract_name_by_position(text)
        if name:
            return name

        return None

    def _extract_name_by_keyword(self, text: str) -> Optional[str]:
        """通过关键词模式提取姓名

        搜索"姓名："、"姓 名："、"Name:"等关键词后的内容

        Args:
            text: 简历文本

        Returns:
            提取的姓名，未找到时返回None
        """
        # 定义关键词模式列表
        keyword_patterns = [
            r"姓\s*名\s*[：:]\s*([^\s\n]{2,4})",  # 姓名：、姓 名：
            r"姓\s*名\s*[：:]\s*([^\s\n]{2,4})",  # 姓名:（英文冒号）
            r"Name\s*[：:]\s*([^\s\n]{2,4})",  # Name:、Name：
            r"name\s*[：:]\s*([^\s\n]{2,4})",  # name:、name：
            r"名\s*字\s*[：:]\s*([^\s\n]{2,4})",  # 名字：
        ]

        for pattern in keyword_patterns:
            match = self.re.search(pattern, text, self.re.IGNORECASE)
            if match:
                candidate = match.group(1).strip()
                # 验证候选姓名
                if self._is_valid_name(candidate):
                    return candidate

        return None

    def _extract_name_by_position(self, text: str) -> Optional[str]:
        """通过位置启发式策略提取姓名

        在简历前200字符中查找2-4个连续中文字符作为候选姓名

        Args:
            text: 简历文本

        Returns:
            提取的姓名，未找到时返回None
        """
        # 只在前200字符中查找
        search_text = text[:200]

        # 查找2-4个连续中文字符
        # \u4e00-\u9fff 是中文字符的Unicode范围
        pattern = r"[\u4e00-\u9fff]{2,4}"

        matches = self.re.findall(pattern, search_text)

        # 遍历所有匹配，找到第一个有效的姓名
        for candidate in matches:
            if self._is_valid_name(candidate):
                return candidate

        return None

    def _is_valid_name(self, candidate: str) -> bool:
        """验证候选文本是否为有效姓名

        过滤规则：
        - 排除常见标题词
        - 排除包含数字的文本
        - 排除包含特殊符号的文本
        - 长度必须在2-4个字符之间

        Args:
            candidate: 候选姓名文本

        Returns:
            True表示有效姓名，False表示无效
        """
        if not candidate:
            return False

        # 去除首尾空白
        candidate = candidate.strip()

        # 检查长度（中文姓名通常2-4个字）
        if len(candidate) < 2 or len(candidate) > 4:
            return False

        # 排除常见标题词
        if candidate in self.common_title_words:
            return False

        # 排除包含数字的文本
        if self.re.search(r"\d", candidate):
            return False

        # 排除包含特殊符号的文本（允许中文字符）
        # 只允许纯中文字符
        if not self.re.match(r"^[\u4e00-\u9fff]+$", candidate):
            return False

        return True

    def extract_gender(self, text: str) -> Optional[str]:
        """提取性别

        搜索"性别："、"性 别："等关键词后的内容

        Args:
            text: 简历文本

        Returns:
            性别字符串（男/女），未找到时返回None
        """
        if not text:
            return None

        # 定义关键词模式列表
        patterns = [
            r"性\s*别\s*[：:]\s*(男|女)",
            r"Gender\s*[：:]\s*(男|女|Male|Female|male|female)",
        ]

        for pattern in patterns:
            match = self.re.search(pattern, text, self.re.IGNORECASE)
            if match:
                gender = match.group(1).strip()
                # 统一转换为中文
                if gender.lower() in ["male", "男"]:
                    return "男"
                elif gender.lower() in ["female", "女"]:
                    return "女"

        return None

    def extract_age(self, text: str) -> Optional[str]:
        """提取年龄

        搜索"年龄："、"Age:"等关键词后的数字，或者"XX岁"格式

        Args:
            text: 简历文本

        Returns:
            年龄字符串，未找到时返回None
        """
        if not text:
            return None

        # 定义关键词模式列表
        patterns = [
            r"年\s*龄\s*[：:]\s*(\d{1,2})",
            r"Age\s*[：:]\s*(\d{1,2})",
            r"(\d{2})岁",  # 匹配"26岁"这种格式
        ]

        for pattern in patterns:
            match = self.re.search(pattern, text, self.re.IGNORECASE)
            if match:
                age = match.group(1).strip()
                # 验证年龄范围合理性（18-70岁）
                if age.isdigit() and 18 <= int(age) <= 70:
                    return age

        return None

    def extract_date(self, text: str) -> Optional[str]:
        """提取日期/时间

        搜索简历中的日期信息（如更新时间、出生日期等）

        Args:
            text: 简历文本

        Returns:
            日期字符串，未找到时返回None
        """
        if not text:
            return None

        # 定义日期模式列表
        patterns = [
            r"更新时间\s*[：:]\s*(\d{4}[-/年]\d{1,2}[-/月]\d{1,2}日?)",
            r"出生日期\s*[：:]\s*(\d{4}[-/年]\d{1,2}[-/月]\d{1,2}日?)",
            r"(\d{4}[-/年]\d{1,2}[-/月]\d{1,2}日?)",
        ]

        for pattern in patterns:
            match = self.re.search(pattern, text)
            if match:
                return match.group(1).strip()

        return None

    def extract_position(self, text: str) -> Optional[str]:
        """提取应聘岗位

        搜索"应聘岗位："、"期望职位："等关键词后的内容
        如果内容包含城市名，只提取岗位部分

        Args:
            text: 简历文本

        Returns:
            岗位字符串，未找到时返回None
        """
        if not text:
            return None

        # 定义关键词模式列表
        patterns = [
            r"应聘岗位\s*[：:]\s*([^\n]+)",
            r"期望职位\s*[：:]\s*([^\n]+)",
            r"求职意向\s*[：:]\s*([^\n]+)",
            r"目标职位\s*[：:]\s*([^\n]+)",
            r"Position\s*[：:]\s*([^\n]+)",
        ]

        for pattern in patterns:
            match = self.re.search(pattern, text, self.re.IGNORECASE)
            if match:
                full_text = match.group(1).strip()
                
                # 如果包含常见分隔符（空格、|、/等），尝试分离岗位和地区
                # 例如："Java 成都" 或 "Java | 成都"
                separators = [r"\s+", r"\|", r"/", r"·"]
                
                for sep in separators:
                    parts = self.re.split(sep, full_text)
                    if len(parts) >= 2:
                        # 取第一部分作为岗位
                        position = parts[0].strip()
                        # 验证岗位不是城市名
                        if position and not self._is_city_name(position):
                            return position
                
                # 如果没有分隔符，返回整个文本（但要验证长度合理）
                if len(full_text) <= 15:
                    return full_text

        return None

    def extract_location(self, text: str) -> Optional[str]:
        """提取地区/城市

        搜索"工作地点："、"期望城市："等关键词后的内容
        或从"求职意向"中提取城市部分

        Args:
            text: 简历文本

        Returns:
            地区字符串，未找到时返回None
        """
        if not text:
            return None

        # 定义关键词模式列表
        patterns = [
            r"工作地点\s*[：:]\s*([^\n]{2,20})",
            r"期望城市\s*[：:]\s*([^\n]{2,20})",
            r"期望地点\s*[：:]\s*([^\n]{2,20})",
            r"所在地\s*[：:]\s*([^\n]{2,20})",
            r"Location\s*[：:]\s*([^\n]{2,20})",
        ]

        for pattern in patterns:
            match = self.re.search(pattern, text, self.re.IGNORECASE)
            if match:
                location = match.group(1).strip()
                # 清理可能的多余空白和换行
                location = self.re.sub(r"\s+", " ", location)
                return location

        # 如果没有找到，尝试从"求职意向"中提取城市
        intention_pattern = r"求职意向\s*[：:]\s*([^\n]+)"
        match = self.re.search(intention_pattern, text)
        if match:
            full_text = match.group(1).strip()
            
            # 尝试分离岗位和地区
            separators = [r"\s+", r"\|", r"/", r"·"]
            
            for sep in separators:
                parts = self.re.split(sep, full_text)
                if len(parts) >= 2:
                    # 取第二部分作为地区
                    location = parts[1].strip()
                    # 验证是否为城市名
                    if location and self._is_city_name(location):
                        return location

        return None

    def _is_city_name(self, text: str) -> bool:
        """判断文本是否为城市名

        Args:
            text: 待判断的文本

        Returns:
            True表示是城市名，False表示不是
        """
        # 常见城市名列表（可以根据需要扩展）
        cities = {
            "北京", "上海", "广州", "深圳", "成都", "重庆", "杭州", "武汉",
            "西安", "天津", "南京", "苏州", "长沙", "郑州", "沈阳", "青岛",
            "宁波", "东莞", "无锡", "佛山", "合肥", "昆明", "福州", "厦门",
            "哈尔滨", "济南", "温州", "长春", "石家庄", "常州", "泉州", "南宁",
            "贵阳", "南昌", "南通", "金华", "徐州", "太原", "嘉兴", "烟台",
            "惠州", "保定", "台州", "中山", "绍兴", "乌鲁木齐", "潍坊", "兰州",
        }
        
        return text in cities

    def extract_salary(self, text: str) -> Optional[str]:
        """提取期望薪资

        搜索"期望薪资："、"薪资要求："等关键词后的内容

        Args:
            text: 简历文本

        Returns:
            薪资字符串，未找到时返回None
        """
        if not text:
            return None

        # 定义关键词模式列表
        patterns = [
            r"期望薪资\s*[：:]\s*([^\n]{2,30})",
            r"薪资要求\s*[：:]\s*([^\n]{2,30})",
            r"期望工资\s*[：:]\s*([^\n]{2,30})",
            r"Salary\s*[：:]\s*([^\n]{2,30})",
        ]

        for pattern in patterns:
            match = self.re.search(pattern, text, self.re.IGNORECASE)
            if match:
                salary = match.group(1).strip()
                # 清理可能的多余空白和换行
                salary = self.re.sub(r"\s+", " ", salary)
                return salary

        return None


# ==================== 主程序入口 ====================

if __name__ == "__main__":
    print("简历信息提取工具 - 测试模块")
    print("=" * 50)

    # 测试FileScanner
    scanner = FileScanner("datas")
    pdf_files = scanner.scan_pdf_files()

    print(f"\n找到 {len(pdf_files)} 个PDF文件:")
    for pdf_file in pdf_files:
        print(f"  - {pdf_file}")

    if len(pdf_files) == 0:
        print("\n提示: 请将PDF文件放入 'datas' 文件夹中进行测试")
    else:
        # 测试PDFExtractor和InfoExtractor
        print("\n" + "=" * 50)
        print("测试PDF文本提取和信息提取模块")
        print("=" * 50)

        pdf_extractor = PDFExtractor()
        info_extractor = InfoExtractor()

        # 测试所有PDF文件
        for i, test_file in enumerate(pdf_files, 1):
            print(f"\n[{i}/{len(pdf_files)}] 正在处理: {test_file.name}")

            try:
                # 提取PDF文本
                text = pdf_extractor.extract_text(test_file)
                print(f"✓ 文本提取成功，长度: {len(text)} 字符")

                # 提取信息
                name = info_extractor.extract_name(text)
                gender = info_extractor.extract_gender(text)
                age = info_extractor.extract_age(text)
                date = info_extractor.extract_date(text)
                phone = info_extractor.extract_phone(text)
                position = info_extractor.extract_position(text)
                location = info_extractor.extract_location(text)
                salary = info_extractor.extract_salary(text)
                email = info_extractor.extract_email(text)

                print(f"  姓名: {name if name else ''}")
                print(f"  性别: {gender if gender else ''}")
                print(f"  年龄: {age if age else ''}")
                print(f"  时间: {date if date else ''}")
                print(f"  电话: {phone if phone else ''}")
                print(f"  岗位: {position if position else ''}")
                print(f"  地区: {location if location else ''}")
                print(f"  工资: {salary if salary else ''}")
                print(f"  邮箱: {email if email else ''}")
                print(f"  文件名: {test_file.name}")

            except PDFExtractionError as e:
                print(f"✗ 提取失败: {e}")
