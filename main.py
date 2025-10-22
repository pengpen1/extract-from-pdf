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

    filename: str  # 文件名
    name: str  # 姓名（未找到时为"未识别"）
    phone: str  # 手机号（未找到时为"未找到"）
    email: str  # 邮箱（未找到时为"未找到"）


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
                phone = info_extractor.extract_phone(text)
                email = info_extractor.extract_email(text)

                print(f"  姓名: {name if name else '未识别'}")
                print(f"  手机号: {phone if phone else '未找到'}")
                print(f"  邮箱: {email if email else '未找到'}")

            except PDFExtractionError as e:
                print(f"✗ 提取失败: {e}")
