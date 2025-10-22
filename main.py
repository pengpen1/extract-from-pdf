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
                phone = info_extractor.extract_phone(text)
                email = info_extractor.extract_email(text)

                print(f"  手机号: {phone if phone else '未找到'}")
                print(f"  邮箱: {email if email else '未找到'}")

            except PDFExtractionError as e:
                print(f"✗ 提取失败: {e}")
