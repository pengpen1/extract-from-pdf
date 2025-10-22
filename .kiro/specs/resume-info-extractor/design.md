# 设计文档

## 概述

简历信息提取工具采用Python开发，使用PyInstaller打包为Windows可执行文件。系统采用模块化设计，包含PDF文本提取、信息识别、数据导出和用户交互四个核心模块。工具设计为命令行应用，通过简单的双击运行即可完成从文件扫描到结果导出的全流程。

## 架构

### 系统架构图

```
┌─────────────────────────────────────────────────┐
│              用户交互层 (CLI)                      │
│  - 启动欢迎信息                                    │
│  - 进度显示                                       │
│  - 结果统计                                       │
└────────────────┬────────────────────────────────┘
                 │
┌────────────────▼────────────────────────────────┐
│            主控制器 (Main Controller)            │
│  - 文件扫描                                       │
│  - 流程编排                                       │
│  - 错误处理                                       │
└────┬───────────┬───────────┬────────────────────┘
     │           │           │
┌────▼─────┐ ┌──▼──────┐ ┌─▼──────────┐
│ PDF提取器 │ │ 信息提取器 │ │ Excel导出器 │
│          │ │          │ │           │
│ PyPDF2/  │ │ - 姓名   │ │ openpyxl  │
│ pdfplumber│ │ - 电话   │ │           │
│          │ │ - 邮箱   │ │           │
└──────────┘ └──────────┘ └───────────┘
```

### 技术栈

- **Python 3.8+**: 核心开发语言
- **PyPDF2 / pdfplumber**: PDF文本提取库
- **openpyxl**: Excel文件生成库
- **re (正则表达式)**: 信息模式匹配
- **PyInstaller**: 打包工具，生成独立可执行文件
- **pathlib**: 文件路径处理

## 组件和接口

### 1. 文件扫描器 (FileScanner)

**职责**: 递归扫描指定目录，收集所有PDF文件路径

**接口**:
```python
class FileScanner:
    def __init__(self, data_folder: str):
        """初始化扫描器
        Args:
            data_folder: 数据文件夹路径
        """
        
    def scan_pdf_files(self) -> List[Path]:
        """扫描所有PDF文件
        Returns:
            PDF文件路径列表
        """
```

### 2. PDF文本提取器 (PDFExtractor)

**职责**: 从PDF文件中提取文本内容

**接口**:
```python
class PDFExtractor:
    def extract_text(self, pdf_path: Path) -> str:
        """提取PDF文本
        Args:
            pdf_path: PDF文件路径
        Returns:
            提取的文本内容
        Raises:
            PDFExtractionError: 提取失败时抛出
        """
```

**实现策略**:
- 优先使用 pdfplumber（更准确）
- 如果失败，回退到 PyPDF2
- 提取前3页内容（简历关键信息通常在前面）

### 3. 信息提取器 (InfoExtractor)

**职责**: 从文本中识别和提取姓名、电话、邮箱

**接口**:
```python
class InfoExtractor:
    def extract_phone(self, text: str) -> Optional[str]:
        """提取手机号码
        Args:
            text: 简历文本
        Returns:
            手机号码或None
        """
    
    def extract_email(self, text: str) -> Optional[str]:
        """提取邮箱地址
        Args:
            text: 简历文本
        Returns:
            邮箱地址或None
        """
    
    def extract_name(self, text: str) -> Optional[str]:
        """提取姓名
        Args:
            text: 简历文本
        Returns:
            姓名或None
        """
```

**提取策略**:

**电话号码**:
- 正则表达式: `1[3-9]\d{9}`
- 匹配11位手机号
- 返回第一个匹配结果

**邮箱地址**:
- 正则表达式: `[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}`
- 返回第一个匹配结果

**姓名识别** (多策略组合):
1. **关键词模式**: 搜索"姓名："、"姓 名："、"Name:"等关键词后的内容
2. **位置启发式**: 在简历前200字符中查找2-4个连续中文字符
3. **过滤规则**: 
   - 排除常见标题词（"个人简历"、"求职意向"等）
   - 排除包含数字或特殊符号的文本
   - 优先选择中文姓名

### 4. Excel导出器 (ExcelExporter)

**职责**: 将提取结果导出为Excel文件

**接口**:
```python
class ExcelExporter:
    def __init__(self, output_dir: Path):
        """初始化导出器
        Args:
            output_dir: 输出目录
        """
    
    def export(self, data: List[ResumeInfo]) -> Path:
        """导出数据到Excel
        Args:
            data: 简历信息列表
        Returns:
            生成的Excel文件路径
        """
```

**输出格式**:
- 文件名: `简历提取结果_YYYYMMDD_HHMMSS.xlsx`
- 表头: 文件名 | 姓名 | 手机号 | 邮箱
- 单元格宽度自动调整

### 5. 主控制器 (ResumeExtractorApp)

**职责**: 协调各组件，控制整体流程

**接口**:
```python
class ResumeExtractorApp:
    def __init__(self, data_folder: str = "datas"):
        """初始化应用
        Args:
            data_folder: 数据文件夹名称
        """
    
    def run(self):
        """运行主流程"""
```

**流程**:
1. 检查/创建数据文件夹
2. 扫描PDF文件
3. 逐个处理文件（显示进度）
4. 收集提取结果
5. 导出Excel
6. 显示统计信息
7. 等待用户按键退出

## 数据模型

### ResumeInfo

```python
@dataclass
class ResumeInfo:
    """简历信息数据类"""
    filename: str          # 文件名
    name: str             # 姓名（未找到时为"未识别"）
    phone: str            # 手机号（未找到时为"未找到"）
    email: str            # 邮箱（未找到时为"未找到"）
```

### ProcessingResult

```python
@dataclass
class ProcessingResult:
    """处理结果统计"""
    total: int            # 总文件数
    success: int          # 成功处理数
    failed: int           # 失败数
    failed_files: List[str]  # 失败文件列表
```

## 错误处理

### 错误类型

1. **文件访问错误**: PDF文件损坏、权限不足
2. **PDF解析错误**: 加密PDF、扫描版PDF（纯图片）
3. **编码错误**: 特殊字符处理
4. **Excel写入错误**: 磁盘空间不足、权限问题

### 处理策略

- **单文件错误**: 记录错误信息，继续处理其他文件
- **致命错误**: 显示错误信息，等待用户确认后退出
- **错误日志**: 在控制台显示详细错误信息
- **容错机制**: 
  - PDF提取失败时尝试备用库
  - 信息提取失败时填充默认值（"未找到"/"未识别"）

### 异常类定义

```python
class PDFExtractionError(Exception):
    """PDF提取异常"""
    pass

class ExcelExportError(Exception):
    """Excel导出异常"""
    pass
```

## 打包配置

### PyInstaller配置

```python
# build.spec 配置要点
a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],  # 无需额外数据文件
    hiddenimports=['openpyxl', 'pdfplumber', 'PyPDF2'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='简历信息提取工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # 压缩可执行文件
    console=True,  # 显示控制台窗口
    icon=None,
)
```

### 打包命令

```bash
pyinstaller --onefile --name="简历信息提取工具" --console main.py
```

## 用户交互设计

### 启动界面

```
========================================
    简历信息提取工具 v1.0
========================================
使用说明：
1. 将PDF简历放入 'datas' 文件夹
2. 双击运行本程序
3. 等待处理完成，查看生成的Excel文件
========================================
```

### 处理进度显示

```
正在扫描文件...
找到 15 个PDF文件

开始处理...
[1/15] 正在处理: resume1.pdf ... ✓
[2/15] 正在处理: resume2.pdf ... ✓
[3/15] 正在处理: resume3.pdf ... ✗ (无法读取文件)
...
```

### 完成界面

```
========================================
处理完成！
========================================
总文件数: 15
成功: 13
失败: 2

失败文件:
  - resume3.pdf (无法读取文件)
  - resume8.pdf (PDF已加密)

输出文件: D:\简历提取结果_20251022_143025.xlsx
========================================
按任意键退出...
```

## 测试策略

### 单元测试

- **InfoExtractor**: 测试各种格式的姓名、电话、邮箱提取
- **FileScanner**: 测试递归扫描功能
- **PDFExtractor**: 测试不同PDF格式的文本提取

### 集成测试

- 准备多种格式的测试简历（10-20份）
- 测试完整流程：扫描 → 提取 → 导出
- 验证Excel输出的正确性

### 边界测试

- 空文件夹
- 单个文件
- 大量文件（100+）
- 损坏的PDF
- 加密的PDF
- 扫描版PDF（纯图片）
- 特殊字符文件名

### 兼容性测试

- Windows 10
- Windows 11
- 不同分辨率和DPI设置

## 性能考虑

- **内存优化**: 逐个处理文件，不一次性加载所有PDF到内存
- **处理速度**: 单个PDF处理时间控制在1-3秒内
- **文件限制**: 支持处理至少100个PDF文件
- **PDF页数限制**: 只提取前3页内容，提高处理速度

## 部署说明

### 交付物

1. `简历信息提取工具.exe` - 可执行文件
2. `README.txt` - 使用说明文档

### 使用流程

1. 将exe文件放到任意目录
2. 首次运行会自动创建 `datas` 文件夹
3. 将PDF简历放入 `datas` 文件夹（支持子文件夹）
4. 双击运行exe文件
5. 等待处理完成
6. 在exe同目录下查看生成的Excel文件

### 系统要求

- 操作系统: Windows 10 或更高版本
- 磁盘空间: 至少100MB可用空间
- 无需安装Python或其他依赖
