# 溯源自动报告生成工具

为南通溯源病理医学实验室定制的自动化 PDF 报告生成工具。

## 项目概述

本工具可从 Excel 测试数据自动生成专业的 PDF 报告，用于食物过敏检测。支持多种项目类型（IgG-F96-1、IgG-F64-1、IgG-F32-1），并创建包含患者信息、测试结果和分类食物过敏原数据的精美格式化报告。

## 功能特性

- **多项目类型支持**：支持 IgG-F96-1（96项）、IgG-F64-1（64项）和 IgG-F32-1（32项）
- **自动分类**：将食物过敏原自动分类到不同组别（肉类、蛋奶类、水产类、谷物类、水果类、蔬菜类、坚果类、调味类）
- **过敏等级评估**：自动计算并显示过敏等级（正常、轻度、中度、重度）
- **专业 PDF 输出**：使用自定义模板生成高质量 PDF 报告
- **批量处理**：从单个 Excel 文件处理多个患者记录
- **自定义签名**：支持检验者和审核者的个性化签名

## 系统要求

- Python 3.7+
- 所需 Python 包：
  - pandas
  - openpyxl
  - jinja2
- WeasyPrint（Windows 版本已包含在 `bin/` 目录中）

## 安装说明

1. 克隆此仓库：
   ```bash
   git clone https://github.com/Lambowmm/suyuan-auto-report.git
   cd suyuan-auto-report
   ```

2. 安装 Python 依赖：
   ```bash
   pip install -r requirements.txt
   ```

## 使用方法

### 准备数据

1. 将测试数据 Excel 文件命名为 `TestResult.xlsx` 并放置在根目录
2. Excel 文件应遵循以下列结构：
   - 第 1 列：（保留）
   - 第 2 列：测试时间
   - 第 3 列：项目类型（IgG-F96-1、IgG-F64-1 或 IgG-F32-1）
   - 第 4 列：患者 ID
   - 第 5 列：患者姓名
   - 第 6 列：性别
   - 第 7 列：年龄
   - 第 15 列：检验者姓名
   - 第 16 列：审核者姓名
   - 第 17 列及以后：食物过敏原测试结果

### 运行报告生成器

**Python 脚本模式：**
```bash
python generate_reports.py
```

**可执行文件模式（Windows）：**
双击 `generate_reports.exe` 文件（如果已构建可执行文件）

### 输出

生成的 PDF 报告将保存在 `output_reports/` 目录中，文件名格式为：
```
{患者姓名}_{项目类型}_Report.pdf
```

## 项目结构

```
suyuan-auto-report/
├── bin/                    # 二进制可执行文件
│   └── weasyprint.exe     # PDF 生成工具
├── images/                # 报告图像资源
│   ├── sign/             # 签名文件（*.bmp）
│   ├── Back.jpg
│   ├── CommonSymptoms.png
│   ├── Cover.jpg
│   ├── foodGuide.png
│   ├── levelInformation.png
│   ├── logo.png
│   ├── process.jpg
│   └── replaceFood*.png
├── templates/             # 报告 HTML 模板
│   ├── IgG-F32-Template.html
│   ├── IgG-F64-Template.html
│   └── IgG-F96-Template.html
├── generate_reports.py    # 主报告生成脚本
├── TestResult.xlsx        # 输入数据文件（不包含在仓库中）
└── output_reports/        # 生成的 PDF 报告（自动创建）
```

## 支持的项目类型

| 项目类型    | 食物项目数量 | 模板文件               |
|------------|-------------|------------------------|
| IgG-F96-1  | 96 项       | IgG-F96-Template.html  |
| IgG-F64-1  | 64 项       | IgG-F64-Template.html  |
| IgG-F32-1  | 32 项       | IgG-F32-Template.html  |

## 过敏等级阈值

工具根据浓度值自动分类过敏原等级：

- **正常**：< 50 U/mL
- **轻度**：50-99 U/mL
- **中度**：100-199 U/mL
- **重度**：≥ 200 U/mL

## 食物分类

食物过敏原自动分类为：

- 肉类
- 蛋奶类
- 水产类
- 谷物类
- 水果类
- 蔬菜类
- 坚果类
- 调味类

## 自定义设置

### 添加自定义签名

1. 创建一个 BMP 格式的签名图像文件
2. 将其命名为 `{检验者姓名}.bmp` 或 `{审核者姓名}.bmp`
3. 将文件放置在 `images/sign/` 目录中
4. 当 Excel 文件中列出该检验者/审核者时，工具将自动使用相应签名

### 修改模板

HTML 模板位于 `templates/` 目录中。您可以通过编辑这些文件来自定义报告布局。模板使用 Jinja2 模板语法。

## 故障排除

### 常见问题

1. **Excel 文件未找到**：确保 `TestResult.xlsx` 位于根目录
2. **WeasyPrint 错误**：验证 `weasyprint.exe` 是否在 `bin/` 目录中
3. **列缺失**：检查 Excel 文件是否在正确位置包含所有必需列
4. **模板未找到**：确保所有模板文件都在 `templates/` 目录中

### 错误信息

- **"未找到 Excel 文件"**：Excel 文件在预期位置未找到
- **"无法执行命令"**：未找到 WeasyPrint 可执行文件或无法正常工作
- **"跳过不支持的项目类型"**：Excel 文件中的项目类型不受支持
- **"模板文件不存在"**：模板文件缺失

## 开发说明

### 从源码构建

创建独立可执行文件：

```bash
pip install pyinstaller
pyinstaller --onefile --add-data "templates;templates" --add-data "images;images" --add-data "bin;bin" generate_reports.py
```

### 代码结构

主脚本 `generate_reports.py` 分为几个部分：

1. **常量定义**：配置值和阈值
2. **路径配置**：不同运行模式的动态路径解析
3. **数据分类配置**：食物类别映射
4. **工具函数**：数据处理辅助函数
5. **数据处理**：从 Excel 提取和处理数据的函数
6. **PDF 生成**：使用 WeasyPrint 将 HTML 转换为 PDF
7. **主逻辑**：报告生成工作流程

## 许可证

本项目是为南通溯源病理医学实验室定制的专有软件。

## 作者

为南通溯源病理医学实验室创建

## 支持

如有问题或疑问，请联系实验室的 IT 支持团队。
