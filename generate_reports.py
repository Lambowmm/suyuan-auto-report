"""
报告生成模块

该模块用于从 Excel 文件中读取测试数据，并根据项目类型生成相应的 PDF 报告。
支持多种项目类型（IgG-F96-1、IgG-F64-1、IgG-F32-1），每种类型对应不同的模板和食物项数量。
"""

import datetime
import os
import subprocess
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any, Union

import pandas as pd
from jinja2 import Environment, FileSystemLoader

# ==================== 常量定义 ====================

# Excel 文件列索引常量
COLUMN_INDEX_PROJECT = 2  # 项目列索引
COLUMN_INDEX_PATIENT_ID = 3  # 患者ID列索引
COLUMN_INDEX_PATIENT_NAME = 4  # 患者名称列索引
COLUMN_INDEX_GENDER = 5  # 性别列索引
COLUMN_INDEX_AGE = 6  # 年龄列索引
COLUMN_INDEX_TEST_TIME = 1  # 测试时间列索引
COLUMN_INDEX_FOOD_START = 15  # 食物数据起始列索引

# 食物分页大小
FOOD_ITEMS_PER_PAGE = 32

# 过敏等级阈值
LEVEL_THRESHOLD_NORMAL = 50
LEVEL_THRESHOLD_MILD = 100
LEVEL_THRESHOLD_MODERATE = 200

# ==================== 路径配置 ====================


def get_base_path() -> Path:
    """
    获取程序的运行根目录。

    兼容 IDE/命令行运行模式 (.py) 和打包后的可执行文件模式 (.exe)。

    Returns:
        Path: 程序运行根目录的路径对象

    Note:
        - 如果是打包后的 exe，sys.executable 指向 exe 文件本身，取其父目录作为基准
        - 如果是普通 python 脚本运行，__file__ 指向脚本文件，取其父目录作为基准
    """
    if getattr(sys, 'frozen', False):
        # 打包后的 exe 模式
        return Path(sys.executable).parent
    else:
        # 普通 Python 脚本运行模式
        return Path(__file__).resolve().parent


# 获取动态基准目录
BASE_DIR = get_base_path()

# 文件路径配置
EXCEL_PATH = BASE_DIR / "TestResult.xlsx"
OUTPUT_DIR = BASE_DIR / "output_reports"
WEASYPRINT_CMD = BASE_DIR / "weasyprint.exe"

# 项目类型到模板文件的映射
PROJECT_TEMPLATE_MAP: Dict[str, Path] = {
    "IgG-F96-1": BASE_DIR / "96-Template.html",
    "IgG-F64-1": BASE_DIR / "64-Template.html",
    "IgG-F32-1": BASE_DIR / "32-Template.html",
}

# 项目类型到食物项数量的映射
PROJECT_ITEM_COUNT: Dict[str, int] = {
    "IgG-F96-1": 96,
    "IgG-F64-1": 64,
    "IgG-F32-1": 32,
}

# ==================== 数据分类配置 ====================

# 食物分类映射字典
CATEGORY_MAP: Dict[str, List[str]] = {
    "肉类": ["火鸡", "鸡肉", "牛肉", "羊肉", "猪肉"],
    "蛋奶类": ["白软干酪", "切达干酪", "酪蛋白", "酸奶", "鸡蛋", "鸡蛋白", "鸡蛋黄", "羊奶", "牛奶",
              "α-乳清蛋白", "β-乳球蛋白"],
    "水产类": ["鳗鱼", "鳕鱼", "三文鱼", "金枪鱼", "草鱼", "鳟鱼", "带鱼", "沙丁鱼", "龙虾", "虾",
              "螃蟹", "墨鱼", "扇贝", "牡蛎", "蛤"],
    "谷物类": ["大麦", "大米", "小米", "小麦", "黑麦", "燕麦", "荞麦", "玉米", "麦芽"],
    "水果类": ["香蕉", "葡萄", "柠檬", "草莓", "柚子", "菠萝", "榴莲", "西瓜", "桃", "芒果",
              "橄榄", "苹果", "蓝莓", "橘子", "哈密瓜"],
    "蔬菜类": ["大豆", "青豆", "豌豆", "菠菜", "小白菜", "生菜", "卷心菜", "菜花", "芹菜",
              "西兰花", "西红柿", "胡萝卜", "黄瓜", "大蒜", "大葱", "香菜", "红辣椒", "青椒",
              "洋葱", "姜", "蘑菇", "甘薯", "马铃薯", "南瓜", "茄子"],
    "坚果类": ["花生", "葵花籽", "杏仁", "腰果", "榛子", "黑胡桃", "芝麻"],
    "调味类": ["肉桂", "红茶", "芥末", "蜂蜜", "黄油", "酵母", "咖啡", "巧克力", "蔗糖"],
}

# ==================== 工具函数 ====================


def get_category(food_name: Any) -> str:
    """
    根据食物名称模糊匹配分类。

    Args:
        food_name: 食物名称，可以是字符串或其他类型

    Returns:
        str: 食物分类名称，如果未匹配到则返回 "其他/未分类"
    """
    if pd.isna(food_name):
        return "其他/未分类"

    name_str = str(food_name)
    for category, keywords in CATEGORY_MAP.items():
        if any(keyword in name_str for keyword in keywords):
            return category

    return "其他/未分类"


def calculate_level(value: Any) -> str:
    """
    根据浓度值计算过敏等级。

    Args:
        value: 浓度值，可以是数字或字符串

    Returns:
        str: 过敏等级，可能的值：
            - "normal": 正常 (< 50)
            - "mild": 轻度 (50-99)
            - "moderate": 中度 (100-199)
            - "severe": 重度 (>= 200)
    """
    try:
        val = float(value)
    except (ValueError, TypeError):
        return "normal"

    if val < LEVEL_THRESHOLD_NORMAL:
        return "normal"
    elif val < LEVEL_THRESHOLD_MILD:
        return "mild"
    elif val < LEVEL_THRESHOLD_MODERATE:
        return "moderate"
    else:
        return "severe"


def chunked(items: List[Any], size: int) -> List[List[Any]]:
    """
    将列表按固定大小分块。

    Args:
        items: 要分块的列表
        size: 每块的大小

    Returns:
        List[List[Any]]: 分块后的列表，如果原列表为空则返回包含空列表的列表
    """
    if not items:
        return [[]]

    return [items[i:i + size] for i in range(0, len(items), size)]


def get_project_info(project_value: Any) -> Tuple[Optional[str], Optional[Path], Optional[int]]:
    """
    根据项目值获取模板文件和食物项数量。

    Args:
        project_value: 项目值，可以是字符串或其他类型

    Returns:
        Tuple[Optional[str], Optional[Path], Optional[int]]:
            - 项目类型字符串（如果项目值无效则为 None）
            - 模板文件路径（如果项目类型不支持则为 None）
            - 食物项数量（如果项目类型不支持则为 None）
    """
    if pd.isna(project_value):
        return None, None, None

    project_str = str(project_value).strip()
    if project_str in PROJECT_TEMPLATE_MAP:
        template_file = PROJECT_TEMPLATE_MAP[project_str]
        item_count = PROJECT_ITEM_COUNT[project_str]
        return project_str, template_file, item_count

    return project_str, None, None


# ==================== 数据处理函数 ====================


def extract_food_data(
    info_row: pd.Series,
    value_row: pd.Series,
    df: pd.DataFrame,
    food_start_col_idx: int,
    item_count: int
) -> List[Dict[str, Any]]:
    """
    从数据行中提取食物数据。

    Args:
        info_row: 包含食物名称的信息行
        value_row: 包含食物数值的数据行
        df: 完整的 DataFrame
        food_start_col_idx: 食物数据起始列索引
        item_count: 需要读取的食物项数量

    Returns:
        List[Dict[str, Any]]: 食物数据列表，每个元素包含：
            - name: 食物名称
            - value: 浓度值
            - level: 过敏等级
            - category: 食物分类
    """
    foods: List[Dict[str, Any]] = []

    end_col_idx = min(food_start_col_idx + item_count, len(df.columns))
    for col_idx in range(food_start_col_idx, end_col_idx):
        col_name = df.columns[col_idx]
        food_name = info_row[col_name]
        food_value = value_row[col_name]

        # 跳过名称和数值都为空的情况
        if pd.isna(food_name) and pd.isna(food_value):
            continue

        # 跳过名称或数值为空的情况
        if pd.isna(food_name) or pd.isna(food_value):
            continue

        try:
            foods.append({
                "name": str(food_name).strip(),
                "value": float(food_value),
                "level": calculate_level(food_value),
                "category": get_category(food_name),
            })
        except (ValueError, TypeError):
            # 数值转换失败，跳过该项
            continue

    return foods


def extract_patient_info(info_row: pd.Series, df: pd.DataFrame) -> Dict[str, str]:
    """
    从信息行中提取患者信息。

    Args:
        info_row: 包含患者信息的数据行
        df: 完整的 DataFrame

    Returns:
        Dict[str, str]: 患者信息字典，包含：
            - gender: 性别
            - age: 年龄
            - lab_id: 患者ID
            - date_received: 测试时间（格式：YYYY-MM-DD）
    """
    gender = ""
    age = ""
    lab_id = ""
    date_received = ""

    # 获取性别
    if len(df.columns) > COLUMN_INDEX_GENDER:
        gender_col = df.columns[COLUMN_INDEX_GENDER]
        if not pd.isna(info_row[gender_col]):
            gender = str(info_row[gender_col])

    # 获取年龄
    if len(df.columns) > COLUMN_INDEX_AGE:
        age_col = df.columns[COLUMN_INDEX_AGE]
        if not pd.isna(info_row[age_col]):
            age = str(info_row[age_col])

    # 获取患者ID
    if len(df.columns) > COLUMN_INDEX_PATIENT_ID:
        id_col = df.columns[COLUMN_INDEX_PATIENT_ID]
        if not pd.isna(info_row[id_col]):
            lab_id = str(info_row[id_col])

    # 获取测试时间
    if len(df.columns) > COLUMN_INDEX_TEST_TIME:
        time_col = df.columns[COLUMN_INDEX_TEST_TIME]
        test_time = info_row[time_col]

        if isinstance(test_time, str):
            date_received = test_time[:10] if len(test_time) >= 10 else ""
        elif pd.notnull(test_time):
            try:
                date_received = test_time.strftime("%Y-%m-%d")
            except AttributeError:
                # 如果不是日期对象，尝试转换为字符串
                date_received = str(test_time)[:10] if len(str(test_time)) >= 10 else ""

    return {
        "gender": gender,
        "age": age,
        "lab_id": lab_id,
        "date_received": date_received,
    }


def process_food_summary(foods: List[Dict[str, Any]]) -> Dict[str, List[str]]:
    """
    处理食物数据，生成按等级分类的摘要。

    Args:
        foods: 食物数据列表

    Returns:
        Dict[str, List[str]]: 按等级分类的食物名称列表，包含：
            - severe: 重度过敏食物列表
            - moderate: 中度过敏食物列表
            - mild: 轻度过敏食物列表
    """
    return {
        "severe": [item["name"] for item in foods if item["level"] == "severe"],
        "moderate": [item["name"] for item in foods if item["level"] == "moderate"],
        "mild": [item["name"] for item in foods if item["level"] == "mild"],
    }


def group_foods_by_category(foods: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
    """
    按分类对食物进行分组。

    Args:
        foods: 食物数据列表

    Returns:
        Dict[str, List[Dict[str, Any]]]: 按分类分组的食物字典
    """
    grouped: Dict[str, List[Dict[str, Any]]] = {}
    for item in foods:
        category = item["category"]
        grouped.setdefault(category, []).append(item)

    return grouped


def generate_pdf_from_html(
    html_content: str,
    output_path: Path,
    patient_name: str,
    project_type: str
) -> bool:
    """
    将 HTML 内容转换为 PDF 文件。

    Args:
        html_content: HTML 内容字符串
        output_path: 输出 PDF 文件路径
        patient_name: 患者名称（用于临时文件命名）
        project_type: 项目类型（用于临时文件命名）

    Returns:
        bool: 转换是否成功
    """
    temp_html_path = BASE_DIR / f"{patient_name}_{project_type}_temp.html"

    try:
        # 保存 HTML 到临时文件
        with open(temp_html_path, "w", encoding="utf-8") as f:
            f.write(html_content)

        # 调用 WeasyPrint 命令行工具转换
        cmd = [str(WEASYPRINT_CMD), str(temp_html_path), str(output_path)]

        subprocess.run(
            cmd,
            check=True,
            cwd=str(BASE_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        return True

    except subprocess.CalledProcessError as e:
        print(f"  -> 生成失败 (WeasyPrint 错误): {e}")
        try:
            error_msg = e.stderr.decode('utf-8', errors='ignore')
            if error_msg:
                print(f"      错误详情: {error_msg}")
        except Exception:
            pass
        return False

    except Exception as exc:
        print(f"  -> 生成失败 (其他错误): {exc}")
        return False

    finally:
        # 清理临时文件
        if temp_html_path.exists():
            try:
                os.remove(temp_html_path)
            except Exception:
                pass


# ==================== 主函数 ====================


def validate_environment() -> bool:
    """
    验证运行环境，检查必要的文件和工具是否存在。

    Returns:
        bool: 环境验证是否通过
    """
    # 检查 Excel 文件是否存在
    if not EXCEL_PATH.exists():
        print(f"错误: 未找到 Excel 文件: {EXCEL_PATH}")
        print(f"当前运行目录: {BASE_DIR}")
        return False

    # 检查 WeasyPrint 命令是否可用
    try:
        subprocess.run(
            [str(WEASYPRINT_CMD), "--version"],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
    except (subprocess.CalledProcessError, FileNotFoundError):
        print(f"错误: 无法执行命令 '{WEASYPRINT_CMD.name}'")
        print(f"请确保 {WEASYPRINT_CMD.name} 位于此文件夹: {BASE_DIR}")
        return False

    return True


def process_single_report(
    df: pd.DataFrame,
    env: Environment,
    row_index: int
) -> int:
    """
    处理单个报告。

    Args:
        df: 完整的 DataFrame
        env: Jinja2 环境对象
        row_index: 当前处理的行索引

    Returns:
        int: 处理完成后应该跳过的行数（通常是 2，表示项目信息行和数值行）
    """
    # 获取项目信息行
    info_row = df.iloc[row_index]
    project_col = df.columns[COLUMN_INDEX_PROJECT]
    patient_name_col = df.columns[COLUMN_INDEX_PATIENT_NAME]
    project_value = info_row[project_col]

    # 检查是否是项目信息行
    if pd.isna(project_value):
        return 1

    # 获取项目类型和模板信息
    project_type, template_file, item_count = get_project_info(project_value)

    # 检查项目类型是否支持
    if template_file is None:
        patient_name = info_row[patient_name_col] if not pd.isna(info_row[patient_name_col]) else "未知"
        supported_types = ", ".join(PROJECT_TEMPLATE_MAP.keys())
        print(f"警告: 跳过不支持的项目类型 '{project_type}' (患者: {patient_name})")
        print(f"      支持的项目类型: {supported_types}")
        return 2

    # 检查模板文件是否存在
    if not template_file.exists():
        patient_name = info_row[patient_name_col] if not pd.isna(info_row[patient_name_col]) else "未知"
        print(f"警告: 模板文件不存在 '{template_file.name}' (患者: {patient_name}, 项目: {project_type})")
        return 2

    # 获取患者名称
    patient_name = info_row[patient_name_col]
    if pd.isna(patient_name):
        print(f"警告: 第 {row_index + 1} 行患者名称为空，跳过...")
        return 2

    print(f"正在处理: {patient_name} ({project_type}) ...")

    # 检查是否有数值行
    if row_index + 1 >= len(df):
        print(f"警告: 第 {row_index + 1} 行缺少对应的数值行，跳过...")
        return 1

    value_row = df.iloc[row_index + 1]

    # 提取食物数据
    foods = extract_food_data(
        info_row,
        value_row,
        df,
        COLUMN_INDEX_FOOD_START,
        item_count
    )

    if not foods:
        print(f"警告: {patient_name} 没有有效的食物数据，跳过...")
        return 2

    # 处理食物分类和摘要
    grouped_foods = group_foods_by_category(foods)
    summary = process_food_summary(foods)

    # 提取患者信息
    patient_info = extract_patient_info(info_row, df)

    # 加载模板并渲染
    template = env.get_template(template_file.name)

    context = {
        "patient_name": str(patient_name),
        "gender": patient_info["gender"],
        "age": patient_info["age"],
        "date_received": patient_info["date_received"],
        "date_report": datetime.date.today().strftime("%Y-%m-%d"),
        "lab_id": patient_info["lab_id"],
        "foods": foods,
        "food_categories": grouped_foods,
        "summary": summary,
        "pages": chunked(foods, FOOD_ITEMS_PER_PAGE),
    }

    html_content = template.render(**context)

    # 生成 PDF
    output_filename = OUTPUT_DIR / f"{patient_name}_{project_type}_Report.pdf"

    success = generate_pdf_from_html(html_content, output_filename, patient_name, project_type)

    if success:
        print(f"  -> 成功生成: {output_filename.name}")

    return 2


def generate_reports() -> None:
    """
    主函数：从 Excel 文件读取数据并生成 PDF 报告。

    处理流程：
    1. 验证运行环境
    2. 读取 Excel 文件
    3. 配置 Jinja2 模板环境
    4. 遍历数据行，为每个患者生成报告
    5. 使用 WeasyPrint 将 HTML 转换为 PDF

    Raises:
        FileNotFoundError: 当 Excel 文件不存在时
        Exception: 当读取 Excel 文件失败时
    """
    # 验证环境
    if not validate_environment():
        return

    # 创建输出目录
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    print(f"正在读取 Excel 文件: {EXCEL_PATH.name} ...")

    # 读取 Excel 文件
    try:
        df = pd.read_excel(EXCEL_PATH, engine="openpyxl")
    except Exception as exc:
        print(f"读取 Excel 失败: {exc}")
        print("请确认已安装 openpyxl: pip install openpyxl")
        return

    # 验证列数
    if len(df.columns) <= COLUMN_INDEX_PATIENT_NAME:
        print("错误: Excel 文件列数不足，请检查文件格式")
        return

    # 配置 Jinja2 环境
    template_dir = BASE_DIR
    env = Environment(loader=FileSystemLoader(str(template_dir)))

    # 遍历处理数据（每次处理两行：项目信息行和数值行）
    row_index = 0
    while row_index < len(df):
        skip_rows = process_single_report(df, env, row_index)
        row_index += skip_rows

    print("\n报告生成完成！")


if __name__ == "__main__":
    generate_reports()
