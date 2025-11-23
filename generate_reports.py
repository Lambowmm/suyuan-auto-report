import datetime
import os
from pathlib import Path

import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML

# ================= 配置区域 =================
# 根据实际文件名修改
EXCEL_PATH = "TestResult(test).xlsx"
TEMPLATE_FILE = "96demo.html"
OUTPUT_DIR = "output_reports"

# 简单的食物分类字典（保持与旧版本一致，方便复用）
CATEGORY_MAP = {
    "肉类": ["火鸡", "鸡肉", "牛肉", "羊肉", "猪肉"],
    "蛋奶类": ["白软干酪","切达干酪","酪蛋白", "酸奶", "鸡蛋", "鸡蛋白", "鸡蛋黄", "羊奶", "牛奶","α-乳清蛋白", "β-乳球蛋白"],
    "水产类": ["鳗鱼","鳕鱼", "三文鱼", "金枪鱼", "草鱼","鳟鱼", "带鱼","沙丁鱼","龙虾","虾","螃蟹","墨鱼","扇贝","牡蛎","蛤"],
    "谷物类": ["大麦","大米", "小米","小麦","黑麦", "燕麦", "荞麦","玉米", "麦芽"],
    "水果类": ["香蕉", "葡萄", "柠檬", "草莓","柚子", "菠萝","榴莲","西瓜", "桃", "芒果","橄榄","苹果","桃","蓝莓","橘子","哈密瓜"],
    "蔬菜类": ["大豆","青豆","豌豆","菠菜","小白菜","生菜","卷心菜","菜花","芹菜", "西兰花","西红柿", "胡萝卜", "黄瓜", "大蒜","大葱","香菜","红辣椒","青椒","洋葱","姜","蘑菇","甘薯","马铃薯","南瓜","茄子"],
    "坚果类": ["花生","葵花籽", "杏仁", "腰果","榛子","黑胡桃","芝麻"],
    "调味类": ["肉桂","红茶", "芥末", "蜂蜜","黄油","酵母","咖啡","巧克力","蔗糖"],
}


def get_category(food_name):
    """根据食物名称模糊匹配分类"""
    if pd.isna(food_name):
        return "其他/未分类"
    name_str = str(food_name)
    for category, keywords in CATEGORY_MAP.items():
        if any(kw in name_str for kw in keywords):
            return category
    return "其他/未分类"


def calculate_level(value):
    """根据浓度计算等级"""
    try:
        val = float(value)
    except (ValueError, TypeError):
        return "normal"

    if val < 50:
        return "normal"
    if 50 <= val < 100:
        return "mild"
    if 100 <= val < 200:
        return "moderate"
    return "severe"


def chunked(items, size):
    """辅助函数，按固定大小切片，供模板分页使用"""
    return [items[i : i + size] for i in range(0, len(items), size)] or [[]]


def generate_reports():
    base_dir = Path(__file__).resolve().parent
    excel_path = base_dir / EXCEL_PATH
    template_path = base_dir / TEMPLATE_FILE
    output_dir = base_dir / OUTPUT_DIR

    if not excel_path.exists():
        raise FileNotFoundError(f"未找到 Excel 文件: {excel_path}")
    if not template_path.exists():
        raise FileNotFoundError(f"未找到模板文件: {template_path}")

    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"正在读取 Excel 文件: {excel_path} ...")

    try:
        df = pd.read_excel(excel_path, engine="openpyxl")
    except Exception as exc:
        print(f"读取 Excel 失败: {exc}")
        print("请确认已安装 openpyxl: pip install openpyxl")
        return

    env = Environment(loader=FileSystemLoader(str(base_dir)))
    template = env.get_template(TEMPLATE_FILE)

    for _, row in df.iterrows():
        p_name = row.get("PatientName")
        if pd.isna(p_name):
            continue

        print(f"正在处理: {p_name} ...")

        foods = []
        for i in range(1, 97):
            name_col = f"Gray{i}"
            val_col = f"Result{i}"
            if name_col not in row or val_col not in row:
                continue

            f_name = row[name_col]
            f_val = row[val_col]
            if pd.isna(f_name) or pd.isna(f_val):
                continue

            foods.append(
                {
                    "name": str(f_name).strip(),
                    "value": float(f_val),
                    "level": calculate_level(f_val),
                    "category": get_category(f_name),
                }
            )

        grouped_foods = {}
        for item in foods:
            grouped_foods.setdefault(item["category"], []).append(item)

        summary = {
            "severe": [x["name"] for x in foods if x["level"] == "severe"],
            "moderate": [x["name"] for x in foods if x["level"] == "moderate"],
            "mild": [x["name"] for x in foods if x["level"] == "mild"],
        }

        test_time = row.get("TestTime", datetime.date.today())
        if isinstance(test_time, str):
            date_str = test_time[:10]
        else:
            date_str = test_time.strftime("%Y-%m-%d") if pd.notnull(test_time) else ""

        context = {
            "patient_name": p_name,
            "gender": row.get("Sex", ""),
            "age": row.get("Age", ""),
            "date_received": date_str,
            "date_report": datetime.date.today().strftime("%Y-%m-%d"),
            "lab_id": str(row.get("ID", "")),
            "foods": foods,
            "food_categories": grouped_foods,
            "summary": summary,
            "pages": chunked(foods, 32),
        }

        html_out = template.render(**context)
        output_filename = output_dir / f"{p_name}_Report.pdf"

        try:
            HTML(string=html_out, base_url=str(base_dir)).write_pdf(str(output_filename))
            print(f"  -> 成功生成: {output_filename}")
        except Exception as exc:
            print(f"  -> 生成失败: {exc}")


if __name__ == "__main__":
    generate_reports()

