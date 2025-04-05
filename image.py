import os
import pandas as pd
from openpyxl import load_workbook
from PIL import Image
import io
import re

# 创建保存图片的目录
output_dir = "product_images"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# 加载Excel文件
file_path = "a.xlsx"
wb = load_workbook(file_path)
ws = wb.active

# 读取Excel数据获取所需列
df = pd.read_excel(file_path)
# 假设第一列是品牌，第二列是型号，第四列是品名
# 如果列的位置不同，请调整下面的索引
brands = df.iloc[:, 0]  # 获取第一列作为品牌
model_numbers = df.iloc[:, 1]  # 获取第二列作为型号
product_names = df.iloc[:, 3]  # 获取第四列作为品名

# 修复文件路径错误问题


# 提取图片
for row_idx, (brand, model, product_name) in enumerate(zip(brands, model_numbers, product_names), start=2):
    # 检查单元格是否含有图片
    if ws._images:
        for img in ws._images:
            # Excel中图片的位置是基于单元格的，检查图片是否在当前行的第五列
            if img.anchor._from.row == row_idx - 1 and img.anchor._from.col == 4:  # 列从0开始索引
                # 提取图片数据
                img_data = img._data()
                # 保存图片
                brand_str = str(brand).strip()
                model_str = str(model).strip()
                product_name_str = str(product_name).strip()
                
                if model_str:  # 确保型号不为空
                    # 创建"品牌-型号-品名"格式的文件名
                    file_name = f"{brand_str}-{model_str}-{product_name_str}"
                    # 替换Windows文件系统不允许的字符，包括斜杠
                    safe_file_name = re.sub(r'[\\/*?:"<>|]', '_', file_name)
                    img_path = os.path.join(output_dir, f"{safe_file_name}.jpg")
                    with open(img_path, "wb") as f:
                        f.write(img_data)
                    print(f"已保存图片: {img_path}")

print("所有图片已提取完成！")
