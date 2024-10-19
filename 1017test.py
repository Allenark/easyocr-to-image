import os
import shutil
import mimetypes  # 用于检测文件类型
from PIL import Image
import easyocr
import pandas as pd
import numpy as np

input("Press Enter to start...")

# --------------------- 图片预处理 ---------------------
# 指定包含原始图片的文件夹路径
input_folder_path = r'C:\Users\allenliu\Desktop\date\testimage'

# 在输入文件夹中新建一个子文件夹作为输出文件夹
output_folder_path = os.path.join(input_folder_path, 'preprocessed')

# 删除旧的输出文件夹并创建新的输出文件夹
if os.path.exists(output_folder_path):
    shutil.rmtree(output_folder_path)
os.makedirs(output_folder_path)

# 需要去除的边栏尺寸
right_padding = 660  # 右边栏宽度
top_padding = 0      # 假设没有上边距需要裁剪，可以修改
bottom_padding = 0   # 假设没有下边距需要裁剪，可以修改

# 遍历文件夹中的所有文件
for file_name in os.listdir(input_folder_path):
    file_path = os.path.join(input_folder_path, file_name)

    # 检查文件是否为图片
    mime_type, _ = mimetypes.guess_type(file_path)
    if mime_type and mime_type.startswith('image'):  # 仅处理图片文件
        print(f"正在处理图片: {file_name}")
        
        # 打开图片
        with Image.open(file_path) as img:
            # 获取图片原始尺寸
            width, height = img.size
            
            # 计算去除边栏后的尺寸
            left = 0
            top = top_padding
            right = width - right_padding
            bottom = height - bottom_padding
            
            # 检查裁剪区域是否合理
            if right > left and bottom > top:
                # 裁剪图片
                img_cropped = img.crop((left, top, right, bottom))
                
                # 压缩图片体积，例如将图片质量设置为70%
                img_cropped = img_cropped.convert('RGB')
                output_image_path = os.path.join(output_folder_path, file_name)
                img_cropped.save(output_image_path, quality=70)  # 设置JPEG质量为70%
            else:
                print(f"无法裁剪图片 {file_name}，检查裁剪区域是否合理。")
    else:
        print(f"跳过非图片文件: {file_name}")

print("图片预处理完成，处理后的图片已保存到指定文件夹。")

# --------------------- OCR 识别 ---------------------
# 创建EasyOCR的Reader对象，指定要识别的语言（只需初始化一次）
reader = easyocr.Reader(['en'])  # 英文模型可以识别数字和小数点

# 指定包含图片的文件夹路径
image_folder_path = output_folder_path

# 初始化一个列表来存储所有图片的识别结果
all_results = []

# 遍历文件夹中的所有文件
for image_file in os.listdir(image_folder_path):
    image_path = os.path.join(image_folder_path, image_file)

    # 检查文件是否为图片
    mime_type, _ = mimetypes.guess_type(image_path)
    if mime_type and mime_type.startswith('image'):
        print(f"正在识别图片: {image_file}")
        
        # 使用EasyOCR读取图片中的文本
        results = reader.readtext(image_path)
        
        # 提取识别的文本和置信度
        detected_texts = []
        for (bbox, text, prob) in results:
            if prob >= 0.5:  # 只考虑置信度大于等于50%的结果
                detected_texts.append(text)
        
        # 将识别的文本添加到列表中
        detected_numbers = ' '.join(detected_texts)
        all_results.append({
            'Image File Name': image_file,
            'Detected Numbers': detected_numbers
        })
    else:
        print(f"跳过非图片文件: {image_file}")

# 创建DataFrame
df = pd.DataFrame(all_results)

# 导出到Excel文件
output_excel_path = r'C:\Users\allenliu\Desktop\date\detected_numbers.xlsx'
df.to_excel(output_excel_path, index=False)

print(f"识别完成，结果已导出到Excel文件: {output_excel_path}")

# --------------------- 数据处理 ---------------------
try:
    # 读取输入的Excel文件
    input_excel_path = output_excel_path
    df = pd.read_excel(input_excel_path)

    # 显示读取的DataFrame内容和列名
    print("读取的DataFrame内容：")
    print(df)

    print("\n列名：")
    print(df.columns)

    # Step 1: 将第二列的多个数据分开
    df_split = df['Detected Numbers'].str.split(' ', expand=True)

    # 显示拆分后的DataFrame
    print("\n拆分后的DataFrame：")
    print(df_split)

    # Step 2: 找到每一行中所有包含 "um" 的列
    um_data = df_split.apply(lambda row: row[row.str.contains('um', na=False, case=False)], axis=1)

    # 显示提取的 "um" 数据
    print("\n提取的 'um' 数据：")
    print(um_data)

    # Step 3: 处理提取的数据，确保每行有两列
    processed_data = []
    for index, row in um_data.iterrows():  # 使用 iterrows() 处理每一行
        # 取出每行中的有效数据
        values = row.dropna().tolist()
        # 只取前两项，并删除其中的 'um' 和其他字母
        values = [str(v).replace('um', '').strip() for v in values[:2]]
        processed_data.append(values + [''] * (2 - len(values)))  # 确保每行有两列

    # 转换为DataFrame
    result_df = pd.DataFrame(processed_data, columns=['横尺寸(um)', '竖尺寸(um)'])

    # 添加序号列
    result_df.insert(0, '序号', range(1, len(result_df) + 1))

    # 将横尺寸和竖尺寸转换为浮点数，便于后续计算
    result_df['横尺寸(um)'] = pd.to_numeric(result_df['横尺寸(um)'], errors='coerce')
    result_df['竖尺寸(um)'] = pd.to_numeric(result_df['竖尺寸(um)'], errors='coerce')

    # Step 4: 当只有一个尺寸有数据时，填充另一个尺寸
    result_df['横尺寸(um)'].fillna(result_df['竖尺寸(um)'], inplace=True)
    result_df['竖尺寸(um)'].fillna(result_df['横尺寸(um)'], inplace=True)

    # 计算实际对角尺寸
    result_df['实际对角尺寸'] = np.sqrt(result_df['横尺寸(um)']**2 + result_df['竖尺寸(um)']**2)

    # 添加比例尺设置边长（厘米）和比例尺（mm）列
    result_df['比例尺设置边长（厘米）'] = 1.37
    result_df['比例尺(mm)'] = 5000

    # 添加图中斜边线长（厘米）列
    result_df['图中斜边线长（厘米）'] = (
        result_df['实际对角尺寸'] * result_df['比例尺设置边长（厘米）'] / result_df['比例尺(mm)']
    )

    # 显示处理后的DataFrame内容
    print("\n处理后的DataFrame内容：")
    print(result_df)

    # Step 5: 保存为新的Excel文件
    output_excel_path = r'C:\Users\allenliu\Desktop\date\split_output_final.xlsx'
    result_df.to_excel(output_excel_path, index=False)

    print(f"数据处理完成，新的Excel文件已保存到: {output_excel_path}")

except Exception as e:
    print(f"程序运行时发生错误: {e}")

# 等待用户按下回车键
input("按下回车键退出...")
