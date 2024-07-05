import os
import shutil

def copy_ppt_files(source_folder: str, destination_folder: str):
    # 检查目标文件夹是否存在，如果不存在则创建
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    
    # 遍历源文件夹中的文件
    for filename in os.listdir(source_folder):
        # 检查文件扩展名是否为ppt或pptx
        if filename.endswith('.ppt') or filename.endswith('.pptx'):
            # 构建文件的完整路径
            source_path = os.path.join(source_folder, filename)
            destination_path = os.path.join(destination_folder, filename)
            
            # 将文件复制到目标文件夹
            shutil.copy2(source_path, destination_path)
            print(f"Copied: {filename}")

source_folder = 'E:\PPT_New'  # 替换为实际的源文件夹路径
destination_folder = 'E:\\PPT_Home'  # 目标文件夹路径

copy_ppt_files(source_folder, destination_folder)