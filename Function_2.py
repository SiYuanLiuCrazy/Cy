from pptx import Presentation
import os

def count_slides_in_folder(folder_path):
    ppt_files_pages = {}
    
    # 遍历文件夹中的每个PPT文件
    for filename in os.listdir(folder_path):
        if filename.endswith('.pptx'):
            pptx_file = os.path.join(folder_path, filename)
            presentation = Presentation(pptx_file)
            num_slides = len(presentation.slides)
            ppt_files_pages[filename] = num_slides
    
    return ppt_files_pages

# 输入文件夹路径
folder_path = r'E:\PPT_New'

# 调用函数获取每个PPT文件的页数
ppt_files_pages = count_slides_in_folder(folder_path)
print('ppt_files_pages[0]')
# 打印每个PPT文件的页数
for filename, num_pages in ppt_files_pages.items():
    print(f"{filename}: {num_pages} 页。")

