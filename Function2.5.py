from spire.presentation import *
License.SetLicenseKey("3vgBAP6gsMpt19Lp5/dLFp/xNXCyz2vrwWXv6RZ9dhyEE6o0AFXtRmuqeHo1LxSwxYqobFgqbUkY/8y0oqHns0AgcGdhUi/gxW3rgxU03m9Je53kYu9SVmpp/9SEEkUhzJ08Ir3+/014o07W0UXzj63zhaQjMxz5Lmx+U08qabOyQmGmnbolIYd70slekKohJLucSw9KL/OxAzIhBdI1+IaElvL2jUsl1w0zGfyVVG56vcENaCh23rDBNUdkTHccDEOW9L3Qa1DZtUpJadmED0scQbbsDHtrJ8sR03yxa2stkpjXUYcVyE9xZphrvWHYvvxAt1UOFOpYxzSwvxS7TIFEhXQBzcHRpEidD1BnBQgIMvo9tdqQQCKCx4k6r9b0++Xs0vwJGUyJGg5a2hVtYnTcPp4ZfPZP8U1XI3sEaVr7BKy44sN4aFRUcOg8E5BVsXJdZQZnWxm0ZPFRazyGL0Ki6Vi0FchpZ1zlqjR15y8GrdZqd2Y4TqYp6gotvPdSH9uirPRDJTPvAvemo491EoPQo83odCDE2MoYu6dm2zNJWc0/ac6sPCt2WR5p9gibTkyjpzW9t+d8YVbjszXSecM+gi5WIBVg98flw/RVhsNg6kxqQ9EqyZXvYxXBCPoygWjJRlC8ETyeMkQy9Cw3t1nUCeFjrpJlWQMuavMa79tLT0anjYYbbl9k/j73i5lG4tmSpCFqgC+WGIzCHDKm5+uLU8MwAzwlFOZ11b2q/rR8Mf/eYczGtKwgvq8HBPx+QFbc3imAYQVhx1Jf2nvcMSD+2x17L+PBoRhvaTRmgcS3+iiB2FBI99OT3qbuIIGsG5X168kbFQH3spxYHUa2Yuik2izLDk+Mki1abBRAyC0cK47EDEtRSc8qp6BbHu4wd2YFYuZ7kYkVcoeB2aSzpKkzVQ3cNy5RU0ovTVxz4asAHUJ0kAzvX17MVsMjQ99rAN0q3WIJufuXAiY8UcSvInRxGIcnPiueETdldin8Cx5ww1yORCHRd41tYEHsGld2nKd2TaE5mGmHiYzFKKqnPFnR8ckY/B7VDslP5HdUkX6V7/aoyKMzh6X99J0PDU51A0dq0aWn6O7J2bBAShj+rjqPT1hljqGoQ37BDDhwD5ab/0Ps2JrrteOiCJbS/KyAFbZ9p6UF7SQZTFzLqaHqZVn0Qz1vgj7PhlSgvBPfGblX5GLUQzAvE8bhh9Xm2RUFTwbwem9rMXJT5hT59GdNXCPmHPpDnSBs0JCi4fB0LVLQgvpnUewkhVVgPA3v6YWBK3JdU3kb6no561XwJ5u0H+TEgXS3hL1qxnnDKlnMVgbf+DG4P0GU0ManBNNM6deXmUks5/DgO4xM2W5EWbCO0+qmGre+c9c+WBt8eflMI+HPjSIcdWeyUauO76+6tesHzIEwTTGYAkMB4KA581Ct5LTYuzv2SA2PS16VflU8mlcn2mya0sDBWwQWmyxct73Dn8NQk9OcZuk9hbBDGjIEl8wHZ161zxhexR4fU51/yDtaGx+E6usezdPVceX8GRNXoLPWWjyCcUeJ19Zax6eL8/nMr70vL03u3nNGeiBFOL4rg5EHSENVFARTAJU2gYZFo6WauXd9N711W2WST5JXOA==")
from spire.presentation import *
from spire.presentation.common import *
import os

# 设置文件夹路径
folder_path = r"E:\PPT_Home"
output_root = r"E:\PPT_Signal_Page"

# 获取文件夹中所有文件的列表
files = os.listdir(folder_path)

# 遍历每个文件
for file_name in files:
    if file_name.endswith(".ppt") or file_name.endswith(".pptx"):
        file_path = os.path.join(folder_path, file_name)
        print(f"Processing file: {file_name}")
        
        # 创建Presentation对象
        presentation = Presentation()
        
        # 加载演示文件
        presentation.LoadFromFile(file_path)
        
        # 获取PPT中的页数
        slide_count = len(presentation.Slides)
        print(f"Total number of slides: {slide_count}")
        
        # 创建以原始文件名命名的输出文件夹
        output_folder_name = os.path.splitext(file_name)[0]
        output_folder_path = os.path.join(output_root, output_folder_name)
        os.makedirs(output_folder_path, exist_ok=True)
        
        # 遍历每个幻灯片
        for i in range(slide_count):
            slide = presentation.Slides[i]
            
            # 构造输出文件名
            slide_filename = f"{output_folder_name}_slide{i + 1}.pptx"
            output_file_path = os.path.join(output_folder_path, slide_filename)
            
            # 保存当前幻灯片到文件
            slide.SaveToFile(output_file_path, FileFormat.Pptx2019)
        
        # 释放Presentation对象
        presentation.Dispose()

print("All files processed.")

