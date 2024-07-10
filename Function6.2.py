from spire.presentation import *
License.SetLicenseKey("3vgBAP6gsMpt19Lp5/dLFp/xNXCyz2vrwWXv6RZ9dhyEE6o0AFXtRmuqeHo1LxSwxYqobFgqbUkY/8y0oqHns0AgcGdhUi/gxW3rgxU03m9Je53kYu9SVmpp/9SEEkUhzJ08Ir3+/014o07W0UXzj63zhaQjMxz5Lmx+U08qabOyQmGmnbolIYd70slekKohJLucSw9KL/OxAzIhBdI1+IaElvL2jUsl1w0zGfyVVG56vcENaCh23rDBNUdkTHccDEOW9L3Qa1DZtUpJadmED0scQbbsDHtrJ8sR03yxa2stkpjXUYcVyE9xZphrvWHYvvxAt1UOFOpYxzSwvxS7TIFEhXQBzcHRpEidD1BnBQgIMvo9tdqQQCKCx4k6r9b0++Xs0vwJGUyJGg5a2hVtYnTcPp4ZfPZP8U1XI3sEaVr7BKy44sN4aFRUcOg8E5BVsXJdZQZnWxm0ZPFRazyGL0Ki6Vi0FchpZ1zlqjR15y8GrdZqd2Y4TqYp6gotvPdSH9uirPRDJTPvAvemo491EoPQo83odCDE2MoYu6dm2zNJWc0/ac6sPCt2WR5p9gibTkyjpzW9t+d8YVbjszXSecM+gi5WIBVg98flw/RVhsNg6kxqQ9EqyZXvYxXBCPoygWjJRlC8ETyeMkQy9Cw3t1nUCeFjrpJlWQMuavMa79tLT0anjYYbbl9k/j73i5lG4tmSpCFqgC+WGIzCHDKm5+uLU8MwAzwlFOZ11b2q/rR8Mf/eYczGtKwgvq8HBPx+QFbc3imAYQVhx1Jf2nvcMSD+2x17L+PBoRhvaTRmgcS3+iiB2FBI99OT3qbuIIGsG5X168kbFQH3spxYHUa2Yuik2izLDk+Mki1abBRAyC0cK47EDEtRSc8qp6BbHu4wd2YFYuZ7kYkVcoeB2aSzpKkzVQ3cNy5RU0ovTVxz4asAHUJ0kAzvX17MVsMjQ99rAN0q3WIJufuXAiY8UcSvInRxGIcnPiueETdldin8Cx5ww1yORCHRd41tYEHsGld2nKd2TaE5mGmHiYzFKKqnPFnR8ckY/B7VDslP5HdUkX6V7/aoyKMzh6X99J0PDU51A0dq0aWn6O7J2bBAShj+rjqPT1hljqGoQ37BDDhwD5ab/0Ps2JrrteOiCJbS/KyAFbZ9p6UF7SQZTFzLqaHqZVn0Qz1vgj7PhlSgvBPfGblX5GLUQzAvE8bhh9Xm2RUFTwbwem9rMXJT5hT59GdNXCPmHPpDnSBs0JCi4fB0LVLQgvpnUewkhVVgPA3v6YWBK3JdU3kb6no561XwJ5u0H+TEgXS3hL1qxnnDKlnMVgbf+DG4P0GU0ManBNNM6deXmUks5/DgO4xM2W5EWbCO0+qmGre+c9c+WBt8eflMI+HPjSIcdWeyUauO76+6tesHzIEwTTGYAkMB4KA581Ct5LTYuzv2SA2PS16VflU8mlcn2mya0sDBWwQWmyxct73Dn8NQk9OcZuk9hbBDGjIEl8wHZ161zxhexR4fU51/yDtaGx+E6usezdPVceX8GRNXoLPWWjyCcUeJ19Zax6eL8/nMr70vL03u3nNGeiBFOL4rg5EHSENVFARTAJU2gYZFo6WauXd9N711W2WST5JXOA==")
from spire.presentation.common import *
from spire.presentation import *
import os
# 定义文件夹路径
directory_path = 'E:\\PPT_ReadyToMerge'

# 获取文件夹中所有的 pptx 文件
ppt_files = [os.path.join(directory_path, file) for file in os.listdir(directory_path) if file.endswith('.pptx')]

# 定义PPT文件夹路径
source_folder = r'E:\PPT_ReadyToMerge'
output_folder = r'E:\PPT_Merge'
output_filename = input("请输入合并后PPT的文件名（不包括扩展名）：") + '.pptx'
output_path = os.path.join(output_folder, output_filename)
page_ppt = len(ppt_files)
print(page_ppt)
Merge1 = Presentation()
Merge2 = Presentation()
Merge3 = Presentation()
Merge1.LoadFromFile(ppt_files[0])
Merge3.LoadFromFile(ppt_files[3])
print(Merge1.Slides[0])
for i in range(page_ppt-1):
    Merge2.LoadFromFile(ppt_files[i+1])
    slide = Merge2.Slides[0]
    Merge1.Slides.AppendByMaster(slide,Merge3.Masters[0])

# 保存第一个演示文稿
Merge1.SaveToFile(output_path,FileFormat.Pptx2019)
Merge1.Dispose()
Merge2.Dispose()
