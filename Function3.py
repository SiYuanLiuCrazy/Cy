from pathlib import Path
from openai import OpenAI
from spire.presentation import *
License.SetLicenseKey("3vgBAP6gsMpt19Lp5/dLFp/xNXCyz2vrwWXv6RZ9dhyEE6o0AFXtRmuqeHo1LxSwxYqobFgqbUkY/8y0oqHns0AgcGdhUi/gxW3rgxU03m9Je53kYu9SVmpp/9SEEkUhzJ08Ir3+/014o07W0UXzj63zhaQjMxz5Lmx+U08qabOyQmGmnbolIYd70slekKohJLucSw9KL/OxAzIhBdI1+IaElvL2jUsl1w0zGfyVVG56vcENaCh23rDBNUdkTHccDEOW9L3Qa1DZtUpJadmED0scQbbsDHtrJ8sR03yxa2stkpjXUYcVyE9xZphrvWHYvvxAt1UOFOpYxzSwvxS7TIFEhXQBzcHRpEidD1BnBQgIMvo9tdqQQCKCx4k6r9b0++Xs0vwJGUyJGg5a2hVtYnTcPp4ZfPZP8U1XI3sEaVr7BKy44sN4aFRUcOg8E5BVsXJdZQZnWxm0ZPFRazyGL0Ki6Vi0FchpZ1zlqjR15y8GrdZqd2Y4TqYp6gotvPdSH9uirPRDJTPvAvemo491EoPQo83odCDE2MoYu6dm2zNJWc0/ac6sPCt2WR5p9gibTkyjpzW9t+d8YVbjszXSecM+gi5WIBVg98flw/RVhsNg6kxqQ9EqyZXvYxXBCPoygWjJRlC8ETyeMkQy9Cw3t1nUCeFjrpJlWQMuavMa79tLT0anjYYbbl9k/j73i5lG4tmSpCFqgC+WGIzCHDKm5+uLU8MwAzwlFOZ11b2q/rR8Mf/eYczGtKwgvq8HBPx+QFbc3imAYQVhx1Jf2nvcMSD+2x17L+PBoRhvaTRmgcS3+iiB2FBI99OT3qbuIIGsG5X168kbFQH3spxYHUa2Yuik2izLDk+Mki1abBRAyC0cK47EDEtRSc8qp6BbHu4wd2YFYuZ7kYkVcoeB2aSzpKkzVQ3cNy5RU0ovTVxz4asAHUJ0kAzvX17MVsMjQ99rAN0q3WIJufuXAiY8UcSvInRxGIcnPiueETdldin8Cx5ww1yORCHRd41tYEHsGld2nKd2TaE5mGmHiYzFKKqnPFnR8ckY/B7VDslP5HdUkX6V7/aoyKMzh6X99J0PDU51A0dq0aWn6O7J2bBAShj+rjqPT1hljqGoQ37BDDhwD5ab/0Ps2JrrteOiCJbS/KyAFbZ9p6UF7SQZTFzLqaHqZVn0Qz1vgj7PhlSgvBPfGblX5GLUQzAvE8bhh9Xm2RUFTwbwem9rMXJT5hT59GdNXCPmHPpDnSBs0JCi4fB0LVLQgvpnUewkhVVgPA3v6YWBK3JdU3kb6no561XwJ5u0H+TEgXS3hL1qxnnDKlnMVgbf+DG4P0GU0ManBNNM6deXmUks5/DgO4xM2W5EWbCO0+qmGre+c9c+WBt8eflMI+HPjSIcdWeyUauO76+6tesHzIEwTTGYAkMB4KA581Ct5LTYuzv2SA2PS16VflU8mlcn2mya0sDBWwQWmyxct73Dn8NQk9OcZuk9hbBDGjIEl8wHZ161zxhexR4fU51/yDtaGx+E6usezdPVceX8GRNXoLPWWjyCcUeJ19Zax6eL8/nMr70vL03u3nNGeiBFOL4rg5EHSENVFARTAJU2gYZFo6WauXd9N711W2WST5JXOA==")
from spire.presentation import *
from spire.presentation.common import *
import os
import shutil

directory_path = 'E:\\PPT_PPTtopdf'
output_folder = r'E:\\PPT_pdf'
output_filename = 'Inverse.pdf'
output_path = os.path.join(output_folder, output_filename)
# 获取文件夹中所有的 pptx 文件
ppt_files = [os.path.join(directory_path, file) for file in os.listdir(directory_path) if file.endswith('.pptx')]

# 创建Presentation类的对象
presentation = Presentation()

# 加载演示文件
presentation.LoadFromFile(ppt_files[0])

# 将演示文件转换为PDF并保存
presentation.SaveToFile(output_path, FileFormat.PDF)
presentation.Dispose()

file_name = "E:\PPT_pdf\Inverse.pdf"

client = OpenAI(
    api_key = "sk-qwEg1ba3qj8JNsUtErVmlDefkK9uRbW60MyEPHeO1d4Lt67y",
    base_url = "https://api.moonshot.cn/v1",
)
 
# xlnet.pdf 是一个示例文件, 我们支持 pdf, doc 以及图片等格式, 对于图片和 pdf 文件，提供 ocr 相关能力
file_object = client.files.create(file=Path(file_name), purpose="file-extract")
 
# 获取结果
# file_content = client.files.retrieve_content(file_id=file_object.id)
# 注意，之前 retrieve_content api 在最新版本标记了 warning, 可以用下面这行代替
# 如果是旧版本，可以用 retrieve_content
file_content = client.files.content(file_id=file_object.id).text
 
# 把它放进请求中
messages = [
    {
        "role": "system",
        "content": "你是 Kimi，由 Moonshot AI 提供的人工智能助手，你更擅长中文和英文的对话。你会为用户提供安全，有帮助，准确的回答。同时，你会拒绝一切涉及恐怖主义，种族歧视，黄色暴力等问题的回答。Moonshot AI 为专有名词，不可翻译成其他语言。",
    },
    {
        "role": "system",
        "content": file_content,
    },
    {"role": "user", "content": "请给我上传的文件起一个10个字以内的名字，如果文件中有主标题，那么选择文件中的主标题，如果没有主标题，请用副标题，如果没有副标题请自己对内容进行总结，注意要精简、准确，请不要自己过度解读，不要添加文件中不存在的内容。你的回复只有名字，请不要有其它内容"},
]
 
# 然后调用 chat-completion, 获取 Kimi 的回答
completion = client.chat.completions.create(
  model="moonshot-v1-32k",
  messages=messages,
  temperature=0.3,
)
print(completion.choices[0].message.content)
# 给PPT文件重新命名
name = completion.choices[0].message.content + '.pptx'

Rename_folder = r'E:\\PPT_Rename'
new_file_path = os.path.join(Rename_folder, name)
# 将文件复制到新位置并使用新名称
shutil.copy(ppt_files[0], new_file_path)


