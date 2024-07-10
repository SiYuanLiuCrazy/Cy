from pathlib import Path
from openai import OpenAI
 
client = OpenAI(
    api_key = "sk-qwEg1ba3qj8JNsUtErVmlDefkK9uRbW60MyEPHeO1d4Lt67y",
    base_url = "https://api.moonshot.cn/v1",
)
 
# xlnet.pdf 是一个示例文件, 我们支持 pdf, doc 以及图片等格式, 对于图片和 pdf 文件，提供 ocr 相关能力
file_object = client.files.create(file=Path("E:\PPT_ReadyToMerge\清华-滴滴项目结题报告_20240705_slide_1.pptx"), purpose="file-extract")
 
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
    {"role": "user", "content": "请给 E:\PPT_ReadyToMerge\清华-滴滴项目结题报告_20240705_slide_1.pptx 起一个7个字以内的名字，请不要自己过度解读，以文件中的内容为准"},
]
 
# 然后调用 chat-completion, 获取 Kimi 的回答
completion = client.chat.completions.create(
  model="moonshot-v1-32k",
  messages=messages,
  temperature=0.3,
)
 
print(completion.choices[0].message)