import re

text = "Это более сложный пример: строка с пунктуацией и числами (123)!"


text = re.sub(r'[^\w\s]', '', text.lower())

print(text)
