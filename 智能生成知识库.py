import os
import openai
import pandas as pd
from docx import Document

openai.api_key = 'sk-SafMVoQpW4bgMdmL9826E890BcEa45D89fA7E954EdF838B9'
openai.api_base = 'https://chatapi.asiainfo-sec.com/v1'

def read_word_file(file_path):
    doc = Document(file_path)
    content = []
    for para in doc.paragraphs:
        content.append(para.text)
    return '\n'.join(content)

def summarize_content(content):
    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=[
            {"role": "system", "content": "你是一个帮助用户梳理文档内容的助手。"},
            {"role": "user", "content": f"请阅读以下内容，并根据语境和上下文生成一组问题和答案，如果找到了问题没有找到答案，你就根据问题生成一个答案。确保每个问题以“问题：”开头，答案以“答案：”开头。\n\n{content}"}
        ]
    )
    return response.choices[0].message['content']

def process_directory(directory):
    data = {'question': [], 'answer': []}
    for filename in os.listdir(directory):
        if filename.startswith("~$"):
            continue
        if filename.endswith(".docx"):
            file_path = os.path.join(directory, filename)
            content = read_word_file(file_path)
            qa_text = summarize_content(content)
            qa_pairs = qa_text.split('\n')
            current_question = None
            for pair in qa_pairs:
                if pair.startswith("问题："):
                    current_question = pair[len("问题："):].strip()
                elif pair.startswith("答案：") and current_question:
                    answer = pair[len("答案："):].strip()
                    data['question'].append(current_question)
                    data['answer'].append(answer)
                    current_question = None
    return data

def save_to_csv(data, output_file):
    df = pd.DataFrame(data)
    df.to_csv(output_file, index=False)

if __name__ == "__main__":
    input_directory = "D:\\Users\\liuya\\download\\shell"
    output_csv = "output.csv"
    data = process_directory(input_directory)
    save_to_csv(data, output_csv)
