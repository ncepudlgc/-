import pandas as pd
import os
import re
from datetime import datetime

# 定义函数 extract_first_occurrence 和 extract_date （保留上一问中的这两个函数）
def extract_first_occurrence(text, target_pattern):
    pattern = r'(?<=。)被告人({0})(?=[，。；？！,\.\?!;\(])'.format(target_pattern)
    recognized_names = set()
    first_occurrence_sentences = []
    native_places = []
    education_levels = []

    
    for match in re.finditer(pattern, text):
        name = match.group(1)
        if name not in recognized_names:
            recognized_names.add(name)
            index = match.start()
            sentence_start = text.rfind('。', 0, index) + 1
            sentence_end = text.find('。', index)
            if sentence_end == -1:
                sentence_end = len(text)
            sentence = text[sentence_start:sentence_end].strip()
            first_occurrence_sentences.append(sentence)

    # 提取被告人后面的名字
    extracted_names = [re.search(r'被告人([\u4E00-\u9FFF]+)', sentence).group(1) for sentence in first_occurrence_sentences]

    # 提取出生日期并推断年龄
    birth_date = []
    for sentence in first_occurrence_sentences:
        birth_match = re.search(r'(生于\d{4}年\d{1,2}月\d{1,2}日)|(\d{4}年\d{1,2}月\d{1,2}日)出生|(\d{4}年\d{1,2}月\d{1,2}日)生|(\d{4})年(\d{1,2})月(\d{1,2})日（农历(\d{1,2})月初(\d{1,2})）出生', sentence)
        if birth_match:
            birth_date_str = birth_match.group()
            date_match = re.search(r'\d{4}年\d{1,2}月\d{1,2}日', birth_date_str)
            if date_match:
                birth_date.append(date_match.group())
            else:
                birth_date.append("无")
        else:
            birth_date.append("无")
    
    # 提取出生地和籍贯
    for sentence in first_occurrence_sentences:
        # Extract birthplace (出生于xxxxx) if present
        hometown ='无'
        birthplace_match = re.search(r'出生于([\u4E00-\u9FFF]+)，', sentence)
        if birthplace_match:
            hometown=birthplace_match.group(1)

        # Extract native place (xxxxx人) if present
        if sentence:
            if not ('人民' in sentence or '人员' in sentence or '工人' in sentence or '故意杀人' in sentence or '故意伤人' in sentence or '辩护人' in sentence):
        # 删除前三个字
                native_place_match = re.search(r'([\u4E00-\u9FFF]+)人', sentence[3:])
                if native_place_match:
                    hometown=native_place_match.group(1)
        # Extract native place from terms like "籍贯"、"籍贯："、"户籍"、"户籍所在地："
        place_match_pre1=re.search(r'(籍贯|户籍)([^，。；？！,\.\?!;\(]+)', sentence)
        if place_match_pre1:
            hometown=place_match_pre1.group(2)

        place_match_pre2=re.search(r'(户籍登记地址|户籍所在地为|户籍地|户籍地址|户籍地为|户籍在|户籍所在地同现住址|户籍所在地及现住址|户籍所在地|户籍所在地及地址|户籍所在地及住址|户籍所在地为)([^，。；？！,\.\?!;\(]+)', sentence)
        if place_match_pre2:
            hometown=place_match_pre2.group(2)

        place_match = re.search(r'(户籍地及居住地：|户籍地和居住地：|户籍:|籍贯:|籍贯：|户籍所在地：|户籍地:|户籍地：|户籍地址:|户籍：|户籍所在地:|户籍地址：|户籍住址：|户籍住址)([^，。；？！,\.\?!;\(]+)', sentence)
        if place_match:
            hometown=place_match.group(2)
        if len(hometown)>2:
            native_places.append(hometown)
        else:
            native_places.append('无')
        # Extract education level
        education_match = re.search(r'(文盲|小学|中学|初中|高中|职高|中专|大专|大学|研究生|博士|硕士)', sentence)
        if education_match:
            education_levels.append(education_match.group())
        else:
            education_levels.append("无")

    return first_occurrence_sentences, extracted_names, birth_date, native_places, education_levels

def extract_date(text):
    date_match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', text)
    if date_match:
        year = int(date_match.group(1))
        month = int(date_match.group(2))
        day = int(date_match.group(3))
        try:
            birth_date_obj = datetime(year, month, day)
            return birth_date_obj
        except ValueError:
            return None
    else:
        return None
# 定义函数 extract_trailing_text （保留上一问中的这个函数）
def extract_trailing_text(text, match_string):
    match_index = text.find(match_string)
    if match_index != -1:
        return text[match_index + len(match_string):]
    else:
        return ""
# 获取pre目录下所有xlsx文件的文件名列表
file_names = [file_name for file_name in os.listdir("pre") if file_name.endswith(".xlsx")]

# 存储所有提取的信息的字典列表
all_data = []

# 处理每个文件
for file_name in file_names:
    print(file_name)
    # 读取当前文件的Excel内容，设置header=None以表示没有标题行
    df = pd.read_excel(os.path.join("pre", file_name), header=None)

    # 处理每一行的数据并提取所需信息
    for index, row in df.iterrows():
        # 获取B列（索引为1）和I列（索引为8）的内容
        text_b = str(row[1])  # 检查B列是否为空
        text_i = str(row[8])  # 检查I列是否为空

        # 在这里执行提取操作并填充J列
        if not pd.isnull(row[9]) or pd.isnull(text_i):
            # J列不为空或I列为空，跳过当前行
            continue
        else:
            # J列为空且I列不为空，提取匹配部分和后面的截断部分，并填充到J列
            extracted_text = extract_trailing_text(text_b, text_i)
            
            # 进一步处理extracted_text
            delete_index = extracted_text.find("一审刑事")
            if delete_index != -1:
                extracted_text = extracted_text[:delete_index]

            df.at[index, 9] = extracted_text

    # 将修改后的DataFrame写回到原始Excel文件，设置header=None以表示没有标题行
    df.to_excel(os.path.join("pre", file_name), index=False, header=None)
    print(file_name)
    # 从第12列（L列）读取全文内容，并将其转换为字符串
    df = pd.read_excel(os.path.join("pre", file_name), usecols="A:C,J,L", skiprows=1)

    # 处理每一行的全文并提取信息
    for index, row in df.iterrows():
        # 从第12列（L列）读取全文内容，并将其转换为字符串
        text = row.iloc[4]  # 第12列的索引是4，因为我们使用了usecols="A:D,L"读取前4列和第12列

        # 检查全文内容是否为空，如果为空则给予默认值空字符串
        text = str(text) if not pd.isnull(text) else ""

        # 在这里执行提取操作
        first_occurrence_sentences, extracted_names, birth_date, native_places, education_levels = extract_first_occurrence(text, r'[\u4E00-\u9FFF]{2,3}')

        # 将提取的信息保存到字典列表中
        for i, (sentence, name, date, native_place, education) in enumerate(zip(first_occurrence_sentences, extracted_names, birth_date, native_places, education_levels)):
            if date == "无":
                age = "无"
            else:
                birth_date_obj = extract_date(date)
                if birth_date_obj:
                    age = datetime.now().year - birth_date_obj.year
                else:
                    age = "无"

            case_info = {
                "案号": row[0],
                "案件名称": row[1],
                "案由": row[3],
                "法院": row[2],
                "被告人": name,
                "出生日期": date,
                "年龄": age,
                "籍贯": native_place,
                "文化程度": education,
                "提取源": sentence
            }
            all_data.append(case_info)

# 将所有提取的信息转换为DataFrame
result_df = pd.DataFrame(all_data)

# 将DataFrame写入到新的Excel文件中
result_df.to_excel("result.xlsx", index=False)
