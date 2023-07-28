import os
import re
import openpyxl

# 创建Excel文件
wb = openpyxl.Workbook()
sheet = wb.active

# 表头
sheet.append(['期刊', '年份', '期号', '文章标题', '发表时间(如非本期首发）', '作者', '全文', '起始页码', '结束页码'])

# 获取当前目录下的所有.txt文件
files = [file for file in os.listdir('pre') if file.endswith('.txt')]

for file in files:
    # 提取期刊、年份和期号
    journal, year, issue_with_ext = file.split('_')[:3]
    issue = os.path.splitext(issue_with_ext)[0]

    # 打开TXT文档
    txt_path = os.path.abspath(os.path.join('pre', file))
    with open(txt_path, 'r', encoding='utf-8') as txt_file:
        # 初始化变量
        title = ''
        publish_time = ''
        author = ''
        full_text = ''
        start_page = ''
        end_page = ''

        # 保存段落的数组
        paragraphs = []

        # 初始化一个标志，指示最后一个段落是否包含全文内容
        last_paragraph_has_full_text = True

        # 遍历文档内容
        for line in txt_file:
            # 这一步我加空行的目的是保持一致性，这样可以让文章分成若干段落
            print(666)
            text = line.strip()

            # 如果遇到空行，则保存当前的段落
            if not text:
                if paragraphs:
                    # 提取全文
                    for paragraph in paragraphs[3:]:
                        full_text += paragraph + '\n'

                    # 将提取的信息写入Excel
                    sheet.append([journal, year, issue, title, publish_time, author, full_text, start_page, end_page])

                    # 重置变量
                    title = ''
                    publish_time = ''
                    author = ''
                    full_text = ''
                    start_page = ''
                    end_page = ''

                # 清空段落数组
                paragraphs = []
                # 标记最后一个段落没有全文内容，等待后续处理
                last_paragraph_has_full_text = False
            else:
                # 添加非空行到段落数组
                paragraphs.append(text)

                # 如果是新的段落，则提取标题、发表时间、作者和页码信息
                if not title:
                    title = text
                elif not publish_time:
                    match = re.match(r'\d{4}年\d{1,2}月', text)  # 正则匹配，这一步目的是为了过滤
                    if match:
                        publish_time = text
                    else:
                        publish_time = '无'
                        author = text
                elif not author:
                    author = text
                elif not start_page:
                    page_text = text
                    page_match = re.match(r'页码(\d+-\d+)', page_text)  # 这一步我觉得可有可无，主要是防止你预处理失误乱码
                    if page_match:
                        start_page, end_page = page_match.group(1).split('-')
                        # 标记最后一个段落包含全文内容，防止它被覆盖
                        last_paragraph_has_full_text = True

        # 加一个保险，感觉可有可无
        if paragraphs and last_paragraph_has_full_text:
            # 提取全文
            for paragraph in paragraphs[3:]:
                full_text += paragraph + '\n'

            # 将提取的信息写入Excel
            sheet.append([journal, year, issue, title, publish_time, author, full_text, start_page, end_page])

# 保存Excel文件
wb.save('提取结果.xlsx')
