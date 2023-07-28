import os
import re
import chardet
from openpyxl import Workbook

pattern = r'"deviceUID":"(\w+)","deviceCnName":"(.*?)","deviceType":"(\w+)",'

def print_result(result, output_sheet):
    for building, devices in result.items():
        for device_type, devices_info in devices.items():
            for uid, name in devices_info.items():
                output_sheet.append([building, device_type, name, uid])


output_dict = {}  # 定义输出字典

for filename in os.listdir('.'):
    file_path = os.path.abspath(filename)
    if os.path.isfile(file_path) and not file_path.endswith('.txt') and not os.path.splitext(file_path)[1]:
        new_file_path = os.path.splitext(file_path)[0] + '.txt'
        os.rename(file_path, new_file_path)
        file_path = new_file_path

    if os.path.isfile(file_path) and file_path.endswith('.txt'):
        building = os.path.splitext(filename)[0]
        # print(f'========== {building} ==========')

        result = {}
        for device_type in ["CWP", "CHWP", "CTW", "CWU", "CCS", "CHU"]:
            result[device_type] = {}

        with open(file_path, 'rb') as f:
            content = f.read()
            encoding = chardet.detect(content)['encoding']

        with open(file_path, 'r', encoding=encoding) as f:
            for line in f:
                m = re.search(pattern, line)
                if m:
                    uid = m.group(1)
                    name = m.group(2)
                    device_type = m.group(3)
                    if device_type in ["CWP", "CHWP", "CTW", "CWU", "CCS", "CHU"]:
                        if device_type not in result:
                            result[device_type] = {}
                        result[device_type][uid] = name

        output_dict[building] = result  # 将结果添加到输出字典中

# 创建Excel工作簿
workbook = Workbook()
output_sheet = workbook.active

# 输出结果到Excel
print_result(output_dict, output_sheet)

# 保存Excel文件
workbook.save("output.xlsx")
