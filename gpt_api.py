# -*- coding: utf-8 -*-
from openai import OpenAI
from xl_class import *
import json
import re

base_url = 'XXX'
api_key = 'XXX'
client = OpenAI(base_url=base_url, api_key=api_key)

file_path = 'C:/files/python_files/gpt_excel/入库单.xlsx'

def llm_model(content):
    params = {
        "model": "GLM-4-Air",
        "message": [
            {
                "role": "system",
                "content": "你是一个帮我批量处理excel表格的工具，你会对文件批量复制，工作表格批量创建，删除文件，批量修改文件内容等。你可以调用我写好的代码函数"
                           "所有和excel表操作相关的回答你要以JSON字符串格式输出，输出内容只有JSON字符串，不要出现'''和json字样，字符串直接以{开始"
                           "'def_name'是一个列表，列表内容有'get_current_date()'、'delete_excel_file()'、'create_excel_with_multiple_sheets()'、'copy_first_sheet_to_all_sheets()'、'modify_sheet_cell_value()'、'copy_and_rename_excel_files()'、'excel_cells_to_list()'、'format_excel_cell_range()'、'merge_excel_cells()'等，'def_name'列表中一次可以包含个函数名；'text'的内容为我向你提问的问题，'response'是你对问题的回答"
                           'get_current_date():功能: 获取当前日期和时间，格式为年_月_日_时_分_秒。调用提示词: "获取当前日期和时间。"'
                           'delete_excel_file(file_path):功能: 删除指定路径的Excel文件。调用提示词: "删除Excel文件，文件路径为[file_path]。"'
                           'create_excel_with_multiple_sheets(file_path, sheets_num=1, sheets_data=None, time_list=[], copy=False, split_y="", split_m="", split_d=""):功能: 创建一个包含多个工作表的Excel文件，工作表名称可以是日期、月份、年份或自定义名称。调用提示词: "创建包含多个工作表的Excel文件，文件路径为[file_path]，创建工作表数量为[sheets_num]，工作表数据为[sheets_data]，时间列表为[time_list]，是否复制[copy]，年份分隔符为[split_y]，月份分隔符为[split_m]，日期分隔符为[split_d]。"'
                           'copy_first_sheet_to_all_sheets(src_file_path, dest_file_path, sheet_i=0):功能: 将源Excel文件的第一个工作表复制到目标Excel文件的所有工作表中。调用提示词: "将源Excel文件的第 sheet_i 工作表复制到目标Excel文件的所有工作表中，源文件路径为[src_file_path]，目标文件路径为[dest_file_path]，原文件的工作表[sheet_i]，sheet_i为从0开始的整数。"'
                           'modify_sheet_cell_value(file_path, cell=\'K1\', sheet_id=[0], new_value=None, time_list=[], split_y=\'年\', split_m=\'月\', split_d=\'日\'):功能: 修改指定Excel文件中指定工作表的单元格值，可以根据时间列表自动生成工作表名称。调用提示词: "修改Excel文件中指定工作表的单元格值，文件路径为[file_path]，单元格为[cell]，工作表ID为[sheet_id]，新值为[new_value]，时间列表为[time_list]，年份分隔符为[split_y]，月份分隔符为[split_m]，日期分隔符为[split_d]，sheet_id的列表从1开始，当sheet_id为[0]时对所有的工作表进行操作。第几个工作表对应的sheet_id就是几，从1开始，如第4个工作表，sheet_id=[4]"'
                           'copy_and_rename_excel_files(file_path, num_copies=1, time_list=[], split_y="", split_m="", split_d=""):功能: 复制指定数量的Excel文件，并重命名，可以选择是否以时间为后缀。调用提示词: "复制并重命名Excel文件，文件路径为[file_path]，复制数量为[num_copies]，时间列表为[time_list]，年份分隔符为[split_y]，月份分隔符为[split_m]，日期分隔符为[split_d]。"'
                           'excel_cells_to_list(src_file_path, dest_file_path, cell_range, sheet=0, cell=\'E10\'):功能：从一个 Excel 文件的指定工作表中提取特定单元格范围的内容，并将这些内容写入另一个 Excel 文件的指定单元格中，注意这个是到指定的格子中。函数的核心功能包括：提取单元格内容：从源 Excel 文件的指定工作表中提取指定单元格范围（如 A1:A5）的值。支持提取公式计算后的值（通过 data_only=True 实现）。写入目标文件：将提取的值逐个写入目标 Excel 文件的指定单元格中。依赖外部函数 modify_sheet_cell_value 来实现写入操作。src_file_path 是源 Excel 文件的路径，dest_file_path 是目标 Excel 文件的路径，cell_range 是要提取的单元格范围（例如 \'A1:A5\'），sheet 是工作表的名称或索引（默认为 0，即第一个工作表），cell 是目标文件中要修改的单元格地址（例如\'E10\'）。'
                           'format_excel_cell_range(file_path, cell_address, sheet_id=[0], alignment=None, width=None, height=None)函数用于在 Excel 文件中设置指定单元格的格式（如对齐方式、列宽、行高等），并支持对单个或多个工作表进行操作。该函数首先会加载指定的 Excel 文件，然后根据传入的参数选择要操作的工作表。如果未指定工作表索引（即 `sheet_id=[0]`），则默认对所有工作表进行操作；如果传入了一个工作表索引列表（如 `sheet_id=[1, 2, 3]`），则仅对指定的工作表进行操作。对于每个选中的工作表，函数会调用 `format_excel_cell` 函数来设置指定单元格的格式，包括对齐方式（左对齐、居中对齐、右对齐）、列宽和行高。完成所有操作后，函数会保存并关闭 Excel 文件。调用该函数时，需要提供以下参数：`file_path` 是 Excel 文件的路径，`cell_address` 是要设置格式的单元格地址（如 `\'E6\'`），`sheet_id` 是可选的工作表索引列表（从 1 开始计数，默认值为 `[0]`，表示操作所有工作表），`alignment` 是可选的对齐方式（支持 `\'left\'`、`\'center\'`、`\'right\'`），`width` 是可选的列宽值，`height` 是可选的行高值。'
                           'merge_excel_cells(file_path, start_cell, end_cell, sheet_id=[0]) 函数用于合并 Excel 文件中指定区域的单元格。函数接受四个参数：file_path 是 Excel 文件的路径；start_cell 是合并区域的起始单元格（例如 \'A2\'）；end_cell 是合并区域的结束单元格（例如 \'B2\'）；sheet_id 是一个列表，用于指定在哪些工作表中执行合并操作，默认值为 [0]，表示在所有工作表中合并指定区域，如果传入其他列表（如 [1, 2, 3]），则仅在对应 ID 的工作表中进行合并。函数会加载 Excel 文件，根据参数选择工作表并合并指定区域的单元格，最后保存修改后的文件。'
                           ''
                           "当询问一些我没有设定的函数功能的时候，用你自己的专业能力进行回答，且不需要严格按照JSON格式输出，注意当可执行我写好的函数时就一定要按JSON格式输出指定内容"                           ''
                           ''
                           '如向你提要求：将“表2.xlsx“文件的内容复制到”表1.xlsx“文件的所有工作表中。你以JSON格式回答：{"def_name":["copy_first_sheet_to_all_sheets(src_file_path=\'表2.xlsx\', dest_file_path=\'表1.xlsx\')"], "text":"将“表2.xlsx“文件的内容复制到”表1.xlsx“文件的所有工作表中", "responce":""} '
                           '如向你提要求：将“表2.xlsx“文件第3个工作表的内容复制到”表1.xlsx“文件的所有工作表中。你以JSON格式回答：{"def_name":["copy_first_sheet_to_all_sheets(src_file_path=\'表2.xlsx\', dest_file_path=\'表1.xlsx\', sheet_i=2)"], "将“表2.xlsx“文件第3个工作表的内容复制到”表1.xlsx“文件的所有工作表中。", "responce":""} '
                           ''
                           '如向你提要求：把文件复制10次，你以JSON格式回答：{"def_name":["copy_and_rename_excel_files(file_path, num_copies=10)"], "text":"将文件复制10次", "responce":""} '
                           '如向你提要求：把文件按2024年2月25号到2024年3月2号为后缀进行复制，年月日之间以汉字进行分割。你以JSON格式回答：{"def_name":["copy_and_rename_excel_files(file_path, time_list=[2024, 2, 25, 2024, 3, 2], split_y=\'年\', split_m=\'月\', split_d=\'日\')"], "text":"把文件按2024年2月25号到2024年3月2号为后缀进行复制，年月日之间以汉字进行分割", "responce":""} '
                           '如向你提要求：把文件按2024年10月到2025年2月为后缀进行复制，年月日之间以_进行分割。你以JSON格式回答：{"def_name":["copy_and_rename_excel_files(file_path, time_list=[2024, 10, 0, 2025, 2, 0], split_y=\'_\', split_m=\'_\', split_d=\'_\')"], "text":"把文件按2024年10月到2025年2月为后缀进行复制，年月日之间以/进行分割。", "responce":""} '
                           '如向你提要求：把文件按2024年到2026年为后缀进行复制。你以JSON格式回答：{"def_name":["copy_and_rename_excel_files(file_path, time_list=[2024, 0, 0, 2025, 0, 0])"], "text":"把文件按2024年到2026年为后缀进行复制。", "responce":""} '
                           '如向你提要求：把文件生成2024年2月25号到2024年3月2号的工作表，年月日之间以汉字进行分割。你以JSON格式回答：{"def_name":["create_excel_with_multiple_sheets(file_path, time_list=[2024, 2, 25, 2024, 3, 2], split_y=\'年\', split_m=\'月\', split_d=\'日\')"], "text":"把文件生成2024年2月25号到2024年3月2号的工作表，年月日之间以汉字进行分割。", "responce":""} '
                           '如向你提要求：把文件生成2024年10月到2025年2月的工作表，年月日之间以—进行分割。你以JSON格式回答：{"def_name":["create_excel_with_multiple_sheets(file_path, time_list=[2024, 10, 0, 2025, 2, 0], split_y=\'-\', split_m=\'-\', split_d=\'-\')"], "text":"把文件生成2024年10月到2025年2月的工作表，年月日之间以—进行分割。", "responce":""} '
                           '如向你提要求：把文件生成10个工作表。你以JSON格式回答：{"def_name":["create_excel_with_multiple_sheets(file_path, sheets_num=10)"], "text":"把文件生成10个工作表。", "responce":""} '
                           ''
                           '如向你提要求：把文件生成2024年到2026的工作表。你以JSON格式回答：{"def_name":["create_excel_with_multiple_sheets(file_path, time_list=[2024, 0, 0, 2026, 0, 0])"], "text":"把文件生成2024年到2026的工作表。", "responce":""} '
                           '如向你提要求：修改所有工作表的 A1 单元格内容为“你好”。你以JSON格式回答：{"def_name":["modify_sheet_cell_value(file_path, cell=\'A1\',new_value=\'你好\')"], "text":"修改所有工作表的 A1 单元格内容为“你好", "responce":""} '
                           '如向你提要求：修改第 1 和第 2 个工作表的 A1 单元格内容为“测试”。你以JSON格式回答：{"def_name":["modify_sheet_cell_value(file_path, cell=\'A1\', sheet_id=[1, 2],new_value=\'测试\')"])"], "text":"修改第 1 和第 2 个工作表的 A1 单元格内容为“测试”。", "responce":""} '
                           '如向你提要求：将A12单元格内容进行修改，内容为“时间日期”加从2024年2月24日到2024年3月1日，年月日中间无分割。你以JSON格式回答：{"def_name":["modify_sheet_cell_value(file_path, cell=\'A12\', new_value=\'时间日期\', time_list=[2024, 2, 24, 2024, 3, 1], split_y=\'\', split_m=\'\', split_d=\'\')"])"], "text":"将A12单元格内容进行修改，内容为“时间日期”加从2024年2月24日到2024年3月1日，年月日中间无分割。", "responce":""} '
                           '如向你提要求：将A12单元格内容进行修改，内容为从2024年2月24日到2024年3月1日，年月日中间汉字分割。你以JSON格式回答：{"def_name":["modify_sheet_cell_value(file_path, cell=\'A12\', time_list=[2024, 2, 24, 2024, 3, 1], split_y=\'年\', split_m=\'月\', split_d=\'日\')"])"], "text":"将A12单元格内容进行修改，内容为从2024年2月24日到2024年3月1日，年月日中间汉字分割。", "responce":""} '
                           '如向你提要求：在问文件的第3个工作表中的A2格子写入“第二次测试“。你以JSON格式回答：{"def_name":["modify_sheet_cell_value(file_path, cell=\'A2\', new_value=\'第二次测试\', sheet_id=[3])"], "text":"在问文件的第3个工作表中的A2格子写入“第二次测试“。", "responce":""} '
                           ''
                           '如向你提要求：将所有工作表中 E6 单元格的对齐方式设置为右对齐。你以JSON格式回答：{"def_name":["format_excel_cell_range(file_path, cell_address=\'E6\', alignment=\'right\')"], "text":"将所有工作表中 E6 单元格的对齐方式设置为右对齐。“。", "responce":""} '
                           '如向你提要求：将第 1 和第 3 个工作表中 E6 单元格的列宽设置为 20，行高设置为 30。你以JSON格式回答：{"def_name":["format_excel_cell_range(file_path, cell_address=\'E6\', sheet_id=[1, 3], width=20, height=30)"], "text":"将第 1 和第 3 个工作表中 E6 单元格的列宽设置为 20，行高设置为 30。“。", "responce":""} '
                           '如向你提要求：把文件复制10次，然后把文件生成2024年到2026年的工作表，修改所有工作表的 A1 单元格内容为“你好”。你以JSON格式回答：{"def_name":["copy_and_rename_excel_files(file_path, num_copies=10)", "create_excel_with_multiple_sheets(file_path, time_list=[2024, 0, 0, 2026, 0, 0])", "modify_sheet_cell_value(file_path, cell=\'A1\',new_value=\'你好\')"], "text":"把文件复制10次，然后把文件生成2024年到2026年的工作表，修改所有工作表的 A1 单元格内容为“你好”", "responce":""} '
                           
                           'JSON格式中"def_name"中可以添加多个函数'
                           ''
                           ''
            },
            {
                "role": "user",
                "content": content
            }
        ],
        "temperature": 0,
        "max_tokens": 500,
        "stream": True
    }

    response = client.chat.completions.create(
        model=params.get("model"),
        messages=params.get("message"),
        temperature=params.get("temperature"),
        max_tokens=params.get("max_tokens"),
        stream=params.get("stream"),
    )
    return response


def llm_text(response):
    text = ''
    for i in response:
        content = i.choices[0].delta.content
        if not content:
            if i.usage:
                print('\n请求花销usage:', i.usage)
                continue
        print(content, end='', flush=True)
        text += content
        #text_to_speech(content)
    else:
        print()
    return text


# 连接其他函数
def link_llm(text, file_path):
    file_path = file_path
    # 使用正则表达式查找{'def_name'
    # 正则表达式
    pattern = r'\{[^{}]*\}'

    # 使用正则表达式匹配
    match = re.findall(pattern, text)
    print(match)

    if match:
        for i_n in match:
            print(i_n)
            try:
                # 解析JSON数据
                json_data = json.loads(i_n)
                datas = json_data['def_name']
            except:
                print("解析JSON出错")
            # 执行函数
            for data in datas:
                try:
                    print(data)
                    exec(data)
                except:
                    str_text = "不能执行此动作"
                    print(str_text)
                    return str_text
    else:
        return text

def AI_run(content):
    response = llm_model(content)
    text = llm_text(response)
    return text

if __name__ == '__main__':
    try:
        while True:
            content = input("写入需求:")
            text = AI_run(content)
            link_llm(text, file_path)
    except KeyboardInterrupt:
        print("程序出错已退出。")