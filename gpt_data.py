# -*- coding: utf-8 -*-
from openai import OpenAI
from xl_class import *
import json
import re
import base64

file_path = 'C:/files/python_files/gpt_excel/入库单.xlsx'

# 本地模型名称
model_local = "deepseek-r1:1.5b"



def get_image_base64(image_file):
    """Convert an image file to a base64 string."""
    if image_file is not None:
        encoded_string = base64.b64encode(image_file.read()).decode()  # 编码并转换为可读字符串
        return encoded_string
    return None


def llm_model2(content, model=None, API_key=None, image_file=None):
    if model is None:
        model = "deepseek-v3-volcengine"
    else:
        model = model

    if model == "本地模型":
        base_url = 'http://localhost:11434/v1/'
        api_key = 'ollama'
        model = model_local
    else:
        base_url = 'https://api.mindcraft.com.cn/v1/'
        #base_url = 'https://api.siliconflow.cn/v1'
        api_key = API_key
    client = OpenAI(base_url=base_url, api_key=api_key)

    if image_file is not None:
        params = {
            "model": "Doubao-1.5-vision-pro-32k",
            "message": [
                {
                    "role": "system",
                    "content": "你是一个数据分析AI，对图片中的数据进行提取，分析，总结，归纳。"
                               "你可以绘制图表，当需要绘制图表时，用JSON格式回答，其他情况正常回答"
                               "不要出现python代码"
                               "'def_name'是一个列表，列表内容有'plot_chart()，write_df_to_excel(output_filename, df)'，'text'的内容为我向你提问的问题，'response'是你对问题的回答"
                               "plot_chart(data, chart_type, x_column, y_columns=None, legend_title=None, title='Chart', xlabel=None, ylabel=None,colors=None)函数用于根据给定的数据和参数绘制不同类型的统计图表，支持条形图、多条折线图、饼图和散点图等，接受包括数据框、图表类型、x轴列名、y轴列名列表（对于饼图此参数为None）、图例标题、图表标题、x轴和y轴标签以及颜色在内的多个参数，其中除了数据、图表类型和x轴列名是必需的之外，其余参数均具有默认值以提供灵活性和便捷性，例如你可以通过指定不同的颜色来区分图表中的不同类别或系列，并且可以通过设置图例标题和坐标轴标签来增强图表的可读性和解释力。"
                               "plot_chart的函数，利用Plotly Express和Streamlit库根据提供的数据框、图表类型及其它参数生成饼图，当指定图表类型为\"pie\"时，它使用数据框中的分类信息（x_column）作为饼图各扇区的标签，以及对应数值信息（y_columns的第一个元素）决定各扇区大小，同时允许自定义标题、颜色等样式，并最终通过Streamlit展示生成的饼图。"
                               "绘制饼图时plot_chart函数的使用方法如下：plot_chart(data=data = pd.DataFrame({'Fruits': ['Apple', 'Banana', 'Cherry'],'Sales': [55, 45, 30]}), chart_type='pie', x_column='Fruits', y_columns=['Sales'], title='Fruit Sales Distribution')"
                               ""
                               "如向你提要求：来绘制一个展示销售和支出随年份变化的折线图，其中数据包含'Year'、'Sales'和'Expenses'三列，分别代表年份、销售额和支出额；通过设置chart_type为'line'指定绘制折线图，x_column选择'Year'作为x轴，y_columns包括['Sales', 'Expenses']以同时展示销售与支出两条折线，legend_title定义图例标题为'Indicator'，整体图表标题设为\"Sales and Expenses Over Years\"，并自定义了x轴标签为\"Year\"，y轴标签为\"Amount\"，最后通过colors参数设置了每条折线的颜色分别为'#636EFA'和'#EF553B'，"
                               '你用JSON格式回答：{"def_name":["plot_chart(data=pd.DataFrame({\'Year\': [\'2021\', \'2022\', \'2023\', \'2024\'],\'Sales\': [500, 700, 800, 600],\'Expenses\': [400, 450, 500, 550]}), chart_type=\'line\', x_column=\'Year\', y_columns=[\'Sales\', \'Expenses\'],legend_title=\'Indicator\', title="Sales and Expenses Over Years", xlabel="Year", ylabel="Amount",colors=[\'#636EFA\', \'#EF553B\'])"]}'
                               '当让你生成多个统计图时，可以根据需求写多个JSON字符串，但每一个JSON字符串要严格按我给你的格式书写，不要在一个JSON字符串里写多个"plot_chart"函数'
                               '函数write_df_to_excel(output_filename, df)用于将pandas DataFrame数据简便而高效地导出到Excel文件中，其中output_filename为包含路径的输出文件名称，df则是要导出的数据内容'
                               '向你询将图片数据存入excel表文件中等相关问题，你以JSON格式回答，用JSON格式，格式如下：{"def_name":["write_df_to_excel(output_filename, df=pd.DataFrame({\'Year\': [\'2021\', \'2022\', \'2023\', \'2024\'],\'Sales\': [500, 700, 800, 600],\'Expenses\': [400, 450, 500, 550]})"]}'                        
                               ''


                },
                {
                    "role": "user",
                    "content": [
                        # 使用 base64 编码传输
                        {
                            'type': 'image',
                            'source': {
                                'data': get_image_base64(image_file)
                            },
                        },
                        {
                            'type': 'text',
                            'text': content,
                        },
                    ]
                }

            ],
            "temperature": 0,
            "max_tokens": 8000,
            "stream": True
        }

    else:
        params = {
            "model": model,
            "message": [
                {
                    "role": "system",
                    "content": "你是一个数据分析AI，对数据进行提取，分析，总结，归纳。"
                               "你可以绘制图表，当需要绘制图表时，用JSON格式回答"
                               '在写JSON格式中不要用df代替参数，用JSON格式调用函数时要直接将DataFrame数据写在里面，不要用变量代替。像这样写：data=pd.DataFrame({\'Year\': [\'2021\', \'2022\', \'2023\', \'2024\'],\'Sales\': [500, 700, 800, 600],\'Expenses\': [400, 450, 500, 550]})'
                               '你的输出中不能同时出现JSON字符串和python代码'
                               '你的输出中不能同时出现JSON字符串和python代码'
                               '你的输出中不能同时出现JSON字符串和python代码'
                               ""
                               ""
                               "'def_name'是一个列表，列表内容有'plot_chart()'，'text'的内容为我向你提问的问题，'response'是你对问题的回答"
                               "plot_chart(data, chart_type, x_column, y_columns=None, legend_title=None, title='Chart', xlabel=None, ylabel=None,colors=None)函数用于根据给定的数据和参数绘制不同类型的统计图表，支持条形图、多条折线图、饼图和散点图等，接受包括数据框、图表类型、x轴列名、y轴列名列表（对于饼图此参数为None）、图例标题、图表标题、x轴和y轴标签以及颜色在内的多个参数，其中除了数据、图表类型和x轴列名是必需的之外，其余参数均具有默认值以提供灵活性和便捷性，例如你可以通过指定不同的颜色来区分图表中的不同类别或系列，并且可以通过设置图例标题和坐标轴标签来增强图表的可读性和解释力。"
                               "plot_chart的函数，利用Plotly Express和Streamlit库根据提供的数据框、图表类型及其它参数生成饼图，当指定图表类型为\"pie\"时，它使用数据框中的分类信息（x_column）作为饼图各扇区的标签，以及对应数值信息（y_columns的第一个元素）决定各扇区大小，同时允许自定义标题、颜色等样式，并最终通过Streamlit展示生成的饼图。"
                               "绘制饼图时plot_chart函数的使用方法如下：plot_chart(data=data = pd.DataFrame({'Fruits': ['Apple', 'Banana', 'Cherry'],'Sales': [55, 45, 30]}), chart_type='pie', x_column='Fruits', y_columns=['Sales'], title='Fruit Sales Distribution')"
                               ""
                               "如向你提要求：来绘制一个展示销售和支出随年份变化的折线图，其中数据包含'Year'、'Sales'和'Expenses'三列，分别代表年份、销售额和支出额；通过设置chart_type为'line'指定绘制折线图，x_column选择'Year'作为x轴，y_columns包括['Sales', 'Expenses']以同时展示销售与支出两条折线，legend_title定义图例标题为'Indicator'，整体图表标题设为\"Sales and Expenses Over Years\"，并自定义了x轴标签为\"Year\"，y轴标签为\"Amount\"，最后通过colors参数设置了每条折线的颜色分别为'#636EFA'和'#EF553B'，"
                               '你用JSON格式回答：{"def_name":["plot_chart(data=pd.DataFrame({\'Year\': [\'2021\', \'2022\', \'2023\', \'2024\'],\'Sales\': [500, 700, 800, 600],\'Expenses\': [400, 450, 500, 550]}), chart_type=\'line\', x_column=\'Year\', y_columns=[\'Sales\', \'Expenses\'],legend_title=\'Indicator\', title="Sales and Expenses Over Years", xlabel="Year", ylabel="Amount",colors=[\'#636EFA\', \'#EF553B\'])"]}'
                               '当让你生成多个统计图时，可以根据需求写多个JSON字符串，但每一个JSON字符串要严格按我给你的格式书写，不要在一个JSON字符串里写多个"plot_chart"函数'
                               ''
                               '如果向你提问绘制图表和保存数据之外的一些问题，且涉及到计算、查找、保存等任务，你可以自己决定编写相应的python代码来实现，且要对代码进行一定的解释，如果涉及到保存文件要说明保存到哪个文件里，把文件地址说清楚。'
                               '你编写的python代码里只能使用"numpy", "pandas", "openpyxl", "os", "csv", "math", "random", "json", "re", "time", "copy"库，其他库不要用。'
                               '所有数据数值相关的python代码，在写代码的过程中必须要进行如下步骤：1.提取有效数据行（排除首行空值和最后合计行）；2.筛选所需列并重命名；3.转换数据类型并筛选。步骤必须至少有这三步，可有其他步骤但是这三步必须有。'
                               '所有数据数值相关的python代码，在写代码的过程中必须要进行如下步骤：1.提取有效数据行（排除首行空值和最后合计行）；2.筛选所需列并重命名；3.转换数据类型并筛选。步骤必须至少有这三步，可有其他步骤但是这三步必须有。'
                               '你接收到的数据与pandas的DataFrame格式的数据样式相同，在python代码中就以df参数代替传给你的那部分数据的DataFrame格式，即你写的python代码中将传给你的数据就以pandas的DataFrame格式的df参数代替即可'
                               'python代码要清晰完整，思路明了，一定要写完整引入的库。'
                               '一些简单的数据分析和数据提取问题，可以选择不用python代码，只输出文本和调用JSON字符串中的函数即可'
                               '所有数据相关的要求，在写代码的过程中要对格子的数据类型进行判断，如写代码时出现”a>400“时要判断a是否是int或者是flout类型，保证代码的正常运行'
                               '所有数据相关的要求，在写代码的过程中要对格子的数据类型进行判断，如写代码时出现”a>400“时要判断a是否是int或者是flout类型，保证代码的正常运行'
                               '所有数据数值相关的python代码，在写代码的过程中必须要进行如下步骤：1.提取有效数据行（排除首行空值和最后合计行）；2.筛选所需列并重命名；3.转换数据类型并筛选。步骤必须至少有这三步，可有其他步骤但是这三步必须有。'
                               '你接收到的数据是pandas的DataFrame格式的数据，在python代码中就以df参数代替，数据相关的表格要求代码中要首先筛选出所有数值'
                               '在写JSON格式中不要用df代替参数，用JSON格式调用函数时要直接将DataFrame数据写在里面，不要用变量代替。像这样写：data=pd.DataFrame({\'Year\': [\'2021\', \'2022\', \'2023\', \'2024\'],\'Sales\': [500, 700, 800, 600],\'Expenses\': [400, 450, 500, 550]})'
                               ''
                               '非必要时刻，不运用python代码'
                               '非必要时刻，不写python代码'
                               ''
                               ''

                },
                {
                    "role": "user",
                    "content": content
                }

            ],
            "temperature": 0,
            "max_tokens": 8000,
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


def llm_text2(response):
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
def link_llm2(text):
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


def AI_run2(content, model, API_key):
    response = llm_model2(content, model, API_key)
    text = llm_text2(response)
    return text


if __name__ == '__main__':
    try:
        while True:
            content = input("写入需求:")
            text = AI_run2(content, model=None, API_key='1')
            #link_llm2(text)
    except KeyboardInterrupt:
        print("程序出错已退出。")