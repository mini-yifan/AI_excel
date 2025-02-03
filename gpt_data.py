# -*- coding: utf-8 -*-
from openai import OpenAI
from xl_class import *
import json
import re
from chart import plot_chart

file_path = 'C:/files/python_files/gpt_excel/入库单.xlsx'

def llm_model2(content, model=None, API_key=None):
    if model is None:
        model = "abab7-chat-preview"
    else:
        model = model
    base_url = 'https://api.mindcraft.com.cn/v1/'
    api_key = API_key
    client = OpenAI(base_url=base_url, api_key=api_key)

    params = {
        "model": model,
        "message": [
            {
                "role": "system",
                "content": "你是一个数据分析AI，对数据进行提取，分析，总结，归纳。"
                           "你可以绘制图表，当需要绘制图表时，用JSON格式回答，其他情况正常回答"
                           "不要出现python代码"
                           "'def_name'是一个列表，列表内容有'plot_chart()'，'text'的内容为我向你提问的问题，'response'是你对问题的回答"
                           "plot_chart(data, chart_type, x_column, y_columns=None, legend_title=None, title='Chart', xlabel=None, ylabel=None,colors=None)函数用于根据给定的数据和参数绘制不同类型的统计图表，支持条形图、多条折线图、饼图和散点图等，接受包括数据框、图表类型、x轴列名、y轴列名列表（对于饼图此参数为None）、图例标题、图表标题、x轴和y轴标签以及颜色在内的多个参数，其中除了数据、图表类型和x轴列名是必需的之外，其余参数均具有默认值以提供灵活性和便捷性，例如你可以通过指定不同的颜色来区分图表中的不同类别或系列，并且可以通过设置图例标题和坐标轴标签来增强图表的可读性和解释力。"
                           "plot_chart的函数，利用Plotly Express和Streamlit库根据提供的数据框、图表类型及其它参数生成饼图，当指定图表类型为\"pie\"时，它使用数据框中的分类信息（x_column）作为饼图各扇区的标签，以及对应数值信息（y_columns的第一个元素）决定各扇区大小，同时允许自定义标题、颜色等样式，并最终通过Streamlit展示生成的饼图。"
                           "绘制饼图时plot_chart函数的使用方法如下：plot_chart(data=data = pd.DataFrame({'Fruits': ['Apple', 'Banana', 'Cherry'],'Sales': [55, 45, 30]}), chart_type='pie', x_column='Fruits', y_columns=['Sales'], title='Fruit Sales Distribution')"
                           ""
                           "如向你提要求：来绘制一个展示销售和支出随年份变化的折线图，其中数据包含'Year'、'Sales'和'Expenses'三列，分别代表年份、销售额和支出额；通过设置chart_type为'line'指定绘制折线图，x_column选择'Year'作为x轴，y_columns包括['Sales', 'Expenses']以同时展示销售与支出两条折线，legend_title定义图例标题为'Indicator'，整体图表标题设为\"Sales and Expenses Over Years\"，并自定义了x轴标签为\"Year\"，y轴标签为\"Amount\"，最后通过colors参数设置了每条折线的颜色分别为'#636EFA'和'#EF553B'，"
                           '你用JSON格式回答：{"def_name":["plot_chart(data=pd.DataFrame({\'Year\': [\'2021\', \'2022\', \'2023\', \'2024\'],\'Sales\': [500, 700, 800, 600],\'Expenses\': [400, 450, 500, 550]}), chart_type=\'line\', x_column=\'Year\', y_columns=[\'Sales\', \'Expenses\'],legend_title=\'Indicator\', title="Sales and Expenses Over Years", xlabel="Year", ylabel="Amount",colors=[\'#636EFA\', \'#EF553B\'])"]}'
                           '当让你生成多个统计图时，可以根据需求写多个JSON字符串，但每一个JSON字符串要严格按我给你的格式书写，不要在一个JSON字符串里写多个"plot_chart"函数'

            },
            {
                "role": "user",
                "content": content
            }
        ],
        "temperature": 0,
        "max_tokens": 3000,
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
            text = AI_run2(content)
            #link_llm2(text)
    except KeyboardInterrupt:
        print("程序出错已退出。")