import streamlit as st
import pandas as pd
from xl_class import *
from gpt_api import AI_run, link_llm
from gpt_data import llm_model2, llm_text2
import time
import copy
import json
import plotly.express as px
import plotly.graph_objects as go
import os


def plot_chart(data, chart_type, x_column, y_columns=None, legend_title=None, title='Chart', xlabel=None, ylabel=None,
               colors=None):
    """
    根据提供的数据和参数绘制不同类型的图表。

    :param data: 输入的数据框
    :param chart_type: 图表类型（"bar", "line", "pie", "scatter"等）
    :param x_column: x轴对应的数据列名
    :param y_columns: y轴对应的数据列名列表，对于饼图此参数为None，对于多条折线图则应提供多个y轴列名
    :param legend_title: 图例标题
    :param title: 图表标题
    :param xlabel: x轴标题
    :param ylabel: y轴标题
    :param colors: 颜色列表或字典，根据图表类型应用到不同的线条或类别
    """
    if chart_type not in ['bar', 'line', 'pie', 'scatter']:
        st.write(f"{chart_type} not supported.")
        return

    if chart_type == 'pie':
        # For pie charts, we assume a single y_column is provided for the values.
        fig = px.pie(data, names=x_column, values=y_columns[0], title=title, color_discrete_sequence=colors)
    else:
        fig = go.Figure()

        for i, y_col in enumerate(y_columns):
            if chart_type == 'bar':
                fig.add_trace(
                    go.Bar(x=data[x_column], y=data[y_col], name=y_col, marker_color=colors[i] if colors else None))
            elif chart_type == 'line':
                fig.add_trace(go.Scatter(x=data[x_column], y=data[y_col], mode='lines', name=y_col,
                                         line=dict(color=colors[i] if colors else None)))
            elif chart_type == 'scatter':
                fig.add_trace(go.Scatter(x=data[x_column], y=data[y_col], mode='markers', name=y_col,
                                         marker=dict(color=colors[i] if colors else None)))

        fig.update_layout(title=title, xaxis_title=xlabel if xlabel else x_column,
                          yaxis_title=ylabel if ylabel else ('Value' if y_columns else ''), legend_title=legend_title)
    st.plotly_chart(fig)


def save_list_to_txt(file_path, data_list):
    """
    将列表存入txt文件
    :param file_path: 文件路径
    :param data_list: 要存储的列表
    """
    with open(file_path, 'w', encoding='utf-8') as file:
        for item in data_list:
            file.write(str(item) + '\n')  # 每个元素写入一行


# 获取当前日期
time_c = get_current_date()

# 初始化状态变量，用于存储历史回答
if 'history' not in st.session_state:
    st.session_state.history = []

st.title("Excel自动化AI工具")

col3_1, col3_2 = st.columns([16, 1])

with col3_1:
    input_str = st.text_input("请输入你的Excel文件路径：")
    file_path = process_path_or_filename(input_str)
    # 检查文件是否存在
    exist_file = os.path.exists(file_path)
    if file_path:
        if exist_file:
            pass
        else:
            st.error(f"文件 '{file_path}' 不存在！请重新输入文件路径！")

with col3_2:
    if file_path:
        st.write(f"### 💡")


tab1, tab2, tab3 = st.tabs(["AI脚本生成", "自定义脚本", "数据分析"])


with tab1:
    col4_1, col4_2, col4_3 = st.columns([6, 8, 2])
    with col4_1:
        # 定义主流大语言模型列表
        models1 = ["GLM-4-Air", "GLM-4-Flash", "Doubao-1.5-lite-32k", "Doubao-1.5-pro-32k", "deepseek-chat", "deepseek-reasoner", "deepseek-coder", "qwen-turbo-latest", "qwen-plus-latest", "qwen-coder-plus-latest"]
        # 创建下拉列表，默认选择 DeepSeek
        selected_model = st.selectbox(
            "选择一个大语言模型",
            models1,
            index=models1.index("GLM-4-Air")  # 设置默认选项为 DeepSeek
        )
    with col4_2:
        API_key_1 = st.text_input("输入API密钥", type="password")
    with col4_3:
        st.markdown("[如何获得API密钥](https://apifox.com/apidoc/shared-0fd7ea54-919e-4c93-b673-c60219bc82e0/doc-4739665)", )

    query = st.text_area("请输入关于你上传文件路径下对该文件的文件的指令，或关于上传文件的需求：")

    col1_1, col1_2 = st.columns([1, 2])
    with col1_1:
        button = st.button("执行任务")
    with col1_2:
        checked = st.checkbox(f"复制 {file_path} 文件备份")

    if button:
        if not file_path:
            st.error("请先输入文件路径")
        elif not query:
            st.error("请输入指令或需求")
        elif not exist_file:
            st.error("文件不存在，请重新输入文件路径")
        elif not API_key_1:
            st.error("请输入API密钥")
        else:
            with st.spinner("AI思考中，请稍等..."):
                time_tab_1_1 = time.time()
                if checked:
                    copy_excel_with_pandas(file_path)
                try:
                    text = AI_run(query, model=selected_model, API_key=API_key_1)
                    resp = link_llm(text, file_path)
                    st.write(text)
                    if resp!=text:
                        st.write(resp)
                    # 将生成的回答插入到历史记录的开头
                    st.session_state.history.insert(0, text)
                except:
                    print("程序出错。")
                    st.error("AI执行出错")
                time_tab_1_2 = time.time()
                st.write("请求用时：", time_tab_1_2-time_tab_1_1, "秒")

# 在侧边栏显示历史回答
with st.sidebar:
    st.markdown("### 历史回答脚本")
    len_history = len(st.session_state.history)
    len_list = list(range(1, len_history+1))
    len_list.reverse()
    for i, answer in enumerate(st.session_state.history):
        with st.expander(f"脚本 {len_list[i]}", expanded=False):
            st.write(answer)


with tab2:
    uploaded_file = st.file_uploader("上传脚本txt文件：", type="txt")

    if uploaded_file:
        column1, column2 = st.columns([7, 2])
        with column1:
            st.write("已接收文件：", uploaded_file.name)
        with column2:
            button_run1 = st.button("连续执行脚本")

        with st.expander(f"{uploaded_file.name}文件内容："):
            lines = uploaded_file.read().decode("utf-8")
            st.write(str(lines))

        if button_run1:
            if not file_path:
                st.error("请先输入文件路径")
            elif not exist_file:
                st.error("文件不存在，请重新输入文件路径")
            else:
                with st.spinner("正在执行脚本，请稍等..."):
                    time_tab_2_1 = time.time()
                    time.sleep(1)
                    try:
                        link_llm(lines, file_path)
                        st.write("脚本执行完毕 ", time_c)
                    except:
                        st.error("脚本执行出错")
                    time_tab_2_2 = time.time()
                    st.write("用时", time_tab_2_2-time_tab_2_1, "秒")


    st.divider()
    with st.expander("当前历史脚本,顺序从下到上", expanded=True):
        st.write(st.session_state.history)
        button_load = st.button("导出到txt文件")
        if button_load:
            if st.session_state.history:
                # 创建一个副本，以防止修改原始列表
                history_list_copy = copy.deepcopy(st.session_state.history)
                # 反转列表，以便从上到下显示
                history_list_copy.reverse()

                txt_file_name = f"history_{time_c}.txt"
                save_list_to_txt(txt_file_name, history_list_copy)
                st.success(f"成功导出为 {txt_file_name} 文件")
            else:
                st.warning("没有历史记录")


def link_llm2(text):
    """
    将字符串用JSON进行解析，并运行JSON中的函数
    :param text:
    :return:
    """
    # 正则表达式
    pattern = r'\{[^{}]*\{.*?\}[^{}]*\}'

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
            data = datas[0]
            try:
                print(data)
                exec(data)
            except:
                str_text = "不能执行此动作"
                print(str_text)
                return str_text
    else:
        return text


# 数据分析Tab
with tab3:
    col5_1, col5_2, col5_3 = st.columns([6, 8, 2])
    with col5_1:
        # 定义主流大语言模型列表
        models2 = ["abab7-chat-preview", "deepseek-coder", "deepseek-chat", "deepseek-reasoner", "qwen-max-latest", "GLM-4-Flash", "GLM-4-Plus", "Doubao-1.5-pro-256k"]
        # 创建下拉列表，默认选择 DeepSeek
        selected_model = st.selectbox(
            "选择一个数据分析大语言模型",
            models2,
            index=models2.index("abab7-chat-preview")  # 设置默认选项为 DeepSeek
        )
    with col5_2:
        API_key_2 = st.text_input("输入数据分析大语言模型API密钥", type="password")
    with col5_3:
        st.markdown("[怎样获得API密钥](https://apifox.com/apidoc/shared-0fd7ea54-919e-4c93-b673-c60219bc82e0/doc-4739665)", )

    data = st.file_uploader("上传你的Excel文件（xlsx格式）：", type="xlsx")
    if data:
        st.session_state["df"] = pd.read_excel(data)
        with st.expander("原始数据", expanded=True):
            st.dataframe(st.session_state["df"])

    query2 = st.text_area("请输入需求")
    if data:
        text_2 = query2 + "数据如下" + st.session_state["df"].to_string()

    col2_1, col2_2 = st.columns([1, 2])
    with col2_1:
        button2 = st.button("生成回答")

    if button2:
        if not data:
            st.error("请先上传文件")
        elif not query2:
            st.error("请输入指令或需求")
        elif not API_key_2:
            st.error("请输入API密钥")
        else:
            with st.spinner("AI思考中，请稍等..."):
                time_tab_3_1 = time.time()
                try:
                    respose2 = llm_model2(text_2, selected_model, API_key_2)
                    #respose2 = AI_run2(text_2)
                    text_t = llm_text2(respose2)
                    st.write(text_t)
                    respose3 = link_llm2(text_t)
                    if respose3 != text_t:
                        st.write(respose3)
                except:
                    print("程序出错。")
                    st.error("程序出错。")
                time_tab_3_2 = time.time()
                st.write("请求用时：", time_tab_3_2-time_tab_3_1, "秒")


