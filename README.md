# AI_excel
这是一个可以自动操作Excel的AI工具

软件效果视频：[只说人话就可以做表格了！我造了一个自动操作Excel表的网站](https://www.bilibili.com/video/BV1JzKpeTEkF/?vd_source=28ba27f4f650db659b1dd1ace9f5fc5c)

本项目调用大语言模型的API接口，并让模型能够调用操作Excel表文件的函数

项目用streamlit编写界面

文件**main.py**：streamlit界面代码

文件**gpt_api.py**：调用API接口并导入操作表格的函数

文件**xl_class.py**：表格操作相关函数

文件**gpt_data.py**：数据分析AI

文件**chart.py**：绘制统计图函数

文件**requirements.txt**：所需要的python依赖库

运行命令 `pip install -r requirements.txt` 即可安装所有所需依赖。

运行命令 `streamlit run main.py` 即可运行网站程序
