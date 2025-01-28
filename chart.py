import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go


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


