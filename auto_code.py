import re
import sys
from io import StringIO

# 允许的库白名单
ALLOWED_MODULES = {
    "numpy", "pandas", "openpyxl", "os", "csv", "math", "random", "json", "re", "time", "copy"
}

'''
def set_resource_limits():
    """
    设置 CPU 和内存使用限制。
    """
    # 设置 CPU 时间限制（单位：秒）
    cpu_time_limit = 8  # 最多允许 8 秒的 CPU 时间
    resource.setrlimit(resource.RLIMIT_CPU, (cpu_time_limit, cpu_time_limit))

    # 设置内存限制（单位：字节）
    memory_limit = 512 * 1024 * 1024  # 最多允许 512 MB 内存
    resource.setrlimit(resource.RLIMIT_AS, (memory_limit, memory_limit))
'''


def check_code_safety(code: str) -> bool:
    """
    检查代码是否只使用了白名单中的库。
    :param code: 生成的 Python 代码
    :return: 如果代码安全返回 True，否则返回 False
    """
    # 匹配 import 语句
    import_pattern = re.compile(r"^\s*(?:from|import)\s+(\w+)", re.MULTILINE)
    imported_modules = set(import_pattern.findall(code))

    # 检查是否有不允许的库
    for module in imported_modules:
        if module not in ALLOWED_MODULES:
            print(f"Error: 不允许的库 '{module}'")
            return False
    return True

def run_generated_code(code: str, **kwargs):
    """
    动态执行生成的 Python 代码，并返回执行结果。
    :param code: 生成的 Python 代码（字符串形式）
    :param kwargs: 传递给代码的参数
    :return: 执行结果或错误信息
    """
    # 捕获标准输出
    original_stdout = sys.stdout
    sys.stdout = StringIO()

    try:
        # 将传入的参数添加到局部变量中
        local_vars = kwargs

        # 动态执行代码，并将参数传递进去
        exec(code, globals(), local_vars)

        # 获取输出
        output = sys.stdout.getvalue()

        # 如果输出为空，尝试从局部变量中获取结果
        if not output and "result" in local_vars:
            output = str(local_vars["result"])
    except SyntaxError as e:
        output = f"SyntaxError: {e}"
    except Exception as e:
        output = f"RuntimeError: {e}"
    finally:
        # 恢复标准输出
        sys.stdout = original_stdout

    return output

def generate_code_with_llm(user_input: str) -> str:
    """
    模拟调用大语言模型生成 Python 代码。
    :param user_input: 用户输入的需求描述
    :return: 生成的 Python 代码
    """
    # 这里可以替换为实际的大语言模型 API 调用
    # 示例：根据用户输入生成代码
    if "read excel" in user_input.lower():
        return """
import pandas as pd
df = pd.read_excel("生产情况.xlsx")
print(df.head())
"""
    elif "write excel" in user_input.lower():
        return """
import pandas as pd
data = {"Name": ["Alice", "Bob"], "Age": [25, 30]}
df = pd.DataFrame(data)
df.to_excel("output.xlsx", index=False)
print("Excel 文件已生成")
"""
    else:
        return """
print("未识别的任务")
"""

def auto_code_main(user_input):
    # 生成代码
    generated_code = generate_code_with_llm(user_input)
    print("生成的代码：")
    print(generated_code)

    # 检查代码安全性
    if not check_code_safety(generated_code):
        print("代码包含不允许的库，终止执行。")
        return

    # 执行生成的代码
    result = run_generated_code(generated_code)
    print("执行结果：")
    print(result)
    return result

if __name__ == "__main__":
    # 用户输入需求
    user_input = input("请输入您的 Excel 自动化需求：")

    auto_code_main(user_input)