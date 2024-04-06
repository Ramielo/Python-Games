from mitmproxy import http
import os
from pathlib import Path

from pathlib import Path

# 获取用户输入的日志文件路径
log_file_path_input = input("Enter the path where you want to save the log file: ")
if not log_file_path_input:
    # 如果用户没有输入路径，使用当前用户的主目录
    log_file_path = Path.home() / "http_traffic_log.txt"
else:
    # 替换所有反斜杠为正斜杠来避免转义问题
    log_file_path_corrected = log_file_path_input.replace("\\", "/")
    # 创建Path对象
    log_file_path = Path(log_file_path_corrected)

# 确保路径包括文件名和扩展名
if log_file_path.is_dir():
    # 如果输入的是一个目录，添加默认的文件名
    log_file_path = log_file_path / "http_traffic_log.txt"

# 转换为字符串用于文件操作
log_file_path = str(log_file_path)


def response(flow: http.HTTPFlow) -> None:
    # 定义要监控的域名
    monitored_domain = "twitter.com"

    # 检查请求的URL是否包含该域名
    if monitored_domain in flow.request.host:
        # 如果是，就写入日志文件
        with open(log_file_path, "a", encoding="utf-8") as log_file:
            log_file.write(f"Request: {flow.request.method} {flow.request.url}\n")
            log_file.write(f"Status code: {flow.response.status_code}\n")
            log_file.write(f"Response: {flow.response.text[:1000]}\n")  # 仅记录响应的前1000个字符
            log_file.write("="*50 + "\n")
