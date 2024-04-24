import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import threading
import pygetwindow as gw
import base64
import tempfile
import pyautogui
import numpy as np
import time

class Empty_CC_1(tk.Tk):
    def __init__(self, window_size="800x600"):
        super().__init__()
        self.title("Cycle Count Automation")
        self.geometry(window_size)

        self.combined_df = None  # 初始化combined_df属性

        # 创建带滚动条的文本框
        self.instructions_text = tk.Text(self, wrap="word", bg="white", height=10)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.instructions_text.yview)
        self.instructions_text.configure(yscrollcommand=self.scrollbar.set)

        # 插入指导步骤文本
        instructions = """This tool automates [Cycle Count].

Retrieve the following data from WMS:

[Pick Location Inquiry]: Set a [Zone] number (such as 01, 09, 15).

Import all data sheets using the [Import Excel Files] function.

Note: The WMS PROD RF interface may freeze, potentially resulting in missed locations in [Cycle Count] outcomes.

Please consult the release notes for data format and additional information.\n"""
        self.instructions_text.insert("1.0", instructions)

        # 布局文本框和滚动条
        self.instructions_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.scrollbar.grid(row=0, column=1, sticky="ns", pady=10)

        # 配置网格行/列的权重，确保文本框可以随窗口大小调整而扩展
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # 创建单选按钮的选中值变量
        self.cycle_count_option = tk.IntVar()
        self.cycle_count_option.set(1)  # 默认选中第一个选项

        # 创建单选按钮
        ttk.Radiobutton(self, text="Cycle Count Empty", variable=self.cycle_count_option, value=1).grid(row=1, column=0, sticky="w", padx=20)
        ttk.Radiobutton(self, text="Cycle Count Everything", variable=self.cycle_count_option, value=2).grid(row=1, column=0, sticky="e", padx=20)

        # 创建导入按钮，并使用grid布局
        self.import_button = ttk.Button(self, text="Import Excel Files", command=self.import_excel)
        # 适当调整按钮的放置位置
        self.import_button.grid(row=2, column=0, columnspan=2, pady=10, sticky="sn")

        # 添加下一步按钮
        self.next_button = ttk.Button(self, text="Next", state="disabled", command=self.goto_next_step)  # 下一步按钮初始不可用
        self.next_button.grid(row=2, column=1, sticky="se", padx=10, pady=10)

        # 配置网格布局权重，确保按钮固定在底部边缘
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(1, weight=1)

       
    def goto_next_step(self):
        if self.combined_df is not None:
            Empty_CC_2(self, self.combined_df)


 
    def import_excel(self):
        self.instructions_text.insert("end", "\nStart to import data...\n")
        self.instructions_text.see("end")  # 确保新消息可见
        
        file_paths = filedialog.askopenfilenames(title="Select Excel Files", filetypes=(("Excel files", "*.xls"),))
        
        if file_paths:
            try:
                new_data_frames = []
                required_columns_base = ["Location", "Current Qty", "Last Count Date"]
                required_columns = required_columns_base + (["Barcode", "Expiry Date"] if self.cycle_count_option.get() == 2 else [])
                
                for file_path in file_paths:
                    df = pd.read_excel(file_path)
                    
                    # 检查是否包含所有必需的列
                    if not all(col in df.columns for col in required_columns):
                        self.show_error_dialog("Critical data missing", "450x150", "Critical data missing in one or more files: Please ensure all imported files contain the following three columns of data:\nLocation, Current Qty, Last Count Date.")
                        return
                    
                    # 根据cycle_count_option的值应用不同的数据处理逻辑
                    if self.cycle_count_option.get() == 1:
                        # 仅保留“Current Qty”为空的行
                        df_filtered = df[required_columns][df["Current Qty"].isna()]
                    else:
                        # cycle_count_option等于2时，不过滤“Current Qty”为空的行，直接使用筛选后的数据
                        df_filtered = df[required_columns]
                    
                    new_data_frames.append(df_filtered)
                
                self.combined_df = pd.concat(new_data_frames).drop_duplicates().reset_index(drop=True)

                # 检查合并后的DataFrame是否为空
                if self.combined_df.empty:
                    self.instructions_text.insert("end", "\nNo empty location can be found.\n")
                else:
                    self.next_button['state'] = 'normal'
                    final_data_message = "\nData import completed, final data has {} rows and {} columns.\n\nClick [Next] to process data.\n".format(self.combined_df.shape[0], self.combined_df.shape[1])
                    self.instructions_text.insert("end", final_data_message)
                self.instructions_text.see("end")  # Auto-scroll to the new text

            except Exception as e:
                print("Unable to read Excel file:", e)
                self.show_error_dialog("Error", "300x100", f"Unable to read Excel file: {e}")


    def show_error_dialog(self, title, size, message):
        error_window = PredefinedWindow(self, title, size)
        message_label = ttk.Label(error_window, text=message)
        message_label.pack(pady=10, padx=10, fill="both", expand=True)
        ttk.Button(error_window, text="Confirm", command=error_window.destroy).pack(pady=(0,10), padx=10)

        def adjust_wraplength(event):
            # 调整wraplength为窗口宽度减去一些像素以留出边距
            message_label.config(wraplength=event.width - 20)

        # 绑定到窗口大小变化事件
        error_window.bind("<Configure>", adjust_wraplength)


class Empty_CC_2(tk.Toplevel):
    def __init__(self, parent, combined_df):
        super().__init__(parent)  # 调用父类构造函数来正确设置父窗口
        self.title("Cycle Count Automation")  # 设置窗口标题
        self.geometry("800x600")  # 设置窗口尺寸
        self.combined_df = combined_df  # 存储传递的DataFrame
        
        self.execution_running = False  # 添加一个执行状态标志
        self.current_location_index = 0  # 添加用于跟踪当前位置的属性

        # 创建带滚动条的文本框
        self.text_box = tk.Text(self, wrap="word", bg="white", height=10)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.text_box.yview)
        self.text_box.configure(yscrollcommand=self.scrollbar.set)

        # 布局文本框和滚动条
        self.text_box.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.scrollbar.grid(row=0, column=1, sticky="ns", pady=10)

        # 添加下拉选择框以选择排序方式
        self.sort_option_var = tk.StringVar(self)
        self.sort_option_var.set("Ascending")  # 默认值
        self.sort_options = ["Ascending", "Descending"]
        self.sort_option_menu = ttk.Combobox(self, textvariable=self.sort_option_var, values=self.sort_options, state="readonly")
        self.sort_option_menu.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        self.sort_option_menu.bind("<<ComboboxSelected>>", self.update_display)

        # 添加日期选择的下拉菜单
        # 首先，将'Last Count Date'列转换为datetime对象
        self.combined_df['Last Count Date'] = pd.to_datetime(self.combined_df['Last Count Date'], format="%d/%m/%Y %H:%M").fillna(pd.Timestamp('1900-01-01'))

        # 最后，提取日期部分。这一步确保所有日期都没有时间部分
        self.combined_df['Last Count Date'] = self.combined_df['Last Count Date'].apply(lambda x: x.date())


        unique_dates = sorted(self.combined_df['Last Count Date'].unique())
        
        # 生成带有前缀和后缀的日期选项
        unique_dates_str = [f"Update to date: {date}" for date in unique_dates]
        if unique_dates_str:
            # 在最后一个日期选项后加上"(All data)"
            unique_dates_str[-1] += " (All data selected)"

        # 设置默认值为最后一天的日期，即所有数据
        self.date_option_var = tk.StringVar(self)
        self.date_option_var.set(unique_dates_str[-1] if unique_dates_str else "")

        self.date_options_menu = ttk.Combobox(self, textvariable=self.date_option_var, values=unique_dates_str, state="readonly")
        self.date_options_menu.grid(row=2, column=0, padx=10, pady=5, sticky="ew")
        self.date_options_menu.bind("<<ComboboxSelected>>", self.update_display)

        # 添加选择 Sleep 时间的提示文字
        self.sleep_time_label = tk.Label(self, text="Please select the script execution speed (sleep time between each step, in milliseconds):")
        self.sleep_time_label.grid(row=3, column=0, padx=10, pady=(10, 0), sticky="nw")

        # 创建 Sleep 时间选项的下拉菜单
        self.sleep_time_var = tk.StringVar(self)
        self.sleep_time_options = ['0.5', '1', '1.5', '2', '2.5', '3', '3.5', '4', '4.5', '5']  # 以毫秒为单位
        self.sleep_time_menu = ttk.Combobox(self, textvariable=self.sleep_time_var, values=self.sleep_time_options, state="readonly")
        self.sleep_time_menu.set('1')  # 设置默认值
        self.sleep_time_menu.grid(row=3, column=0, padx=10, pady=5, sticky="e")
        
        # 确保添加对应的 grid_rowconfigure 调用以适应新组件...

        # 初始数据显示
        self.update_display()

        # 配置网格行/列的权重，确保文本框和下拉框可以随窗口大小调整而扩展
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # 添加导出至Autokey脚本文件按钮
        self.export_button = ttk.Button(self, text="Start Execution", command=self.start_export_thread)
        self.export_button.grid(row=3, column=1, sticky="se", padx=10, pady=10)


        # 配置网格布局权重，确保按钮固定在底部边缘
        self.grid_rowconfigure(3, weight=1)
        self.grid_columnconfigure(1, weight=1)

    def start_export_thread(self):
        export_thread = threading.Thread(target=self.export_to_txt, name="ExportThread")
        export_thread.daemon = True  # 设置为守护线程，这样当主程序退出时线程也会被终止
        export_thread.start()

    def update_display(self, event=None):
        # 解析选择的日期
        selected_date_str = self.date_option_var.get().replace("Update to date: ", "").replace(" (All data)", "")
        try:
            selected_date = pd.to_datetime(selected_date_str).date()
        except ValueError:
            selected_date = self.combined_df['Last Count Date'].max()

        ascending = True if self.sort_option_var.get() == "Ascending" else False
        
        # 使用新属性来存储过滤和排序后的DataFrame
        self.filtered_df = self.combined_df[self.combined_df['Last Count Date'] <= selected_date].sort_values(by="Location", ascending=ascending)
        self.filtered_df['Location'] = self.filtered_df['Location'].apply(lambda x: f"{x:09}")
        if 'Barcode' in self.filtered_df.columns:
            # 如果存在，则对 'Barcode' 列应用格式化函数，CC empty没有这一列
            self.filtered_df['Barcode'] = self.filtered_df['Barcode'].apply(lambda x: format(x, '.0f') if isinstance(x, (int, float)) else x)
        else:
            pass
        rows_count_info = f"Total rows displayed: {len(self.filtered_df)}\n" # 计算总行数
        df_string = self.filtered_df.to_string(index=False) # 数据文本化
        display_string = rows_count_info + df_string # 把行数和数据合起来准备显示

        self.text_box.delete("1.0", tk.END)
        self.text_box.insert("1.0", display_string)


    def export_to_txt(self):

        try:  # 这是按下运行按钮时的窗口激活判断
            wms_window = gw.getWindowsWithTitle("vipdjwmsapp.davidjones.com.au - PuTTY")[0]
            if not wms_window.isActive:
                if wms_window.isMinimized:
                    wms_window.restore()
                wms_window.activate()  # 未激活则激活
        except IndexError:
            message = "\n\n[vipdjwmsapp.davidjones.com.au - PuTTY] window is not active or not open."
            self.text_box.insert(tk.END, message)
            self.text_box.see(tk.END)  # Scroll to the bottom
            return # 激活不了直接返回

        # 锁定下拉选择框
        self.sort_option_menu.config(state='disabled')
        self.date_options_menu.config(state='disabled')
        
        # 使用过滤后的DataFrame更新self.combined_df，过滤是在display环节完成的
        # 复制 DataFrame 并转换所有列为字符串类型
        self.combined_df = self.filtered_df.copy().astype(str)

        start_message = "\n\nExecution has started. Processing locations..."
        self.text_box.insert(tk.END, start_message)
        self.text_box.see(tk.END)  # Ensure the message is visible

        self.export_button.config(text="Executing", state="disabled")
        self.execution_running = True

        # pd.set_option('future.no_silent_downcasting', True) # 避免在如 where、mask 和 clip 等方法中无提示地进行数据类型降级。
        sleep_time = float(self.sleep_time_var.get())
        locations = self.combined_df['Location'].apply(lambda x: str(x).zfill(9)).tolist() # 'Location'的列，转换为字符串，前面填0到9位，最后写成list。
        barcode_data_mark = False  # 在循环开始前定义，是否需要处理barcode，即不是empty CC
        base64_image_data = "iVBORw0KGgoAAAANSUhEUgAAADgAAAAICAYAAACs/DyzAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsEAAA7BAbiRa+0AAAJpSURBVEhLnVWvjvswDP7uHqOKThrNA0zVSNFQWUHJQHHRxgaKC8Y2VFwwMlBWVDQyVXuA0JOmqK/R+9J/6+7W3353n2QpsWPHjh37LS6Keg6DCtk6wNlJcfCslpOtESQzxMUG7RmiyrAMkm5DuDGKzSBtUF0z7KIEims3LmDErS0FGd7t47rHMsqHMw/4JvuXPiARxlvYwoJFkbm/hA0cA8rcuE5DWXM1IlnHaVzLMU+GdZqGtRundShHfBIvHdmQPFM82DQ6Ke253b631ct7erRzp1f6Rh67Iz1p/G/9fMenBsSMfMlMFShimoHZ6yYDPaRjQ58S5BcN22Hok1DIj1ln846yBPzG9t8wqS9D+KJElI+8VQpREIAJx3vDEB8mAoiqwpVrKT8g9K0RtZBwbI2jqYb8CG075ExAsly2fIyLOTzCOcIJPviqf8OU/kww+nO3+Yl3qBtzRb+agyeubTgO1/qzPWHgrmDrS5dRhYu2sfr2mJZ3QGEq4LCFOO34op1ghPzIn7ENpx/nBV7ru20V9sSMM4MMxBJYLebQtxxnloJtM8ARXMqs+WZQ3MwtzBePEZomsFwusd5rCH/Fq55AJdiVfByn2/8Wz/SHL2aQI6IPxo/9lU2TJccAFXQlIMQVpqrUjQpsRfrW17SLhciw7hRbWiMTi6dBqDyiEwKbif+mkhNf0EPXB3+NH/oMuhQsXfcxr/x0Dd5IputhwyDa9m/S7ENzZCQIkR56Y+0YSZSR92ODvEzD87oeP4wQ/sP0AA8Z9tobRkDf6pvR4uth3PzPmDCY0m/u45jwWFkt2Ev25psofAGNAGgVBZyj1QAAAABJRU5ErkJggg=="
        image_data = base64.b64decode(base64_image_data) # 输出临时图片文件以判断WARNING
        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_image:
            tmp_image.write(image_data)
            tmp_image_path = tmp_image.name

        def is_notepad_plus_plus_active(): # 这是运行循环里的窗口激活监控
            try:
                wms_window = gw.getWindowsWithTitle("vipdjwmsapp.davidjones.com.au - PuTTY")[0]
                return wms_window.isActive
            except IndexError:
                return False

        for index in range(self.current_location_index, len(locations)): # current_location_index在class最开始已初始化为0
            # print("Start Location：",self.current_location_index, "/", len(locations))

            if not is_notepad_plus_plus_active():
                message = f"\n\n[vipdjwmsapp.davidjones.com.au - PuTTY] window is not active or not open.\nExecution paused at Location: {locations[self.current_location_index-1]}.\nPlease return to Cycle Count entry, then continue execution."
                self.text_box.insert(tk.END, message)
                self.text_box.see(tk.END)  # Scroll to the bottom
                self.export_button.config(text="Continue", state="normal")
                break

            barcode_data_mark = False # 初始化循环假设为空location，即没有barcode
            current_location = locations[self.current_location_index] # 取当前index的location值
            location_df = self.combined_df[self.combined_df['Location'] == current_location] # 把所有与当前location相同的df整个取出来
            # print(location_df)
            
            if location_df['Current Qty'].str.match(r'^\d+\s+units$').any(): # 'Current Qty'的值符合“数字+空格+units”的格式
                barcode_data_mark = True # 说明有barcodes需要处理
                location_df.loc[:, 'Barcode'] = location_df['Barcode'].replace('nan', np.nan)# 确保 'Barcode' 列中的 'nan' 文本被正确视为 NaN
                self.current_location_index += len(location_df) # 去到下一个location，即index加上location的df的长度
                
                if location_df['Barcode'].isna().any() or location_df['Expiry Date'].notna().any(): # 检查 'Barcode' 为空或 'Expiry Date' 不为空的情况
                    message = f"\n\n{locations[self.current_location_index-1]} has been skipped, due to empty barcode or existing expiry date."
                    self.text_box.insert(tk.END, message)
                    self.text_box.see(tk.END)  # Scroll to the bottom
                    continue  # 可以直接去下一个循环

            else:
                self.current_location_index += 1 # 如果当前location没有Current Qty，即emplty location

            pyautogui.write(current_location)
            pyautogui.press('enter')
            time.sleep(sleep_time)
            # print ("Location", current_location, "Enter, 等待", sleep_time, "秒")
            try:
                image_location = pyautogui.locateCenterOnScreen(tmp_image_path, confidence=0.5)
                if image_location:
                    message = f"\n\nWARNING message in location {locations[self.current_location_index-1]}, Ctrl + A sent to continue."
                    self.text_box.insert(tk.END, message)
                    self.text_box.see(tk.END)  # Scroll to the bottom
                    pyautogui.hotkey('ctrl', 'a')
                    time.sleep(sleep_time)
                    # print ("Ctrl+A, 等待", sleep_time, "秒")
            except pyautogui.ImageNotFoundException:
                pass # 忽略未找到图像的异常
            except Exception as e:
                print("发生错误：", e) # 输出其他类型的错误信息和堆栈跟踪
                
            # 准备输入barcodes
            if barcode_data_mark:
                current_location_df = self.combined_df[self.combined_df['Location'] == current_location] # 筛选当前Location的数据
                for _, row in current_location_df.iterrows(): # 遍历current_location_df这个DataFrame中的每一行
                    current_qty_data = row['Current Qty'] # 从Current Qty列中提取数量
                    current_bc_qty = int(current_qty_data.split(' ')[0])  # 从格式是数字+空格+units中提取数量
                    current_bc = str(row['Barcode']).zfill(13)  # 从Barcode列中提取条形码并格式化, 在前面加0到达13位

                    # 对于每个条形码，执行指定次数的发送操作
                    for _ in range(current_bc_qty):
                        pyautogui.write(current_bc)
                        pyautogui.press('enter')
                        time.sleep(sleep_time)
                        # print ("Barcode", current_bc, "Enter, 等待", sleep_time, "秒")

            pyautogui.hotkey('ctrl', 'n') # 实际操作时激活这一行
            time.sleep(sleep_time)
            # print ("Ctrl+N, 等待", sleep_time, "秒")

            # self.update() # 好像没用

            if self.current_location_index==len(locations):  # If the loop finishes normally without a break
                self.current_location_index = 0  # Reset position for next execution
                complete_message = "\n\nExecution complete. All locations have been processed."
                self.text_box.insert(tk.END, complete_message)
                self.text_box.see(tk.END)
                self.export_button.config(text="Start Execution", state="normal")
                self.execution_running = False
                break

if __name__ == "__main__":
    app = Empty_CC_1()
    app.mainloop()
