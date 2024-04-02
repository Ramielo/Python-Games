import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import os
import webbrowser
import pandas as pd
import xlrd
import openpyxl
import csv

# 定义一个预定义的窗口类
class PredefinedWindow(tk.Toplevel):
    def __init__(self, parent, title, window_size):
        super().__init__(parent)
        self.title(title)
        self.geometry(window_size)

# 修改MainWindow类，继承自tk.Tk不变，因为它是主窗口
class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Automation Toolkit 1.0")
        self.geometry("400x300")  # 设置窗口大小
        
        self.open_subwindow_button = ttk.Button(self, text="AutoKey coding for Empty Location C/C", command=self.open_subwindow)
        self.open_subwindow_button.pack(pady=20)

        # Button to open the second subwindow for SIM barcode typing
        self.open_sim_bc_button = ttk.Button(self, text="AutoKey coding for SIM barcode typing", command=self.open_sim_bc_subwindow)
        self.open_sim_bc_button.pack(pady=20)

    def open_subwindow(self):
        subwindow = Empty_CC_1(self, "800x600")
        # subwindow.grab_set()
        self.iconify()

        # 子窗口关闭时销毁子窗口并呼出主窗口
        subwindow.protocol("WM_DELETE_WINDOW", lambda: (subwindow.destroy(), self.deiconify()))

    def open_sim_bc_subwindow(self):
        subwindow = SIM_BC_1(self)
        self.iconify()

        # 子窗口关闭时销毁子窗口并呼出主窗口
        subwindow.protocol("WM_DELETE_WINDOW", lambda: (subwindow.destroy(), self.deiconify()))

class SIM_BC_1(PredefinedWindow):
    def __init__(self, parent, window_size="800x600"):
        super().__init__(parent, "AutoKey coding for SIM barcode typing", window_size)
        # Initialize any necessary attributes here, such as for storing data

        # Create a textbox with a scrollbar for instructions
        self.instructions_text = tk.Text(self, wrap="word", bg="white", height=10)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.instructions_text.yview)
        self.instructions_text.configure(yscrollcommand=self.scrollbar.set)

        # Insert guidance text
        instructions = """This tool is used to generate ".ahk" automation scripts compatible with the Autokey software. The scripts generated in this interface are used for typing numerous of barcodes in SIM.

It accepts input files exclusively in .csv format, from one of the following 4 resources:
1 - Data exported from Claim Enquiry
2 - Data exported from Transfer Enquiry
3 - Data exported from RMS Enquiry
4 - A clean data list consisting solely of Barcodes and Quantities, the first column should be Barcodes Data titled as [UPC], and the second column should be Quantities data titaled as [Qty].

To import files, click the [Import CSV Files] button. Note that importing new data will overwrite any previously imported data within the same interface window.\n"""
        self.instructions_text.insert("1.0", instructions)

        # Layout the textbox and scrollbar
        self.instructions_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.scrollbar.grid(row=0, column=1, sticky="ns", pady=10)

        # Configure grid row/column weights to ensure the textbox expands with the window size
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Create the "Import Excel Files" button and layout using grid
        self.import_button = ttk.Button(self, text="Import CSV Files", command=self.import_excel)
        self.import_button.grid(row=1, column=0, columnspan=2, pady=10, sticky="ew")
        self.data_dict = {}

        # Create the "Export Data" button and layout using grid
        self.export_button = ttk.Button(self, text="Export Data", command=self.export_data)
        self.export_button.grid(row=2, column=0, columnspan=2, pady=10, sticky="ew")
        
        # Optionally, add additional UI components as needed, similar to the setup in Empty_CC_1

    def import_excel(self):
        # Clear the textbox and display the starting message
        self.instructions_text.insert(tk.END, "\nStarting to import data...\n")

        # 打开文件对话框选择CSV文件
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not file_path:
            self.instructions_text.insert(tk.END, "\nData import cancelled.\n")
            return  # 用户取消了对话框

        self.data_dict.clear()  # 清空之前的数据
        rows_imported = 0  # 跟踪导入的行数

        # 读取CSV文件并定位所需的列
        with open(file_path, newline='', encoding='utf-8-sig') as csvfile:
            reader = csv.DictReader(csvfile)
            qty_header = None

            # 检查文件中的特定数量列标题
            for header in ("Claim Qty", "Qty", "QTY", "SOH", "TransferQty"):
                if header in reader.fieldnames:
                    qty_header = header
                    break

            if not qty_header:
                self.instructions_text.insert(tk.END, "\nNone of the specified qty columns found in the file.\n")
                return

            # 在字典中存储数据
            for row in reader:
                upc = row['UPC']
                qty = row[qty_header]
                # 检查UPC和数量是否非空，然后添加到字典中
                if upc and qty:
                    self.data_dict[upc] = qty  # 使用self.data_dict而不是data_dict
                    rows_imported += 1

        # 显示导入过程的摘要
        self.instructions_text.insert(tk.END, f"\nData import complete. Total rows imported: {rows_imported}\n")
        self.instructions_text.see("end")  

    def export_data(self):
        if not hasattr(self, 'data_dict') or not self.data_dict:
            self.instructions_text.insert(tk.END, "\nNo data to export.\n")
            return

        export_text = "#n::{\n"
        for upc, qty in self.data_dict.items():
            for _ in range(int(qty)):  # 确保数量是整数类型
                export_text += f'SendText "{upc}"\nSleep 500\n'
        export_text += "}"

        # 弹出保存文件对话框，让用户选择保存位置和文件名，指定后缀为.ahk
        filepath = filedialog.asksaveasfilename(defaultextension=".ahk", filetypes=[("AutoHotkey Scripts", "*.ahk")])
        if filepath:
            # 将脚本内容写入到文件
            with open(filepath, "w") as file:
                file.write(export_text)
            self.instructions_text.insert(tk.END, "\n\nPlease download and install the Autokey software from the official website. [.ahk] files can only be executed after installing the Autokey software.\n\
After double-clicking the [.ahk] file, a green [A] icon will appear in the bottom right corner of the taskbar. Press the [WIN] + [N] simultaneously to start running the script.\n\
The function of this script is very simple - it automates the typing of barcodes from your list according to the specified quantities. This is useful when you have a large number of barcodes that need to be entered into Claim or Transfer. Please note that for transfers, it is necessary to manually enter the first barcode to initiate the transfer, which is limited by SIM.\n\
As both this tool and SIM lack screen locking functionality, you must ensure that SIM remains on the top layer of the screen throughout the entire process. It's recommended not to perform any keyboard or mouse operations during the entire process.\n\
If you wish to terminate the execution during the process, you can right-click on the green [A] icon in the bottom right corner and directly close the opened [.ahk] file, or set a shortcut key to terminate the execution through the settings option of Autokey.\n\
It is advisable to close any opened [.ahk] files before running a new [.ahk] file.\n\
The installation of this software will not be blocked by the company's system policies. However, if used on a personal computer, be aware that some online games may detect the Autokey software as cheating software, leading to account bans.\n")
            self.instructions_text.see(tk.END)
            
            # 使用默认程序打开文件
            if os.path.exists(filepath):
                webbrowser.open(filepath)


# 修改SubWindow类，使其继承自新的PredefinedWindow类
class Empty_CC_1(PredefinedWindow):
    def __init__(self, parent, window_size):
        super().__init__(parent, "AutoKey coding for Empty Location C/C", window_size)
        self.combined_df = None  # 初始化combined_df属性

        # 创建带滚动条的文本框
        self.instructions_text = tk.Text(self, wrap="word", bg="white", height=10)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.instructions_text.yview)
        self.instructions_text.configure(yscrollcommand=self.scrollbar.set)

        # 插入指导步骤文本
        instructions = """This tool is used to generate ".ahk" automation scripts compatible with the Autokey software. The scripts generated in this interface are used to batch process [Cycle Count] tasks for [Empty Locations].

Follow the steps below to obtain the original Excel data:
Step 1 - Enter the WMS's [Pick Location Inquiry] interface, only input the zone number (e.g., 01, 09, 15, etc.), and press [Enter] key.
Step 2 - If the result generates <5000 rows of data, proceed to Step 5. If the result generates >5000 rows of data and <10000 rows of data, proceed to Step 4. If the result generates >10000 rows of data, proceed to Step 3.
Step 3 - Restrict the [Last Count Date] in the [Pick Location Inquiry] interface, advance the date step by step until pressing enter generates <10000 rows of data.
Step 4 - Due to WMS system limitations, the maximum amount of data that can be exported is 5000 rows (excluding the title). Export the data and store it locally. After completion, click the [Location column title] in the current WMS data interface to arrange the Location in descending order.
Step 5 - Export the data and store it locally (supported formats: [*.xlsx] [*.xls]). If you arrived here from Step 4, you will save a second file.

Now click the [Import Excel Files] button, select all the saved Excel data files.
Note: You can also submit any additional number of picklocation require data results, and this tool will automatically process, keeping all valid data.\n"""
        self.instructions_text.insert("1.0", instructions)

        # 布局文本框和滚动条
        self.instructions_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.scrollbar.grid(row=0, column=1, sticky="ns", pady=10)

        # 配置网格行/列的权重，确保文本框可以随窗口大小调整而扩展
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # 创建导入按钮，并使用grid布局
        self.import_button = ttk.Button(self, text="Import Excel Files", command=self.import_excel)
        # 适当调整按钮的放置位置
        self.import_button.grid(row=1, column=0, columnspan=2, pady=10, sticky="ew")

        # 添加上一步和下一步按钮
        # self.prev_button = ttk.Button(self, text="Previous", state="disabled")  # 上一步按钮始终不可用
        self.next_button = ttk.Button(self, text="Next", state="disabled", command=self.goto_next_step)  # 下一步按钮初始不可用
        
        # 布局按钮到窗口右下方和左下方
        # self.prev_button.grid(row=2, column=0, sticky="sw", padx=10, pady=10)
        self.next_button.grid(row=2, column=1, sticky="se", padx=10, pady=10)

        # 配置网格布局权重，确保按钮固定在底部边缘
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(1, weight=1)
        
    def goto_next_step(self):
        if self.combined_df is not None:
            self.destroy()  # 关闭当前窗口
            next_window = Empty_CC_2(self.master, "800x600", self.combined_df)
            # next_window.grab_set()
            # 子窗口关闭时销毁子窗口并呼出主窗口
            next_window.protocol("WM_DELETE_WINDOW", lambda: (next_window.destroy(), self.master.deiconify()))
        else:
            messagebox.showwarning("Warning", "No data to proceed.")

 
    def import_excel(self):
        self.instructions_text.insert("end", "\nStart to import data...\n")
        self.instructions_text.see("end")  # 确保新消息可见
        
        file_paths = filedialog.askopenfilenames(title="Select Excel Files", filetypes=(("Excel files", "*.xlsx;*.xls"),))
        
        if file_paths:
            try:
                new_data_frames = []
                required_columns = ["Location", "Current Qty", "Last Count Date"]
                
                for file_path in file_paths:
                    df = pd.read_excel(file_path)
                    
                    # 检查是否包含所有必需的列
                    if not all(col in df.columns for col in required_columns):
                        self.show_error_dialog("Critical data missing", "450x150", "Critical data missing in one or more files: Please ensure all imported files contain the following three columns of data:\nLocation, Current Qty, Last Count Date.")
                        return
                    
                    # 仅保留必需的列，并筛选出“Current Qty”为空的行
                    df_filtered = df[required_columns][df["Current Qty"].isna()]  # 确保只有“Current Qty”为空的行被保留
                    
                    new_data_frames.append(df_filtered)
                
                # 合并新导入的数据
                new_combined_df = pd.concat(new_data_frames).drop_duplicates().reset_index(drop=True)
                
                # 如果之前已经有数据，合并旧数据和新数据
                if self.combined_df is not None:
                    self.combined_df = pd.concat([self.combined_df, new_combined_df]).drop_duplicates().reset_index(drop=True)
                else:
                    self.combined_df = new_combined_df
                
                # 检查合并后的DataFrame是否为空
                if self.combined_df.empty:
                    self.instructions_text.insert("end", "\nNo empty location can be found.\n")
                else:
                    self.next_button['state'] = 'normal'
                    final_data_message = "\nData import completed, final data has {} rows and {} columns.\n".format(self.combined_df.shape[0], self.combined_df.shape[1])
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


class Empty_CC_2(PredefinedWindow):
    def __init__(self, parent, window_size, combined_df):
        super().__init__(parent, "Next Step for Empty Location C/C", window_size)
        self.combined_df = combined_df  # 存储传递的DataFrame
        self.initialize_ui()

    def initialize_ui(self):
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
        # 提取"Last Count Date"列中所有独特的日期，并按日期排序
        self.combined_df['Last Count Date'] = pd.to_datetime(self.combined_df['Last Count Date'], format="%d/%m/%Y %H:%M").dt.date
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
        self.sleep_time_options = ['500', '1000', '1500', '2000', '2500', '3000', '3500', '4000', '4500', '5000']  # 以毫秒为单位
        self.sleep_time_menu = ttk.Combobox(self, textvariable=self.sleep_time_var, values=self.sleep_time_options, state="readonly")
        self.sleep_time_menu.set('1000')  # 设置默认值
        self.sleep_time_menu.grid(row=3, column=0, padx=10, pady=5, sticky="e")
        
        # 确保添加对应的 grid_rowconfigure 调用以适应新组件...

        # 初始数据显示
        self.update_display()

        # 配置网格行/列的权重，确保文本框和下拉框可以随窗口大小调整而扩展
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # 添加前一步按钮
        self.prev_button = ttk.Button(self, text="Previous", command=self.go_back_to_empty_cc_1)
        self.prev_button.grid(row=3, column=0, sticky="sw", padx=10, pady=10)

        # 添加导出至Autokey脚本文件按钮
        self.export_button = ttk.Button(self, text="Export to Autokey script file", command=self.export_to_txt)
        self.export_button.grid(row=3, column=1, sticky="se", padx=10, pady=10)

        # 配置网格布局权重，确保按钮固定在底部边缘
        self.grid_rowconfigure(3, weight=1)
        self.grid_columnconfigure(1, weight=1)

    def update_display(self, event=None):
        # 解析选择的日期
        selected_date_str = self.date_option_var.get().replace("Update to date: ", "").replace(" (All data)", "")
        try:
            # 如果选择了有效的日期，则进行解析
            selected_date = pd.to_datetime(selected_date_str).date()
        except ValueError:
            # 如果未选择日期（或选择的是初始提示字符串），使用最大日期确保显示所有数据
            selected_date = self.combined_df['Last Count Date'].max()

        # 确定排序方式
        ascending = True if self.sort_option_var.get() == "Ascending" else False
        
        # 根据选择的日期过滤数据，并根据Location排序
        filtered_sorted_df = self.combined_df[self.combined_df['Last Count Date'] <= selected_date].sort_values(by="Location", ascending=ascending)

        # 构造显示的字符串：首先是行数信息，然后是过滤和排序后的DataFrame
        rows_count_info = f"Total rows displayed: {len(filtered_sorted_df)}\n"
        df_string = filtered_sorted_df.to_string(index=False)  # 转换DataFrame为字符串，省略索引
        display_string = rows_count_info + df_string  # 将行数信息和数据字符串合并

        # 更新文本框内容
        self.text_box.delete("1.0", tk.END)  # 清空当前文本框内容
        self.text_box.insert("1.0", display_string)  # 插入合并后的字符串

    def go_back_to_empty_cc_1(self):
        # 关闭当前窗口
        self.destroy()
        # 直接创建并显示一个Empty_CC_1窗口的实例
        # 注意: 这假设Empty_CC_1类已经在这个文件中被定义或正确导入
        empty_cc_1_window = Empty_CC_1(self.master, "800x600")
        # empty_cc_1_window.grab_set()


    def export_to_txt(self):
        # 获取用户选择的 Sleep 时间
        sleep_time = self.sleep_time_var.get()
        
        # 获取Location列的数据
        locations = self.combined_df['Location'].tolist()
        
        # 根据用户选择的 Sleep 时间构建脚本内容
        script_content = "#n::{\n"
        for location in locations:
            script_content += f'SendText "{location}"\n'
            script_content += 'Send "{Enter}"\n'
            script_content += f'Sleep {sleep_time}\n'
            script_content += 'Send "^A"\n'
            script_content += f'Sleep {sleep_time}\n'
            script_content += 'Send "^N"\n'
            script_content += f'Sleep {sleep_time}\n'
        script_content += "}"

        # 弹出保存文件对话框，让用户选择保存位置和文件名，指定后缀为.ahk
        filepath = filedialog.asksaveasfilename(defaultextension=".ahk", filetypes=[("AutoHotkey Scripts", "*.ahk")])
        if filepath:
            # 将脚本内容写入到文件
            with open(filepath, "w") as file:
                file.write(script_content)
            self.text_box.insert(tk.END, "\n\nPlease download and install the Autokey software from the official website. [.ahk] files can only be executed after installing the Autokey software.\n\
After double-clicking the [.ahk] file, a green [A] icon will appear in the bottom right corner of the taskbar. Press the [WIN] + [N] simultaneously to start running the script.\n\
This script needs to open the command-line interface of PDE gun in a Windows environment, select [1] -> [7], and then proceed with the script smoothly.\n\
As both this tool and the Autokey software lack screen locking functionality, you must ensure that the command-line interface of PDE gun remains on the top layer of the screen throughout the entire process. It's recommended not to perform any keyboard or mouse operations during the entire process.\n\
If you wish to terminate the execution during the process, you can right-click on the green [A] icon in the bottom right corner and directly close the opened [.ahk] file, or set a shortcut key to terminate the execution through the settings option of Autokey.\n\
It is advisable to close any opened [.ahk] files before running a new [.ahk] file.\n\
The installation of this software will not be blocked by the company's system policies. However, if used on a personal computer, be aware that some online games may detect the Autokey software as cheating software, leading to account bans.\n")
            self.text_box.see(tk.END)
            
            # 使用默认程序打开文件
            if os.path.exists(filepath):
                webbrowser.open(filepath)

if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
