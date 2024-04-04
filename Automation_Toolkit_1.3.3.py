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
from collections import Counter
import re
from datetime import datetime
import numpy as np
import calendar

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
        self.title("Automation Toolkit")
        self.geometry("400x300")  # 设置窗口大小
        
        self.open_subwindow_button = ttk.Button(self, text="AutoKey coding for Empty Location C/C", command=self.open_subwindow)
        self.open_subwindow_button.pack(pady=20)

        # Button to open the second subwindow for SIM barcode typing
        self.open_sim_bc_button = ttk.Button(self, text="AutoKey coding for SIM barcode typing", command=self.open_sim_bc_subwindow)
        self.open_sim_bc_button.pack(pady=20)

        # 新增按钮用于打开 Shipment Verify-able Finder 窗口
        self.open_svf_button = ttk.Button(self, text="Shipment Verify-able Finder", command=self.open_svf_subwindow)
        self.open_svf_button.pack(pady=20)

        # 新增按钮用于打开 AutoKey coding for C/C Everything 窗口
        self.open_et_cc_button = ttk.Button(self, text="AutoKey coding for C/C Everything", command=self.open_et_cc_subwindow)
        self.open_et_cc_button.pack(pady=20)

        # 创建一个标签用作联系开发者的链接
        self.contact_developer_label = tk.Label(self, text="Contact developer: charles.liang@davidjones.com.au", fg="blue", cursor="hand2")
        self.contact_developer_label.pack(side="bottom", pady=10)
        self.contact_developer_label.bind("<Button-1>", self.open_mail_to)

    def open_mail_to(self, event=None):
        # 使用默认邮件客户端打开一个新邮件窗口，预填充收件人地址
        webbrowser.open("mailto:charles.liang@davidjones.com.au")

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

    def open_svf_subwindow(self):
        # 实例化 SVF_1 类
        subwindow = SVF_1(self)
        self.iconify()
        subwindow.protocol("WM_DELETE_WINDOW", lambda: (subwindow.destroy(), self.deiconify()))

    def open_et_cc_subwindow(self):
        # 实例化 ET_CC_1 类
        subwindow = ET_CC_1(self)
        self.iconify()
        subwindow.protocol("WM_DELETE_WINDOW", lambda: (subwindow.destroy(), self.deiconify()))

class ET_CC_1(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("AutoKey coding for C/C Everything")
        self.geometry("800x600")

        self.data_dict = {'Location': [], 'Barcode': [], 'Current Qty': []}  # 新建一个空字典

        # 创建一个文本框和一个滚动条
        self.text_box = tk.Text(self, wrap=tk.WORD, height=20, width=80)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.text_box.yview)
        self.text_box.configure(yscrollcommand=self.scrollbar.set)

        # 布局文本框和滚动条
        self.text_box.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.scrollbar.grid(row=0, column=1, sticky="ns")

        # 配置网格行/列权重，确保文本框随窗口大小调整而扩展
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # 在文本框中添加初始描述文本
        initial_text = """
This tool is designed for skimping on cycle counts; it automatically generates a script based on the stock quantities displayed in the system for all locations imported. Once executed, it will perform the cycle count automatically according to the quantities shown in the system. However, do not attempt to cut corners unless you and your manager have decided to do so.

Even then, since the WMS PROD RF interface often freezes, you might still encounter descrepencies or locations that weren't counted in the cycle count results. You will need to find a way to extract the stock data for these locations and re-perform another round by this tool. You can mitigate this risk by slowing down the execution speed through adjusting the interval milliseconds.

WARNING!!!

Please make sure you know what you are doing before using this script; otherwise, please revert to the previous step.\n"""
        self.text_box.insert(tk.END, initial_text)

        # 创建一个标签和下拉选项
        self.speed_label = ttk.Label(self, text="Script execution speed (ms):")
        self.speed_label.grid(row=1, column=0, sticky="e", pady=10)
        self.sleep_time_combobox = ttk.Combobox(self, values=[500, 1000, 1500, 2000, 2500, 3000, 3500, 4000, 4500, 5000])
        self.sleep_time_combobox.grid(row=1, column=1, sticky="w", pady=10)
        self.sleep_time_combobox.set(1000)  # 默认值为1000毫秒

        # 创建一个按钮用于导入.xls文件，并使用grid布局
        self.import_button = ttk.Button(self, text="Import .xls files", command=self.import_xls_file)
        self.import_button.grid(row=1, column=0, columnspan=2, pady=10, sticky="ew")


    def import_xls_file(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xls")])
        if not file_paths:
            return  # 用户取消了选择

        self.data_dict = {'Location': [], 'Barcode': [], 'Current Qty': []}

        for file_path in file_paths:
            try:
                df = pd.read_excel(file_path, usecols=['Location', 'Barcode', 'Current Qty'], dtype={'Barcode': str, 'Location': str})
                for col in self.data_dict:
                    self.data_dict[col].extend(df[col].tolist())
            except Exception as e:
                self.text_box.insert(tk.END, f"Error importing file {file_path}: {e}\n")

        df = pd.DataFrame(self.data_dict).drop_duplicates().sort_values(by='Location')
        self.data_dict = df.to_dict(orient='list')

        # 让用户选择保存commands.ahk的位置
        output_filename = filedialog.asksaveasfilename(defaultextension=".ahk", filetypes=[("AutoHotkey Script", "*.ahk")])
        if not output_filename:  # 用户取消了选择
            return

        sleep_time = self.sleep_time_combobox.get()  # 获取用户选择的sleep时间
        
        with open(output_filename, 'w') as file:
            file.write("#n::{\n")

            last_location = None
            for i in range(len(df)):
                location = df.iloc[i]['Location']
                barcode = df.iloc[i]['Barcode']
                qty_str = str(df.iloc[i]['Current Qty'])
                qty_str = qty_str if qty_str != 'nan' else '0'
                qty = int(qty_str.split()[0])

                if location != last_location and last_location is not None:
                    # 如果location变化了，添加Send "^N"和Sleep，使用用户选择的sleep时间
                    file.write('Send "^N"\nSleep {}\n'.format(sleep_time))

                if location != last_location:
                    file.write(f'SendText "{location}"\nSend "{{Enter}}"\nSleep {sleep_time}\n')
                    file.write(f'Send "^A"\nSleep {sleep_time}\nSend "^A"\nSleep {sleep_time}\n')

                if pd.isna(barcode) or barcode == '':
                    file.write(f'Send "^N"\nSleep {sleep_time}\n')
                else:
                    for _ in range(qty):
                        file.write(f'SendText "{barcode}"\nSend "{{Enter}}"\nSleep {sleep_time}\n')

                last_location = location

            file.write("}\n")

        self.text_box.insert(tk.END, f"""\nCommands have been written to {output_filename}.

WARNING!!!

Please make sure you know what you are doing before execute this script; otherwise, please close the [.ahk] file immediately to terminate the script, \n""")
        self.text_box.see(tk.END)

        try:
            file_path = os.path.abspath(output_filename)
            webbrowser.open(file_path)
        except Exception as e:
            self.text_box.insert(tk.END, f"Error opening file {output_filename}: {e}\n")






class SVF_1(PredefinedWindow):
    def __init__(self, parent, window_size="800x600"):
        super().__init__(parent, "Shipment Verify-able Finder", window_size)

        # Create a textbox with a scrollbar for instructions
        self.instructions_text = tk.Text(self, wrap="word", bg="white", height=10)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.instructions_text.yview)
        self.instructions_text.configure(yscrollcommand=self.scrollbar.set)

        # Insert guidance text
        instructions = """This tool helps automatically identify shipments ready for verification.

Retrieve the following data from WMS and save it locally:

Inbound Shipment Inquiry: Set "Status From" and "Status To" to "In receiving". Export all data. (*1)

Case Inquiry: Set "From Status" and "To Status" to "Consumed". Use "More Criteria" to set "From Consume Priority Date" to the same day a week ago. The result should be under 5000 rows. (*2)

Import all data sheets via the "Import Excel files" button for analysis.

Additional notice:

(*1) If needed for more coverage, set "From Consume Priority Date" even earlier, but manage how to export all since WMS limits exports to 5000 rows at a time. This tool auto-removes duplicates.

(*2) To eliminate virtual case errors, set "From Status" and "To Status" to "In Inventory, Not Putaway", without setting "From Consume Priority Date". Import as additional data. Skipping this step won't result in incorrectly verifying shipments that shouldn't be verified; it will only potentially miss shipments that could have been verified.

About the data format: I am trying to achieve no specific requirement for the data format. If you ever encounter any data format error, send all related data sheets to charles.liang@davidjones.com.au. I believe most of them will be solved very quickly.\n"""
        self.instructions_text.insert("1.0", instructions)

        # Layout the textbox and scrollbar
        self.instructions_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.scrollbar.grid(row=0, column=1, sticky="ns", pady=10)

        # Configure grid row/column weights to ensure the textbox expands with the window size
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Create the "Import Excel Files" button and layout using grid
        self.import_button = ttk.Button(self, text="Import Excel Files", command=self.import_excel)
        self.import_button.grid(row=1, column=0, columnspan=2, pady=10, sticky="ew")
        self.data_dict = {}

    def import_excel(self):
        # Clear the textbox and display the starting message
        self.instructions_text.insert(tk.END, "\nStarting to import data...\n")
        self.update_idletasks()  # Update UI

        # Open file dialog to select Excel files
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel 97-2003 Workbook", "*.xls")])
        if not file_paths:
            self.instructions_text.insert(tk.END, "\nData import cancelled.\n")
            return

        case_inquiry_data = {}
        Inbound_Shipment_inquiry_data = {}

        inbound_shipment_inquiry_files = [file for file in file_paths if "Inbound Shipment Inquiry" in file]
        if len(inbound_shipment_inquiry_files) > 1:
            self.instructions_text.insert(tk.END, "\nError: More than one 'Inbound Shipment Inquiry' file found.\n")
            return
        elif len(inbound_shipment_inquiry_files) == 1:
            try:
                df_inbound = pd.read_excel(inbound_shipment_inquiry_files[0], usecols=["Inbound Shipment", "Cases Shipped", "Cases Received"])
                Inbound_Shipment_inquiry_data = df_inbound.to_dict(orient='list')
            except Exception as e:
                self.instructions_text.insert(tk.END, f"\nError processing file {inbound_shipment_inquiry_files[0]}: {e}\n")

        for file_path in file_paths:
            if "Case Inquiry" in file_path:
                try:
                    df = pd.read_excel(file_path)
                    
                    # 确定“Inbound Shipment”列
                    shipment_col = None
                    for col in df.columns:
                        if "shipment" in col.lower() or "shpmt" in col.lower():
                            shipment_col = col
                            break

                    if shipment_col is None:
                        raise ValueError(f"Shipment column not found in {file_path}")
                    
                    # 对“Case”列和“Inbound Shipment”列进行数据处理
                    df = df.drop_duplicates(subset=['Case'])
                    df = df[df['Case'].astype(str).str.match(r'^\d{10,}$')]
                    
                    df[shipment_col] = df[shipment_col].apply(lambda x: str(x).rstrip('.0') if pd.notnull(x) else None)
                    df.dropna(subset=[shipment_col], inplace=True)
                    
                    counts = df[shipment_col].value_counts().to_dict()

                    for key, value in counts.items():
                        if key in case_inquiry_data:
                            case_inquiry_data[key] += value
                        else:
                            case_inquiry_data[key] = value
                except Exception as e:
                    self.instructions_text.insert(tk.END, f"\nError processing file {file_path}: {e}\n")


        # Add the fourth column for matches with Case Inquiry data
        if Inbound_Shipment_inquiry_data:
            Inbound_Shipment_inquiry_data['Matched Case Counts'] = [
                case_inquiry_data.get(str(shipment), 0) for shipment in Inbound_Shipment_inquiry_data['Inbound Shipment']
            ]

        if len(inbound_shipment_inquiry_files) == 1:
            try:
                # 加载整个文件以查找特定列
                df_full = pd.read_excel(inbound_shipment_inquiry_files[0])
                # 寻找同时包含"first", "date", "time"的列标题
                datetime_col = [col for col in df_full.columns if all(keyword in col.lower() for keyword in ["first", "date", "time"]) or all(keyword in col.lower() for keyword in ["received", "date"])]
                
                if datetime_col:
                    # 假定只有一个符合条件的列，获取其数据
                    datetime_data = df_full[datetime_col[0]].tolist()
                    # 添加到Inbound_Shipment_inquiry_data字典中
                    Inbound_Shipment_inquiry_data['First DateTime'] = datetime_data
            except Exception as e:
                self.instructions_text.insert(tk.END, f"\nError processing additional data from file {inbound_shipment_inquiry_files[0]}: {e}\n")


        # Display results
        self.instructions_text.insert(tk.END, "\nCase Inquiry Data:\n")
        for key, value in case_inquiry_data.items():
            self.instructions_text.insert(tk.END, f"{key}: {value}\n")

        if Inbound_Shipment_inquiry_data:
            # 将字典转换为DataFrame
            df_inbound = pd.DataFrame(Inbound_Shipment_inquiry_data)
            # 删除重复行
            df_inbound = df_inbound.drop_duplicates()
            # 更新Inbound_Shipment_inquiry_data字典，以便于展示和之后的处理
            Inbound_Shipment_inquiry_data = df_inbound.to_dict(orient='list')

            self.instructions_text.insert(tk.END, "\nInbound Shipment Inquiry Data:\n")
            # 检查是否存在'First DateTime'列
            has_datetime = 'First DateTime' in Inbound_Shipment_inquiry_data
            for i in range(len(Inbound_Shipment_inquiry_data['Inbound Shipment'])):
                # 构建要插入的文本行
                text_line = f"{Inbound_Shipment_inquiry_data['Inbound Shipment'][i]}, {Inbound_Shipment_inquiry_data['Cases Shipped'][i]}, {Inbound_Shipment_inquiry_data['Cases Received'][i]}, {Inbound_Shipment_inquiry_data['Matched Case Counts'][i]}"
                # 如果存在'First DateTime'数据，则添加到文本行
                if has_datetime:
                    text_line += f", {Inbound_Shipment_inquiry_data['First DateTime'][i]}"
                self.instructions_text.insert(tk.END, text_line + "\n")

        self.instructions_text.insert(tk.END, "\nData import completed.\n")

        # 在你的方法最末尾加入以下代码，紧接在展示结果到GUI的代码之后

        def summarize_and_write_to_txt(Inbound_Shipment_inquiry_data):
            # Convert dictionary to DataFrame
            df = pd.DataFrame(Inbound_Shipment_inquiry_data)
            
            # Ensure correct datetime format
            df['First DateTime'] = pd.to_datetime(df['First DateTime'], format='%d/%m/%Y %H:%M')
            
            # Calculate today's date
            today = pd.to_datetime("today")
            
            # Shipments that can be VERIFIED
            verified_shipments = df[(df['Cases Shipped'] == df['Cases Received']) & (df['Cases Received'] == df['Matched Case Counts'])]['Inbound Shipment']
            
            # Shipments that can be verified but not all have been received
            partial_verified_shipments = df[(df['Cases Received'] == df['Matched Case Counts']) & (df['Cases Received'] < df['Cases Shipped'])]['Inbound Shipment']
            
            # Shipments that are about to expire
            upcoming_expired_shipments = df[df['First DateTime'] < (today - pd.Timedelta(days=2))]['Inbound Shipment']
            
            # Write summary to TXT file
            txt_file_path = 'shipment_summary.txt'
            with open(txt_file_path, 'w', encoding='utf-8') as file:
                file.write("Shipments that can be VERIFIED:\n")
                file.write('\n'.join(verified_shipments.astype(str)) + '\n\n')
                
                file.write("Shipments that can be verified but not all have been received:\n")
                file.write('\n'.join(partial_verified_shipments.astype(str)) + '\n\n')
                
                file.write("Shipments that are about to expire:\n")
                file.write('\n'.join(upcoming_expired_shipments.astype(str)) + '\n\n')

            # Insert summary information into the GUI
            self.instructions_text.insert(tk.END, "\nSummary information has been written to 'shipment_summary.txt'.\n")

            # Open the TXT file automatically after saving
            webbrowser.open(txt_file_path)

        # Assuming Inbound_Shipment_inquiry_data is ready
        # Call this function at the end of the method
        if Inbound_Shipment_inquiry_data:
            summarize_and_write_to_txt(Inbound_Shipment_inquiry_data)

        self.instructions_text.insert(tk.END, "\nData import and summarization completed.\n")

        # Scroll to the bottom of the textbox
        self.instructions_text.see(tk.END)


class SIM_BC_1(PredefinedWindow):
    def __init__(self, parent, window_size="800x600"):
        super().__init__(parent, "AutoKey coding for SIM barcode typing", window_size)
        
        self.instructions_text = tk.Text(self, wrap="word", bg="white", height=10)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.instructions_text.yview)
        self.instructions_text.configure(yscrollcommand=self.scrollbar.set)
        
        instructions = """Enter a list of barcodes in the text box below and click "Generate Script" to create and save a script.
        
If Autokey software is installed on your computer, the script will automatically load. If not, you will be prompted to open it in another way (e.g., text editor).
        
This script is useful for entering a large number of different barcodes at once in SIM, especially suitable for creating Claims or Transfers.\n"""
        
        self.instructions_text.insert("1.0", instructions)
        
        self.instructions_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.scrollbar.grid(row=0, column=1, sticky="ns", pady=10)
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        self.data_input_text = tk.Text(self, wrap="word", bg="lightgray", height=15)
        self.data_input_scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.data_input_text.yview)
        self.data_input_text.configure(yscrollcommand=self.data_input_scrollbar.set)
        self.data_input_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.data_input_scrollbar.grid(row=1, column=1, sticky="ns", pady=10)
        
        self.grid_rowconfigure(1, weight=2)
        self.grid_columnconfigure(0, weight=1)
        
        self.generate_script_button = ttk.Button(self, text="Generate Script", command=self.generate_script)
        self.generate_script_button.grid(row=2, column=0, columnspan=2, pady=10, sticky="ew")

    def generate_script(self):
        barcodes = self.data_input_text.get("1.0", "end-1c").split("\n")
        
        # 检查所有输入是否为数字
        if not all(barcode.isdigit() for barcode in barcodes if barcode.strip()):
            # 在指令文本框内容后追加错误消息
            error_message = "\nInput error: Please ensure all inputs are numbers.\n"
            self.instructions_text.insert("end", error_message)  # 更新了这里
            # 滚动到指令文本框底部
            self.instructions_text.see("end")  # 更新了这里
            return
        
        script_content = "#n::{\n" + "\n".join(f'SendText "{barcode}"\nSleep 500' for barcode in barcodes if barcode.strip()) + "\n}"
        
        # 保存文件对话框
        filepath = filedialog.asksaveasfilename(defaultextension=".ahk", filetypes=[("AutoHotkey scripts", "*.ahk")])
        if not filepath:
            return  # 用户取消了保存操作
        
        with open(filepath, 'w') as file:
            file.write(script_content)
        
        # 在指令文本框内容后追加成功消息
        success_message = "\nScript generated successfully.\n\nSaved to: " + filepath + "\n\nTo execute the script, press [WIN] + [N].\n"
        self.instructions_text.insert("end", success_message)  # 更新了这里
        # 滚动到指令文本框底部
        self.instructions_text.see("end")  # 更新了这里
        
        # 打开保存的文件
        os.startfile(filepath)



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
        instructions = """This tool automates Cycle Counts for empty locations.

In WMS's [Pick Location Inquiry], enter a zone number (such as 01, 09, 15) and press the [Enter] key.

Due to WMS system limits, the maximum exportable rows per file is 5000, excluding the title row.

For zones with 5000-10000 rows, first export locations in ascending order, then descending. Reverse order by clicking the [Location] column title.

For zones with over 10000 rows, export in batches by setting [Bay] from 0 to 9.

Save data locally, then select [Import Excel Files]. The tool supports importing multiple Excel files and automatically removes duplicates.\n"""
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

        # 格式化Location列以保留前导零，假设Location的最大长度为9位
        filtered_sorted_df['Location'] = filtered_sorted_df['Location'].apply(lambda x: f"{x:09}")

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
        
        # 获取并格式化Location列的数据，确保保留前导零
        # 这里使用zfill()方法确保字符串长度，以9为例
        locations = self.combined_df['Location'].apply(lambda x: str(x).zfill(9)).tolist()
        
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
            self.text_box.insert(tk.END, """\n\nExport complete. You should now see a green [H] icon in the taskbar's notification area, indicating that the script is loaded.

Ensure Autokey software is installed to load the script. If not, you have the option to open the script with Notepad.

It is recommended to close any previously opened [.ahk] files before loading a new one.

To execute the script, log into WMS PROD RF on a Windows system, select [1. DJ Inbound] [7. DJ CC Manual], and press [WIN] + [N].

During the script's execution, keep the WMS PROD RF interface on top and avoid any keyboard or mouse use.

To terminate the script, close the [.ahk] file or use a pre-set Autokey shortcut key.\n""")
            self.text_box.see(tk.END)
            
            # 使用默认程序打开文件
            if os.path.exists(filepath):
                webbrowser.open(filepath)

if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
