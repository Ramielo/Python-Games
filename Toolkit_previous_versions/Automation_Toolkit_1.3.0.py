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
        self.title("Automation Toolkit 1.0")
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
        self.title("ET_CC_1")
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

        # 创建一个按钮用于导入.xls文件，并使用grid布局
        self.import_button = ttk.Button(self, text="Import .xls file", command=self.import_xls_file)
        self.import_button.grid(row=1, column=0, columnspan=2, pady=10, sticky="ew")

    def import_xls_file(self):
        # 一次性选择多个.xls文件
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xls")])
        if not file_paths:
            return  # 用户取消了选择

        # 清空之前的数据
        self.data_dict = {'Location': [], 'Barcode': [], 'Current Qty': []}
        self.text_box.delete('1.0', tk.END)

        for file_path in file_paths:
            try:
                # 读取指定的三列，确保Barcode作为字符串读取
                df = pd.read_excel(file_path, usecols=['Location', 'Barcode', 'Current Qty'], dtype={'Barcode': str})

                # 将数据追加到字典中
                for col in self.data_dict:
                    self.data_dict[col].extend(df[col].tolist())

            except Exception as e:
                self.text_box.insert(tk.END, f"Error importing file {file_path}: {e}\n")

        # 转换为DataFrame进行去重和排序
        df = pd.DataFrame(self.data_dict).drop_duplicates().sort_values(by='Location')

        # 将去重和排序后的DataFrame转换回字典
        self.data_dict = df.to_dict(orient='list')     

        # 在文本框中显示结果
        self.text_box.insert(tk.END, f"Total Rows Imported and Processed: {len(df)}\n")
        self.text_box.insert(tk.END, "First 50 rows (if available) after removing duplicates and sorting:\n")
        for i in range(min(50, len(df))):
            self.text_box.insert(tk.END, f"{df.iloc[i]['Location']}, {df.iloc[i]['Barcode']}, {df.iloc[i]['Current Qty']}\n")

        # 滚动到文本框底部
        self.text_box.see(tk.END)


class SVF_1(PredefinedWindow):
    def __init__(self, parent, window_size="800x600"):
        super().__init__(parent, "Shipment Verify-able Finder", window_size)

        # Create a textbox with a scrollbar for instructions
        self.instructions_text = tk.Text(self, wrap="word", bg="white", height=10)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.instructions_text.yview)
        self.instructions_text.configure(yscrollcommand=self.scrollbar.set)

        # Insert guidance text
        instructions = """This tool helps automatically identify shipments ready for verification, still in testing, so please sample check the results for accuracy. Report any errors to the developer.

Retrieve the following data from WMS and save it locally:

1. Inbound Shipment Inquiry: Set "Status From" and "Status To" to "In receiving". Export all data.

2. Case Inquiry: Follow 2 steps below.

Step 1 - Set "From Status" and "To Status" to "Consumed". Use "More Criteria" to set "From Consume Priority Date" to the same day a week ago. Result should be under 5000 rows.

* Notes for Step 1 - If needed for more coverage, set "From Consume Priority Date" even earlier, but manage how to export all since WMS limits exports to 5000 rows at a time. This tool auto-removes duplicates.

Step 2 - Set "From Status" and "To Status" to "In Inventory, Not Putaway", without setting "From Consume Priority Date". Export all data.

* Notes for Step 2 - This is to eliminate virtual case errors. Skipping Step 2 won't result in incorrectly verifying shipments that shouldn't be verified; it will only potentially miss shipments that could have been verified.

Import all data via the "Import Excel files" button for analysis.\n"""
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

        # Process each selected file for Case Inquiry
        for file_path in file_paths:
            if "Case Inquiry" in file_path:
                try:
                    df = pd.read_excel(file_path, usecols=["Case", "Inbound Shipment"])
                    # Remove duplicates and filter rows where 'Case' is at least 10 digits long
                    df = df.drop_duplicates(subset=['Case'])
                    df = df[df['Case'].astype(str).str.match(r'^\d{10,}$')]
                    
                    if "Inbound Shipment" not in df.columns:
                        raise ValueError(f"Column title does not match in {file_path}")

                    df["Inbound Shipment"] = df["Inbound Shipment"].apply(lambda x: str(x).rstrip('.0') if pd.notnull(x) else None)
                    df.dropna(subset=["Inbound Shipment"], inplace=True)
                    counts = df["Inbound Shipment"].value_counts().to_dict()

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
                datetime_col = [col for col in df_full.columns if all(keyword in col.lower() for keyword in ["first", "date", "time"])]
                
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
        # Initialize any necessary attributes here, such as for storing data

        # Create a textbox with a scrollbar for instructions
        self.instructions_text = tk.Text(self, wrap="word", bg="white", height=10)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.instructions_text.yview)
        self.instructions_text.configure(yscrollcommand=self.scrollbar.set)

        # Insert guidance text
        instructions = """This tool creates '.ahk' scripts for Autokey software, automating the typing of barcodes in SIM.

It only accepts .csv input files from these 4 sources:
1 - Claim Enquiry export
2 - Transfer Enquiry export
3 - RMS Enquiry export
4 - A clean list with [UPC] as the first column for Barcodes and [Qty] as the second for Quantities.

To import, click [Import CSV Files]. Be aware that new imports will replace existing data.\n"""
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
            self.instructions_text.insert(tk.END, """\nExport complete. You should now see a green [H] icon in the taskbar's notification area, indicating that the script is loaded.

Ensure Autokey software is installed to load the script. If not, you have the option to open the script with Notepad.

It is recommended to close any previously opened [.ahk] files before loading a new one.

This script automates the entry of barcode data for bulk processing in SIM's Claims or Transfers. For transfers, manually enter the first barcode due to system restrictions. To execute the script, press [WIN] + [N].

During the script's execution, keep the WMS PROD RF interface on top and avoid any keyboard or mouse use.

To terminate the script, close the [.ahk] file or use a pre-set Autokey shortcut key.\n""")
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
