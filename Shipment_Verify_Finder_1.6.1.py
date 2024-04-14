import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
from datetime import datetime
import os
import webbrowser
import numpy as np

class SVF_1(tk.Tk):
    def __init__(self, window_size="800x600"):
        super().__init__()
        self.title("Shipment Verify-able Finder")
        self.geometry(window_size)


        # Create a textbox with a scrollbar for instructions
        self.instructions_text = tk.Text(self, wrap="word", bg="white", height=10)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.instructions_text.yview)
        self.instructions_text.configure(yscrollcommand=self.scrollbar.set)

        # Insert guidance text
        instructions = """This tool identifies shipments ready for verification.

Retrieve the following data from WMS:

[Inbound Shipment Inquiry]: Set [Status From] and [Status To] to [In receiving]. 

[Case Inquiry]: Set [From Status] and [To Status] to [Consumed]. Use [More Criteria] to set [From Consume Priority Date].

Import all data sheets at once using the [Import Excel Files] function.

Please consult the release notes for data format and additional information.\n"""

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
        # 预分类文件
        files_cases_consumed = []
        files_cases_virtual = []
        files_shipment = []

        for file_path in file_paths:
            df = pd.read_excel(file_path, nrows=2)  # 只加载前两行进行检查
            if 'Consumed' in df.iloc[1].to_string():
                files_cases_consumed.append(file_path)
            elif 'In Inventory, Not Putaway' in df.iloc[1].to_string():
                files_cases_virtual.append(file_path)
            elif 'In Receiving' in df.iloc[1].to_string():
                files_shipment.append(file_path)

        if files_cases_consumed == []:
            self.instructions_text.insert(tk.END, "\nError: 'Case Inquiry' files include no consumed data.\n")
            return
        
        if files_shipment == []:
            self.instructions_text.insert(tk.END, "\nError: 'Inbound Shipment' file should include only 'In Receiving' status.\n")
            return

        Inbound_Shipment_inquiry_data = {}

        inbound_shipment_inquiry_files = [file for file in file_paths if "Inbound Shipment Inquiry" in file]
        if len(inbound_shipment_inquiry_files) != 1:
            self.instructions_text.insert(tk.END, "\nError: 'Inbound Shipment Inquiry' file must exist and be unique.\n")
            return
        elif len(inbound_shipment_inquiry_files) == 1:
            try:
                df_inbound = pd.read_excel(inbound_shipment_inquiry_files[0], usecols=["Inbound Shipment", "Cases Shipped", "Cases Received"])
                Inbound_Shipment_inquiry_data = df_inbound.to_dict(orient='list')

            except Exception as e:
                self.instructions_text.insert(tk.END, f"\nError processing file {inbound_shipment_inquiry_files[0]}: {e}\n")
                return

        df_cases_consumed_combined = pd.concat([pd.read_excel(file, dtype=str) for file in files_cases_consumed], ignore_index=True).drop_duplicates()
        # 读取files_cases_consumed中的excel，忽略标题行，去重，读取为str
        df_cases_consumed_combined = df_cases_consumed_combined.drop_duplicates(subset=['Case']) # 按第一列Case number去重
        df_cases_consumed_combined = df_cases_consumed_combined[df_cases_consumed_combined['Case'].astype(str).str.match(r'^\d{10,}$')]
        # 仅保留Case列=[长度至少为10位的纯数字字符串]的行

        shipment_col = None
        for col in df_cases_consumed_combined.columns:
            if "shipment" in col.lower() or "shpmt" in col.lower():
                shipment_col = col # 找到cases_consumed表的Inbound Shipment列
                break
        if shipment_col is None:
            raise ValueError(f"Shipment column not found in {file_path}")
        
        df_cases_consumed_combined.dropna(subset=[shipment_col], inplace=True) # 删除Inbound Shipment列为空的行
        counts_cases_consumed = df_cases_consumed_combined[shipment_col].value_counts().to_dict() # 统计Inbound Shipment列

        if len(files_cases_virtual):
            df_cases_virtual_combined = pd.concat([pd.read_excel(file, dtype=str) for file in files_cases_virtual], ignore_index=True).drop_duplicates()
            # 读取files_cases_virtual中的excel，忽略标题行，去重，读取为str
            df_cases_virtual_combined = df_cases_virtual_combined.drop_duplicates(subset=['Case']) # 按第一列Case number去重
            df_cases_virtual_combined = df_cases_virtual_combined[df_cases_virtual_combined['Case'].astype(str).str.match(r'^[a-zA-Z0-9]{1,9}$')]
            # 筛选出“Case”列中符合新格式要求的行：包含数字或字母，长度不足10
            shipment_col_v = None
            for col in df_cases_virtual_combined.columns:
                if "shipment" in col.lower() or "shpmt" in col.lower():
                    shipment_col_v = col # 找到cases_virtual表的Inbound Shipment列
                    break
            if shipment_col_v is None:
                raise ValueError(f"Shipment column not found in {file_path}")

            df_cases_virtual_combined.dropna(subset=[shipment_col_v], inplace=True) # 删除Inbound Shipment列为空的行
            counts_cases_virtual = df_cases_virtual_combined[shipment_col_v].value_counts().to_dict() # 统计Inbound Shipment列
        
        Inbound_Shipment_inquiry_data['Cases Consumed'] = [
            counts_cases_consumed.get(str(shipment), 0) for shipment in Inbound_Shipment_inquiry_data['Inbound Shipment']
        ] # 写入Cases Consumed列

        received_rate=[int((a / b) * 100) if b != 0 else None for a, b in zip(Inbound_Shipment_inquiry_data['Cases Received'], Inbound_Shipment_inquiry_data['Cases Shipped'])]
        consumed_rate=[int((a / b) * 100) if b != 0 else None for a, b in zip(Inbound_Shipment_inquiry_data['Cases Consumed'], Inbound_Shipment_inquiry_data['Cases Received'])]
        Inbound_Shipment_inquiry_data['Consumed %'] = consumed_rate # 写入Consumed %列
        Inbound_Shipment_inquiry_data['Received %'] = received_rate # Received %列

        if len(files_cases_virtual):
            Inbound_Shipment_inquiry_data['Virtual Cases'] = [
                counts_cases_virtual.get(str(shipment), 0) for shipment in Inbound_Shipment_inquiry_data['Inbound Shipment']
            ] # 写入Cases Virtual列


        df_full = pd.read_excel(inbound_shipment_inquiry_files[0])

        datetime_data = []
        for column in df_full.columns:
            if all(word in column for word in ['first', 'date', 'time']) or all(word in column for word in ['received', 'date']):
                datetime_data = df_full[column]
                break  # 找到符合条件的第一个列，获取其数据后终止循环


        if len(datetime_data) > 0:
            now = datetime.now() # 现在时刻
            time_differences = [] # 用以存储时间差数据

            for date_str in datetime_data:
                try:
                    date_time = datetime.strptime(date_str, "%d/%m/%Y %H:%M") # 将字符串转换为datetime对象
                    workday_count = np.busday_count(date_time.strftime("%Y-%m-%d"), now.strftime("%Y-%m-%d")) # 计算工作日差
                    time_diff_days = ((now.hour - date_time.hour) * 3600 + (now.minute - date_time.minute) * 60) / 86400
                    delta = round(workday_count + time_diff_days, 2) # 计算时间差，以day为单位的浮点数

                except:
                    delta = None

                time_differences.append(delta) # 写入时间差数据

            # 添加到Inbound_Shipment_inquiry_data字典中
            Inbound_Shipment_inquiry_data['Days after received'] = time_differences
        else:
            self.instructions_text.insert(tk.END, "\nNo datetime column found in 'Inbound Shipment Inquiry' file.\n")
        
        df = pd.DataFrame(Inbound_Shipment_inquiry_data)
        df['Consumed %'] = df['Consumed %'].astype(str) + ' %'
        df['Received %'] = df['Received %'].astype(str) + ' %'
        if len(datetime_data) > 0:
            df_sorted = df.sort_values(by='Days after received', ascending=False)
        else: df_sorted = df
        df_final = df_sorted.drop_duplicates()
        
        self.instructions_text.insert(tk.END, "\nData import completed.\n")

        txt_file_path = 'shipment_summary.txt' # 开始生成report
        full_path = os.path.abspath(txt_file_path)  # 转换为绝对路径

        with open(txt_file_path, 'w', encoding='utf-8') as file:
            file.write("Shipments that can be VERIFIED:\n\n")
            filtered_verify_OK = df_final[(df_final['Consumed %'] == '100 %') & (df_final['Received %'] == '100 %')]
            file.write(filtered_verify_OK.to_string(index=False, header=True))
            
            filtered_verify_check = df_final[(df_final['Consumed %'] == '100 %') & (df_final['Received %'] != '100 %')]
            if len(filtered_verify_check):
                file.write("\n\n\nShipments that can be verified but not all have been received:\n\n")
                file.write(filtered_verify_check.to_string(index=False, header=True))

            indexes_to_remove = pd.Index([]) 
            if not filtered_verify_OK.empty:
                indexes_to_remove = indexes_to_remove.union(filtered_verify_OK.index)
            if not filtered_verify_check.empty:
                indexes_to_remove = indexes_to_remove.union(filtered_verify_check.index)
            df_final_filtered = df_final[~df_final.index.isin(indexes_to_remove)] # 移除已经输出的行

            filtered_SLA = []
            if len(datetime_data) > 0:
                file.write("\n\n\nShipments over SLA = 0.85 day:\n\n")
                filtered_SLA = df_final_filtered[df_final_filtered['Days after received'].astype(float) > 0.85]
                file.write(filtered_SLA.to_string(index=False, header=True))

            if len(filtered_SLA) > 0:
                indexes_to_remove = indexes_to_remove.union(filtered_SLA.index)
            df_final_filtered = df_final[~df_final.index.isin(indexes_to_remove)] # 再次移除已经输出的行

            file.write("\n\n\nAll other Shipments:\n\n")
            file.write(df_final_filtered.to_string(index=False, header=True))

            # Open the TXT file automatically after saving
            webbrowser.open(txt_file_path)

if __name__ == "__main__":
    app = SVF_1()  # 直接创建 SVF_1 实例，这将是主窗口
    app.mainloop()  # 启动事件循环