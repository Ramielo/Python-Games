import tkinter as tk
from tkinter import ttk
import threading
import time
import pyautogui
import pygetwindow as gw
import base64
import tempfile

class RSV_1(tk.Tk): # 类定义和初始化
    def __init__(self):
        super().__init__()
        self.title("Receiving") # 窗口基础设置
        self.geometry("800x600")

        self.current_barcode_index = 0
        self.palletise_barcode_index = 0
        self.formatted_barcodes = []

        self.instructions_text = tk.Text(self, wrap="word", bg="white", height=10) # 提示文本框和滚动条
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.instructions_text.yview)
        self.instructions_text.configure(yscrollcommand=self.scrollbar.set)
        instructions = """Enter a list of SSCC in the text box below and click [Start receiving].
        
You should run this tool under: [WMS PROD RF][1. DJ Inbound], not under any of its subdirectories.\n"""
        self.instructions_text.insert("1.0", instructions)
        self.instructions_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.scrollbar.grid(row=0, column=1, sticky="ns", pady=10)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.data_input_text = tk.Text(self, wrap="word", bg="lightgray", height=15) # 数据输入文本框和滚动条
        self.data_input_scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.data_input_text.yview)
        self.data_input_text.configure(yscrollcommand=self.data_input_scrollbar.set)
        self.data_input_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.data_input_scrollbar.grid(row=1, column=1, sticky="ns", pady=10)
        self.grid_rowconfigure(1, weight=2)
        self.grid_columnconfigure(0, weight=1)

        # 创建 Sleep 时间选项的下拉菜单
        self.sleep_time_label = tk.Label(self, text="Please select the script execution speed (sleep time between each step, in milliseconds):")
        self.sleep_time_label.grid(row=2, column=0, padx=10, pady=(10, 0), sticky="nw")
        self.sleep_time_var = tk.StringVar(self)
        self.sleep_time_options = ['0.5', '1', '1.5', '2', '2.5', '3', '3.5', '4', '4.5', '5']  # 以毫秒为单位
        self.sleep_time_menu = ttk.Combobox(self, textvariable=self.sleep_time_var, values=self.sleep_time_options, state="readonly")
        self.sleep_time_menu.set('1')  # 设置默认值
        self.sleep_time_menu.grid(row=2, column=0, padx=10, pady=5, sticky="e")

        # 创建单选按钮的选中值变量
        self.palletise_label = tk.Label(self, text="Palletise all SSCC into PA000000 after receieving completed?")
        self.palletise_label.grid(row=3, column=0, padx=10, pady=(10, 0), sticky="nw")
        self.palletise_option = tk.IntVar()
        self.palletise_option.set(2)  # 默认选中第一个选项
        ttk.Radiobutton(self, text="Yes", variable=self.palletise_option, value=1).grid(row=3, column=0, sticky="ns", padx=20)
        ttk.Radiobutton(self, text="No", variable=self.palletise_option, value=2).grid(row=3, column=0, sticky="e", padx=20)

        self.generate_script_button = ttk.Button(self, text="Start receiving", command=self.generate_script)
        self.generate_script_button.grid(row=4, column=0, columnspan=2, pady=10, sticky="ns")

    def generate_script(self):
        # 从文本框获取数据并检查是否为空
        input_text = self.data_input_text.get("1.0", "end-1c").strip()
        if not input_text:
            # 直接在instructions_text文本框中显示警告信息
            self.instructions_text.insert("end", "\nNo SSCC list specified, please check the input box below.\n")
            self.instructions_text.see("end")  # 确保警告信息滚动到可视区域
            return  # 结束方法的执行
        
        # 如果输入框不为空，则继续处理数据
        self.barcodes = input_text.split("\n")
        self.formatted_barcodes = ["0" * (13 - len(barcode)) + barcode if len(barcode) < 13 else barcode for barcode in self.barcodes if barcode.strip()]
        
        threading.Thread(target=self.input_barcodes).start()

    def input_barcodes(self):
        sleep_time = float(self.sleep_time_var.get())
        base64_image_data = "iVBORw0KGgoAAAANSUhEUgAAADgAAAAICAYAAACs/DyzAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsEAAA7BAbiRa+0AAAJpSURBVEhLnVWvjvswDP7uHqOKThrNA0zVSNFQWUHJQHHRxgaKC8Y2VFwwMlBWVDQyVXuA0JOmqK/R+9J/6+7W3353n2QpsWPHjh37LS6Keg6DCtk6wNlJcfCslpOtESQzxMUG7RmiyrAMkm5DuDGKzSBtUF0z7KIEims3LmDErS0FGd7t47rHMsqHMw/4JvuXPiARxlvYwoJFkbm/hA0cA8rcuE5DWXM1IlnHaVzLMU+GdZqGtRundShHfBIvHdmQPFM82DQ6Ke253b631ct7erRzp1f6Rh67Iz1p/G/9fMenBsSMfMlMFShimoHZ6yYDPaRjQ58S5BcN22Hok1DIj1ln846yBPzG9t8wqS9D+KJElI+8VQpREIAJx3vDEB8mAoiqwpVrKT8g9K0RtZBwbI2jqYb8CG075ExAsly2fIyLOTzCOcIJPviqf8OU/kww+nO3+Yl3qBtzRb+agyeubTgO1/qzPWHgrmDrS5dRhYu2sfr2mJZ3QGEq4LCFOO34op1ghPzIn7ENpx/nBV7ru20V9sSMM4MMxBJYLebQtxxnloJtM8ARXMqs+WZQ3MwtzBePEZomsFwusd5rCH/Fq55AJdiVfByn2/8Wz/SHL2aQI6IPxo/9lU2TJccAFXQlIMQVpqrUjQpsRfrW17SLhciw7hRbWiMTi6dBqDyiEwKbif+mkhNf0EPXB3+NH/oMuhQsXfcxr/x0Dd5IputhwyDa9m/S7ENzZCQIkR56Y+0YSZSR92ODvEzD87oeP4wQ/sP0AA8Z9tobRkDf6pvR4uth3PzPmDCY0m/u45jwWFkt2Ev25psofAGNAGgVBZyj1QAAAABJRU5ErkJggg=="
        image_data = base64.b64decode(base64_image_data) # 输出临时图片文件以判断WARNING
        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_image:
            tmp_image.write(image_data)
            tmp_image_path = tmp_image.name

        if not self.activate_notepad_plus_plus(): # 找到wms window
            self.instructions_text.insert("end", "\nNo running [vipdjwmsapp.davidjones.com.au - PuTTY] window detected.\n")
            self.instructions_text.see("end")
            self.data_input_text.config(state='normal')
            return

        self.generate_script_button.config(state='disabled', text='Executing')
        self.data_input_text.config(state='disabled')
        
        time.sleep(sleep_time)
        while self.current_barcode_index < len(self.formatted_barcodes):
            barcode_to_send = self.formatted_barcodes[self.current_barcode_index]
            # pyautogui.write('1')
            # pyautogui.press('enter')
            # time.sleep(sleep_time)
            print ("type 1, Enter, sleep" , sleep_time," sec")
            # pyautogui.write(barcode_to_send)
            # pyautogui.press('enter')
            # time.sleep(sleep_time)
            print ("type barcode: ", barcode_to_send,", Enter, sleep" , sleep_time," sec")

            try:
                image_location = pyautogui.locateCenterOnScreen(tmp_image_path, confidence=0.5)
                if image_location:
                    self.text_box.insert(tk.END, "\n\nWARNING found, Ctrl + A sent to skip")
                    self.text_box.see(tk.END)
                    # pyautogui.hotkey('ctrl', 'a')
                    # time.sleep(sleep_time)
                    print ("Ctrl+A, 等待", sleep_time, "秒")
            except pyautogui.ImageNotFoundException: # 若有WARNING, 按Ctrl+A, 否则等2秒后继续
                # time.sleep(sleep_time * 2)
                print ("sleep" , sleep_time * 2," sec") 
            except Exception as e:
                print("发生错误：", e) # 输出其他类型的错误信息和堆栈跟踪
            
            # pyautogui.hotkey('ctrl', 'x')
            # time.sleep(sleep_time)
            print ("Ctrl+X, sleep" , sleep_time," sec")

            if not self.is_notepad_plus_plus_active():
                self.instructions_text.insert("end", f"\n[vipdjwmsapp.davidjones.com.au - PuTTY] window is deactivated or closed, execution proceeded up to SSCC: {barcode_to_send}.\n")
                self.instructions_text.see("end")
                self.generate_script_button.config(state='normal', text='Continue')
                break
            self.current_barcode_index += 1

        if self.palletise_option == 1 and self.palletise_barcode_index >= len(self.formatted_barcodes): # 开始Palletise
                        
            # pyautogui.write('5')
            # pyautogui.press('enter')
            # time.sleep(sleep_time)
            print ("type 5, Enter, sleep" , sleep_time," sec")
            # pyautogui.write('PA000000')
            # pyautogui.press('enter')
            # time.sleep(sleep_time)
            print ("type 5, Enter, sleep" , sleep_time," sec")

            try:
                image_location = pyautogui.locateCenterOnScreen(tmp_image_path, confidence=0.5)
                if image_location:
                    self.text_box.insert(tk.END, "\n\nWARNING found, Ctrl + A sent to skip")
                    self.text_box.see(tk.END)
                    # pyautogui.hotkey('ctrl', 'a')
                    # time.sleep(sleep_time)
                    print ("Ctrl+A, 等待", sleep_time, "秒")
            except pyautogui.ImageNotFoundException:
                pass # 忽略未找到图像的异常
            except Exception as e:
                print("发生错误：", e) # 输出其他类型的错误信息和堆栈跟踪

            while self.palletise_barcode_index < len(self.formatted_barcodes):
                barcode_to_send = self.formatted_barcodes[self.current_barcode_index]
                # pyautogui.write(barcode_to_send)
                # pyautogui.press('enter')
                # time.sleep(sleep_time)
                print ("type barcode: ", barcode_to_send,", Enter, sleep" , sleep_time," sec")

                try:
                    image_location = pyautogui.locateCenterOnScreen(tmp_image_path, confidence=0.5)
                    if image_location:
                        self.text_box.insert(tk.END, "\n\nWARNING found, Ctrl + A sent to skip")
                        self.text_box.see(tk.END)
                        # pyautogui.hotkey('ctrl', 'a')
                        # time.sleep(sleep_time)
                        print ("Ctrl+A, 等待", sleep_time, "秒")
                except pyautogui.ImageNotFoundException:
                    pass # 忽略未找到图像的异常
                except Exception as e:
                    print("发生错误：", e) # 输出其他类型的错误信息和堆栈跟踪

                try:
                    image_location = pyautogui.locateCenterOnScreen(tmp_image_path, confidence=0.5)
                    if image_location:
                        self.text_box.insert(tk.END, "\n\nWARNING found, Ctrl + A sent to skip")
                        self.text_box.see(tk.END)
                        # pyautogui.hotkey('ctrl', 'a')
                        # time.sleep(sleep_time)
                        print ("Ctrl+A, 等待", sleep_time, "秒")
                except pyautogui.ImageNotFoundException:
                    pass # 忽略未找到图像的异常
                except Exception as e:
                    print("发生错误：", e) # 输出其他类型的错误信息和堆栈跟踪

                if not self.is_notepad_plus_plus_active():
                    self.instructions_text.insert("end", f"\n[vipdjwmsapp.davidjones.com.au - PuTTY] window is deactivated or closed, execution proceeded up to SSCC: {barcode_to_send}.\n")
                    self.instructions_text.see("end")
                    break
                self.palletise_barcode_index += 1

        self.data_input_text.config(state='normal')
        if self.current_barcode_index >= len(self.formatted_barcodes):
            self.instructions_text.insert("end", "\nAll SSCC receiving completed.\n")
            self.instructions_text.see("end")
            self.current_barcode_index = 0  # 重置当前处理的条形码索引
            self.data_input_text.config(state='normal')
            self.generate_script_button.config(state='normal', text='Start receiving')

    def is_notepad_plus_plus_active(self):
        # 这个方法需要根据实际情况来实现，以检查vipdjwmsapp.davidjones.com.au - PuTTY是否处于激活状态
        try:
            notepad_window = gw.getWindowsWithTitle("vipdjwmsapp.davidjones.com.au - PuTTY")[0]
            return notepad_window.isActive
        except IndexError:
            return False

    def activate_notepad_plus_plus(self):
        try:
            notepad_window = gw.getWindowsWithTitle("vipdjwmsapp.davidjones.com.au - PuTTY")[0]
            notepad_window.activate()
            time.sleep(1)  # Give the window some time to become active
            return self.is_notepad_plus_plus_active()
        except IndexError:
            return False

    def on_window_close(self):
        self.thread.join()  # 等待线程结束
        self.destroy()  # 销毁窗口

if __name__ == "__main__":
    app = RSV_1()
    app.mainloop()
