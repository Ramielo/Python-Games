import tkinter as tk
from tkinter import filedialog
from PIL import Image
import base64
from io import BytesIO

def convert_image():
    file_path = filedialog.askopenfilename(filetypes=[("PNG files", "*.png")])
    if file_path:
        with Image.open(file_path) as img:
            buffered = BytesIO()
            img.save(buffered, format="PNG")
            img_base64 = base64.b64encode(buffered.getvalue()).decode("utf-8")
        
        text.delete('1.0', tk.END)
        text.insert('1.0', img_base64)

app = tk.Tk()
app.title('PNG to Base64 Converter')

text = tk.Text(app, height=10, width=80)
text.pack()

convert_button = tk.Button(app, text='Convert PNG to Base64', command=convert_image)
convert_button.pack()

app.mainloop()
