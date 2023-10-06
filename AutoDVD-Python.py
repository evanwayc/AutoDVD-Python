import pandas as pd
import datetime
import time
import tkinter as tk
from tkinter.filedialog import askopenfilename

sample = ("[JOB]\n"
          "SourceType = \"ISO\"\n"
          "DiscType = \"DVD\"\n"
          "VolumeName = \"*\"\n"
          "Description = \"(片名)\"\n"
          "WritingSpeed = 0\n"
          "Quantity = (數量)\n"
          "[LABEL]\n"
          "Image = <Src: \"\\\\Art\\光碟出版機\\(片名).jpg\">\n"
          "[DATA]\n"
          "<Src: \"\\\\Art\\光碟出版機\\(片名).ISO\">\n")


def select_file():
    file_path = askopenfilename()
    entry.delete(0, tk.END)
    entry.insert(tk.END, file_path)
    return


def btn_start():
    df = pd.read_excel(entry.get())
    alist = df.values.tolist()

    # print(alist)

    id = 0

    for iten in alist:
        if str(iten[0]) == "nan":
            break
        name = str(iten[0])
        number = str(iten[1])
        writeword = sample.replace("(片名)", name)
        writeword = writeword.replace("(數量)", number)
        id += 1
        # nowtime = datetime.datetime.now().strftime('%Y-%m-%d_(%H-%M-%S)')
        with open(f'{str(id).zfill(2)}_{name}_{number}.job', 'w', encoding='utf-8') as f:
        # with open(f'{str(id).zfill(2)}_{nowtime}_{name}_{number}.job', 'w', encoding='utf-8') as f:
            f.write(f'{writeword}')
            # time.sleep(1)
    return


root = tk.Tk()
root.title("AutoDVD-Python")

entry = tk.Entry(root, width=20, font=("Arial", 16))
entry.grid(row=0, column=0, columnspan=4)

btn_select = tk.Button(root, text="瀏覽", command=select_file)
btn_select.grid(row=0, column=4, columnspan=1)

btn_start = tk.Button(root, text="執行", command=btn_start)
btn_start.grid(row=3, column=2, columnspan=1)

root.mainloop()
