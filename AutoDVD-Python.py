import pandas as pd
import datetime
import time

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

df = pd.read_excel('@ 項目_Python.xlsx')
alist = df.values.tolist()

name = ""
number = ""
writeword = sample

print(alist)

for iten in alist:
    if str(iten[0]) == "nan":
        break
    name = str(iten[0])
    number = str(iten[1])
    writeword = sample.replace("(片名)", name)
    writeword = writeword.replace("(數量)", number)
    nowtime = datetime.datetime.now().strftime('%Y-%m-%d_(%H-%M-%S)')
    with open(f'{nowtime}_{name}_{number}.job', 'w', encoding='utf-8') as f:
        f.write(f'{writeword}')
        time.sleep(1)
