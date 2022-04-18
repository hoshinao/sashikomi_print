import glob
from tkinter import messagebox
from docx import Document
import win32com.client
import openpyxl as px
import os
import os, tkinter, tkinter.filedialog, tkinter.messagebox
from tqdm import tk
import shutil
from pathlib import Path
import PyPDF2
import subprocess
import time

# 前回のpdf保存ファイルをクリア
shutil.rmtree("print_list")
os.mkdir("print_list")

# ファイル選択ダイアログの表示
root = tkinter.Tk()
root.withdraw()
fTyp = [("","*")]
iDir = os.path.abspath(os.path.dirname("L:/Z99_個別ﾌｫﾙﾀﾞ/230_星川/差し込み印刷pg/"))
tkinter.messagebox.showinfo('ファイル選択','処理データ(Excel)を選択してください')
filepath = tkinter.filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)


# 「.xlsx」を開いて「Sheet1」を読み込む
wb = px.load_workbook(filename=filepath)
ws1 = wb['Sheet1']

# Excelの「Sheet1」の全データをリストとして取得
values1 = [[cell.value for cell in row1] for row1 in ws1]

# Pythonファイルが保管されているフォルダパスを取得
curdir = os.getcwd()

# win32comでwordアプリケーションを起動
word = win32com.client.Dispatch("Word.Application")
word.Visible = True
word.DisplayAlerts = False

# Excelリストを2行目から最終行まで繰り返す
for i in range(1, len(values1)):

    # テンプレートとなるwordファイルを呼び出す
    templatepath = '差し込み印刷サンプルword.docx'
    doc = Document(templatepath)

    # プログラム3のリストを辞書に変換
    dic = dict(zip(values1[0], values1[i]))

    # テンプレートファイル(Word)の文言をプログラム8の辞書データで置換
    for key, value in dic.items():
        for paragraph in doc.paragraphs:
            paragraph.text = paragraph.text.replace(key, value)

    # Wordファイルを名前を変えて保存
    word_newFilePath = f'{i}_{values1[i][0]}_{values1[i][1]}.docx'
    doc.save(word_newFilePath)
    wordfile = os.path.join(curdir, word_newFilePath)


    # WordファイルをPDFに変換して保存
    wordDoc = word.Documents.Open(wordfile)
    pdf_newFileName = f'{i}_{values1[i][0]}_{values1[i][1]}.pdf'
    pdf_fileFullPath = os.path.join("C:/Users/obi21703/PycharmProjects/差し込み印刷/print_list", pdf_newFileName)
    wordDoc.SaveAs2(pdf_fileFullPath, FileFormat=17)

    # Wordファイルを閉じる
    wordDoc.Close()

word.DisplayAlerts = True
word.Quit()

#生成されたwordを削除
file_list = glob.glob("*_*docx")
for file in file_list:
    print("remove:{}".format(file))
    os.remove(file)


# #１つのPFDファイルにまとめる
l = glob.glob(os.path.join("C:/Users/obi21703/PycharmProjects/差し込み印刷/print_list","*.pdf"))
l.sort()
merger = PyPDF2.PdfFileMerger()
for p in l:
    if not PyPDF2.PdfFileReader(p).isEncrypted:
        merger.append(p)
merger.write("C:/Users/obi21703/PycharmProjects/差し込み印刷/print_list/印刷データ.pdf")
merger.close()

#生成元の各pdfファイルを削除する
file_list = glob.glob("C:/Users/obi21703/PycharmProjects/差し込み印刷/print_list/*_*pdf")
for file in file_list:
    print("remove:{}".format(file))
    os.remove(file)

#生成されたpdf帳票を確認する
pdf_check = subprocess.Popen(["start","",r"C:/Users/obi21703/PycharmProjects/差し込み印刷/print_list/印刷データ.pdf"],shell=True)
time.sleep(1)



#帳票出力処理を行うかどうか選択
res = messagebox.askquestion("question","帳票出力を行いますか？")
if res == "yes":
    res = messagebox.showinfo("print_start","印刷を開始します")
    print("印刷を開始します")
    from auto_print import auto_print_module, file_check
    auto_print_module("C:/Users/obi21703/PycharmProjects/差し込み印刷")
    file_check("C:/Users/obi21703/PycharmProjects/差し込み印刷/print_list")

else:
    res = messagebox.showinfo("exit","プログラムを終了します。")
    print("プログラムを終了します")
