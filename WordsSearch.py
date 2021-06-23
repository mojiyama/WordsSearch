#! python3
# -WordsSearch.py 指定したフォルダ階層の中のExcelとPDF内のテキストから
#   検索する文字列に該当するものがあるか調べ、ファイル名とパスを返す。
#   なお、フォルダ内のショートカットまではアクセスしない。

import os,re,openpyxl,sys

#検索するファイルの拡張子
excel_regex = re.compile(r'xlsx$')
pdf_regex = re.compile(r'pdf$')
zip_regex = re.compile(r'zip$')
#TODO 検索する文字列（最終的にはsys.argv[2]に文字列を入れるようにする）
text = "流木発生場所"
text_regex  = re.compile(r".*"+text+".*")

#テストとして光明谷右支川の報告書作成データを検索フォルダに選択
#TODO 最終的にはsys.argv[1]にパス名を入れるようにする。*sys.argv[0]は.batファイル名*
pri_path = r'\\192.168.11.14\share\作業フォルダ\0020-0722_光明谷砂防堰堤詳細(2基)'
#pri_path = r'C:\Users\honjo\Desktop\sample'

#開始の表示
print('****検索開始****')

#選択したフォルダ内のフォルダ、サブフォルダ、ファイル名を取得する
for folder,sub_folders,files in os.walk(pri_path):
    os.chdir(pri_path)
    if pri_path != folder:
        sub_path = os.path.join(pri_path,folder)
        os.chdir(sub_path)
    #print(folder) #親フォルダの確認用
    #TODO もしファイルが.xlsか.xlsxならファイルを開く
    for file in files:
        flag = False #forループbreak用のフラグ
        if re.search(excel_regex,file):
            try:
                wb = openpyxl.load_workbook(file)
                #TODO ファイルからシートを取得し、シート内の各行を検索する
            except PermissionError:
                continue
            for sheet in wb:
                if flag:
                    break
                for row in sheet.rows:
                    if flag:
                        break
                    for cell in row:
                        if re.search(text_regex,str(cell.value)):
                            print(f'ファイル：「{file}」内に「{text}」が見つかりました。')
                            print(f'ファイルの場所は、\n{folder}\nです。\n')
                            flag = True
                            break
print("*****検索終了*****")
