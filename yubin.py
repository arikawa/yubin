# 郵便局アンケート集計 2017/8/12
import xlrd
import os.path
import pandas as pd
xlfile = "data/data_20170808.xlsx"

file = pd.ExcelFile(xlfile)
# for sheet in file.sheet_names:
    # df_list.append(file.parse(sheet))
fm = file.parse("Format")
# フォーマットを穴埋め・フィールド名のダブリをリネーム
# fm.to_csv('output/Format.csv')
fm_af = fm.fillna(method='ffill')
fm_af.Question = fm_af.Question.where(fm_af.Column != 130, "Q37_2.1")
fm_af.Question = fm_af.Question.where(fm_af.Column != 134, "Q38_2.1")
fm_af.CtgNo = fm.CtgNo
fm_af.to_csv('output/Format_af.csv')

# 変換マスタを作成
type_master = fm_af.drop_duplicates(["Question", "Type"])[["Question", "Type", "Title"]]
type_master = type_master.set_index("Question")
type_master.to_csv('output/type_master.csv')
# type_master
S_master = fm_af[fm_af.Type == "S"][fm_af.CtgNo > 0]
S_master = S_master.set_index(S_master.Question + "_" + S_master.CtgNo.astype(str))
S_master.to_csv('output/S_master.csv')

# データの選択肢をフォーマットの値に変換
data = file.parse("data")
data.to_csv('output/data.csv')
data_af = data.copy()
for q in data_af.columns:
    if type_master.ix[q].Type == "S":
        try:
            data_af[q] = data[q].apply(lambda x: S_master.ix[q + "_" + str(float(x))].Title)
        except KeyError:
            True
    data_af = data_af.rename(columns = {q:q+":"+type_master.ix[q].Title})
data_af
data_af.to_csv('output/data_af.csv')
