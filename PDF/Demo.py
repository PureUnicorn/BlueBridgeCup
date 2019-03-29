import tabula

df = tabula.read_pdf("软件类-湖北赛区获奖名单.pdf", encoding='utf-8', pages='all')
print(df)
for indexs in df.index:
    # 遍历打印企业名称
    print(df.loc[indexs].values[1].strip())