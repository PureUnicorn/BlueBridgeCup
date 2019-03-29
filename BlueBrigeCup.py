import xlwt
import pandas as pd

# 湖北省高校排名
table = {"武汉大学":1,
         "华中科技大学":2,
         "中南财经政法大学": 3,
         "武汉理工大学": 4,
         "华中师范大学": 5,
         "中国地质大学": 6,
         "华中农业大学": 7,
         "武汉科技大学": 8,
         "湖北大学": 9,
         "中南民族大学": 10,
         "湖北工业大学": 11,
         "武汉工程大学": 12,
         "三峡大学": 13,
         "湖北医药学院": 14,
         "湖北经济学院": 15,
         "湖北中医药大学": 16,
         "武汉纺织大学": 17,
         "湖北第二师范学院": 18,
         "武汉轻工大学": 19,
         "江汉大学": 20,
         "武汉学院": 21,
         "长江大学": 22,
         "湖北汽车工业学院": 23
         }

#设置表格样式
def set_style(name,height,bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    return style

def readFile():
    with open("湖北.txt","r") as f :
        e = xlwt.Workbook()
        sheet1 = e.add_sheet('湖北', cell_overwrite_ok=True)
        row0 = ["学校", "比赛类型", "奖项","是否进入决赛","排名"]

        for i in range(0, len(row0)):
            sheet1.write(0, i, row0[i], set_style('Times New Roman', 220, True))

        # for x in range(0,11):
        x= 0
        while True:
            line = f.readline()
            if not len(line):
                break
            # print(line)
            school = line.split("10")[0].split("(")[0]
            number = line.split("组")[0][-2:-1]
            up = ""

            temp = line.split("组")
            # print(temp[1])
            # print(len(temp[1]))
            if(len(temp[1])>4):
                prize = temp[1][0:3]
                # print(prize)
                up = "是"
            else:
                prize = temp[1]

            sheet1.write(x+1,0,school)
            sheet1.write(x+1,1,number)
            sheet1.write(x+1,2,prize.strip())
            sheet1.write(x+1,3,up)
            sheet1.write(x+1,4,"30")
            x+=1

            # print(school)
    e.save('test.xls')

def excelrank():
    writer = pd.ExcelWriter('data.xls')
    df = pd.read_excel("test.xls",names=["学校", "比赛类型", "奖项","排名"])
    rank = pd.Series()

    for school in df["学校"] :
        exception = 0
        try:
            if (table.get(school[2:])!=None):
                p = pd.Series(table.get(school[2:]))
            else:
                p = pd.Series(30)
            rank = rank.append(p,ignore_index=True)
        except Exception as e:
            exception+=1
    df["排名"] = rank
    df.to_excel(writer,"ultimate")

    writer.save()
    writer.close()

def analysis():
    df = pd.read_excel("data.xls")
    list = []

    # 获奖人数前十学校
    # print(df["学校"].value_counts()[:10])

    # 学校排名前十获奖人数
    # print(df[df["排名"] <11]["排名"].value_counts())

    # 奖项分布
    # print(df["奖项"].value_counts())

    # 三等奖
    # print(df[df["奖项"]=='三等奖']['学校'].value_counts().head(10))
    # 二等奖
    # print(df[df["奖项"] == '二等奖']['学校'].value_counts().head(10))
    # 一等奖
    # print(df[df["奖项"] == '一等奖']['学校'].value_counts().head(10))

    # A组二等奖
    # print(df[(df["奖项"]=='二等奖')&(df['比赛类型']=="A")]['学校'].value_counts().head(10))
    # A组一等奖
    # print(df[(df["奖项"]=='一等奖')&(df['比赛类型']=="A")]['学校'].value_counts().head(10))


    # B组二等奖
    # print(df[(df["奖项"]=='二等奖')&(df['比赛类型']=="B")]['学校'].value_counts().head(10))
    # B组一等奖
    # print(df[(df["奖项"]=='一等奖')&(df['比赛类型']=="B")]['学校'].value_counts().head(10))


    # 一等奖B组
    # print(df[(df["奖项"]=='一等奖')&(df['比赛类型']=="B")]['学校'].value_counts().head(10))
    # 一等奖C组
    # print(df[(df["奖项"]=='一等奖')&(df['比赛类型']=="C")]['学校'].value_counts().head(10))

    # A组获奖学校数量排名
    # print(df[df["比赛类型"]=="A"]["学校"].value_counts().head(10))
    # B组获奖学校数量排名
    # print(df[df["比赛类型"]=="B"]["学校"].value_counts().head(10))
    # C组获奖学校数量排名
    # print(df[df["比赛类型"]=="C"]["学校"].value_counts().head(10))

    a = df[df["比赛类型"]=="A"]["学校"].value_counts().head(10)
    for n, v in a.to_dict().items():
        # list.append(json.dumps({'name': n, 'value': v},ensure_ascii=False))
        list.append("{value:" + str(v) + ", name:\"" + n + "\"}")
        # list.append(" "+str(v))
        # list.append("\"" + n[2:] + "\"")

    for l in list:
        print(l, end="")
        print(",")
        # print(",",end="")

if __name__ == "__main__":
    # 1，将网上的PDF文件提取出来保存为excel
    readFile()

    # 2，将提取出来的数据标记学校排名
    excelrank()

    # 3，对数据进行清洗、分析
    analysis()

