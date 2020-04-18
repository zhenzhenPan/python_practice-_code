import xlrd

data = xlrd.open_workbook("SPMS_api.xlsx") # 打开Excel文件读取数据
table = data.sheet_by_name("Sheet1") # 读取sheet1表单
nrows = table.nrows  # 获取该sheet中的有效行数

colNum =table.ncols # 获取总列数

#table.row_values(rowx, start_colx=0, end_colx=None)   返回由该行中所有单元格的数据组成的列表

keys = table.row_values(0) # 获取第一行作为key值
print (keys)
print("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
print("\n")
if nrows <= 1:
    print("总行数小于1")
else:
    r = []
    j = 1
    for i in list(range(nrows-1)): # range() 取随机数，list() 将元组转换为列表
        s = {}
        s['nrows'] = i+2

        values = table.row_values(j)
        for x in list(range(colNum)):
            s[keys[x]] = values[x] # 把每个单元格的内容 存到 字典 s 里
        r.append(s)
        print(r)
        print("******************************************")
        j += 1
