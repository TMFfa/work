from datetime import datetime
import openpyxl as op

time1 = datetime.now()
wb1 = op.load_workbook('#D座高二周分.xlsx')  # 周分表******
ws1 = wb1['第十一周']  # 第几周##############################**************************
wb2 = op.load_workbook('#D座高二男生个人分.xlsx')  # 个人扣分总表*********
ws2 = wb2['个人分']
time2 = datetime.now()
print(time2-time1)


# 周分表的函数
def collect_class(cell_list):
    li = []
    for cell in cell_list:
        if cell.value is not None:
            li.append(cell)
    return li[2:]  # 这里的意思是只返回班级，上面的表头不要（具体情况具体分析）*****


# 个人扣分总表的函数
def get_data(cell_list):
    li = []
    for cell in cell_list:
        if cell.value is not None:
            li.append(cell)
    li = li[1:]  # ***************************同上，只要除表头外的内容, 如果是第一周第一天，用3， 其他每周第一天用2，其余用1
    data = [[], [], []]  # 班级，宿舍，扣分
    for cell in li:
        row = str(cell.row)
        data[0].append(str(ws2['D'+row].value)+'班')  # 根据班级列返回其值********班级那一列的值
        data[1].append(ws2['B'+row].value)  # 宿舍列的值*********
        data[2].append(cell.value)  # 扣分内容
    return data


# 载入扣分内容，也就是个人扣分表
ws2_cells = ws2['BH']  # 扣分那一列****************************################################
ws2_data = get_data(ws2_cells)
print(ws2_data)

# 周分表班级列表（注意里面的值是‘x班’）
class_cells = ws1['A']  # 周分表班级列*******
class_cells = collect_class(class_cells)


# 开始查找并录入数据
def parse(class_cell_list, data):
    while len(data[0]):
        for class_cell in class_cell_list:
            class_ = class_cell.value
            print(class_)
            if class_ in data[0]:
                index = data[0].index(class_cell.value)
                dormitory = data[1][index]
                for i in range(10):
                    row = 'B' + str(class_cell.row + i)
                    dor = ws1[row].value
                    try:
                        if '（' in dor:
                            dor = int(dor.split('（')[0])
                        elif '(' in dor:
                            dor = int(dor.split('(')[0])
                        else:
                            pass
                    except:
                        pass
                    print(dor, dormitory)

                    if dor == dormitory:
                        cell = ws1['G'+str(class_cell.row + i)]  # 周分要扣的那一列，每次都要改***********##############
                        if cell.value is not None:
                            cell.value += data[2][index]
                        else:
                            cell.value = data[2][index]
                        data[0].pop(index)
                        data[1].pop(index)
                        data[2].pop(index)
                        break  # 打断这个for循环
                    else:
                        pass
            else:
                pass
        print(ws2_data)


parse(class_cells, ws2_data)

print(ws2_data)

wb1.save('#D座高二周分.xlsx')  # 保存路径，可以随便写，确认之后再复制到正表
wb1.close()
wb2.close()
