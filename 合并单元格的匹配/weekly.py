import openpyxl as op

wb1 = op.load_workbook('demo1.xlsx')  # 周分表
ws1 = wb1.active
wb2 = op.load_workbook('source.xlsx')  # 个人扣分总表
ws2 = wb2.active


def collect_class(cell_list):
    li = []
    for cell in cell_list:
        if cell.value is not None:
            li.append(cell)
    return li[1:]  # 这里的意思是只返回班级，上面的表头不要（具体情况具体分析）


def get_data(cell_list):
    li = []
    for cell in cell_list:
        if cell.value is not None:
            li.append(cell)
    li = li[1:]  # 同上，只要除表头外的内容
    data = [[], [], []]  # 班级，宿舍，扣分
    for cell in li:
        row = str(cell.row)
        data[0].append(str(ws2['B'+row].value)+'班')  # 根据班级列返回其值
        data[1].append(ws2['C'+row].value)  # 宿舍列
        data[2].append(cell.value)  # 扣分内容
    return data


# 载入扣分内容
ws2_cells = ws2['F']  # 扣分那一列
ws2_data = get_data(ws2_cells)

# 周分表班级列表（注意里面的值是‘x班’）
class_cells = ws1['A']
class_cells = collect_class(class_cells)


# 开始查找并录入数据
def parse(class_cell_list, data):
    for cell in class_cell_list:
        # p = 1
        while cell.value in data[0]:  # 因为这里用的while循环，如果有个值没有确定的话会卡住，所以前面28行最好确定排除多少个值
            index = data[0].index(cell.value)
            try:
                if '（' in data[1][index]:  # 注意这个括号是中文的，表单要一致
                    dormitory = data[1][index].strip('（')[0]
                elif '(' in data[1][index]:
                    dormitory = data[1][index].strip('(')[0]  # 这里是英文括号，现在兼容了，暂时无bug了
                else:
                    dormitory = data[1][index]
            except:
                dormitory = data[1][index]
            i = 0
            while i < 10:
                if ws1['B'+str(cell.row+i)].value == dormitory:  # 周分宿舍列********
                    row = str(cell.row+i)
                    # 定位到目标后，检测是否已有值，有就要相加
                    if ws1['C'+row].value is not None:  # 周分要扣的那一列，每次都要改*********#############
                        ws1['C' + row].value += data[2][index]
                    else:
                        ws1['C' + row].value = data[2][index]
                    data[0].pop(index)
                    data[1].pop(index)
                    data[2].pop(index)
                    break
                else:
                    i += 1
            # p += 1
            # if p > 10000:  # 防卡死装置
            #     break


parse(class_cells, ws2_data)

wb1.save('demo1.xlsx')
wb1.close()
wb2.close()
