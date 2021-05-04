# 在source.xlsx表中添加扣分
import openpyxl as op


# 用于解析单条数据为多人扣分的情况，传入的参数为单条数据拆分后的列表
def parse_data(dt):
    people_num = len(dt)-2
    for i in range(people_num):
        dormitory.append(int(dt[0]))
        nums.append(int(dt[i+1]))
        marks.append(float(dt[-1])/people_num)


# 用于找到宿舍时分析床号，确认后添加扣分
def parse_cell(cell):
    index = dormitory.index(cell.value)
    row = str(cell.row)
    num = ws2['D'+row].value
    if num == nums[index]:
        ws2['F'+row] = marks[index]
        dormitory.pop(index)
        nums.pop(index)
        marks.pop(index)
    else:
        pass


# 经测试，打开这两个文件的速度有点慢
wb1 = op.load_workbook('today.xlsx')
wb2 = op.load_workbook('source.xlsx')
ws1 = wb1['today']
ws2 = wb2['个人扣分']

# 将扣分数据按顺序排入列表，输入时也要按顺序
datas = []
for cell in ws1['A']:
    datas.append(cell.value)
print(datas)

dormitory = []
nums = []
marks = []
for data in datas:
    dt = data.split('/')
    parse_data(dt)
print(dormitory, nums, marks)

# 在总表中搜索宿舍和床号，确认后写入，用parse_cell()确认
dor = ws2['C']
for cell in dor:
    if cell.value in dormitory:
        parse_cell(cell)
    else:
        pass

# 关闭文件， 艹，记得保存
wb2.save('source.xlsx')
wb1.close()
wb2.close()
