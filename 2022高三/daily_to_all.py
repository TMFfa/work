# 添加扣分到个人扣分总表，注意扣分时不要有表头
import openpyxl as op
import time

kou_fen = 'O'
today_file = '9.18.xlsx'

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
    num = ws2['C'+row].value  # 床号那一列********
    if num == nums[index]:
        if ws2[kou_fen+row].value is not None:  # 每次必改，要扣分的那一列***********************########################
            ws2[kou_fen + row].value += marks[index]  # 每次必改，要扣分的那一列***********************##############
        else:
            ws2[kou_fen+row].value = marks[index]  # 每次必改，要扣分的那一列***********************######################
        dormitory.pop(index)
        nums.pop(index)
        marks.pop(index)
    else:
        pass


# 经测试，打开这两个文件的速度有点慢
wb1 = op.load_workbook(today_file)
wb2 = op.load_workbook('#高三男生个人扣分.xlsx')
ws1 = wb1.active
ws2 = wb2['Sheet1']

# 将扣分数据按顺序排入列表，输入时也要按顺序
datas = []
for cell in ws1['A']:
    datas.append(cell.value)
print(datas)

dormitory = []
nums = []
marks = []
for data in datas:
    try:  # 防止输入有空白
        dt = data.split('/')
        parse_data(dt)
    except AttributeError:
        pass
print(dormitory, nums, marks)

# 在总表中搜索宿舍和床号，确认后写入，用parse_cell()确认
dor = ws2['B']  # 宿舍那一列**********************

for i in range(len(dormitory)):  # 这要求录分时不能空行
    if not dormitory:
        break
    for cell in dor:
        if cell.value in dormitory:
            parse_cell(cell)
        else:
            pass


print(dormitory, nums, marks)
# 关闭文件， 记得保存
wb2.save('#高三男生个人扣分.xlsx')
wb1.close()
wb2.close()
