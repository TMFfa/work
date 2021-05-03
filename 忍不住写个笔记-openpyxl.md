# openpyxl 模块总结

## 1、安装与导入

- pip install openpyxl  # 安装命令
- import openpyxl as op  # 个人喜好as op，方便写，虽然就创建工作簿要用

## 2、核心思想

- Excel由**工作簿**(`Workbook`)、**工作表**(`worksheet`)、**单元格**(`cell`)组成，这个模块就是操作这三个东西
- tips:  Excel操作的单元格都在**内存**里，且打开文件时有点慢，应该和电脑性能有关

## 3、基本操作

#### （1）工作簿

- wb = op.Workbook()  # wb就是一个工作簿对象（注意是**新建**的）
- wb = op.load_workbook('`file path`')  # 加载一个**已有**工作簿

#### （2）工作表

- 创建工作表：ws = wb.create_sheet('`sheet name`')

- 活跃工作表，一般默认是第一个，可直接使用 ws = wb.active 调用，且**新创建**的工作簿**默认会有一个sheet**

- 获取工作簿的所有工作表名称：wb.sheetnames  # 这是一个**列表**，包含所有工作表名称

- 调用指定工作表： ws = wb['`sheet name`']  # 相当于字典的键

    ​				或者用：wb.get_sheet_by_name('`sheet name`')

#### （3）单元格

- 单元格（cell）是一个**cell对象**

- 对cell**写入值**有两种方法：

    ​				1、ws['A5'] = 'value'  # 用键的方式直接写入

    ​				2、ws.cell(row=`行数`, column=`列数`, value=`值`)  # 值可以不输入，默认为None

- **访问**cell的方法：

    ​				1、ws.cell(row, column) 循环遍历

    ​				2、矩阵式访问：ws['A1':'E4]  # 返回的是**元组**，而且里面**每一行**又是一个**元组**

    ​				3、最基本的：ws['A4']  # 这是一个**cell对象**，以上访问方法返回的最小元素都是对象

- cell对象的属性（最小单位了，没有方法）：

    ​				`row`: 该单元格的行数

    ​				`column`：该单元格的列数

    ​				`value`：该单元格的值，空单元格值为None

    ​				`style`：单元格风格，默认为normal，就最基本的

    ​				`style_id`：单元格风格对应的id

    ​	用cell.row这样的代码就可以访问其属性了

    tips：cell还有很多属性，我只用了前三个，可以通过**dir(cell)**来查看（cell得是个具体的**实例**）

#### （4）保存

- wb.save('`file path`')  # 注意如果是新建的表格，而file path已经存在，它会直接覆盖
- 如果是用load_workbook()载入的工作簿，修改后**`记得保存!!!`**，保存时的file path可以任意，里面的原始内容还在，如果写原文件的路径，就会更新原文件
- 类比open函数， 最后记得wb.close()