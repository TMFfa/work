# openpyxl 再使用总结

需要用到的基本类实例：wb，ws，cell

## 1、起因

​	宿舍周分需要另一种格式，于是打算用Python直接从总分复制过去

## 2、cell的属性介绍

- cell.row，cell.column可以返回该单元格的行列

- cell.column_letter可以返回该单元格列的字母序号

- cell.number_format可以查看单元格的数字格式，也可以直接=设置新值

    ！我在复制日期时，它返回的值始终不是Excel显示的值，几番搜索才发现，这个日期用的是自定义的格式，因此，复制过去的值要还复制他的格式。但是，原表用代码显示是general格式，实际是自定义的。这只能自己手动复制自定义格式，并在复制后用cell.number_format = 'm"月"d"日"' （这是我的格式）来设置格式。

- cell.max_row, cell.max_column返回最大行列
- cell.dimensions返回表格数据范围，是个字符串，类似"A1:K9"
- cell.rows, cell.columns可以返回一行一行或一列一列的迭代器

## 3、一些新学到的知识

1、注意事项：新建的表格用 wb = openpyxl.Workbook()来创建

​			注意：括号里不要输入要保存的文件名，否则他就是writeonlyworkbook，而我们要普通的workbook才能进行多样操作，保存文件放在最后wb.save()那里即可。

2、合并单元格：其实很简单，就是用ws.merge_cells('A1:B4')，里面的值就是要合并单元格形成的矩形。合并后最好设置居中（反正我设置了）

3、居中：这需要from openpyxl.styles import Alignment

用cell的属性值设置：cell.alignment = Alignment(horizontal='center', vertical='center')

4、单元格边框设置：我只是设置实线，其实还有很多种边框都可以设置

```python
from openpyxl.styles import Border, Side, colors
# 实例化一个border类作为cell的属性
border = Border(
	left=Side(style='thin', color=colors.BLACK),
    right=Side(style='thin', color=colors.BLACK),
    top=Side(style='thin', color=colors.BLACK),
    bottom=Side(style='thin', color=colors.BLACK),
)
cell.border = border
"""
border很好理解，一个单元格矩形的四方，但是注意上下是top和bottom
border的每个参数又是一个Side对象，就是方格的一边
side对象的参数决定是设置什么边框：
	虚线：dashDot dashDotDot dashed hair(不同的点线)
	实线：thin thick medium(thin就是我们常用的实线，其他两个比较粗)
color可以设置颜色，这里的颜色用colors库的参数设置
"""
```



好了，大致就这么多了，每次进步一点点。
