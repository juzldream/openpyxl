### 指导

#### 创建一个工作簿

不需要在文件系统上创建文件即可开始使用openpyxl。 只需导入Workbook类并开始工作：

```
>>> from openpyxl import Workbook
>>> wb = Workbook()
```


至少需要一个工作薄然后创建工作表，你可以使用 `Workbook.active` 获取活跃工作表

`>>> ws = wb.active`

*默认情况下下标为0。 除非你修改其值，否则始终将使用此方法获得第一个工作表。*


你可以使用 `Workbook.create_sheet() ` 方法创建工作表：

```
ws1 = wb.create_sheet("Mysheet") # insert at the end (default)

ws2 = wb.create_sheet("Mysheet", 0) # insert at first position

ws3 = wb.create_sheet("Mysheet", -1) # insert at the penultimate position
```

创建工作表，不指定工作表名称，默认命名（Sheet, Sheet1, Sheet2, …），你可以通过 `Worksheet.title` 方法制定工作表名称

`ws.title = "New Title"`

默认创建的工作表的标题是白色背景，你可以通过 `Worksheet.sheet_properties.tabColor` 属性`RRGGBB`十六进制方式设定工作表标题背景颜色

`ws.sheet_properties.tabColor = '{0:06X}'.format(random.randint(0, 0xFFFFFF))`

你也可以使用 `Workbook.sheetname` 属性获取工作簿中所有工作表的名称。

```
print(wb.sheetnames)
['1月销售清单', '2月销售清单', '3月销售清单', '4月销售清单', '5月销售清单', '6月销售清单', '7月销售清单', '8月销售清单', '9月销售清单', '10月销售清单', '11月销售清单', '12月销售清单']
```

你也可以使用 `Workbook.copy_worksheet() ` 方法拷贝工作表

```
source = wb.active
target = wb.copy_worksheet(source)
```


#### 打开一个Excel文件

和写一个文件一样，你可以使用 `openpyxl.load_workbook()` 方法打开一个工作薄。

```
from openpyxl import load_workbook
wb = load_workbook('test.xlsx')
print wb.sheetnames
['Sheet2', 'New Title', 'Sheet1']
```

#### 保存一个文件

报错工作薄最简单、最安全的方法是 `Workbook.save()`

```
wb = Workbook()
wb.save('balances.xlsx')
```

#### 删除一个工作表

`wb.remove(wb.active)`

#### 样例代码

```
path = "D:\\local_python_learn\\data\\video-file\\002 工作簿与工作表操作\\sales.xlsx"
wb = load_workbook(path)
for i in range(12):
    source = wb.active
    target = wb.copy_worksheet(source)
    target.title = '{}月{}'.format(i+1, source.title)
    # target.sheet_properties.tabColor = '{0:6X}'.format(random.randint(0, 0xFFFFFF))
    print('{0:6X}'.format(random.randint(0, 0xFFFFFF)))
    target.sheet_properties.tabColor = '{0:06X}'.format(random.randint(0, 0xFFFFFF))
wb.remove(wb.active)
wb.save(path.split(".")[0] + "test.xlsx")
```

指导到此结束，你可以进入[简易使用](chapter1/SimpleUsage.md)章节。