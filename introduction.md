#### [openpyxl]() - 一个读/写 Excel 2010 xlsx/xlsm Python 库

作者 ： 小白鞋

源代码 ： http://bitbucket.org/openpyxl/openpyxl/src

Issues ： http://bitbucket.org/openpyxl/openpyxl/issues

翻译时间： June 20, 2020

License ： MIT/Expat

Version : 0.0.1

![](images/coveralls_95.svg)

##### 介绍

openpyxl 是一个 Python 库用于读/写 Excel 2010 xlsx / xlsm / xltx / xltm文件。

It was born from lack of existing library to read/write natively from Python the Office Open XML format.

All kudos to the PHPExcel team as openpyxl was initially based on PHPExcel.

#### 安全

By default openpyxl does not guard against quadratic blowup or billion laughs xml attacks. To guard against these attacks install defusedxml.

#### 邮件列表

The user list can be found on http://groups.google.com/group/openpyxl-users

代码样例:

```
from openpyxl import Workbook
wb = Workbook()

# grab the active worksheet
ws = wb.active

# Data can be assigned directly to cells
ws['A1'] = 42

# Rows can also be appended
ws.append([1, 2, 3])

# Python types will automatically be converted
import datetime
ws['A2'] = datetime.datetime.now()

# Save the file
wb.save("sample.xlsx")

```

#### 文档

The documentation is at: https://openpyxl.readthedocs.io

- installation methods
- code examples
- instructions for contributing

Release notes: https://openpyxl.readthedocs.io/en/stable/changes.html


#### 支持


#### 如何贡献


#### 其它帮助方式


#### 安装

Install openpyxl using pip. It is advisable to do this in a Python virtualenv without system packages:

`$ pip install openpyxl`


#### 