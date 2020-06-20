### [openpyxl]() - ***一个读/写 Excel 2010 xlsx/xlsm Python 库***


|||
|:-|:-|
|作者：|小白鞋|
|源代码：|http://bitbucket.org/openpyxl/openpyxl/src|
|ISSUES：|http://bitbucket.org/openpyxl/openpyxl/issues|
|翻译时间：|June 20, 2020|
|License：|MIT/Expat|
|Version：|0.1.0|



![](https://s3.amazonaws.com/assets.coveralls.io/badges/coveralls_95.svg)

#### 介绍

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

文档位于: https://openpyxl.readthedocs.io

- 安装方式
- 代码示例
- 贡献说明

发行说明: https://openpyxl.readthedocs.io/en/stable/changes.html


#### 支持

This is an open source project, maintained by volunteers in their spare time. This may well mean that particular features or functions that you would like are missing. But things don’t have to stay that way. You can contribute the project Development yourself or contract a developer for particular features.

Professional support for openpyxl is available from Clark Consulting & Research and Adimian. Donations to the project to support further development and maintenance are welcome.

Bug reports and feature requests should be submitted using the issue tracker. Please provide a full traceback of any error you see and if possible a sample file. If for reasons of confidentiality you are unable to make a file publicly available then contact of one the developers.


#### 如何贡献

Any help will be greatly appreciated, just follow those steps:

1. Please start a new fork (https://bitbucket.org/openpyxl/openpyxl/fork) for each independent feature, don’t try to fix all problems at the same time, it’s easier for those who will review and merge your changes ;-)

2. Hack hack hack

3. Don’t forget to add unit tests for your changes! (YES, even if it’s a one-liner, changes without tests will not be accepted.) There are plenty of examples in the source if you lack know-how or inspiration.

4. If you added a whole new feature, or just improved something, you can be proud of it, so add yourself to the AUTHORS file :-)

5. Let people know about the shiny thing you just implemented, update the docs!

6. When it’s done, just issue a pull request (click on the large “pull request” button on your repository) and wait for your code to be reviewed, and, if you followed all theses steps, merged into the main repository.

For further information see Development

#### 其它帮助方式

There are several ways to contribute, even if you can’t code (or can’t code well):

- triaging bugs on the bug tracker: closing bugs that have already been closed, are not relevant, cannot be reproduced, …

- updating documentation in virtually every area: many large features have been added (mainly about charts and images at the moment) but without any documentation, it’s pretty hard to do anything with it

- proposing compatibility fixes for different versions of Python: we support 2.7, 3.4, 3.5, 3.6 and 3.7

#### 安装


使用 pip 安装 openpyxl 。 建议在不带系统软件包的Python virtualenv中执行此操作：

`$ pip install openpyxl`


:o: 注意

支持流行的lxml库（如果已安装）将使用该库。 这在创建大文件时特别有用。

:heavy_exclamation_mark: 警告

To be able to include images (jpeg, png, bmp,…) into an openpyxl file, you will also need the “pillow” library that can be installed with:

`$ pip install pillow`

or browse https://pypi.python.org/pypi/Pillow/, pick the latest version and head to the bottom of the page for Windows binaries.

#### Working with a checkout

Sometimes you might want to work with the checkout of a particular version. This may be the case if bugs have been fixed but a release has not yet been made.

`$ pip install -e hg+https://bitbucket.org/openpyxl/openpyxl@3.0#egg=openpyxl`

#### 使用示例

##### 指导

- 指导
	- [创建一个工作簿]()
	- [玩儿数据]()
		- [访问一个单元格]()
		- [访问多个单元]()
		- [只获取值]()
	- [存储数据]()
		- [保存到文件]()
		- [另存为流]()
	- [从文件加载]()

##### 菜单

- 简易使用
	- [写一个工作簿]()
	- [读一个工作簿]()
	- Using number formats
	- Using formulae
	- Merge / Unmerge cells
	- Inserting an image
	- Fold (outline)

##### 性能

- [性能]()
	-[Benchmarks]()
		- Write Performance
		- Read Performance
		- Parallelisation


##### 发行说明



