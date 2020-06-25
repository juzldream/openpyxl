### 命令

:exclamation: 警告

Openpyxl currently supports the reading and writing of comment text only. Formatting information is lost. Comment dimensions are lost upon reading, but can be written. Comments are not currently supported if read_only=True is used.


#### Adding a comment to a cell

Comments have a text attribute and an author attribute, which must both be set

```
>>> from openpyxl import Workbook
>>> from openpyxl.comments import Comment
>>> wb = Workbook()
>>> ws = wb.active
>>> comment = ws["A1"].comment
>>> comment = Comment('This is the comment text', 'Comment Author')
>>> comment.text
'This is the comment text'
>>> comment.author
'Comment Author'

```