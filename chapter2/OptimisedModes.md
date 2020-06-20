### 优化模式

#### 只读模式

Sometimes, you will need to open or write extremely large XLSX files, and the common routines in openpyxl won’t be able to handle that load. Fortunately, there are two modes that enable you to read and write unlimited amounts of data with (near) constant memory consumption.

Introducing openpyxl.worksheet._read_only.ReadOnlyWorksheet:

```
from openpyxl import load_workbook
wb = load_workbook(filename='large_file.xlsx', read_only=True)
ws = wb['big_data']

for row in ws.rows:
    for cell in row:
        print(cell.value)
```

