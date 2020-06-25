### 图表 

#### 图表类型

The following charts are available:

- Area Charts
	- 2D Area Charts
	- 3D Area Charts
- Bar and Column Charts
	- Vertical, Horizontal and Stacked Bar Charts
	- 3D Bar Charts
	- Bubble Charts

- Bubble Charts
- Line Charts
	- Line Charts
	- 3D Line Charts

#### 创建一个图表

Charts are composed of at least one series of one or more data points. Series themselves are comprised of references to cell ranges.

```
>>> from openpyxl import Workbook
>>> wb = Workbook()
>>> ws = wb.active
>>> for i in range(10):
...     ws.append([i])
>>>
>>> from openpyxl.chart import BarChart, Reference, Series
>>> values = Reference(ws, min_col=1, min_row=1, max_col=1, max_row=10)
>>> chart = BarChart()
>>> chart.add_data(values)
>>> ws.add_chart(chart, "E15")
>>> wb.save("SampleChart.xlsx")
```

By default the top-left corner of a chart is anchored to cell E15 and the size is 15 x 7.5 cm (approximately 5 columns by 14 rows). This can be changed by setting the anchor, width and height properties of the chart. The actual size will depend on operating system and device. Other anchors are possible; see openpyxl.drawing.spreadsheet_drawing for further information.

#### Working with axes

- Axis Limits and Scale
	- Minima and Maxima
	- Logarithmic Scaling
	- Axis Orientation
- Adding a second axis