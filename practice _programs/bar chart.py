# import Workbook from openpyxl
from openpyxl import Workbook

# import RadarChart, Reference from openpyxl.chart sub_module .
from openpyxl.chart import BarChart, Reference

# Call a Workbook() function of openpyxl
# to create a new blank Workbook object
workbook = Workbook()
workbook1=workbook.active

values = Reference(workbook1, min_col=1, min_row=1, max_col=2, max_row=3)

chart = BarChart()
chart.add_data(values)
chart.title = " BAR-CHART "
chart.x_axis.title = " X_AXIS "
chart.y_axis.title = " Y_AXIS "
chart.set_categories(values)
workbook1.add_chart(chart, "E2")
workbook.save("master.xlsx")