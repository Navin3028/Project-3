# import win32com.client as win32

# def create_bar_chart(wb, sheet, axis_range, value_range, chart_title, chart_left, chart_top):
#     try:
#         chart_sheet = wb.Worksheets.Add()  # Add a new worksheet for each chart
#         chart = chart_sheet.Shapes.AddChart2(251, 1) 
#         chart.Chart.SetSourceData(sheet.Range(value_range))
#         chart.Chart.HasTitle = True
#         chart.Chart.ChartTitle.Text = chart_title
#         chart.Chart.Axes(win32.constants.xlCategory, win32.constants.xlPrimary).CategoryNames = sheet.Range(axis_range).Value
#         chart.Chart.ChartType = win32.constants.xlColumnClustered
#         chart.Left = chart_left
#         chart.Top = chart_top
#     except Exception as e:
#         print(f"Error: {e}")

# def main():
#     excel = win32.Dispatch("Excel.Application")
#     wb = excel.Workbooks.Open(r"C:\Users\v.jeevinee\Documents\intern\Consolidate Cloud Expense January 2024_updated.xlsx")
#     sheet = wb.Worksheets("Monthly Summary")

#     axis_range = input("Enter the range for the axis (e.g., H1:AA1): ")
#     value_range1 = input("Enter the range for the values for chart 1: ")
#     value_range2 = input("Enter the range for the values for chart 2: ")
#     value_range3 = input("Enter the range for the values for chart 3: ")
#     value_range4 = input("Enter the range for the values for chart 4: ")

#     # Define the chart positions
#     chart1_left = sheet.Cells(1, 1).Left
#     chart1_top = sheet.Cells(1, 1).Top

#     chart2_left = sheet.Cells(11, 1).Left
#     chart2_top = sheet.Cells(11, 1).Top

#     chart3_left = sheet.Cells(31, 3).Left
#     chart3_top = sheet.Cells(31, 3).Top

#     chart4_left = sheet.Cells(11, 3).Left
#     chart4_top = sheet.Cells(11, 3).Top

    

#     create_bar_chart(wb, sheet, axis_range, value_range1, 'AWS Corporate Cost Per Day', chart1_left, chart1_top)
#     create_bar_chart(wb, sheet, axis_range, value_range2, 'PNC Cloud Cost Per Day', chart2_left, chart2_top)
#     create_bar_chart(wb, sheet, axis_range, value_range3, 'PartnerLinq Cost Per Day', chart3_left, chart3_top)
#     create_bar_chart(wb, sheet, axis_range, value_range4, 'Average Cloud Cost Per Day', chart4_left, chart4_top)

#     wb.Save()
#     wb.Close()
#     excel.Quit()

# if __name__ == "__main__":
#     main()

# import win32com.client as win32

# def create_bar_chart(wb, chart_sheet, axis_range, value_range, chart_title, cell):
#     try:
#         chart = chart_sheet.Shapes.AddChart2(251, 1) 
#         chart.Chart.SetSourceData(chart_sheet.Range(value_range))
#         chart.Chart.HasTitle = True
#         chart.Chart.ChartTitle.Text = chart_title
#         chart.Chart.ChartType = win32.constants.xlColumnClustered
#         chart.Left = cell.Left
#         chart.Top = cell.Top
        
#         # Get the category names directly from the specified range
#         category_range = chart_sheet.Range(axis_range)
#         chart.Chart.Axes(win32.constants.xlCategory, win32.constants.xlPrimary).CategoryNames = "=sub_monthly_sum!" + category_range.GetAddress()
#     except Exception as e:
#         print(f"Error: {e}")


# def main():
#     excel = win32.Dispatch("Excel.Application")
#     wb = excel.Workbooks.Open(r"C:\Users\v.jeevinee\Documents\intern\Consolidate Cloud Expense January 2024_updated.xlsx")
#     sheet = wb.Worksheets("Monthly Summary")
    
#     # Add a new worksheet for the charts
#     chart_sheet = wb.Worksheets.Add()
#     chart_sheet.Name = "sub_monthly_sum"
    
#     axis_range = input("Enter the range for the axis (e.g., H1:AA1): ")
#     value_range1 = input("Enter the range for the values for chart 1: ")
#     value_range2 = input("Enter the range for the values for chart 2: ")
#     value_range3 = input("Enter the range for the values for chart 3: ")
#     value_range4 = input("Enter the range for the values for chart 4: ")

#     # Define the cell positions for the charts
#     cell1 = chart_sheet.Cells(1, 1)  # A1
#     cell2 = chart_sheet.Cells(1, 6)  # F1
#     cell3 = chart_sheet.Cells(30, 1) # A30
#     cell4 = chart_sheet.Cells(30, 6) # F30

#     # Create the charts in the specified cell positions
#     create_bar_chart(wb, chart_sheet, axis_range, value_range1, 'AWS Corporate Cost Per Day', cell1)
#     create_bar_chart(wb, chart_sheet, axis_range, value_range2, 'PNC Cloud Cost Per Day', cell2)
#     create_bar_chart(wb, chart_sheet, axis_range, value_range3, 'PartnerLinq Cost Per Day', cell3)
#     create_bar_chart(wb, chart_sheet, axis_range, value_range4, 'Average Cloud Cost Per Day', cell4)

#     wb.Save()
#     wb.Close()
#     excel.Quit()

# if __name__ == "__main__":
#     main()



import win32com.client as win32

def create_bar_chart(wb, sheet, axis_range, value_range, chart_title, chart_left, chart_top):
    try:
        chart = sheet.Shapes.AddChart2(251, 1, chart_left, chart_top, 300, 200).Chart
        chart.SetSourceData(sheet.Range(value_range))
        chart.HasTitle = True
        chart.ChartTitle.Text = chart_title
        chart.Axes(win32.constants.xlCategory, win32.constants.xlPrimary).CategoryNames = sheet.Range(axis_range).Value
        chart.ChartType = win32.constants.xlColumnClustered
    except Exception as e:
        print(f"Error: {e}")

def main():
    excel = win32.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(r"C:\Users\v.jeevinee\Documents\intern\Consolidate Cloud Expense January 2024_updated.xlsx")
    sheet = wb.Worksheets("Monthly Summary")

    axis_range = input("Enter the range for the axis (e.g., H1:AA1): ")
    value_range1 = input("Enter the range for the values for chart 1: ")
    value_range2 = input("Enter the range for the values for chart 2: ")
    value_range3 = input("Enter the range for the values for chart 3: ")
    value_range4 = input("Enter the range for the values for chart 4: ")

    # Define the chart positions
    chart1_left = 50
    chart1_top = 50

    chart2_left = 400
    chart2_top = 50

    chart3_left = 50
    chart3_top = 300

    chart4_left = 400
    chart4_top = 300

    create_bar_chart(wb, sheet, axis_range, value_range1, 'AWS Corporate Cost Per Day', chart1_left, chart1_top)
    create_bar_chart(wb, sheet, axis_range, value_range2, 'PNC Cloud Cost Per Day', chart2_left, chart2_top)
    create_bar_chart(wb, sheet, axis_range, value_range3, 'PartnerLinq Cost Per Day', chart3_left, chart3_top)
    create_bar_chart(wb, sheet, axis_range, value_range4, 'Average Cloud Cost Per Day', chart4_left, chart4_top)

    wb.Save()
    wb.Close()
    excel.Quit()

if __name__ == "__main__":
    main()


