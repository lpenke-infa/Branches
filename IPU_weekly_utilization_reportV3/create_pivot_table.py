import win32com.client as win32
import time

xlApp = win32.Dispatch("Excel.Application")
xlApp.Visible = True
# xlApp.Visible = False

def getxlApp():
    return xlApp
    
def createPivotTableStyle(path='',tableName="MyReportSummary",dataSheet="",summarySheet="", cellNumber="B3"):

    workbook = xlApp.Workbooks.Open(path)

    ws_data = workbook.Worksheets(dataSheet)
    ws_report = workbook.Worksheets(summarySheet)

    #create pivot table cache connection 
    pt_cache = workbook.PivotCaches().Create(1,ws_data.Range('A1').CurrentRegion)

    #create pivot table designer/editor
    pt = pt_cache.CreatePivotTable(ws_report.Range(cellNumber), tableName)

    #row and column grandtotals
    pt.ColumnGrand = True
    pt.RowGrand = True

    #change report layout
    pt.RowAxisLayout(0)   # RowAxisLayout() 0 1 2

    #change pivot table style
    # pt.TableStyle2 = "pivotStyleMedium9"
    pt.TableStyle2 = "pivotStyleLight16"

    return pt, ws_report

def createPivotTable(pt='',selectedSubOrgs=[],months=[]):
    field_rows = {} 
    field_rows["Month"] = pt.PivotFields("Month")
   
    # field_rows["Month"].AutoSort(1, 2)
    # print(dir(field_rows["Month"].PivotItems('Nov 2023')))
    # field_rows["Month"].AutoSort(2, "Descending", 2, None)
    # exit(0)
    # if(len(months) > 0):
    #     # field_rows["Month"].AutoSort = False
    #     for i, item_name in enumerate(months, start=1):
    #         field_rows["Month"].PivotItems(item_name).Position = i
        # month_field.PivotItems(item_name).Position = i
    
    # if(len(months) > 0):
    #     field_rows["Month"].ClearAllFilters()
    #     for item in field_rows["Month"].PivotItems():
    #         item.Visible = item.Name in months
            
    field_rows["Meter"] = pt.PivotFields("Meter")

    field_rows["Month"].Orientation = 1    # 1 for row orientation
    field_rows["Month"].Position = 1

    field_rows["Meter"].Orientation = 1    # 1 for row orientation
    field_rows["Meter"].Position = 2
    if(len(selectedSubOrgs) > 0):
        field_rows["Meter"].ClearAllFilters()
        for item in field_rows["Meter"].PivotItems():
            item.Visible = item.Name in selectedSubOrgs

    field_values = {}
    field_values["IPUs Consumed"] = pt.PivotFields("IPUs Consumed")
    field_values["IPUs Consumed"].Orientation = 4   # 4 for data/value
    field_values["IPUs Consumed"].Function = -4157  # -4112 for xlCount 4157
    
    time.sleep(1)

def createPivotTableSubOrg(pt='', numericalRepresentation=0, month='', selectedSubOrgs=[]):
    field_rows = {} 
    field_rows["Meter"] = pt.PivotFields("Meter")
    field_rows["Meter"].Orientation = 1    # 1 for row orientation
    field_rows["Meter"].Position = 1
    if(len(selectedSubOrgs) > 0):
        field_rows["Meter"].ClearAllFilters()
        for item in field_rows["Meter"].PivotItems():
            item.Visible = item.Name in selectedSubOrgs
    
    field_filter = {}
    field_filter["Month"] = pt.PivotFields("Month")
    field_filter["Month"].Orientation = 3  # 3 for PageFields
    field_filter["Month"].ClearAllFilters
    field_filter["Month"].CurrentPage = month

    field_values = {}
    field_values["IPUs Consumed"] = pt.PivotFields("IPUs Consumed")
    field_values["IPUs Consumed"].Orientation = 4   # 4 for data/value
    field_values["IPUs Consumed"].Function = numericalRepresentation  # -4112 for xlCount 4157
    
    time.sleep(1)


def getPivotTableData(workbook_path, sheet_name, pivot_table_name):
    workbook = xlApp.Workbooks.Open(workbook_path)
    sheet = workbook.Sheets(sheet_name)
    pivot_table = sheet.PivotTables(pivot_table_name)
    data_range = pivot_table.TableRange1
    data = [row.Value for row in data_range.Rows]
    return data
        
        
