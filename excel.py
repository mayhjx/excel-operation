import win32com.client
import os.path
from os import getcwd


def create_excel():
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible =1
        return excel
    except Exception as e:
        print(e)

    
def get_workbook(excel, filepath):

    if excel is None: return
    
    if not "\\" in filepath:
        filepath = os.path.join(getcwd(), filepath)
    
    if not ".xls" in os.path.basename(filepath):
        os.path.join(filepath, ".xlsx")

    try:
        wb = excel.Workbooks.Open(filepath)
        return wb
    except Exception as e:
        print("Creating Workbook: {0}".format(os.path.basename(filepath)))
        wb = excel.Workbooks.Add()
        wb.SaveAs(filepath)
        return wb


def get_sheet(wb, sheetname):
    
    if wb is None: return
    
    try:
        ws = wb.Worksheets(str(sheetname))
    except Exception as e:
        ws = wb.Worksheets.Add()
        ws.Name = str(sheetname)        
    return ws


def find(ws, what):

    if ws is None: return

    data = ws.UsedRange.Value

    #type convert
    for da in data:
        for d in da:
            if type(d)(what) == d:
                print("Find {0} in ({1}, {2})".format(what, data.index(da)+1, da.index(d)+1))
                

if __name__ == "__main__":

    path = r"test"
    new_sheet_name = "newsheet"
    
    excel = create_excel()
    wb = get_workbook(excel, path)
    ws = get_sheet(wb, new_sheet_name)
    
    if not ws is None:
        for i in range(1,11):
            for j in range(1,11):
                ws.Cells(i,j).Value = i * j
        
        
    find(ws, "6")
