import win32com.client
from win32com.client import Dispatch
__version__ = '0.0.1'

class ExcelLib():

    def __init__(self):
        ROBOT_LIBRARY_SCOPE = 'GLOBAL'
        ROBOT_LIBRARY_VERSION = __version__
        self.wb = None
        self.xl = None

    def Open_Excel(self, ExcelPath):
        self.xl = win32com.client.Dispatch('Excel.Application')
        self.wb = self.xl.Workbooks.Open(ExcelPath)

    def GetSheetCount(self):
        return self.wb.Worksheets.Count

    def GetSheetName(self):
        return self.wb.Activesheet.Name

    def GetRowCount(self,SheetName):
        ws = self.wb.Worksheets(SheetName)
        allData = ws.UsedRange
        return allData.Rows.Count

    def GetColumnCount(self,SheetName):
        ws = self.wb.Worksheets(SheetName)
        allData = ws.UsedRange
        return allData.Columns.Count

    def GetCellData(self,SheetName,iRow,iCol):
        ws = self.wb.Worksheets(SheetName)
        allData = ws.UsedRange
        return allData.Cells(int(iRow),int(iCol))

    def GetCellData_By_Name(self,SheetName,CellName):
        ws = self.wb.Worksheets(SheetName)
        allData = ws.UsedRange
        return allData.Range(CellName)

    def Write_Cell_Data(self,SheetName,iRow,iCol,TestData):
        ws = self.wb.Worksheets(SheetName)
        ws.Cells(int(iRow),int(iCol)).Value = TestData

    def Write_Cell_Data_By_Name(self,SheetName,CellName,TestData):
        ws = self.wb.Worksheets(SheetName)
        ws.Range(CellName).Value = TestData

    def Clear_Cell_Data(self,SheetName,iRow,iCol):
        ws = self.wb.Worksheets(SheetName)
        ws.Cells(int(iRow),int(iCol)).Value = ''

    def Clear_Cell_Data_By_Name(self,SheetName,CellName):
        ws = self.wb.Worksheets(SheetName)
        ws.Range(CellName).Value = ''

    def Save_Excel(self):
        self.wb.Save()

    def SaveAs_Excel(self,ExcelPath):
        self.wb.Save(ExcelPath)

    def Close_Excel(self):
        self.wb.Close()
        self.xl.Quit()
        self.xl = None
