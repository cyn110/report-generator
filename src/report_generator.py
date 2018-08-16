import mysql.connector
import win32com.client
from config import *
from xlConstants import *
from datetime import datetime


class ReportGenerator:
    QUERY = "SELECT Name, District, Population FROM city WHERE CountryCode = 'USA';"

    def __init__(self, output_directory: str):
        self._path = r'{}\report_{}.xlsx'.format(output_directory, datetime.now().strftime('%Y%m%d%H%M%S'))
        self._excel = None
        self._workbook = None
        self._ws_data = None
        self._ws_pivot = None

    def generate(self):
        try:
            self._excel = win32com.client.DispatchEx('Excel.Application')
        except Exception:
            raise ReportGeneratorException('Error in dispatch Excel.')
        else:
            try:
                self._prepare_workbook()
                self._load_data()
                self._draw_pivot()
                self._workbook.SaveAs(self._path)
            except Exception as err:
                raise err
            finally:
                self._excel.Application.DisplayAlerts = False
                self._excel.Application.Quit()

    def _prepare_workbook(self):
        self._workbook = self._excel.Workbooks.Add()
        self._ws_data = self._workbook.Worksheets(1)
        self._ws_data.Name = 'Data'
        self._ws_pivot = self._workbook.Worksheets.Add()
        self._ws_pivot.Name = 'Pivot'

    def _load_data(self):
        con = None
        cur = None
        try:
            con = mysql.connector.connect(user=DB_USER, password=DB_PASSWORD, host=DB_HOST, database=DB_SCHEMA)
            cur = con.cursor()
            cur.execute(self.QUERY)
        except Exception as err:
            raise ReportGeneratorException('Error in communicating with the database. Error: ' + str(err))
        else:
            self._ws_data.Range('A1:C1').Value = ['Name', 'District', 'Population']
            for i, (Name, District, Population) in enumerate(cur):
                self._ws_data.Range('A' + str(i + 2) + ':C' + str(i + 2)).Value = [Name, District, Population]

            self._ws_data.ListObjects.Add(SourceType=xlSrcRange,
                                          Source=self._ws_data.Range('A1:C' + str(cur.rowcount + 1)),
                                          XlListObjectHasHeaders=xlYes).Name = 'us_city'
        finally:
            if cur is not None:
                cur.close()
            if con is not None:
                con.close()

    def _draw_pivot(self):
        try:
            pc = self._workbook.PivotCaches().Create(SourceType=xlDatabase,
                                                     SourceData=self._ws_data.Range('us_city[#All]'))
            pt = pc.CreatePivotTable(TableDestination=self._ws_pivot.Range('A1'), TableName='us_city_pivot')

            pt.PivotFields('District').Orientation = xlRowField
            pt.PivotFields('District').Position = 1
            pt.PivotFields('Name').Orientation = xlRowField
            pt.PivotFields('Name').Position = 2

            pt.AddDataField(Field=pt.PivotFields('Population'),
                            Caption='Sum of Population',
                            Function=xlSum)
            pt.AddDataField(Field=pt.PivotFields('Population'),
                            Caption='Percentage of Population')
            pt.PivotFields('Percentage of Population').Calculation = xlPercentOfTotal
            pt.PivotFields('Percentage of Population').NumberFormat = '0.00%'

        except Exception as err:
            raise ReportGeneratorException('Error in drawing PivotTable. Error: ' + str(err))


class ReportGeneratorException(Exception):
    pass
