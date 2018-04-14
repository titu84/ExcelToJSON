using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadDataFromExcel
{
    public class ExcelReader
    {
        string patch;
        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        Excel._Worksheet xlWorksheet;
        Excel.Range xlRange;
        public ExcelReader(string _patch)
        {
            patch = _patch;
        }
        public string GetJsonFormExcel(int sheetID = 1)
        {
            try
            {
                StringBuilder s = new StringBuilder();
                List<object[]> lo = new List<object[]>();
                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(patch);
                xlWorksheet = xlWorkbook.Sheets[sheetID];
                xlRange = xlWorksheet.UsedRange;
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                for (int i = 1; i <= rowCount; i++)
                {
                    object[] o = new object[colCount];
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            o[j - 1] = xlRange.Cells[i, j].Value2;
                    }
                    lo.Add(o);
                }
                return new JavaScriptSerializer().Serialize(lo);
            }
            catch (NullReferenceException nex)
            {
                return nex.Message;
            }
            catch (COMException ex)
            {
                return ex.Message;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                try
                {
                    Marshal.ReleaseComObject(xlRange);
                }
                catch { }
                try
                {
                    Marshal.ReleaseComObject(xlWorksheet);
                    xlWorkbook.Close(false);
                }
                catch { }
                try
                {
                    Marshal.ReleaseComObject(xlWorkbook);
                }
                catch { }
                try
                {
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                }
                catch { }
            }
        }
    }
}
