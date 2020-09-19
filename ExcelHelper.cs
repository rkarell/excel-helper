using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace ExcelHelp
{
    public class ExcelHelper
    {
        public Excel.Workbook OpenWb(string path)
        {
            Excel.Application excelApp = GetExcelAppIfOpen();
            if (excelApp == null)
            {
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                return excelApp.Workbooks.Open(path);
            }
            else
            {
                if (IsWbOpen(excelApp, path))
                    return excelApp.Workbooks[Path.GetFileName(path)];
                
                return excelApp.Workbooks.Open(path);
            }
        }

        public Excel.Workbook CreateWb()
        {
            Excel.Application excelApp = GetExcelAppIfOpen();
            if (excelApp == null)
            {
                excelApp = new Excel.Application();
                excelApp.Visible = true;
            }
            return excelApp.Workbooks.Add();
        }

        public void DontUpdate()
        {
            Excel.Application excelApp = GetExcelAppIfOpen();
            if (excelApp != null)
                excelApp.ScreenUpdating = false;
        }

        public void Update()
        {
            Excel.Application excelApp = GetExcelAppIfOpen();
            if (excelApp != null)
                excelApp.ScreenUpdating = true;
        }

        public Excel.Application GetExcelAppIfOpen()
        {
            try
            {
                Excel.Application excelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                return excelApp;
            }

            catch (Exception)
            {
                return null;
            }
        }

        public bool IsWbOpen(Excel.Application excelApp, string path)
        {
            foreach (Excel.Workbook wb in excelApp.Workbooks)
            {
                if (path == wb.FullName)
                    return true;
            }
            return false;
        }

        public static Excel.Range SelectEverythingOnSheet(Excel.Worksheet sheet)
        {
            long lastRow = FindLastRow(sheet);
            long lastCol = FindLastCol(sheet);

            if (lastRow != 0)
                return (sheet.Range["A3", sheet.Cells[lastRow, lastCol]]);
            else
                return (null);
        }

        public static int FindLastRow(Excel.Worksheet sheet)
        {
            try
            {
                return (sheet.Cells.Find("*", After: sheet.Range["A1"], SearchOrder: Excel.XlSearchOrder.xlByRows, SearchDirection: Excel.XlSearchDirection.xlPrevious).Row);
            }
                
            catch (Exception)
            {
                return (0);
            } 
        }

        public static int FindLastCol(Excel.Worksheet sheet)
        {
            try
            {
                return (sheet.Cells.Find("*", After: sheet.Range["A1"], SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlPrevious).Column);
            }

            catch (Exception)
            {
                return (0);
            }
        }
    }
}
