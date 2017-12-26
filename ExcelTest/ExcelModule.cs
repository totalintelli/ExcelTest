using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTest
{
    class ExcelModule
    {
        public static void RunTest()
        {
            List<string> testData = new List<string>()
            { "캔버스의 Y값", "마우스의 수직 이동거리", "도형의 회전 각도", "tranformOrigin" };

            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                // Excel 첫번째 워크시트 가져오기                
                excelApp = new Excel.Application();
                wb = excelApp.Workbooks.Add();
                ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;

                // 데이타 넣기
                int r = 1;
                foreach (var d in testData)
                {
                    ws.Cells[1, r] = d;
                    r++;
                }

                // 엑셀파일 저장
                wb.SaveAs(@"C:\temp\test.xls", Excel.XlFileFormat.xlWorkbookNormal);
                wb.Close(true);
                excelApp.Quit();
            }
            finally
            {
                // Clean up
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(excelApp);
            }
        }

        private static void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
