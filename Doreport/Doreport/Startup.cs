using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using ConsoleApp1;

namespace DoReport
{
    public class Startup
    {
        private Excel.Application excelapp;
        public async Task<object> GenerateReoprtAsync(dynamic input)
        {
            
            try
            {
                string strFilePath = (string)input.filename;
                dynamic sheets = (dynamic)input.sheets;

                excelapp = new Excel.Application();

                Excel.Workbooks wbs = excelapp.Workbooks;
                Excel.Workbook wbResultReoprt = CreateWorkbook(strFilePath, excelapp);

                if (null != wbResultReoprt)
                {
                    //if (sheets.Length > 0)
                    if (sheets.Count> 0)
                    {
                        foreach (dynamic sheet in sheets)
                        {
                            string sheetname = (string)sheet.name;
                            dynamic content = (dynamic)sheet.content;
                            
                            HandleWorksheet(sheetname, wbResultReoprt,content);
                        }
                    }

                    Excel.Worksheet wsStab = wbResultReoprt.Worksheets[1];
                    Excel.Shape shape = wsStab.Shapes.AddChart(XlChartType.xl3DArea,25,25,400,300);
                    Excel.Shape shape1 = wsStab.Shapes.AddChart(XlChartType.xlBarStacked,450,25,400,300);

                    Excel.Chart chart1 = shape.Chart;
                    chart1.HasTitle = true;
                    chart1.ChartTitle.Text = "test";
                    chart1.ChartTitle.Interior.Color = "Red";

                    
                    wbResultReoprt.SaveAs(strFilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    CloseWorkbook(wbResultReoprt);

                    excelapp.Quit();
                    GC.Collect();
                    KillExcelProcess.Kill(excelapp);

                    return true;
                }
                return false;

            }
            catch (Exception e)
            {
                throw e;
                if (null != excelapp)
                {
                    excelapp.Quit();
                    GC.Collect();
                    KillExcelProcess.Kill(excelapp);
                }
                //File.AppendAllText(@"C:\Users\97901\Project\ZD Data Logger", e.Message);
                return false;
            }
        }

        private Excel.Workbook CreateWorkbook(string strSavePath,Excel.Application excelapp)
        {

            Excel.Workbook wb = excelapp.Workbooks.Add(true);      //for create a new workbook,the argument must be true

            return wb;
        }

        private Excel.Workbook OpenWorkbook(string path,Excel.Application excelapp)
        {

            Excel.Workbook wb = excelapp.Workbooks.Open(path);

            return wb;
        }

        private void HandleWorksheet(string strWSName,Excel.Workbook wb,dynamic content)
        {
            Excel.Worksheet ws = wb.Worksheets.Add();
            ws.Name = strWSName;

            foreach (dynamic cellorchart in content)
            {
                if ("cell" == (string)cellorchart.type)
                {
                    int row = (int)cellorchart.row;
                    int col = (int)cellorchart.column;
                    string value = (string)cellorchart.value;
                    string formula = (string)cellorchart.formula;

                    ws.Cells[row, col] = value;
                    if ("" != formula)
                    {
                        ws.Cells[row, col].Formula = formula;
                    }
                }
                else if("chart" == (string)cellorchart.type)
                {
                    string title = (string)cellorchart.title;
                    string legend = (string)cellorchart.legend;
                    string datasrc = (string)cellorchart.datasrc;
                }
            }
        }

        private void CloseWorkbook(Excel.Workbook ExcelWorkbook)
        {
            ExcelWorkbook.Close(false, Type.Missing, Type.Missing);
            ExcelWorkbook = null;
        }

        private class KillExcelProcess
        {
            [DllImport("User32.dll", CharSet = CharSet.Auto)]
            public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
            public static void Kill(Microsoft.Office.Interop.Excel.Application excel)
            {
                try
                {
                    IntPtr t = new IntPtr(excel.Hwnd);    //get the Handler of excel app，具体作用是得到这块内存入口
                    int k = 0;
                    GetWindowThreadProcessId(t, out k);   //get the Process id
                    System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //get the reference of this process
                    p.Kill();     //kill the process k
                }
                catch (System.Exception ex)
                {
                    throw ex;
                }
            }
        }
    }


}
