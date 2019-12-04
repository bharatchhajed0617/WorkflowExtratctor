using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CodeAnalyserConsoleApp
{
    class Excel
    {

        public bool ExportDataTableToExcel(System.Data.DataTable dt, string filepath, string sheetName)
        {
            bool isOpned;
            Application app;
            Workbook wb;
            Worksheet ws;
            Range oRange;

            try
            {
                app = null;
                wb = null;
                try
                {
                    app = Marshal.GetActiveObject("Excel.Application") as Application;
                    foreach (Workbook workbook in app.Workbooks)
                    {
                        if (filepath.ToLower().Equals(workbook.FullName.ToLower()))
                        {
                            wb = workbook;
                            isOpned = true;
                            break;
                        }
                    }
                }
                catch (Exception)
                {
                }
                // Start Excel and get Application object. 
                if (wb == null)
                {
                    object missing = Type.Missing;
                    app = new Application();
                    app.Visible = true;

                    bool excelExist = File.Exists(filepath);
                    if (excelExist)
                    {
                        // open the workbook. 
                        wb = app.Workbooks.Open(
                        filepath, true
                      , false, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                    }
                    else
                    {
                        string dir = Path.GetDirectoryName(filepath);
                        bool dirExist = Directory.Exists(dir);
                        if (!dirExist)
                        {
                            dirExist = Directory.Exists(Path.Combine(Environment.CurrentDirectory, dir));
                            if (dirExist)
                                filepath = Path.Combine(Environment.CurrentDirectory, filepath);
                        }
                        wb = app.Workbooks.Add();
                        wb.SaveAs(filepath);
                    }
                }


                // Get the Active sheet 
             
                        ws = wb.Worksheets.Add();
                        ws.Name = sheetName;
                ws.Activate();

              


                int rowCount = 1;
            foreach (DataRow dr in dt.Rows)
            {
                rowCount += 1;
                for (int i = 1; i < dt.Columns.Count + 1; i++)
                {
                    // Add the header the first time through 
                    if (rowCount == 2)
                    {
                        ws.Cells[1, i] = dt.Columns[i - 1].ColumnName;
                    }
                    ws.Cells[rowCount, i] = dr[i - 1].ToString();
                }
            }

            // Resize the columns 
            oRange = ws.UsedRange;
            oRange.EntireColumn.AutoFit();

            // Save the sheet and close 
            ws = null;
            oRange = null;
                wb.Save();
            wb.Close(Missing.Value, Missing.Value, Missing.Value);
            wb = null;
            app.Quit();
        }
            catch
            {
                throw;
            }
            finally
            {
                // Clean up 
                // NOTE: When in release mode, this does the trick 
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

            return true;
        }

        public bool ExportDataTableToExcel(System.Data.DataSet ds, string filepath)
        {
            bool isOpned;
            Application app;
            Workbook wb;
            Worksheet ws;
            Range oRange;

            try
            {
                app = null;
                wb = null;
                try
                {
                    app = Marshal.GetActiveObject("Excel.Application") as Application;
                    foreach (Workbook workbook in app.Workbooks)
                    {
                        if (filepath.ToLower().Equals(workbook.FullName.ToLower()))
                        {
                            wb = workbook;
                            isOpned = true;
                            break;
                        }
                    }
                }
                catch (Exception)
                {
                }
                // Start Excel and get Application object. 
                if (wb == null)
                {
                    object missing = Type.Missing;
                    app = new Application();
                    app.Visible = true;

                    bool excelExist = File.Exists(filepath);
                    if (excelExist)
                    {
                        // open the workbook. 
                        wb = app.Workbooks.Open(
                        filepath, true
                      , false, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                    }
                    else
                    {
                        string dir = Path.GetDirectoryName(filepath);
                        bool dirExist = Directory.Exists(dir);
                        if (!dirExist)
                        {
                            dirExist = Directory.Exists(Path.Combine(Environment.CurrentDirectory, dir));
                            if (dirExist)
                                filepath = Path.Combine(Environment.CurrentDirectory, filepath);
                        }
                        wb = app.Workbooks.Add();
                        wb.SaveAs(filepath);
                    }
                }


                // Get the Active sheet 

       


                foreach (System.Data.DataTable dt in ds.Tables)
                {
                    ws = wb.Worksheets.Add();
                    ws.Name = dt.TableName;
                    ws.Activate();
                    int rowCount = 1;
                    foreach (DataRow dr in dt.Rows)
                    {
                        rowCount += 1;
                        for (int i = 1; i < dt.Columns.Count + 1; i++)
                        {
                            // Add the header the first time through 
                            if (rowCount == 2)
                            {
                                ws.Cells[1, i] = dt.Columns[i - 1].ColumnName;
                            }
                            ws.Cells[rowCount, i] = dr[i - 1].ToString();
                        }
                    }
                   
                    oRange = ws.Range["A1"].EntireRow;
                    Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);

                    String address = last.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    oRange=ws.Range["A1", address.Split('$')[1]+"1"];
                    oRange.Interior.Color= System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Tomato);
                    oRange = ws.UsedRange;

                    Borders borders = oRange.Borders;
                    borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
                    borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
                    borders[XlBordersIndex.xlDiagonalDown].LineStyle = XlLineStyle.xlDot;
                    // oRange.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    oRange.EntireColumn.AutoFit();

                    // Save the sheet and close 
                    ws = null;
                    oRange = null;
                }

          

                // Resize the columns 

                wb.Save();
                wb.Close(Missing.Value, Missing.Value, Missing.Value);
                wb = null;
                app.Quit();
            }
            catch
            {
                throw;
            }
            finally
            {
                // Clean up 
                // NOTE: When in release mode, this does the trick 
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

            return true;
        }


    }
}
