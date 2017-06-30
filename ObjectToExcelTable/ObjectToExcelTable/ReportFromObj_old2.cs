using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XmlWorkbook = DocumentFormat.OpenXml.Spreadsheet.Workbook;
using XmlWorksheet = DocumentFormat.OpenXml.Spreadsheet.Worksheet;
using XmlSheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;
using XmlSheets = DocumentFormat.OpenXml.Spreadsheet.Sheets;
//using mvcTests.Models;
using OfficeOpenXml;
using mvcTests.Models;

namespace ObjectToExcelTable
{
    public class ReportFromObj
    {
        private static Dictionary<string, List<string>> _ObjItems { get; set; } = new Dictionary<string, List<string>>();
        private static List<Dictionary<string, List<string>>> _ListItems { get; set; } = new List<Dictionary<string, List<string>>>();
        public ReportFromObj(object o)
        {
            try
            {
                GetPropertiesOneByOne(o);
            }
            catch (ParameterNotValidException e)
            {
                throw e;
            }
        }

        private void GetPropertiesOneByOne(object o)
        {
            Type t = o.GetType();
            
            if (typeof(IEnumerable).IsAssignableFrom(o.GetType()) && !(o is String))
            {
                throw new ParameterNotValidException("The given parameter cannot be of type" + o.GetType());
                /*foreach (var item in (IEnumerable)o)
                {
                    Type itemT = item.GetType();
                    PropertyInfo[] p = itemT.GetProperties(BindingFlags.Instance | BindingFlags.Public);
                    //За всяко пропърти правим проверка дали си нямаме работа с колекция
                    foreach (PropertyInfo pi in p)
                    {
                        ProcessEachProp(pi, item);
                    }
                }//*/
            }
            else
            {
                PropertyInfo[] p = t.GetProperties(BindingFlags.Instance | BindingFlags.Public);
                foreach (PropertyInfo pi in p)
                {
                    ProcessEachProp(pi, o);
                }
            }
        }
        private void ProcessEachProp(PropertyInfo pi, object o)
        {
            //Ако пропъртито е от тип IEnumerable, извъртаме колекцията и подаваме всеки един обект от нея отново на нашия метод (изключваме String от сметките)
            if (typeof(IEnumerable).IsAssignableFrom(pi.PropertyType)&& pi.GetValue(o)!= null && !(pi.GetValue(o) is String) && pi.CanRead)
            {
                _ListItems.Add(new Dictionary<string, List<string>>());
                foreach (var enumPi in (IEnumerable)pi.GetValue(o))
                {
                    Type propertyType = enumPi.GetType();
                    PropertyInfo[] propInfos = propertyType.GetProperties(BindingFlags.Instance | BindingFlags.Public);
                    foreach (PropertyInfo propInfo in propInfos)
                    {
                        if (propInfo.GetValue(enumPi) is String || !(typeof(IEnumerable).IsAssignableFrom(propInfo.PropertyType)))
                        {
                            if (GetDispAttribute(propInfo.GetCustomAttributes()))
                            {
                                string AttrName = propInfo.GetCustomAttribute<DisplayNameAttribute>().DisplayName;
                                ProcessSimpleTypeProperty(propInfo, enumPi, true, AttrName);
                            }
                        }
                    }
                    //GetPropertiesOneByOne(enumPi);
                }
            }
            else if (pi.CanRead)
            {
                if (GetDispAttribute(pi.GetCustomAttributes()))
                {
                    string AttrName = pi.GetCustomAttribute<DisplayNameAttribute>().DisplayName;
                    ProcessSimpleTypeProperty(pi, o, false, AttrName);
                }
                //ProcessSimpleTypeProperty(pi, o, false);
            }
        }
        private void ProcessSimpleTypeProperty(PropertyInfo pi, object o, bool isCollectionProp, string propDispName)
        {
            if (isCollectionProp)
            {
                if (string.IsNullOrEmpty(propDispName))
                {
                    if (!_ListItems.Last().ContainsKey(pi.Name))
                        _ListItems.Last()[pi.Name] = new List<string>() { pi.GetValue(o, null).ToString() };
                    else
                        _ListItems.Last()[pi.Name].Add(pi.GetValue(o, null).ToString());
                }
                else
                {
                    if (!_ListItems.Last().ContainsKey(propDispName))
                        _ListItems.Last()[propDispName] = new List<string>() { pi.GetValue(o, null).ToString() };
                    else
                        _ListItems.Last()[propDispName].Add(pi.GetValue(o, null).ToString());
                }
            }
            else
            {
                if (string.IsNullOrEmpty(propDispName))
                {
                    if (!_ObjItems.ContainsKey(pi.Name))
                        _ObjItems[pi.Name] = new List<string>() { pi.GetValue(o, null).ToString() };
                    else
                        _ObjItems[pi.Name].Add(pi.GetValue(o, null).ToString());
                }
                else
                {
                    if (!_ObjItems.ContainsKey(propDispName))
                        _ObjItems[propDispName] = new List<string>() { pi.GetValue(o, null).ToString() };
                    else
                        _ObjItems[propDispName].Add(pi.GetValue(o, null).ToString());
                }
            }
        }
        private bool GetDispAttribute(IEnumerable<object> o)
        {
            foreach (var objs in o)
            {
                if (objs is DisplayNameAttribute)
                    return true;
            }
            return false;
        }

        public void ExportByXml()
        {
            ExcelPackage ep = new ExcelPackage();
            ExcelWorkbook xlWBook = ep.Workbook;
            xlWBook.Worksheets.Add("Test Sheet");
            ExcelWorksheet xlWsheet = xlWBook.Worksheets.FirstOrDefault();

            //File path... I have to do that with a Method.....
            string str = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\testXML.xlsx";
            FileInfo fi = new FileInfo(str);

            //Initial row and column index
            int r = 1;
            int c = 1;

            //This will follow the last row and column that has been used..
            int lastRow = 1;
            int lastCol = 1;

            int objStartRow = 1;
            int objStartCol = 1;
            
            //Simple type properties...
            foreach (string key in _ObjItems.Keys)
            {
                try
                {
                    //Value of header cells
                    xlWsheet.SetValue(r, c++, separateWords(key));

                    foreach (string value in _ObjItems[key])
                        xlWsheet.SetValue(r, c++, value);
                }
                catch (Exception)
                {
                    ep.SaveAs(fi);
                    return;
                }
                finally
                {
                    ep.SaveAs(fi);
                    Console.WriteLine("File Saved at: " + fi.DirectoryName);
                }
                
                if (lastCol < c)
                    lastCol = c;
                
                lastRow = ++r;
                c = 1;
            }
            r = 1;
            ExcelRange er = xlWsheet.Cells[r, c, lastRow, lastCol-1];
            er.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
            
            foreach (var cell in er)
            {                
                cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
            }
            
            xlWsheet.Cells.AutoFitColumns();
            ep.SaveAs(fi);
            er[lastRow, c, lastRow, lastCol - 1].Merge = true;
            ep.SaveAs(fi);
            
            //Lets Try to export the List of Items......

            int objEndRow = lastRow;
            int objEndCol = lastCol - 1;

            int listEndCol = 1;
            List<ExcelRange> EmptyRows = new List<ExcelRange>();
            List<ExcelRange> listRanges = new List<ExcelRange>();
            int listStartRow = lastRow;
            int listStartCol = 1;
            foreach (var item in _ListItems)
            {
                //lastRow++;
                listStartCol = 1;
                r = ++listStartRow;
                int numberOfRows = 0;
                //Range tempListStartCell = xlSheet.Cells[r, listStartCol];
                foreach (var key in item.Keys)
                {
                    try
                    {
                        xlWsheet.SetValue(r++, listStartCol, separateWords(key));

                        foreach (string value in item[key])
                        {
                            xlWsheet.SetValue(r++, listStartCol, value);
                        }
                        if (objEndCol < listStartCol)
                            objEndCol = listStartCol;
                        if (listEndCol < listStartCol)
                            listEndCol = listStartCol;
                        listStartCol++;
                        r = listStartRow;
                        if (numberOfRows < item[key].Count)
                        {
                            numberOfRows = item[key].Count;
                        }
                    }
                    catch (Exception e)
                    {
                        
                        return;
                    }
                }

                listStartRow += (numberOfRows + 1);
                //Range tempListEndCell = xlSheet.Cells[listStartRow, listEndCol];
                //Range tempListRange = xlSheet.Range[tempListStartCell, tempListEndCell];
                ExcelRange tempListRange = xlWsheet.Cells[r, listStartCol, listStartRow, listEndCol];
                listRanges.Add(tempListRange);
            }
            ep.SaveAs(fi);
        }
        private void InputValuesToXml(SheetData xmlSheetData)
        {
            foreach(var key in _ObjItems.Keys)
            {
                Row row = new Row();
                row.Append(ConstructCell(key, CellValues.String));
                xmlSheetData.AppendChild(row);
            }
            
        }

        private Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }
        public void ExportToExcel()
        {
            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet xlSheet = xlWorkbook.Sheets[1];

            int r = 1;
            int c = 1;
            int lastRow = 1;
            int lastCol = 1;
            int objStartRow = 1;
            int objStartCol = 1;
            
            foreach (string key in _ObjItems.Keys)
            {
                try
                {
                    (xlSheet.Cells[r, c++] as Range).Value = separateWords(key);
                    foreach (string value in _ObjItems[key])
                        (xlSheet.Cells[r, c++] as Range).Value = value;
                    
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    Console.WriteLine(e.StackTrace);
                    xlWorkbook.Close(false);
                    xlApp.Quit();
                    releaseObject(xlSheet);
                    releaseObject(xlWorkbook);
                    releaseObject(xlApp);
                    return;
                }
                if (lastCol < c)
                    lastCol = c;
                c = 1;
                lastRow = ++r;
            }
            int objEndRow = lastRow;
            int objEndCol = lastCol - 1;

            int listEndCol = 1;
            List<Range> EmptyRows = new List<Range>();
            List<Range> listRanges = new List<Range>();
            int listStartRow = lastRow;
            int listStartCol = 1;
            foreach (var item in _ListItems)
            {
                //lastRow++;
                listStartCol = 1;
                r = ++listStartRow;
                int numberOfRows = 0;
                Range tempListStartCell = xlSheet.Cells[r, listStartCol];
                foreach (var key in item.Keys)
                {
                    try
                    {
                        (xlSheet.Cells[r++, listStartCol] as Range).Value = separateWords(key);

                        foreach (string value in item[key])
                        {
                            (xlSheet.Cells[r++, listStartCol] as Range).Value = value;
                        }
                        if (objEndCol < listStartCol)
                            objEndCol = listStartCol;
                        if (listEndCol < listStartCol)
                            listEndCol = listStartCol;
                        listStartCol++;
                        r = listStartRow;
                        if (numberOfRows < item[key].Count)
                        {
                            numberOfRows = item[key].Count;
                        }                            
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        Console.WriteLine(e.StackTrace);
                        xlWorkbook.Close(false);
                        xlApp.Quit();
                        releaseObject(xlSheet);
                        releaseObject(xlWorkbook);
                        releaseObject(xlApp);
                        return;
                    }                    
                }
                
                listStartRow += (numberOfRows + 1);
                Range tempListEndCell = xlSheet.Cells[listStartRow, listEndCol];
                Range tempListRange = xlSheet.Range[tempListStartCell, tempListEndCell];
                listRanges.Add(tempListRange);
            }

            try
            {
                FormatTable(xlSheet);
                Range startObjRangeCell = xlSheet.Cells[objStartRow, objStartCol];
                Range endObjRangeCell = xlSheet.Cells[objEndRow, objEndCol];
                Range ObjRange = xlSheet.Range[startObjRangeCell, endObjRangeCell];
                EmptyRows.Add(FormatObjRange(ObjRange, xlSheet));

                foreach (Range row in listRanges)
                {
                    if(row != listRanges.Last())
                        EmptyRows.Add(FormatListRange(row, xlSheet));
                    else
                    {
                        FormatListRange(row, xlSheet);
                    }
                }

                foreach (Range row in EmptyRows)
                {
                    row.Merge(true);
                }
                    
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }//*/
            try
            {
                string str = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                Console.WriteLine(str);
                xlWorkbook.SaveAs(str + "\\test.xlsx");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
                xlWorkbook.Close(false);
                xlApp.Quit();
                releaseObject(xlSheet);
                releaseObject(xlWorkbook);
                releaseObject(xlApp);
            }
        }

        private Range FormatListRange(Range listRange, Worksheet xlSheet)
        {
            var r1 = listRange.Cells[1, 1];
            var r2 = listRange.Cells[1, listRange.Columns.Count];

            Range titleRow = xlSheet.Range[r1, r2];

            titleRow.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRow.VerticalAlignment = XlVAlign.xlVAlignCenter;
            titleRow.Font.Bold = true;

            int lastRow = listRange.Rows.Count;
            int lastCol = listRange.Columns.Count;
            Range EmptyRow = xlSheet.Range[listRange.Cells[lastRow, 1], listRange.Cells[lastRow, lastCol]];
            return EmptyRow;
        }

        private Range FormatObjRange(Range objRange, Worksheet xlSheet)
        {            
            Range titleRow = objRange.Range[objRange.Cells[1, 1], objRange.Cells[objRange.Rows.Count, 1]];
            titleRow.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRow.VerticalAlignment = XlVAlign.xlVAlignCenter;
            titleRow.Font.Bold = true;
            
            int lastRow = objRange.Rows.Count;
            int lastCol = objRange.Columns.Count;
            Range EmptyRow = xlSheet.Range[objRange.Cells[lastRow, 1], objRange.Cells[lastRow, lastCol]];
            return EmptyRow;
        }

        private void FormatTable(Worksheet xlWorksheet)
        {
            //Align and autofit the whole table
            Range xlRange = xlWorksheet.UsedRange;
            xlRange.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
            xlRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            xlRange.Columns.AutoFit();
            
            //Places border on each cell            
            foreach (Range c in xlRange.Cells)
            {
                c.BorderAround(XlLineStyle.xlContinuous,
                                        XlBorderWeight.xlThin,
                                        XlColorIndex.xlColorIndexAutomatic,
                                        XlColorIndex.xlColorIndexAutomatic);
            }
        }
        private string separateWords(string str)
        {
            str = str.Trim();
            for (int i = 0; i < str.Length; ++i)
            {
                if (i != 0 && i < str.Length - 1
                    && str[i + 1].ToString() != " "
                    && (str[i + 1].ToString() == (str[i + 1]).ToString().ToUpper())
                    && str[i].ToString() != str[i].ToString().ToUpper()
                    && str[i].ToString() != " ")
                {
                    str = str.Insert(++i, " ");
                }
            }
            return str;
        }

        private void releaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.Write("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }

        }
    }
}
