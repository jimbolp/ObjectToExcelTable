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

namespace ObjectToExcelTable
{
    class ReportFromObj
    {
        private static Dictionary<string, List<string>> _ObjItems { get; set; } = new Dictionary<string, List<string>>();
        private static List<Dictionary<string, List<string>>> _ListItems { get; set; } = new List<Dictionary<string, List<string>>>();
        public ReportFromObj(object o)
        {
            GetPropertiesOneByOne(o);
        }

        private void GetPropertiesOneByOne(object o)
        {
            Type t = o.GetType();
            PropertyInfo[] p = t.GetProperties(BindingFlags.Instance | BindingFlags.Public);

            //За всяко пропърти правим проверка дали си нямаме работа с колекция
            foreach (PropertyInfo pi in p)
            {

                //Ако пропъртито е от тип IEnumerable, извъртаме колекцията и подаваме всеки един обект от нея отново на нашия метод (изключваме String от сметките)
                if (typeof(IEnumerable).IsAssignableFrom(pi.PropertyType) && !(pi.GetValue(o) is String) && pi.CanRead)
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
        }

        private void ProcessSimpleTypeProperty(PropertyInfo pi, object o, bool isCollectionProp, string propDispName)
        {
            if (isCollectionProp)
            {
                if (!_ListItems.Last().ContainsKey(propDispName))
                {
                    if (string.IsNullOrEmpty(propDispName))
                        _ListItems.Last()[pi.Name] = new List<string>() { pi.GetValue(o, null).ToString() };
                    else
                        _ListItems.Last()[propDispName] = new List<string>() { pi.GetValue(o, null).ToString() };
                }
                else
                {
                    if (string.IsNullOrEmpty(propDispName))
                        _ListItems.Last()[pi.Name].Add(pi.GetValue(o, null).ToString());
                    else
                        _ListItems.Last()[propDispName].Add(pi.GetValue(o, null).ToString());
                }
            }
            else
            {
                if (!_ObjItems.ContainsKey(pi.Name))
                {
                    if (string.IsNullOrEmpty(propDispName))
                        _ObjItems[pi.Name] = new List<string>() { pi.GetValue(o, null).ToString() };
                    else
                        _ObjItems[propDispName] = new List<string>() { pi.GetValue(o, null).ToString() };
                }
                else
                {
                    if (string.IsNullOrEmpty(propDispName))
                        _ObjItems[pi.Name].Add(pi.GetValue(o, null).ToString());
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

        public void ExportToHtml()
        {
            string str = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "testXML.xlsx";
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(str, SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new XmlWorkbook();
                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new XmlWorksheet(new SheetData());
                XmlSheets sheets = workbookPart.Workbook.AppendChild(new XmlSheets());
                XmlSheet sheet = new XmlSheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Test Sheet" };
                sheets.Append(sheet);

                workbookPart.Workbook.Save();
            }
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
            int objEndCol = lastCol;

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
