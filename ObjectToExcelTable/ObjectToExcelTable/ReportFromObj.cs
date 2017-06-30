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
using OfficeOpenXml;

namespace ObjectToExcelTable
{
    public class ReportFromObj
    {
        private Dictionary<string, List<string>> _ObjItems { get; set; } = new Dictionary<string, List<string>>();
        private List<Dictionary<string, List<string>>> _ListItems { get; set; } = new List<Dictionary<string, List<string>>>();
        public ReportFromObj(object o)
        {
            try
            {
                GetPropertiesOneByOne(o);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private void GetPropertiesOneByOne(object o)
        {
            Type t = o.GetType();

            if (typeof(IEnumerable).IsAssignableFrom(o.GetType()) && !(o is String))
            {
                throw new Exception("The given parameter cannot be of type" + o.GetType());
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
            if (typeof(IEnumerable).IsAssignableFrom(pi.PropertyType) && pi.GetValue(o) != null && !(pi.GetValue(o) is String) && pi.CanRead)
            {
                _ListItems.Add(new Dictionary<string, List<string>>());
                foreach (var enumPi in (IEnumerable)pi.GetValue(o))
                {
                    Type propertyType = enumPi.GetType();
                    PropertyInfo[] propInfos = propertyType.GetProperties(BindingFlags.Instance | BindingFlags.Public);
                    foreach (PropertyInfo propInfo in propInfos)
                    {
                        var test = propInfo.GetValue(enumPi);
                        var testType = typeof(IEnumerable).IsAssignableFrom(propInfo.PropertyType);

                        if (propInfo.GetValue(enumPi) is String || !(typeof(IEnumerable).IsAssignableFrom(propInfo.PropertyType)))
                        {
                            string AttrName = propInfo.Name;
                            if (GetDispAttribute(propInfo.GetCustomAttributes()))
                            {
                                AttrName = propInfo.GetCustomAttribute<DisplayNameAttribute>().DisplayName;
                            }
                            ProcessSimpleTypeProperty(propInfo, enumPi, true, AttrName);
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
            string PropValue = "";
            if (pi.GetValue(o, null) != null)
                PropValue = pi.GetValue(o, null).ToString();

            if (isCollectionProp)
            {
                if (string.IsNullOrEmpty(propDispName))
                {
                    if (!_ListItems.Last().ContainsKey(pi.Name))
                        _ListItems.Last()[pi.Name] = new List<string>() { PropValue };
                    else
                        _ListItems.Last()[pi.Name].Add(PropValue);
                }
                else
                {
                    if (!_ListItems.Last().ContainsKey(propDispName))
                        _ListItems.Last()[propDispName] = new List<string>() { PropValue };
                    else
                        _ListItems.Last()[propDispName].Add(PropValue);
                }
            }
            else
            {
                if (string.IsNullOrEmpty(propDispName))
                {
                    if (!_ObjItems.ContainsKey(pi.Name))
                        _ObjItems[pi.Name] = new List<string>() { PropValue };
                    else
                        _ObjItems[pi.Name].Add(PropValue);
                }
                else
                {
                    if (!_ObjItems.ContainsKey(propDispName))
                        _ObjItems[propDispName] = new List<string>() { PropValue };
                    else
                        _ObjItems[propDispName].Add(PropValue);
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

        public MemoryStream ExportByXml()
        {
            MemoryStream ms = new MemoryStream();
            ExcelPackage ep = new ExcelPackage(ms);
            ExcelWorkbook xlWBook = ep.Workbook;
            xlWBook.Worksheets.Add("Report");
            ExcelWorksheet xlWsheet = xlWBook.Worksheets.FirstOrDefault();

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
                    return new MemoryStream(ep.GetAsByteArray()); ;
                }

                if (lastCol < c)
                    lastCol = c;

                lastRow = ++r;
                c = 1;
            }

            //Lets Try to export the List of Items......

            int objEndRow = lastRow;
            int objEndCol = lastCol - 1;

            int listEndCol = 1;
            List<ExcelRangeBase> EmptyRows = new List<ExcelRangeBase>();
            List<ExcelRange> listRanges = new List<ExcelRange>();
            int listStartRow = lastRow;
            int listStartCol = 1;
            foreach (var item in _ListItems)
            {
                //lastRow++;
                listStartCol = 1;
                r = ++listStartRow;
                int numberOfRows = 0;
                int tempListStartRow = r;
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
                    catch (Exception)
                    {
                        return new MemoryStream(ep.GetAsByteArray());
                    }
                }

                listStartRow += (numberOfRows + 1);
                ExcelRange tempListRange = xlWsheet.Cells[tempListStartRow, 1, listStartRow, listEndCol];
                listRanges.Add(tempListRange);
            }
            //ep.SaveAs(fi);

            try
            {
                FormatTable(xlWsheet);

                ExcelRange ObjRange = xlWsheet.Cells[objStartRow, objStartCol, objEndRow, objEndCol];
                var tempStr = ObjRange.Address;
                EmptyRows.Add(FormatObjRange(ObjRange, xlWsheet));

                foreach (ExcelRange row in listRanges)
                {
                    if (row != listRanges.Last())
                    {
                        EmptyRows.Add(FormatListRange(row, xlWsheet));
                    }
                    else
                    {
                        FormatListRange(row, xlWsheet);
                    }
                }

                foreach (ExcelRange row in EmptyRows)
                {
                    row.Merge = true;
                }
                //ep.SaveAs(fi);
            }
            catch (Exception)
            {
                ms = new MemoryStream(ep.GetAsByteArray());
                return ms;
                //Console.WriteLine(e);
            }//*/

            return new MemoryStream(ep.GetAsByteArray());
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
                    if (row != listRanges.Last())
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
        private ExcelRange FormatListRange(ExcelRange listRange, ExcelWorksheet xlSheet)
        {
            string keepInitialAddress = listRange.Address;
            int firstRow = listRange.Start.Row;
            int firstCol = listRange.Start.Column;
            ExcelRange title = listRange[firstRow, firstCol, firstRow, listRange.Columns];
            title.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            title.Style.Font.Bold = true;

            listRange.Address = keepInitialAddress;
            int lastRow = listRange.End.Row;
            int lastCol = listRange.End.Column;
            return listRange[lastRow, firstCol, lastRow, lastCol];
        }
        private Range FormatListRange(Range listRange, Worksheet xlSheet)
        {
            Range r1 = listRange.Cells[1, 1];
            Range r2 = listRange.Cells[1, listRange.Columns.Count];

            Range titleRow = xlSheet.Range[r1, r2];

            titleRow.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRow.VerticalAlignment = XlVAlign.xlVAlignCenter;
            titleRow.Font.Bold = true;

            int lastRow = listRange.Rows.Count;
            int lastCol = listRange.Columns.Count;
            Range EmptyRow = xlSheet.Range[listRange.Cells[lastRow, 1], listRange.Cells[lastRow, lastCol]];
            return EmptyRow;
        }
        private ExcelRange FormatObjRange(ExcelRange xlRange, ExcelWorksheet xlWsheet)
        {
            string keepInitialAddress = xlRange.Address;
            int firstRow = xlRange.Start.Row;
            int firstCol = xlRange.Start.Column;
            ExcelRange title = xlRange[firstRow, firstCol, xlRange.Rows, firstCol];
            title.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            title.Style.Font.Bold = true;

            xlRange.Address = keepInitialAddress;
            int lastRow = xlRange.End.Row;
            int lastCol = xlRange.End.Column;

            return xlRange[lastRow, firstCol, lastRow, lastCol];
        }
        private Range FormatObjRange(Range objRange, Worksheet xlSheet)
        {
            Range r1 = objRange.Cells[1, 1];
            Range r2 = objRange.Cells[objRange.Rows.Count, 1];
            Range titleRow = objRange.Range[r1, r2];
            titleRow.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRow.VerticalAlignment = XlVAlign.xlVAlignCenter;
            titleRow.Font.Bold = true;

            int lastRow = objRange.Rows.Count;
            int lastCol = objRange.Columns.Count;
            r1 = objRange.Cells[lastRow, 1];
            r2 = objRange.Cells[lastRow, lastCol];
            Range EmptyRow = xlSheet.Range[r1, r2];
            return EmptyRow;
        }

        //Uses EPPlus lib
        private void FormatTable(ExcelWorksheet xlWsheet)
        {
            ExcelRange xlRange = xlWsheet.SelectedRange[xlWsheet.Dimension.Address];
            xlRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
            xlRange.AutoFitColumns();
            foreach (var cell in xlRange)
            {
                cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
            }
        }

        //Uses COM Object Excel
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

