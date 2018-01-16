using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ObjectToExcelTable
{
    class ObjFromXlFile<T> where T : new()
    {
        /// <summary>
        /// Holds the header and index of the columns to map them with the respective properties
        /// </summary>
        private Dictionary<string, int> ColumnNumber = new Dictionary<string, int>();
        private List<string> NameAttribs = new List<string>();        
        int? _headerRowNumber = null;
        

        public ObjFromXlFile()
        {
            Type t = typeof(T);
            PropertyInfo[] pInfos = t.GetProperties(BindingFlags.Instance | BindingFlags.Public);

            NameAttribs = GetAttrList(pInfos); //Get collection of the DisplayName attributes as string
        }

        /// <summary>
        /// Opens Excel file from memorystream and returns the rows as object from the requested Type 
        /// if the columns match with the properties of the Type
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="ms"></param>
        /// <returns></returns>
        public List<T> PosCodeFromStream(MemoryStream ms)
        {
            if (NameAttribs.Count == 0)
                return null;

            ExcelPackage ep;
            ExcelWorksheet xlWsheet;
            try
            {
                ep = new ExcelPackage(ms);
                xlWsheet = ep.Workbook.Worksheets.FirstOrDefault();
                _headerRowNumber = null;
                foreach(string name in NameAttribs)
                {
                    if (_headerRowNumber != null)
                        break;
                    _headerRowNumber = GetHeaderRow(xlWsheet, name);
                }
                if (!MapColNameAndIdx(xlWsheet))
                    return null;                
            }
            catch (ArgumentNullException)
            {
                throw new ArgumentNullException("Файлът е празен!");
            }
            catch (Exception)
            {
                throw;
            }
            List<T> items = new List<T>();
            try
            {
                items = TakeRange(xlWsheet);

            }
            catch(Exception)
            {
                throw;
            }
            return items;
        }
        /// <summary>
        /// Check each property for DisplayNameAttribute and return list of the DisplayNames
        /// </summary>
        /// <param name="pInfos"></param>
        /// <returns></returns>
        private List<string> GetAttrList(PropertyInfo[] pInfos)
        {
            List<DisplayNameAttribute> attributes = new List<DisplayNameAttribute>();
            foreach(PropertyInfo pi in pInfos)
            {
                if (HasNameAttr(pi.GetCustomAttributes()))
                    attributes.Add(pi.GetCustomAttribute<DisplayNameAttribute>());
            }
            if(attributes.Count != 0)
            {
                List<string> temp = new List<string>();
                foreach(var a in attributes)
                {
                    temp.Add(a.DisplayName);
                }
                return temp;
            }
            else
            {
                return new List<string>();
            }
        }
        private bool MapColNameAndIdx(ExcelWorksheet ws)
        {
            if (_headerRowNumber == null || NameAttribs.Count == 0)
                return false;
            ExcelRangeBase range = ws.Cells[ws.Dimension.Address];

            foreach(string attrName in NameAttribs)
            {
                for (int col = 1; col <= range.Columns; ++col)
                {
                    if (attrName == ws.Cells[_headerRowNumber.Value, col].Value.ToString())
                    {
                        if (!ColumnNumber.Keys.Any(k => k == attrName))
                            ColumnNumber[attrName] = col;
                    }
                }
            }
            return true;
        }
        /// <summary>
        /// Check if there are any DisplayName Attributes
        /// </summary>
        /// <param name="objs"></param>
        /// <returns></returns>
        private bool HasNameAttr(IEnumerable<object> objs)
        {
            if (objs == null || objs.Count() == 0)
                return false;
            foreach(object o in objs)
            {
                if (o is DisplayNameAttribute)
                    return true;
            }
            return false;
        }

        private int? GetHeaderRow(ExcelWorksheet ws, string columnName)
        {
            ExcelRangeBase range = ws.Cells[ws.Dimension.Address];
            if(range == null)
                return null;
            for(int i = 1; i <= range.Rows; ++i)
            {
                for(int j = 1; j <= range.Columns; ++j)
                {
                    string _cellValue = "";
                    if (ws.Cells[i, j].Value != null)
                        _cellValue = ws.Cells[i, j].Value.ToString();
                    //else continue;
                    if (_cellValue.Trim() == columnName.Trim())
                        return i;
                }
            }
            return null;
        }
        
        private List<T> TakeRange(ExcelWorksheet xlWSheet)
        {
            string range = xlWSheet.Dimension.Address;
            ExcelRangeBase er = xlWSheet.Cells[range];
            int startRow = _headerRowNumber.Value + 1;
            int endRow = er.Rows;
            List<T> listT = new List<T>();
            for (int i = startRow; i <= endRow; ++i)
            {
                T tempObj = new T();
                foreach(KeyValuePair<string, int> d in ColumnNumber)
                {
                    try
                    {
                        PropertyInfo pi = GetPropertyInfoFromName(d.Key, tempObj);
                        Type t = pi.PropertyType;
                        if (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(System.Nullable<>))
                        {
                            t = Nullable.GetUnderlyingType(t);
                        }
                        switch (Type.GetTypeCode(t))
                        {
                            case TypeCode.String:
                                pi.SetValue(tempObj, xlWSheet.Cells[i, d.Value].Value.ToString());
                                break;
                            case TypeCode.Empty:
                                break;
                            case TypeCode.Int32:
                                int pInt = 0;
                                if (int.TryParse(xlWSheet.Cells[i, d.Value].Value.ToString(), out pInt))
                                {
                                    pi.SetValue(tempObj, pInt);
                                }
                                break;
                            case TypeCode.DateTime:
                                DateTime dTime;
                                if (DateTime.TryParse(xlWSheet.Cells[i, d.Value].Value.ToString(), out dTime))
                                {
                                    pi.SetValue(tempObj, dTime);
                                }
                                break;
                            case TypeCode.Decimal:
                                decimal dec;
                                if (decimal.TryParse(xlWSheet.Cells[i, d.Value].Value.ToString(), out dec))
                                {
                                    pi.SetValue(tempObj, dec);
                                }
                                break;
                            case TypeCode.Double:
                                double dVal;
                                if (double.TryParse(xlWSheet.Cells[i, d.Value].Value.ToString(), out dVal))
                                {
                                    pi.SetValue(tempObj, dVal);
                                }
                                break;
                            case TypeCode.Object:
                                {
                                    //TODO...
                                }
                                break;
                            default:
                                break;                                
                        }
                                              
                        //listT.Add(FillTheList(d.Key, xlWSheet.Cells[i, d.Value].Value));
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }
                listT.Add(tempObj);
            }
            return listT;
        }

        private PropertyInfo GetPropertyInfoFromName(string key, T tempObj)
        {
            Type t = tempObj.GetType();
            
            foreach(PropertyInfo pi in t.GetProperties())
            {
                try
                {
                    if (pi.GetCustomAttribute<DisplayNameAttribute>().DisplayName == key)
                    {
                        return pi;
                    }
                }
                catch (ArgumentNullException)
                {
                    continue;
                }
            }
            throw new Exception("No property with DisplayName Attribute");
            
        }

        /*
private List<T> TakeRange(ExcelWorksheet xlWSheet)
{
   string range = xlWSheet.Dimension.Address;
   ExcelRangeBase er = xlWSheet.Cells[range];
   int startRow = _headerRowNumber.Value + 1;
   int endRow = er.Rows;

   PosCodeItemsSql Doc = new PosCodeItemsSql(true);
   for (int i = startRow; i <= endRow; ++i)
   {
       int? palletID = null;
       int? articleID = null;
       int? qty = null;
       int? qtyFound = null;
       DateTime? expiryDate = null;
       int nullableInt = 0;

       if (int.TryParse((xlWSheet.Cells[i, ColumnName.PalletID, i, ColumnName.PalletID].Value).ToString(), out nullableInt))
       {
           palletID = nullableInt;
       }
       if (int.TryParse((xlWSheet.Cells[i, ColumnName.ArticleID, i, ColumnName.ArticleID].Value).ToString(), out nullableInt))
       {
           articleID = nullableInt;
       }                
       DateTime d;
       if (DateTime.TryParse((xlWSheet.Cells[i, ColumnName.ExpiryDate, i, ColumnName.ExpiryDate].Value).ToString(), out d))
       {
           expiryDate = d;
       }
       if (int.TryParse((xlWSheet.Cells[i, ColumnName.Qty, i, ColumnName.Qty].Value).ToString(), out nullableInt))
       {
           qty = nullableInt;
       }
       if (int.TryParse((xlWSheet.Cells[i, ColumnName.QtyFound, i, ColumnName.QtyFound].Value).ToString(), out nullableInt))
       {
           qtyFound = nullableInt;
       }

       try
       {
           Doc.items.Add(new PosCodeItemSql()
           {
               StoreName = (xlWSheet.Cells[i, ColumnName.StoreName, i, ColumnName.StoreName].Value).ToString(),
               PosCodeName = (xlWSheet.Cells[i, ColumnName.PosCodeName, i, ColumnName.PosCodeName].Value).ToString(),
               PalletID = palletID,
               ArticleID = articleID,
               Producer = (xlWSheet.Cells[i, ColumnName.Producer, i, ColumnName.Producer].Value).ToString(),
               ArticleName = (xlWSheet.Cells[i, ColumnName.ArticleName, i, ColumnName.ArticleName].Value).ToString(),
               ParcelNo = (xlWSheet.Cells[i, ColumnName.ParcelNo, i, ColumnName.ParcelNo].Value).ToString(),
               ExpiryDate = expiryDate,
               Qty = qty,
               QtyFound = qtyFound
           });
       }
       catch (Exception)
       {
           throw;
       }
   }
   return Doc;
}//*/
        private static int? ColumnCheck(PosCodeItemSql item)
        {
            Type t = item.GetType();
            PropertyInfo[] pInfs = t.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo pi in pInfs)
            {

            }
            return null;
        }
    }
}
