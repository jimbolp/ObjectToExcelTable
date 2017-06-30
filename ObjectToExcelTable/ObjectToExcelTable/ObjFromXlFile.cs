using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ObjectToExcelTable
{
    class ObjFromXlFile
    {
        /// <summary>
        /// Holds the header and index of the columns to map them with the respective properties
        /// </summary>
        private Dictionary<string, int> ColumnNumber = new Dictionary<string, int>();


        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="ms"></param>
        /// <returns></returns>
        public List<T> PosCodeFromStream<T>(MemoryStream ms)
        {
            //PosCodeItemsSql Items = new PosCodeItemsSql(true);
            Type t = typeof(T);
            PropertyInfo[] pInfos = t.GetProperties(BindingFlags.Instance | BindingFlags.Public);
            
            ExcelPackage ep;
            ExcelWorksheet xlWsheet;
            try
            {
                ep = new ExcelPackage(ms);
                xlWsheet = ep.Workbook.Worksheets.FirstOrDefault();
            }
            catch (Exception)
            {
                throw;
            }
            List<T> items = new List<T>();
            try
            {
                //items = TakeRange<T>(xlWsheet);
            }
            catch(Exception)
            {
                throw;
            }
            return items;
        }
        private void testReadPropInfo<T>(PropertyInfo pi, T o)
        {
            int i = 0;
            pi.SetValue(o, i);
        }
        /*
        private static List<T> TakeRange<T>(ExcelWorksheet xlWSheet)
        {
            string range = xlWSheet.Dimension.Address;
            ExcelRangeBase er = xlWSheet.Cells[range];
            int startRow = 4;
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
