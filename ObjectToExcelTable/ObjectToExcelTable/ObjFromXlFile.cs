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
    static class ObjFromXlFile
    {
        private static class ColumnName
        {
            public const int StoreName = 1;
            public const int PosCodeName = 2;
            public const int PalletID = 3;
            public const int ArticleID = 4;
            public const int Producer = 5;
            public const int ArticleName = 6;
            public const int ParcelNo = 7;
            public const int ExpiryDate = 8;
            public const int Qty = 9;
            public const int QtyFound = 10;
        }
        public static PosCodeItemsSql PosCodeFromStream(MemoryStream ms)
        {
            PosCodeItemsSql Items = new PosCodeItemsSql(true);
            //FileInfo fi;
            //MemoryStream ms;

            ExcelPackage ep;
            //ExcelWorkbook xlWBook;
            ExcelWorksheet xlWsheet;
            try
            {
                //fi = new FileInfo(path);
                ep = new ExcelPackage(ms);

                xlWsheet = ep.Workbook.Worksheets.FirstOrDefault();
            }
            catch (Exception)
            {
                throw;
            }
            PosCodeItemsSql pcItems;
            try
            {
                pcItems = TakeRange(xlWsheet);
            }
            catch(Exception)
            {
                throw;
            }
            return pcItems;
        }
        private static PosCodeItemsSql TakeRange(ExcelWorksheet xlWSheet)
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
        }
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
