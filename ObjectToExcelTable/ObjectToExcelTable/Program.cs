using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Application = Microsoft.Office.Interop.Excel.Application;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.ComponentModel;
using OfficeOpenXml;

namespace ObjectToExcelTable
{
    class Program
    {
        const string filePath = @"C:\Users\yavor.georgiev\Documents\GitHub\ObjectToExcelTable\ObjectToExcelTable\ObjectToExcelTable\bin\Debug\ExcelFile.xlsx";
        //public static Dictionary<string, List<string> > LinkedObjHeaderAndContent { get; set; } = new Dictionary<string, List<string> >();
        //public static Dictionary<string, List<string> > LinkedListHeaderAndContent { get; set; } = new Dictionary<string, List<string> >();
        public static void Main(string[] args)
        {
            PosCodeItemsSql items = new PosCodeItemsSql(true);
            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        file.CopyTo(ms);
                        ObjFromXlFile<PosCodeItemSql> objFromF = new ObjFromXlFile<PosCodeItemSql>();
                        items.items = objFromF.PosCodeFromStream(ms);

                    }
                }
                    
                string tempFilePath = @"C:\Users\yavor.georgiev\Documents\GitHub\ObjectToExcelTable\ObjectToExcelTable\ObjectToExcelTable\bin\Debug\temp.txt";
                //Console.WriteLine(items.Caption);
                
                File.WriteAllText(tempFilePath, items.Caption);
                /*foreach(var item in items.items)
                {
                    Type t = item.GetType();
                    PropertyInfo[] propInfos = t.GetProperties(BindingFlags.Instance | BindingFlags.Public);
                    foreach(PropertyInfo pi in propInfos)
                    {
                        //Console.Write(pi.Name + " -> ");
                        //Console.WriteLine(pi.GetValue(item));
                        
                        File.AppendAllText(tempFilePath, pi.Name + " -> " + pi.GetValue(item) + Environment.NewLine);
                    }
                }//*/
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
                Console.ReadLine();
            }
            Console.ReadLine();
            
            /*PackingListItem pli = new PackingListItem()
            {
                ArticleID = 6263,
                ArticleName = "Аналгин",
                ParcelNo = "32545",
                ExpiryDate = DateTime.Now,
                Qty = 1000,
                PalletID = 4325,
                PalletBarcode = 4324,
                PosCodeID = 12,
                PosCodeName = "Нещо си :)",
                StoreID = 22,
                StoreName = "София"
            };
            PackingListItem pli2 = new PackingListItem()
            {
                ArticleID = 63644,
                ArticleName = "Зопиклон",
                ParcelNo = "3221855",
                ExpiryDate = DateTime.Now,
                Qty = 50000,
                PalletID = 1254444,
                PalletBarcode = 235656456,
                PosCodeID = 50,
                PosCodeName = "Каса",
                StoreID = 23,
                StoreName = "Бургас"
            };
            PackingListItem pli3 = new PackingListItem()
            {
                ArticleID = 12546,
                ArticleName = "Парацетмол",
                ParcelNo = "133532",
                ExpiryDate = DateTime.Now,
                Qty = 3546,
                PalletID = 123,
                PalletBarcode = 4321,
                PosCodeID = 50,
                PosCodeName = "Каса",
                StoreID = 25,
                StoreName = "Пловдив"
            };
            List<PackingListItem> lPli = new List<PackingListItem>();
            lPli.Add(pli);
            lPli.Add(pli2);
            lPli.Add(pli3);
            List<PackingListItem> lPli2 = new List<PackingListItem>();
            lPli2.Add(pli);
            lPli2.Add(pli2);
            lPli2.Add(pli3);
            PackingList pl = new PackingList()
            {
                AppID = 1,
                AppName = "TestApp",
                CustomerID = 2012101,
                CustomerName = "Test Pharmacy",
                DeliveryAddress = "Test str.",
                DocNo = 201,
                DocName = "Test Doc",
                DocDate = DateTime.Now,
                items = lPli,
                items2 = lPli2
            };
            ReportFromObj rfo = new ReportFromObj(pl);
            FileStream fs = new FileStream("\\tempXML.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            MemoryStream ms = new MemoryStream();
            try
            {
                ms = rfo.ExportByXml();
                if (ms != null)
                {
                    ms.WriteTo(fs);
                }
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                Console.ReadLine();
            }//*/
        }
        
    }
}
