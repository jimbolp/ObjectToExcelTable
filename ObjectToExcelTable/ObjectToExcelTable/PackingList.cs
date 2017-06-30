using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
//using System.Web;

namespace ObjectToExcelTable
{
    public class PropDisplay
    {
        public string PropName { get; set; }
        public string PropDisplayName { get; set; }
        public bool isVisible { get; set; }
        public string PropFormat { get; set; } // "0.00 %"
        public bool isFormated
        { get
            {
                if (string.IsNullOrEmpty(this.PropFormat))
                    return false;
                else
                    return true;
            }
                
        }
    }


    public class PropsDisplay
    {
        List<PropDisplay> props { get; set; }

        public PropsDisplay()
        {
            props = new List<PropDisplay>();
        }

        public void export(Object obj, PropsDisplay props = null)
        {
            if(props != null)
            {
                // imash imena na poleta
            }

        }
    }

    

    public class PackingList
    {
        private class _sale
        {
            public int CustomerID { get; set; }
            public string CustomerName { get; set; }
            public string DeliveryAddress { get; set; }
            public string DocName { get; set; }
            public int DocNo { get; set; }
            public DateTime DocDate { get; set; }
        }

        private const string sqlSale = "select s.CustomerID, c.Name as CustomerName, c.ShipmentAddress as DeliveryAddress, " +
                                       "d.ShortName as DocName, s.DocNo, s.DocDate from {0}.dbo.Sale s with(nolock) " +
                                       "left join {0}.Dbo.Customer c with(nolock) on s.CustomerID = c.ID " +
                                       "left join {0}.dbo.DocType d with (nolock) on d.ID = s.DocTypeID " +
                                       "where s.ID= {1}";

        private const string sqlSaleItem = "";

        public int AppID { get; set; }
        //[DisplayName("")]
        public string AppName { get; set; }
        [DisplayName("Клиент Номер")]
        public int CustomerID { get; set; }
        [DisplayName("Име на Клиент")]
        public string CustomerName { get; set; }
        [DisplayName("Адрес на доставка")]
        public string DeliveryAddress { get; set; }
        [DisplayName("Име на Документ")]
        public string DocName { get; set; }
        [DisplayName("Номер на Документ")]
        public int DocNo { get; set; }
        [DisplayName("Дата на Документ")]
        public DateTime DocDate { get; set; }

        public List<PackingListItem> items { get; set; }
        public List<PackingListItem> items2 { get; set; }


        public PackingList()
        {
            items = new List<PackingListItem>();
            items2 = new List<PackingListItem>();
        }

        /*public PackingList(int appID, int saleID)
        {
            PositioningEntities2 db = new PositioningEntities2();
            App app = db.Apps.Find(appID);
            if (db != null && app != null)
            {
                AppID = appID;
                AppName = app.Name;

                string sql = string.Format(sqlSale, app.DBName, saleID);
                _sale sale = db.Database.SqlQuery<_sale>(sql).FirstOrDefault();

                if (sale != null)
                {
                    CustomerID = sale.CustomerID;
                    CustomerName = sale.CustomerName;
                    DeliveryAddress = sale.DeliveryAddress;
                    DocName = sale.DocName;
                    DocNo = sale.DocNo;
                    DocDate = sale.DocDate;

                    items = new List<PackingListItem>();



                }
                else
                {
                    throw new Exception("Търсения документ за продажба не беше открит!");
                }

            }
            else
            {
                throw new Exception("Грешка в инициализацията на базата данни");
            }


        }//*/
    }

    public class PackingListItem
    {
        //private PackingListItem() { }
        [DisplayName("Артикул №")]
        public int ArticleID { get; set; }
        [DisplayName("Име на Артикул")]
        public string ArticleName { get; set; }
        [DisplayName("Партида")]
        public string ParcelNo { get; set; }
        [DisplayName("Годен до")]
        public DateTime ExpiryDate { get; set; }
        [DisplayName("Наличност")]
        public int Qty { get; set; }
        public int PalletID { get; set; }
        [DisplayName("Баркод на Пале")]
        public int PalletBarcode { get; set; }
        [DisplayName("Поз. Код №")]
        public int PosCodeID { get; set; }
        [DisplayName("Поз Код Име")]
        public string PosCodeName { get; set; }
        public int StoreID { get; set; }
        [DisplayName("Склад")]
        public string StoreName { get; set; }
        
    }

}
