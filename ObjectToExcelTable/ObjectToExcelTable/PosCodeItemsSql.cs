using System;
using System.Collections.Generic;
using System.Linq;
//using System.Web;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace ObjectToExcelTable
{
    public class PosCodeItemsSql
    {
        [DisplayName("Експорт")]
        public string Caption { get; set; }

        public List<PosCodeItemSql> items { get; set; }

        public PosCodeItemsSql(bool emptyClass)
        {
            if (emptyClass)
            {
                items = new List<PosCodeItemSql>();
            }
            else
            {
                //SQL Query here
            }
            Caption = "Позиционни кодове с артикули"; 
        }
    }

    public class PosCodeItemSql
    {
        [DisplayName("Склад")]
        public string StoreName { get; set; } = null;

        [DisplayName("Позиционен код")]
        public string PosCodeName { get; set; } = null;

        [DisplayName("PalletID")]
        public int? PalletID { get; set; } = null;

        [DisplayName("ArticleID")]
        public int? ArticleID { get; set; } = null;

        [DisplayName("Производител")]
        public string Producer { get; set; } = null;

        [DisplayName("Артикул")]
        public string ArticleName { get; set; } = null;

        [DisplayName("Партида")]
        public string ParcelNo { get; set; } = null;

        [DisplayName("Срок на годност")]
        [DisplayFormat(DataFormatString = "MM.yyyy")]
        public DateTime? ExpiryDate { get; set; } = null;

        [DisplayName("К-во")]
        public int? Qty { get; set; } = null;

        [DisplayName("К-во Намерено")]
        public int? QtyFound { get; set; } = null;

    }
}
