using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelApp.Models
{
    public class Store
    {
        public int STOREID { get; set; }
        public string STORE_NAME { get; set; }
        public int STORE_ATTRIBUTE_TYPE_CODE { get; set; }
        public string STORE_ATTRIBUTE_TYPE_NAME { get; set; }
        public int STORE_ATTRIBUTE_CODE { get; set; }
        public string STORE_ATTRIBUTE_NAME { get; set; }
    }
}