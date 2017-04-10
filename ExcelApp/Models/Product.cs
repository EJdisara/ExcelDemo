using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;


namespace ExcelApp.Models
{
    public class Product
    {
        public long APO_SUBCLASS { get; set; }
        public int SUBCLASS { get; set; }
        public string SUB_NAME { get; set; }
        public string addIn { get; set; }
        public int CLASS { get; set; }
        public string CLASS_NAME { get; set; }
        public int DEPT { get; set; }
        public string DEPT_NAME { get; set; }
        public int GROUP_NO { get; set; }
        public string GROUP_NAME { get; set; }
        public int DIVISION { get; set; }
        public string DIV_NAME { get; set; }
    }

    
}