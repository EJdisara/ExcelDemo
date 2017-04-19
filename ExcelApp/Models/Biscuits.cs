using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelApp.Models
{
    public class Biscuits
    {
        public string apoCategory { get; set; }
        public long id { get; set; }
        public string name { get; set; }
        public string descJ { get; set; }
        public string descE { get; set; }
        public string packTypeDescription { get; set; }
        public string segmentDescription { get; set; }
        public string segment1Description { get; set; }
        public string segment2Description { get; set; }
        public string segment3Description { get; set; }
        public string importLocalDescription { get; set; }
        public string brand { get; set; }
        public string thaiTouristDescription { get; set; }

        public List<Biscuits> biscuiutsList { get; set; }
    }
}