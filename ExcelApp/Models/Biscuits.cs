using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelApp.Models
{
    public class Biscuits
    {
        public string APO_CATEGORY { get; set; }
        public long ID { get; set; }
        public string NAME { get; set; }
        public string DESC_J { get; set; }
        public string DESC_E { get; set; }
        public string Pack_Type_Description { get; set; }
        public string Segment_Description { get; set; }
        public string Subsegment_1_Description { get; set; }
        public string Subsegment_2_Description { get; set; }
        public string Subsegment_3_Description { get; set; }
        public string Import_Local_Description { get; set; }
        public string Brand { get; set; }
        public string Thai_Tourist_Description { get; set; }
    }
}