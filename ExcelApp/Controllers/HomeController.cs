using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using ExcelApp.Models;
using System.IO;

namespace ExcelApp.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        
        public ActionResult Index()
        {
            return View();

        }
        private List<Store> allStore()
        {

            List<Store> stores = new List<Store>();
            Store s0 = new Store()
            {
                STOREID = 004,
                STORE_NAME = "TMR BANGRAK _004",
                STORE_ATTRIBUTE_TYPE_CODE = 005,
                STORE_ATTRIBUTE_TYPE_NAME = "Format Cluster MKT",
                STORE_ATTRIBUTE_CODE = 1,
                STORE_ATTRIBUTE_NAME = "BKK Inter"
            };
            Store s1 = new Store()
            {
                STOREID = 005,
                STORE_NAME = "TMR SUKHUMIVT_005",
                STORE_ATTRIBUTE_TYPE_CODE = 005,
                STORE_ATTRIBUTE_TYPE_NAME = "Format Cluster MKT",
                STORE_ATTRIBUTE_CODE = 2,
                STORE_ATTRIBUTE_NAME = " BKK Premium"
            };
            Store s2 = new Store()
            {
                STOREID = 007,
                STORE_NAME = "TMR FASHION_007",
                STORE_ATTRIBUTE_TYPE_CODE = 005,
                STORE_ATTRIBUTE_TYPE_NAME = "Format Cluster MKT",
                STORE_ATTRIBUTE_CODE = 1,
                STORE_ATTRIBUTE_NAME = "BKK Inter"
            };
            Store s3 = new Store()
            {
                STOREID = 008,
                STORE_NAME = "TMR RANGSIT_008",
                STORE_ATTRIBUTE_TYPE_CODE = 005,
                STORE_ATTRIBUTE_TYPE_NAME = "Format Cluster MKT",
                STORE_ATTRIBUTE_CODE = 2,
                STORE_ATTRIBUTE_NAME = "BKK Premium"
            };
            Store s4 = new Store()
            {
                STOREID = 009,
                STORE_NAME = "TMR SRINAKARIN_009",
                STORE_ATTRIBUTE_TYPE_CODE = 005,
                STORE_ATTRIBUTE_TYPE_NAME = "Format Cluster MKT",
                STORE_ATTRIBUTE_CODE = 1,
                STORE_ATTRIBUTE_NAME = "BKK Inter"
            };
            Store s5 = new Store()
            {
                STOREID = 011,
                STORE_NAME = "TMS SILOM_011",
                STORE_ATTRIBUTE_TYPE_CODE = 005,
                STORE_ATTRIBUTE_TYPE_NAME = "Format Cluster MKT",
                STORE_ATTRIBUTE_CODE = 1,
                STORE_ATTRIBUTE_NAME = "BKK Inter"
            };
            Store s6 = new Store()
            {
                STOREID = 012,
                STORE_NAME = "TMC LADPRAO_012",
                STORE_ATTRIBUTE_TYPE_CODE = 005,
                STORE_ATTRIBUTE_TYPE_NAME = "Format Cluster MKT",
                STORE_ATTRIBUTE_CODE = 2,
                STORE_ATTRIBUTE_NAME = "BKK Premium"
            };
            Store s7 = new Store()
            {
                STOREID = 017,
                STORE_NAME = "TMC RAMINTRA_017",
                STORE_ATTRIBUTE_TYPE_CODE = 005,
                STORE_ATTRIBUTE_TYPE_NAME = "Format Cluster MKT",
                STORE_ATTRIBUTE_CODE = 1,
                STORE_ATTRIBUTE_NAME = "BKK Inter"
            };
            Store s8 = new Store()
            {
                STOREID = 020,
                STORE_NAME = "TMS RCA_020",
                STORE_ATTRIBUTE_TYPE_CODE = 005,
                STORE_ATTRIBUTE_TYPE_NAME = "Format Cluster MKT",
                STORE_ATTRIBUTE_CODE = 1,
                STORE_ATTRIBUTE_NAME = "BKK Inter"
            };
            Store s9 = new Store()
            {
                STOREID = 029,
                STORE_NAME = "TMR SRIRACHA_029",
                STORE_ATTRIBUTE_TYPE_CODE = 005,
                STORE_ATTRIBUTE_TYPE_NAME = "Format Cluster MKT",
                STORE_ATTRIBUTE_CODE = 1,
                STORE_ATTRIBUTE_NAME = "BKK Inter"
            };
            stores.Add(s0);
            stores.Add(s1);
            stores.Add(s2);
            stores.Add(s3);
            stores.Add(s4);
            stores.Add(s5);
            stores.Add(s6);
            stores.Add(s7);
            stores.Add(s8);
            stores.Add(s9);

            return stores;
        }

        private List<Subclass> allSubclass()
        {

            List<Subclass> subclasss = new List<Subclass>();
            Subclass s0 = new Subclass()
            {
                APO_SUBCLASS = Convert.ToInt64(043300010001),
                APO_SUBCLASS_NAME = "Adult Super Premium (>66)",
                MARKET_DATA_SUBCATEGORY = "Adult",
                MARKET_DATA_CATEGORY = "Disposable Diapers",
            };
            Subclass s1 = new Subclass()
            {
                APO_SUBCLASS = Convert.ToInt64(021300100002),
                APO_SUBCLASS_NAME = "Aerosol Air Fresheners",
                MARKET_DATA_SUBCATEGORY = "Spray",
                MARKET_DATA_CATEGORY = "Air Care",
            };
            Subclass s2 = new Subclass()
            {
                APO_SUBCLASS = Convert.ToInt64(042700010004),
                APO_SUBCLASS_NAME = "After-Shave Lotion",
                MARKET_DATA_SUBCATEGORY = "After shaving",
                MARKET_DATA_CATEGORY = "Shaving Preparation",
            };
            Subclass s3 = new Subclass()
            {
                APO_SUBCLASS = Convert.ToInt64(042300020016),
                APO_SUBCLASS_NAME = "Dermo Facial Acne Medication",
                MARKET_DATA_SUBCATEGORY = "Anti Acne",
                MARKET_DATA_CATEGORY = "Anti Acne",
            };
            Subclass s4 = new Subclass()
            {
                APO_SUBCLASS = Convert.ToInt64(044100010004),
                APO_SUBCLASS_NAME = "Natural/Herbal Shampoos",
                MARKET_DATA_SUBCATEGORY = "Anti Ddf.",
                MARKET_DATA_CATEGORY = "Shampoo",
            };
            Subclass s5 = new Subclass()
            {
                APO_SUBCLASS = Convert.ToInt64(042300020001),
                APO_SUBCLASS_NAME = "Mass Anti-aging",
                MARKET_DATA_SUBCATEGORY = "Anti-aging",
                MARKET_DATA_CATEGORY = "Moisturizer for Face",
            };
            Subclass s6 = new Subclass()
            {
                APO_SUBCLASS = Convert.ToInt64(021300100004),
                APO_SUBCLASS_NAME = "Car Air Fresheners",
                MARKET_DATA_SUBCATEGORY = "Automobile",
                MARKET_DATA_CATEGORY = "Air Care",
            };
            Subclass s7 = new Subclass()
            {
                APO_SUBCLASS = Convert.ToInt64(042100010005),
                APO_SUBCLASS_NAME = "Baby Bar Soap",
                MARKET_DATA_SUBCATEGORY = "Baby",
                MARKET_DATA_CATEGORY = "Toilet Soap",
            };
            Subclass s8 = new Subclass()
            {
                APO_SUBCLASS = Convert.ToInt64(042100020005),
                APO_SUBCLASS_NAME = "Baby Liquid Soap",
                MARKET_DATA_SUBCATEGORY = "Baby",
                MARKET_DATA_CATEGORY = "Liquid Soap",
            };
            Subclass s9 = new Subclass()
            {
                APO_SUBCLASS = Convert.ToInt64(042500030001),
                APO_SUBCLASS_NAME = "Baby Powder",
                MARKET_DATA_SUBCATEGORY = "Baby",
                MARKET_DATA_CATEGORY = "Talcum Powder",
            };


            subclasss.Add(s0);
            subclasss.Add(s1);
            subclasss.Add(s2);
            subclasss.Add(s3);
            subclasss.Add(s4);
            subclasss.Add(s5);
            subclasss.Add(s6);
            subclasss.Add(s7);
            subclasss.Add(s8);
            subclasss.Add(s9);

            return subclasss;
        }

        private List<Biscuits> allBiscuit()
        {

            List<Biscuits> biscuitss = new List<Biscuits>();
            Biscuits b0 = new Biscuits()
            {
                APO_CATEGORY = "Biscuits",
                ID = 4893049120046,
                NAME = "ริทซ์แซนวิชชีส 27ก",
                DESC_J = "RITZ",
                DESC_E = "Discontinue_Y",
                Pack_Type_Description = "Single",
                Segment_Description = "Sweet",
                Subsegment_1_Description = "Sweet Biscuit",
                Subsegment_2_Description = "Sandwich",
                Subsegment_3_Description = "Normal",
                Import_Local_Description = "Imported",
                Brand = "RITZ",
                Thai_Tourist_Description = "NO",
            };
            Biscuits b1 = new Biscuits()
            {
                APO_CATEGORY = "Biscuits",
                ID = 7622300438661,
                NAME = "ริทซ์แครกเกอร์รสมะนาว 118ก",
                DESC_J = "RITZ",
                DESC_E = "Active",
                Pack_Type_Description = "Single",
                Segment_Description = "Sweet",
                Subsegment_1_Description = "Sweet Biscuit",
                Subsegment_2_Description = "Flavored Biscuit",
                Subsegment_3_Description = "Normal",
                Import_Local_Description = "Imported",
                Brand = "RITZ",
                Thai_Tourist_Description = "NO",
            };
            Biscuits b2 = new Biscuits()
            {
                APO_CATEGORY = "Biscuits",
                ID = 0044000012526,
                NAME = "ริทซ์ขนมปังอบกรอบพร้อมชีส 162ก",
                DESC_J = "RITZ",
                DESC_E = "Active",
                Pack_Type_Description = "Single",
                Segment_Description = "Salty",
                Subsegment_1_Description = "Savory Cracker",
                Subsegment_2_Description = "Savory Cracker",
                Subsegment_3_Description = "Normal",
                Import_Local_Description = "Imported",
                Brand = "RITZ",
                Thai_Tourist_Description = "NO",
            };
            Biscuits b3 = new Biscuits()
            {
                APO_CATEGORY = "Biscuits",
                ID = 0044000035457,
                NAME = "นาบิสโก้ริชส์บิทส์ชีส 249ก",
                DESC_J = "RITZ",
                DESC_E = "Active",
                Pack_Type_Description = "Single",
                Segment_Description = "Salty",
                Subsegment_1_Description = "Savory Cracker",
                Subsegment_2_Description = "Savory Cracker",
                Subsegment_3_Description = "Normal",
                Import_Local_Description = "Imported",
                Brand = "RITZ",
                Thai_Tourist_Description = "NO",
            };
            Biscuits b4 = new Biscuits()
            {
                APO_CATEGORY = "Biscuits",
                ID = 8992760211036,
                NAME = "ริทซ์แครกเกอร์ 300ก",
                DESC_J = "RITZ",
                DESC_E = "Active",
                Pack_Type_Description = "Single",
                Segment_Description = "Salty",
                Subsegment_1_Description = "Savory Cracker",
                Subsegment_2_Description = "Savory Cracker",
                Subsegment_3_Description = "Normal",
                Import_Local_Description = "Imported",
                Brand = "RITZ",
                Thai_Tourist_Description = "NO",
            };
            Biscuits b5 = new Biscuits()
            {
                APO_CATEGORY = "Biscuits",
                ID = 4893049120084,
                NAME = "ริทซ์แซนวิชชีส 27X12",
                DESC_J = "RITZ",
                DESC_E = "Active",
                Pack_Type_Description = "Bulk",
                Segment_Description = "Salty",
                Subsegment_1_Description = "Savory Cracker",
                Subsegment_2_Description = "Savory Cracker",
                Subsegment_3_Description = "Normal",
                Import_Local_Description = "Imported",
                Brand = "RITZ",
                Thai_Tourist_Description = "NO",
            };
            Biscuits b6 = new Biscuits()
            {
                APO_CATEGORY = "Biscuits",
                ID = 4893049120008,
                NAME = "ริทซ์ชีส 118ก",
                DESC_J = "RITZ",
                DESC_E = "Active",
                Pack_Type_Description = "Single",
                Segment_Description = "Salty",
                Subsegment_1_Description = "Savory Cracker",
                Subsegment_2_Description = "Savory Cracker",
                Subsegment_3_Description = "Normal",
                Import_Local_Description = "Imported",
                Brand = "RITZ",
                Thai_Tourist_Description = "NO",
            };
            Biscuits b7 = new Biscuits()
            {
                APO_CATEGORY = "Biscuits",
                ID = 8992760211029,
                NAME = "ริทซ์แครกเกอร์ 100ก",
                DESC_J = "RITZ",
                DESC_E = "Active",
                Pack_Type_Description = "Single",
                Segment_Description = "Salty",
                Subsegment_1_Description = "Savory Cracker",
                Subsegment_2_Description = "Savory Cracker",
                Subsegment_3_Description = "Normal",
                Import_Local_Description = "Imported",
                Brand = "RITZ",
                Thai_Tourist_Description = "NO",
            };


            biscuitss.Add(b0);
            biscuitss.Add(b1);
            biscuitss.Add(b2);
            biscuitss.Add(b3);
            biscuitss.Add(b4);
            biscuitss.Add(b5);
            biscuitss.Add(b6);
            biscuitss.Add(b7);

            return biscuitss;
        }

        private List<Product> allProduct()
        {

            List<Product> product = new List<Product>();
            Product p0 = new Product()
            {
                
                    APO_SUBCLASS = 012900010004,
                    SUBCLASS = 4,
                    SUB_NAME = "OTOP",
                    addIn = "OTOP",
                    CLASS = 1,
                    CLASS_NAME = "Seasonal",
                    DEPT = 129,
                    DEPT_NAME = "Seasonal",
                    GROUP_NO = 12,
                    GROUP_NAME = "Packaged",
                    DIVISION = 1,
                    DIV_NAME = "Food"
            };
            Product p1 = new Product()
            {

                APO_SUBCLASS = 11100010002,
                SUBCLASS = 2,
                SUB_NAME = "Import Beer",
                addIn = "Import Beer",
                CLASS = 1,
                CLASS_NAME = "Beer",
                DEPT = 111,
                DEPT_NAME = "Beverages",
                GROUP_NO = 11,
                GROUP_NAME = "Snack & Beverage",
                DIVISION = 1,
                DIV_NAME = "Food"
            };
            Product p2 = new Product()
            {

                APO_SUBCLASS = 011100030001,
                SUBCLASS = 1,
                SUB_NAME = "Cola Soda",
                addIn = "Cola Soda",
                CLASS = 3,
                CLASS_NAME = "Soft Drinks",
                DEPT = 111,
                DEPT_NAME = "Beverages",
                GROUP_NO = 11,
                GROUP_NAME = "Snack & Beverage",
                DIVISION = 1,
                DIV_NAME = "Food"
            };
            Product p3 = new Product()
            {

                APO_SUBCLASS = 011100030004,
                SUBCLASS = 4,
                SUB_NAME = "Mixers",
                addIn = "Mixers",
                CLASS = 3,
                CLASS_NAME = "Soft Drinks",
                DEPT = 111,
                DEPT_NAME = "Beverages",
                GROUP_NO = 11,
                GROUP_NAME = "Snack & Beverage",
                DIVISION = 1,
                DIV_NAME = "Food"
            };
            Product p4 = new Product()
            {

                APO_SUBCLASS = 011100040001,
                SUBCLASS = 1,
                SUB_NAME = "Energy Drinks",
                addIn = "Energy Drinks",
                CLASS = 4,
                CLASS_NAME = "Energy/Sports Drinks",
                DEPT = 111,
                DEPT_NAME = "Beverages",
                GROUP_NO = 11,
                GROUP_NAME = "Snack & Beverage",
                DIVISION = 1,
                DIV_NAME = "Food"
            };
            Product p5 = new Product()
            {

                APO_SUBCLASS = 011100030003,
                SUBCLASS = 3,
                SUB_NAME = "Fruit Flavored Soda",
                addIn = "Fruit Flavored Soda",
                CLASS = 3,
                CLASS_NAME = "Soft Drinks",
                DEPT = 111,
                DEPT_NAME = "Beverages",
                GROUP_NO = 11,
                GROUP_NAME = "Snack & Beverage",
                DIVISION = 1,
                DIV_NAME = "Food"
            };
            Product p6 = new Product()
            {

                APO_SUBCLASS = 011100030002,
                SUBCLASS = 2,
                SUB_NAME = "Clear Soda",
                addIn = "Clear Soda",
                CLASS = 3,
                CLASS_NAME = "Soft Drinks",
                DEPT = 111,
                DEPT_NAME = "Beverages",
                GROUP_NO = 11,
                GROUP_NAME = "Snack & Beverage",
                DIVISION = 1,
                DIV_NAME = "Food"
            };
            Product p7 = new Product()
            {

                APO_SUBCLASS = 011100040002,
                SUBCLASS = 2,
                SUB_NAME = "Isotonics",
                addIn = "Isotonics",
                CLASS = 4,
                CLASS_NAME = "Energy/Sports Drinks",
                DEPT = 111,
                DEPT_NAME = "Beverages",
                GROUP_NO = 11,
                GROUP_NAME = "Snack & Beverage",
                DIVISION = 1,
                DIV_NAME = "Food"
            };
            Product p8 = new Product()
            {

                APO_SUBCLASS = 011100060001,
                SUBCLASS = 1,
                SUB_NAME = "Still Mineral Water",
                addIn = "Still Mineral Water",
                CLASS = 6,
                CLASS_NAME = "Mineral Water",
                DEPT = 111,
                DEPT_NAME = "Beverages",
                GROUP_NO = 11,
                GROUP_NAME = "Snack & Beverage",
                DIVISION = 1,
                DIV_NAME = "Food"
            };
            Product p9 = new Product()
            {

                APO_SUBCLASS = 011100080007,
                SUBCLASS = 7,
                SUB_NAME = "Economy Tomato Juice",
                addIn = "Economy Tomato Juice",
                CLASS = 8,
                CLASS_NAME = "Economy Juice",
                DEPT = 111,
                DEPT_NAME = "Beverages",
                GROUP_NO = 11,
                GROUP_NAME = "Snack & Beverage",
                DIVISION = 1,
                DIV_NAME = "Food"
            };

            product.Add(p0);
            product.Add(p1);
            product.Add(p2);
            product.Add(p3);
            product.Add(p4);
            product.Add(p5);
            product.Add(p6);
            product.Add(p7);
            product.Add(p8);
            product.Add(p9);

            return product;
        }
        
        //Market data
        public ActionResult MarketData()
        {
            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("Store Detail");
                ExcelWorksheet ws = package.Workbook.Worksheets[1];
                ws.Name = "Store Detail"; //Setting Sheet's name
                ws.Cells.Style.Font.Size = 9; //Default font size for whole sheet
                ws.Cells.Style.Font.Name = "Arial"; //Default Font name for whole sheet 

                ws.Cells["B1:C1"].Merge = true;
                ws.Cells[1, 1].Style.Font.Bold = true;

                var cell3 = ws.Cells[3, 1];
                cell3.Style.Font.Bold = true;
                cell3.Style.Font.Size = 12;

                ws.Cells[1, 1].Value = "Target Table Name:";
                ws.Cells[1, 2].Value = "CRC_MARKET_APO_HIER_XREF_LD";

                ws.Cells[3, 1].Value = "Data Values";
                ws.Cells[3, 3].Value = "this shuld be unique";
                ws.Cells[3, 5].Value = "Should not over 7";

                int headerRow = 4;
                ws.Cells[headerRow, 1].Value = "STORE";
                ws.Cells[headerRow, 2].Value = "STORE_NAME";
                ws.Cells[headerRow, 3].Value = "STORE_ATTRIBUTE_TYPE_CODE";
                ws.Cells[headerRow, 4].Value = "STORE_ATTRIBUTE_TYPE_NAME";
                ws.Cells[headerRow, 5].Value = "STORE_ATTRIBUTE_CODE";
                ws.Cells[headerRow, 6].Value = "STORE_ATTRIBUTE_NAME";

                for (int c = 1; c <= 6; c++)
                {
                    if (c % 2 == 1)
                    {
                        var columnWS = ws.Cells[headerRow, c];
                        columnWS.Style.Font.Bold = true;
                        columnWS.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        columnWS.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    }
                    else
                    {
                        var columnWS = ws.Cells[headerRow, c];
                        columnWS.Style.Font.Bold = true;
                        columnWS.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        columnWS.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                    }

                    using (var range = ws.Cells[headerRow, 1, headerRow, 6])
                    {
                        // Assign borders
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    }
                }

                var s = allStore();
                int i = 0;
                int row = 4;
                foreach (var item in s)
                {
                    i++;
                    row++;
                    var dataColumn1 = ws.Cells[row, 1];
                    dataColumn1.Value = string.Format("{0:000}", item.STOREID);
                    dataColumn1.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    ws.Cells[row, 2].Value = item.STORE_NAME;

                    var dataColumn3 = ws.Cells[row, 3];
                    dataColumn3.Value = string.Format("{0:000}", item.STORE_ATTRIBUTE_TYPE_CODE);
                    dataColumn3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    ws.Cells[row, 4].Value = item.STORE_ATTRIBUTE_TYPE_NAME;
                    ws.Cells[row, 5].Value = item.STORE_ATTRIBUTE_CODE;
                    ws.Cells[row, 6].Value = item.STORE_ATTRIBUTE_NAME;
                }

                ws.Cells.AutoFitColumns();

                /////////////////////////////////////////////////////////////////////////////////////
                package.Workbook.Worksheets.Add("Market Data");
                ExcelWorksheet ws2 = package.Workbook.Worksheets[2];
                ws2.Name = "Market Data"; //Setting Sheet's name
                ws2.Cells.Style.Font.Size = 9; //Default font size for whole sheet
                ws2.Cells.Style.Font.Name = "Arial"; //Default Font name for whole sheet 

                ws2.Cells["B1:C1"].Merge = true;
                ws2.Cells[1, 1].Style.Font.Bold = true;

                var ws2Cell3 = ws2.Cells[3, 1];
                ws2Cell3.Style.Font.Bold = true;
                ws2Cell3.Style.Font.Size = 12;

                ws2.Cells[1, 1].Value = "Target Table Name:";
                ws2.Cells[1, 2].Value = "CRC_MARKET_APO_HIER_XREF_LD";

                ws2.Cells[3, 1].Value = "Data Values";

                int ws2headerRow = 4;
                ws2.Cells[ws2headerRow, 1].Value = "APO_SUBCLASS";
                ws2.Cells[ws2headerRow, 2].Value = "APO_SUBCLASS_NAME";
                ws2.Cells[ws2headerRow, 3].Value = "MARKET_DATA_SUBCATEGORY";
                ws2.Cells[ws2headerRow, 4].Value = "MARKET_DATA_CATEGORY";

                using (var range = ws2.Cells[ws2headerRow, 1, ws2headerRow, 4])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    // Assign borders
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }

                var a = allSubclass();
                int j = 0;
                int ws2Row = 4;
                foreach (var item in a)
                {
                    j++;
                    ws2Row++;
                    ws2.Cells[ws2Row, 1].Value = string.Format("{0:000000000000}", item.APO_SUBCLASS) ;
                    ws2.Cells[ws2Row, 2].Value = item.APO_SUBCLASS_NAME;
                    ws2.Cells[ws2Row, 3].Value = item.MARKET_DATA_SUBCATEGORY;
                    ws2.Cells[ws2Row, 4].Value = item.MARKET_DATA_CATEGORY;
                }

                ws2.Cells.AutoFitColumns();

                var memoryStream = package.GetAsByteArray();
                var fileName = "attributes maintainance.xlsx";
                return base.File(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }

        //public ActionResult StoreDetail()
        //{
        //    using (var package = new ExcelPackage())
        //    {
        //        package.Workbook.Worksheets.Add("Sheet1");
        //        ExcelWorksheet ws = package.Workbook.Worksheets[1];
        //        ws.Name = "Sheet1"; //Setting Sheet's name
        //        ws.Cells.Style.Font.Size = 9; //Default font size for whole sheet
        //        ws.Cells.Style.Font.Name = "Arial"; //Default Font name for whole sheet 

        //        ws.Cells["B1:C1"].Merge = true;
        //        ws.Cells[1, 1].Style.Font.Bold = true;

        //        var cell3 = ws.Cells[3, 1];
        //        cell3.Style.Font.Bold = true;
        //        cell3.Style.Font.Size = 12;

        //        ws.Cells[1, 1].Value = "Target Table Name:";
        //        ws.Cells[1, 2].Value = "CRC_MARKET_APO_HIER_XREF_LD";
               
        //        ws.Cells[3, 1].Value = "Data Values";
        //        ws.Cells[3, 3].Value = "this shuld be unique";
        //        ws.Cells[3, 5].Value = "Should not over 7";

        //        int headerRow = 4;
        //        ws.Cells[headerRow, 1].Value = "STORE";
        //        ws.Cells[headerRow, 2].Value = "STORE_NAME";
        //        ws.Cells[headerRow, 3].Value = "STORE_ATTRIBUTE_TYPE_CODE";
        //        ws.Cells[headerRow, 4].Value = "STORE_ATTRIBUTE_TYPE_NAME";
        //        ws.Cells[headerRow, 5].Value = "STORE_ATTRIBUTE_CODE";
        //        ws.Cells[headerRow, 6].Value = "STORE_ATTRIBUTE_NAME";

        //        for (int c = 1; c <= 6; c++)
        //        {
        //            if (c % 2 == 1)
        //            {
        //                var columnWS = ws.Cells[headerRow, c];
        //                columnWS.Style.Font.Bold = true;
        //                columnWS.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //                columnWS.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
        //            }
        //            else
        //            {
        //                var columnWS = ws.Cells[headerRow, c];
        //                columnWS.Style.Font.Bold = true;
        //                columnWS.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //                columnWS.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
        //            }

        //            using (var range = ws.Cells[headerRow, 1, headerRow, 6])
        //            {
        //                // Assign borders
        //                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
        //                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
        //                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        //                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        //            }
        //        }

        //        var s = allStore();
        //        int i = 0;
        //        int row = 4;
        //        foreach (var item in s)
        //        {
        //            i++;
        //            row++;
        //            var dataColumn1 = ws.Cells[row, 1];
        //            dataColumn1.Value = string.Format("{0:000}", item.STOREID);
        //            dataColumn1.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right; 

        //            ws.Cells[row, 2].Value = item.STORE_NAME;

        //            var dataColumn3 = ws.Cells[row, 3];
        //            dataColumn3.Value = string.Format("{0:000}", item.STORE_ATTRIBUTE_TYPE_CODE);
        //            dataColumn3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

        //            ws.Cells[row, 4].Value = item.STORE_ATTRIBUTE_TYPE_NAME;
        //            ws.Cells[row, 5].Value = item.STORE_ATTRIBUTE_CODE;
        //            ws.Cells[row, 6].Value = item.STORE_ATTRIBUTE_NAME;
        //        }

        //        ws.Cells.AutoFitColumns();
        //        var memoryStream = package.GetAsByteArray();
        //        var fileName = "Store Detail.xlsx";
        //        return base.File(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        //    }
        //}

        public ActionResult FinalCut()
        {

            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("Final Cut");
                ExcelWorksheet ws = package.Workbook.Worksheets[1];
                ws.Name = "Final Cut"; //Setting Sheet's name
                ws.Cells.Style.Font.Size = 9; //Default font size for whole sheet
                ws.Cells.Style.Font.Name = "Arial"; //Default Font name for whole sheet 
                ws.View.FreezePanes(5,4);

                var cell1 = ws.Cells["A1:C1"];
                cell1.Merge = true;
                cell1.Style.Font.Size = 12;
                cell1.Style.Font.Bold = true;

                ws.Cells[1, 1].Value = "BISCUITS (with Active & Discon with conditon)";
                ws.Cells[2, 1].Value = "As of Nov4";

                int headerRow = 4;
                ws.Cells[headerRow, 1].Value = "APO CATEGORY";
                ws.Cells[headerRow, 2].Value = "ID";
                ws.Cells[headerRow, 3].Value = "NAME";
                ws.Cells[headerRow, 4].Value = "DESC_J";
                ws.Cells[headerRow, 5].Value = "DESC_J";
                ws.Cells[headerRow, 6].Value = "Pack Type Description";
                ws.Cells[headerRow, 7].Value = "Segment Description";
                ws.Cells[headerRow, 8].Value = "Subsegment 1 Description";
                ws.Cells[headerRow, 9].Value = "Subsegment 2 Description";
                ws.Cells[headerRow, 10].Value = "Subsegment 3 Description";
                ws.Cells[headerRow, 11].Value = "Import Local Description";
                ws.Cells[headerRow, 12].Value = "Brand";
                ws.Cells[headerRow, 13].Value = "Thai Tourist Description";

                using (var range = ws.Cells[headerRow, 1, headerRow, 5])
                {
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.CadetBlue);
                    // Assign borders
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }

                var data_export = allBiscuit();
                int i = 0;
                int row = 4;
                foreach (var item in data_export)
                {
                    i++;
                    row++;
                    ws.Cells[row, 1].Value = item.APO_CATEGORY;
                    ws.Cells[row, 2].Value = string.Format("{0:0000000000000}", item.ID);
                    ws.Cells[row, 3].Value = item.NAME;
                    ws.Cells[row, 4].Value = item.DESC_J;
                    ws.Cells[row, 5].Value = item.DESC_E;
                    ws.Cells[row, 6].Value = item.Pack_Type_Description;
                    ws.Cells[row, 7].Value = item.Segment_Description;
                    ws.Cells[row, 8].Value = item.Subsegment_1_Description;
                    ws.Cells[row, 9].Value = item.Subsegment_2_Description;
                    ws.Cells[row, 10].Value = item.Subsegment_3_Description;
                    ws.Cells[row, 11].Value = item.Import_Local_Description;
                    ws.Cells[row, 12].Value = item.Brand;
                    var dataColumn13 = ws.Cells[row, 13];
                    dataColumn13.Value = item.Thai_Tourist_Description;
                    dataColumn13.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    dataColumn13.Style.Font.Color.SetColor(Color.Silver);

                }

                ws.Cells.AutoFitColumns();
                var memoryStream = package.GetAsByteArray();
                var fileName = "BISCUITS.xlsx";
                return base.File(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }

        public ActionResult GridResults()
        {
            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("Grid Results");
                ExcelWorksheet ws = package.Workbook.Worksheets[1];
                ws.Name = "Grid Results"; //Setting Sheet's name
                ws.Cells.Style.Font.Size = 10; //Default font size for whole sheet
                ws.Cells.Style.Font.Name = "Arial"; //Default Font name for whole sheet 

                ws.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

                using (var range = ws.Cells[1, 2, 1, 12])
                {
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.Silver);
                }

                int cell1 = 1;
                ws.Cells[cell1, 1].Value = "APO_SUBCLASS";
                ws.Cells[cell1, 2].Value = "SUBCLASS";
                ws.Cells[cell1, 3].Value = "SUB_NAME";
                ws.Cells[cell1, 4].Value = "add in";
                ws.Cells[cell1, 5].Value = "CLASS";
                ws.Cells[cell1, 6].Value = "CLASS_NAME";
                ws.Cells[cell1, 7].Value = "DEPT";
                ws.Cells[cell1, 8].Value = "DEPT_NAME";
                ws.Cells[cell1, 9].Value = "GROUP_NO";
                ws.Cells[cell1, 10].Value = "GROUP_NAME";
                ws.Cells[cell1, 11].Value = "DIVISION";
                ws.Cells[cell1, 12].Value = "DIV_NAME";

                var data_export = allProduct();

                int i = 0;
                int row = 1;
                foreach (var item in data_export)
                {
                    i++;
                    row++;
                    if (row % 2 == 1)
                    {
                        var columnColor1 = ws.Cells[row, 1, row, 3];
                        columnColor1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        columnColor1.Style.Fill.BackgroundColor.SetColor(Color.PowderBlue);

                        var columnColor2 = ws.Cells[row, 5, row, 12];
                        columnColor2.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        columnColor2.Style.Fill.BackgroundColor.SetColor(Color.PowderBlue);

                        ws.Cells[row, 1].Value = string.Format("{0:000000000000}", item.APO_SUBCLASS);
                        ws.Cells[row, 2].Value = item.SUBCLASS;
                        ws.Cells[row, 3].Value = item.SUB_NAME;
                        ws.Cells[row, 4].Value = item.addIn;
                        ws.Cells[row, 5].Value = item.CLASS;
                        ws.Cells[row, 6].Value = item.CLASS;
                        ws.Cells[row, 7].Value = item.DEPT;
                        ws.Cells[row, 8].Value = item.DEPT_NAME;
                        ws.Cells[row, 9].Value = item.GROUP_NO;
                        ws.Cells[row, 10].Value = item.GROUP_NAME;
                        ws.Cells[row, 11].Value = item.DIVISION;
                        ws.Cells[row, 12].Value = item.DIV_NAME;
                    }
                    else
                    {
                        var columnColor1 = ws.Cells[row, 1, row, 3];
                        columnColor1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        columnColor1.Style.Fill.BackgroundColor.SetColor(Color.White);

                        var columnColor2 = ws.Cells[row, 5, row, 12];
                        columnColor2.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        columnColor2.Style.Fill.BackgroundColor.SetColor(Color.White);

                        ws.Cells[row, 1].Value = string.Format("{0:000000000000}", item.APO_SUBCLASS);
                        ws.Cells[row, 2].Value = item.SUBCLASS;
                        ws.Cells[row, 3].Value = item.SUB_NAME;
                        ws.Cells[row, 4].Value = item.addIn;
                        ws.Cells[row, 5].Value = item.CLASS;
                        ws.Cells[row, 6].Value = item.CLASS;
                        ws.Cells[row, 7].Value = item.DEPT;
                        ws.Cells[row, 8].Value = item.DEPT_NAME;
                        ws.Cells[row, 9].Value = item.GROUP_NO;
                        ws.Cells[row, 10].Value = item.GROUP_NAME;
                        ws.Cells[row, 11].Value = item.DIVISION;
                        ws.Cells[row, 12].Value = item.DIV_NAME;
                    }

                    using (var range = ws.Cells[row, 1, row, 12])
                    {
                        // Assign borders
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        range.Style.Border.Top.Color.SetColor(Color.PowderBlue);
                        range.Style.Border.Left.Color.SetColor(Color.PowderBlue);
                        range.Style.Border.Right.Color.SetColor(Color.PowderBlue);
                        range.Style.Border.Bottom.Color.SetColor(Color.PowderBlue);
                    }
                }

                ws.Cells.AutoFitColumns();
                var memoryStream = package.GetAsByteArray();
                var fileName = "rms_hierachy_14mar2017.xlsx";
                return base.File(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }

        public ActionResult ImportExcel(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
            {
                if(Path.GetExtension(file.FileName) == ".xlsx")
                {
                    string filename = file.FileName;
                    string path = Server.MapPath("~/Upload/");
                    file.SaveAs(path + filename);
                    

                    List<Subclass> s = new List<Subclass>();
                    using (var package = new ExcelPackage(file.InputStream))
                    {
                        ExcelWorksheet ws = package.Workbook.Worksheets.First();
                        int row = 5;
                        for (row = 5; row <= ws.Dimension.End.Row; row++)
                        {
                            Subclass sub = new Subclass()
                            {
                                APO_SUBCLASS = Convert.ToInt64(ws.Cells[row, 1].Value),
                                APO_SUBCLASS_NAME = Convert.ToString(ws.Cells[row, 2].Value),
                                MARKET_DATA_SUBCATEGORY = Convert.ToString(ws.Cells[row, 3].Value),
                                MARKET_DATA_CATEGORY = Convert.ToString(ws.Cells[row, 4].Value)
                            };
                            s.Add(sub);
                        }
                    }
                    ViewBag.Error = null;
                    return View("Index");
                }
                else
                {
                    ViewBag.Error = "Not excel file!!";
                    return View("Index");
                }
                
            }
            else
            {
                ViewBag.Error = "select file !!";
                return View("Index");
            }
            
        }

    }
}