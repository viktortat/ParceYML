using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Diagnostics;
using OfficeOpenXml.Style;

namespace ParceYmlApp
{
    public partial class frmMain : Form
    {


        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            Program.connectionStr =
                Program.connectionStr = ConfigurationManager.ConnectionStrings["TbnProd.Local"].ConnectionString;
            Program.PathExcelFileImport = AppDomain.CurrentDomain.BaseDirectory + @"testXml\soap.xml";
            txbPathSelector.Text = Program.PathExcelFileImport;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            XmlDocument doc = new XmlDocument();
            //doc.Load(@"d:\5552\yml.xml");
            //doc.Load(@"d:\5552\ymlTCN.xml");
            doc.Load(@"d:\5552\soap.xml");
            XmlNodeList nodeList;
            XmlElement root = doc.DocumentElement;
            nodeList = root.SelectNodes("/yml_catalog/shop/offers/offer");

            string fileName = "test12.xlsx";
            string outputDir = @"d:\5552\";

            var file = new FileInfo(outputDir + fileName);
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet ws = package.Workbook.Worksheets.Add("Распарсен");
                int cRow = 2;
                //ExcelWorksheet wsPlus = package.Workbook.Worksheets[1]; //package.Workbook.Worksheets.Add("Правильные");

                ws.Cells[1, 1].Value = "Артикул";
                ws.Cells[1, 2].Value = "url";
                ws.Cells[1, 3].Value = "price";
                ws.Cells[1, 4].Value = "currencyId";
                ws.Cells[1, 5].Value = "categoryId";
                ws.Cells[1, 6].Value = "picture";
                ws.Cells[1, 7].Value = "name";
                ws.Cells[1, 8].Value = "vendor";
                ws.Cells[1, 9].Value = "description";
                ws.Cells[1, 10].Value = "country_of_origin";
                ws.Cells[1, 11].Value = "barcode";

                ws.Cells[1, 12].Value = "производитель";
                ws.Cells[1, 13].Value = "регион";
                ws.Cells[1, 14].Value = "страна";
                ws.Cells[1, 15].Value = "тип";
                ws.Cells[1, 16].Value = "цвет";
                ws.Cells[1, 17].Value = "сладость";
                ws.Cells[1, 18].Value = "алкоголь";
                ws.Cells[1, 19].Value = "сахар";
                ws.Cells[1, 20].Value = "Сорт винограда";
                ws.Cells[1, 21].Value = "год";
                ws.Cells[1, 22].Value = "объём";
                ws.Cells[1, 23].Value = "Масса / Объем";
                ws.Cells[1, 24].Value = "Вид";
                ws.Cells[1, 25].Value = "Страна";
                ws.Cells[1, 26].Value = "классификация";



                foreach (XmlNode isbn in nodeList)
                {
                    ws.Cells[cRow, 1].Value = isbn.Attributes["id"].InnerText;
                    ws.Cells[cRow, 2].Value = isbn["url"].InnerText;
                    ws.Cells[cRow, 3].Value = isbn["price"].InnerText;
                    //ws.Cells[cRow, 4].Value = isbn["currencyId"].InnerText;
                    ws.Cells[cRow, 5].Value = isbn["categoryId"].InnerText;
                    ws.Cells[cRow, 6].Value = isbn["picture"].InnerText;
                    ws.Cells[cRow, 7].Value = isbn["name"].InnerText;
                    if (isbn["vendor"] != null)
                    {
                        ws.Cells[cRow, 8].Value = isbn["vendor"].InnerText;
                    }
                    ws.Cells[cRow, 9].Value = isbn["description"].InnerText;

                    if (isbn["country_of_origin"] != null)
                    {
                        ws.Cells[cRow, 10].Value = isbn["country_of_origin"].InnerText;
                    }
                    if (isbn["barcode"] != null)
                    {
                        ws.Cells[cRow, 11].Value = isbn["barcode"].InnerText;
                    }

                    int cCol = 26;
                    XmlNodeList nodeParams = isbn.SelectNodes("param");
                    foreach (XmlNode p in nodeParams)
                    {

                        switch (p.Attributes["name"].InnerText)
                        {
                            case "производитель":
                                ws.Cells[cRow, 12].Value = p.InnerText;
                                break;
                            case "Производитель":
                                ws.Cells[cRow, 12].Value = p.InnerText;
                                break;
                            case "регион":
                                ws.Cells[cRow, 13].Value = p.InnerText;
                                break;
                            case "страна":
                                ws.Cells[cRow, 14].Value = p.InnerText;
                                break;
                            case "тип":
                                ws.Cells[cRow, 15].Value = p.InnerText;
                                break;
                            case "цвет":
                                ws.Cells[cRow, 16].Value = p.InnerText;
                                break;
                            case "сладость":
                                ws.Cells[cRow, 17].Value = p.InnerText;
                                break;
                            case "алкоголь":
                                ws.Cells[cRow, 18].Value = p.InnerText;
                                break;
                            case "сахар":
                                ws.Cells[cRow, 19].Value = p.InnerText;
                                break;
                            case "Сорт винограда":
                                ws.Cells[cRow, 20].Value = p.InnerText;
                                break;
                            case "год":
                                ws.Cells[cRow, 21].Value = p.InnerText;
                                break;
                            case "объём":
                                ws.Cells[cRow, 22].Value = p.InnerText;
                                break;
                            case "Масса / Объем":
                                ws.Cells[cRow, 23].Value = p.InnerText;
                                break;
                            case "Вид":
                                ws.Cells[cRow, 24].Value = p.InnerText;
                                break;
                            case "Страна":
                                ws.Cells[cRow, 25].Value = p.InnerText;
                                break;
                            case "классификация":
                                ws.Cells[cRow, 26].Value = p.InnerText;
                                break;
                            default:
                                ws.Cells[cRow, cCol++].Value = p.InnerText;
                                break;
                        }

                    }



                    cRow++;
                    //Console.WriteLine(isbn.Attributes["id"].InnerText);
                    //Console.WriteLine(isbn["url"].InnerText);
                    //Console.WriteLine(isbn["price"].InnerText);
                    //Console.WriteLine("----------------------------------");
                }
                package.SaveAs(file);

                Process.Start(file.FullName);
                //MessageBox.Show("Все!");
            }

        }


        private void btnSelFile_Click(object sender, EventArgs e)
        {
            SelPathExcelFileImport();
        }


        private void SelPathExcelFileImport()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            /*
            openFileDialog1.InitialDirectory = Directory.Exists(Path.GetDirectoryName(PathExcelFile))
               ? Path.GetDirectoryName(PathExcelFile)
               : "c:\\";
             */
            openFileDialog1.Filter = "Все файлы (*.*)|*.*|YML файлы (*.xml)|*.xml";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Program.PathExcelFileImport = openFileDialog1.FileName;
                txbPathSelector.Text = Path.GetFullPath(Program.PathExcelFileImport);

                //lblTopFileName.Text = Path.GetFileName(Program.PathExcelFileImport);
                //ExcelFolderName = Path.GetDirectoryName(Program.PathExcelFileImport);
                //PathExcelFile = txbPathSelector.Text;

                //ReloadPackage();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            XmlDocument doc = new XmlDocument();
            doc.Load(@"d:\5552\yml.xml");
            XmlNodeList nodeList;
            XmlElement root = doc.DocumentElement;
            nodeList = root.SelectNodes("/yml_catalog/shop/offers/offer");

            string fileName = "Pricat.xlsx";
            string fileNameT = "Торговое предложение.xltx";

            string outputDir = @"d:\5552\";

            var file = new FileInfo(outputDir + fileName);
            var newTFile = new FileInfo(Path.GetDirectoryName(Application.ExecutablePath) + @"\Tml\" + fileNameT);

            //Path.GetDirectoryName(Application.ExecutablePath)

            //ExcelPackage package = null;
            //if (File.Exists(newTFile.FullName))
            //{
            //    if (package != null) package.Dispose();
            //    package = new ExcelPackage(newTFile);
            //}
            //else
            //{
            //    package = new ExcelPackage();
            //    Console.WriteLine("Не найден путь к шаблонному Excel");
            //}

            FileInfo empTpl = null;
            var basePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Tmpl", "Торговое предложение.xltx");
            //var locPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", _appContext.UserCulture.TwoLetterISOLanguageName, "EmptyTemplate.xlsx");

            if (File.Exists(basePath))
                empTpl = new FileInfo(basePath);
            else
                empTpl = new FileInfo(basePath);

            using (ExcelPackage package = new ExcelPackage(empTpl, true))
            {
                //Tmpl/Торговое предложение.xltx
                //_excelPackage = new ExcelPackage(newFile);
                ExcelWorksheet ws = package.Workbook.Worksheets.First(); //.Add("Прайс лист");
                //ExcelWorksheet wsPlus = package.Workbook.Worksheets[1]; //package.Workbook.Worksheets.Add("Правильные");

                //ws.Cells[2, 1].Value = "name";
                //ws.Cells[2, 2].Value = "barcode";
                //ws.Cells[2, 3].Value = "Артикул продавца";
                //ws.Cells[2, 4].Value = "Артикул покупателя";
                //ws.Cells[2, 5].Value = "Мин кол-во";
                //ws.Cells[2, 6].Value = "НДС";
                //ws.Cells[2, 7].Value = "price";
                //ws.Cells[2, 8].Value = "categoryId";

                int cRow = 4;
                foreach (XmlNode isbn in nodeList)
                {
                    ws.Cells[cRow, 1].Value = isbn["name"].InnerText;
                    if (isbn["barcode"] != null)
                    {
                        ws.Cells[cRow, 2].Value = isbn["barcode"].InnerText;
                    }
                    ws.Cells[cRow, 3].Value = isbn.Attributes["id"].InnerText;
                    ws.Cells[cRow, 4].Value = "";
                    ws.Cells[cRow, 5].Value = "1";
                    ws.Cells[cRow, 6].Value = "0";
                    ws.Cells[cRow, 7].Value = isbn["price"].InnerText;
                    ws.Cells[cRow, 8].Value = isbn["categoryId"].InnerText;

                    cRow++;
                }
                package.SaveAs(file);

                MessageBox.Show("Все!");
            }
            Process.Start(file.FullName);
        }

        private void btnParce2_Click(object sender, EventArgs e)
        {
            XmlDocument doc = new XmlDocument();
            string FileName = Program.PathExcelFileImport;

            doc.Load(FileName);
            XmlElement root = doc.DocumentElement;
            var nodeList = GetOffer(root);
            var brandArr = GetBrandColl(root);
            var ManufactureArr = GetManufacturerColl(root);
            var CategoriesColl = GetCategoriesColl(root);

            if (chbCopyToDB.Checked) BulkСopyToDB(root);

            string fileName = Path.GetFileNameWithoutExtension(FileName) + ".xlsx";
            string outputDir = Path.GetDirectoryName(FileName);

            Dictionary<string, int> dAttr = new Dictionary<string, int>();

            var file = new FileInfo(outputDir + '\\' + fileName);
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet ws = package.Workbook.Worksheets.Add("Распарсен");
                ExcelWorksheet ws2 = package.Workbook.Worksheets.Add("Категории");
                ExcelWorksheet ws3 = package.Workbook.Worksheets.Add("Фильтры");
                ExcelWorksheet ws4 = package.Workbook.Worksheets.Add("Производители");
                ExcelWorksheet ws5 = package.Workbook.Worksheets.Add("Бренды");

                int cRow = 2;
                int cRowAtr = 2;
                int cRowCat = 2;
                int cRowFact = 2;
                int cRowBrand = 2;

                Color clrHead = Color.LightSkyBlue;

                SetCellHeader(ws4.Cells[1, 1], clrHead, "id");
                SetCellHeader(ws4.Cells[1, 2], clrHead, "Название");

                SetCellHeader(ws5.Cells[1, 1], clrHead, "id");
                SetCellHeader(ws5.Cells[1, 2], clrHead, "Название");

                SetCellHeader(ws3.Cells[1, 1], clrHead, "id");
                SetCellHeader(ws3.Cells[1, 2], clrHead, "Название");
                SetCellHeader(ws3.Cells[1, 3], clrHead, "НазваниеTBN");
                SetCellHeader(ws3.Cells[1, 4], clrHead, "Type");

                SetCellHeader(ws2.Cells[1, 1], clrHead, "id");
                SetCellHeader(ws2.Cells[1, 2], clrHead, "parentId");
                SetCellHeader(ws2.Cells[1, 3], clrHead, "parentName");
                SetCellHeader(ws2.Cells[1, 4], clrHead, "Name");
                SetCellHeader(ws2.Cells[1, 5], clrHead, "НашId");
                SetCellHeader(ws2.Cells[1, 6], clrHead, "НашаКатегория");

                SetCellHeader(ws.Cells[1, 1], clrHead, "available");
                SetCellHeader(ws.Cells[1, 2], clrHead, "id");
                SetCellHeader(ws.Cells[1, 3], clrHead, "name");
                SetCellHeader(ws.Cells[1, 4], clrHead, "url");
                SetCellHeader(ws.Cells[1, 5], clrHead, "price");
                SetCellHeader(ws.Cells[1, 6], clrHead, "currencyId");
                SetCellHeader(ws.Cells[1, 7], clrHead, "categoryId");
                SetCellHeader(ws.Cells[1, 8], clrHead, "categoryName");
                SetCellHeader(ws.Cells[1, 9], clrHead, "delivery");
                SetCellHeader(ws.Cells[1, 10], clrHead, "vendorCode");
                SetCellHeader(ws.Cells[1, 11], clrHead, "vendor");
                SetCellHeader(ws.Cells[1, 12], clrHead, "description");
                SetCellHeader(ws.Cells[1, 13], clrHead, "picture");

                foreach (var item in ManufactureArr)
                {
                    ws4.Cells[cRowFact, 1].Value = item.factory;
                    ws4.Cells[cRowFact, 2].Value = item.country;
                    cRowFact++;
                }

                var cR = 1;
                foreach (var item in brandArr)
                {
                    ws5.Cells[cRowBrand, 1].Value = cR;
                    ws5.Cells[cRowBrand, 2].Value = item.brand;
                    cR++;
                    cRowBrand++;
                }


                foreach (var item in CategoriesColl)
                {
                    ws2.Cells[cRowCat, 1].Value = item.id;
                    ws2.Cells[cRowCat, 2].Value = item.parentId;
                    ws2.Cells[cRowCat, 3].Value =
                        CategoriesColl.Where(x => x.id == item.parentId).Select(x => x.Name).FirstOrDefault();
                    ws2.Cells[cRowCat, 4].Value = item.Name;
                    cRowCat++;
                }

                //ws.Cells[1, 10].Value = "country_of_origin";
                //ws.Cells[1, 11].Value = "barcode";


                var startCol = 14;

                foreach (XmlNode isbn in nodeList)
                {
                    XmlNodeList nodeParams = isbn.SelectNodes("param");
                    foreach (XmlNode p in nodeParams)
                    {
                        if (!dAttr.ContainsKey(p.Attributes["name"].InnerText))
                        {
                            dAttr.Add(p.Attributes["name"].InnerText, startCol);
                            startCol++;
                        }
                    }
                }

                var d = from x in dAttr
                        select new
                        {
                            Name = x.Key,
                            Val = x.Value
                        };


                foreach (var item in d)
                {
                    SetCellHeader(ws.Cells[1, item.Val], Color.LightBlue, item.Name);
                    ws3.Cells[cRowAtr, 1].Value = item.Val;
                    ws3.Cells[cRowAtr, 2].Value = item.Name;
                    ws3.Cells[cRowAtr, 3].Value = item.Name;
                    ws3.Cells[cRowAtr, 4].Value = "f";
                    cRowAtr++;
                }

                ws3.Cells[4, 7].Value = "f - затягиваем в фильтры";
                ws3.Cells[5, 7].Value = "v - затягиваем в ВГХ";
                ws3.Cells[6, 7].Value = "fv - затягиваем в фильтры и в ВГХ";
                ws3.Cells[7, 7].Value = "n - не затягивать";

                foreach (XmlNode isbn in nodeList)
                {
                    ws.Cells[cRow, 1].Value = isbn.Attributes["available"].InnerText;
                    ws.Cells[cRow, 2].Value = isbn.Attributes["id"].InnerText;
                    ws.Cells[cRow, 3].Value = isbn["name"].InnerText;
                    ws.Cells[cRow, 4].Value = isbn["url"].InnerText;
                    ws.Cells[cRow, 5].Value = isbn["price"].InnerText;
                    ws.Cells[cRow, 6].Value = isbn["currencyId"].InnerText;
                    ws.Cells[cRow, 7].Value = isbn["categoryId"].InnerText;
                    ws.Cells[cRow, 8].Value =
                        CategoriesColl.Where(x => x.id == isbn["categoryId"].InnerText)
                            .Select(x => x.Name)
                            .FirstOrDefault();
                    ws.Cells[cRow, 9].Value = isbn["delivery"].InnerText;
                    ws.Cells[cRow, 10].Value = isbn["vendorCode"].InnerText;
                    ws.Cells[cRow, 11].Value = isbn["vendor"]?.InnerText ?? "";
                    ws.Cells[cRow, 12].Value = isbn["description"]?.InnerText ?? "";
                    ws.Cells[cRow, 13].Value = isbn["picture"].InnerText;

                    XmlNodeList nodeParams = isbn.SelectNodes("param");
                    foreach (XmlNode p in nodeParams)
                    {
                        var val = p.InnerText;
                        var unit = p.Attributes["unit"]?.InnerText ?? "";
                        unit = unit.Length > 0 ? @"###" + unit : "";
                        ws.Cells[cRow, dAttr[p.Attributes["name"].InnerText]].Value = $"{val} {unit}";

                        if (p.Attributes["name"].InnerText == "Производитель")
                        {
                        }

                        if (p.Attributes["name"].InnerText == "Бренд")
                        {
                        }

                    }
                    cRow++;
                }

                ws.Cells[ws.Dimension.Address].AutoFilter = true;

                ws2.Cells[ws2.Dimension.Address].AutoFilter = true;
                ws2.Cells[ws2.Dimension.Address].AutoFitColumns();

                ws3.Cells[ws3.Dimension.Address].AutoFilter = true;
                ws3.Cells[ws3.Dimension.Address].AutoFitColumns();

                ws4.Cells[ws4.Dimension.Address].AutoFilter = true;
                ws4.Cells[ws4.Dimension.Address].AutoFitColumns();

                ws5.Cells[ws5.Dimension.Address].AutoFilter = true;
                ws5.Cells[ws5.Dimension.Address].AutoFitColumns();

                package.SaveAs(file);
                Process.Start(file.FullName);
            }
        }
        
        private static void SetCellHeader(ExcelRange rg, Color clr, string val)
        {
            rg.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rg.Style.Fill.BackgroundColor.SetColor(clr);
            rg.Value = val;
        }

        private static IEnumerable<Category> GetCategoriesColl_old(XmlElement root)
        {
            return root.SelectNodes("/yml_catalog/shop/categories/category")
                .Cast<XmlNode>().Select(x => new Category
                {
                    id = x.Attributes["id"].InnerText,
                    parentId = x.Attributes["parentId"]?.InnerText ?? "",
                    Name = x.InnerText
                });
        }

        private static IEnumerable<Category> GetCategoriesColl(XmlElement root)
        {

            var tCat = root.SelectNodes("/yml_catalog/shop/categories/category")
                .Cast<XmlNode>().Select(x => new
                {
                    id = x.Attributes["id"].InnerText,
                    parentId = x.Attributes["parentId"]?.InnerText ?? "",
                    Name = x.InnerText
                });

            var catReturn = tCat.Select(x => new Category
                {
                    parentId = x.parentId,
                    id = x.id,
                    Name = x.Name,
                    NameNew = "",
                    ParentName = tCat.Where(c => c.id == x.parentId)
                        .Select(c => x.Name)
                        .FirstOrDefault(),
                    idInDB = ""
                }
            );
            return catReturn;
        }

        private static IEnumerable<Manufacture> GetManufacturerColl(XmlElement root)
        {
            return root.SelectNodes("/yml_catalog/shop/offers/offer/param[@name='Производитель']")
                .Cast<XmlNode>()
                .Select(x => new Manufacture
                {
                    factory = x.InnerText,
                    country = x.ParentNode["country"]?.InnerText ?? ""
                }).OrderBy(x => x.factory).Distinct();
        }

        private static IEnumerable<Brand> GetBrandColl(XmlElement root)
        {
            return root.SelectNodes("/yml_catalog/shop/offers/offer/param[@name='Бренд']")
                .Cast<XmlNode>()
                .Select(x => new Brand
                {
                    brand = x.InnerText
                }).OrderBy(x => x.brand).Distinct();
        }

        private static IEnumerable<XmlNode> GetOffer(XmlElement root)
        {
            return root.SelectNodes("/yml_catalog/shop/offers/offer")
                .Cast<XmlNode>()
                /*
                .Where(x =>
                {
                    var xmlElement = x["vendor"];
                    return xmlElement != null && xmlElement.InnerText == "Garnier";
                })
                */
                .OrderBy(r => r["name"].InnerText);
        }

        private void BulkСopyToDB(XmlElement root)
        {


            var nodeList = GetOffer(root);

            var brandArr = GetBrandColl(root);

            var factoryArr = GetManufacturerColl(root);

            var CategoriesColl = GetCategoriesColl(root);

            var DT = new DataTable();


            return;//Пока не понятно что и как...

            List<Item> MainTableColl = new List<Item>
            {
                new Item() {id = 1,Name = "available"},
                new Item() {id = 2,Name = "id"},
                new Item() {id = 3,Name = "name"},
                new Item() {id = 4,Name = "url"},
                new Item() {id = 5,Name = "price"},
                new Item() {id = 6,Name = "currencyId"},
                new Item() {id = 7,Name = "categoryId"},
                new Item() {id = 8,Name = "categoryName"},
                new Item() {id = 9,Name = "delivery"},
                new Item() {id = 10,Name = "vendorCode"},
                new Item() {id = 11,Name = "vendor"},
                new Item() {id = 12,Name = "description"},
                new Item() {id = 13,Name = "picture"}
            };

            //CategoriesColl.Where(x => x.id == isbn["categoryId"].InnerText).Select(x => x.Name).FirstOrDefault();


            using (var connection = new SqlConnection(Program.connectionStr))
            {
                connection.Open();
                /*
                var tSQL = "truncate table tmp_YML2";
                var cmd = new SqlCommand(tSQL, connection);
                cmd.CommandType = CommandType.Text;
                cmd.Connection = connection;
                cmd.ExecuteNonQuery();
                */
                using (var bulkCopy = new SqlBulkCopy(connection))
                {
                    //TODO Доделать заливку
                }
            }
        }



        private void button3_Click(object sender, EventArgs e)
        {
            XmlDocument doc = new XmlDocument();
            string FileName = Program.PathExcelFileImport;

            doc.Load(FileName);
            XmlElement root = doc.DocumentElement;
            var coll = GetCategoriesColl(root);
            dataGridView1.DataSource=SqlHelper.ToDataTable(coll.ToList());
        }
    }
}
