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
using System.Runtime.CompilerServices;
using OfficeOpenXml.Style;
using ParceYmlApp.Enums;

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
            Program.PathFolderBase = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"testXml");
            txbPathSelector.Text = Program.PathExcelFileImport;

            //(DateTime.Now).Subtract(new DateTime(1970, 1, 1)).TotalSeconds
            //var random = new Random((int)DateTime.Now.Ticks);
            //var s = DateTime.Now.ToString("yyyyMMdd");
            //Random rand = new Random(Guid.NewGuid().GetHashCode());
        }

        private void btnParseFromExcel_Click(object sender, EventArgs e)
        {
            //Program.PathFolderBase
            var fName = "TestOut.xlsx";

            var FileNameOut = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory + @"..\\..\\..\\" + fName);
            //var FileNameIn = Path.GetFileNameWithoutExtension(Program.PathExcelFileImport) + ".xlsx";
            var FileNameIn = @"c:\333\ParceYML\soap.xlsx";
            DataTable dt = new DataTable();

            FileInfo existingFile = new FileInfo(FileNameIn);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //ExcelWorksheet ws = package.Workbook.Worksheets["Фильтры"];
                ExcelWorksheet ws = package.Workbook.Worksheets[(int)enWsName.Распарсен];

                //TODO Обработать ошибку задвоения колонок 
                dt = GetDataTableFromWS(ws);
                dataGridView1.DataSource = dt;

                //dt.Columns.Cast<DataColumn>().GroupBy(v => v).Where(g => g.Count() > 1).ToList()
                //dt.Columns.Cast<DataColumn>().Select(x => x.ColumnName.ToUpper()).GroupBy(n=>n).Where(g=>g.Count()>1).ToList()
                var arrName = dt.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToArray();
                //lblInfo.Text =  $"Добавлено - {BulkСopyToDB(dt)} строки";
            }
        }
        //list.GroupBy(v => v).Where(g => g.Count() > 1).Select(g => g.Key)

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
            var productColl = GetOffer(root);

            var param1 = GetParam(root);

            var brandColl = GetBrandColl(root);
            var manufactureColl = GetManufacturerColl(root);
            var categoriesColl = GetCategoriesColl(root);

            //if (chbCopyToDB.Checked) BulkСopyToDB(root);

            string fileName = Path.GetFileNameWithoutExtension(FileName) + ".xlsx";
            string outputDir = Path.GetDirectoryName(FileName);

            Dictionary<string, int> dAttr = new Dictionary<string, int>();

            var file = new FileInfo(outputDir + '\\' + fileName);
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet ws = package.Workbook.Worksheets.Add("Распарсен");
                ExcelWorksheet wsCatigoty = package.Workbook.Worksheets.Add(enWsName.Категории.ToString());
                ExcelWorksheet wsParam = package.Workbook.Worksheets.Add(enWsName.Фильтры.ToString());
                ExcelWorksheet wsManuf = package.Workbook.Worksheets.Add(enWsName.Производители.ToString());
                ExcelWorksheet wsBrand = package.Workbook.Worksheets.Add(enWsName.Бренды.ToString());

                int cRow = 3;
                int cRowAtr = 3;
                int cRowCat = 3;
                int cRowManuf = 3;
                int cRowBrand = 3;

                var clrHead = Color.LightSkyBlue;
                var sCol = 1;
                var sRow = 2;
                
                List<RowItem> lstWsTitle = new List<RowItem>
                {
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "available", Name = "available", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "id", Name = "id", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "name", Name = "name", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "url", Name = "url", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "price", Name = "price", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "currencyId", Name = "currencyId", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "categoryId", Name = "categoryId", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "categoryName", Name = "categoryName", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "delivery", Name = "delivery", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "vendorCode", Name = "vendorCode", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "vendor", Name = "vendor", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "description", Name = "description", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "picture", Name = "picture", Color = clrHead}
                };
                InitTitleWS(lstWsTitle, ws);

                sCol = 1;
                List<RowItem> lstTitleWsManuf = new List<RowItem>
                {
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Nom", Name = "Nom", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Название", Name = "name", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "НазваниеTBN", Name = "NameTbn", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Страна", Name = "country", Color = clrHead}
                };
                InitTitleWS(lstTitleWsManuf, wsManuf);


                sCol = 1;
                List<RowItem> lstTitleWsBrand = new List<RowItem>
                {
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "№", Name = "Nom", Color = clrHead},
                    //new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "КодБренда", Name = "BrendCode", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Название", Name = "ame", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "НазваниеTBN", Name = "NameTbn", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Страна", Name = "Country", Color = clrHead}
                };
                InitTitleWS(lstTitleWsBrand, wsBrand);


                sCol = 1;
                List<RowItem> lstTitleWsParam = new List<RowItem>
                {
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "ID", Name = "Id", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Название", Name = "Name", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "НазваниеTBN", Name = "NameTbn", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Тип", Name = "ParamType", Color = clrHead}
                };
                InitTitleWS(lstTitleWsParam, wsParam);


                sCol = 1;
                List<RowItem> lstTitleWsCatigoty = new List<RowItem>
                {
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "ID", Name = "id", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "ParentId", Name = "parentId", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "parentName", Name = "parentName", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Name", Name = "Name", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "НашId", Name = "row_id", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "НашаКатегория", Name = "CatId", Color = clrHead}
                };
                InitTitleWS(lstTitleWsCatigoty, wsCatigoty);



                var cRowNom = 1;
                foreach (var item in manufactureColl)
                {
                    wsManuf.Cells[cRowManuf, 1].Value = cRowNom;
                    //wsManuf.Cells[cRowManuf, 2].Value = item.Name;
                    SetCellHeader(wsManuf.Cells[cRowManuf, 2], Color.LightGray, item.Name);
                    wsManuf.Cells[cRowManuf, 3].Value = item.Name;
                    wsManuf.Cells[cRowManuf, 4].Value = item.country;
                    cRowNom++;
                    cRowManuf++;
                }

                cRowNom = 1;
                foreach (Brand item in brandColl)
                {
                    wsBrand.Cells[cRowBrand, 1].Value = cRowNom;
                    //wsBrand.Cells[cRowBrand, 2].Value = item.brandCode;
                    //ws5.Cells[cRowBrand, 3].Value = item.brandName;
                    SetCellHeader(wsBrand.Cells[cRowBrand, 2], Color.LightGray, item.brandName);
                    wsBrand.Cells[cRowBrand, 3].Value = item.brandName;
                    wsBrand.Cells[cRowBrand, 4].Value = item.country;
                    cRowNom++;
                    cRowBrand++;
                }


                foreach (var item in categoriesColl)
                {
                    wsCatigoty.Cells[cRowCat, 1].Value = item.id;
                    wsCatigoty.Cells[cRowCat, 2].Value = item.parentId;
                    wsCatigoty.Cells[cRowCat, 3].Value =
                        categoriesColl.Where(x => x.id == item.parentId).Select(x => x.Name).FirstOrDefault();
                    wsCatigoty.Cells[cRowCat, 4].Value = item.Name;
                    SetCellHeader(wsCatigoty.Cells[cRowCat, 5], Color.LightGoldenrodYellow, "");
                    //SetCellHeader(wsCatigoty.Cells[cRowCat, 6], Color.LightGoldenrodYellow, "");
                    cRowCat++;
                }

                //ws.Cells[1, 10].Value = "country_of_origin";
                //ws.Cells[1, 11].Value = "barcode";


                var startCol = 14;

                foreach (XmlNode isbn in productColl)
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
                    SetCellHeader(ws.Cells[2, item.Val], Color.LightBlue, item.Name);

                    wsParam.Cells[cRowAtr, 1].Value = item.Val;
                    //wsParam.Cells[cRowAtr, 2].Value = item.Name;
                    SetCellHeader(wsParam.Cells[cRowAtr, 2], Color.LightGray, item.Name);
                    wsParam.Cells[cRowAtr, 3].Value = item.Name;
                    wsParam.Cells[cRowAtr, 4].Value = "f";
                    cRowAtr++;
                }

                wsParam.Cells[4, 7].Value = "f - затягиваем в фильтры";
                wsParam.Cells[5, 7].Value = "v - затягиваем в ВГХ";
                wsParam.Cells[6, 7].Value = "fv - затягиваем в фильтры и в ВГХ";
                wsParam.Cells[7, 7].Value = "n - не затягивать";

                foreach (XmlNode isbn in productColl)
                {
                    ws.Cells[cRow, 1].Value = isbn.Attributes["available"].InnerText;
                    ws.Cells[cRow, 2].Value = isbn.Attributes["id"].InnerText;
                    ws.Cells[cRow, 3].Value = isbn["name"].InnerText;
                    ws.Cells[cRow, 4].Value = isbn["url"].InnerText;
                    ws.Cells[cRow, 5].Value = isbn["price"].InnerText;
                    ws.Cells[cRow, 6].Value = isbn["currencyId"].InnerText;
                    ws.Cells[cRow, 7].Value = isbn["categoryId"].InnerText;
                    ws.Cells[cRow, 8].Value =
                        categoriesColl.Where(x => x.id == isbn["categoryId"].InnerText)
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

                wsCatigoty.Cells[wsCatigoty.Dimension.Address].AutoFilter = true;
                wsCatigoty.Cells[wsCatigoty.Dimension.Address].AutoFitColumns();

                wsParam.Cells[wsParam.Dimension.Address].AutoFilter = true;
                wsParam.Cells[wsParam.Dimension.Address].AutoFitColumns();

                wsManuf.Cells[wsManuf.Dimension.Address].AutoFilter = true;
                wsManuf.Cells[wsManuf.Dimension.Address].AutoFitColumns();

                wsBrand.Cells[wsBrand.Dimension.Address].AutoFilter = true;
                wsBrand.Cells[wsBrand.Dimension.Address].AutoFitColumns();

                package.SaveAs(file);
                Process.Start(file.FullName);
            }
        }

        private static void InitTitleWS(List<RowItem> lst, ExcelWorksheet ws)
        {
            foreach (var row in lst)
            {
                SetCellHeader(ws.Cells[row.RowNom, row.ColNom], row.Color, row.NameCol);
                SetCellHeader(ws.Cells[row.RowNom-1, row.ColNom], Color.Lavender, row.Name);
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

        private static IEnumerable<RowItem> GetParam(XmlElement root)
        {

            var ret = root.SelectNodes("/yml_catalog/shop/offers/offer/param")
                .Cast<XmlNode>()
                .Select(r => new RowItem
                {
                    InnerText = r.InnerXml,
                    Name = r.Attributes["name"]?.InnerText??"",
                    Unit = r.Attributes["unit"]?.InnerText??""
                });

            var duplicates = ret.Select(r => r.Name).Distinct().Select(r=>r.ToUpper())
                                .GroupBy(g => g).Where(w => w.Count() > 1).Select(w=>w.First()).ToList();
            if (duplicates.Count>0)
            {
                var sErr = "Найдены дубликаты названий фильтров! \n" +
                           "Исправьте исходный файл и повторите затяжку\n";
                foreach (var r in duplicates)
                {
                    sErr += $"\t'{r}'\n";
                }
                sErr += $"Выйти из программы?";

                DialogResult dialogResult  = MessageBox.Show(sErr,"Ошибка!",MessageBoxButtons.OKCancel,MessageBoxIcon.Error);
                Console.WriteLine(sErr);
                if (dialogResult == DialogResult.OK)
                {
                    Application.Exit();
                    return null;
                }
            }
            return ret;
        }

        private static IEnumerable<Brand> GetBrandColl(XmlElement root)
        {
            var random = new Random((int)DateTime.Now.Ticks);
            var s = DateTime.Now.ToString("yyyyMMdd");

            var ret = root.SelectNodes("/yml_catalog/shop/offers/offer/param[@name='Бренд']")
                .Cast<XmlNode>()
                .Select(x => new
                {
                    brandName = x.InnerText,
                    //country = x.ParentNode["country"]?.InnerText ?? ""
                }).Distinct().OrderBy(x => x.brandName);

            var ret1 = ret.Select(x => new Brand
            {
                brandName = x.brandName
                //,brandCode = s + random.Next(1000, 9999),
            });

            return ret1;
        }

        private static IEnumerable<Manufacture> GetManufacturerColl(XmlElement root)
        {
            var ret = root.SelectNodes("/yml_catalog/shop/offers/offer/param[@name='Производитель']")
                .Cast<XmlNode>()
                .Select(x => new
                {
                    factory = x.InnerText
                    //country = x.ParentNode["country"]?.InnerText ?? ""
                }).OrderBy(x => x.factory).Distinct();

            var ret1 = ret.Select(x => new Manufacture
            {
                Name = x.factory
            }).OrderBy(f => f.Name);
            return ret1;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            XmlDocument doc = new XmlDocument();
            string FileName = Program.PathExcelFileImport;

            doc.Load(FileName);
            XmlElement root = doc.DocumentElement;
            //var coll = GetCategoriesColl(root);
            //var coll = GetBrandColl(root);
            var coll = GetManufacturerColl(root);
            lblInfo.Text = $"Выбрано - {coll.Count()} строк";
            dataGridView1.DataSource = SqlHelper.ToDataTable(coll.ToList());
        }


        /// <summary>
        /// Поля возвращаемой таблицы соответствуют названиям колонок первой строки из листа Excel
        /// </summary>
        private static DataTable GetDataTableFromWS(ExcelWorksheet ws)
        {
            DataTable dtResult = new DataTable();
            List<object> WorksheetRowsColl = new List<object>();
            var rowCount = ws.Dimension.End.Row; //Utils.GetLastUsedRow(ws);
            var сolCount = ws.Dimension.End.Column;

            // TODO Обработать ошибку задвоения колонок
            //ws.Cells[1, 1, 1, сolCount]

            for (var rowNum = 1; rowNum <= rowCount; rowNum++)
            {
                var row = ws.Cells[rowNum, 1, rowNum, сolCount];
                bool allEmpty = row.All(c => string.IsNullOrWhiteSpace(c.Text));
                if (allEmpty) continue;
                WorksheetRowsColl.Add(row.Value);
            }

            //dt.Columns.Cast<DataColumn>().Select(x => x.ColumnName.ToUpper()).GroupBy(n=>n).Where(g=>g.Count()>1).ToList()

            var titleColName = (object[,])WorksheetRowsColl[0];
            for (int k = 0; k < titleColName.Length; k++)
            {

                if (titleColName[0, k] == null) continue;
                dtResult.Columns.Add(titleColName[0, k].ToString(), typeof(string));
            }

            dtResult.Columns.Add("row_id", typeof(int));

            var i = 1;
            foreach (object[,] row in WorksheetRowsColl)
            {
                var dr = dtResult.Rows.Add();
                for (int j = 0; j < row.Length; j++)
                {
                    if (titleColName[0, j] == null) continue;
                    dr[(string)titleColName[0, j]] = row[0, j];
                }
                dr["row_id"] = i;
                i++;
            }


            return dtResult;
        }
        
        private long BulkСopyToDB(DataTable dt)
        {
            //DateTime.Now.Millisecond
            var insRowsCount = 0;
            using (var connection = new SqlConnection(Program.connectionStr))
            {
                connection.Open();
                string tTable = "dbo.tmp_YML2";
                SqlCommand commandRowCount = new SqlCommand($"select count(*) from {tTable}", connection);

                string tSQL = $" IF EXISTS (select * from dbo.sysobjects where id = object_id('{tTable}')) " +
                              $" drop table {tTable}";
                var cmd = new SqlCommand(tSQL, connection);
                cmd.CommandType = CommandType.Text;
                cmd.Connection = connection;
                cmd.ExecuteNonQuery();


                GreateTable(dt, connection, tTable);

                using (var bulkCopy = new SqlBulkCopy(connection))
                {
                    bulkCopy.DestinationTableName = tTable;
                    bulkCopy.WriteToServer(dt);
                }
                insRowsCount = System.Convert.ToInt32(commandRowCount.ExecuteScalar());
            }
            return insRowsCount;
            /*
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
            */
        }

        private void GreateTable(DataTable dt, SqlConnection connection, string tTable)
        {

            var sSQL = $"CREATE TABLE {tTable} (";
            for (int i = 0; i < dt.Columns.Count - 1; i++)
            {
                sSQL += $"[{dt.Columns[i].Caption}] nvarchar(MAX) NULL,";
            }
            sSQL += "[row_id][int] IDENTITY(1, 1) NOT NULL";
            sSQL +=
                $" CONSTRAINT [PK_{tTable}] PRIMARY KEY CLUSTERED ([row_id] ASC) " +
                $" WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF" +
                $", ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON[PRIMARY]) ON[PRIMARY] TEXTIMAGE_ON[PRIMARY]";

            SqlCommand createtable = new SqlCommand(sSQL, connection);
            createtable.ExecuteNonQuery();

        }
    }
}
