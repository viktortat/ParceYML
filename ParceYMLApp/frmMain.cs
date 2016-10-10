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
using System.Globalization;
using System.Runtime.CompilerServices;
using OfficeOpenXml.Style;
using ParceYmlApp.Enums;

namespace ParceYmlApp
{
    /*
    public static class ExcelWorksheetExtension
    {
        public static string[] GetHeaderColumns(this ExcelWorksheet sheet)
        {
            List<string> columnNames = new List<string>();
            //foreach (var firstRowCell in sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, 1, sheet.Dimension.End.Column])
            foreach (var firstRowCell in sheet.Cells[1, sheet.Dimension.Start.Column, 1, sheet.Dimension.End.Column])
                columnNames.Add(firstRowCell.Text);
            return columnNames.ToArray();
        }
    }
    */
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            Program.connectionStr = Program.connectionStr = ConfigurationManager.ConnectionStrings["TbnProd.Local"].ConnectionString;
            //Program.PathExcelFile = AppDomain.CurrentDomain.BaseDirectory + @"testXml\soap.xml";
            //Program.PathFolderBase = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"testXml");
            //Program.PathFolderBase
            //var fName = "TestOut.xlsx";
            //var FileNameOut = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory + @"..\\..\\..\\" + fName);
            ////var FileNameIn = Path.GetFileNameWithoutExtension(Program.PathExcelFile) + ".xlsx";
            ////var FileNameIn = @"c:\333\ParceYML\soap.xlsx";
            //var FileNameIn = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory + @"testXml\soapIdeal.xlsx");
            ////var FileNameIn = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory+ @"testXml\soap.xlsx");

            Program.PathFolderBase = AppDomain.CurrentDomain.BaseDirectory;
            Program.InsertToDB = chbCopyToDB.Checked;
            txbPathSelector.Text = Program.PathExcelFile;

            //(DateTime.Now).Subtract(new DateTime(1970, 1, 1)).TotalSeconds
            //var random = new Random((int)DateTime.Now.Ticks);
            //var s = DateTime.Now.ToString("yyyyMMdd");
            //Random rand = new Random(Guid.NewGuid().GetHashCode());
        }

        private void btnParseFromExcel_Click(object sender, EventArgs e)
        {
            SelPathExcelFileImport(Filter: "Все файлы (*.*)|*.*|Excel файлы (*.xlsx)|*.xlsx"
             , Title: "Выберите EXCEL файл для записи в БД");

            DataTable dtMain = new DataTable();

            FileInfo existingFile = new FileInfo(Program.PathExcelFile);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet ws = package.Workbook.Worksheets[(int)enWsName.Распарсен];

                List<string> columnNames = new List<string>();
                foreach (var firstRowCell in ws.Cells[ws.Dimension.Start.Row, ws.Dimension.Start.Column, 1, ws.Dimension.End.Column])
                    columnNames.Add(firstRowCell.Text);

                var duplicates = columnNames.Select(x => x.ToUpper()).GroupBy(v => v).Where(g => g.Count() > 1).ToList();
                if (duplicates.Count > 0)
                {
                    var sErr = "Найдены дубликаты названий колонок! \n" +
                               "Исправьте исходный файл и повторите затяжку\n";
                    foreach (var r in duplicates)
                    {
                        sErr += $"\t'{r.Key}'\n";
                    }
                    lblInfo.Text = sErr;
                    DialogResult dialogResult = MessageBox.Show(sErr, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Console.WriteLine(sErr);
                    if (dialogResult == DialogResult.OK)
                    {
                        Application.Exit();
                        return;
                    }
                }

                dtMain = GetDataTableFromWS(ws);
                string tableName = "tmp_YML2";
                if (Program.InsertToDB)
                    lblInfo.Text = $"Добавлено - {BulkСopyToDB(dtMain, "dbo.tmp_YML2")} строки в {tableName}";
                Console.WriteLine($"select * from {tableName}");

                InsParamsInDb(package);
                InsProductInDb(package);

            }
        }

        private void btnSelFile_Click(object sender, EventArgs e)
        {
            SelPathExcelFileImport();
        }


        private void SelPathExcelFileImport(
            string Filter = "Все файлы (*.*)|*.*|Excel файлы (*.xlsx)|*.xlsx|YML файлы (*.xml)|*.xml"
            , string Title = "Выбор файла"
            )
        {
            OpenFileDialog dlgSelFile = new OpenFileDialog();
            dlgSelFile.InitialDirectory = Program.PathFolderBase;
            dlgSelFile.Filter = Filter;
            dlgSelFile.FilterIndex = 2;
            dlgSelFile.RestoreDirectory = true;
            dlgSelFile.Title = Title;
            if (dlgSelFile.ShowDialog() == DialogResult.OK)
            {
                var selParh = dlgSelFile.FileName;
                if (Path.GetExtension(selParh)?.ToUpper() == ".XML")
                    Program.PathXmlFile = selParh;
                else
                    Program.PathExcelFile = selParh;

                txbPathSelector.Text = Path.GetFullPath(selParh);
            }
        }

        private void btnCreatePricat_Click(object sender, EventArgs e)
        {
            int startRow = 4;

            SelPathExcelFileImport(Filter: "Все файлы (*.*)|*.*|YML файлы (*.xml)|*.xml"
    , Title: "Выберите YML-XML файл для разбора данных в EXCEL");

            XmlDocument doc = new XmlDocument();
            string FileName = Program.PathXmlFile;
            doc.Load(FileName);
            XmlNodeList nodeList;
            XmlElement root = doc.DocumentElement;
            nodeList = root.SelectNodes("/yml_catalog/shop/offers/offer");

            string fileName = "Pricat.xlsx";
            string fileNameT = "Торговое предложение.xltx";

            string outputDir = Program.PathFolderBase;

            var file = new FileInfo(outputDir + fileName);
            var newTFile = new FileInfo(Path.GetDirectoryName(Application.ExecutablePath) + @"\Tmpl\" + fileNameT);

            FileInfo empTpl = null;
            var basePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Tmpl", "Торговое предложение.xltx");


            if (File.Exists(basePath))
                empTpl = new FileInfo(basePath);
            else
                empTpl = new FileInfo(basePath);

            using (ExcelPackage package = new ExcelPackage(empTpl, true))
            {
                //_excelPackage = new ExcelPackage(newFile);
                ExcelWorksheet ws = package.Workbook.Worksheets.First(); //.Add("Прайс лист");
                foreach (XmlNode isbn in nodeList)
                {
                    ws.Cells[startRow, 1].Value = isbn["name"].InnerText;
                    if (isbn["barcode"] != null)
                    {
                        ws.Cells[startRow, 2].Value = isbn["barcode"].InnerText;
                    }
                    ws.Cells[startRow, 3].Value = isbn.Attributes["id"].InnerText;
                    ws.Cells[startRow, 4].Value = "";
                    ws.Cells[startRow, 5].Value = "1";
                    ws.Cells[startRow, 6].Value = "0";

                    decimal tPrice;
                    decimal.TryParse(((string)isbn["price"].InnerText)?.Replace(",", "."), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out tPrice);
                    ws.Cells[startRow, 7].Value = tPrice.ToString();
                    //ws.Cells[startRow, 8].Value = isbn["categoryId"].InnerText;

                    startRow++;
                }
                package.SaveAs(file);
            }
            Process.Start(file.FullName);
        }

        private void btnParseInExcel_Click(object sender, EventArgs e)
        {
            var startCol = 19;
            int cRow = 3;
            int cRowAtr = 3;
            int cRowCat = 3;
            int cRowManuf = 3;
            int cRowBrand = 3;

            var sCol = 1;
            var sRow = 2;

            var clrHead = Color.LightSkyBlue;

            SelPathExcelFileImport(Filter: "Все файлы (*.*)|*.*|YML файлы (*.xml)|*.xml"
                , Title: "Выберите YML-XML файл для разбора данных в EXCEL");

            XmlDocument doc = new XmlDocument();
            string FileName = Program.PathXmlFile;

            doc.Load(FileName);
            XmlElement root = doc.DocumentElement;
            var productColl = GetProductColl(root);

            var brandColl = GetBrandColl(root);
            var manufactureColl = GetManufacturerColl(root);
            var categoriesColl = GetCategoriesColl(root);

            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet ws = package.Workbook.Worksheets.Add("Распарсен");
                ExcelWorksheet wsCatigoty = package.Workbook.Worksheets.Add(enWsName.Категории.ToString());
                ExcelWorksheet wsParam = package.Workbook.Worksheets.Add(enWsName.Фильтры.ToString());
                ExcelWorksheet wsManuf = package.Workbook.Worksheets.Add(enWsName.Производители.ToString());
                ExcelWorksheet wsBrand = package.Workbook.Worksheets.Add(enWsName.Бренды.ToString());

                List<RowItem> lstWsTitle = new List<RowItem>
                {
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Nom", Name = "Row_id", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Доступен", Name = "Available", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Код-продукта", Name = "ProductId", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Название", Name = "Name", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Тип", Name = "ProdType", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Вид", Name = "ProdKind", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "url", Name = "Url", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Цена", Name = "Price", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Валюта", Name = "CurrencyId", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Категория-код", Name = "CategoryId", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Категория", Name = "CategoryName", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Доставка", Name = "Delivery", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Продавец-код", Name = "VendorCode", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Продавец", Name = "Vendor", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Описание", Name = "Description", Color = clrHead },
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Фото", Name = "Picture", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Штрих-код", Name = "BarCode", Color = clrHead}
                };
                InitTitleWS(lstWsTitle, ws);

                sCol = 1;
                List<RowItem> lstTitleWsManuf = new List<RowItem>
                {
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Nom", Name = "Row_id", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Название", Name = "Name", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "НазваниеTBN", Name = "NameTbn", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Страна", Name = "country", Color = clrHead}
                };
                InitTitleWS(lstTitleWsManuf, wsManuf);


                sCol = 1;
                List<RowItem> lstTitleWsBrand = new List<RowItem>
                {
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Nom", Name = "Row_id", Color = clrHead},
                    //new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "КодБренда", Name = "BrendCode", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Название", Name = "Name", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "НазваниеTBN", Name = "NameTbn", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Страна", Name = "Country", Color = clrHead}
                };
                InitTitleWS(lstTitleWsBrand, wsBrand);


                sCol = 1;
                List<RowItem> lstTitleWsParam = new List<RowItem>
                {
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Nom", Name = "Row_id", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "ParamId", Name = "ParamId", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Название", Name = "Name", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "НазваниеTBN", Name = "NameTbn", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Тип", Name = "ParamType", Color = clrHead}
                };
                InitTitleWS(lstTitleWsParam, wsParam);


                sCol = 1;
                List<RowItem> lstTitleWsCatigoty = new List<RowItem>
                {
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Nom", Name = "Row_id", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "CatID", Name = "CatId", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "ParentId", Name = "ParentId", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "ParentName", Name = "ParentName", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "Название", Name = "Name", Color = clrHead},
                    new RowItem() {RowNom = sRow, ColNom = sCol++, NameCol = "НашId", Name = "CatIdDB", Color = clrHead}
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

                cRowNom = 1;
                foreach (var item in categoriesColl)
                {
                    wsCatigoty.Cells[cRowCat, 1].Value = cRowNom;
                    wsCatigoty.Cells[cRowCat, 2].Value = item.id;
                    wsCatigoty.Cells[cRowCat, 3].Value = item.parentId;
                    wsCatigoty.Cells[cRowCat, 4].Value =
                        categoriesColl.Where(x => x.id == item.parentId).Select(x => x.Name).FirstOrDefault();
                    wsCatigoty.Cells[cRowCat, 5].Value = item.Name;
                    SetCellHeader(wsCatigoty.Cells[cRowCat, 6], Color.LightGoldenrodYellow, "");
                    //SetCellHeader(wsCatigoty.Cells[cRowCat, 6], Color.LightGoldenrodYellow, "");
                    cRowNom++;
                    cRowCat++;
                }

                //ws.Cells[1, 10].Value = "country_of_origin";
                //ws.Cells[1, 11].Value = "barcode";

                Dictionary<string, int> dicParam = new Dictionary<string, int>();
                foreach (XmlNode isbn in productColl)
                {
                    XmlNodeList nodeParams = isbn.SelectNodes("param");
                    foreach (XmlNode p in nodeParams)
                    {
                        if (!dicParam.ContainsKey(p.Attributes["name"].InnerText))
                        {
                            dicParam.Add(p.Attributes["name"].InnerText, startCol);
                            startCol++;
                        }
                    }
                }

                var dParam = (from x in dicParam
                              select new
                              {
                                  Name = x.Key,
                                  Val = x.Value
                              }).OrderBy(x => x.Name);

                cRowNom = 1;
                foreach (var item in dParam)
                {
                    SetCellHeader(ws.Cells[1, item.Val], Color.AntiqueWhite, item.Name);
                    SetCellHeader(ws.Cells[2, item.Val], Color.LightBlue, item.Name);

                    wsParam.Cells[cRowAtr, 1].Value = cRowNom;
                    wsParam.Cells[cRowAtr, 2].Value = item.Val;
                    //wsParam.Cells[cRowAtr, 2].Value = item.Name;
                    SetCellHeader(wsParam.Cells[cRowAtr, 3], Color.LightGray, item.Name);
                    wsParam.Cells[cRowAtr, 4].Value = item.Name;
                    wsParam.Cells[cRowAtr, 5].Value = "f";
                    cRowNom++;
                    cRowAtr++;
                }

                wsParam.Cells[4, 7].Value = "f - затягиваем в фильтры";
                wsParam.Cells[5, 7].Value = "v - затягиваем в ВГХ";
                wsParam.Cells[6, 7].Value = "fv - затягиваем в фильтры и в ВГХ";
                wsParam.Cells[7, 7].Value = "n - не затягивать";

                cRowNom = 1;
                foreach (XmlNode isbn in productColl)
                {
                    ws.Cells[cRow, 1].Value = cRowNom;
                    ws.Cells[cRow, 2].Value = isbn.Attributes["available"].InnerText;
                    ws.Cells[cRow, 3].Value = isbn.Attributes["id"].InnerText;
                    ws.Cells[cRow, 4].Value = isbn["name"].InnerText;
                    ws.Cells[cRow, 5].Value = "";
                    ws.Cells[cRow, 6].Value = "";
                    ws.Cells[cRow, 7].Value = isbn["url"].InnerText;
                    ws.Cells[cRow, 8].Value = isbn["price"].InnerText;
                    ws.Cells[cRow, 9].Value = isbn["currencyId"].InnerText;
                    ws.Cells[cRow, 10].Value = isbn["categoryId"].InnerText;
                    ws.Cells[cRow, 11].Value =
                        categoriesColl.Where(x => x.id == isbn["categoryId"].InnerText)
                            .Select(x => x.Name)
                            .FirstOrDefault();
                    ws.Cells[cRow, 12].Value = isbn["delivery"]?.InnerText ?? "";
                    ws.Cells[cRow, 13].Value = isbn["vendorCode"]?.InnerText ?? "";
                    ws.Cells[cRow, 14].Value = isbn["vendor"]?.InnerText ?? "";
                    ws.Cells[cRow, 15].Value = isbn["description"]?.InnerText ?? "";
                    ws.Cells[cRow, 16].Value = isbn["picture"].InnerText;
                    ws.Cells[cRow, 17].Value = isbn["barcode"]?.InnerText ?? "";

                    XmlNodeList nodeParams = isbn.SelectNodes("param");
                    foreach (XmlNode p in nodeParams)
                    {
                        var val = p.InnerText;
                        var unit = p.Attributes["unit"]?.InnerText ?? "";
                        unit = unit.Length > 0 ? @"###" + unit : "";
                        ws.Cells[cRow, dicParam[p.Attributes["name"].InnerText]].Value = $"{val} {unit}";
                    }
                    cRowNom++;
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


                string fNameOut = Path.GetFileNameWithoutExtension(FileName) + ".xlsx";
                string outputDir = Path.GetDirectoryName(FileName);

                var fileOut = new FileInfo(outputDir+"/"+fNameOut);
                package.SaveAs(fileOut);
                Process.Start(fileOut.FullName);
            }
        }

        private static void InitTitleWS(List<RowItem> lst, ExcelWorksheet ws)
        {
            foreach (var row in lst)
            {
                SetCellHeader(ws.Cells[row.RowNom, row.ColNom], row.Color, row.NameCol);
                SetCellHeader(ws.Cells[row.RowNom - 1, row.ColNom], Color.Lavender, row.Name);
            }
        }

        private static void SetCellHeader(ExcelRange rg, Color clr, string val)
        {
            rg.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rg.Style.Fill.BackgroundColor.SetColor(clr);
            if (!string.IsNullOrEmpty(val)) rg.Value = val;
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

        private static IEnumerable<XmlNode> GetProductColl(XmlElement root)
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
                    Name = r.Attributes["name"]?.InnerText ?? "",
                    Unit = r.Attributes["unit"]?.InnerText ?? ""
                });

            var duplicates = ret.Select(r => r.Name).Distinct().Select(r => r.ToUpper())
                                .GroupBy(g => g).Where(w => w.Count() > 1).Select(w => w.First()).ToList();
            if (duplicates.Count > 0)
            {
                var sErr = "Найдены дубликаты названий фильтров! \n" +
                           "Исправьте исходный файл и повторите затяжку\n";
                foreach (var r in duplicates)
                {
                    sErr += $"\t'{r}'\n";
                }
                sErr += $"Выйти из программы?";

                DialogResult dialogResult = MessageBox.Show(sErr, "Ошибка!", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
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
            string FileName = Program.PathExcelFile;

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

            dtResult.Columns.Add("rowNom", typeof(int));

            var i = 1;

            //TODO - как по свободе переделать этот костыль! Не нравится!!!
            var startRow = 2;
            for (int k = startRow; k < WorksheetRowsColl.Count; k++)
            {
                var dr = dtResult.Rows.Add();
                var row = WorksheetRowsColl[k];
                for (int j = 0; j < ((object[,])row).Length; j++)
                {
                    if (titleColName[0, j] == null) continue;
                    dr[(string)titleColName[0, j]] = ((object[,])row)[0, j];
                }
                dr["rowNom"] = i;
                i++;
            }

            /*
            foreach (object[,] row in WorksheetRowsColl)
            {
                var dr = dtResult.Rows.Add();
                for (int j = 0; j < row.Length; j++)
                {
                    if (titleColName[0, j] == null) continue;
                    dr[(string)titleColName[0, j]] = row[0, j];
                }
                dr["rowNom"] = i;
                i++;
            }
            */

            return dtResult;
        }

        private long BulkСopyToDB(DataTable dt, string tTable)
        {
            var insRowsCount = 0;
            using (var connection = new SqlConnection(Program.connectionStr))
            {
                connection.Open();
                SqlCommand commandRowCount = new SqlCommand($"select count(*) from {tTable}", connection);

                CreateTable(dt, connection, tTable);

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

        private void InsParamsInDb(ExcelPackage ex)
        {
            ExcelWorksheet wsParam = ex.Workbook.Worksheets[(int)enWsName.Фильтры];
            var dtParams = GetDataTableFromWS(wsParam);
            List<RowItemParam> lp = new List<RowItemParam>();
            foreach (DataRow row in dtParams.Rows)
            {
                lp.Add(new RowItemParam()
                {
                    Row_id = Convert.ToInt64(row["Row_id"]),
                    ParamId = Convert.ToInt64(row["ParamId"]),
                    Name = (string)row["Name"],
                    NameTbn = (string)row["NameTbn"],
                    ParamType = (string)row["ParamType"]
                });
            }

            var dtParam = SqlHelper.ToDataTable(lp.ToList());
            dataGridView1.DataSource = dtParam;
            string tableName = "tmp_YML2Params";
            if (Program.InsertToDB)
                lblInfo.Text = $"Добавлено - {BulkСopyToDB(dtParam, tableName)} строки в {tableName}";
            Console.WriteLine($"select * from {tableName}");

            /*
            IEnumerable<RowItem> query = dtParams.AsEnumerable()
                //where order.Field<DateTime>("OrderDate") > new DateTime(2001, 8, 1)
                .Select((c, index) => new RowItem
                {
                    RowNom = index,
                    Name = c[0].ToString(),
                    NameCol = c[1].ToString()
                }).Where(b => b.RowNom > 0).ToList();
            */
        }
        private void InsProductInDb(ExcelPackage ex)
        {
            ExcelWorksheet wsProduct = ex.Workbook.Worksheets[(int)enWsName.Распарсен];
            var dtMain = GetDataTableFromWS(wsProduct);
            List<OfferItem> lp = new List<OfferItem>();
            foreach (DataRow row in dtMain.Rows)
            {
                double tPrice;
                double.TryParse(((string)row["Price"])?.Replace(",", "."), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out tPrice);

                lp.Add(new OfferItem()
                {
                    Row_id = Convert.ToInt64(row["Row_id"]),
                    Available = Convert.ToBoolean(row["Available"]),
                    ProductId = (string)row["ProductId"],
                    Name = (string)row["Name"],
                    ProdType = (string)row["ProdType"],
                    ProdKind = (string)row["ProdKind"],
                    Url = (string)row["Url"],
                    Price = tPrice,//Convert.ToDouble(row["Price"]),
                    CurrencyId = (string)row["CurrencyId"],
                    CategoryId = (string)row["CategoryId"],
                    CategoryName = (string)row["CategoryName"],
                    Delivery = (string)row["Delivery"],
                    VendorCode = (string)row["VendorCode"],
                    Vendor = (string)row["Vendor"],
                    Description = (string)row["Description"],
                    Picture = (string)row["Picture"]
                });
            }
            var dtPoduct = SqlHelper.ToDataTable(lp.ToList());
            //dataGridView1.DataSource = dtPoduct;
            string tableName = "tmp_YML2Product";
            if (Program.InsertToDB)
                lblInfo.Text = $"Добавлено - {BulkСopyToDB(dtPoduct, tableName)} строки в {tableName}";
            Console.WriteLine($"select * from {tableName}");
        }

        private void CreateTable(DataTable dt, SqlConnection connection, string tTable)
        {

            string sSQL = $" IF EXISTS (select * from dbo.sysobjects where id = object_id('{tTable}')) " +
                              $" drop table {tTable}";
            var cmd = new SqlCommand(sSQL, connection);
            cmd.CommandType = CommandType.Text;
            cmd.Connection = connection;
            cmd.ExecuteNonQuery();

            sSQL = $"CREATE TABLE {tTable} (";
            for (int i = 0; i < dt.Columns.Count - 1; i++)
            {
                sSQL += $"[{dt.Columns[i].Caption}] nvarchar(MAX) NULL,";
            }
            sSQL += "[rowNom][int] IDENTITY(1, 1) NOT NULL);";

            #region MaxSQL

            /*
            var sSQL = $"CREATE TABLE {tTable} (";
            for (int i = 0; i < dt.Columns.Count - 1; i++)
            {
                sSQL += $"[{dt.Columns[i].Caption}] nvarchar(MAX) NULL,";
            }
            sSQL += "[rowNom][int] IDENTITY(1, 1) NOT NULL";
            sSQL +=
                $" CONSTRAINT [PK_{tTable}] PRIMARY KEY CLUSTERED ([rowNom] ASC) " +
                $" WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF" +
                $", ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON[PRIMARY]) ON[PRIMARY] TEXTIMAGE_ON[PRIMARY]";
            */
            #endregion

            SqlCommand createtable = new SqlCommand(sSQL, connection);
            createtable.ExecuteNonQuery();
        }
        private void chbCopyToDB_CheckedChanged(object sender, EventArgs e)
        {
            Program.InsertToDB = chbCopyToDB.Checked;
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
