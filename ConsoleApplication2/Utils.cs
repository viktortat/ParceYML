﻿using System;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace UtilsTbn
{
    public static class Utils
    {

            /*
             public static DataTable GetDTNew(String pathExcelFile, int StartRow, string SheetName)
            {
                if (!File.Exists(pathExcelFile))
                {
                    return null;
                }

                var DT = new DataTable();

                var notOpenFile = true;
                do
                {
                    try
                    {
                        using (var p = new ExcelPackage(new FileInfo(pathExcelFile)))
                        {
                            if (p.Workbook != null)
                            {
                                var ws = p.Workbook.Worksheets[SheetName]; //First();

                                var rowCount = GetLastUsedRow(ws);
                                var сolCount = ws.Dimension.End.Column;

                                DT.Columns.Add("cid", typeof(int));
                                for (var col = 1; col < сolCount + 1; col++)
                                {
                                    var Column = new DataColumn();
                                    Column.DataType = typeof(string);
                                    Column.ColumnName = (string)ws.Cells[1, col].Value ?? (string)ws.Cells[2, col].Value;
                                    DT.Columns.Add(Column);
                                }

                                for (var i = StartRow; i <= rowCount; i++)
                                {
                                    var dr = DT.Rows.Add();
                                    dr["cid"] = i;
                                    for (var k = 1; k < сolCount + 1; k++)
                                    {
                                        string error = null;
                                        var cellVal = ws.Cells[i, k].Value;
                                        if (cellVal != null)
                                        {
                                            //&& string.IsNullOrWhiteSpace(cellVal)
                                            //dr[k] = cellVal.ToString().Replace(Environment.NewLine, " ").Trim();
                                            dr[k] = UtilsTbn.ReplSpaces(cellVal.ToString().Trim());
                                        }
                                    }
                                }
                                notOpenFile = false;
                            }
                        }
                    }
                    catch (IOException exception)
                    {
                        MessageBox.Show($"закройте документ {pathExcelFile} и нажмите OK");
                    }
                } while (notOpenFile);
                return DT;
            }
            */
        public static DataTable GetDTNew2(ExcelWorksheet ws)
        {
            var StartRow = 3;
            var DT = new DataTable();
            using (ws)
            {

                var rowCount = GetLastUsedRow(ws);
                var сolCount = ws.Dimension.End.Column;

                DT.Columns.Add("cid", typeof(int));
                for (var col = 1; col < сolCount + 1; col++)
                {
                    var column = new DataColumn();
                    column.DataType = typeof(string);
                    column.ColumnName = (string)ws.Cells[1, col].Value ?? (string)ws.Cells[2, col].Value;
                    DT.Columns.Add(column);
                }

                for (var i = StartRow; i <= rowCount; i++)
                {
                    var dr = DT.Rows.Add();
                    dr["cid"] = i;
                    for (var k = 1; k < сolCount + 1; k++)
                    {
                        var cellVal = ws.Cells[i, k].Value;
                        if (cellVal != null)
                        {
                            dr[k] = ReplSpaces(cellVal.ToString().Trim());
                        }
                    }
                }


            }
            return DT;
        }

        public static int GetLastUsedRow(ExcelWorksheet sheet)
        {
            var row = sheet.Dimension.End.Row;
            var col = sheet.Dimension.End.Column;
            object[,] valueArray = sheet.Cells.GetValue<object[,]>();

            while (row >= 1)
            {
                var cellValue = "";
                for (int i = 1; i <= col; i++)
                {
                    cellValue += (valueArray[row - 1, i - 1] ?? string.Empty).ToString();
                }

                if (cellValue.Replace("0", "").Length > 2)
                {
                    break;
                }
                row--;
            }

            return row;
        }

        public static string ReplSpaces(string text)
        {
            text = Regex.Replace(text, @"\s+", " ");
            return text.Trim();

            #region "По старорому"
            /*
            string s = "";
            string[] whitespaceChars = new string[] {
                char.ConvertFromUtf32(9),
                char.ConvertFromUtf32(10),
                char.ConvertFromUtf32(11),
                char.ConvertFromUtf32(12),
                char.ConvertFromUtf32(13),
                char.ConvertFromUtf32(32),
                char.ConvertFromUtf32(133),
                char.ConvertFromUtf32(160),
                char.ConvertFromUtf32(5760),
                char.ConvertFromUtf32(8192),
                char.ConvertFromUtf32(8193),
                char.ConvertFromUtf32(8194),
                char.ConvertFromUtf32(8195),
                char.ConvertFromUtf32(8196),
                char.ConvertFromUtf32(8197),
                char.ConvertFromUtf32(8198),
                char.ConvertFromUtf32(8199),
                char.ConvertFromUtf32(8200),
                char.ConvertFromUtf32(8201),
                char.ConvertFromUtf32(8202),
                char.ConvertFromUtf32(8203),
                char.ConvertFromUtf32(8232),
                char.ConvertFromUtf32(8233),
                char.ConvertFromUtf32(12288),
                char.ConvertFromUtf32(65279)};
            StringBuilder sb = new StringBuilder(text);
            foreach (var wc in whitespaceChars)
            {
                sb.Replace(wc, " ");
                s = sb.ToString();
            }
            return s;
            //return text.Replace(char.ConvertFromUtf32(160), "");
            */
            #endregion
        }
    }
}
