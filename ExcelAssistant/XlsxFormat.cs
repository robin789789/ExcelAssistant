using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace ExcelAssistant
{
    public class ExcelSetting
    {
        public AppSetting Setting = new AppSetting();

        public List<RegionSheetFormat> RegionList { get; set; } = new List<RegionSheetFormat>();

        public List<PerformenceSheetFormat> PerformenceList { get; set; } = new List<PerformenceSheetFormat>();

        public ExcelSetting JsonToObject(string json)
        {
            try
            {
                return JsonConvert.DeserializeObject<ExcelSetting>(json, new JsonSerializerSettings
                {
                    MissingMemberHandling = MissingMemberHandling.Ignore
                });
            }
            catch (JsonSerializationException ex)
            {
                MessageBox.Show(ex.Message);
            }
            var temp = new ExcelSetting();
            temp.MakeExample();
            return temp;
        }

        public void MakeExample()
        {
            RegionList.Add(new RegionSheetFormat("崑山", "K", 5, "V", 116)
            {
                TargetStart = new Cell("G", 5),
                TargetEnd = new Cell("G", 116)
            });

            PerformenceList.Add(new PerformenceSheetFormat("崑山績效", "W", 5, "W", 116)
            {
                CompleteStart = new Cell("X", 5),
                CompleteEnd = new Cell("X", 116),
            });

            Setting.OpenExcelAfterGenerate = true;
            Setting.CloseAfterGenerate = true;
        }

        public class AppSetting
        {
            public string DefaultPath { get; set; } = String.Empty;

            public bool OpenExcelAfterGenerate { get; set; } = false;

            public bool CloseAfterGenerate { get; set; } = false;
        }
    }

    public class XlsxFormat
    {
        public XlsxFormat(string name, string startCol, int startRow, string endCol, int endRow)
        {
            SheetName = name;
            RangeStart = new Cell(startCol, startRow);
            RangeEnd = new Cell(endCol, endRow);
        }

        public string SheetName { get; set; }

        public Cell RangeStart { get; set; }

        public Cell RangeEnd { get; set; }

        public string GetRangeStartCell()
        {
            return RangeStart.ToString();
        }

        public string GetRangeEndCell()
        {
            return RangeEnd.ToString();
        }

        public string GetColCellByIndex(int idx)
        {
            int startFormZero = 1;
            int num = Util.GetNumberFromExcelColumn(RangeStart.Column);
            return Util.GetExcelColumnName(num + idx - startFormZero);
        }

        public int GetRowsCount()
        {
            return RangeEnd.Row - RangeStart.Row;
        }

    }

    public class RegionSheetFormat : XlsxFormat
    {
        public RegionSheetFormat(string name, string startCol, int startRow, string endCol, int endRow) : base(name, startCol, startRow, endCol, endRow)
        {
        }

        public Cell TargetStart { get; set; }

        public Cell TargetEnd { get; set; }

    }

    public class PerformenceSheetFormat : XlsxFormat
    {
        public PerformenceSheetFormat(string name, string startCol, int startRow, string endCol, int endRow) : base(name, startCol, startRow, endCol, endRow)
        {
        }

        public Cell CompleteStart { get; set; }

        public Cell CompleteEnd { get; set; }
    }

    public class Cell
    {
        public Cell(string col, int row)
        {
            Column = col;
            Row = row;
        }
        public string Column { get; set; }
        public int Row { get; set; }
        public string DotFormat { get; set; } = "0.000";

        public override string ToString()
        {
            return Column + Row.ToString();
        }

    }

    public static class Util
    {
        public const string xlsxFilter = "Excel Work|*.xlsx";
        public const string xlsxTitle = "Open 績效.xlsx File";
        public const string excelSettingPath = @".\ExcelSetting.json";

        public static int Green = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
        public static int Yellow = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
        public static int Orange = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
        public static int Red = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
        public static int White = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

        public static string GetOpenFileName(string title, string filter)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = filter,
                FilterIndex = 2,
                RestoreDirectory = true,
                Title = title
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog.FileName;
            }
            return "";
        }

        public static int GetDateFromExcel(string dateValue)
        {
            return DateTime.FromOADate(Convert.ToUInt32(dateValue)).Month;
        }

        public static int GetNumberFromExcelColumn(string column)
        {
            int retVal = 0;
            string col = column.ToUpper();
            for (int iChar = col.Length - 1; iChar >= 0; iChar--)
            {
                char colPiece = col[iChar];
                int colNum = colPiece - 64;
                retVal = retVal + colNum * (int)Math.Pow(26, col.Length - (iChar + 1));
            }
            return retVal;
        }

        public static string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }

        public static int GetColorByCondition(double target)
        {
            if (target > 0 && target < 6)
            {
                return Green;
            }
            else if (target >= 6 && target < 9)
            {
                return Yellow;
            }
            else if (target >= 9 && target < 12)
            {
                return Orange;
            }
            else if (target > 12)
            {
                return Red;
            }
            else
            {
                return White;
            }
        }

        public static bool GetJsonByPath(string path, ref string result)
        {
            if (File.Exists(path))
            {
                StreamReader streamReader = new StreamReader(path, new UTF8Encoding());
                result = streamReader.ReadToEnd().ToString();
                streamReader.Close();
                return true;
            }
            else MessageBox.Show($"{path} does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }

        public static string ObjectToJson(dynamic dynamic)
        {
            return JsonConvert.SerializeObject(dynamic, Formatting.Indented);
        }

        public static void OpenExcelByPath(string path)
        {
            System.Diagnostics.Process.Start(path);
        }

        public static bool isFileOpen(string filename)
        {
            bool isOpen = true;
            while (isOpen)
            {
                try
                {
                    System.IO.FileStream stream = System.IO.File.OpenWrite(filename);
                    stream.Close();
                    isOpen = false;
                }
                catch
                {
                    var result = MessageBox.Show("檔案目前處於開啟狀態，請關閉後重試。", "Warning", MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);

                    if (result == DialogResult.Cancel)
                        return true;
                }
            }
            return false;
        }
    }
}
