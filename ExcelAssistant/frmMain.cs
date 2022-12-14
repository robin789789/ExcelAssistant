using System;
using System.IO;
using System.Windows.Forms;
using msExcel = Microsoft.Office.Interop.Excel;
using static ExcelAssistant.Util;

namespace ExcelAssistant
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
            loadingSetting();
            CenterToScreen();
            this.FormClosing += frmMain_FormClosing;
        }

        private msExcel.Application application = null;
        private msExcel.Workbook workBook = null;
        private string settingJson = string.Empty;
        private ExcelSetting excelSetting = new ExcelSetting();

        private void loadingSetting()
        {
            if (!GetJsonByPath(ExcelSettingPath, ref settingJson))
            {
                MessageBox.Show("Missing file : ExcelSetting.json.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (!string.IsNullOrEmpty(settingJson))
                excelSetting = excelSetting.JsonToObject(settingJson);
            else
                excelSetting.MakeExample();

            tbFilePath.Text = excelSetting.Setting.DefaultPath;
        }

        private void saveSetting()
        {
            settingJson = ObjectToJson(excelSetting);

            File.WriteAllText(ExcelSettingPath, settingJson);
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            excelSetting.Setting.DefaultPath = GetOpenFileName(XlsxTitle, XlsxFilter);

            if (excelSetting.Setting.DefaultPath == string.Empty)
                return;

            tbFilePath.Text = excelSetting.Setting.DefaultPath;
        }

        private void bntGenerate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(excelSetting.Setting.DefaultPath) || IsFileOpen(excelSetting.Setting.DefaultPath))
                return;

            try
            {
                getExcelApp(excelSetting.Setting.DefaultPath);

                if (excelSetting.PerformenceList.Count != excelSetting.RegionList.Count)
                {
                    throw new Exception("Data not Match.");
                }

                for (int i = 0; i < excelSetting.PerformenceList.Count; i++)
                {
                    generatePerformenceColor(excelSetting.RegionList[i], excelSetting.PerformenceList[i]);
                }

                MessageBox.Show("Done", "Information", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                exitExcelApp();

                if (excelSetting.Setting.OpenExcelAfterGenerate)
                    OpenExcelByPath(excelSetting.Setting.DefaultPath);

                if (excelSetting.Setting.CloseAfterGenerate)
                    this.Close();
            }
            catch (Exception ex)
            {
                exitExcelApp();
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void getExcelApp(string path)
        {
            application = new msExcel.Application();
            workBook = application.Workbooks.Open(path);

            FileInfo xlsAttribute = new FileInfo(path)
            {
                Attributes = FileAttributes.Normal
            };
        }

        private void exitExcelApp()
        {
            if (null == application)
                return;
            if (null == workBook)
                return;

            workBook.Save();
            workBook.Close();
            application.Quit();
        }

        private void generatePerformenceColor(RegionSheetFormat regionFormat, PerformenceSheetFormat performenceFormat)
        {
            msExcel.Worksheet regionSheet = (msExcel.Worksheet)workBook.Sheets[regionFormat.SheetName];
            msExcel.Worksheet performenceSheet = (msExcel.Worksheet)workBook.Sheets[performenceFormat.SheetName];
            msExcel.Range paintRange;
            double totalStartDays = 0, totalEndDays, totalMonth = 0;
            int color, paintStartMonth, paintEndMonth;
            string leftTop, rightBottom, currentPaintRow;
            int totalRows = regionFormat.GetRowsCount();
            for (int i = 0; i < totalRows; i++)
            {
                var tempTotalMonth = performenceSheet.get_Range(performenceFormat.RangeStart.Column + (performenceFormat.RangeStart.Row + i).ToString()).Value2;//已裝月份

                if (null == tempTotalMonth)
                    return;

                totalMonth = tempTotalMonth;

                color = GetColorByCondition(totalMonth);//選擇顏色

                var temp = regionSheet.get_Range(regionFormat.TargetStart.Column + (regionFormat.TargetStart.Row + i).ToString()).Value2;//得到日子總數，要轉Date
                if (null != temp)
                    totalStartDays = temp;
                paintStartMonth = GetDateFromExcel(totalStartDays.ToString());//轉換後，從哪一個月份開始有顏色

                var temp2 = regionSheet.get_Range(performenceFormat.CompleteStart.Column + (performenceFormat.CompleteStart.Row + i).ToString()).Value2;//得到日子總數，要轉Date
                if (null != temp2)
                {
                    totalEndDays = temp2;
                    paintEndMonth = GetDateFromExcel(totalEndDays.ToString());//驗收完成，沒有就用now
                }
                else
                {
                    paintEndMonth = DateTime.Now.Month;
                }

                if (color == White || color == Red)
                {
                    paintStartMonth = 1; paintEndMonth = 12;
                }

                //range 1-12
                //eg.5~7 (5,6,7)
                //eg.12~3 (12,1,2,3)

                currentPaintRow = (regionFormat.RangeStart.Row + i).ToString();
                leftTop = regionFormat.GetColCellByIndex(1) + currentPaintRow;
                rightBottom = regionFormat.GetColCellByIndex(12) + currentPaintRow;
                paintRange = (msExcel.Range)regionSheet.get_Range(leftTop, rightBottom);

                if (paintEndMonth > paintStartMonth)
                {
                    paintRange.Interior.Color = White;//先1~12整列塗白

                    leftTop = regionFormat.GetColCellByIndex(paintStartMonth) + currentPaintRow;
                    rightBottom = regionFormat.GetColCellByIndex(paintEndMonth) + currentPaintRow;

                    paintRange = (msExcel.Range)regionSheet.get_Range(leftTop, rightBottom);
                    paintRange.Interior.Color = color;//範圍內的月份上色
                }
                else if (paintEndMonth < paintStartMonth)
                {
                    paintRange.Interior.Color = color;//整列變色

                    if (Math.Abs(paintStartMonth - paintEndMonth) > 1)//避免互相抵消 //eg. start at 10 ,end at 9
                    {
                        leftTop = regionFormat.GetColCellByIndex(paintStartMonth - 1) + currentPaintRow; //開始月份不塗白色
                        rightBottom = regionFormat.GetColCellByIndex(paintEndMonth + 1) + currentPaintRow;//結束月份不塗白色

                        paintRange = (msExcel.Range)regionSheet.get_Range(leftTop, rightBottom);
                        paintRange.Interior.Color = White;//原區域塗白
                    }
                }
                else if (paintEndMonth == paintStartMonth)
                {
                    paintRange.Interior.Color = White;

                    if (color == Green)//    1 > 裝機總時間 > 0
                    {
                        leftTop = regionFormat.GetColCellByIndex(paintStartMonth) + currentPaintRow;
                        rightBottom = regionFormat.GetColCellByIndex(paintEndMonth) + currentPaintRow;
                    }
                    else if (color == Orange)//    12 > 裝機總時間 > 11
                    {
                        leftTop = regionFormat.GetColCellByIndex(1) + currentPaintRow;
                        rightBottom = regionFormat.GetColCellByIndex(12) + currentPaintRow;
                    }
                    paintRange = (msExcel.Range)regionSheet.get_Range(leftTop, rightBottom);
                    paintRange.Interior.Color = color;
                }
            }
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            saveSetting();
        }

    }
}
