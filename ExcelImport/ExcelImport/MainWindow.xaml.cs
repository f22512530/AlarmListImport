using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace ExcelImport
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        List<List<ICell>> dataList = new List<List<ICell>>();
        _Core _C;

        public MainWindow()
        {
            InitializeComponent();
            _C = new _Core();
            if (!_C.IsExist("alarmlist"))
            {
                CreateDTRemind.Text = "請先建立資料表";
                CreateDTBtn.IsEnabled = true;
                fileLoad.IsEnabled = false;
            }
            else
            {
                CreateDTBtn.IsEnabled = false;
                fileLoad.IsEnabled = true;
                CreateDTRemind.Text = "資料表已存在";
            }
        }
        
        private void CreateDTBtn_Click(object sender, RoutedEventArgs e)
        {
            _C.CreateTable("alarmlist");
            CreateDTRemind.Text = "資料表建立成功";
        }

        private void FileImport_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Title = "開啟檔案";
            dialog.Filter = "xlsx file (*.*)|*.xlsx";
            if (dialog.ShowDialog() == true)
            {
                dataList.Clear();
                dataPath.Text = dialog.FileName;
                using (FileStream stream = new FileStream(dialog.FileName, FileMode.Open))
                {
                    XSSFWorkbook workbook = new XSSFWorkbook(stream);
                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        string sheetName = workbook.GetSheetName(i);
                        XSSFSheet sheet = (XSSFSheet)workbook.GetSheetAt(i);
                        for (int j = 1; j <= sheet.LastRowNum; j++)
                        {
                            IRow headerRow = sheet.GetRow(j);
                            Tuple<string, string> engResult = alarmMsgSplit(headerRow.Cells[0].StringCellValue);
                            Tuple<string, string> chtResult = alarmMsgSplit(headerRow.Cells[1].StringCellValue);
                            bool IsSend = headerRow.Cells[2].StringCellValue == "Y" ? true : false;
                            _C.InsertRow(engResult.Item1, chtResult.Item2, engResult.Item2, IsSend, sheetName);
                            InsertOverflow.Text += $"{engResult.Item1}, {chtResult.Item2}, {engResult.Item2}, {IsSend}, {sheetName} \r\n";
                        }
                    }
                    workbook.Close();
                    InsertOverflow.Text += "Finish!!";
                }
            }
        }

        private Tuple<string, string> alarmMsgSplit(string str)
        {
            Tuple<string, string> tuple = null;
            try
            {
                string[] splitString = Regex.Split(str, ",");
                if(splitString.Length == 4)
                {
                    splitString[3] = splitString[3].Remove(0, splitString[2].Length).Trim();
                }
                else if(splitString.Length > 4)
                {
                    splitString[3] = splitString[3].Remove(0, splitString[2].Length).Trim();

                    int index = splitString.Length - 3;
                    string[] msg = new string[index];
                    Array.Copy(splitString, 3, msg, 0, splitString.Length - 3);
                    splitString[3] = string.Join(",", msg);
                }
                else if(splitString.Length < 4)
                {
                    tuple = new Tuple<string, string>(string.Empty, string.Empty);
                    return tuple;
                }
                tuple = new Tuple<string, string>(splitString[2], splitString[3]);
            }
            catch (Exception e)
            {
                MessageBox.Show($"alarmMsgSplit error : {e.ToString()}");
            }
            return tuple;
        }
    }
}
