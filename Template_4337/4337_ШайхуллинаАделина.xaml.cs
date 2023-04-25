using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.Json;
using Word = Microsoft.Office.Interop.Word;

namespace Template_4337
{
    /// <summary>
    /// Логика взаимодействия для _4337_ШайхуллинаАделина.xaml
    /// </summary>
    public partial class _4337_ШайхуллинаАделина : Window
    {
        private const int _sheetsCount = 3;
        public _4337_ШайхуллинаАделина()
        {
            InitializeComponent();
        }
        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
            {
                return;
            }

            string[,] list; //for data in excel
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (Entities entities = new Entities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    entities.import.Add(new import() { Role = list[i, 1], FIO = list[i, 2], E_mail = list[i, 3], Password = list[i, 4], LastEntry = DateTime.Parse(list[i, 5]), TypeEntry = list[i, 6] });
                }
                MessageBox.Show("Успешно!");
                entities.SaveChanges();
            }
        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<import> staffs;

            using (Entities entities = new Entities())
            {
                staffs = entities.import.ToList();
            }

            List<string[]> RoleCategories = new List<string[]>() { //for sheets name
                new string[]{ "Администратор" },
                new string[]{ "Старший смены" },
                new string[]{ "Продавец" },
            };

            var app = new Microsoft.Office.Interop.Excel.Application();
            app.SheetsInNewWorkbook = _sheetsCount;
            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < _sheetsCount; i++)
            {
                int startRowIndex = 1;
                Microsoft.Office.Interop.Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = $"Категория - {RoleCategories[i][0]}";

                Microsoft.Office.Interop.Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][1]];
                headerRange.Merge();
                headerRange.Value = $"Категория - {RoleCategories[i][0]}";
                headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                headerRange.Font.Bold = true;
                startRowIndex++;

                worksheet.Cells[1][startRowIndex] = "Код клиента";
                worksheet.Cells[2][startRowIndex] = "ФИО";
                worksheet.Cells[3][startRowIndex] = "Логин";

                startRowIndex++;

                foreach (import import in staffs)
                {
                    if (import.Role == RoleCategories[i][0])
                    {
                        worksheet.Cells[1][startRowIndex] = import.IdClient;
                        worksheet.Cells[2][startRowIndex] = import.FIO;
                        worksheet.Cells[3][startRowIndex] = import.E_mail;
                        startRowIndex++;
                    }
                }

                Microsoft.Office.Interop.Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][startRowIndex - 1]];
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        }

        class StaffJSON
        {
            public int Id { get; set; }
            public string CodeStaff { get; set; }
            public string Position { get; set; }
            public string FullName { get; set; }
            public string Log { get; set; }
            public string Password { get; set; }
            public string LastEnter { get; set; }
            public string TypeEnter { get; set; }
        }
        private void BnImportJSON_Click(object sender, RoutedEventArgs e)
        {
            string json = File.ReadAllText(@"C:\Users\ashai\Desktop\Импорт\4.json");
            var staffs = JsonSerializer.Deserialize<List<StaffJSON>>(json);
            using (Entities entities = new Entities())
            {
                foreach(StaffJSON staffJSON in staffs)
                {
                    try
                    {
                        entities.import3.Add(new import3() 
                        { 
                            Id = Convert.ToInt32(staffJSON.Id),
                            CodeStaff = staffJSON.CodeStaff,
                            Position = staffJSON.Position,
                            FullName = staffJSON.FullName,
                            Log = staffJSON.Log,
                            Password = staffJSON.Password,
                            LastEnter = staffJSON.LastEnter,
                            TypeEnter = staffJSON.TypeEnter
                        });

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                MessageBox.Show("Успешно!");
                entities.SaveChanges();
            }
        }

        private void BnExportWord_Click(object sender, RoutedEventArgs e)
        {
            List<import3> staffs;

            using (Entities entities = new Entities())
            {
                staffs = entities.import3.ToList();
            }

            var app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document document = app.Documents.Add();

            for (int i = 0; i < _sheetsCount; i++)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph = document.Paragraphs.Add();
                Microsoft.Office.Interop.Word.Range range = paragraph.Range;

                List<string[]> RoleCategories = new List<string[]>() { //for sheets name
                    new string[]{ "Администратор" },
                    new string[]{ "Старший смены" },
                    new string[]{ "Продавец" },
                };

                var data = i == 0 ? staffs.Where(o => o.Position == "Администратор")
                        : i == 1 ? staffs.Where(o => o.Position == "Старший смены")
                        : i == 2 ? staffs.Where(o => o.Position == "Продавец") : staffs; //sort for task
                List<import3> currentStaffs = data.ToList();
                int countStaffsInCategory = currentStaffs.Count();

                Microsoft.Office.Interop.Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Microsoft.Office.Interop.Word.Range tableRange = tableParagraph.Range;
                Microsoft.Office.Interop.Word.Table staffsTable = document.Tables.Add(tableRange, countStaffsInCategory + 1, 3);
                staffsTable.Borders.InsideLineStyle =
                staffsTable.Borders.OutsideLineStyle =
                Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                staffsTable.Range.Cells.VerticalAlignment =
                Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                range.Text = Convert.ToString($"Категория - {RoleCategories[i][0]}");
                range.InsertParagraphAfter();

                Microsoft.Office.Interop.Word.Range cellRange = staffsTable.Cell(1, 1).Range;
                cellRange.Text = "Код сотрудника";
                cellRange = staffsTable.Cell(1, 2).Range;
                cellRange.Text = "ФИО";
                cellRange = staffsTable.Cell(1, 3).Range;
                cellRange.Text = "Логин";
                staffsTable.Rows[1].Range.Bold = 1;
                staffsTable.Rows[1].Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                int j = 1;
                foreach (var currentStaff in currentStaffs)
                {
                    cellRange = staffsTable.Cell(j + 1, 1).Range;
                    cellRange.Text = $"{currentStaff.CodeStaff}";
                    cellRange.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = staffsTable.Cell(j + 1, 2).Range;
                    cellRange.Text = currentStaff.FullName;
                    cellRange.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = staffsTable.Cell(j + 1, 3).Range;
                    cellRange.Text = currentStaff.Log;
                    cellRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    j++;
                }
                
                if (i > 0)
                {
                    range.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);
                }
            }
            app.Visible = true;
        }
    }
}
