using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.Json;
using Word = Microsoft.Office.Interop.Word;
using System.IO;


namespace Template_4337
{
    /// <summary>
    /// Логика взаимодействия для ShaykhullinaAlbina_4337.xaml
    /// </summary>
    public partial class ShaykhullinaAlbina_4337 : Window
    {
        private const int _sheetsCount = 47;
        public ShaykhullinaAlbina_4337()
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
                    entities.import.Add(new import() { FullName = list[i, 0], CodeClient = list[i, 1], BirthDay = list[i, 2], PostCode = list[i, 3], City = list[i, 4], Street = list[i, 5], Home = list[i, 6], Apartment = list[i, 7], Email = list[i, 8] });
                }
                MessageBox.Show("Данные успешно добавлены!");
                entities.SaveChanges();
            }
        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<import> clients;

            using (Entities entities = new Entities())
            {
                clients = entities.import.ToList();
            }

            List<string[]> StreetCategories = new List<string[]>() { //for sheets name
                new string[]{ " Чехова" },
                new string[]{ " Степная" },
                new string[]{ " Коммунистическая" },
                new string[]{ " Солнечная" },
                new string[]{ " Шоссейная" },
                new string[]{ " Партизанская" },
                new string[]{ " Победы" },
                new string[]{ " Молодежная" },
                new string[]{ " Новая" },
                new string[]{ " Октябрьская" },
                new string[]{ " Садовая" },
                new string[]{ " Комсомольская" },
                new string[]{ " Дзержинского" },
                new string[]{ " Набережная" },
                new string[]{ " Фрунзе" },
                new string[]{ " Школьная" },
                new string[]{ " 8 Марта" },
                new string[]{ " Зеленая" },
                new string[]{ " Маяковского" },
                new string[]{ " Светлая" },
                new string[]{ " Цветочная" },
                new string[]{ " Спортивная" },
                new string[]{ " Гоголя" },
                new string[]{ " Северная" },
                new string[]{ " Вишневая" },
                new string[]{ " Подгорная" },
                new string[]{ " Полевая" },
                new string[]{ " Клубная" },
                new string[]{ " Некрасова" },
                new string[]{ " Мичурина" },
                new string[]{ " Парковая" },
                new string[]{ " Дорожная" },
                new string[]{ " Первомайская" },
                new string[]{ " Красноармейская" },
                new string[]{ " Чкалова" },
                new string[]{ " Заводская" },
                new string[]{ " Больничная" },
                new string[]{ " Гагарина" },
                new string[]{ " Вокзальная" },
                new string[]{ " Западная" },
                new string[]{ " Механизаторов" },
                new string[]{ " Свердлова" },
                new string[]{ " Матросова" },
                new string[]{ " Красная" },
                new string[]{ " Дачная" },
                new string[]{ " Нагорная" },
                new string[]{ " Весенняя" },
            };

            var app = new Microsoft.Office.Interop.Excel.Application();
            app.SheetsInNewWorkbook = _sheetsCount;
            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < _sheetsCount; i++)
            {
                int startRowIndex = 1;
                Microsoft.Office.Interop.Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = $"Категория - {StreetCategories[i][0]}";

                Microsoft.Office.Interop.Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][1]];
                headerRange.Merge();
                headerRange.Value = $"Категория - {StreetCategories[i][0]}";
                headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                headerRange.Font.Bold = true;
                startRowIndex++;

                worksheet.Cells[1][startRowIndex] = "Код клиента";
                worksheet.Cells[2][startRowIndex] = "ФИО";
                worksheet.Cells[3][startRowIndex] = "Email";

                startRowIndex++;

                foreach (import import in clients.OrderBy(a => a.FullName))
                {
                    if (import.Street == StreetCategories[i][0])
                    {
                        worksheet.Cells[1][startRowIndex] = import.CodeClient;
                        worksheet.Cells[2][startRowIndex] = import.FullName;
                        worksheet.Cells[3][startRowIndex] = import.Email;
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

        class StreetJSON
        {
            public string FullName { get; set; }
            public string CodeClient { get; set; }
            public string BirthDate { get; set; }
            public string Index { get; set; }
            public string City { get; set; }
            public string Street { get; set; }
            public int Home { get; set; }
            public int Kvartira { get; set; }
            public string E_mail { get; set; }
            public int Id { get; set; }
        }
        private void BnImportJSON_Click(object sender, RoutedEventArgs e)
        {
            string json = File.ReadAllText(@"C:\Users\Albin\OneDrive\Рабочий стол\Импорт\3.json");
            var streets = JsonSerializer.Deserialize<List<StreetJSON>>(json);
            using (Entities entities = new Entities())
            {
                foreach (StreetJSON streetJSON in streets)
                {
                    try
                    {
                        entities.import.Add(new import()
                        {
                            CodeClient = streetJSON.CodeClient,
                            BirthDay = streetJSON.BirthDate,
                            FullName = streetJSON.FullName,
                            PostCode = streetJSON.Index,
                            City = streetJSON.City,
                            Street = streetJSON.Street,
                            Home = streetJSON.Home.ToString(),
                            Apartment = streetJSON.Kvartira.ToString(),
                            Email = streetJSON.E_mail
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                MessageBox.Show("Данные успешно давлены!");
                entities.SaveChanges();
            }
        }

        private void BnExportWord_Click(object sender, RoutedEventArgs e)
        {
            List<import> streets;

            using (Entities entities = new Entities())
            {
                streets = entities.import.ToList();
            }

            var app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document document = app.Documents.Add();

            for (int i = 0; i < _sheetsCount; i++)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph = document.Paragraphs.Add();
                Microsoft.Office.Interop.Word.Range range = paragraph.Range;

                List<string[]> StreetCategories = new List<string[]>() { //for sheets name
                    new string[]{ " Чехова" },
                    new string[]{ " Степная" },
                    new string[]{ " Коммунистическая" },
                    new string[]{ " Солнечная" },
                    new string[]{ " Шоссейная" },
                    new string[]{ " Партизанская" },
                    new string[]{ " Победы" },
                    new string[]{ " Молодежная" },
                    new string[]{ " Новая" },
                    new string[]{ " Октябрьская" },
                    new string[]{ " Садовая" },
                    new string[]{ " Комсомольская" },
                    new string[]{ " Дзержинского" },
                    new string[]{ " Набережная" },
                    new string[]{ " Фрунзе" },
                    new string[]{ " Школьная" },
                    new string[]{ " 8 Марта" },
                    new string[]{ " Зеленая" },
                    new string[]{ " Маяковского" },
                    new string[]{ " Светлая" },
                    new string[]{ " Цветочная" },
                    new string[]{ " Спортивная" },
                    new string[]{ " Гоголя" },
                    new string[]{ " Северная" },
                    new string[]{ " Вишневая" },
                    new string[]{ " Подгорная" },
                    new string[]{ " Полевая" },
                    new string[]{ " Клубная" },
                    new string[]{ " Некрасова" },
                    new string[]{ " Мичурина" },
                    new string[]{ " Парковая" },
                    new string[]{ " Дорожная" },
                    new string[]{ " Первомайская" },
                    new string[]{ " Красноармейская" },
                    new string[]{ " Чкалова" },
                    new string[]{ " Заводская" },
                    new string[]{ " Больничная" },
                    new string[]{ " Гагарина" },
                    new string[]{ " Вокзальная" },
                    new string[]{ " Западная" },
                    new string[]{ " Механизаторов" },
                    new string[]{ " Свердлова" },
                    new string[]{ " Матросова" },
                    new string[]{ " Красная" },
                    new string[]{ " Дачная" },
                    new string[]{ " Нагорная" },
                    new string[]{ " Весенняя" },
                };


                

                var data = i == 0 ? streets.Where(o => o.Street == " Чехова")
                        : i == 1 ? streets.Where(o => o.Street == " Степная")
                        : i == 2 ? streets.Where(o => o.Street == " Коммунистическая")
                        : i == 3 ? streets.Where(o => o.Street == " Солнечная")
                        : i == 4 ? streets.Where(o => o.Street == " Шоссейная")
                        : i == 5 ? streets.Where(o => o.Street == " Партизанская")
                        : i == 6 ? streets.Where(o => o.Street == " Победы")
                        : i == 7 ? streets.Where(o => o.Street == " Молодежная")
                        : i == 8 ? streets.Where(o => o.Street == " Новая")
                        : i == 9 ? streets.Where(o => o.Street == " Октябрьская")
                        : i == 10 ? streets.Where(o => o.Street == " Садовая")
                        : i == 11 ? streets.Where(o => o.Street == " Комсомольская")
                        : i == 12 ? streets.Where(o => o.Street == " Дзержинского")
                        : i == 13 ? streets.Where(o => o.Street == " Набережная")
                        : i == 14 ? streets.Where(o => o.Street == " Фрунзе")
                        : i == 15 ? streets.Where(o => o.Street == " Школьная")
                        : i == 16 ? streets.Where(o => o.Street == " 8 Марта")
                        : i == 17 ? streets.Where(o => o.Street == " Зеленая")
                        : i == 18 ? streets.Where(o => o.Street == " Маяковского")
                        : i == 19 ? streets.Where(o => o.Street == " Светлая")
                        : i == 20 ? streets.Where(o => o.Street == " Цветочная")
                        : i == 21 ? streets.Where(o => o.Street == " Спортивная")
                        : i == 22 ? streets.Where(o => o.Street == " Гоголя")
                        : i == 23 ? streets.Where(o => o.Street == " Северная")
                        : i == 24 ? streets.Where(o => o.Street == " Вишневая")
                        : i == 25 ? streets.Where(o => o.Street == " Подгорная")
                        : i == 26 ? streets.Where(o => o.Street == " Полевая")
                        : i == 27 ? streets.Where(o => o.Street == " Клубная")
                        : i == 28 ? streets.Where(o => o.Street == " Некрасова")
                        : i == 29 ? streets.Where(o => o.Street == " Мичурина")
                        : i == 30 ? streets.Where(o => o.Street == " Парковая")
                        : i == 31 ? streets.Where(o => o.Street == " Дорожная")
                        : i == 32 ? streets.Where(o => o.Street == " Первомайская")
                        : i == 33 ? streets.Where(o => o.Street == " Красноармейская")
                        : i == 34 ? streets.Where(o => o.Street == " Чкалова")
                        : i == 35 ? streets.Where(o => o.Street == " Заводская")
                        : i == 36 ? streets.Where(o => o.Street == " Больничная")
                        : i == 37 ? streets.Where(o => o.Street == " Гагарина")
                        : i == 38 ? streets.Where(o => o.Street == " Вокзальная")
                        : i == 39 ? streets.Where(o => o.Street == " Западная")
                        : i == 40 ? streets.Where(o => o.Street == " Механизаторов")
                        : i == 41 ? streets.Where(o => o.Street == " Свердлова")
                        : i == 42 ? streets.Where(o => o.Street == " Матросова")
                        : i == 43 ? streets.Where(o => o.Street == " Красная")
                        : i == 44 ? streets.Where(o => o.Street == " Дачная")
                        : i == 45 ? streets.Where(o => o.Street == " Нагорная")
                        : i == 46 ? streets.Where(o => o.Street == " Весенняя")
                        : streets; //sort for task

                //for(int y = 0; y < StreetCategories.Count; y++)
                //{
                //    if(i == y)
                //    {
                //        streets = streets.Where(b => b.Street == StreetCategories.ElementAt(y));
                //    }
                //}
                List<import> currentStreet = data.ToList();
                int countStreetInCategory = currentStreet.Count();

                Microsoft.Office.Interop.Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Microsoft.Office.Interop.Word.Range tableRange = tableParagraph.Range;
                Microsoft.Office.Interop.Word.Table strettTable = document.Tables.Add(tableRange, countStreetInCategory + 1, 3);
                strettTable.Borders.InsideLineStyle =
                strettTable.Borders.OutsideLineStyle =
                Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                strettTable.Range.Cells.VerticalAlignment =
                Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                

                Microsoft.Office.Interop.Word.Range cellRange = strettTable.Cell(1, 1).Range;
                cellRange.Text = "Код клиента";
                cellRange = strettTable.Cell(1, 2).Range;
                cellRange.Text = "ФИО";
                cellRange = strettTable.Cell(1, 3).Range;
                cellRange.Text = "Email";
                strettTable.Rows[1].Range.Bold = 1;
                strettTable.Rows[1].Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                int j = 1;
                foreach (var currentStaff in currentStreet.OrderBy(a => a.FullName))
                {
                    cellRange = strettTable.Cell(j + 1, 1).Range;
                    cellRange.Text = $"{currentStaff.CodeClient}";
                    cellRange.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = strettTable.Cell(j + 1, 2).Range;
                    cellRange.Text = currentStaff.FullName;
                    cellRange.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = strettTable.Cell(j + 1, 3).Range;
                    cellRange.Text = currentStaff.Email;
                    cellRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    j++;
                }

                for (int t = 0; t < _sheetsCount; t++)
                {
                    range.Text = Convert.ToString($"Улица - {StreetCategories[i][0]}");
                    range.InsertParagraphAfter();
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
