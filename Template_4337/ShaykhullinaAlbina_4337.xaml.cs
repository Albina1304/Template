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

        private void BnImportJSON_Click(object sender, RoutedEventArgs e)
        {

        }

        private void BnExportWord_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
