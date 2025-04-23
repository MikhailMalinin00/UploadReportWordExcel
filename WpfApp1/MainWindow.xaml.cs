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
using System.Windows.Navigation;
using System.Windows.Shapes;

using System.Windows.Forms.DataVisualization.Charting;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Entities _context = new Entities();
        public MainWindow()
        {
            InitializeComponent();
            ChartPayments.ChartAreas.Add(new ChartArea("Main"));
            var currentSeries = new Series("Платежи")
            {
                IsValueShownAsLabel = true
            };
            ChartPayments.Series.Add(currentSeries);

            CmbUser.ItemsSource = _context.User.ToList();
            CmbDiagram.ItemsSource = Enum.GetValues(typeof(SeriesChartType));
        }

        private void UpdateChart(object sender, SelectionChangedEventArgs e)
        {
            if(CmbUser.SelectedItem is User currentUser && CmbDiagram.SelectedItem is SeriesChartType currentType)
            {
                Series currentSeries = ChartPayments.Series.FirstOrDefault();
                currentSeries.ChartType = currentType;
                currentSeries.Points.Clear();

                var categoriesList = _context.Category.ToList();
                foreach(var category in categoriesList)
                {
                    currentSeries.Points.AddXY(category.Name, _context.Payment.ToList().Where(u => u.User == currentUser && u.Category == category).Sum(u => u.Price));
                }
            }
        }

        private void buttonExcel_Click(object sender, RoutedEventArgs e)
        {
            // Получаем список пользователей с одновременной сортировкой по ФИО
            var allUsers = _context.User.ToList().OrderBy(u => u.FIO).ToList();

            // Создаем новую книгу Excel, указывая количество листов (sheets) равным количеству пользователей в базе данных, и добавляем рабочую книгу (workbook)
            var application = new Excel.Application();
            application.SheetsInNewWorkbook = allUsers.Count();
            Excel.Workbook workbook = application.Workbooks.Add(Type.Missing);

            // Запускаем цикл по пользователям
            for (int i = 0; i < allUsers.Count(); i++)
            {
                // Устанавливаем счетчик строк и называем листы рабочей книги ExcelУстанавливаем счетчик строк и называем листы рабочей книги Excel
                int startRowIndex = 1;
                Excel.Worksheet worksheet = application.Worksheets.Item[i + 1];
                worksheet.Name = allUsers[i].FIO;

                // Добавляем названия колонок и форматируем их
                worksheet.Cells[1][startRowIndex] = "Дата платежа";
                worksheet.Cells[2][startRowIndex] = "Название";
                worksheet.Cells[3][startRowIndex] = "Стоимость";
                worksheet.Cells[4][startRowIndex] = "Количество";
                worksheet.Cells[5][startRowIndex] = "Сумма";
                Excel.Range columlHeaderRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][1]];
                columlHeaderRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                columlHeaderRange.Font.Bold = true;
                startRowIndex++;

                // Группируем платежи текущего пользователя и сортируемм по дате и названию категории
                var userCategories = allUsers[i].Payment.OrderBy(u => u.Date).GroupBy(u => u.Category).OrderBy(u => u.Key.Name);

                // Вложенный цикл по категориям платежей
                foreach (var groupCategory in userCategories)
                {
                    Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[5][startRowIndex]];
                    headerRange.Merge();
                    headerRange.Value = groupCategory.Key.Name;
                    headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    headerRange.Font.Italic = true;
                    startRowIndex++;

                    // Вложенный цикл по платежам, расчет величины платежа по каждой категории
                    foreach (var payment in groupCategory)
                    {
                        worksheet.Cells[1][startRowIndex] = string.Format("{0:dd.MM.yyyy}", payment.Date);
                        worksheet.Cells[2][startRowIndex] = payment.Name;
                        worksheet.Cells[3][startRowIndex] = payment.Price;
                        (worksheet.Cells[3][startRowIndex] as Excel.Range).NumberFormat = "0.00";
                        worksheet.Cells[4][startRowIndex] = payment.Num;
                        worksheet.Cells[5][startRowIndex].Formula = $"=C{startRowIndex}*D{startRowIndex}";
                        (worksheet.Cells[5][startRowIndex] as Excel.Range).NumberFormat = "0.00";
                        startRowIndex++;
                    } // Завершение цикла по платежам

                    // Добавляем название к ячейкам и форматируем их
                    Excel.Range sumRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[4][startRowIndex]];
                    sumRange.Merge();
                    sumRange.Value = "ИТОГО:";
                    sumRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                    // Рассчитываем величину общих затрат и форматируем ячейку
                    worksheet.Cells[5][startRowIndex].Formula = $"=SUM(E{startRowIndex - groupCategory.Count()}:" + $"E{startRowIndex - 1})";
                    sumRange.Font.Bold = worksheet.Cells[5][startRowIndex].Font.Bold = true;
                    startRowIndex++;

                    // Добавляем границы таблицы платежей (внешние и внутренние)
                    Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][startRowIndex - 1]]; 
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = 
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = 
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = 
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = 
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = 
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = 
                        Excel.XlLineStyle.xlContinuous;

                    // Устанавливаем автоширину всех столбцов листа
                    worksheet.Columns.AutoFit();
                } // Завершение цикла по категориям платежей

                // Разрешаем отобразить таблицу по завершении экспорта
                application.Visible = true;
            } // Завершение цикла по пользователям
        }

        private void buttonWord_Click(object sender, RoutedEventArgs e)
        {
            // Получаем список пользователей и категорий из базы данных
            var allUsers = _context.User.ToList();
            var allCategories = _context.Category.ToList();

            // Создаем новый документ Word
            var application = new Word.Application();
            Word.Document document = application.Documents.Add();

            // Запускаем цикл по пользователям
            foreach(var user in allUsers)
            {
                // Внутри цикла создаем параграф для хранения названий страниц и добавляем названия страниц
                Word.Paragraph userParagraph = document.Paragraphs.Add();
                Word.Range userRange = userParagraph.Range;
                userRange.Text = user.FIO;
                userParagraph.set_Style("Заголовок");
                userRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                userRange.InsertParagraphAfter();
                document.Paragraphs.Add();

                // Добавляем новый параграф для таблицы с платежами, создаем и форматируем саму таблицу
                Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Word.Range tableRange = tableParagraph.Range;
                Word.Table paymentsTable = document.Tables.Add(tableRange, allCategories.Count() + 1, 2);
                paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                // Добавляем названия колонок и их форматирование
                Word.Range cellRange;

                cellRange = paymentsTable.Cell(1, 1).Range;
                cellRange.Text = "Категория";
                cellRange = paymentsTable.Cell(1, 2).Range;
                cellRange.Text = "Сумма расходов";

                paymentsTable.Rows[1].Range.Font.Name = "Times New Roman";
                paymentsTable.Rows[1].Range.Font.Size = 14;
                paymentsTable.Rows[1].Range.Bold = 1;
                paymentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                // Вложенный цикл по строкам таблицы
                for (int i = 0; i < allCategories.Count(); i++)
                {
                    var currentCategory = allCategories[i];
                    cellRange = paymentsTable.Cell(i + 2, 1).Range;
                    cellRange.Text = currentCategory.Name;
                    cellRange.Font.Name = "Times New Roman";
                    cellRange.Font.Size = 12;

                    cellRange = paymentsTable.Cell(i + 2, 2).Range;
                    cellRange.Text = string.Format("{0:N2}", user.Payment.ToList().Where(u => u.Category == currentCategory).Sum(u => u.Num * u.Price)) + " руб.";
                    cellRange.Font.Name = "Times New Roman";
                    cellRange.Font.Size = 12;
                } // Завершение цикла по строкам таблицы
                document.Paragraphs.Add(); // Пустая строка

                // Для каждого пользователя добавляем максимальную величину платежакаждого пользователя добавляем максимальную величину платежа
                Payment maxPayment = user.Payment.OrderByDescending(u => u.Price * u.Num).FirstOrDefault();
                if (maxPayment != null)
                {
                    Word.Paragraph maxPaymentParagraph = document.Paragraphs.Add();
                    Word.Range maxPaymentRange = maxPaymentParagraph.Range;
                    maxPaymentRange.Text = $"Самый дорогостоящий платеж - {maxPayment.Name} за {maxPayment.Price * maxPayment.Num:N2} руб. от {maxPayment.Date:dd.MM.yyyy}";
                    maxPaymentParagraph.set_Style("Подзаголовок");
                    maxPaymentRange.Font.Color = Word.WdColor.wdColorDarkRed;
                    maxPaymentRange.InsertParagraphAfter();
                }
                document.Paragraphs.Add(); // Пустая строка

                // Аналогично добавляем минимальную величину платежа 
                Payment minPayment = user.Payment.OrderBy(u => u.Price * u.Num).FirstOrDefault();
                if (maxPayment != null)
                {
                    Word.Paragraph minPaymentParagraph = document.Paragraphs.Add();
                    Word.Range minPaymentRange = minPaymentParagraph.Range;
                    minPaymentRange.Text = $"Самый дешевый платеж - {minPayment.Name} за {string.Format("{0:N2}", minPayment.Price * minPayment.Num)} руб. от {minPayment.Date:dd.MM.yyyy}";
                    minPaymentParagraph.set_Style("Подзаголовок");
                    minPaymentRange.Font.Color = Word.WdColor.wdColorDarkGreen;
                    minPaymentRange.InsertParagraphAfter();
                }

                // Добавляем разрыв страницы в документ 
                if (user != allUsers.LastOrDefault()) document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);

                // Разрешаем отображение таблицы по завершении экспорта
                application.Visible = true;

                // Сохраняем документ в формате .docx и .pdf и завершаем цикл по пользователям
                document.SaveAs2(@"C:\Users\Mikhail\Downloads\Payment.docx"); 
                document.SaveAs2(@"C:\Users\Mikhail\Downloads\Payments.pdf", Word.WdExportFormat.wdExportFormatPDF);
            }
        }
    }
}
