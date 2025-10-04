using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Data.Entity;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            DG.ItemsSource = PaymentsBaseEntities.GetContext().Payment.ToList();

            cbFIO.ItemsSource = PaymentsBaseEntities.GetContext().User.Select(x => x.FIO).ToList();
            cbCat.ItemsSource = PaymentsBaseEntities.GetContext().Category.Select(x => x.Name).ToList();

            UpdateStats();
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            AddEditWindow addEditWindow = new AddEditWindow(null);
            addEditWindow.ShowDialog();
            RefreshData();
        }

        private void Edit_OnClick(object sender, RoutedEventArgs e)
        {
            AddEditWindow addEditWindow = new AddEditWindow((sender as Button).DataContext as Payment);
            addEditWindow.ShowDialog();
            RefreshData();
        }

        private void Delete_OnClick(object sender, RoutedEventArgs e)
        {
            var paymentForDel = DG.SelectedItems.Cast<Payment>().ToList();

            if (MessageBox.Show($"Вы точно хотите удалить следующие элементы ({paymentForDel.Count()})",
                    "Внимание!", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    foreach (var p in paymentForDel)
                    {
                        PaymentsBaseEntities.GetContext().Payment.Remove(p);
                    }
                    PaymentsBaseEntities.GetContext().SaveChanges();
                    MessageBox.Show("Данные удалены!");

                    RefreshData();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    Console.WriteLine(ex.StackTrace);
                }
            }
        }

        private void applyFilters_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyAllFilters();
        }

        private void UpdateStats()
        {
            var filteredPayments = DG.Items.OfType<Payment>().ToList();

            itemsCount.Text = "Выбрано " + filteredPayments.Count +
                              " из " + PaymentsBaseEntities.GetContext().Payment.Count();

            decimal sum = filteredPayments.Sum(x => x.Summ);
            itemsSum.Text = "Сумма выбранных платежей: " + sum.ToString("0,00");
        }

        private void ApplyAllFilters()
        {
            var context = PaymentsBaseEntities.GetContext();

            // Загружаем ВСЕ данные с навигационными свойствами
            var allPayments = context.Payment
                .Include(p => p.User)
                .Include(p => p.Category)
                .ToList(); // Сначала загружаем в память

            // Теперь фильтруем в памяти
            var filtered = allPayments.AsEnumerable();

            if (cbFIO.SelectedValue != null)
            {
                string fio = cbFIO.SelectedValue.ToString();
                filtered = filtered.Where(x => x.User != null && x.User.FIO == fio);
            }

            if (cbCat.SelectedValue != null)
            {
                string cat = cbCat.SelectedValue.ToString();
                filtered = filtered.Where(x => x.Category != null && x.Category.Name == cat);
            }

            if (fromDate.SelectedDate != null)
            {
                var date = fromDate.SelectedDate.Value;
                filtered = filtered.Where(x => x.Date >= date);
            }

            if (toDate.SelectedDate != null)
            {
                var date = toDate.SelectedDate.Value;
                filtered = filtered.Where(x => x.Date <= date);
            }

            DG.ItemsSource = filtered.ToList();
            UpdateStats();
        }

        private void RefreshData()
        {
            cbFIO.SelectedValue = null;
            cbCat.SelectedValue = null;
            fromDate.SelectedDate = null;
            toDate.SelectedDate = null;

            DG.ItemsSource = PaymentsBaseEntities.GetContext().Payment
                .Include(p => p.User)
                .Include(p => p.Category)
                .ToList();

            UpdateStats();
        }

        private void clear_onClick(object sender, RoutedEventArgs e)
        {
            RefreshData();
        }

        private void FromDate_OnSelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyAllFilters();
        }

        private void excelImport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Excel.Application app = new Excel.Application
                {
                    Visible = true,
                    SheetsInNewWorkbook = 1
                };
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                app.DisplayAlerts = false;
                Excel.Worksheet sheet = (Excel.Worksheet)app.Worksheets.get_Item(1);
                sheet.Name = "Платежи";

                // Заголовок
                sheet.Cells[1, 2] = "Платежи";
                sheet.Cells[1, 3] = DateTime.Now.ToString("dd.MM.yyyy");

                Excel.Range headerRange = sheet.Range["B1:C1"];
                headerRange.Font.Bold = true;
                headerRange.Font.Size = 12;
                headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int currentRow = 2;

                var payments = DG.Items.OfType<Payment>().ToList();

                var groupedPayments = payments
                    .Where(p => p.User != null && p.Category != null)
                    .GroupBy(p => p.User.FIO)
                    .OrderBy(g => g.Key);

                foreach (var userGroup in groupedPayments)
                {
                    // Строка с именем пользователя
                    sheet.Cells[currentRow, 1] = userGroup.Key;
                    Excel.Range userNameCell = sheet.Range[$"A{currentRow}"];
                    userNameCell.Font.Bold = true;
                    userNameCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    // Пустые ячейки для единообразия
                    Excel.Range emptyRange = sheet.Range[$"B{currentRow}:C{currentRow}"];
                    emptyRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    currentRow++;

                    var categoryGroups = userGroup
                        .Where(p => p.Category != null)
                        .GroupBy(p => p.Category.Name)
                        .OrderBy(g => g.Key);

                    foreach (var categoryGroup in categoryGroups)
                    {
                        decimal categorySum = categoryGroup.Sum(p => p.Summ);

                        // Категория и сумма
                        sheet.Cells[currentRow, 2] = categoryGroup.Key;
                        sheet.Cells[currentRow, 3] = categorySum;

                        Excel.Range categoryRange = sheet.Range[$"A{currentRow}:C{currentRow}"];
                        categoryRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                        currentRow++;
                    }

                    // Строка "Итого:"
                    sheet.Cells[currentRow, 2] = "Итого:";
                    decimal userTotal = userGroup.Sum(p => p.Summ);
                    sheet.Cells[currentRow, 3] = userTotal;

                    Excel.Range totalRange = sheet.Range[$"A{currentRow}:C{currentRow}"];
                    totalRange.Font.Bold = true;
                    totalRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    currentRow++;
                }

                // Автоподбор ширины колонок
                sheet.Columns["A:A"].ColumnWidth = 20;
                sheet.Columns["B:B"].ColumnWidth = 30;
                sheet.Columns["C:C"].ColumnWidth = 15;

                // Форматирование чисел
                Excel.Range amountRange = sheet.Range[$"C2:C{currentRow}"];
                amountRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                amountRange.NumberFormat = "0,00";

                MessageBox.Show("Данные успешно экспортированы в Excel!", "Экспорт завершен",
                               MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Excel: {ex.Message}", "Ошибка",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}