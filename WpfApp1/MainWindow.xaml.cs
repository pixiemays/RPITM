using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private PaymentsBaseEntities _context = new PaymentsBaseEntities();
        
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
        }

        private void Edit_OnClick(object sender, RoutedEventArgs e)
        {
            AddEditWindow addEditWindow = new AddEditWindow((sender as Button).DataContext as Payment);
            addEditWindow.ShowDialog();
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

                    DG.ItemsSource = PaymentsBaseEntities.GetContext().Payment.ToList();
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
            var context = PaymentsBaseEntities.GetContext();
            var query = context.Payment.AsQueryable();
            
            if (cbFIO.SelectedValue != null)
            {
                string fio = cbFIO.SelectedValue.ToString();
                query = query.Where(x => x.User.FIO == fio);
            }
            
            if (cbCat.SelectedValue != null)
            {
                string cat = cbCat.SelectedValue.ToString();
                query = query.Where(x => x.Category.Name == cat);
            }

            DG.ItemsSource = query.ToList();
            UpdateStats();
        }
        
        private void UpdateStats()
        {
            itemsCount.Text = "Выбрано " + DG.Items.Count + 
                              " из " + PaymentsBaseEntities.GetContext().Payment.Count();

            decimal sum = DG.Items.OfType<Payment>().Sum(x => x.Summ);
            itemsSum.Text = "Сумма выбранных платежей: " + sum;
        }

        private void clear_onClick(object sender, RoutedEventArgs e)
        {
            cbFIO.SelectedValue = null;
            cbCat.SelectedValue = null;
            
            DG.ItemsSource = PaymentsBaseEntities.GetContext().Payment.ToList();
            UpdateStats();
        }
        
        private void FromDate_OnSelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            var context = PaymentsBaseEntities.GetContext();
            var query = context.Payment.AsQueryable();

            if (fromDate.SelectedDate != null)
            {
                var date = fromDate.SelectedDate;
                query = query.Where(x => x.Date >= date);
            }

            if (toDate.SelectedDate != null)
            {
                var date = toDate.SelectedDate;
                query = query.Where(x => x.Date <= date);
            } 
            
            DG.ItemsSource = query.ToList();
            UpdateStats();
        }
    }
}
