using System;
using System.Globalization;
using System.Linq;
using System.Windows;

namespace WpfApp1
{
    public partial class AddEditWindow : Window
    {
        private Payment _currentPaym = new Payment();
        
        public AddEditWindow(Payment selectedPayment)
        {
            InitializeComponent();

            if (selectedPayment != null)
            {
                _currentPaym = selectedPayment;
            }
            
            DataContext = _currentPaym;
            
            cbCategory.ItemsSource = PaymentsBaseEntities.GetContext().Category.ToList();
            cbFIO.ItemsSource = PaymentsBaseEntities.GetContext().User.ToList();
            
            if (_currentPaym.id == 0)
            {
                //PaymentsBaseEntities.GetContext().Payment.Add(_currentPaym);
                DatePicker.Text = DateTime.Today.ToString(CultureInfo.InvariantCulture);
            }
        }
        
        private void SaveBtn(object sender, RoutedEventArgs e)
        {
            try
            {
                var context = PaymentsBaseEntities.GetContext();

                if (_currentPaym.id == 0)
                {
                    context.Payment.Add(_currentPaym);
                }

                context.SaveChanges();
                MessageBox.Show("Данные сохранены!");
                Close();
            }
            catch (Exception ex)
            {
                string error = ex.InnerException?.InnerException?.Message ?? ex.Message;
                MessageBox.Show(error);
            }
        }

    }
}