using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
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

namespace TradeAnalyzer
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public DataTable dealsCollection;


        public MainWindow()
        {
            InitializeComponent();
            
        }

        private void SetRomanDeals(string tradeInstrumentName)
        {
            Dictionary<DateTime, Deal> romanDeals = StatisticsFiles.GetRomanDeals(tradeInstrumentName);
            List <DealDataViewer> deals = new List <DealDataViewer>();
            foreach (KeyValuePair <DateTime, Deal> dealPair in romanDeals)
            {
                DealDataViewer dealParams = new DealDataViewer();
                dealParams.OpenDate = dealPair.Key.ToShortDateString();
                Deal deal = dealPair.Value;
                dealParams.OpenValue = deal.OpenValue;
                dealParams.IsLong = deal.IsLong;
                dealParams.CloseDate = deal.CloseDate.ToShortDateString();
                dealParams.CloseValue = deal.CloseValue;
                TimeSpan dur = deal.CloseDate - deal.OpenDate;
                dealParams.Duration = dur.Days;
                dealParams.Profit = deal.ProfitProcent;
                dealParams.ProfitComis = 0.0;
                deals.Add(dealParams);
            }
            CollectionViewSource itemCollectionViewSource = (CollectionViewSource)(FindResource("ItemCollectionViewSource"));
            itemCollectionViewSource.Source = deals;
            //DgRomanDeals.ItemsSource = deals;
        }

        //---------------------------------------------------------------------------------

        private void miHelp_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine(MiHelp.Header.ToString());
        }

        private void BCheck_Click(object sender, RoutedEventArgs e)
        {
            StatisticsFiles.Check();
        }

        private void Window_Loaded_1(object sender, RoutedEventArgs e)
        {
            List <string> issuerNames = StatisticsFiles.SetIssuerFiles();
            foreach (string issuerName in issuerNames)
            {
                cbInstrs.Items.Add(issuerName);
            }
        }

        private void cbInstrs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SetRomanDeals(cbInstrs.SelectedValue.ToString());
        }

    }
}
