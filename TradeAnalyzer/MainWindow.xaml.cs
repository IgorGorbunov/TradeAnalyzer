using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.DataVisualization.Charting;
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
        private Dictionary <DateTime, Deal> _romanDeals;
        public MainWindow()
        {
            InitializeComponent();
            
        }

        private void SetRomanDeals(string tradeInstrumentName)
        {
            _romanDeals = StatisticsFiles.GetRomanDeals(tradeInstrumentName);
            List <DealDataViewer> deals = new List <DealDataViewer>();
            foreach (KeyValuePair<DateTime, Deal> dealPair in _romanDeals)
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
                CbInstrs.Items.Add(issuerName);
            }
        }

        private void cbInstrs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SetRomanDeals(CbInstrs.SelectedValue.ToString());
        }

        private void CbDurationSelection_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Dictionary <DateTime, double?> points = new Dictionary <DateTime, double?>();
            double? startValue = 100;
            foreach (KeyValuePair <DateTime, Deal> deal in _romanDeals)
            {
                startValue = startValue * deal.Value.ProfitProcent / 100 + startValue;
                points.Add(deal.Key, startValue);
            }
            ((LineSeries) ChartProfitness.Series[0]).ItemsSource = points;
        }


    }
}
