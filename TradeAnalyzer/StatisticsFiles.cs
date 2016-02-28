using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows;
using Microsoft.Win32;

namespace TradeAnalyzer
{
    public static class StatisticsFiles
    {
        private static readonly List <string> Files;

        private static List <TradeInstrument> issuers; 

        static StatisticsFiles()
        {
            Files = new List <string>();
        }

        public static void Check()
        {
            bool selectFiles = SetFiles();
            if (selectFiles)
            {
                foreach (string file in Files)
                {
                    TradeInstrument instrument = new TradeInstrument(file);
                    instrument.ReadAllQuotes();
                    instrument.ReadAllDeals();
                    instrument.WriteAllDeals();
                    instrument.WriteSimpleDeals();
                }
            }
            MessageBox.Show("vse!");
        }

        public static List <string> SetIssuerFiles()
        {
            string exePath = Assembly.GetExecutingAssembly().Location;
            string folder = Path.GetDirectoryName(exePath);
            string[] files = Directory.GetFiles(folder, "*.xls");
            issuers = new List <TradeInstrument>();
            foreach (string file in files)
            {
                issuers.Add(new TradeInstrument(file));
            }
            List <string> issuerNames = new List <string>();
            foreach (TradeInstrument issuer in issuers)
            {
                issuerNames.Add(issuer.Name);
            }
            issuerNames.Sort();
            return issuerNames;
        }

        private static bool SetFiles()
        {
            if (Files.Count > 0)
                return true;

            OpenFileDialog dlg = new OpenFileDialog
                {
                        FileName = "Выберите файлы",
                        Multiselect = true,
                        DefaultExt = ".xlsx",
                        Filter = "Электронные таблицы Excel (.xlsx, .xls)|*.xls;*.xlsx|All Files|*.*"
                };
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                foreach (string fileName in dlg.FileNames)
                {
                    Files.Add(fileName);
                }
                return true;
            }
            return false;
        }
    }
}
