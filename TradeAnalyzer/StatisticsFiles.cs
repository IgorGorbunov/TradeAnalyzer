﻿using System.Collections.Generic;
using System.Windows;
using Microsoft.Win32;

namespace TradeAnalyzer
{
    public static class StatisticsFiles
    {
        private static readonly List <string> Files; 

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
