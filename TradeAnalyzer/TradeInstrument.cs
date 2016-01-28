using System;
using System.Collections.Generic;
using System.Globalization;

public class TradeInstrument
{
    public string Code
    {
        get;
        private set;
    }
    public string Name
    {
        get;
        private set;
    }

    private readonly string _quotesFileName;
    private Dictionary <DateTime, Quote> _quotes;
    private Dictionary<DateTime, Deal> _deals;

    private const string DateCol = "C";
    private const string OpenCol = "E";
    private const string CloseCol = "H";
    private const string HighCol = "F";
    private const string LowCol = "G";
    private const string VolumeCol = "I";

    private const string DirectionDealCol = "K";
    private const string OpenDealCol = "L";
    private const string ReverseDealCol = "M";

    private const string NewDirectionDealCol = "O";
    private const string NewOpenDealCol = "P";
    private const string NewReverseDealCol = "Q";
    private const string NewProfitLostCol = "R";

    private const int FirstRow = 2;

    private const string DateFormat = "ddMMyy";

    public TradeInstrument(string code, string name, string quotesFileName)
    {
        Code = code;
        Name = name;
        _quotesFileName = quotesFileName;
    }

    public void ReadAllQuotes()
    {
        DateTime fDateTime = new DateTime(1, 1, 1);
        DateTime tDateTime = DateTime.Today;
        ReadQuotes(fDateTime, tDateTime);
    }

    public void ReadQuotes(DateTime fromDate, DateTime toDate)
    {
        using (ExcelClass xls = new ExcelClass())
        {
            try
            {
                _quotes = new Dictionary <DateTime, Quote>();
                xls.OpenDocument(_quotesFileName, false);
                string sDate = xls.GetCellStringValue(DateCol, FirstRow);
                int i = FirstRow;
                while (!string.IsNullOrEmpty(sDate))
                {
                    DateTime date = StringFunctions.GetDate(sDate, DateFormat);
                    if (date > fromDate && date < toDate)
                    {
                        double open = double.Parse(xls.GetCellStringValue(OpenCol, i));
                        double close = double.Parse(xls.GetCellStringValue(CloseCol, i));
                        double high = double.Parse(xls.GetCellStringValue(HighCol, i));
                        double low = double.Parse(xls.GetCellStringValue(LowCol, i));
                        int volume = int.Parse(xls.GetCellStringValue(VolumeCol, i));
                        Quote quote = new Quote(date, open, close, high, low, volume);
                        _quotes.Add(date, quote);
                    }
                    i++;
                    sDate = xls.GetCellStringValue(DateCol, i);
                }
            }
            finally 
            {
                xls.CloseDocument(false);
            }
        }
    }

    public void ReadAllDeals()
    {
        DateTime fDateTime = new DateTime(1, 1, 1);
        DateTime tDateTime = DateTime.Today;
        ReadDeals(fDateTime, tDateTime);
    }

    public void ReadDeals(DateTime fromDate, DateTime toDate)
    {
        using (ExcelClass xls = new ExcelClass())
        {
            try
            {
                _deals = new Dictionary<DateTime, Deal>();
                xls.OpenDocument(_quotesFileName, false);
                string sDate = xls.GetCellStringValue(DateCol, FirstRow);
                int i = FirstRow;
                Deal deal = null;
                bool isFirstDeal = true;
                while (!string.IsNullOrEmpty(sDate))
                {
                    DateTime date = StringFunctions.GetDate(sDate, DateFormat);
                    if (date > fromDate && date < toDate)
                    {
                        string sDir = xls.GetCellStringValue(DirectionDealCol, i);
                        string sOpen = xls.GetCellStringValue(OpenDealCol, i);
                        string sReverse = xls.GetCellStringValue(ReverseDealCol, i);
                        if (!string.IsNullOrEmpty(sDir) &&
                            !string.IsNullOrEmpty(sOpen) &&
                            !string.IsNullOrEmpty(sReverse))
                        {
                            double? reverse = StringFunctions.TryParse(sReverse);
                            if (isFirstDeal)
                            {
                                double? open = StringFunctions.TryParse(sOpen);
                                deal = new Deal(sDir, date.AddDays(-1), open);
                                deal.SetStopReverse(date, reverse);
                                isFirstDeal = false;
                            }
                            else
                            {
                                if (deal.IsSameDirection(sDir))
                                {
                                    deal.SetStopReverse(date, reverse);
                                }
                                else
                                {
                                    double? open = StringFunctions.TryParse(sOpen);
                                    _deals.Add(deal.OpenDate, deal);
                                    deal = deal.Reverse(date.AddDays(-1), open);
                                    deal.SetStopReverse(date, reverse);
                                }
                            }
                        }
                    }
                    i++;
                    sDate = xls.GetCellStringValue(DateCol, i);
                }
            }
            finally
            {
                xls.CloseDocument(false);
            }
        }
    }

    public void WriteAllDeals()
    {
        using (ExcelClass xls = new ExcelClass())
        {
            try
            {
                xls.OpenDocument(_quotesFileName, false);
                int i = FirstRow;
                foreach (KeyValuePair <DateTime, Deal> pair in _deals)
                {
                    string sDate = xls.GetCellStringValue(DateCol, i);
                    while (!string.IsNullOrEmpty(sDate))
                    {
                        DateTime date = StringFunctions.GetDate(sDate, DateFormat);
                        if (date.AddDays(-1) == pair.Key)
                        {
                            SetDealStops(xls, ref i, pair.Value);
                            break;
                        }
                        i++;
                        sDate = xls.GetCellStringValue(DateCol, i);
                    }
                }
            }
            finally
            {
                xls.CloseDocumentSave();
            }
        }
    }

    private void SetDealStops(ExcelClass xls, ref int iRow, Deal deal)
    {
        foreach (KeyValuePair <DateTime, double?> stop in deal.Stops)
        {
            string sDate = xls.GetCellStringValue(DateCol, iRow);
            DateTime date = StringFunctions.GetDate(sDate, DateFormat);
            while (stop.Key != date)
            {
                iRow++;
                sDate = xls.GetCellStringValue(DateCol, iRow);
                date = StringFunctions.GetDate(sDate, DateFormat);
            }
            xls.SetCellValue(NewDirectionDealCol, iRow, deal.DirectionStr);
            xls.SetCellValue(NewOpenDealCol, iRow, deal.OpenValue.ToString());
            xls.SetCellValue(NewReverseDealCol, iRow, stop.Value.ToString());
            iRow++;
        }
        double? profitProcentNul;
        int directCoef = 1;
        if (!deal.IsLong)
        {
            directCoef = -1;
        }
        profitProcentNul = (deal.CloseValue - deal.OpenValue)*directCoef*100/deal.OpenValue;
        char c = '+';
        if (profitProcentNul < 0)
        {
            c = '-';
        }
        double profitProcent = Math.Round(Math.Abs((double) profitProcentNul), 2);
        xls.SetCellValue(NewProfitLostCol, iRow, profitProcent.ToString() + c);
    }

}

