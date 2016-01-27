using System;
using System.Collections.Generic;


public class Deal
{
    public bool IsLong
    {
        get;
        private set;
    }
    public DateTime OpenDate
    {
        get;
        private set;
    }
    public DateTime CloseDate
    {
        get;
        private set;
    }
    public double OpenValue
    {
        get;
        private set;
    }
    public double CloseValue
    {
        get;
        private set;
    }

    private readonly Dictionary <DateTime, double> _stops;

    private const string LongTitle = "ЛOHГ";
    private const string ShortTitle = "ШOPT";

    public Deal(bool isLong, DateTime openDate, double openValue)
    {
        IsLong = isLong;
        OpenDate = openDate;
        OpenValue = openValue;
        _stops = new Dictionary <DateTime, double>();
    }

    public void SetStopReverse(DateTime date, double stopValue)
    {
        if (!_stops.ContainsKey(date))
        {
            _stops.Add(date, stopValue);
        }
    }

    public bool IsSameDirection(string directionTitle)
    {
        string title = StringFunctions.GetClearlyDefinedString(directionTitle.Trim());
        if (IsLong && title == LongTitle)
        {
            return true;
        }
        if (!IsLong && title == ShortTitle)
        {
            return true;
        }
        return false;
    }

    public Deal Reverse(DateTime date, double closeValue)
    {
        CloseDate = date;
        CloseValue = closeValue;
        return new Deal(!IsLong, date, closeValue);
    }


}

