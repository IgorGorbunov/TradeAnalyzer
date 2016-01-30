using System;
using System.Collections.Generic;


public class Deal
{
    public enum Direction
    {
        Short = 0,
        Long = 1
    }

    public bool IsLong
    {
        get;
        private set;
    }
    public Direction DirectionEnum
    {
        get;
        private set;
    }
    public string DirectionStr
    {
        get
        {
            switch (DirectionEnum)
            {
                case Direction.Long:
                    return "ЛОНГ";
                case Direction.Short:
                    return "шорт";
            }
            return "";
        }
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
    public double? OpenValue
    {
        get;
        private set;
    }
    public double? CloseValue
    {
        get;
        private set;
    }
    public double? ProfitProcent
    {
        get
        {
            int directCoef = 1;
            if (IsLong)
            {
                directCoef = -1;
            }
            double? profitProcentNul = (CloseValue - OpenValue) * directCoef * 100 / OpenValue;
            return Math.Round((double)profitProcentNul, 2);
        }
    }
    public string ProfitProcentStr
    {
        get
        {
            int directCoef = 1;
            if (IsLong)
            {
                directCoef = -1;
            }
            double? profitProcentNul = (CloseValue - OpenValue) * directCoef * 100 / OpenValue;
            char c = '+';
            if (profitProcentNul < 0)
            {
                c = '-';
            }
            double profitProcent = Math.Round(Math.Abs((double)profitProcentNul), 2);
            return profitProcent.ToString() + c;
        }
    }

    public Dictionary <DateTime, double?> Stops
    {
        get { return _stops; }
    }

    private readonly Dictionary <DateTime, double?> _stops;

    private const string LongTitle = "ЛOHГ";
    private const string ShortTitle = "ШOPT";

    public Deal(bool isLong, DateTime openDate, double? openValue) 
        : this (openDate, openValue)
    {
        IsLong = isLong;
        if (isLong)
        {
            DirectionEnum = Direction.Long;
        }
        else
        {
            DirectionEnum = Direction.Short;
        }
    }

    public Deal(string direction, DateTime openDate, double? openValue)
        : this(openDate, openValue)
    {
        SetDirection(direction);
    }

    private Deal(DateTime openDate, double? openValue)
    {
        OpenDate = openDate;
        OpenValue = openValue;
        _stops = new Dictionary<DateTime, double?>();
    }

    public void SetStopReverse(DateTime date, double? stopValue)
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

    public Deal Reverse(DateTime closeDate, double? closeValue)
    {
        CloseDate = closeDate;
        CloseValue = closeValue;
        return new Deal(!IsLong, closeDate, closeValue);
    }

    private void SetDirection(string dir)
    {
        string title = StringFunctions.GetClearlyDefinedString(dir.Trim());
        if (title == LongTitle)
        {
            IsLong = true;
            DirectionEnum = Direction.Long;
        }
        if (title == ShortTitle)
        {
            IsLong = false;
            DirectionEnum = Direction.Short;
        }
    }

}

