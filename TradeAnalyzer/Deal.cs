using System;
using System.Collections.Generic;


public class Deal
{
    public enum Direction
    {
        Short = 0,
        Long = 1
    }

    public static string GetEnumStr(Direction direction)
    {
        switch (direction)
        {
            case Direction.Long:
                return "ЛОНГ";
            case Direction.Short:
                return "шорт";
        }
        return "";
    }

    public bool IsLong
    {
        get;
        private set;
    }
    public Direction DirectionDeal
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

    private readonly Dictionary <DateTime, double?> _stops;

    private const string LongTitle = "ЛOHГ";
    private const string ShortTitle = "ШOPT";

    public Deal(bool isLong, DateTime openDate, double? openValue) 
        : this (openDate, openValue)
    {
        IsLong = isLong;
        if (isLong)
        {
            DirectionDeal = Direction.Long;
        }
        else
        {
            DirectionDeal = Direction.Short;
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
            DirectionDeal = Direction.Long;
        }
        if (title == ShortTitle)
        {
            IsLong = false;
            DirectionDeal = Direction.Short;
        }
    }

}

