using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


public static class StringFunctions
{

    public static string GetClearlyDefinedString(string s)
    {
        return s.Aggregate("", (current, c) => current + GetClearDefinedStringChar(c));
    }

    private static string GetClearDefinedStringChar(char c)
    {
        string ss = c.ToString().ToUpper();
        if (string.IsNullOrWhiteSpace(ss))
        {
            return " ";
        }
        switch (ss)
        {
            case "К":
                return "K";
            case "Е":
                return "E";
            case "Н":
                return "H";
            case "З":
                return "3";
            case "Х":
                return "X";
            case "В":
                return "B";
            case "А":
                return "A";
            case "Р":
                return "P";
            case "О":
                return "O";
            case "С":
                return "C";
            case "М":
                return "M";
            case "Т":
                return "T";
            case "1":
                return "L";
            case "8":
                return "S";
            case "0":
                return "O";
        }
        return ss;
    }
}

