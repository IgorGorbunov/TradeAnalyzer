using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

    private string _quotesFileName;

    public TradeInstrument(string code, string name, string quotesFileName)
    {
        Code = code;
        Name = name;
        _quotesFileName = quotesFileName;
    }

    public void ReadQuotes()
    {
        
    }
}

