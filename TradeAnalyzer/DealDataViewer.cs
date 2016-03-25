using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


public class DealDataViewer
{
    public string OpenDate { get; set; }
    public double? OpenValue { get; set; }
    public bool IsLong { get; set; }
    public string CloseDate { get; set; }
    public double? CloseValue { get; set; }
    public int Duration { get; set; }
    public double Profit { get; set; }
    public double ProfitComis { get; set; }
}

