using System;
using System.Collections.Generic;
using System.Text;

using Microsoft.Office.Excel.Server.Udf;

namespace SimpleFunction
{
    [UdfClass()]
    public class Dice
    {
        [UdfMethod()]
        public Int32 Roll()
        {
            Random r = new Random();
            return (Int32)((r.NextDouble() * 5) + 1);
        }

    }
}
