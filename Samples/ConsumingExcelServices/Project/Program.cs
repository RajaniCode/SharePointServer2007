using System;
using System.Collections.Generic;
using System.Text;

using Project.com.tarangdemo.www; 

namespace Project
{
    class Program
    {
        private static string[] inputCells = { "BlackPenQty", "BlackPenRate", "GreenPenQty", "GreenPenRate" };
        private static string[] outputCells = { "Items", "Total","Average" };
        private static string workBookPath = "http://www.tarangdemo.com/ExcelLIbrary/Book1.xlsx";
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                args = new string[] { "1000","30","30","30"};
                Console.WriteLine("Using Default values of 5 and 30000");
            }
            else
            {
                Console.WriteLine("Using Values passed in as arguments"); 
            }
            ExcelService xlSrv = new ExcelService();
            xlSrv.Credentials = System.Net.CredentialCache.DefaultCredentials;
            Status[] status;
            string sessionId = xlSrv.OpenWorkbook(workBookPath, String.Empty, String.Empty, out status);

            for (int i = 0; i < inputCells.Length; i++)
            {
                status = xlSrv.SetCellA1(sessionId, "Sheet1", inputCells[i], args[i]); 
            }
            status = xlSrv.CalculateWorkbook(sessionId, CalculateType.Recalculate);

            foreach (string cellName in outputCells)
            {
                object result = xlSrv.GetCellA1(sessionId, "Sheet1", cellName, true, out status);
                Console.WriteLine("{0} is : {1}", cellName, result);
            }

            status = xlSrv.CloseWorkbook(sessionId);
        
        }
    }
}
