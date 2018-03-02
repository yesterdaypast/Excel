using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DoReport;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Startup st = new Startup();
            dynobj dj  = new dynobj();

            st.GenerateReoprtAsync("C:\\Users\\97901\\Project\\ZD Data Logger\\test1.xlsx");
        }

        private class dynobj
        {
            public string filepath = "C:\\Users\\97901\\Project\\ZD Data Logger\\test1.xlsx";
        }
    }
}
