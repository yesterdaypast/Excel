using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DoReport;
using System.Linq.Expressions;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)

        {
            string strFilePath = @"C:\Users\97901\Project\Audi\Stab&Perf\test.xlsx";
            Startup st = new Startup();

            dynamic xlsx = new excel();
            
            xlsx.filename = strFilePath;
            xlsx.sheets = new List<sheet>();
            dynamic sheet1 = new sheet("stab");
            dynamic sheet2 = new sheet("perf");

            dynamic cell1 = new cell(1, 1, "1", "");
            dynamic cell2 = new cell(1, 2, "2", "");

            dynamic chart1 = new chart("test", "3D", "");

            sheet1.content.Add(cell1);
            sheet1.content.Add(cell2);

            sheet2.content.Add(chart1);

            xlsx.sheets.Add(sheet1);
            xlsx.sheets.Add(sheet2);

            dynamic dynXlsx = xlsx;

            st.GenerateReoprtAsync(dynXlsx);
        }
    }

    public class excel
    {
        public string filename;
        public List<sheet> sheets;

        public excel()
        { }

        public excel(string strFilePath)
        {
            this.filename = strFilePath;
            this.sheets = new List<sheet>();
        }
    }

    public class sheet
    {
        public string name;
        public List<object> content;

        public sheet(string strSheetName)
        {
            this.name = strSheetName;
            this.content = new List<object>();
        }
    }

    public class cell
    {
        public int row;
        public int column;
        public string value;
        public string formula;
        public string type = "cell";

        public cell(int row, int column, string value, string formula)
        {
            this.row = row;
            this.column = column;
            this.value = value;
            this.formula = formula;
        }
    }

    public class chart
    {
        public string title;
        public string legend;
        public string datasrc;
        public string type = "chart";

        public chart(string title, string legend, string datasrc)
        {

            this.title = title;
            this.legend = legend;
            this.datasrc = datasrc;
        }
    }
}
