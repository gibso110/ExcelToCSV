using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToCsv
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ExcelToCsvUI ui = new ExcelToCsvUI();

            ui.Run();
        }
    }
}
