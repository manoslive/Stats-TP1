using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace tp1_echantillonnage
{
    class Workbook
    {
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkBook;
    }
}
