using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel=Microsoft.Office.Interop.Excel;

namespace ExcelTester
{
    class ExcelCreatorClass
    {

        public void CreateSpread()
        {
            Excel.Application MyApp=null;
            Excel.Worksheet MySheet=null;
            Excel.Workbook MyBook=null;


            MyApp = new Excel.Application();
            MyApp.Visible = true;

            if (MyApp == null)
            {
                MessageBox.Show("Excel is not properly installed on your system!");
                return;
            }

            DirectoryInfo dirInfo = new DirectoryInfo(Directory.GetCurrentDirectory());
            String temp = dirInfo.Parent.Parent.FullName;

            String excelPath = Path.Combine(temp, "Resources", "tmpIStarPanelMap.xlsx");

            MyBook = MyApp.Workbooks.Open(excelPath);
            MySheet = MyBook.Worksheets.get_Item(2);
            
            //Excel.Range er = MySheet.get_Range("B:B", System.Type.Missing);
            //er.EntireColumn.ColumnWidth = 1;

            MySheet.Cells[1, 2] = "Test Bitches";

            //MyBook.Save();

        }

        


    }
}
