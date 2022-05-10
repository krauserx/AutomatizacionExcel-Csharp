using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;


namespace DEV_TEST
{
    class clsLogica
    {
       
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        Workbook wb1;
        Worksheet ws1;
        int cont1;

        public clsLogica(string path,int Sheet,int cont)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
            cont1 = cont;
        }

        public void createNewSheet()
        {
            Worksheet temptsheet = wb.Worksheets.Add(After: ws);
            
        }

        public void CountRows()
        {

            Range last = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Range rangeX = ws.get_Range("A1",last);

            wb1 = excel.Workbooks.Open(@"C:\Users\kraus\Desktop\Pruebas Genpact\DEV TEST\CarpetaMonitoreo\LibroMaestro\LibroMaestro");
            ws1 = wb1.Worksheets[cont1];

            Range last1 = ws1.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Range rangeX1 = ws1.get_Range("A1:A1");

            rangeX.Copy(rangeX1);

            Worksheet temptsheet1 = wb1.Worksheets.Add(After: ws1);
            cont1++;
            wb1.Save();
            wb1.Close(true, Type.Missing, Type.Missing);
            wb.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
        }


    }
}
