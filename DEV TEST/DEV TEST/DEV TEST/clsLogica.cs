using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
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
        int cont;
        int cont1;

        public clsLogica(string path)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            cont = wb.Sheets.Count;
        }

        public void createNewSheet()
        {
            Worksheet temptsheet = wb1.Worksheets.Add(After: ws1);
            
        }

        public void openLibroMaestro()
        {
         
                wb1 = excel.Workbooks.Open(@"C:\Users\kraus\Desktop\Pruebas Genpact\DEV TEST\CarpetaMonitoreo\LibroMaestro\LibroMaestro");
                cont1 = wb1.Sheets.Count;
                ws1 = wb1.Worksheets[cont1];
            
    
        }

        public void savesAndCloses()
        {
            wb1.Save();
            wb1.Close(true, Type.Missing, Type.Missing);
            wb.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
        }


        public void CountRows()
        {
            for (int i = 1; i <= cont; i++)
            {
             ws = wb.Worksheets[i];

            Range last = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Range rangeX = ws.get_Range("A1",last);

            openLibroMaestro();

            Range last1 = ws1.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Range rangeX1 = ws1.get_Range("A1:A1");

            rangeX.Copy(rangeX1);

            createNewSheet();
            }
            savesAndCloses();

        }


    }
}
