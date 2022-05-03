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


        public clsLogica(string path,int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
        }

        public void CreateNewFile()
        {
            this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.ws = wb.Worksheets[1];
        }

        public void createNewSheet()
        {
            Worksheet temptsheet = wb.Worksheets.Add(After: ws);
            
        }
        public string[,] ReadRange(int starti,int starty,int endi,int endy)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            object[,] holder = range.Value2;
            string[,] returnstring = new string[endi - starti, endy - starty];
            for (int p= 1; p <= endi - starti; p++)
            {
                for (int q = 1; q <= endy - starty; q++)
                {
                    returnstring[p - 1, q - 1] = holder[p, q].ToString();
                }
            }
            return returnstring;
        }

       
        public void WriteRange(int starti, int starty, int endi, int endy, string [,] writestring)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            range.Value2 = writestring;
        }

        public void save()
        {
            wb.Save();
        }

        public void SelectWorksheet(int sheetNumber)
        {
            this.ws = wb.Worksheets[sheetNumber];
        }

        public void deleteWorksheet(int SheetNumber)
        {
            wb.Worksheets[SheetNumber].Delete();
        }

        public void close()
        {
            wb.Close();
        }

    }
}
