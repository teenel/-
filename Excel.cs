using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;


namespace Программа_для_военки
{
    class Excel
    {
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        string path = "";

        object _missingObj = System.Reflection.Missing.Value;

        public Excel(string path)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
        }

        public void CreateTableScore(double[,] Score, int sheet, int length)
        {
            ws = wb.Worksheets[sheet];

            for (int i = 2; i < length; i++)
            {
                if (ws.Cells[i, 1].Value2 != null)
                    Score[i - 2, 0] = (int)ws.Cells[i, 1].Value2;
                if (ws.Cells[i, 2].Value2 != null)
                    Score[i - 2, 1] = (int)ws.Cells[i, 2].Value2;
                if (ws.Cells[i, 3].Value2 != null)
                    Score[i - 2, 2] = (double)ws.Cells[i, 3].Value2;
                if (ws.Cells[i, 4].Value2 != null)
                    Score[i - 2, 3] = (double)ws.Cells[i, 4].Value2;
            }
        }
        
        public void CreateTableResults(int[,] Score, int sheet, int length)
        {
            ws = wb.Worksheets[sheet];

            for (int i = 2; i < length; i++)
            {
                if (ws.Cells[i, 1].Value2 != null)
                    Score[i - 2, 0] = (int)ws.Cells[i, 1].Value2;
                if (ws.Cells[i, 2].Value2 != null)
                    Score[i - 2, 1] = (int)ws.Cells[i, 2].Value2;
            }
        }

        public int Lenght(int sheet)
        {
            ws = wb.Worksheets[sheet];

            int len = 2;
            bool t = true;

            while (t == true)
            {
                if (ws.Cells[len, 1].Value2 == null)
                    t = false;
                else
                    len++;
            }
            return len;
        }

        public void Close()
        {
            
            wb.Close(false, _missingObj, _missingObj);
            excel.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

            excel = null;
            wb = null;
            ws = null;
            System.GC.Collect();
        }
    }
}
