using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using _Word = Microsoft.Office.Interop.Word;


namespace Программа_для_военки
{
    class Word
    {
        _Word.Application word = new _Word.Application();
        Document worddocument;
        object oMissing = System.Reflection.Missing.Value;

        string path;

        public Word(string path)
        {
            this.path = path;
            worddocument = word.Documents.Open(path);
            word.Visible = false;
        }
        
        public void AddFromWord(DataGridView dataGridView)
        {
            _Word.Range wordcellrange;
            _Word.Table table = worddocument.Tables[1];

            int n = table.Rows.Count;
            int m = table.Columns.Count;
            string stroka;
            int t = 0;

            for (int i = 5; i < n; i++)
            {
                dataGridView.Rows.Add(i - 4, null, null, null, null, null, null, null, null, null);

                t = 0;

                for (int j = 1; j < m - 3; j = j + 2)
                {
                    if (j == 3)
                        t++;

                    wordcellrange = table.Cell(i, j + 1).Range;
                    stroka = wordcellrange.Text.Replace("\r\a", "");
                    dataGridView.Rows[i - 5].Cells[j - t].Value = stroka;
                }
            }

            dataGridView.Rows[n - 5].Cells[0].Value = n - 4;

            t = 0;

            for (int j = 1; j < m - 3; j = j + 2)
            {
                if (j == 3)
                    t++;

                wordcellrange = table.Cell(n, j + 1).Range;
                stroka = wordcellrange.Text.Replace("\r\a", "");
                dataGridView.Rows[n - 5].Cells[j - t].Value = stroka;
            }
        }

        public void Save(DataGridView dataGridView)
        {
            _Word.Range wordcellrange;
            _Word.Table table = worddocument.Tables[1];

            
            int n = table.Rows.Count;
            int m = table.Columns.Count;

            for (int i = 5; i <= n; i++)
            {
                for (int j = 4; j < m - 1; j++)
                {
                    if (j % 2 == 0 || j == 9)
                        if (dataGridView.Rows[i - 5].Cells[j - 1].Value != null)
                        {
                            wordcellrange = table.Cell(i, j + 1).Range;
                            wordcellrange.Text = dataGridView.Rows[i - 5].Cells[j - 1].Value.ToString();
                        }
                }
            }

            worddocument.Save();
        }
        
        public void Close()
        {
            worddocument.Close();
            word.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(word);
        }
    }
}
