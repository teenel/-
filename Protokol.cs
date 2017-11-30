using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using _Word = Microsoft.Office.Interop.Word;

namespace Программа_для_военки
{
    public partial class Protokol : Form
    {
        _Word.Application word;
        Document protokol;
        object oMissing = System.Reflection.Missing.Value;

        _Word.Range wordcellrange;

        int n;
        int m;
        string[,] table;
        int vsego;

        public Protokol()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                Clear();

                word = new _Word.Application();

                protokol = word.Documents.Open(textBox_protokolWay.Text);
                Document vedomost = word.Documents.Open(textBox_path.Text);
                
                _Word.Table table_vedomost = vedomost.Tables[1];
                _Word.Table table_protokol = protokol.Tables[1];
                
                n = table_protokol.Rows.Count;
                m = table_protokol.Columns.Count;
                table = new string[n - 1, m - 1];

                int t = 0;

                #region
                for (int i = 5; i <= n + 1; i++)
                {
                    t = 0;
                    string stroka;

                    for (int j = 2; j < 12; j++)
                    {
                        if (j < 5)
                        {
                            wordcellrange = table_protokol.Cell(i - 1, j).Range;
                            stroka = replace(wordcellrange);
                            if (stroka == "")
                                stroka = "-";
                            table[i - 5, j - 2] = stroka;
                        }

                        else if (j % 2 == 1)
                        {
                            wordcellrange = table_vedomost.Cell(i, j).Range;
                            stroka = replace(wordcellrange);

                            if (stroka == "")
                                stroka = wordcellrange.Text.Replace("\r\a", "-");

                            table[i - 5, 3 + t] = stroka;

                            t++;
                        }

                        else if (j == 10)
                        {
                            wordcellrange = table_vedomost.Cell(i, j).Range;
                            stroka = replace(wordcellrange);

                            if (stroka == "")
                                stroka = wordcellrange.Text.Replace("\r\a", "-");

                            table[i - 5, 3 + t] = stroka;

                            t++;

                            wordcellrange = table_protokol.Cell(i - 1, j).Range;
                            stroka = replace(wordcellrange);

                            if (stroka == "")
                                stroka = wordcellrange.Text.Replace("\r\a", "-");

                            table[i - 5, 8] = stroka;
                        }
                    }

                    wordcellrange = table_protokol.Cell(i - 1, 12).Range;
                    stroka = replace(wordcellrange);

                    if (stroka == "")
                        stroka = wordcellrange.Text.Replace("\r\a", "-");

                    table[i - 5, 10] = stroka;


                    if (table[i - 5, 8] != "-")
                        table[i - 5, 9] = Convert.ToString(Convert.ToDouble(table[i - 5, 8]) * 20); //зачетка
                    else
                        table[i - 5, 9] = "-";


                    if (table[i - 5, 9] != "-" && table[i - 5, 7] != "-")
                        table[i - 5, 11] = Convert.ToString(Convert.ToInt32(table[i - 5, 9]) + Convert.ToInt32(table[i - 5, 7]));
                    else
                        table[i - 5, 11] = "-";
                }
                #endregion

                vedomost.Close();
                protokol.Close();
                word.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(word);

                #region

                table[n - 3, 0] = "Граждане, не выполнившие пороговый минимум по физической подготовленности";

                t = 0;

                for (int i = 0; i <= n - 3 - t; i++)
                {
                    for (int j = 3; j < 6; j++)
                    {
                        if (table[i, j] != "" && table[i, j] != null && table[i, j] != "-"
                            && Convert.ToInt16(table[i, j]) < 26)
                            {
                                change(table, n - 2, i);
                                i--;
                                t++;
                                break;
                            }
                    }
                }
                

                table[n - 2, 0] = "Граждане, не прошедшие предварительный отбор";

                t = 0;

                for (int i = 0; i <= n - 3 - t; i++)
                {
                    if (table[i, 0] == "Граждане, не выполнившие пороговый минимум по физической подготовленности")
                        continue;

                    else
                    {
                        if (table[i, 2] == "-" || table[i, 2].IndexOf("не годен") > 0 ||
                            table[i, 2].IndexOf("В") > 0 || table[i, 2].IndexOf("Г") > 0)
                        {
                            change(table, n - 1, i);
                            i--;
                            t++;
                            continue;
                        }

                        else if (table[i, 10] == "IV категория" || table[i, 10] == "-")
                        {
                            change(table, n - 1, i);
                            i--;
                            t++;
                            continue;
                        }

                        else
                            for (int j = 3; j < m - 3; j++)
                                if (table[i, j] == "-")
                                {
                                    change(table, n - 1, i);
                                    i--;
                                    t++;
                                    break;
                                }
                    }
                }

                for (int i = 0; i < n - 3; i++)
                    if (table[i, 0] == "Граждане, не выполнившие пороговый минимум по физической подготовленности")
                    {
                        t = i;
                        break;
                    }

                #endregion

                sortirovka(table, t);

                Cursor.Current = Cursors.WaitCursor;

                DisplayGrid(dataGridView1, table, n - 1, m - 1);

                Cursor.Current = Cursors.WaitCursor;

                if (textBox1.Text != "")
                {
                    for (int i = 0; i < n - 3; i++)
                    {
                        if (i < Convert.ToInt16(textBox1.Text) && i < t)
                        {
                            dataGridView1.Rows[i].Cells[12].Style.BackColor = Color.Green;
                            table[i, 12] = "Допустить";
                            dataGridView1.Rows[i].Cells[12].Value = table[i, 12];
                        }
                        else
                        {
                            dataGridView1.Rows[i].Cells[12].Style.BackColor = Color.Red;
                            table[i, 12] = "Отказать";
                            dataGridView1.Rows[i].Cells[12].Value = table[i, 12];
                        }
                    }

                    table[n - 2, 12] = "Отказать";
                    table[n - 3, 12] = "Отказать";

                    if (Convert.ToInt16(textBox1.Text) < t)
                        vsego = Convert.ToInt16(textBox1.Text);
                    else
                        vsego = t;
                }
                else
                    MessageBox.Show("Не введено кол-во поступающих", "Предупреждение");

                Cursor.Current = Cursors.Default;
            }
            catch
            {
                MessageBox.Show("Что-то пошло не так. \r\rПроверьте путь до файла или вид word-документа.", "Ошибка");
            }
        }

        private void DisplayGrid(DataGridView dataGridView, string[,] table, int n, int m)
        {
            int t = 0;
            int q = 0;

            for (int i = 0; i < n - 1; i++)
            {
                if (table[i, 0] == "Граждане, не выполнившие пороговый минимум по физической подготовленности")
                {
                    q++;
                    continue;
                }
                else if (table[i, 0] == "Граждане, не прошедшие предварительный отбор")
                {
                    q++;
                    continue;
                }

                t = 0;

                dataGridView.Rows.Add(i + 1 - q, null, null, null, null, null, null, null, null, null, null, null, null, null);

                if (q == 1)
                    dataGridView.Rows[i - q].DefaultCellStyle.BackColor = Color.Yellow;
                else if (q == 2)
                    dataGridView.Rows[i - q].DefaultCellStyle.BackColor = Color.SteelBlue;

                for (int j = 0; j < m; j++)
                {
                    if (j == 2)
                        t++;

                    dataGridView.Rows[i - q].Cells[j + 1 - t].Value = table[i, j];
                }
            }

            dataGridView.Rows[n - 1 - q].Cells[0].Value = n - 2;
            
            if (q == 1)
                dataGridView.Rows[n - 1 - q].DefaultCellStyle.BackColor = Color.Yellow;
            else if (q == 2)
                dataGridView.Rows[n - 1 - q].DefaultCellStyle.BackColor = Color.SteelBlue;
            t = 0;

            for (int j = 0; j < m; j++)
            {
                if (j == 2)
                    t++;

                dataGridView.Rows[n - 1 - q].Cells[j + 1 - t].Value = table[n - 1, j];
            }
        }

        private void button_way_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Выберите Word файл";
            openFileDialog1.Filter = "Word files |*.doc; *.docx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.FileName = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                textBox_path.Text = openFileDialog1.FileName;
        }

        private void button_protokolWay_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Выберите Word файл";
            openFileDialog1.Filter = "Word files |*.doc; *.docx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.FileName = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                textBox_protokolWay.Text = openFileDialog1.FileName;
        }

        private string replace(_Word.Range wordcellrange)
        {
            string st1 = wordcellrange.Text.Replace("\r\a", "");
            string st = st1.Replace("\r", "");
            return st;
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                word = new _Word.Application();

                protokol = word.Documents.Open(textBox_protokolWay.Text);
                _Word.Table table_protokol = protokol.Tables[1];

                table_protokol.Rows.Add();
                table_protokol.Rows.Add();

                for (int i = 4; i <= n + 2; i++)
                {
                    for (int j = 2; j <= m; j++)
                    {
                        wordcellrange = table_protokol.Cell(i, j).Range;
                        wordcellrange.Text = table[i - 4, j - 2];
                        wordcellrange.Font.Size = 14;
                    }

                    if (table[i - 4, 0] == "Граждане, не выполнившие пороговый минимум по физической подготовленности" ||
                        table[i - 4, 0] == "Граждане, не прошедшие предварительный отбор")
                    {
                        for (int x = 1; x < table_protokol.Columns.Count; x++)
                            table_protokol.Cell(i, 1).Merge(table_protokol.Cell(i, 2));
                        wordcellrange = table_protokol.Cell(i, 1).Range;
                        wordcellrange.Text = table[i - 4, 0];
                        wordcellrange.Bold = 1;
                    }

                }

                try
                {
                    _Word.Table table_Finish = protokol.Tables[2];

                    wordcellrange = table_Finish.Cell(1, 2).Range;
                    wordcellrange.Text = "– " + (n - 3) + "  чел.";

                    wordcellrange = table_Finish.Cell(1, 5).Range;
                    wordcellrange.Text = "– " + vsego + "  чел.";

                    int s = 0;

                    for (int x = 2; x <= table_Finish.Rows.Count; x++)
                    {
                        s = 0;
                        wordcellrange = table_Finish.Cell(x, 4).Range;
                        string stroka1 = replace(wordcellrange);

                        for (int i = 0; i < vsego; i++)
                        {
                            wordcellrange = table_protokol.Cell(i + 4, 3).Range;
                            string stroka = replace(wordcellrange);
                            if (stroka.IndexOf(stroka1) > 0)
                                s++;
                        }

                        wordcellrange = table_Finish.Cell(x, 5).Range;
                        wordcellrange.Text = "– " + Convert.ToString(s) + "  чел.";
                    }
                }
                catch
                {
                    MessageBox.Show("Не удалось посчитать кол-во поступивших по факультетам", "Предупреждение");
                }
                
                protokol.Save();
                protokol.Close();
                word.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(word);

                Cursor.Current = Cursors.Default;
            }
            catch
            {
                MessageBox.Show("Что-то пошло не так. \r\rПроверьте путь до файла и закройте word-документ.", "Ошибка");
            }
        }

        private void change(string[,] table, int n, int x)
        {
            string temp;

            for (int j = 0; j < 13; j++)
            {
                temp = table[x, j];

                for (int i = x; i < n - 1; i++)
                    table[i, j] = table[i + 1, j];

                table[n - 1, j] = temp;
            }
        }

        private void очиститьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clear();
        }

        private void Clear()
        {
            int leng = dataGridView1.Rows.Count;
            for (int i = 1; i < leng; i++)
                dataGridView1.Rows.RemoveAt(0);

            int weith = dataGridView1.Columns.Count;
            for (int j = 0; j < weith; j++)
                dataGridView1.Rows[0].Cells[j].Value = "";

            dataGridView1.Rows[0].DefaultCellStyle.BackColor = Color.White;
            dataGridView1.Rows[0].Cells[12].Style.BackColor = Color.White;
        }

        private void выйтиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void sortirovka(string [,] table, int n)
        {
            for (int i = 0; i < n; i++)
            {
                for (int j = i; j > 0 && Convert.ToInt16(table[j - 1, 11]) <= Convert.ToInt16(table[j, 11]); j--) // пока j > 0 и элемент j-1 > j, x-массив int
                {
                    if (table[j - 1, 11] == table[j, 11])
                    {
                        if (table[j, 10] == table[j - 1, 10])
                        {
                            if (Convert.ToInt16(table[j, 9]) > Convert.ToInt16(table[j - 1, 9]))
                                for (int x = 0; x < 13; x++)
                                {
                                    string temp = table[j - 1, x];
                                    table[j - 1, x] = table[j, x];
                                    table[j, x] = temp;
                                }
                        }
                        else if (table[j - 1, 10].Length >= table[j, 10].Length)
                            for (int x = 0; x < 13; x++)
                            {
                                string temp = table[j - 1, x];
                                table[j - 1, x] = table[j, x];
                                table[j, x] = temp;
                            }
                    }
                    else
                        for (int x = 0; x < 13; x++)
                        {
                            string temp = table[j - 1, x];
                            table[j - 1, x] = table[j, x];
                            table[j, x] = temp;
                        }

                }
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }
    }
}