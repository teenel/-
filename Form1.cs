using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Программа_для_военки
{
    public partial class Form1 : Form
    {
        double[,] ScoreMale = new double[100, 4];
        double[,] ScoreFemale = new double[100, 4];
        int[,] Results = new int[100, 2];
        int length_male, length_female, length_results;
        public bool first = true;

        public Form1()
        {
            InitializeComponent();
        }

        private void CreateTabels()
        {
            if (first == true)
            {
                try
                {
                    first = false;

                    Setting setting = new Setting();

                    Excel excel = new Excel(setting.textBox1.Text);

                    length_male = excel.Lenght(1);
                    excel.CreateTableScore(ScoreMale, 1, length_male);

                    length_female = excel.Lenght(2);
                    excel.CreateTableScore(ScoreFemale, 2, length_female);

                    length_results = excel.Lenght(3);
                    excel.CreateTableResults(Results, 3, length_results);

                    excel.Close();
                    setting.Close();
                }
                catch
                {
                    MessageBox.Show("Ошибка с документом Excel", "Ошибка");
                }
            }
        } //Создание массивов баллов

        private void button1_Click(object sender, EventArgs e) //Кнопка рассчета
        {
            Cursor.Current = Cursors.WaitCursor;

            CreateTabels();
            Process();

            Cursor.Current = Cursors.Default;
        }

        public void Process()
        {
            if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text != "")
            {
                int count_pod = Convert.ToInt16(textBox1.Text);
                double count_run100 = Convert.ToDouble(textBox2.Text);
                double count_run3 = Convert.ToDouble(textBox3.Text);

                if (Юноши.Checked == true)
                {
                    textBox4.Text = Convert.ToString(Podtyag(count_pod, length_male, ScoreMale));
                    textBox5.Text = Convert.ToString(run(count_run100, length_male, ScoreMale, 2));
                    textBox6.Text = Convert.ToString(run(count_run3, length_male, ScoreMale, 3));

                }

                else if (Девушки.Checked == true)
                {
                    textBox4.Text = Convert.ToString(Podtyag(count_pod, length_female, ScoreFemale));
                    textBox5.Text = Convert.ToString(run(count_run100, length_female, ScoreFemale, 2));
                    textBox6.Text = Convert.ToString(run(count_run3, length_female, ScoreFemale, 3));
                }

                textBox_sum.Text = Convert.ToString(Convert.ToInt16(textBox4.Text) + Convert.ToInt16(textBox5.Text)
                    + Convert.ToInt16(textBox6.Text));
                
                textBox_result.Text = Convert.ToString(Result(Convert.ToInt16(textBox_sum.Text), length_results, Results));
            }
            else
                MessageBox.Show("Не введены данные нормативов", "Ошибка");
        } //Процесс идет сюда после кнопки1
        
        private int Podtyag(int count, int length, double[,] Score)
        {
            for (int i = 0; i < length - 1; i++)
            {
                if (count != 0)
                {
                    if (count > Score[0, 1])
                        return (int)Score[0, 0];
                    else if (count == Score[i, 1])
                        return (int)Score[i, 0];
                }
                else
                    return 0;
            }
            return 0;
        } //Подсчет баллов подтягиваний или пресса 

        private int run(double count, int length, double[,] Score, int col)
        {
            for (int i = 0; i < length - 1; i++)
            {
                if (count != 0)
                {
                    if (count < Score[0, col])
                        return (int)Score[0, 0];
                    else if (count == Score[i, col])
                        return (int)Score[i, 0];
                    else if (Score[i, col] < count && count < Score[i + 1, col])
                        return (int)Score[i + 1, 0];
                }
                else
                    return 0;
            }
            return 0;

        } //Подсчет баллов бега на 100м или 3км

        private int Result(int count, int length, int [,] Score)
        {
            if (count < Score[0, 0])
                return 0;

            else if (count < Score[length - 3 , 0])
                for (int i = 0; i < length - 1; i++)
                    if (count == Score[i, 0])
                        return Score[i, 1];
                
            return Score[length - 3, 1];
        } //Подсчет баллов

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Выберите файл с нормативами";
            openFileDialog1.Filter = "Word files |*.doc; *.docx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.FileName = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                textBox_path.Text = openFileDialog1.FileName;
        } //Путь к Word

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                Clear();

                Word word = new Word(textBox_path.Text);

                word.AddFromWord(dataGridView1);

                CreateTabels();

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (Юноши.Checked == true)
                    {
                        if (dataGridView1.Rows[i].Cells[2].Value.ToString() != "")
                            if (dataGridView1.Rows[i].Cells[2].Value.ToString() != "-")
                                dataGridView1.Rows[i].Cells[3].Value = Podtyag
                                (Convert.ToInt16(dataGridView1.Rows[i].Cells[2].Value), length_male, ScoreMale);

                        if (dataGridView1.Rows[i].Cells[4].Value.ToString() != "")
                            if (dataGridView1.Rows[i].Cells[4].Value.ToString() != "-")
                                dataGridView1.Rows[i].Cells[5].Value = run
                                (Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value), length_male, ScoreMale, 2);

                        if (dataGridView1.Rows[i].Cells[6].Value.ToString() != "")
                            if (dataGridView1.Rows[i].Cells[6].Value.ToString() != "-")
                                dataGridView1.Rows[i].Cells[7].Value = run
                                (Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value), length_male, ScoreMale, 3);
                    }

                    else if (Девушки.Checked == true)
                    {
                        if (dataGridView1.Rows[i].Cells[2].Value.ToString() != "")
                            if (dataGridView1.Rows[i].Cells[2].Value.ToString() != "-")
                                dataGridView1.Rows[i].Cells[3].Value = Podtyag
                                (Convert.ToInt16(dataGridView1.Rows[i].Cells[2].Value), length_female, ScoreFemale);

                        if (dataGridView1.Rows[i].Cells[4].Value.ToString() != "")
                            if (dataGridView1.Rows[i].Cells[4].Value.ToString() != "-")
                                dataGridView1.Rows[i].Cells[5].Value = run
                                (Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value), length_female, ScoreFemale, 2);

                        if (dataGridView1.Rows[i].Cells[6].Value.ToString() != "")
                            if (dataGridView1.Rows[i].Cells[6].Value.ToString() != "-")
                                dataGridView1.Rows[i].Cells[7].Value = run
                                (Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value), length_female, ScoreFemale, 3);
                    }

                    if (dataGridView1.Rows[i].Cells[6].Value.ToString() != "" && dataGridView1.Rows[i].Cells[4].Value.ToString() != "" && dataGridView1.Rows[i].Cells[2].Value.ToString() != ""
                        && dataGridView1.Rows[i].Cells[6].Value.ToString() != "-" && dataGridView1.Rows[i].Cells[4].Value.ToString() != "-" && dataGridView1.Rows[i].Cells[2].Value.ToString() != "-")
                    {
                        dataGridView1.Rows[i].Cells[8].Value = Convert.ToInt16(dataGridView1.Rows[i].Cells[3].Value) +
                            Convert.ToInt16(dataGridView1.Rows[i].Cells[5].Value) +
                            Convert.ToInt16(dataGridView1.Rows[i].Cells[7].Value);
                        dataGridView1.Rows[i].Cells[9].Value = Result(Convert.ToInt16(dataGridView1.Rows[i].Cells[8].Value), length_results, Results);
                    }
                }
                word.Close();

                Cursor.Current = Cursors.Default;
            }
            catch
            {
                MessageBox.Show("Что-то пошло не так. \r\rПроверьте путь до файла или вид word-документа.", "Ошибка");
            }
        } //Работа с word
        
        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                Word word = new Word(textBox_path.Text);

                word.Save(dataGridView1);

                word.Close();

                Cursor.Current = Cursors.Default;
            }
            catch
            {
                MessageBox.Show("Что-то пошло не так. \r\rПроверьте путь до файла и закройте word-документ.", "Ошибка");
            }
        } // Сохранить
        
        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Setting fm = new Setting();
            fm.Show();
        } // Открыть настройки

        private void textBox_Enter(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            textBox.SelectAll();
        } //Хождение Табом
        
        public void Clear()
        {
            int leng = dataGridView1.Rows.Count;
            for (int i = 1; i < leng; i++)
                dataGridView1.Rows.RemoveAt(0);

            int weith = dataGridView1.Columns.Count;
            for (int j = 0; j < weith; j++)
                dataGridView1.Rows[0].Cells[j].Value = "";
        }

        private void очиститьВсеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clear();
        }
        
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.BeginEdit(true);
        }

        private void работаСПротоколомToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Protokol fm = new Protokol();
            fm.Show();
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About fm = new About();
            fm.Show();
        }
        
        private void Юноши_Click(object sender, EventArgs e)
        {
            if (Девушки.Checked == true)
            {
                Юноши.Checked = true;
                Девушки.Checked = false;


                label1.Text = "Подтягивание:";
                label11.Text = "Бег на 3 км:";
                label1.TextAlign = ContentAlignment.MiddleRight;
                label1.Location = new Point(label1.Location.X - 25, label1.Location.Y);
            }
            Юноши.Checked = true;
        }

        private void Девушки_Click(object sender, EventArgs e)
        {
            if (Юноши.Checked == true)
            {
                Юноши.Checked = false;
                Девушки.Checked = true;

                label1.Text = "         Пресс:";
                label11.Text = "Бег на 1 км:";
                label1.TextAlign = ContentAlignment.MiddleRight;
                label1.Location = new Point(label1.Location.X + 25, label1.Location.Y);
            }
            Девушки.Checked = true;
        }

        private void textBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8 && e.KeyChar != 44)
            {
                e.Handled = true;
            }
            
            TextBox textBox = (TextBox)sender;

            if (textBox.Text.IndexOf(",") != -1 && e.KeyChar !=8)
            {
                if (e.KeyChar == 44)
                    e.Handled = true;

                if (textBox.Text.Length == textBox.Text.IndexOf(",") + 3)
                {
                    e.Handled = true;
                }
            }
        } //Ограничение по числам в textBox

        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        } //Закрытие приложения
    }
}