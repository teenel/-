using System;
using System.Windows.Forms;
using Программа_для_военки.Properties;

namespace Программа_для_военки
{
    public partial class Setting : Form
    {
        public Setting()
        {
            InitializeComponent();
            textBox1.Text = Settings.Default["Excel_path"].ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button_save_Click(object sender, EventArgs e)
        {
            Settings.Default["Excel_path"] = textBox1.Text;
            Settings.Default.Save();

            Form1 form1 = new Form1();
            form1.first = false;

            Close();
        }

        private void button_path_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Выберите Excel файл";
            openFileDialog1.InitialDirectory = "d:\\";
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.FileName = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                textBox1.Text = openFileDialog1.FileName;

        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }
    }
}
