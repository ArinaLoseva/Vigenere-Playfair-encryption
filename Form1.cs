using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace WindowsFormsApp8
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        public void ExcelTabel()
        {
            string filePath = "D:\\Cipher3.xlsx";
            string sheetName = "Sheet1";

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filePath);
            Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[sheetName];
            Excel.Range excelRange = excelWorksheet.UsedRange;

            DataTable dt = new DataTable();

            for (int i = 1; i <= excelRange.Columns.Count; i++)
            {
                dt.Columns.Add(excelRange.Cells[1, i].Value2.ToString());
            }

            for (int i = 2; i <= excelRange.Rows.Count; i++)
            {
                DataRow row = dt.NewRow();
                for (int j = 1; j <= excelRange.Columns.Count; j++)
                {
                    row[j - 1] = excelRange.Cells[i, j].Value2;
                }
                dt.Rows.Add(row);
            }

            dataGridView1.DataSource = dt;

            excelWorkbook.Close();
            excelApp.Quit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelTabel();
            string text = textBox1.Text;
            string key = textBox2.Text;
            textBox3.Text = "";

            for (int i = 0; i < text.Length; i++)
            {
                for (int m = 0; m < dataGridView1.RowCount - 1; m++)
                {
                    for (int n = 0; n < dataGridView1.ColumnCount; n++)
                    {
                        string t = Convert.ToString(text[i]);
                        string k = Convert.ToString(key[(i % key.Length)]);
                        if (dataGridView1[0, m].Value.ToString() == t && dataGridView1[n, 0].Value.ToString() == k)
                        {
                            textBox3.Text = textBox3.Text + dataGridView1[n, m].Value;
                        }
                    }
                }
            }

            SaveFileDialog sfd = new SaveFileDialog();
            if (sfd.ShowDialog() == DialogResult.OK)
                System.IO.File.WriteAllText(sfd.FileName, textBox3.Text);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog o = new OpenFileDialog();
            o.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            if (o.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = System.IO.File.ReadAllText(o.FileName, Encoding.Default);
            }

            ExcelTabel();
            string text = textBox1.Text;
            string key = textBox2.Text;
            textBox3.Text = "";

            for (int i = 0; i < text.Length; i++)
            {
                for (int n = 0; n < dataGridView1.ColumnCount; n++)
                {
                    string t = Convert.ToString(text[i]);
                    string k = Convert.ToString(key[(i % key.Length)]);
                    if (dataGridView1[n, 0].Value.ToString() == k)
                    {
                        for (int m = 0; m < dataGridView1.RowCount - 1; m++)
                        {
                            if(dataGridView1[n, m].Value.ToString() == t)
                            {
                                textBox3.Text = textBox3.Text + dataGridView1[0, m].Value;
                            }
                        }
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string alphabet = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ";
            string cipher = textBox4.Text;
            string key = textBox5.Text;
            int[] text = {1, 2, 3, 4, 5, 6, 7, 8, 9, 0};
            int count = 0;

            for(int i = 0; i < cipher.Length; i++)
            {
                for (int j = 0; j < alphabet.Length; j++)
                {
                    if (alphabet[j] == cipher[i])
                    {
                        text[count] = j;
                        count++;
                    }
                }
            }

            string inputWord = key.ToUpper(); 
            string resultString = inputWord;

            foreach (char letter in alphabet)
            {
                if (!inputWord.Contains(letter)) // Проверяем, содержит ли введенное слово текущую букву алфавита
                {
                    resultString += letter;
                }
            }

            for (int i = 0; i < cipher.Length; i++)
            {
                textBox2.Text = textBox2.Text + resultString[text[i]];
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }
    }
}
