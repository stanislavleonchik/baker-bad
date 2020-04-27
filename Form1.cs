using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {

            InitializeComponent();
        }
        double weight, count;
        private void button1_Click(object sender, EventArgs e)
        {
           if (textBox1.Text != "" || textBox2.Text != "")
           {
               try
               {
                   weight = double.Parse(textBox1.Text);
                   count = double.Parse(textBox2.Text);
               }
               catch (FormatException)
               {
                   MessageBox.Show("Отсутствую или ошибочные данные!");
                   return;
               }
               bread breakinBad = new bread(weight, count);
               textBox3.Text = breakinBad.CountingProducts();
                button2.Visible = true;
                checkBox1.Visible = true;
           }
        }
        private bool closeCheckBox = true;
        private void button2_Click(object sender, EventArgs e)
        {
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = true;
            oDoc = oWord.Documents.Open(Environment.CurrentDirectory + "\\blank-spravka.doc");
            bread breakinBad = new bread(weight, count);
            oDoc.Bookmarks["Вес"].Range.Text = Convert.ToString(breakinBad.weight);
            oDoc.Bookmarks["Количество"].Range.Text = Convert.ToString(breakinBad.count);
            oDoc.Bookmarks["Продукты"].Range.Text = breakinBad.CountingProducts();
            oDoc.SaveAs2(FileName: Environment.CurrentDirectory + "\\Владимир Бланк");
            if (closeCheckBox == true)
            {
                oDoc.Close();
                oWord.Quit();
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            closeCheckBox = false;
        }
    }
    public class bread
    {
        public double weight;
        public double count;
        public bread(double weight, double count)
        {
            if ((weight > 0) || (count > 0))
            {
                this.weight = weight;
                this.count = count;
            }
            else MessageBox.Show("Должно быть положительное число");
        }
        public string CountingProducts()
        {
            string bufer = Convert.ToString(weight * 0.55 * count) + "гр. муки\r\n";
            bufer += Convert.ToString(weight * 0.4 * count) + "гр. воды\r\n";
            bufer += Convert.ToString(weight * 0.05 * count) + "гр. дрожжей";
            return bufer;
        }
    }
}
