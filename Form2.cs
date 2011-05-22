using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace myexcel
{
    public partial class Form2 : Form
    {
        public List<string> listaFunkcija = new List<string>();
        public List<string> opisiFunkcija = new List<string>();
        public MyExcel.Form1 excel;

        public Form2()
        {
            InitializeComponent();
            listaFunkcija.Add("MIN");
            opisiFunkcija.Add("Minimum \n\nKorištenje \nMIN(x1; x2; ..., xn) ili \nMIN(x1:xn)");
            listaFunkcija.Add("MAX");
            opisiFunkcija.Add("Maksimum \n\nKorištenje \nMIN(x1; x2; ..., xn) ili \nMIN(x1:xn)");
            listaFunkcija.Add("SUM");
            opisiFunkcija.Add("Suma \n\nKorištenje \nMIN(x1; x2; ..., xn) ili \nMIN(x1:xn)");
            listaFunkcija.Add("AVERAGE");
            opisiFunkcija.Add("Prosjek \n\nKorištenje \nMIN(x1; x2; ..., xn) ili \nMIN(x1:xn)");
            foreach (string fja in listaFunkcija)
                listBox1.Items.Add(fja);
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            label4.Text = opisiFunkcija[listBox1.SelectedIndex];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
//listBox1.Items[i].ToString());
