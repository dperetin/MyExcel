using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MyExcel
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
            opisiFunkcija.Add("Minimum \n\nKorištenje \nMIN(x1; x2; ...; xn) ili \nMIN(x1: xn)");
            listaFunkcija.Add("MAX");
            opisiFunkcija.Add("Maksimum \n\nKorištenje \nMAX(x1; x2; ...; xn) ili \nMAX(x1: xn)");
            listaFunkcija.Add("SUM");
            opisiFunkcija.Add("Suma \n\nKorištenje \nSUM(x1; x2; ...; xn) ili \nSUM(x1: xn)");
            listaFunkcija.Add("AVERAGE");
            opisiFunkcija.Add("Prosjek \n\nKorištenje \nAVERAGE(x1; x2; ...; xn) ili \nAVERAGE(x1: xn)");
            foreach (string fja in listaFunkcija)
                listBox1.Items.Add(fja);
            listBox1.SelectedIndexChanged += new EventHandler(listBox1_SelectedIndexChanged);
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            label4.Text = opisiFunkcija[listBox1.SelectedIndex];
            string t = "= " + listaFunkcija[listBox1.SelectedIndex] + "( )";
            textBox1.Text = t;
            textBox1.Select(1,5);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }     
    }
}
