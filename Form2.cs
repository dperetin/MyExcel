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
        DataGridViewCell odabrana;
        bool okPressed = false;
        public Form2()
        {
            
            InitializeComponent();
           /* listaFunkcija.Add("MIN");
            opisiFunkcija.Add("Minimum \n\nKorištenje \nMIN(x1; x2; ...; xn) ili \nMIN(x1: xn)");
            listaFunkcija.Add("MAX");
            opisiFunkcija.Add("Maksimum \n\nKorištenje \nMAX(x1; x2; ...; xn) ili \nMAX(x1: xn)");
            listaFunkcija.Add("SUM");
            opisiFunkcija.Add("Suma \n\nKorištenje \nSUM(x1; x2; ...; xn) ili \nSUM(x1: xn)");
            listaFunkcija.Add("AVERAGE");
            opisiFunkcija.Add("Prosjek \n\nKorištenje \nAVERAGE(x1; x2; ...; xn) ili \nAVERAGE(x1: xn)");*/
            
            listBox1.SelectedIndexChanged += new EventHandler(listBox1_SelectedIndexChanged);
               
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (excel.fje.opisi.Count > listBox1.SelectedIndex && listBox1.SelectedIndex != -1)
            {
                label4.Text = excel.fje.opisi[listBox1.SelectedItem.ToString().ToLower()];
                string t = "= " + listBox1.SelectedItem.ToString().ToUpper() + "( )";
                textBox1.Text = t;
            }
            //textBox1.Select(1,5);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            excel.celijaIzForme = odabrana;
            excel.toolStripTextBox1.Text = textBox1.Text;
            excel.toolStripButton1_Click(null, null);
            excel.toolStripTextBox1.Clear();
            okPressed = true;
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            excel.otvorenaFormula = false;
            Close();
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            excel.otvorenIzbor = false;
            if (!okPressed)
                excel.otvorenaFormula = false;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            odabrana = excel.tGrid.SelectedCells[0];
            foreach (KeyValuePair<string, string> f in excel.fje.opisi)
                listBox1.Items.Add(f.Key.ToUpper());
        }

     
    }
}
