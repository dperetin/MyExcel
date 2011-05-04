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
    public partial class Form1 : Form
    {
        Celije ListaCelija = new Celije();

        public Form1()
        {
            InitializeComponent();
            // popunjavam tablicu praznom redovima
            string[] red = {};
            for (int i = 1; i < 100; i++)
            {
                tablica.Rows.Add(red);
                tablica.Rows[i-1].HeaderCell.Value = i.ToString();
            }
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void tablica_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
           
        }

        private void tablica_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1 && e.ColumnIndex != -1)
            {
                ListaCelija.Dodaj(e.RowIndex, e.ColumnIndex);
            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            return;
        }

        private void tablica_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //ako kliknuta celija nije prazna, ispisuje se i njen sadrzaj, inace samo koordinate
            KeyValuePair<int, int> index = new KeyValuePair<int, int>(e.RowIndex, e.ColumnIndex);
            if (ListaCelija.sveCelije.ContainsKey(index))
            {
                if (ListaCelija.sveCelije[index].DajVrijednostCelije() != "")
                    statusLabel.Text = "Sadrzaj celije (" + e.RowIndex.ToString() + ", " 
                        + e.ColumnIndex.ToString() + "): " + ListaCelija.sveCelije[index].DajVrijednostCelije();
                else statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() + 
                    ", " + e.ColumnIndex.ToString() + ")";
            }
            else statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() + ", " + e.ColumnIndex.ToString() + ")";
        }

        void tablica_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //spremam podatke upisane u celiju
            string s = tablica.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            ListaCelija.DodajVrijednost(e.RowIndex, e.ColumnIndex, s);
        }


    }
}

//ovo mi treba
//string s = tablica.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();