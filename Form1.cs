﻿using System;
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
        int brojRedaka = 1; //ima ih n, od 0 do n-1
        int brojStupaca = 25; //isto od 0

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

            //izbrisi boju svih celija koje su prije bile kliknute
            for (int i = 0; i < brojStupaca; i++)
                for (int j = 0; j < brojRedaka; j++) 
                    tablica.Rows[j].Cells[i].Style.BackColor = Color.White;
            for (int i = 0; i < brojStupaca; i++)
                tablica.Columns[i].HeaderCell.Style.BackColor = Control.DefaultBackColor;
            for (int j = 0; j < brojRedaka; j++)
                tablica.Rows[j].HeaderCell.Style.BackColor = Control.DefaultBackColor;
        
        }

        void tablica_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (tablica.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null) return;

            //stvori novu celiju ako vec ne postoji
            KeyValuePair<int, int> index = new KeyValuePair<int, int>(e.RowIndex, e.ColumnIndex);
            if (!ListaCelija.sveCelije.ContainsKey(index) && e.RowIndex != -1 && e.ColumnIndex != -1)
            {
                ListaCelija.Dodaj(e.RowIndex, e.ColumnIndex);
            }
           
            //spremam podatke upisane u celiju
            string s = tablica.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            ListaCelija.DodajVrijednost(e.RowIndex, e.ColumnIndex, s);
        }

        private void tablica_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            //kad se doda novi redak, ispisi redni broj u header
            tablica.Rows[brojRedaka - 1].HeaderCell.Value = brojRedaka.ToString();
            brojRedaka++;
        }

        private void tablica_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            //ako izbacimo i-ti redak, moramo promijeniti brojeve headera od i+1 nadalje
            brojRedaka--;
            for (int j = 0; j < brojRedaka - 1; j++)
                tablica.Rows[j].HeaderCell.Value = Convert.ToString(j + 1);
        }

        private void tablica_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            tablica.Columns[e.ColumnIndex].HeaderCell.Style.BackColor = Color.SlateBlue;
            for (int i = 0; i < brojRedaka; i++)
                for (int j = 0; j < brojStupaca; j++)
                    if (j == e.ColumnIndex) tablica.Rows[i].Cells[j].Style.BackColor = Color.LightSteelBlue;
                    else tablica.Rows[i].Cells[j].Style.BackColor = Color.White;
        }

        private void tablica_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            tablica.Rows[e.RowIndex].HeaderCell.Style.BackColor = Color.SlateBlue;
            for (int i = 0; i < brojStupaca; i++)
                for (int j = 0; j < brojRedaka; j++)
                    if (j == e.RowIndex) tablica.Rows[j].Cells[i].Style.BackColor = Color.LightSteelBlue;
                    else tablica.Rows[j].Cells[i].Style.BackColor = Color.White;
        }

    

    }
}

//ovo mi treba
//string s = tablica.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();


