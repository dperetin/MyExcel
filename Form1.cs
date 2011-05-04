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
                ListaCelija.Dodaj(e.ColumnIndex.ToString(), e.RowIndex);
                //statusLabel.Text = tablica[e.ColumnIndex, e.RowIndex].Value.ToString();
            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            return;
        }

        private void tablica_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            statusLabel.Text = "Koordinate celije: (" + e.ColumnIndex.ToString() + ", " + e.RowIndex.ToString() + ")";
        }

        private void tablica_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            ///blablabla
        }

       

    }
}
