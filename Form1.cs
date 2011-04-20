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

       

    }
}
