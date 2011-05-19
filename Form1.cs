using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing; 
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MyExcel
{
    public partial class Form1 : Form
    {
        List<DataGridView> gridovi = new List<DataGridView>();
        int broj_gridova = 0;

        List<Celije> ListaCelija = new List<Celije>();
        Funkcije fje = new Funkcije();
        
        //int brojRedaka = 1; //ima ih n, od 0 do n-1
        //int brojStupaca = 25; //isto od 0
        
        public Form1()
        {
            InitializeComponent();

            gridovi.Add(new DataGridView());

            tabControl1.TabPages[0].Controls.Add(gridovi[0]);
            gridovi[0].CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellClick);
            gridovi[0].CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellValueChanged);
            gridovi[0].CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellEndEdit);
            gridovi[broj_gridova].SelectionChanged += new EventHandler(this.tablica_SelectionChanged);

            Celije noviTab = new Celije();
            ListaCelija.Add(noviTab);

            gridovi[0].Dock = DockStyle.Fill;
            tabControl1.TabPages[0].Text = "Sheet1";
            gridovi[0].RowHeadersWidth = 60;
            for (int i = 65; i <= 90; i++)
            {
                DataGridViewColumn newCol = new DataGridViewColumn();
                newCol.HeaderText = Convert.ToChar(i).ToString();
                newCol.Visible = true;
                newCol.Width = 100;
                DataGridViewCell cell = new DataGridViewTextBoxCell();
                newCol.CellTemplate = cell;
                gridovi[0].Columns.Add(newCol);
            }
            string[] red = { };
            for (int i = 1; i < 100; i++)
            {
                gridovi[0].Rows.Add(red);
                gridovi[0].Rows[i - 1].HeaderCell.Value = i.ToString();
            }
            broj_gridova++;
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            int indexTaba = tabControl1.SelectedIndex;
            string s = "";
            foreach (KeyValuePair<KeyValuePair<int, int>, Cell> c in ListaCelija[indexTaba].sveCelije)
            {
                s += c.Value.sadrzaj + " ";
            }
            MessageBox.Show(s);
        }

        private void tablica_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //ako kliknuta celija nije prazna, ispisuje se i njen sadrzaj, inace samo koordinate
            int indexTaba = tabControl1.SelectedIndex;
            KeyValuePair<int, int> index = new KeyValuePair<int, int>(e.RowIndex, e.ColumnIndex);
            if (ListaCelija[indexTaba].sveCelije.ContainsKey(index))
            {
                if (ListaCelija[indexTaba].sveCelije[index].DajVrijednostFormule() != null)
                    toolStripTextBox1.Text = ListaCelija[indexTaba].sveCelije[index].formula;
                else toolStripTextBox1.Text = ListaCelija[indexTaba].sveCelije[index].sadrzaj;

                if (ListaCelija[indexTaba].sveCelije[index].DajVrijednostFormule() != null)
                        statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() + ", " +
                                e.ColumnIndex.ToString() + "); Sadrzaj celije: " +
                                ListaCelija[indexTaba].sveCelije[index].DajVrijednostCelije() +
                                "; Formula: " + ListaCelija[indexTaba].sveCelije[index].DajVrijednostFormule();

                else if (ListaCelija[indexTaba].sveCelije[index].DajVrijednostCelije() != "")
                    statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() + ", " +
                                e.ColumnIndex.ToString() + "); Sadrzaj celije: " +
                                ListaCelija[indexTaba].sveCelije[index].DajVrijednostCelije(); 
                    
                    else statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() +
                                ", " + e.ColumnIndex.ToString() + ")";
            }
            else
            {
                statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() + ", " + e.ColumnIndex.ToString() + ")";
                toolStripTextBox1.Text = "";
            }

            /*
            //izbrisi boju svih celija koje su prije bile kliknute
            for (int i = 0; i < brojStupaca; i++)
                for (int j = 0; j < brojRedaka; j++) 
                    tablica.Rows[j].Cells[i].Style.BackColor = Color.White;
            for (int i = 0; i < brojStupaca; i++)
                tablica.Columns[i].HeaderCell.Style.BackColor = Control.DefaultBackColor;
            for (int j = 0; j < brojRedaka; j++)
                tablica.Rows[j].HeaderCell.Style.BackColor = Control.DefaultBackColor;
            if ((e.ColumnIndex != -1) && (e.RowIndex != -1))
            tablica.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.LightSteelBlue;
            */
        }

        void tablica_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int indexTaba = tabControl1.SelectedIndex;
            if (gridovi[indexTaba].Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null) return;

            //ako se radi o formuli
            if (gridovi[indexTaba].Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()[0] == '=')
            {
                toolStripTextBox1.Text = gridovi[indexTaba].Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                toolStripButton1_Click(null, null);
                toolStripTextBox1.Clear();
            }

            //inace, ako se radi o obicnim brojevima
            else
            {
                //stvori novu celiju ako vec ne postoji
                KeyValuePair<int, int> index = new KeyValuePair<int, int>(e.RowIndex, e.ColumnIndex);
                if (!ListaCelija[indexTaba].sveCelije.ContainsKey(index) && e.RowIndex != -1 && e.ColumnIndex != -1)
                {
                    ListaCelija[indexTaba].Dodaj(e.RowIndex, e.ColumnIndex);
                }

                //spremam podatke upisane u celiju
                string s = gridovi[indexTaba].Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                ListaCelija[indexTaba].DodajVrijednost(e.RowIndex, e.ColumnIndex, s);
            }
        }

        private void tablica_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            //kad se doda novi redak, ispisi redni broj u header
          /*  tablica.Rows[brojRedaka - 1].HeaderCell.Value = brojRedaka.ToString();
            brojRedaka++;*/
        }

        private void tablica_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            //ako izbacimo i-ti redak, moramo promijeniti brojeve headera od i+1 nadalje
         /*   brojRedaka--;
            for (int j = 0; j < brojRedaka - 1; j++)
                tablica.Rows[j].HeaderCell.Value = Convert.ToString(j + 1);*/
        }

        private void tablica_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
        /*    tablica.Columns[e.ColumnIndex].HeaderCell.Style.BackColor = Color.SlateBlue;
            for (int i = 0; i < brojRedaka; i++)
                for (int j = 0; j < brojStupaca; j++)
                    if (j == e.ColumnIndex) tablica.Rows[i].Cells[j].Style.BackColor = Color.LightSteelBlue;
                    else tablica.Rows[i].Cells[j].Style.BackColor = Color.White;*/
        }

        private void tablica_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
         /*   tablica.Rows[e.RowIndex].HeaderCell.Style.BackColor = Color.SlateBlue;
            for (int i = 0; i < brojStupaca; i++)
                for (int j = 0; j < brojRedaka; j++)
                    if (j == e.RowIndex) tablica.Rows[j].Cells[i].Style.BackColor = Color.LightSteelBlue;
                    else tablica.Rows[j].Cells[i].Style.BackColor = Color.White; */
        }

        private void toolStripButton1_Click(object sender, EventArgs e) // GO! Izracunaj formulu
        {
            int indexTaba = tabControl1.SelectedIndex;
            int stupac = gridovi[indexTaba].SelectedCells[0].ColumnIndex;
            int redak = gridovi[indexTaba].SelectedCells[0].RowIndex;
            KeyValuePair<int, int> koordinate = new KeyValuePair<int, int>(redak, stupac);

            if (!ListaCelija[indexTaba].sveCelije.ContainsKey(koordinate) && redak != -1 && stupac != -1)
            {
                ListaCelija[indexTaba].Dodaj(redak, stupac);
            }

            Cell celija = ListaCelija[indexTaba].sveCelije[koordinate];
            //gridovi[indexTaba].SelectedCells[0].Value = celija.sadrzaj;

            string formula = toolStripTextBox1.Text.ToLower();
            celija.formula = formula;

            string fja = Regex.Match(formula, @"=\w*[(]").Value;
            string rje = Regex.Match(formula, "[(].*[)]").Value;
            string a, b;
            a = rje.TrimEnd(')');
            a = a.TrimStart('(');
            b = fja.TrimEnd('(');
            b = b.TrimStart('=');

            celija.sadrzaj = fje.SveFunkcije[b](ListaCelija[indexTaba].parsiraj(a)).ToString();
            gridovi[indexTaba].SelectedCells[0].Value = celija.sadrzaj;
            ListaCelija[indexTaba].DodajVrijednost(gridovi[indexTaba].SelectedCells[0].RowIndex, 
                gridovi[indexTaba].SelectedCells[0].ColumnIndex, celija.sadrzaj);
            ListaCelija[indexTaba].DodajFormulu(gridovi[indexTaba].SelectedCells[0].RowIndex,
                gridovi[indexTaba].SelectedCells[0].ColumnIndex, celija.formula);
        }

        private void tablica_SelectionChanged(object sender, EventArgs e)
        {
            List<Cell> argument = new List<Cell>();
            int indexTaba = tabControl1.SelectedIndex;
            for (int c = 0; c < gridovi[indexTaba].SelectedCells.Count; c++ )
            {
                KeyValuePair<int, int> index =
                    new KeyValuePair<int, int>(gridovi[indexTaba].SelectedCells[c].RowIndex,
                                                gridovi[indexTaba].SelectedCells[c].ColumnIndex);
                if (ListaCelija[indexTaba].sveCelije.ContainsKey(index))
                {
                    argument.Add(ListaCelija[indexTaba].sveCelije[index]);
                }

            }
            LabelSuma.Text = fje.SveFunkcije["sum"](argument).ToString(); 
        }

        private void toolStripButton4_Click(object sender, EventArgs e) //novi tab
        {
            string s = "Sheet" + (broj_gridova + 1);
            TabPage newPage = new TabPage(s);
            tabControl1.TabPages.Add(newPage);
            gridovi.Add(new DataGridView());

            gridovi[broj_gridova].RowHeadersWidth = 60;
            tabControl1.TabPages[broj_gridova].Controls.Add(gridovi[broj_gridova]);
            gridovi[broj_gridova].CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellClick);
            gridovi[broj_gridova].CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellEndEdit);
            gridovi[broj_gridova].CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellValueChanged);
            gridovi[broj_gridova].SelectionChanged += new EventHandler(this.tablica_SelectionChanged);
            Celije noviTab = new Celije();
            ListaCelija.Add(noviTab);

            gridovi[broj_gridova].Dock = DockStyle.Fill;
            for (int i = 65; i <= 90; i++)
            {
                DataGridViewColumn newCol = new DataGridViewColumn();
                newCol.HeaderText = Convert.ToChar(i).ToString();
                newCol.Visible = true;
                newCol.Width = 100;
                DataGridViewCell cell = new DataGridViewTextBoxCell();
                newCol.CellTemplate = cell;
                gridovi[broj_gridova].Columns.Add(newCol);
            }
            string[] red = { };
            for (int i = 1; i < 100; i++)
            {
                gridovi[broj_gridova].Rows.Add(red);
                gridovi[broj_gridova].Rows[i - 1].HeaderCell.Value = i.ToString();
            }
            broj_gridova++;
        }

        private void toolStripTextBox1_KeyPress(object sender, KeyPressEventArgs e) //enter u textboxu
        {
            if (!(e.KeyChar == (char)Keys.Enter)) return;
            toolStripButton1_Click(null, null);
            int indexTaba = tabControl1.SelectedIndex;
            KeyValuePair<int, int> index = new KeyValuePair<int, int>
                (gridovi[indexTaba].SelectedCells[0].RowIndex, gridovi[indexTaba].SelectedCells[0].ColumnIndex);
            statusLabel.Text = "Koordinate celije: (" + gridovi[indexTaba].SelectedCells[0].ColumnIndex.ToString() + ", " +
                                gridovi[indexTaba].SelectedCells[0].ColumnIndex.ToString() + "); Sadrzaj celije: " +
                                ListaCelija[indexTaba].sveCelije[index].DajVrijednostCelije() +
                                "; Formula: " + ListaCelija[indexTaba].sveCelije[index].DajVrijednostFormule();
            toolStripTextBox1.Text = "";
            int r = gridovi[indexTaba].SelectedCells[0].RowIndex;
            int s = gridovi[indexTaba].SelectedCells[0].ColumnIndex;
            gridovi[indexTaba].ClearSelection();
            gridovi[indexTaba].Rows[r + 1].Cells[s].Selected = true;
            
        }

        private void tablica_CellValueChanged(object sender, DataGridViewCellEventArgs e) //ne radi
        {
            //nesto ne valja s ovim e
            int indexTaba = tabControl1.SelectedIndex;
            //toolStripTextBox1.Text = gridovi[indexTaba].Rows[e.RowIndex].Cells[e.ColumnIndex].ToString();
            //statusLabel.Text = e.RowIndex + " " + e.ColumnIndex;
        }

        
    }
}
