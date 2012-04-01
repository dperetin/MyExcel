using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing; 
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Schema;
using System.IO;
using System.Drawing.Drawing2D;

namespace MyExcel
{
    public partial class Form1 : Form
    {
        List<DataGridView> gridovi = new List<DataGridView>();
        int broj_gridova = 0;
        string imeFilea;
        
        List<Celije> ListaCelija = new List<Celije>();
        public Funkcije fje = new Funkcije();
        public bool loading = false;
        public DataGridView tGrid; // trenutno aktivni grid
        Celije tCell;       // trenutno aktivni skup celija
        public DataGridViewCell celijaIzForme;

        public bool otvorenIzbor = false;
        public bool otvorenaFormula = false;
        public List<DataGridViewCell> roze = new List<DataGridViewCell>();
        bool smijemZagasiti = false;
        MyExcel.Form2 funkcije;
        
        public Form1()
        {
            InitializeComponent();
            
            gridovi.Add(new DataGridView());
            tGrid = gridovi[0];
            tabControl1.KeyDown += new KeyEventHandler(delete);
            tabControl1.TabPages[0].Controls.Add(gridovi[0]);
            tabControl1.Selected += new TabControlEventHandler(promjenaTaba);
            gridovi[0].BorderStyle = BorderStyle.None;
            tabControl1.TabPages[0].BorderStyle = BorderStyle.None;
            gridovi[0].CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellClick);
            gridovi[0].CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellValueChanged);
            gridovi[0].CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellEndEdit);
            gridovi[0].SelectionChanged += new EventHandler(this.tablica_SelectionChanged);
            gridovi[0].CellEnter += new DataGridViewCellEventHandler(this.tablica_CellEnter);
            
            gridovi[0].RowsAdded += new DataGridViewRowsAddedEventHandler(tablica_RowsAdded);
            
            Celije noviTab = new Celije();
            ListaCelija.Add(noviTab);
            tCell = ListaCelija[0];

            gridovi[0].Dock = DockStyle.Fill;
            tabControl1.TabPages[0].Text = "Sheet1";
            gridovi[0].RowHeadersWidth = 60;
            gridovi[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
           
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


            gridovi[0].TabIndex = 0;
            gridovi[0].CurrentCell = gridovi[0][0, 0];
           
            toolStripTextBox1.Anchor = AnchorStyles.Right;
            toolStripTextBox1.Dock = DockStyle.Fill;
            
            
        }

        public void delete(object o, KeyEventArgs e)
        {
            //loading = true;
            if ((e.KeyValue == (int)Keys.Delete))
            {
                foreach (DataGridViewCell c in tGrid.SelectedCells)
                {
                    KeyValuePair<int, int> index = new KeyValuePair<int, int>(c.RowIndex, c.ColumnIndex);
                    if (tCell.sveCelije.ContainsKey(index))
                    {
                        tCell.sveCelije[index].Sadrzaj = "";
                        tCell.sveCelije[index].Formula = "";

                        c.Value = "";
                        tablica_CellValueChanged(c, new DataGridViewCellEventArgs(c.ColumnIndex, c.RowIndex));
                        
                    }
                }
            }
            //loading = false;   
        }

        public void promjenaTaba(object o, EventArgs e)
        {
            TabControl tab = (TabControl)o;
            if (tab.SelectedIndex >= 0)
            {
                tGrid = gridovi[tab.SelectedIndex];
                tCell = ListaCelija[tab.SelectedIndex];
            }
        }
        //
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            int indexTaba = tabControl1.SelectedIndex;
            string s = "";
            foreach (KeyValuePair<KeyValuePair<int, int>, Cell> c in ListaCelija[indexTaba].sveCelije)
            {
                s += c.Value.Sadrzaj + " ";
            }
            MessageBox.Show(s);
        }

        private void tablica_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //ako kliknuta celija nije prazna, ispisuje se i njen sadrzaj i formula, inace samo koordinate
            //int indexTaba = tabControl1.SelectedIndex;
            KeyValuePair<int, int> index = new KeyValuePair<int, int>(e.RowIndex, e.ColumnIndex);
            if (tCell.sveCelije.ContainsKey(index))
            {
                if (tCell.sveCelije[index].Formula != "")
                    toolStripTextBox1.Text = tCell.sveCelije[index].Formula;
                else
                    toolStripTextBox1.Text = tCell.sveCelije[index].Sadrzaj;

                if (tCell.sveCelije[index].Formula != "")
                        statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() + ", " +
                                e.ColumnIndex.ToString() + "); Sadrzaj celije: " +
                                tCell.sveCelije[index].Sadrzaj +
                                "; Formula: " + tCell.sveCelije[index].Formula;

                else if (tCell.sveCelije[index].Sadrzaj != "")
                    statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() + ", " +
                                e.ColumnIndex.ToString() + "); Sadrzaj celije: " +
                                tCell.sveCelije[index].Sadrzaj; 
                    
                    else statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() +
                                ", " + e.ColumnIndex.ToString() + ")";
            }
            else
            {
                statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() + ", " + e.ColumnIndex.ToString() + ")";
                toolStripTextBox1.Text = "";
            }

            
        }

        private void tablica_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            smijemZagasiti = false;
            KeyValuePair<int, int> index = new KeyValuePair<int, int>(e.RowIndex, e.ColumnIndex);
            //int indexTaba = tabControl1.SelectedIndex;
            if (tGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null)
            {
                return;
            }

            //ako se radi o formuli
            if (tGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()[0] == '=')
            {
                if (!tCell.sveCelije.ContainsKey(index) && e.RowIndex != -1 && e.ColumnIndex != -1)
                {
                    tCell.Dodaj(e.RowIndex, e.ColumnIndex);
                }

                //spremam podatke upisane u celiju
                //string s = tGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                //tCell.sveCelije[index].Sadrzaj = s;
                toolStripTextBox1.Text = tGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                //////////////////////////////////////////////////////////////toolStripButton1_Click(null, null);
                if (toolStripTextBox1.Text == "")
                {
                    return;
                }
                int redak, stupac;
                //int indexTaba = tabControl1.SelectedIndex;
                if (otvorenaFormula)
                {
                    stupac = celijaIzForme.ColumnIndex;
                    redak = celijaIzForme.RowIndex;
                    //otvorenaFormula = false;
                }
                else
                {
                    stupac = e.ColumnIndex;
                    redak = e.RowIndex;

                }
                KeyValuePair<int, int> koordinate = new KeyValuePair<int, int>(redak, stupac);

                if (!tCell.sveCelije.ContainsKey(koordinate) && redak != -1 && stupac != -1)
                {
                    tCell.Dodaj(redak, stupac);
                }

                Cell celija = tCell.sveCelije[koordinate];

                string formula = toolStripTextBox1.Text;

                if (formula.StartsWith("=") == false)
                {
                    tCell.sveCelije[koordinate].Sadrzaj = formula;
                    if (otvorenaFormula)
                    {
                        celijaIzForme.Value = formula;
                    }
                    else
                    {
                        tGrid.SelectedCells[0].Value = formula;
                    }
                    return;
                }

                formula = formula.ToUpper();
                celija.Formula = formula;
                /////////////////////////////
                //Console.WriteLine(tmp.Pop().ToString());
                try
                {
                    celija.evaluateFormula(tCell, fje);
                }
                catch
                {
                    MessageBox.Show("Neispravna formula!");
                }
                //celija.sadrzaj = tmp.Pop().ToString();
                if (otvorenaFormula)
                {
                    celijaIzForme.Value = celija.Sadrzaj;
                    otvorenaFormula = false;
                }
                else
                {
                    tGrid[stupac, redak].Value = celija.Sadrzaj;
                }
                toolStripTextBox1.Clear();
            }

            //inace, ako se radi o obicnim brojevima
            else
            {
                //stvori novu celiju ako vec ne postoji
                //KeyValuePair<int, int> index = new KeyValuePair<int, int>(e.RowIndex, e.ColumnIndex);
                if (!tCell.sveCelije.ContainsKey(index) && e.RowIndex != -1 && e.ColumnIndex != -1)
                {
                    tCell.Dodaj(e.RowIndex, e.ColumnIndex);
                }

                //spremam podatke upisane u celiju
                string s = tGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                tCell.sveCelije[index].Sadrzaj = s; //DodajVrijednost(e.RowIndex, e.ColumnIndex, s);
                //tCell.sveCelije[index].Formula = "";
            }

            //poravnjanje brojeva i teksta

            if (tCell.sveCelije[index].Numerical)
            {
                //KeyValuePair<int, int> index = new KeyValuePair<int, int>(e.RowIndex, e.ColumnIndex);
                tGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.Alignment = DataGridViewContentAlignment.BottomRight;
            }
            else
            {
                tGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.Alignment = DataGridViewContentAlignment.BottomLeft;
            }
            foreach (Cell c in tCell.sveCelije[index].uFormuli)
            {
                try {
                    c.evaluateFormula(tCell, fje);
                
                    tGrid.Rows[c.red].Cells[c.stupac].Value = c.Sadrzaj;
                }
                catch
                {
                    MessageBox.Show("Neispravna formula!");
                }
            }
        }
         
            
        private void tablica_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            //kad se doda novi redak, ispisi redni broj u header
            //int indexTaba = tabControl1.SelectedIndex;
            tGrid.Rows[tGrid.RowCount - 1].HeaderCell.Value = tGrid.RowCount.ToString();
        }



        private void tablica_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //int indexTaba = tabControl1.SelectedIndex;
            tGrid.Columns[e.ColumnIndex].HeaderCell.Style.BackColor = Color.SlateBlue;
            for (int i = 0; i < tGrid.RowCount; i++)
                for (int j = 0; j < tGrid.ColumnCount; j++)
                    if (j == e.ColumnIndex)
                        tGrid.Rows[i].Cells[j].Style.BackColor = Color.LightSteelBlue;
                    else
                        tGrid.Rows[i].Cells[j].Style.BackColor = Color.White;
        }

        private void tablica_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //int indexTaba = tabControl1.SelectedIndex;
            tGrid.Rows[e.RowIndex].HeaderCell.Style.BackColor = Color.SlateBlue;
            for (int i = 0; i < tGrid.ColumnCount; i++)
                for (int j = 0; j < tGrid.RowCount; j++)
                    if (j == e.RowIndex)
                        tGrid.Rows[j].Cells[i].Style.BackColor = Color.LightSteelBlue;
                    else
                        tGrid.Rows[j].Cells[i].Style.BackColor = Color.White; 
        }

        public void toolStripButton1_Click(object sender, EventArgs e) // GO! Izracunaj formulu
        {
            // klik na kvacicu (ili enter) nakon unosa formule u textbox
            if (toolStripTextBox1.Text == "")
            {
                return;
            }
            int redak, stupac;
            //int indexTaba = tabControl1.SelectedIndex;
            if (otvorenaFormula)
            {
                stupac = celijaIzForme.ColumnIndex;
                redak = celijaIzForme.RowIndex;
                //otvorenaFormula = false;
            }
            else
            {
                stupac = tGrid.SelectedCells[0].ColumnIndex;
                redak = tGrid.SelectedCells[0].RowIndex;
                
            }
            KeyValuePair<int, int> koordinate = new KeyValuePair<int, int>(redak, stupac);

            if (!tCell.sveCelije.ContainsKey(koordinate) && redak != -1 && stupac != -1)
            {
                tCell.Dodaj(redak, stupac);
            }

            Cell celija = tCell.sveCelije[koordinate];

            string formula = toolStripTextBox1.Text;

            if (formula.StartsWith("=") == false)
            {
                tCell.sveCelije[koordinate].Sadrzaj = formula;
                if (otvorenaFormula)
                {
                    celijaIzForme.Value = formula;
                }
                else
                {
                    tGrid.SelectedCells[0].Value = formula;
                }
                return;
            }

            formula = formula.ToUpper();
            celija.Formula = formula;
            /////////////////////////////
                //Console.WriteLine(tmp.Pop().ToString());
            try
            {
                celija.evaluateFormula(tCell, fje);
            }
            catch
            {
                MessageBox.Show("Neispravna formula!");
            }
                //celija.sadrzaj = tmp.Pop().ToString();
            if (otvorenaFormula)
            {
                celijaIzForme.Value = celija.Sadrzaj;
                otvorenaFormula = false;
            }
            else
            {
                tGrid.SelectedCells[0].Value = celija.Sadrzaj;
            }

 
        }

        private void tablica_SelectionChanged(object sender, EventArgs e)
        {
            
            
            foreach (DataGridViewCell c in roze)
            {
                c.Style.BackColor = Color.White;
            }
                        
            List<double> argument = new List<double>();
            //int indexTaba = tabControl1.SelectedIndex;
            //double rez;
            for (int c = 0; c < tGrid.SelectedCells.Count; c++ )
            {
                KeyValuePair<int, int> index =
                    new KeyValuePair<int, int>(tGrid.SelectedCells[c].RowIndex,
                                                tGrid.SelectedCells[c].ColumnIndex);
                if (tCell.sveCelije.ContainsKey(index))
                {
                    if (tCell.sveCelije[index].Numerical)
                        argument.Add(Double.Parse(tCell.sveCelije[index].Sadrzaj));
                }

            }
            LabelSuma.Text = fje.SveFunkcije["sum"](argument).ToString();
            if (otvorenIzbor)
            {
                int index = funkcije.textBox1.Text.IndexOf("(");
                string arg = "";
                if (index > 0)
                {
                    foreach(DataGridViewCell c in tGrid.SelectedCells)
                    {
                        arg += Char.ConvertFromUtf32(c.ColumnIndex + 65) + (c.RowIndex + 1).ToString() + ";";
                    }
                    arg = arg.TrimEnd(';');
                    Match m = Regex.Match(funkcije.textBox1.Text, @"\(.*\)");
                    funkcije.textBox1.Text = funkcije.textBox1.Text.Replace(m.Value, "("+arg+")");
                }
            }
          
            
        }

        private void toolStripButton4_Click(object sender, EventArgs e) //novi tab
        {
            string s = "Sheet" + (broj_gridova + 1);
            TabPage newPage = new TabPage(s);
            tabControl1.TabPages.Add(newPage);
            gridovi.Add(new DataGridView());

            gridovi[broj_gridova].RowHeadersWidth = 60;
            tabControl1.TabPages[broj_gridova].Controls.Add(gridovi[broj_gridova]);
            gridovi[broj_gridova].BorderStyle = BorderStyle.None;
            gridovi[broj_gridova].CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellClick);
            gridovi[broj_gridova].CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellEndEdit);
            gridovi[broj_gridova].CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellValueChanged);
            gridovi[broj_gridova].SelectionChanged += new EventHandler(this.tablica_SelectionChanged);
            gridovi[broj_gridova].CellEnter += new DataGridViewCellEventHandler(this.tablica_CellEnter);
            gridovi[broj_gridova].RowsAdded += new DataGridViewRowsAddedEventHandler(tablica_RowsAdded);
            
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
            if (!(e.KeyChar == (char)Keys.Enter)) 
                return;

            toolStripButton1_Click(null, null);

            //int indexTaba = tabControl1.SelectedIndex;
            KeyValuePair<int, int> index = new KeyValuePair<int, int> (tGrid.SelectedCells[0].RowIndex, tGrid.SelectedCells[0].ColumnIndex);
            statusLabel.Text = "Koordinate celije: (" + tGrid.SelectedCells[0].ColumnIndex.ToString() + ", " +
                                tGrid.SelectedCells[0].ColumnIndex.ToString() + "); Sadrzaj celije: " +
                                tCell.sveCelije[index].Sadrzaj +
                                "; Formula: " + tCell.sveCelije[index].Formula;
            toolStripTextBox1.Text = "";
            int r = tGrid.SelectedCells[0].RowIndex;
            int s = tGrid.SelectedCells[0].ColumnIndex;
            tGrid.ClearSelection();
            tGrid.Rows[r + 1].Cells[s].Selected = true;
            
        }
      
        private void tablica_CellValueChanged(object sender, DataGridViewCellEventArgs e) //ne radi
        {
            

            KeyValuePair<int, int> index = new KeyValuePair<int, int>(e.RowIndex, e.ColumnIndex);
            if (tCell.sveCelije.Count != 0 && tCell.sveCelije.ContainsKey(index) && !loading)
            {
                foreach (Cell c in tCell.sveCelije[index].uFormuli)
                {
                    try
                    {
                        c.evaluateFormula(tCell, fje);

                        tGrid.Rows[c.red].Cells[c.stupac].Value = c.Sadrzaj;
                        if (c.Numerical)
                        {
                            tGrid.Rows[c.red].Cells[c.stupac].Style.Alignment = DataGridViewContentAlignment.BottomRight;
                        }
                        else
                        {
                            tGrid.Rows[c.red].Cells[c.stupac].Style.Alignment = DataGridViewContentAlignment.BottomLeft;
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Neispravna formula!");
                    }
                }
                
            }
        }

        private void tablica_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
             tablica_CellClick(null, e);
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            if (!otvorenIzbor)
            {

                otvorenIzbor = true;
                otvorenaFormula = true;
                funkcije = new MyExcel.Form2();
                funkcije.excel = this;

                funkcije.Show();
                
            }
            
        }

        private void fontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //klik na Font u izborniku
            //primijeni odabrano formatiranje na odabrane celije

            //int indexTaba = tabControl1.SelectedIndex;
            fontDialog1.ShowColor = true;
            fontDialog1.Font = tGrid.SelectedCells[0].Style.Font;
            fontDialog1.Color = tGrid.SelectedCells[0].Style.ForeColor;
            
            if (fontDialog1.ShowDialog() != DialogResult.Cancel)
            {
                for (int i = 0; i < tGrid.SelectedCells.Count; i++)
                {
                    tGrid.SelectedCells[i].Style.Font = fontDialog1.Font;
                    tGrid.SelectedCells[i].Style.ForeColor = fontDialog1.Color;
                }
            }
        }

        // dovrsiti praznjenje
        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //klik na New u izborniku
            loading = true;
            //pitaj korisnika zeli li spremiti promjene prije izlaza iz trenutne tablice
            MyExcel.Form3 izlaz = new MyExcel.Form3();
            izlaz.excel = this;
            DialogResult rez = izlaz.ShowDialog();
            if (rez == DialogResult.Yes)
            {
                //napravi saveAs ili save pa izadji
                if (imeFilea == null) saveAsToolStripMenuItem_Click(null, null);
                else saveToolStripMenuItem_Click(null, null);
            }
            if (rez == DialogResult.Cancel)
            {
                loading = false;
                return;
            }
            //otvori novu praznu tablicu

            while (broj_gridova > 1)
            {
                
                int tmp = broj_gridova - 1;
                
                ListaCelija.RemoveAt(tmp);
                

                tabControl1.TabPages.RemoveAt(tmp);
                //gridovi.Remove(gridovi[t.SelectedIndex]);
                gridovi.RemoveAt(tmp);
                broj_gridova--;
            }
            //isprazni tablicu prvog taba
            for (int j = 0; j < gridovi[0].Rows.Count; j++)
                for (int k = 0; k < gridovi[0].Columns.Count; k++)
                {




                    gridovi[0].Rows[j].Cells[k].Value = null;
                }


            //toolStripButton4_Click(null, null);
            ListaCelija[0].sveCelije.Clear(); //sve Cell ostaju ili nestaju?!!
            //vratiti fokus na (0,0)    
            toolStripTextBox1.Text = "";
            statusLabel.Text = "Koordinate celije: (0, 0)";
            imeFilea = "";
            gridovi[0].ClearSelection();
            gridovi[0].CurrentCell = gridovi[0][0, 0];
            loading = false;
        }



        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            //klik na Open u izborniku

            //pitaj korisnika zeli li spremiti promjene prije izlaza iz trenutne tablice
            //otvori novu praznu tablicu
            newToolStripMenuItem_Click(null, null);
            loading = true;
            //otvori open dialog
            //procitaj i prepisi tablicu iz xml-a
            openFileDialog1.Filter = "Extensible Markup Language|*.xml";
            broj_gridova = 1;
            bool pass = false;
            int red = 0, stupac = 0;
            string sadrzaj = "", formula = "", ovisnosti = "";
            bool numerical = false;
            if (openFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                imeFilea = openFileDialog1.FileName;
                //XmlTextReader reader = new XmlTextReader(imeFilea);
                XmlReader reader = XmlReader.Create(imeFilea);
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        if (reader.Name == "grid")
                        {
                            if (pass)
                            {
                                toolStripButton4_Click(null, null);
                            }
                            pass = true;
                        }
                        if (reader.Name == "celija")
                        {
                            XmlReader cell = reader.ReadSubtree();
                            while (cell.Read())
                            {
                                if (cell.NodeType == XmlNodeType.Element)
                                {
                                    if (cell.Name == "red")
                                    {
                                        red = Convert.ToInt32(cell.ReadString());
                                    }
                                    if (cell.Name == "stupac")
                                    {
                                        stupac = Convert.ToInt32(cell.ReadString());
                                    }
                                    if (cell.Name == "sadrzaj")
                                    {
                                        sadrzaj = cell.ReadString();
                                    }
                                    if (cell.Name == "formula")
                                    {
                                        formula = cell.ReadString();
                                    }
                                    if (cell.Name == "numerical")
                                    {
                                        numerical = Convert.ToBoolean(cell.ReadString());
                                    }
                                    if (cell.Name == "ovisnosti")
                                    {
                                        ovisnosti = cell.ReadString();
                                    }

                                }
                            }
                            KeyValuePair<int, int> i = new KeyValuePair<int, int>(red, stupac);
                            if (ListaCelija[broj_gridova - 1].sveCelije.ContainsKey(i) == false)
                            {
                                ListaCelija[broj_gridova - 1].Dodaj(red, stupac);
                            }
                            ListaCelija[broj_gridova - 1].sveCelije[i].Sadrzaj = sadrzaj;
                            ListaCelija[broj_gridova - 1].sveCelije[i].Formula = formula;
                            gridovi[broj_gridova - 1].Rows[red].Cells[stupac].Value = sadrzaj;
                            ListaCelija[broj_gridova - 1].sveCelije[i].PostaviOvisnosti(ListaCelija[broj_gridova - 1], ovisnosti);

                            if (ListaCelija[broj_gridova - 1].sveCelije[i].Numerical)
                            {
                                gridovi[broj_gridova - 1].Rows[red].Cells[stupac].Style.Alignment = DataGridViewContentAlignment.BottomRight;
                            }
                            else
                            {
                                gridovi[broj_gridova - 1].Rows[red].Cells[stupac].Style.Alignment = DataGridViewContentAlignment.BottomLeft;
                                
                            }

                        }
                    }
                }
                                    
                                   
                    
                
                reader.Close();
            }
            loading = false;
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //klik na Save As u izborniku
            //pitaj korisnika gdje zeli spremiti tablicu
            //spremi tablicu u xml

            saveFileDialog1.Filter = "Extensible Markup Language|*.xml";
            saveFileDialog1.Title = "Save As";
            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                //ako je za ime file-a upisano nesto razlicito od praznog stringa, spremi tablicu u xml
                if (saveFileDialog1.FileName != "")
                {
                    imeFilea = saveFileDialog1.FileName;
                    saveToolStripMenuItem_Click(null, null);
                }
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //klik na Save u izborniku
            //spremi tablicu u xml naziva imeFilea

            if (imeFilea == null || imeFilea == "")
            {
                saveAsToolStripMenuItem_Click(null, null);
                return;
            }
            XmlTextWriter xmlw = new XmlTextWriter(imeFilea, null);
            xmlw.Formatting = Formatting.Indented;
            xmlw.WriteStartDocument();
            xmlw.WriteStartElement("tablica");
            for (int i = 0; i < broj_gridova; i++)
            {
                xmlw.WriteStartElement("grid"); 
                foreach (KeyValuePair<KeyValuePair<int, int>, Cell> par in ListaCelija[i].sveCelije)
                {
                    Cell c = par.Value;
                    xmlw.WriteStartElement("celija");

                    xmlw.WriteStartElement("red");                  
                    xmlw.WriteString(c.red.ToString());
                    xmlw.WriteEndElement();

                    xmlw.WriteStartElement("stupac");
                    xmlw.WriteString(c.stupac.ToString());
                    xmlw.WriteEndElement();

                    xmlw.WriteStartElement("sadrzaj");
                    xmlw.WriteString(c.Sadrzaj);
                    xmlw.WriteEndElement();

                    xmlw.WriteStartElement("formula");
                    xmlw.WriteString(c.Formula);
                    xmlw.WriteEndElement();

                    xmlw.WriteStartElement("numerical");
                    xmlw.WriteString(c.Numerical.ToString());
                    xmlw.WriteEndElement();

                    string s="";
                    foreach (Cell a in c.uFormuli)
                    {
                        s += a.ID+";";
                    }
                    s = s.TrimEnd(';');
                    xmlw.WriteStartElement("ovisnosti");
                    xmlw.WriteString(s);
                    xmlw.WriteEndElement();

                    xmlw.WriteEndElement();
                }
                xmlw.WriteEndElement();
            }
            xmlw.WriteEndElement();
            xmlw.WriteEndDocument();
            xmlw.Close();
            smijemZagasiti = true;
        }
        
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // klik na Exit u izborniku
            // pitaj korisnika zeli li spremiti promjene prije izlaza
            // izadji iz programa

            MyExcel.Form3 izlaz = new MyExcel.Form3();
            izlaz.excel = this;
            DialogResult rez = izlaz.ShowDialog();
            if (rez == DialogResult.Yes)
            {
                //napravi saveAs ili save pa izadji
                if (imeFilea == null)
                {
                    saveAsToolStripMenuItem_Click(null, null);
                }
                else
                {
                    saveToolStripMenuItem_Click(null, null);
                }
                
            }
            else if (rez == DialogResult.No)
            {
                //samo izadji
                smijemZagasiti = true;
                Form1.ActiveForm.Close();
            }
            if (smijemZagasiti)
            {
                Form1.ActiveForm.Close();
            }
        }

     
       

      
        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            bool numTest = false;
            //int tab = tabControl1.SelectedIndex;
            foreach (DataGridViewCell c in tGrid.SelectedCells)
            {

                KeyValuePair<int, int> index = new KeyValuePair<int, int>(c.RowIndex, c.ColumnIndex);
                if (tCell.sveCelije.ContainsKey(index) && tCell.sveCelije[index].Numerical)
                {
                    numTest = true;
                    break;
                }
            }
            if (numTest)
            {
                Grafovi Slika = new Grafovi(this, tGrid, tCell);
                Slika.drawHistogram();
                

            }
           
            
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            bool numTest = false;
            //int tab = tabControl1.SelectedIndex;
            foreach (DataGridViewCell c in tGrid.SelectedCells)
            {

                KeyValuePair<int, int> index = new KeyValuePair<int, int>(c.RowIndex, c.ColumnIndex);
                if (tCell.sveCelije.ContainsKey(index) && tCell.sveCelije[index].Numerical)
                {
                    numTest = true;
                    break;
                }
            }
            if (numTest)
            {
                Grafovi Slika = new Grafovi(this, tGrid, tCell);
                Slika.drawPieChart();
            }
        
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            bool numTest = false;
            //int tab = tabControl1.SelectedIndex;
            foreach (DataGridViewCell c in tGrid.SelectedCells)
            {

                KeyValuePair<int, int> index = new KeyValuePair<int, int>(c.RowIndex,c.ColumnIndex);
                if (tCell.sveCelije.ContainsKey(index) && tCell.sveCelije[index].Numerical)
                {
                    numTest = true;
                    break;
                }
            }
            if (numTest)
            {
                Grafovi Slika = new Grafovi(this, tGrid, tCell);
                Slika.drawLineChart();
            }
            
        }

        private void toolStripTextBox1_Click(object sender, EventArgs e)
        {

        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form5 about = new Form5();
            about.ShowDialog();
        }

        private void tabControl1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            TabControl t = (TabControl)sender;
            //MessageBox.Show(t.TabPages[t.SelectedIndex].Text);
            if (t.SelectedIndex != 0)
            {
                int tmp = t.SelectedIndex;
             
                ListaCelija.RemoveAt(tmp);
               

                tabControl1.TabPages.RemoveAt(tmp);
                //gridovi.Remove(gridovi[t.SelectedIndex]);
                gridovi.RemoveAt(tmp);
                broj_gridova--;
            }
            //tGrid = gridovi[tabControl1.SelectedIndex];
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (smijemZagasiti == false)
            {
                exitToolStripMenuItem_Click(null, null);
                e.Cancel = true;
            }
        }

 

        private void Form1_Load(object sender, EventArgs e)
        {
            gridovi[0].CurrentCell = gridovi[0][0, 0];
            gridovi[0].BeginEdit(true);
        }
    }
}
