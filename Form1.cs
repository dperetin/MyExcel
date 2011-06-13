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
        Funkcije fje = new Funkcije();
        
        public Form1()
        {
            InitializeComponent();

            tabControl1.TabPages[0].Text = "Sheet1";
            tabControl1.SelectedTab = tabControl1.TabPages[0];
            gridovi.Add(new DataGridView());
            tabControl1.TabPages[0].Controls.Add(gridovi[0]);
            
            gridovi[0].BorderStyle = BorderStyle.None;
            gridovi[0].Dock = DockStyle.Fill;
            gridovi[0].RowHeadersWidth = 60;
            gridovi[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

            gridovi[0].CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellClick);
            gridovi[0].CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellValueChanged);
            gridovi[0].CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellEndEdit);
            gridovi[0].SelectionChanged += new EventHandler(this.tablica_SelectionChanged);
            gridovi[0].CellEnter += new DataGridViewCellEventHandler(this.tablica_CellEnter);
            gridovi[0].RowsAdded += new DataGridViewRowsAddedEventHandler(tablica_RowsAdded);
            //gridovi[0].ColumnHeaderMouseClick += new DataGridViewCellMouseEventHandler(tablica_ColumnHeaderMouseClick);
            //gridovi[0].RowHeaderMouseClick += new DataGridViewCellMouseEventHandler(tablica_RowHeaderMouseClick);
            
            Celije noviTab = new Celije();
            ListaCelija.Add(noviTab);

            
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

            //kad se otvori tablica, stavi fokus na (0, 0)
            gridovi[0].Focus();
            //gridovi[0].CurrentCell = gridovi[0][0, 0];
            //gridovi[0].BeginEdit(false);
            //gridovi[0].TabIndex = 0;
            //gridovi[0].BeginEdit(true);
            //gridovi[0][0, 0].Selected = true;
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
            //ako kliknuta celija nije prazna, ispisuje se i njen sadrzaj i formula, inace samo koordinate
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

        private void tablica_CellEndEdit(object sender, DataGridViewCellEventArgs e)
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

            //poravnjanje brojeva i teksta
            double r; 
            if (System.Double.TryParse(gridovi[indexTaba].Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), out r))
                gridovi[indexTaba].Rows[e.RowIndex].Cells[e.ColumnIndex].Style.Alignment = DataGridViewContentAlignment.BottomRight;
            else gridovi[indexTaba].Rows[e.RowIndex].Cells[e.ColumnIndex].Style.Alignment = DataGridViewContentAlignment.BottomLeft;
        }

        private void tablica_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            //kad se doda novi redak, ispisi redni broj u header
            int indexTaba = tabControl1.SelectedIndex;
            gridovi[indexTaba].Rows[gridovi[indexTaba].RowCount - 1].HeaderCell.Value = gridovi[indexTaba].RowCount.ToString();
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
            int indexTaba = tabControl1.SelectedIndex;
            gridovi[indexTaba].Columns[e.ColumnIndex].HeaderCell.Style.BackColor = Color.SlateBlue;
            for (int i = 0; i < gridovi[indexTaba].RowCount; i++)
                for (int j = 0; j < gridovi[indexTaba].ColumnCount; j++)
                    if (j == e.ColumnIndex) gridovi[indexTaba].Rows[i].Cells[j].Style.BackColor = Color.LightSteelBlue;
                    else gridovi[indexTaba].Rows[i].Cells[j].Style.BackColor = Color.White;
        }

        private void tablica_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int indexTaba = tabControl1.SelectedIndex;
            gridovi[indexTaba].Rows[e.RowIndex].HeaderCell.Style.BackColor = Color.SlateBlue;
            for (int i = 0; i < gridovi[indexTaba].ColumnCount; i++)
                for (int j = 0; j < gridovi[indexTaba].RowCount; j++)
                    if (j == e.RowIndex) gridovi[indexTaba].Rows[j].Cells[i].Style.BackColor = Color.LightSteelBlue;
                    else gridovi[indexTaba].Rows[j].Cells[i].Style.BackColor = Color.White; 
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            // klik na kvacicu (ili enter) nakon unosa formule u textbox
            if (toolStripTextBox1.Text == "") return;

            int indexTaba = tabControl1.SelectedIndex;
            int stupac = gridovi[indexTaba].SelectedCells[0].ColumnIndex;
            int redak = gridovi[indexTaba].SelectedCells[0].RowIndex;
            KeyValuePair<int, int> koordinate = new KeyValuePair<int, int>(redak, stupac);

            if (!ListaCelija[indexTaba].sveCelije.ContainsKey(koordinate) && redak != -1 && stupac != -1)
            {
                ListaCelija[indexTaba].Dodaj(redak, stupac);
            }

            Cell celija = ListaCelija[indexTaba].sveCelije[koordinate];
            string formula = toolStripTextBox1.Text.ToLower();
            
            double rez;
            if (System.Double.TryParse(formula, out rez))
            {
                ListaCelija[indexTaba].sveCelije[koordinate].sadrzaj = Convert.ToString(rez);
                gridovi[indexTaba].SelectedCells[0].Value = rez.ToString();
                return;
            }
            
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
            double rez;
            for (int c = 0; c < gridovi[indexTaba].SelectedCells.Count; c++ )
            {
                KeyValuePair<int, int> index =
                    new KeyValuePair<int, int>(gridovi[indexTaba].SelectedCells[c].RowIndex,
                                                gridovi[indexTaba].SelectedCells[c].ColumnIndex);
                if (ListaCelija[indexTaba].sveCelije.ContainsKey(index))
                {
                    if (System.Double.TryParse(ListaCelija[indexTaba].sveCelije[index].sadrzaj, out rez))
                    argument.Add(ListaCelija[indexTaba].sveCelije[index]);
                }

            }
            LabelSuma.Text = fje.SveFunkcije["sum"](argument).ToString(); 
        }

        private void toolStripButton4_Click(object sender, EventArgs e) //novi tab
        {
            //klik na new tab
            string s = "Sheet" + (broj_gridova + 1);
            TabPage newPage = new TabPage(s);
            tabControl1.TabPages.Add(newPage);
            tabControl1.SelectedTab = tabControl1.TabPages[broj_gridova];
            gridovi.Add(new DataGridView());
            tabControl1.TabPages[broj_gridova].Controls.Add(gridovi[broj_gridova]);
            
            gridovi[broj_gridova].RowHeadersWidth = 60;
            gridovi[broj_gridova].Dock = DockStyle.Fill;
            gridovi[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            gridovi[broj_gridova].BorderStyle = BorderStyle.None;
            
            gridovi[broj_gridova].CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellClick);
            gridovi[broj_gridova].CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellEndEdit);
            gridovi[broj_gridova].CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.tablica_CellValueChanged);
            gridovi[broj_gridova].SelectionChanged += new EventHandler(this.tablica_SelectionChanged);
            gridovi[broj_gridova].CellEnter += new DataGridViewCellEventHandler(this.tablica_CellEnter);
            gridovi[broj_gridova].RowsAdded += new DataGridViewRowsAddedEventHandler(tablica_RowsAdded);
            
            Celije noviTab = new Celije();
            ListaCelija.Add(noviTab);

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

            gridovi[broj_gridova].Focus();
            broj_gridova++;
        }

        private void toolStripTextBox1_KeyPress(object sender, KeyPressEventArgs e) 
        {
            //enter u textboxu za formule
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

        private void tablica_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
             tablica_CellClick(null, e);
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {   
            
            MyExcel.Form2 funkcije = new MyExcel.Form2();
            funkcije.excel = this;
            if (funkcije.ShowDialog() == DialogResult.OK)
            {
                toolStripTextBox1.Text = funkcije.textBox1.Text;
                toolStripButton1_Click(null, null);
                toolStripTextBox1.Clear();
            }
        }

        private void fontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //klik na Font u izborniku
            //primijeni odabrano formatiranje na odabrane celije

            int indexTaba = tabControl1.SelectedIndex;
            fontDialog1.ShowColor = true;
            fontDialog1.Font = gridovi[indexTaba].SelectedCells[0].Style.Font;
            fontDialog1.Color = gridovi[indexTaba].SelectedCells[0].Style.ForeColor;
            
            if (fontDialog1.ShowDialog() != DialogResult.Cancel)
            {
                for (int i = 0; i < gridovi[indexTaba].SelectedCells.Count; i++ )
                {
                    gridovi[indexTaba].SelectedCells[i].Style.Font = fontDialog1.Font;
                    gridovi[indexTaba].SelectedCells[i].Style.ForeColor = fontDialog1.Color;
                }
            }
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //klik na New u izborniku

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

            //otvori novu praznu tablicu
            while (broj_gridova > 1)
            {
                //izbaci sve tabove osim nultog
                ListaCelija[broj_gridova - 1].sveCelije.Clear(); //sve Cell ostaju ili nestaju?!!
                ListaCelija.Remove(ListaCelija[broj_gridova - 1]);
                gridovi[broj_gridova - 1].Controls.Remove(gridovi[broj_gridova - 1]);
                tabControl1.TabPages[broj_gridova - 1].Controls.Remove(gridovi[broj_gridova - 1]);
                tabControl1.TabPages.Remove(tabControl1.TabPages[broj_gridova - 1]);
                gridovi.Remove(gridovi[broj_gridova - 1]);
                broj_gridova--;
            }
            //isprazni tablicu prvog taba
            for (int j = 0; j < gridovi[0].Rows.Count; j++)
                for (int k = 0; k < gridovi[0].Columns.Count; k++)
                {
                    gridovi[0].Rows[j].Cells[k].Value = null;
                }
            ListaCelija[0].sveCelije.Clear(); //sve Cell ostaju ili nestaju?!!
            //vratiti fokus na (0,0)    
            toolStripTextBox1.Text = "";
            statusLabel.Text = "Koordinate celije: (0, 0)";
            imeFilea = "";
            gridovi[0].ClearSelection();
            gridovi[0].CurrentCell = gridovi[0][0, 0];
            
        }
        
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //klik na Open u izborniku

            //pitaj korisnika zeli li spremiti promjene prije izlaza iz trenutne tablice
            //otvori novu praznu tablicu
            newToolStripMenuItem_Click(null, null);

            //otvori open dialog
            //procitaj i prepisi tablicu iz xml-a
            openFileDialog1.Filter = "Extensible Markup Language|*.xml";
            broj_gridova = 1;
            bool vise_gridova = false;
            if (openFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                imeFilea = openFileDialog1.FileName;
                //XmlTextReader reader = new XmlTextReader(imeFilea);
                XmlReader reader = XmlReader.Create(imeFilea);
                while (reader.Read()) 
                {
                    if (reader.IsStartElement() && reader.Name == "tablica") reader.Read(); // Read the start tag.
                    if (reader.IsStartElement() && reader.Name == "grid") 
                    {
                        if (vise_gridova)
                            //napravi novi tab
                            toolStripButton4_Click(null, null);
                        while (reader.Read())
                        {
                            if (reader.IsStartElement() && reader.Name == "celija")
                            {
                                //prepisi elemente u celije
                                int red = Convert.ToInt32(reader.GetAttribute(0));
                                int stupac = Convert.ToInt32(reader.GetAttribute(1));
                                string sadrzaj = reader.GetAttribute(2);
                                string formula = reader.GetAttribute(3);
                                KeyValuePair<int, int> index = new KeyValuePair<int, int>(red, stupac);
                                ListaCelija[broj_gridova - 1].Dodaj(red, stupac);
                                ListaCelija[broj_gridova - 1].DodajVrijednost(red, stupac, sadrzaj);
                                ListaCelija[broj_gridova - 1].DodajFormulu(red, stupac, formula);
                                gridovi[broj_gridova - 1].Rows[red].Cells[stupac].Value = sadrzaj;

                                double r;
                                if (System.Double.TryParse(sadrzaj, out r))
                                    gridovi[broj_gridova - 1].Rows[red].Cells[stupac].Style.Alignment = DataGridViewContentAlignment.BottomRight;
                                else gridovi[broj_gridova - 1].Rows[red].Cells[stupac].Style.Alignment = DataGridViewContentAlignment.BottomLeft;
                            }
                            else
                            {
                                vise_gridova = true;
                                break;
                            }
                        }
                    }
                    
                }
                reader.Close();
            }
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

            if (imeFilea == null)
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
                    xmlw.WriteAttributeString("red", c.red.ToString());
                    xmlw.WriteAttributeString("stupac", c.stupac.ToString());
                    xmlw.WriteAttributeString("sadrzaj", c.sadrzaj);
                    xmlw.WriteAttributeString("formula", c.formula);
                    xmlw.WriteEndElement();
                }
                xmlw.WriteEndElement();
            }
            xmlw.WriteEndElement();
            xmlw.WriteEndDocument();
            xmlw.Close();
        }
        
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //klik na Exit u izborniku
            //pitaj korisnika zeli li spremiti promjene prije izlaza
            //izadji iz programa

            MyExcel.Form3 izlaz = new MyExcel.Form3();
            izlaz.excel = this;
            DialogResult rez = izlaz.ShowDialog();
            if (rez == DialogResult.Yes)
            {
                //napravi saveAs ili save pa izadji
                if (imeFilea == null) saveAsToolStripMenuItem_Click(null, null);
                else saveToolStripMenuItem_Click(null, null);
                Form1.ActiveForm.Close();
            }
            else if (rez == DialogResult.No)
            {
                //samo izadji
                Form1.ActiveForm.Close();
            }
        }
  
        bool smijesVuci = false;
        void vuci(object o, EventArgs e)
        {
            if(!smijesVuci)
                smijesVuci = true;
            else 
                smijesVuci = false;
            Panel p = (Panel)o;
            p.BringToFront();
        }

        void pomakni(object o, EventArgs e)
        {
            if (smijesVuci)
            {
                Panel tGraf = (Panel)o;
                tGraf.Location = new Point(MousePosition.X - Location.X - (150), MousePosition.Y - Location.Y - 300);
            }
        }

        void zagasi(object o, EventArgs e)
        {
            Button b = (Button)o;
            b.Parent.Dispose();
        }

      
        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            int tab = tabControl1.SelectedIndex;
            List<double> vrijednosti = new List<double>(); ;

            Panel graf = new Panel();
            graf.Size = new Size(330, 330);           
            graf.BorderStyle = BorderStyle.FixedSingle;
            graf.Location = new Point(ClientSize.Width - 400, 50);
            graf.Parent = gridovi[tab];
            graf.MouseDown += new MouseEventHandler(vuci);
            graf.MouseUp += new MouseEventHandler(vuci);
            graf.MouseMove += new MouseEventHandler(pomakni);


            Button close = new Button();
            close.Size = new Size(16, 16);
            close.FlatStyle = FlatStyle.Flat;
            
            close.Parent = graf;
            close.Location = new Point(close.Parent.Width - 19, 1);
            close.Click += new EventHandler(zagasi);
            close.BackgroundImage = imageList1.Images[0];
            close.FlatAppearance.BorderSize = 0;
            foreach (DataGridViewCell c in gridovi[tab].SelectedCells)
            {
                double r; 
            
                KeyValuePair<int, int> index = new KeyValuePair<int, int>(c.RowIndex, c.ColumnIndex);
                if (ListaCelija[tab].sveCelije.ContainsKey(index))
                {
                    if(!Double.TryParse(c.Value.ToString(), out r)) 
                        continue;
                    vrijednosti.Add(r);
                }
            }

            int hStep = 270 / (int)vrijednosti.Max();
            int sirina = 270 / vrijednosti.Count;

            List<Color> boje = new List<Color>();
            boje.Add(Color.FromArgb(62, 87, 145));
            boje.Add(Color.FromArgb(186, 61, 59));
            boje.Add(Color.FromArgb(74, 122, 69));
            boje.Add(Color.FromArgb(197, 97, 68));
            boje.Add(Color.FromArgb(111, 145, 62));
            boje.Add(Color.FromArgb(214, 154, 80));
            boje.Add(Color.FromArgb(203, 193, 76));
       

            Graphics g = graf.CreateGraphics();
            
            //GraphicsPath path = new GraphicsPath();
            Pen pen = new Pen(Color.Black, 1);
            Pen myPen = new Pen(Color.Gray, 1);
            g.DrawLine(pen, 20, 300, 310, 300);
            g.DrawLine(pen, 20, 300, 20, 20);
            int broj = 0;
            Brush crni = new SolidBrush(Color.Black);
            Font myFont = new System.Drawing.Font("Helvetica", 10);
            for (int j = 300; j >= 30; j-=hStep)
            {
                g.DrawLine(myPen, 20, j, 310, j);
                g.DrawString(broj.ToString(), myFont, crni, 10-7*((int)Math.Floor(Math.Log10(broj))), j-8);
                broj++;
            }

            int i = 0;
            foreach (double d in vrijednosti)
            {
                 Brush myBrush = new SolidBrush(boje[i%boje.Count]);
                Rectangle r = new Rectangle(30 + (i*sirina),300 - (int)d * hStep, sirina, (int)d * hStep);
                i++;
                g.FillRectangle(myBrush, r);
                g.DrawRectangle(pen, r);
            }
            
           
            
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            int tab = tabControl1.SelectedIndex;
            List<double> vrijednosti = new List<double>(); ;

            Panel graf = new Panel();
            graf.Size = new Size(330, 330);
            graf.BorderStyle = BorderStyle.FixedSingle;
            graf.Location = new Point(ClientSize.Width - 400, 50);
            graf.Parent = gridovi[tab];
            graf.MouseDown += new MouseEventHandler(vuci);
            graf.MouseUp += new MouseEventHandler(vuci);
            graf.MouseMove += new MouseEventHandler(pomakni);

            Button close = new Button();
            close.Size = new Size(16, 16);
            close.FlatStyle = FlatStyle.Flat;
            close.Parent = graf;
            close.Location = new Point(close.Parent.Width - 19, 1);
            close.Click += new EventHandler(zagasi);
            close.BackgroundImage = imageList1.Images[0];
            close.FlatAppearance.BorderSize = 0;
            foreach (DataGridViewCell c in gridovi[tab].SelectedCells)
            {
                double r;

                KeyValuePair<int, int> index = new KeyValuePair<int, int>(c.RowIndex, c.ColumnIndex);
                if (ListaCelija[tab].sveCelije.ContainsKey(index))
                {
                    if (!Double.TryParse(c.Value.ToString(), out r))
                        continue;
                    vrijednosti.Add(r);
                }
            }

            
            List<Color> boje = new List<Color>();
            boje.Add(Color.FromArgb(62, 87, 145));
            boje.Add(Color.FromArgb(186, 61, 59));
            boje.Add(Color.FromArgb(74, 122, 69));
            boje.Add(Color.FromArgb(197, 97, 68));
            boje.Add(Color.FromArgb(111, 145, 62));
            boje.Add(Color.FromArgb(214, 154, 80));
            boje.Add(Color.FromArgb(203, 193, 76));


            Graphics g = graf.CreateGraphics();
            g.SmoothingMode = SmoothingMode.AntiAlias;
            int n = (int)vrijednosti.Sum();
            Rectangle rect = new Rectangle(50, 50, 230, 230);
            int i = 0;
            float startAngle = 0.0F;
            float sweepAngle = 360.0F / n * (float)vrijednosti[0];
            for (int k = 1; k < vrijednosti.Count; k++ )
            {
                Brush myBrush = new SolidBrush(boje[i % boje.Count]);



                // Fill pie to screen.
                g.FillPie(myBrush, rect, startAngle, sweepAngle);

                i++;
                startAngle += sweepAngle;
                sweepAngle = 360.0F / n * (float)vrijednosti[k];
            }
            Brush myBrush2 = new SolidBrush(boje[i % boje.Count]);
            


            // Fill pie to screen.
            g.FillPie(myBrush2, rect, startAngle, sweepAngle);
            //GraphicsPath path = new GraphicsPath();
           
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            int tab = tabControl1.SelectedIndex;
            List<double> vrijednosti = new List<double>(); ;

            Panel graf = new Panel();
            graf.Size = new Size(330, 330);
            graf.BorderStyle = BorderStyle.FixedSingle;
            graf.Location = new Point(ClientSize.Width - 400, 50);
            graf.Parent = gridovi[tab];
            graf.MouseDown += new MouseEventHandler(vuci);
            graf.MouseUp += new MouseEventHandler(vuci);
            graf.MouseMove += new MouseEventHandler(pomakni);
            
            Button close = new Button();
            close.Size = new Size(16, 16);
            close.FlatStyle = FlatStyle.Flat;
            close.Parent = graf;
            close.Location = new Point(close.Parent.Width - 19, 1);
            close.Click += new EventHandler(zagasi);
            close.BackgroundImage = imageList1.Images[0];
            close.FlatAppearance.BorderSize = 0;
         
     
            foreach (DataGridViewCell c in gridovi[tab].SelectedCells)
            {
                double r;

                KeyValuePair<int, int> index = new KeyValuePair<int, int>(c.RowIndex, c.ColumnIndex);
                if (ListaCelija[tab].sveCelije.ContainsKey(index))
                {
                    if (!Double.TryParse(c.Value.ToString(), out r))
                        continue;
                    vrijednosti.Add(r);
                }
            }

            int hStep = 270 / (int)vrijednosti.Max();
            int sirina = 270 / vrijednosti.Count;

            List<Color> boje = new List<Color>();
            boje.Add(Color.FromArgb(62, 87, 145));
            boje.Add(Color.FromArgb(186, 61, 59));
            boje.Add(Color.FromArgb(74, 122, 69));
            boje.Add(Color.FromArgb(197, 97, 68));
            boje.Add(Color.FromArgb(111, 145, 62));
            boje.Add(Color.FromArgb(214, 154, 80));
            boje.Add(Color.FromArgb(203, 193, 76));


            Graphics g = graf.CreateGraphics();
            g.SmoothingMode = SmoothingMode.AntiAlias;
            //GraphicsPath path = new GraphicsPath();
            Pen pen = new Pen(Color.Black, 1);
            Pen myPen = new Pen(Color.Gray, 1);
            g.DrawLine(pen, 20, 300, 310, 300);
            g.DrawLine(pen, 20, 300, 20, 20);
            int broj = 0;
            Brush crni = new SolidBrush(Color.Black);
            Font myFont = new System.Drawing.Font("Helvetica", 10);
            for (int j = 300; j >= 30; j -= hStep)
            {
                g.DrawLine(myPen, 20, j, 310, j);
                g.DrawString(broj.ToString(), myFont, crni, 10 - 7 * ((int)Math.Floor(Math.Log10(broj))), j - 8);
                broj++;
            }

            int i = 0;
            for(int d = 0; d < vrijednosti.Count-1; d++)
            {
                Pen crta = new Pen(boje[0], 3);
                //Pen tocka = new Pen(boje[1], 5);
                g.DrawLine(crta, 30 + (sirina / 2) + (i * sirina), 300 - (int)vrijednosti[d] * hStep, 30 + (sirina / 2) + ((i + 1) * sirina), 300 - (int)vrijednosti[d + 1] * hStep);
                
                //g.DrawEllipse(tocka, 30 + (sirina / 2) + ((i + 1) * sirina), 300 - (int)vrijednosti[d + 1] * hStep, 5, 5);
                i++;
                //g.FillRectangle(myBrush, r);
                //g.DrawRectangle(pen, r);
            }
            //Rectangle myRectangle = new Rectangle(20, 20, 250, 200);
            i = 0;
            Pen tocka = new Pen(boje[1], 4);
            g.DrawEllipse(tocka, 28 + (sirina / 2) + (i * sirina), 298 - (int)vrijednosti[0] * hStep, 4, 4);
            for (int d = 0; d < vrijednosti.Count - 1; d++)
            {
               
               
               

                g.DrawEllipse(tocka, 28 + (sirina / 2) + ((i + 1) * sirina), 298 - (int)vrijednosti[d + 1] * hStep, 4, 4);
                i++;
                //g.FillRectangle(myBrush, r);
                //g.DrawRectangle(pen, r);
            }



            // g.DrawString("Hello C#", myFont, myBrush, 30, 30);

            //g.DrawRectangle(pen, myRectangle);

            // g.DrawPath(pen, path);
        }

    }
}
