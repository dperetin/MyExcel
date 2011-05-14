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
        List<Celije> ListaTablica = new List<Celije>();
        Celije ListaCelija = new Celije();
        Funkcije fje = new Funkcije();
        private int brojTabova = 1;
        private TabControl tabControlNew;
        private TabPage tabPageNew;
        private DataGridView novaTablica;

        public Form1()
        {
            InitializeComponent();
            // popunjavam tablicu praznom redovima
            string[] red = {};
            for (int i = 1; i < 100; i++)
            {
                tablica1.Rows.Add(red);
                tablica1.Rows[i - 1].HeaderCell.Value = i.ToString(); 
            }
            ListaTablica.Add(ListaCelija);
        
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            string s = "";
            int indexTaba = this.tabControl1.SelectedTab.TabIndex;
            foreach (KeyValuePair<KeyValuePair<int, int>, Cell> c in ListaTablica[indexTaba].sveCelije)
            {
                s += c.Value.sadrzaj + " ";
            }
            MessageBox.Show(s);
        }

        private void tablica1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int indexTaba = this.tabControl1.SelectedTab.TabIndex;
            //ako kliknuta celija nije prazna, ispisuje se i njen sadrzaj, inace samo koordinate
            KeyValuePair<int, int> index = new KeyValuePair<int, int>(e.RowIndex, e.ColumnIndex);
            if (ListaTablica[indexTaba].sveCelije.ContainsKey(index))
            {
                if (ListaTablica[indexTaba].sveCelije[index].DajFormulu() != null)
                    statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() + ", "
                        + e.ColumnIndex.ToString() + "); Formula " + ListaTablica[indexTaba].sveCelije[index].DajFormulu();
                else if (ListaTablica[indexTaba].sveCelije[index].DajVrijednostCelije() != "")
                    statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() + ", "
                        + e.ColumnIndex.ToString() + "); Sadrzaj celije: " + ListaTablica[indexTaba].sveCelije[index].DajVrijednostCelije();
                    else statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() + 
                    ", " + e.ColumnIndex.ToString() + ")";
            }
            else statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() + ", " + e.ColumnIndex.ToString() + ")";

            //izbrisi boju svih celija koje su prije bile kliknute
            //for (int i = 0; i < ListaCelija.brojStupaca; i++)
            //    for (int j = 0; j < ListaCelija.brojRedaka; j++) 
            //        tablica.Rows[j].Cells[i].Style.BackColor = Color.White;
            //for (int i = 0; i < ListaCelija.brojStupaca; i++)
            //    tablica.Columns[i].HeaderCell.Style.BackColor = Control.DefaultBackColor;
            //for (int j = 0; j < ListaCelija.brojRedaka; j++)
            //    tablica.Rows[j].HeaderCell.Style.BackColor = Control.DefaultBackColor;
            //if ((e.ColumnIndex != -1) && (e.RowIndex != -1))
            //tablica.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.LightSteelBlue;
            
        }

     
        private void novaTablica_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int indexTaba = this.tabControl1.SelectedTab.TabIndex;
            //ako kliknuta celija nije prazna, ispisuje se i njen sadrzaj, inace samo koordinate
            KeyValuePair<int, int> index = new KeyValuePair<int, int>(e.RowIndex, e.ColumnIndex);
            if (ListaTablica[indexTaba].sveCelije.ContainsKey(index))
            {
                //toolStripTextBox1.Text = ListaTablica[indexTaba].sveCelije[index].formula;
                if (ListaTablica[indexTaba].sveCelije[index].DajFormulu() != null)
                    statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() + ", "
                        + e.ColumnIndex.ToString() + "); Formula " + ListaTablica[indexTaba].sveCelije[index].DajFormulu();
                else if (ListaTablica[indexTaba].sveCelije[index].DajVrijednostCelije() != "")
                    statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() + ", "
                        + e.ColumnIndex.ToString() + "); Sadrzaj celije: " + ListaTablica[indexTaba].sveCelije[index].DajVrijednostCelije();
                else statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() +
                ", " + e.ColumnIndex.ToString() + ")";
            }
            else statusLabel.Text = "Koordinate celije: (" + e.RowIndex.ToString() + ", " + e.ColumnIndex.ToString() + ")";
        }

        void tablica1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int indexTaba = this.tabControl1.SelectedTab.TabIndex;
            if ( indexTaba == 0 && tablica1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null) return;
            
            //stvori novu celiju ako vec ne postoji
            KeyValuePair<int, int> index = new KeyValuePair<int, int>(e.RowIndex, e.ColumnIndex);
            if (!ListaTablica[indexTaba].sveCelije.ContainsKey(index) && e.RowIndex != -1 && e.ColumnIndex != -1)
            {
                ListaTablica[indexTaba].Dodaj(e.RowIndex, e.ColumnIndex);
            }
           
            //spremam podatke upisane u celiju
            string s = tablica1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            ListaTablica[indexTaba].DodajVrijednost(e.RowIndex, e.ColumnIndex, s);
        }

        void novaTablica_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int indexTaba = this.tabControl1.SelectedTab.TabIndex;
            string imeTablice = "tablica" + (indexTaba + 1);
            if (novaTablica.Name == imeTablice && novaTablica.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null) return;

            //stvori novu celiju ako vec ne postoji
            KeyValuePair<int, int> index = new KeyValuePair<int, int>(e.RowIndex, e.ColumnIndex);
            if (!ListaTablica[indexTaba].sveCelije.ContainsKey(index) && e.RowIndex != -1 && e.ColumnIndex != -1)
            {
                ListaTablica[indexTaba].Dodaj(e.RowIndex, e.ColumnIndex);
            }

            //spremam podatke upisane u celiju
            string s = novaTablica.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(); // !!!!! null string
            ListaTablica[indexTaba].DodajVrijednost(e.RowIndex, e.ColumnIndex, s);
        }
        //private void tablica_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        //{
        //    //kad se doda novi redak, ispisi redni broj u header
        //    tablica.Rows[ListaCelija.brojRedaka - 1].HeaderCell.Value = ListaCelija.brojRedaka.ToString();
        //    ListaCelija.brojRedaka++;
        //}

        //private void tablica_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        //{
        //    //ako izbacimo i-ti redak, moramo promijeniti brojeve headera od i+1 nadalje
        //    ListaCelija.brojRedaka--;
        //    for (int j = 0; j < ListaCelija.brojRedaka - 1; j++)
        //        tablica.Rows[j].HeaderCell.Value = Convert.ToString(j + 1);
        //}

        //private void tablica_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        //{
        //    tablica.Columns[e.ColumnIndex].HeaderCell.Style.BackColor = Color.SlateBlue;
        //    for (int i = 0; i < ListaCelija.brojRedaka; i++)
        //        for (int j = 0; j < ListaCelija.brojStupaca; j++)
        //            if (j == e.ColumnIndex) tablica.Rows[i].Cells[j].Style.BackColor = Color.LightSteelBlue;
        //            else tablica.Rows[i].Cells[j].Style.BackColor = Color.White;
        //}

        //private void tablica_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        //{
        //    tablica.Rows[e.RowIndex].HeaderCell.Style.BackColor = Color.SlateBlue;
        //    for (int i = 0; i < ListaCelija.brojStupaca; i++)
        //        for (int j = 0; j < ListaCelija.brojRedaka; j++)
        //            if (j == e.RowIndex) tablica.Rows[j].Cells[i].Style.BackColor = Color.LightSteelBlue;
        //            else tablica.Rows[j].Cells[i].Style.BackColor = Color.White;
        //}

       
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            int indexTaba = this.tabControl1.SelectedTab.TabIndex;
           
            int stupac = tablica1.SelectedCells[0].ColumnIndex;
            int redak = tablica1.SelectedCells[0].RowIndex;

            KeyValuePair<int, int> koordinate = new KeyValuePair<int, int>(redak, stupac);

            if (!ListaTablica[indexTaba].sveCelije.ContainsKey(koordinate) && redak != -1 && stupac != -1)
            {
                ListaTablica[indexTaba].Dodaj(redak, stupac);
            }

            Cell celija = ListaTablica[indexTaba].sveCelije[koordinate];

            //tablica.SelectedCells[0].Value = celija.sadrzaj;

            string formula = toolStripTextBox1.Text;
            celija.formula = formula;
            string fja = Regex.Match(formula, @"=\w*[(]").Value;
            string rje = Regex.Match(formula, "[(].*[)]").Value;
            string a, b;
            a = rje.TrimEnd(')');
            a = a.TrimStart('(');
            b = fja.TrimEnd('(');
            b = b.TrimStart('=');
            celija.sadrzaj = fje.SveFunkcije[b](ListaTablica[indexTaba].parsiraj(a)).ToString();

            tablica1.SelectedCells[0].Value = celija.sadrzaj; // ovo ce raditi samo za prvu tablicu
        }


        //dupli klik na neki od tabova stvara novi
        //iz nekog razloga na drugi tab ne stavi tablicu, tek na treci i dalje

        private void tabControl1_DoubleClick(object sender, EventArgs e)
        {
            //dodaj novi tab i ubaci tablicu u njega
            tabPageNew = new TabPage();
            tabControlNew = new TabControl();
            tabControl1.TabPages.Add(tabPageNew);
            Controls.Add(tabControlNew);
            brojTabova++;

            tabPageNew.Controls.Add(novaTablica);
            tabPageNew.Location = new System.Drawing.Point(4, 22);
            tabPageNew.Name = "tabPage" + brojTabova;
            tabPageNew.Padding = new System.Windows.Forms.Padding(3);
            tabPageNew.Size = new System.Drawing.Size(645, 401);
            tabPageNew.TabIndex = brojTabova - 1;
            tabPageNew.Text = "Sheet" + brojTabova;
            tabPageNew.UseVisualStyleBackColor = true;
            
            // ubaci novu tablicu u tab
            A1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            B1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            C1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            D1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            E1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            F1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            G1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            H1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            I1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            J1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            K1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            L1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            M1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            N1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            O1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            P1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            Q1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            R1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            S1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            T1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            U1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            V1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            W1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            X1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            Y1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            Z1 = new System.Windows.Forms.DataGridViewTextBoxColumn();

            novaTablica = new System.Windows.Forms.DataGridView();
            novaTablica.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            novaTablica.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            A1, B1, C1, D1, E1, F1, G1, H1, I1, J1, K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, U1, V1, W1 ,X1, Y1, Z1});
            novaTablica.Dock = DockStyle.Fill;
            novaTablica.Name = "tablica" + brojTabova;
            // gore je bilo tabPageNew.Text = "Sheet" + brojTabova; i on napravi tablica2, sheet3
            novaTablica.Location = new System.Drawing.Point(3, 3);
            novaTablica.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.Raised;
            novaTablica.Size = new System.Drawing.Size(639, 395);
            novaTablica.TabIndex = brojTabova - 1; 

            A1.Name = "A";            A1.HeaderText = "A";            B1.Name = "B";            B1.HeaderText = "B"; 
            C1.Name = "C";            C1.HeaderText = "C";            D1.Name = "D";            D1.HeaderText = "D";
            E1.Name = "E";            E1.HeaderText = "E";            F1.Name = "F";            F1.HeaderText = "F";
            G1.Name = "G";            G1.HeaderText = "G";            H1.Name = "H";            H1.HeaderText = "H";
            I1.Name = "I";            I1.HeaderText = "I";            J1.Name = "J";            J1.HeaderText = "J";
            K1.Name = "K";            K1.HeaderText = "K";            L1.Name = "L";            L1.HeaderText = "L";
            M1.Name = "M";            M1.HeaderText = "M";            N1.Name = "N";            N1.HeaderText = "N";
            O1.Name = "O";            O1.HeaderText = "O";            P1.Name = "P";            P1.HeaderText = "P";
            Q1.Name = "Q";            Q1.HeaderText = "Q";            R1.Name = "R";            R1.HeaderText = "R";
            S1.Name = "S";            S1.HeaderText = "S";            T1.Name = "T";            T1.HeaderText = "T";
            U1.Name = "U";            U1.HeaderText = "U";            V1.Name = "V";            V1.HeaderText = "V";
            W1.Name = "W";            W1.HeaderText = "W";            X1.Name = "X";            X1.HeaderText = "X";
            Y1.Name = "Y";            Y1.HeaderText = "Y";            Z1.Name = "Z";            Z1.HeaderText = "Z";

            novaTablica.EnableHeadersVisualStyles = false;
            novaTablica.RowHeadersWidth = 60;
            Controls.Add(novaTablica); // nisam sigurna je li potrebno oboje ili samo jedno od ovoga
            novaTablica.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.novaTablica_CellClick);
            novaTablica.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.novaTablica_CellEndEdit);
            //novaTablica.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(this.tablica_RowsAdded);
            //novaTablica.RowsRemoved += new System.Windows.Forms.DataGridViewRowsRemovedEventHandler(this.tablica_RowsRemoved);
            //novaTablica.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.tablica_ColumnHeaderMouseClick);
            //novaTablica.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.tablica_RowHeaderMouseClick);

            string[] red = { };
            for (int i = 1; i < 100; i++)
            {
                novaTablica.Rows.Add(red);
                novaTablica.Rows[i - 1].HeaderCell.Value = i.ToString();
            }
            Celije novaListaCelija = new Celije();
            ListaTablica.Add(novaListaCelija);  
        }
    }
}

// ovo mi treba
// string s = tablica.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
// tablica.Columns[i].HeaderCell.Style.BackColor = Control.DefaultBackColor;

