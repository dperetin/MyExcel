﻿using System;
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
    class Grafovi
    {
        // za pomicanje grafa
        Point mouseDownPoint;
        bool dragging;

        // labele
        public string naslov = "";
        public string xOs = "";
        public string yOs = "";
        bool prviStupacOznake = false;
        
        int prviStupac = 1000;
        
        int tipGrafa;
        // vrijednosti koje se crtaju na grafu
        List<double> vrijednosti = new List<double>();
        static int brGrafova;
        // desni klik meni
        ContextMenuStrip strip = new ContextMenuStrip();

        // celije ciji sadrzaj se crta
        public Celije ListaCelija;
        private List<Cell> CelijeZaPlot = new List<Cell>();
        List<Cell> Kategorije = new List<Cell>();
        List<string> imenaKategorija = new List<string>();

        // forma na kojoj se prikazuje panel
        public Form1 f;

        // forma za odabir svojstava grafa
        public Form4 svojstva = new Form4();

        // PictureBox na kojem se crta graf
        PictureBox graf = new PictureBox();

        // sama slika
        Bitmap b;
        Graphics g;
        
        // gumb za gasenje panela
        Button close = new Button();

        // boje koje koristimo u crtanju grafova
        List<Color> boje = new List<Color>();

        // fja za crtanje grafa, postavlja se u odnosu na tip grafa,
        // da se prilikom osvjezavanja uvijek nacrta pravi graf

        nacrtajGraf fjaZaCrtanje;

        public delegate void nacrtajGraf();

        void zagasi(object o, EventArgs e)
        {
            //Button b = (Button)o;
            graf.Dispose();
        }

        void desniKlikMeni(object o, MouseEventArgs e)
        {
            
            if (e.Button == MouseButtons.Right)
            {
                Point pt = graf.PointToScreen(e.Location);
                strip.Show(pt);
            }

        }
        
        void izbor(object o, EventArgs e)
        {
            svojstva = new Form4();
            svojstva.textBox1.Text = naslov;
            svojstva.textBox2.Text = xOs;
            svojstva.textBox3.Text = yOs;
            svojstva.checkBox1.Checked = prviStupacOznake;
            if (tipGrafa == 3)
            {
                svojstva.textBox2.Enabled = false;
                svojstva.textBox3.Enabled = false;
                svojstva.checkBox1.Enabled = false;
            }
            svojstva.Show();
            svojstva.button1.Click += new EventHandler(postaviNaslov);
            svojstva.checkBox1.CheckedChanged += new EventHandler(kvacica);
           
        }
        void kvacica(object o, EventArgs e)
        {
            if (prviStupacOznake) 
                prviStupacOznake = false;
            else
                prviStupacOznake = true;
        }
        void postaviNaslov(object o, EventArgs e)
        {
            naslov = svojstva.textBox1.Text;
            xOs = svojstva.textBox2.Text;
            yOs = svojstva.textBox3.Text;
            
            svojstva.Close();
            //g.Restore(stanjePrijeOznaka);
            g.Clear(Color.White);
            try
            {
                fjaZaCrtanje();
            }
            catch
            {
                MessageBox.Show("Nije selektirano dovoljno podataka!");
            }
            Brush crni = new SolidBrush(Color.Black);
            Font naslovFont = new System.Drawing.Font("Helvetica", 10);
            Font textFont = new System.Drawing.Font("Helvetica", 8);
            g.DrawString(naslov, naslovFont, crni, 215 - naslov.Length * 5 / 2, 10);
            g.DrawString(xOs, textFont, crni, 60 + 135 - xOs.Length * 5 / 2, 314);
            g.DrawString(yOs, textFont, crni, 1, 135 - yOs.Length * 5 / 2, new StringFormat(StringFormatFlags.DirectionVertical));
            graf.Refresh();
        }
        void kopiraj(object o, EventArgs e)
        {
            Clipboard.SetImage(b);
        }

        void spremi(object o, EventArgs e)
        {
            SaveFileDialog s = new SaveFileDialog();
            s.FileName = naslov + ".bmp";
            s.Title = "Spremanje slike";
            s.Filter = "Bitmap slika|*.bmp";

            if (s.ShowDialog() != DialogResult.Cancel)
            {
                //ako je za ime file-a upisano nesto razlicito od praznog stringa, spremi tablicu u xml
                if (s.FileName != "")
                {
                    b.Save(s.FileName);
                }
            }
        }

        public Grafovi(Form1 f, DataGridView grid, Celije celije)
        {
            
            this.ListaCelija = celije;
            this.f = f;
            // inicijaliziramo boje
            boje.Add(Color.FromArgb(62, 87, 145));
            boje.Add(Color.FromArgb(186, 61, 59));
            boje.Add(Color.FromArgb(74, 122, 69));
            boje.Add(Color.FromArgb(197, 97, 68));
            boje.Add(Color.FromArgb(111, 145, 62));
            boje.Add(Color.FromArgb(214, 154, 80));
            boje.Add(Color.FromArgb(203, 193, 76));

            // pomicanje
            graf.MouseDown += _MouseDown;
            graf.MouseMove += _MouseMove;
            graf.MouseUp += _MouseUp;
            
            graf.Size = new Size(430, 350);
            graf.BorderStyle = BorderStyle.FixedSingle;
            graf.Location = new Point(f.ClientSize.Width - 500 - brGrafova * 20, 50 + brGrafova * 20);
            graf.Parent = grid;
            graf.MouseClick += new MouseEventHandler(desniKlikMeni);
            graf.BringToFront();

            strip.Items.Add("Postavke");
            strip.Items.Add("Kopiraj");
            strip.Items.Add("Spremi u datoteku");
            strip.Items.Add("Izbrisi");
            
            strip.Items[0].Image = f.imageList1.Images[1];
            strip.Items[1].Image = f.imageList1.Images[2];
            strip.Items[2].Image = f.imageList1.Images[3];
            strip.Items[3].Image = f.imageList1.Images[0];

            strip.Items[0].Click += new EventHandler(izbor);
            strip.Items[1].Click += new EventHandler(kopiraj);
            strip.Items[2].Click += new EventHandler(spremi);
            strip.Items[3].Click += new EventHandler(zagasi);
           // graf.Controls.Add(strip);

            close.Size = new Size(16, 16);
            close.FlatStyle = FlatStyle.Flat;
            close.Parent = graf;
            close.Location = new Point(close.Parent.Width - 19, 1);
            close.Click += new EventHandler(zagasi);
            close.BackgroundImage = f.imageList1.Images[0];
            close.FlatAppearance.BorderSize = 0;
            foreach (DataGridViewCell c in grid.SelectedCells)
            {
                //double r;

                KeyValuePair<int, int> index = new KeyValuePair<int, int>(c.RowIndex, c.ColumnIndex);
                if (ListaCelija.sveCelije.ContainsKey(index))
                {
                    if (ListaCelija.sveCelije[index].Numerical == false)
                    {
                        Kategorije.Add(ListaCelija.sveCelije[index]);
                        //if (ListaCelija.sveCelije[index].stupac < prviStupac)
                            //prviStupac = ListaCelija.sveCelije[index].stupac;
                        continue;
                    }
                    vrijednosti.Add(Double.Parse(ListaCelija.sveCelije[index].Sadrzaj));
                    CelijeZaPlot.Add(ListaCelija.sveCelije[index]);
                    if (ListaCelija.sveCelije[index].stupac < prviStupac)
                        prviStupac = ListaCelija.sveCelije[index].stupac;
                }
            }
            CelijeZaPlot.Sort();
            brGrafova++;
        }
        public void drawHistogram()
        {
            tipGrafa = 1;
            fjaZaCrtanje = new nacrtajGraf(drawHistogram);
            //int stupacOznaka = ;
            List<int> stupci = new List<int>();

            // broj tocaka u svakom od tih stupaca, da znamo odrediti
            // koliko ce graf biti sirok
            List<int> brojTocaka = new List<int>();
            List<Cell> oznakeXosi = new List<Cell>();
            bool stringOznake = false;
            int pomakniStupac = 0;
            // popunjavanje gornjih listi
            stupci.Add(CelijeZaPlot[0].stupac);
            brojTocaka.Add(1);
            foreach (Cell c in CelijeZaPlot)
            {
                if (stupci.Contains(c.stupac))
                {
                   // brojTocaka[stupci.IndexOf(c.stupac)]++;
                    continue;
                }
                else
                {
                    stupci.Add(c.stupac);
                    brojTocaka.Add(1);
                }
            }
            stupci.Sort();
           // bool skip = false;
            if (prviStupacOznake)
            {
                pomakniStupac = 1;
                foreach (KeyValuePair<KeyValuePair<int, int>, Cell> c in ListaCelija.sveCelije)
                {
                    if (c.Value.stupac == prviStupac)
                    {
                        oznakeXosi.Add(c.Value);
                        if (c.Value.Numerical == false)
                        {
                            stringOznake = true;
                        }
                    }
                }
                if (!stringOznake)
                {
                    vrijednosti.Clear();
                    foreach (Cell c in CelijeZaPlot)
                    {
                        if (c.stupac != prviStupac)
                        {
                            vrijednosti.Add(Double.Parse(c.Sadrzaj));
                        }
                    }
                }
            }
            else
            {
                foreach (Cell c in CelijeZaPlot)
                {

                    vrijednosti.Add(Double.Parse(c.Sadrzaj));

                }
            }
            if (prviStupacOznake)
            {
                // stupacOznaka = stupci[0];
                stupci.RemoveAt(0);
                brojTocaka.RemoveAt(0);
            }
            foreach (Cell c in CelijeZaPlot)
            {

                if (stupci.IndexOf(c.stupac)>=0)
                    brojTocaka[stupci.IndexOf(c.stupac)]++;
                
                
            }
            
            foreach (int i in stupci)
            {
                bool found = false;
                foreach (Cell c in Kategorije)
                {
                    if (c.stupac == i)
                    {
                        imenaKategorija.Add(c.Sadrzaj);
                        found = true;
                    }
                }
                if (!found)
                {
                    imenaKategorija.Add("Stupac " + Convert.ToChar(i + 65 + pomakniStupac).ToString());
                }
            }

            //double absMin = Math.Abs()
            double span = vrijednosti.Max();// - vrijednosti.Min(); // raspon vrijednosti
            if (span == 0) span = 1;
            double pixelStep = 270 / span;                          // koliko vrijednosi nosi jedan pixel na grafu
            int sirina = (270 - brojTocaka.Max() * 10) / ((brojTocaka.Max() - 1) * brojTocaka.Count);                 // razmak izmedu tocaka

            
            //Graphics g = graf.CreateGraphics();
            Brush crni = new SolidBrush(Color.Black);
            Font naslovFont = new System.Drawing.Font("Helvetica", 10);
            Font textFont = new System.Drawing.Font("Helvetica", 8);
            //g.DrawString(naslov, naslovFont, crni, 215 - naslov.Length * 5 / 2, 10);
            //g.DrawString(xOs, textFont, crni, 60 + 135 - xOs.Length * 5 / 2, 310);
            //g.DrawString(yOs, textFont, crni, 1, 215 - yOs.Length * 5 / 2, new System.Drawing.StringFormat(StringFormatFlags.DirectionVertical));
            //GraphicsPath path = new GraphicsPath();
            Pen pen = new Pen(Color.Black, 1);
            Pen pozadina = new Pen(Color.White, 1);
            Pen myPen = new Pen(Color.Gray, 1);
            b = new Bitmap(430, 350);
            //Graphics g = graf.CreateGraphics();
            g = Graphics.FromImage(b);
            graf.Image = b;
            g.FillRectangle(new SolidBrush(Color.White), 0, 0, 430, 350);
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            g.DrawLine(pen, 50, 300, 310, 300);
            g.DrawLine(pen, 50, 300, 50, 20);
            double broj = vrijednosti.Max();
            double korak = (vrijednosti.Max()) / 10;
            
            //int brPixela = 270 / (int)((vrijednosti.Max() - vrijednosti.Min()) / 10);
            int j;
            for (j = 30; j < 300; j += 27)
            {
                string displayBroj = (broj).ToString().Substring(0, Math.Min((broj).ToString().Length, 5));
                displayBroj = displayBroj.TrimEnd('.');
                g.DrawLine(myPen, 50, j, 310, j);
                g.DrawString(displayBroj, textFont, crni, 15, j - 8);
                broj -= korak;
              
            }
            g.DrawLine(myPen, 50, j, 310, j);
            g.DrawString((0).ToString(), textFont, crni, 15, j - 8);
           // g.DrawString((broj).ToString().Substring(0, 4), textFont, crni, 10, j - 8);

            if (!prviStupacOznake)
            {
                for (int i = 1; i < brojTocaka.Max(); i++)
                {
                    g.DrawString(i.ToString(), textFont, crni,
                        45 + sirina * stupci.Count / 2 + (i - 1) * sirina * stupci.Count + 10 * i, 300);
                }
            }
            else
            {
                int i = 1;
                foreach (Cell c in oznakeXosi)
                {
                   // g.DrawString(c.Sadrzaj, textFont, crni,
                       // 56 + sirina * (i - 1) + (sirina / 2), 304);
                    g.DrawString(c.Sadrzaj, textFont, crni,
                        45 + sirina * stupci.Count / 2 + (i - 1) * sirina * stupci.Count + 10 * i, 300);
                    i++;
                }
            }
            
            int k = 0;
            foreach (int stupac in stupci)
            {
             /*   if (skip)
                {
                    skip = false;
                    continue;
                }*/
                List<double> vrStup = new List<double>();

                for (int In = CelijeZaPlot.Count - 1; In >= 0; In--)
                {
                    if (CelijeZaPlot[In].stupac == stupac)
                    {
                        vrStup.Add(Double.Parse(CelijeZaPlot[In].Sadrzaj));
                    }
                }
                Brush myBrush = new SolidBrush(boje[k % boje.Count]);
                int i = 0;
                for (int d = 0; d < vrStup.Count; d++)
                {
                    
                    Rectangle r = new Rectangle(60 + (i * sirina)+ 10*i + sirina*(brojTocaka.Count-1)*i +k*sirina, 
                                                300 - (int)((vrStup[d]) * pixelStep),
                                                sirina, 
                                                (int)((vrStup[d]) * pixelStep));
                    i++;
                    g.FillRectangle(myBrush, r);
                    g.DrawRectangle(pen, r);
                    //g.DrawLine(crta, 30 + (sirina / 2) + (i * sirina), 300 - (int)((vrStup[d] - vrijednosti.Min()) * pixelStep),
                        //30 + (sirina / 2) + ((i + 1) * sirina), 300 - (int)((vrStup[d + 1] - vrijednosti.Min()) * pixelStep));
                }
                
                //string naziv = "";
               /* if (k < imenaKategorija.Count)
                {
                    naziv = imenaKategorija[k];
                }
                else
                    naziv = "Stupac " + (k+1).ToString();*/
                Rectangle legend = new Rectangle(330, 100 + k * 15, 10, 10);
                g.DrawString(imenaKategorija[k + pomakniStupac], textFont, crni, 345, 98 + k * 15);
                g.FillRectangle(myBrush, legend);
                k++;
            }
        }
        public void drawLineChart()
        {
            tipGrafa = 2;
            fjaZaCrtanje = new nacrtajGraf(drawLineChart);

            List<Cell> oznakeXosi = new List<Cell>();
            bool stringOznake = false;
            int pomakniStupac = 0;
            if (prviStupacOznake)
            {
                pomakniStupac = 1;
                foreach (KeyValuePair<KeyValuePair<int, int>, Cell> c in ListaCelija.sveCelije)
                {
                    if (c.Value.stupac == prviStupac)
                    {
                        oznakeXosi.Add(c.Value);
                        if (c.Value.Numerical == false)
                        {
                            stringOznake = true;
                        }
                    }
                }
                if (!stringOznake)
                {
                    vrijednosti.Clear();
                    foreach (Cell c in CelijeZaPlot)
                    {
                        if (c.stupac != prviStupac)
                        {
                            vrijednosti.Add(Double.Parse(c.Sadrzaj));
                        }
                    }
                }
            }
            else
            {
                foreach (Cell c in CelijeZaPlot)
                {
                    
                        vrijednosti.Add(Double.Parse(c.Sadrzaj));
                    
                }
            }

            //oznakeXosi.Sort();

            // lista stupaca u kojima se nalaze vrijednosti koje crtamo
            List<int> stupci = new List<int>();
            
            // broj tocaka u svakom od tih stupaca, da znamo odrediti
            // koliko ce graf biti sirok
            List<int> brojTocaka = new List<int>();

            // popunjavanje gornjih listi
            stupci.Add(CelijeZaPlot[0].stupac);
            brojTocaka.Add(1);
            foreach (Cell c in CelijeZaPlot)
            {
                if (stupci.Contains(c.stupac))
                {
                    //brojTocaka[stupci.IndexOf(c.stupac)]++;
                    continue;
                }
                else
                {
                    stupci.Add(c.stupac);
                    brojTocaka.Add(1);
                }
            }
            stupci.Sort();
            if (prviStupacOznake)
            {
                // stupacOznaka = stupci[0];
                stupci.RemoveAt(0);
                brojTocaka.RemoveAt(0);
            }
            foreach (Cell c in CelijeZaPlot)
            {

                if (stupci.IndexOf(c.stupac) >= 0)
                    brojTocaka[stupci.IndexOf(c.stupac)]++;


            }

            foreach (int i in stupci)
            {
                bool found = false;
                foreach (Cell c in Kategorije)
                {
                    if (c.stupac == i)
                    {
                        imenaKategorija.Add(c.Sadrzaj);
                        found = true;
                    }
                }
                if (!found)
                {
                    imenaKategorija.Add("Stupac " + Convert.ToChar(i + 65 + pomakniStupac).ToString());
                }
            }
            //double absMin = Math.Abs()
            double span = vrijednosti.Max() - vrijednosti.Min(); // raspon vrijednosti
            if (span == 0) span = 1;
            double pixelStep = 270 / span;                          // koliko vrijednosi nosi jedan pixel na grafu
            int sirina = 270 / (brojTocaka.Max()-1);                 // razmak izmedu tocaka
            b = new Bitmap(430, 350);
            //Graphics g = graf.CreateGraphics();
            g = Graphics.FromImage(b);
            graf.Image = b;
            g.FillRectangle(new SolidBrush(Color.White), 0, 0, 430, 350);
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            Brush crni = new SolidBrush(Color.Black);
            Font naslovFont = new System.Drawing.Font("Helvetica", 10);
            Font textFont = new System.Drawing.Font("Helvetica", 8);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            
            Pen pen = new Pen(Color.Black, 1);
            Pen myPen = new Pen(Color.Gray, 1);
            g.DrawLine(pen, 50, 300, 310, 300);
            g.DrawLine(pen, 50, 300, 50, 20);
            double broj = vrijednosti.Min();
            //Brush crni = new SolidBrush(Color.Black);
            Font myFont = new System.Drawing.Font("Helvetica", 10);
            double brojZaPlot = vrijednosti.Max();
            double korak = (vrijednosti.Max()-vrijednosti.Min()) / 10;

            if (korak != 0)
            {



                //int brPixela = 270 / (int)((vrijednosti.Max() - vrijednosti.Min()) / 10);
                int j;
                string displayBroj = "";
                for (j = 30; j < 300; j += 27)
                {
                    displayBroj = (brojZaPlot).ToString().Substring(0, Math.Min((brojZaPlot).ToString().Length, 5));
                    displayBroj = displayBroj.TrimEnd('.');
                    g.DrawLine(myPen, 50, j, 310, j);
                    g.DrawString(displayBroj, textFont, crni, 15, j - 8);
                    brojZaPlot -= korak;

                }
                displayBroj = vrijednosti.Min().ToString().Substring(0, Math.Min((vrijednosti.Min()).ToString().Length, 5)); ;
                g.DrawLine(myPen, 50, j, 310, j);
                g.DrawString(displayBroj, textFont, crni, 15, j - 8);
            }
            if (!prviStupacOznake)
            {
                for (int i = 1; i < brojTocaka.Max(); i++)
                {
                    g.DrawString(i.ToString(), textFont, crni,
                        56 + sirina * (i - 1) + (sirina / 2), 304);
                }
            }
            else
            {
                int i = 1;
                foreach (Cell c in oznakeXosi)
                {
                    g.DrawString(c.Sadrzaj, textFont, crni,
                        56 + sirina * (i - 1) + (sirina / 2), 304);
                    i++;
                }
            }
            int k = 0;
            foreach (int stupac in stupci)
            {
                Brush myBrush = new SolidBrush(boje[k % boje.Count]);
                Pen crta = new Pen(boje[k%boje.Count], 3); 

                // vrijednosti u stupcu koji se trenutno crta
                List<double> vrStup = new List<double>();

                for (int In = CelijeZaPlot.Count - 1; In >= 0; In--)
                {
                    if (CelijeZaPlot[In].stupac == stupac)
                    {
                        vrStup.Add(Double.Parse(CelijeZaPlot[In].Sadrzaj));
                    }
                }
                int i = 0;
                for (int d = 0; d < vrStup.Count - 1; d++)
                {

                    g.DrawLine(crta, 60 + (sirina / 2) + (i * sirina), 300 - (int)((vrStup[d] - vrijednosti.Min()) * pixelStep),
                        60 + (sirina / 2) + ((i + 1) * sirina), 300 - (int)((vrStup[d + 1] - vrijednosti.Min()) * pixelStep));

                    i++;

                }
                //Rectangle myRectangle = new Rectangle(20, 20, 250, 200);
                i = 0;
                Pen tocka = new Pen(Color.Black, 4);
                g.DrawEllipse(tocka, 58 + (sirina / 2) + (i * sirina), 298 - (int)((vrStup[0] - vrijednosti.Min()) * pixelStep), 4, 4);
                for (int d = 0; d < vrStup.Count - 1; d++)
                {
                    g.DrawEllipse(tocka, 58 + (sirina / 2) + ((i + 1) * sirina), 298 - (int)((vrStup[d+1] - vrijednosti.Min()) * pixelStep), 4, 4);
                    i++;

                }
                Rectangle legend = new Rectangle(330, 100 + k * 15, 10, 10);
                
                g.DrawString(imenaKategorija[k + pomakniStupac], textFont, crni, 345, 98 + k * 15);
                g.FillRectangle(myBrush, legend);
                k++;
            }
            //stanjePrijeOznaka = g.Save();
        }

        public void drawPieChart()
        {
            tipGrafa = 3;
            fjaZaCrtanje = new nacrtajGraf(drawPieChart);
            List<Cell> oznake = new List<Cell>();
            List<Cell> TxtOznake = new List<Cell>();
            List<double> plotVrijednosti = new List<double>();
            //bool stringOznake = false;

            List<int> stupci = new List<int>();

            // broj tocaka u svakom od tih stupaca, da znamo odrediti
            // koliko ce graf biti sirok
            List<int> brojTocaka = new List<int>();

            // popunjavanje gornjih listi
            stupci.Add(CelijeZaPlot[0].stupac);
            brojTocaka.Add(1);
            foreach (Cell c in CelijeZaPlot)
            {
                if (stupci.Contains(c.stupac))
                {
                    //brojTocaka[stupci.IndexOf(c.stupac)]++;
                    continue;
                }
                else
                {
                    stupci.Add(c.stupac);
                    brojTocaka.Add(1);
                }
            }
            stupci.Sort();

            int prviTxtStupac = prviStupac;

            foreach (KeyValuePair<KeyValuePair<int, int>, Cell> c in ListaCelija.sveCelije)
            {
                if (c.Value.Numerical == false && c.Value.stupac < prviTxtStupac && c.Value.stupac < prviStupac)
                {
                    prviTxtStupac = c.Value.stupac;
                }
            }

            
            //oznake.Sort();

            int stOznake;
            int plotStupac;

            stOznake = prviTxtStupac;//stupci[0];
            plotStupac = stupci[0];
            //}
            
            foreach (KeyValuePair<KeyValuePair<int, int>, Cell> c in ListaCelija.sveCelije)
            {
                if (c.Value.stupac == plotStupac && c.Value.Numerical)
                {
                    plotVrijednosti.Add(Double.Parse(c.Value.Sadrzaj));
                    oznake.Add(c.Value);
                }
            }
            foreach (KeyValuePair<KeyValuePair<int, int>, Cell> c in ListaCelija.sveCelije)
            {
                if (c.Value.stupac == stOznake)
                {
                    TxtOznake.Add(c.Value);

                }
            }
            for (int j = 0; j < TxtOznake.Count; j++)
            {
                for (int jj = 0; jj < oznake.Count; jj++)
                {
                    if (oznake[jj].red == TxtOznake[j].red)
                    {
                        oznake[jj] = TxtOznake[j];
                    }
                }
                
            }
                b = new Bitmap(430, 350);
            //Graphics g = graf.CreateGraphics();
            g = Graphics.FromImage(b);
            graf.Image = b;
            g.FillRectangle(new SolidBrush(Color.White), 0, 0, 430, 350);
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            int n = (int)plotVrijednosti.Sum();
            Rectangle rect = new Rectangle(50, 50, 230, 230);
            int i = 0;
            float startAngle = 0.0F;
            float sweepAngle = 360.0F / n * (float)plotVrijednosti[0];
            Brush crni = new SolidBrush(Color.Black);
            Font naslovFont = new System.Drawing.Font("Helvetica", 10);
            Font textFont = new System.Drawing.Font("Helvetica", 8);
            Rectangle legend;
            int k = 1;
            for (k = 1; k < plotVrijednosti.Count; k++)
            {
                Brush myBrush = new SolidBrush(boje[i % boje.Count]);



                // Fill pie to screen.
                g.FillPie(myBrush, rect, startAngle, sweepAngle);

                
                startAngle += sweepAngle;
                sweepAngle = 360.0F / n * (float)plotVrijednosti[k];

                legend = new Rectangle(330, 100 + k * 15, 10, 10);
                g.DrawString(oznake[i].Sadrzaj, textFont, crni, 345, 98 + k * 15);
                g.FillRectangle(myBrush, legend);
                i++;
            }
            legend = new Rectangle(330, 100 + k * 15, 10, 10);
            Brush myBrush2 = new SolidBrush(boje[i % boje.Count]);
            g.DrawString(oznake[i].Sadrzaj, textFont, crni, 345, 98 + k * 15);
            g.FillRectangle(myBrush2, legend);


            // Fill pie to screen.
            g.FillPie(myBrush2, rect, startAngle, sweepAngle);
            //GraphicsPath path = new GraphicsPath();

        }

        private void _MouseDown(object sender, MouseEventArgs e)
        {
            PictureBox p = (PictureBox)sender;
            dragging = true;
            mouseDownPoint = new Point(e.X, e.Y);
            p.BringToFront();
        }

        private void _MouseMove(object sender, MouseEventArgs e)
        {
            if (dragging)
                graf.Location += new Size(e.X - mouseDownPoint.X, e.Y - mouseDownPoint.Y);
        }

        private void _MouseUp(object sender, MouseEventArgs e)
        {
            dragging = false;

        }
    }
}
