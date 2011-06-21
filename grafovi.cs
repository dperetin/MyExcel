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
   
    class grafovi
    {
        Point mouseDownPoint;
        bool dragging;
        public string naslov = "";
        public string xOs = "";
        public string yOs = "";
        bool prviStupacOznake = false;
        bool prviRedakOznake = false;
        int prviStupac = 1000;
        int prviRedak  = 1000;
        // vrijednosti koje se crtaju na grafu
        List<double> vrijednosti = new List<double>();
        ContextMenuStrip strip = new ContextMenuStrip();
        // celije ciji sadrzaj se crta
        public Celije ListaCelija;
        private List<Cell> CelijeZaPlot = new List<Cell>();
        List<Cell> Kategorije = new List<Cell>();
        List<string> imenaKategorija = new List<string>();
        // forma na kojoj se prikazuje panel
        public Form1 f;
        public Form4 svojstva = new Form4();
        //panel na kojem se crta graf
        Panel graf = new Panel();
        
        // gumb za gasenje panela
        Button close = new Button();

        // boje koje koristimo u crtanju grafova
        List<Color> boje = new List<Color>();

        void zagasi(object o, EventArgs e)
        {
            Button b = (Button)o;
            b.Parent.Dispose();
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
            graf.Refresh();
            svojstva.Close();
        }
        public grafovi(Form1 f, DataGridView grid, Celije celije)
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
            graf.Location = new Point(f.ClientSize.Width - 500, 50);
            graf.Parent = grid;
            graf.MouseClick += new MouseEventHandler(desniKlikMeni);
            
            strip.Items.Add("stavka 1");
            strip.Items.Add("stavka 2");
            strip.Items.Add("stavka 3");
            strip.Items[0].Click += new EventHandler(izbor);
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
                double r;

                KeyValuePair<int, int> index = new KeyValuePair<int, int>(c.RowIndex, c.ColumnIndex);
                if (ListaCelija.sveCelije.ContainsKey(index))
                {
                    if (ListaCelija.sveCelije[index].Numerical == false)
                    {
                        Kategorije.Add(ListaCelija.sveCelije[index]);
                        if (ListaCelija.sveCelije[index].stupac < prviStupac)
                            prviStupac = ListaCelija.sveCelije[index].stupac;
                        continue;
                    }
                    vrijednosti.Add(Double.Parse(ListaCelija.sveCelije[index].sadrzaj));
                    CelijeZaPlot.Add(ListaCelija.sveCelije[index]);
                }
            }
            CelijeZaPlot.Sort();
        }
        void histogram(object o, EventArgs e)
        {
            //int stupacOznaka = ;
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
                        imenaKategorija.Add(c.sadrzaj);
                        found = true;
                    }
                }
                if (!found)
                {
                    imenaKategorija.Add("Stupac " + Convert.ToChar(i+65).ToString());
                }
            }

            //double absMin = Math.Abs()
            double span = vrijednosti.Max();// - vrijednosti.Min(); // raspon vrijednosti
            if (span == 0) span = 1;
            double pixelStep = 270 / span;                          // koliko vrijednosi nosi jedan pixel na grafu
            int sirina = (270 - brojTocaka.Max() * 10) / ((brojTocaka.Max() - 1) * brojTocaka.Count);                 // razmak izmedu tocaka

            
            Graphics g = graf.CreateGraphics();
            Brush crni = new SolidBrush(Color.Black);
            Font naslovFont = new System.Drawing.Font("Helvetica", 10);
            Font textFont = new System.Drawing.Font("Helvetica", 8);
            g.DrawString(naslov, naslovFont, crni, 215 - naslov.Length * 5 / 2, 10);
            g.DrawString(xOs, textFont, crni, 60 + 135 - xOs.Length * 5 / 2, 310);
            g.DrawString(yOs, textFont, crni, 1, 215 - yOs.Length * 5 / 2, new System.Drawing.StringFormat(StringFormatFlags.DirectionVertical));
            //GraphicsPath path = new GraphicsPath();
            Pen pen = new Pen(Color.Black, 1);
            Pen myPen = new Pen(Color.Gray, 1);
            g.DrawLine(pen, 50, 300, 310, 300);
            g.DrawLine(pen, 50, 300, 50, 20);
            int broj = 0;
            
           /* for (int j = 300; j >= 30; j -= hStep)
            {
                g.DrawLine(myPen, 20, j, 310, j);
                g.DrawString(broj.ToString(), myFont, crni, 10 - 7 * ((int)Math.Floor(Math.Log10(broj))), j - 8);
                broj++;
            }*/

            if (!prviStupacOznake)
            {
                for (int i = 1; i < brojTocaka.Max(); i++)
                {
                    g.DrawString(i.ToString(), textFont, crni,
                        45 + sirina * stupci.Count / 2 + (i - 1) * sirina * stupci.Count + 10 * i, 300);
                }
            }
            /*else
            {
                List<Cell> oznake = new List<Cell>();
                int i = 1;
                foreach (Cell c in Kategorije)
                {
                    
                    if (c.stupac == prviStupac)
                    {
                        oznake.Add(c);
                    }
                   

                }
                oznake.Sort();
                foreach (Cell c in oznake)
                {
                    g.DrawString(c.sadrzaj, textFont, crni,
                           45 + sirina * stupci.Count / 2 + (i - 1) * sirina * stupci.Count + 10 * i, 300);
                    i++;
                }
                
            }*/
            
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
                        vrStup.Add(Double.Parse(CelijeZaPlot[In].sadrzaj));
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
                Rectangle legend = new Rectangle(330, 100 + k * 15, 10, 10);
                string naziv = "";
               /* if (k < imenaKategorija.Count)
                {
                    naziv = imenaKategorija[k];
                }
                else
                    naziv = "Stupac " + (k+1).ToString();*/

                g.DrawString(imenaKategorija[k], textFont, crni, 345, 98 + k * 15);
                g.FillRectangle(myBrush, legend);
                k++;
            }
        }
        void line(object o, EventArgs e)
        {
            List<Cell> oznakeXosi = new List<Cell>();
            bool stringOznake = false;
            if (prviStupacOznake)
            {
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
            }

            oznakeXosi.Sort();

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
                        imenaKategorija.Add(c.sadrzaj);
                        found = true;
                    }
                }
                if (!found)
                {
                    imenaKategorija.Add("Stupac " + Convert.ToChar(i + 65).ToString());
                }
            }
            //double absMin = Math.Abs()
            double span = vrijednosti.Max() - vrijednosti.Min(); // raspon vrijednosti
            if (span == 0) span = 1;
            double pixelStep = 270 / span;                          // koliko vrijednosi nosi jedan pixel na grafu
            int sirina = 270 / (brojTocaka.Max()-1);                 // razmak izmedu tocaka

            Graphics g = graf.CreateGraphics();
            Brush crni = new SolidBrush(Color.Black);
            Font naslovFont = new System.Drawing.Font("Helvetica", 10);
            Font textFont = new System.Drawing.Font("Helvetica", 8);
            g.DrawString(naslov, naslovFont, crni, 215 - naslov.Length * 5 / 2, 10);
            g.DrawString(xOs, textFont, crni, 60 + 135 - xOs.Length * 5 / 2, 314);
            g.DrawString(yOs, textFont, crni, 1, 215 - yOs.Length * 5 / 2, new System.Drawing.StringFormat(StringFormatFlags.DirectionVertical));
            g.SmoothingMode = SmoothingMode.AntiAlias;
            
            Pen pen = new Pen(Color.Black, 1);
            Pen myPen = new Pen(Color.Gray, 1);
            g.DrawLine(pen, 50, 300, 310, 300);
            g.DrawLine(pen, 50, 300, 50, 20);
            double broj = vrijednosti.Min();
            //Brush crni = new SolidBrush(Color.Black);
            Font myFont = new System.Drawing.Font("Helvetica", 10);
           /* for (int j = 300; j >= 30; j -= 60)
            {
                g.DrawLine(myPen, 20, j, 310, j);
                g.DrawString(broj.ToString(), myFont, crni, 10 - 7 * broj.ToString().Length, j - 8);
                broj += (span / 5);
            } */

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
                    g.DrawString(c.sadrzaj, textFont, crni,
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
                        vrStup.Add(Double.Parse(CelijeZaPlot[In].sadrzaj));
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
                string naziv = "";
                /* if (k < imenaKategorija.Count)
                 {
                     naziv = imenaKategorija[k];
                 }
                 else
                     naziv = "Stupac " + (k+1).ToString();*/

                g.DrawString(imenaKategorija[k], textFont, crni, 345, 98 + k * 15);
                g.FillRectangle(myBrush, legend);
                k++;
            }
            
        }

        void pieChart(object o, EventArgs e)
        {
            Panel graf = (Panel)o;
     
            Graphics g = graf.CreateGraphics();
            g.SmoothingMode = SmoothingMode.AntiAlias;
            int n = (int)vrijednosti.Sum();
            Rectangle rect = new Rectangle(50, 50, 230, 230);
            int i = 0;
            float startAngle = 0.0F;
            float sweepAngle = 360.0F / n * (float)vrijednosti[0];
            for (int k = 1; k < vrijednosti.Count; k++)
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

        private void _MouseDown(object sender, MouseEventArgs e)
        {
            Panel p = (Panel)sender;
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

        public void drawPieChart()
        {
            graf.Paint += new PaintEventHandler(pieChart);
        }
        public void drawHistogram()
        {
            graf.Paint += new PaintEventHandler(histogram);          
        }
        public void drawLineChart()
        {
            graf.Paint += new PaintEventHandler(line);
        }
    }
}
