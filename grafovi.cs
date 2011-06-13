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
    class grafovi
    {
        Point mouseDownPoint;
        bool dragging;

        // vrijednosti koje se crtaju na grafu
        List<double> vrijednosti = new List<double>();

        // celije ciji sadrzaj se crta
        public Celije ListaCelija;

        // forma na kojoj se prikazuje panel
        public Form1 f;

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

        public grafovi(Form1 f, DataGridView grid, Celije ListaCelija)
        {
            this.ListaCelija = ListaCelija;
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
            
            graf.Size = new Size(330, 330);
            graf.BorderStyle = BorderStyle.FixedSingle;
            graf.Location = new Point(f.ClientSize.Width - 400, 50);
            graf.Parent = grid;
            

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
                    if (!Double.TryParse(c.Value.ToString(), out r))
                        continue;
                    vrijednosti.Add(r);
                }
            }
        }
        void histogram(object o, EventArgs e)
        {
            Panel graf = (Panel)o;
            int hStep = 270 / (int)vrijednosti.Max();
            int sirina = 270 / vrijednosti.Count;
            Graphics g = graf.CreateGraphics();

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
            foreach (double d in vrijednosti)
            {
                Brush myBrush = new SolidBrush(boje[i % boje.Count]);
                Rectangle r = new Rectangle(30 + (i * sirina), 300 - (int)d * hStep, sirina, (int)d * hStep);
                i++;
                g.FillRectangle(myBrush, r);
                g.DrawRectangle(pen, r);
            }
        }
        void line(object o, EventArgs e)
        {
            int hStep = 270 / (int)vrijednosti.Max();
            int sirina = 270 / vrijednosti.Count;
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
            for (int d = 0; d < vrijednosti.Count - 1; d++)
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

            // g.DrawPath(pen, path);*/
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
