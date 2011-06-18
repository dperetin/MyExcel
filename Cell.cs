using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace MyExcel
{
    public class Cell : IComparable
    {
        public int red;
        public int stupac;
        public string sadrzaj;
        public string formula;

        private bool numerical = false;

        public bool Numerical
        {
            set { numerical = value; }
            get { return numerical; } 
        }

        int IComparable.CompareTo(object obj)
        {
            Cell o = (Cell)obj;

            if (this.red > o.red)
                return -1;
            if (this.red < o.red)
                return 1;

            return 0;
        }

        private Cell(int r, int s)
        {
            red = r;
            stupac = s;
            sadrzaj = "";
        }

        public static Cell NapraviCeliju(int r, int s)
        {
            return new Cell(r, s);
        }

        public void DodajVrijednostCeliji(string v)
        {
            sadrzaj = v;
        }
        public string DajVrijednostCelije()
        {
            return sadrzaj;
        }

        public void DodajVrijednostFormuli(string f)
        {
            formula = f;
        }
        public string DajVrijednostFormule()
        {
            return formula;
        }

    }

    class Celije
    {
        public Dictionary<KeyValuePair<int, int>, Cell> sveCelije = new Dictionary<KeyValuePair<int, int>, Cell>();

        public void Dodaj(int r, int s)
        {
            KeyValuePair<int, int> index = new KeyValuePair<int, int>(r, s);
            sveCelije.Add(index, Cell.NapraviCeliju(r, s));
        }

        public void DodajVrijednost(int r, int s, string v)
        {
            KeyValuePair<int, int> index = new KeyValuePair<int, int>(r, s);
            sveCelije[index].DodajVrijednostCeliji(v);
        }

        public void DodajFormulu(int r, int s, string f)
        {
            KeyValuePair<int, int> index = new KeyValuePair<int, int>(r, s);
            sveCelije[index].DodajVrijednostFormuli(f);
        }

        //public  parsiraj(string s)
        public List<Cell> parsiraj(string s)
        {
            string slovo, broj;
            double rez; 
            List<Cell> celije = new List<Cell>();
            string[] koordinate = s.Split(';');
            foreach (string k in koordinate)
            {

                if (k.Contains(":"))
                {
                    string[] oddo = s.Split(':');
                    slovo = Regex.Match(oddo[0], @"[a-z]+").Value;
                    broj = Regex.Match(oddo[0], "[0-9]+").Value;
                    int c1 = slovo[0] - 97;
                    int r1 = Convert.ToInt32(broj) - 1;

                    slovo = Regex.Match(oddo[1], @"[a-z]+").Value;
                    broj = Regex.Match(oddo[1], "[0-9]+").Value;
                    int c2 = slovo[0] - 97;
                    int r2 = Convert.ToInt32(broj) - 1;
                        
                    if (c1 == c2 && r1 != r2)
                    {
                        for (int i = r1; i <= r2; i++)
                        {
                            KeyValuePair<int, int> index = new KeyValuePair<int, int>(i, c1);
                            if (sveCelije.ContainsKey(index) &&
                                System.Double.TryParse(sveCelije[index].sadrzaj, out rez))
                                celije.Add(sveCelije[index]);
                        }
                    }
                    else if (r1 == r2 && c1 != c2)
                    {
                        for (int i = c1; i <= c2; i++)
                        {
                            KeyValuePair<int, int> index = new KeyValuePair<int, int>(r1, i);
                            if (sveCelije.ContainsKey(index) &&
                                System.Double.TryParse(sveCelije[index].sadrzaj, out rez)) 
                                    celije.Add(sveCelije[index]);
                        }
                    }
                    else
                    {
                        for (int i = r1; i <= r2; i++)
                            for (int j = c1; j <= c2; j++)
                            {
                                KeyValuePair<int, int> index = new KeyValuePair<int, int>(i, j);
                                //if (sveCelije.ContainsKey(index)) celije.Add(sveCelije[index]);
                                if (sveCelije.ContainsKey(index) &&
                                    System.Double.TryParse(sveCelije[index].sadrzaj, out rez))
                                        celije.Add(sveCelije[index]);
                            }
                    }
                }
                else
                {
                    slovo = Regex.Match(k, @"[a-z]+").Value;
                    broj = Regex.Match(k, "[0-9]+").Value;
                    //celije.Add(sveCelije[])
                    int c = slovo[0] - 97;
                    int r = Convert.ToInt32(broj) - 1;
                    KeyValuePair<int, int> index = new KeyValuePair<int, int>(r, c);
                    //celije.Add(sveCelije[index]);
                    if (sveCelije.ContainsKey(index) &&
                        System.Double.TryParse(sveCelije[index].sadrzaj, out rez))
                            celije.Add(sveCelije[index]);
                }
            }
            return celije;
        }

    }

    public class Funkcije
    {
        public Dictionary<string, izracunaj> SveFunkcije = new Dictionary<string, izracunaj>();
        public Funkcije()
        {
           
            SveFunkcije.Add("average", new izracunaj(Average));
            SveFunkcije.Add("min", new izracunaj(Minimum));
            SveFunkcije.Add("max", new izracunaj(Maximum));
            SveFunkcije.Add("sum", new izracunaj(Suma));
        }

        private static double Average(List<Cell> celije)
        {
            double suma = 0;

            foreach (Cell c in celije)
            {
                suma += Double.Parse(c.sadrzaj);
            }
            return suma / celije.Count;
        }
        private static double Maximum(List<Cell> celije)
        {
            double max = Double.Parse(celije.First().sadrzaj);

            foreach (Cell c in celije)
            {
                if (Double.Parse(c.sadrzaj) > max)
                    max = Double.Parse(c.sadrzaj);
            }
            return max;
        }
        private static double Minimum(List<Cell> celije)
        {
            double min = Double.Parse(celije.First().sadrzaj);

            foreach (Cell c in celije)
            {
                if (Double.Parse(c.sadrzaj) < min)
                    min = Double.Parse(c.sadrzaj);
            }
            return min;
        }
        private static double Suma(List<Cell> celije)
        {
            double suma = 0;

            foreach (Cell c in celije)
            {
                suma += Double.Parse(c.sadrzaj);
            }
            return suma;
        }

    }
    public delegate double izracunaj(List<Cell> o);
}
