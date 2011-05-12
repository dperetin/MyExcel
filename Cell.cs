using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace MyExcel
{
    public class Cell
    {
        private int red;
        private int stupac;
        public string sadrzaj;
        public string formula;

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

        //public  parsiraj(string s)
        public List<Cell> parsiraj(string s)
        {
            string slovo, broj;

            List<Cell> celije = new List<Cell>();
            string[] koordinate = s.Split(';');
            foreach (string k in koordinate)
            {

                
                slovo = Regex.Match(k, @"[a-z]+").Value;
                broj = Regex.Match(k, "[0-9]+").Value;
                //celije.Add(sveCelije[])
                int c = slovo[0] - 97;
                int r = Convert.ToInt32(broj) - 1;
                KeyValuePair<int, int> index = new KeyValuePair<int, int>(r, c);
                celije.Add(sveCelije[index]);
            }
            return celije;
        }

    }

    public class Funkcije
    {
        public Dictionary<string, izracunaj> SveFunkcije = new Dictionary<string, izracunaj>();
        public Funkcije()
        {
            izracunaj a = new izracunaj(Average);
            izracunaj b = new izracunaj(Minimum);
            izracunaj c = new izracunaj(Maximum);
            SveFunkcije.Add("average", a);
            SveFunkcije.Add("max", c);
            SveFunkcije.Add("min", b);
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

    }
    public delegate double izracunaj(List<Cell> o);
}
