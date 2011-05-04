using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MyExcel
{
    class Cell
    {
        private int red;
        private int stupac;
        private string sadrzaj;

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

    }
}
