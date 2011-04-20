using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MyExcel
{
    class Cell
    {
        private int red;
        private string stupac;
        private string sadrzaj;

        private Cell(string s, int r)
        {
            red = r;
            stupac = s;
        }
        public static Cell NapraviCeliju(string s, int r)
        {
            return new Cell(s, r);
        }
    }

    class Celije
    {
        Dictionary<KeyValuePair<string, int>, Cell> sveCelije = new Dictionary<KeyValuePair<string,int>,Cell>();

        public void Dodaj(string s, int r)
        {
            KeyValuePair<string, int> index = new KeyValuePair<string, int>(s, r);
            sveCelije.Add(index, Cell.NapraviCeliju(s, r));
        }
    }
}
