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

        private Cell(int s, int r)
        {
            red = r;
            stupac = s;
        }
        public static Cell NapraviCeliju(int s, int r)
        {
            return new Cell(s, r);
        }
    }

    class Celije
    {
        Dictionary<KeyValuePair<int, int>, Cell> sveCelije = new Dictionary<KeyValuePair<int, int>,Cell>();

        public void Dodaj(int s, int r)
        {
            KeyValuePair<int, int> index = new KeyValuePair<int, int>(s, r);
            sveCelije.Add(index, Cell.NapraviCeliju(s, r));
        }
    }
}
