using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MyExcel
{
    class Token
    {
        public string value;
        public string tip;
        public int brArg = 1;
        public int prioritet;
        public string asoc;

        public Token(string t, string v)
        {
            tip = t;
            value = v;
        }
        public override string ToString()
        {
            return tip + "\t" + value + " " + brArg.ToString();

        }
    }
}
