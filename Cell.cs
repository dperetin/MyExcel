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
        public void evaluateFormula(Dictionary<KeyValuePair<int, int>, Cell> sveCelije, Funkcije fje)
        {
            string formula;
            formula = this.formula.ToLower();
            // zamjenjujem oznake celija konkretnim vrijednostima
            // PRETPOSTAVKA: nema razmaka nije dosega
            formula = formula.Replace(" ", "");

            // rasirivanje :
           // try
            //{
                while (true)
                {
                    Match m = Regex.Match(formula, @"[a-z]+[0-9]+:[a-z]+[0-9]+");
                    string slovo, broj;
                    if (m.Success)
                    {
                        string rep = "";
                        string sa = m.Value;
                        string[] oddo = sa.Split(':');
                        slovo = Regex.Match(oddo[0], @"[a-z]+").Value;
                        broj = Regex.Match(oddo[0], "[0-9]+").Value;
                        int c1 = slovo[0] - 97;
                        int r1 = Convert.ToInt32(broj);

                        slovo = Regex.Match(oddo[1], @"[a-z]+").Value;
                        broj = Regex.Match(oddo[1], "[0-9]+").Value;
                        int c2 = slovo[0] - 97;
                        int r2 = Convert.ToInt32(broj);

                        if (c1 == c2 && r1 != r2)
                        {
                            for (int i = r1; i <= r2; i++)
                            {
                                rep += Char.ConvertFromUtf32(c1 + 65) + i.ToString() + ";";

                            }
                        }
                        else if (r1 == r2 && c1 != c2)
                        {
                            for (int i = c1; i <= c2; i++)
                            {
                                rep += Char.ConvertFromUtf32(i + 65) + r1.ToString() + ";";
                            }
                        }
                        else
                        {
                            for (int i = r1; i <= r2; i++)
                                for (int j = c1; j <= c2; j++)
                                {
                                    rep += Char.ConvertFromUtf32(j + 65) + i.ToString() + ";";
                                }
                        }
                        rep.TrimEnd(';');
                        formula = formula.Replace(sa, rep);
                    }
                    else
                    {
                        break;
                    }
                }
                //MessageBox.Show(formula);
                formula = formula.ToLower();
                while (true)
                {
                    Match m = Regex.Match(formula, @"[a-z]+[0-9]+");
                    if (m.Success)
                    {
                        string cel = m.Value;
                        string slovo = Regex.Match(cel, @"[a-z]+").Value;
                        string broj = Regex.Match(cel, "[0-9]+").Value;
                        int c1 = slovo[0] - 97;
                        int r1 = Convert.ToInt32(broj) - 1;
                        KeyValuePair<int, int> koo = new KeyValuePair<int, int>(r1, c1);

                        if (sveCelije.ContainsKey(koo) && sveCelije[koo].Numerical)
                        {
                            formula = formula.Replace(cel, sveCelije[koo].sadrzaj);
                        }
                        else
                        {
                            int aa = 0;
                            formula = formula.Replace(cel, /*aa.ToString()*/"");
                        }


                    }
                    else
                    {
                        break;
                    }
                }

                //MessageBox.Show(formula);
                // razdvajam formulu na tokene

                List<Token> listaTokena = new List<Token>();
                string s = formula.TrimStart('=');
                string function_token = "";
                string zagrada_token = "";
                string separator_token = "";
                string celija_token = "";
                string op_token = "";
                Queue<Token> q = new Queue<Token>();
                Stack<Token> st = new Stack<Token>();
                while (s != "")
                {
                    Match m1 = Regex.Match(s, @"^\s*[a-zA-Z]+\s*");
                    Match m2 = Regex.Match(s, @"^[)(]");
                    Match m3 = Regex.Match(s, @"^;");
                    Match m4 = Regex.Match(s, @"^[0-9.]+");
                    Match m5 = Regex.Match(s, @"^[\+\-\*\/\^]");
                    if (m1.Success) // FUNKCIJA
                    {
                        function_token = m1.Value;
                        s = s.Substring(function_token.Length, s.Length - function_token.Length);
                        listaTokena.Add(new Token("funk", function_token));
                        //continue;
                    }

                    
                    else if (m2.Success) // ZAGRADA
                    {
                        zagrada_token = m2.Value;
                        s = s.Substring(zagrada_token.Length, s.Length - zagrada_token.Length);
                        listaTokena.Add(new Token("zagr", zagrada_token));
                        //continue;
                    }



                    else if (m3.Success) // SEPARATOR
                    {
                        separator_token = m3.Value;
                        s = s.Substring(separator_token.Length, s.Length - separator_token.Length);
                        listaTokena.Add(new Token("sepa", separator_token));
                        //continue;
                    }


                    else if (m4.Success)
                    {
                        celija_token = m4.Value;
                        s = s.Substring(celija_token.Length, s.Length - celija_token.Length);
                        listaTokena.Add(new Token("broj", celija_token));
                        //continue;
                    }

                    else if (m5.Success) // FUNKCIJA
                    {
                        op_token = m5.Value;
                        //st.Push(function_token);
                        s = s.Substring(op_token.Length, s.Length - op_token.Length);
                        Token op = new Token("oper", op_token);
                        op.brArg = 2;
                        if (op_token == "+")
                        {
                            op.prioritet = 0;
                            op.asoc = "L";
                        }
                        else if (op_token == "-")
                        {
                            op.prioritet = 0;
                            op.asoc = "L";
                        }
                        else if (op_token == "*")
                        {
                            op.prioritet = 1;
                            op.asoc = "L";
                        }
                        else if (op_token == "/")
                        {
                            op.prioritet = 1;
                            op.asoc = "L";
                        }
                        else if (op_token == "^")
                        {
                            op.prioritet = 2;
                            op.asoc = "D";
                        }
                        listaTokena.Add(op);
                        //continue;
                    }
                    else
                    {
                        throw new Exception();
                    }
                }

                // brojim argumente fja jer imaju varijabilan broj argumenata
                for (int i = 0; i < listaTokena.Count; i++)
                {
                    if (listaTokena[i].tip == "funk")
                    {
                        int smijemBrojat = -1;
                        for (int j = i + 1; j < listaTokena.Count; j++)
                        {
                            if (smijemBrojat == 0 && listaTokena[j].tip == "sepa")
                            {
                                listaTokena[i].brArg++;
                            }
                            if (listaTokena[j].value == "(")
                            {
                                smijemBrojat++;

                            }
                            if (listaTokena[j].value == ")")
                            {
                                smijemBrojat--;
                            }
                            if (smijemBrojat == -1)
                                break;
                        }
                    }
                }

                // Shunting-yard

                foreach (Token t in listaTokena)
                {
                    if (t.tip == "oper") // OPERATOR
                    {

                        while (st.Count != 0 && st.Peek().tip == "oper" &&
                              ((t.asoc == "L" && t.prioritet <= st.Peek().prioritet) ||
                              (t.asoc == "R" && t.prioritet < st.Peek().prioritet)))
                        {
                            q.Enqueue(st.Pop());
                        }
                        st.Push(t);
                        continue;
                    }
                    if (t.tip == "funk") // FUNKCIJA
                    {

                        st.Push(t);

                        continue;
                    }


                    if (t.tip == "zagr") // ZAGRADA
                    {

                        if (t.value == "(")
                        {
                            st.Push(t);
                        }
                        else if (t.value == ")")
                        {
                            while (st.Peek().value != "(")
                            {
                                q.Enqueue(st.Pop());
                            }
                            st.Pop();

                            if (st.Count != 0 && st.Peek().tip == "funk")
                            {
                                q.Enqueue(st.Pop());
                            }
                        }

                        continue;
                    }



                    if (t.tip == "sepa") // SEPARATOR
                    {

                        while (st.Peek().value != "(")
                        {
                            q.Enqueue(st.Pop());
                        }

                        continue;
                    }


                    if (t.tip == "broj")
                    {

                        q.Enqueue(t);

                        continue;
                    }
                }

                while (st.Count != 0)
                {
                    q.Enqueue(st.Pop());
                }

                // racunanje izraza

                Stack<double> tmp = new Stack<double>();

                while (q.Count != 0)
                {
                    List<double> arg = new List<double>();
                    if (q.Peek().tip == "broj")
                    {
                        tmp.Push(Double.Parse(q.Dequeue().value));
                        continue;
                    }
                    if (q.Peek().tip == "funk" || q.Peek().tip == "oper")
                    {
                        for (int i = 0; i < q.Peek().brArg; i++)
                        {
                            if (tmp.Count != 0)
                                arg.Add(tmp.Pop());
                        }
                        tmp.Push(fje.SveFunkcije[q.Dequeue().value](arg));
                    }
                }
                sadrzaj = tmp.Pop().ToString();

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
            SveFunkcije.Add("+", new izracunaj(Suma));
            SveFunkcije.Add("-", new izracunaj(Razlika));
            SveFunkcije.Add("*", new izracunaj(Produkt));
            SveFunkcije.Add("/", new izracunaj(Kvocjent));
            SveFunkcije.Add("^", new izracunaj(Potencija));
        }

        private static double Average(List<double> celije)
        {
            double suma = 0;

            foreach (double c in celije)
            {
                suma += c;
            }
            return suma / celije.Count;
        }
        private static double Maximum(List<double> celije)
        {
            double max = celije[0];

            foreach (double c in celije)
            {
                if (c > max)
                    max = c;
            }
            return max;
        }
        private static double Minimum(List<double> celije)
        {
            double min = celije[0];

            foreach (double c in celije)
            {
                if (c < min)
                    min = c;
            }
            return min;
        }
        private static double Suma(List<double> celije)
        {
            double suma = 0;

            foreach (double c in celije)
            {
                suma += c;
            }
            return suma;
        }
        private static double Razlika(List<double> celije)
        {
            return celije[1] - celije[0];
        }
        private static double Produkt(List<double> celije)
        {
            double prod = 1;

            foreach (double c in celije)
            {
                prod *= c;
            }
            return prod;
        }
        private static double Kvocjent(List<double> celije)
        {
            return celije[1] / celije[0];
        }
        private static double Potencija(List<double> celije)
        {
            return Math.Pow(celije[1], celije[0]);
        }

    }
    public delegate double izracunaj(List<double> o);
}
