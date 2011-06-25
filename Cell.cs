using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace MyExcel
{
    public class Cell : IComparable
    {
        private string id;
        private string sadrzaj;
        private string formula = "";

        public int red;
        public int stupac;

             
        public List<Cell> uFormuli = new List<Cell>();
        private bool numerical = false;

        public bool Numerical
        {
            set { numerical = value; }
            get { return numerical; }
        }

        public string Formula
        {
            get { return formula; }
            set { formula = value; }
        }

        public string ID
        {
            get { return id; }
        }

        public void PostaviOvisnosti(Celije cells, string s)
        {
            s = s.ToLower();
            if (s != "")
            {
                string[] celije = s.Split(';');
                foreach (string koo in celije)
                {
                    string slovo, broj;
                    slovo = Regex.Match(koo, @"[a-z]+").Value;
                    broj = Regex.Match(koo, "[0-9]+").Value;
                    int c = slovo[0] - 97;
                    int r = Convert.ToInt32(broj) - 1;
                    KeyValuePair<int, int> i = new KeyValuePair<int, int>(r, c);
                    //KeyValuePair<int, int> ova = new KeyValuePair<int, int>(red, stupac);
                    if (cells.sveCelije.ContainsKey(i))
                    {
                        uFormuli.Add(cells.sveCelije[i]);
                    }
                    else
                    {
                        cells.Dodaj(r, c);
                        uFormuli.Add(cells.sveCelije[i]);
                    }
                }
            }
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
            id = Char.ConvertFromUtf32(s + 65) + (r + 1).ToString();
        }

        public static Cell NapraviCeliju(int r, int s)
        {
            return new Cell(r, s);
        }

        /* public void DodajVrijednostCeliji(string v)
         {
             sadrzaj = v;
         }*/
        public string Sadrzaj
        {
            get { return sadrzaj; }
            set
            {
                sadrzaj = value;
                double r;
                if (System.Double.TryParse(sadrzaj, out r))
                {
                    Numerical = true;
                }

            }
        }
        /*public string DajVrijednostCelije()
        {
            return sadrzaj;
        }*/

       /* public void DodajVrijednostFormuli(string f)
        {
            formula = f;
        }
        public string DajVrijednostFormule()
        {
            return formula;
        }*/
        public void evaluateFormula(Celije tCell, Funkcije fje)
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
                    rep = rep.TrimEnd(';');
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

                    if (tCell.sveCelije.ContainsKey(koo) && tCell.sveCelije[koo].Numerical)
                    {
                        string arg = tCell.sveCelije[koo].sadrzaj;
                        double d = double.Parse(arg);
                        if (d < 0)
                        {
                            arg = "0" + arg;
                        }
                        formula = formula.Replace(cel, arg);
                        if (tCell.sveCelije[koo].uFormuli.Contains(this) == false)
                            tCell.sveCelije[koo].uFormuli.Add(this);
                    }
                    else
                    {
                        //int aa = 0;
                        formula = formula.Replace(cel, /*aa.ToString()*/"");
                        if (tCell.sveCelije.ContainsKey(koo) == false)
                        {
                            tCell.Dodaj(r1, c1);
                            tCell.sveCelije[koo].uFormuli.Add(this);
                        }
                        throw new Exception();
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
            Sadrzaj = tmp.Pop().ToString();

        }
    }

    public class Celije
    {
        public Dictionary<KeyValuePair<int, int>, Cell> sveCelije = new Dictionary<KeyValuePair<int, int>, Cell>();

        public void Dodaj(int r, int s)
        {
            KeyValuePair<int, int> index = new KeyValuePair<int, int>(r, s);
            sveCelije.Add(index, Cell.NapraviCeliju(r, s));

        }

        /*public void DodajVrijednost(int r, int s, string v)
        {
            KeyValuePair<int, int> index = new KeyValuePair<int, int>(r, s);
            sveCelije[index].Sadrzaj = v;
        }*/

        /*public void DodajFormulu(int r, int s, string f)
        {
            KeyValuePair<int, int> index = new KeyValuePair<int, int>(r, s);
            sveCelije[index].Formula = f;
        }*/
    }
}
    
