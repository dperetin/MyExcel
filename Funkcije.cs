using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MyExcel
{
    public class Funkcije
    {
        public Dictionary<string, izracunaj> SveFunkcije = new Dictionary<string, izracunaj>();
        public Dictionary<string, string> opisi = new Dictionary<string, string>();
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
            SveFunkcije.Add("log", new izracunaj(Logaritam));
            SveFunkcije.Add("mod", new izracunaj(Mod));
            SveFunkcije.Add("sin", new izracunaj(Sinus));
            SveFunkcije.Add("cos", new izracunaj(Kosinus));
            SveFunkcije.Add("tan", new izracunaj(Tangens));
            SveFunkcije.Add("abs", new izracunaj(Apsolutno));
            SveFunkcije.Add("sqrt", new izracunaj(Korjen));
            SveFunkcije.Add("ceil", new izracunaj(Ceil));
            SveFunkcije.Add("floor", new izracunaj(Floor));
            SveFunkcije.Add("count", new izracunaj(Count));

            opisi.Add("min", "Minimum\n\n" +
                             "    Korištenje:\n    MIN ( x1; x2; ...; xn ) \n    MIN ( x1: xn )\n\n" + 
                             "    Opis:\n    Vraća minimum vrijednosti x1 ... xn");
            opisi.Add("max", "Maksimum\n\n" +
                             "    Korištenje:\n    MAX ( x1; x2; ...; xn )  \n    MAX ( x1: xn )\n\n" +
                             "    Opis:\n    Vraća maksimum vrijednosti x1 ... xn");
            opisi.Add("sum", "Suma\n\n" + 
                             "    Korištenje:\n    SUM ( x1; x2; ...; xn )  \n    SUM ( x1: xn )\n\n" +
                             "    Opis:\n    Vraća sumu vrijednosti x1 ... xn");
            opisi.Add("average", "Aritmeticka sredina\n\n" + 
                             "    Korištenje:\n    AVERAGE ( x1; x2; ...; xn ) ili \n    AVERAGE ( x1: xn )\n\n" + 
                             "    Opis:\n    Vraća aritmetičku sredinu vrijednosti x1 ... xn");
            opisi.Add("log", "Logaritam\n\n" + 
                             "    Korištenje:\n    LOG ( x )  \n    LOG ( x; b )\n\n" +
                             "    Opis:\n    Vraća logaritam broja x po bazi b (Default b = 10)");
            opisi.Add("mod", "Modulo\n\n" + 
                             "    Korištenje:\n    MOD ( a; b ) \n\n"+
                             "    Opis:\n    Vraća ostatak dijeljenja a sa b");
            opisi.Add("sin", "Sinus\n\n" + 
                             "    Korištenje:\n    SIN ( x ) \n\n" + 
                             "    Opis:\n    Vraća sinus broja x");
            opisi.Add("cos", "Kosinus\n\n" + 
                             "    Korištenje:\n    COS ( x ) \n\n" +
                             "    Opis:\n    Vraća kosinus broja x");
            opisi.Add("tan", "Tangens\n\n" + 
                             "    Korištenje:\n    TAN ( x ) \n\n" +
                             "    Opis:\n    Vraća tangens broja x");
            opisi.Add("abs", "Apsolutno\n\n" + 
                             "    Korištenje:\n    ABS ( x ) \n\n" +
                             "    Opis:\n    Vraća apsolutnu vrijednost broja x");
            opisi.Add("sqrt", "Apsolutno\n\n" + 
                             "    Korištenje:\n    SQRT ( x ) \n\n" +
                             "    Opis:\n    Vraća drugi korjen broja x");
            opisi.Add("ceil", "Najmanje cijelo\n\n" + 
                              "    Korištenje:\n    CEIL ( x ) \n\n" +
                              "    Opis:\n    Vraća najmanji cijeli broj veći od broja x");
            opisi.Add("floor", "Najveće cijelo\n\n" + 
                               "    Korištenje:\n    FLOOR ( x ) \n\n" +
                               "    Opis:\n    Vraća najveći cijeli broj manji od broja x");
            opisi.Add("count", "Brojanje\n\n" + 
                               "    Korištenje:\n    COUNT ( v; x1; ...; xn ) \n\n" +
                               "    Opis:\n    Vraća koliko od vrijednosti x1 ... xn je jednako v");
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
            if (celije.Count == 0)
                return 0;
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
            if (celije.Count == 0)
                return 0;
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
        private static double Sinus(List<double> celije)
        {
            return Math.Sin(celije[0]);
        }
        private static double Kosinus(List<double> celije)
        {
            return Math.Cos(celije[0]);
        }
        private static double Tangens(List<double> celije)
        {
            return Math.Tan(celije[0]);
        }
        private static double Apsolutno(List<double> celije)
        {
            return Math.Abs(celije[0]);
        }
        private static double Korjen(List<double> celije)
        {
            return Math.Sqrt(celije[0]);
        }
        private static double Logaritam(List<double> celije)
        {
            if (celije.Count == 0)
                return 0;
            double baza = 10;
            if (celije.Count > 1)
            {

                baza = celije[1];
                return Math.Log(baza, celije[0]);
            }
            return Math.Log(celije[0], baza);
        }
        private static double Mod(List<double> celije)
        {
            if (celije.Count != 2)
                return 0;
            return celije[1] % celije[0];
        }
        private static double Ceil(List<double> celije)
        {
            return Math.Ceiling(celije[0]);
        }
        private static double Floor(List<double> celije)
        {
            return Math.Floor(celije[0]);
        }
        private static double Count(List<double> celije)
        {
            double count = -1;
            double uvjet = celije[celije.Count - 1];
            foreach (double d in celije)
            {
                if (d == uvjet)
                    count++;
            }
            return count;
        }


    }
    public delegate double izracunaj(List<double> o);
}

