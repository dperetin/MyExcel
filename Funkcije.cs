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

            opisi.Add("min", "Minimum\n\n\tKorištenje:\nMIN(x1; x2; ...; xn) ili \n\tMIN(x1: xn)");
           
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

