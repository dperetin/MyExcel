using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MyExcel
{
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

