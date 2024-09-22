using System.Collections.ObjectModel;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System;
using System.Linq;
//using Microsoft.Office.Interop.Excel;
using MathNet.Numerics.Distributions;

namespace ST201
{
    public static class MyFunctions
    {
        private const string unequalLengthError = "Všechny sloupce musí mít stejný počet buněk";        
        [ExcelFunction(Description = "My first .NET function")]
        public static string SayHello(string name)
        {
            return "Hello " + name;
        }

            [ExcelFunction(Name = "PRŮMĚR.W", Description = "Spočítá vážený aritmetický průměr.")]
            public static double MeanWeighted(
            [ExcelArgument(Name = "hodnoty", Description = "oblast buněk s hodnotami")]
            object[] values,
            [ExcelArgument(Name = "váhy", Description = "oblast buněk s váhami")]
            object[] weights
        )
        {
            
            if (values.Length != weights.Length)
            {
                throw new ArgumentException(unequalLengthError);
            }
            
            double weightedSum = 0;
            double totalWeight = 0;

            for (int i = 0; i < values.Length; i++)
            {
                double value = Convert.ToDouble(values[i]);
                double weight = Convert.ToDouble(weights[i]);

                weightedSum += value * weight;
                totalWeight += weight;
            }

            if (totalWeight == 0)
            {
                throw new DivideByZeroException("Total weight cannot be zero.");
            }

            return weightedSum / totalWeight;
        }

        [ExcelFunction(Name = "GEOMEAN.W", Description = "Spočítá vážený geometrický průměr.")]
        public static double GeomeanWeighted(
            [ExcelArgument(Name = "hodnoty", Description = "oblast buněk s hodnotami")]
                    object[] values,
            [ExcelArgument(Name = "váhy", Description = "oblast buněk s váhami")]
                    object[] weights
        )
        {
            if (values.Length != weights.Length)
            {
                throw new ArgumentException(unequalLengthError);
            }

            double weightedProduct = 1;
            double totalWeight = 0;

            for (int i = 0; i < values.Length; i++)
            {
                double value = Convert.ToDouble(values[i]);
                double weight = Convert.ToDouble(weights[i]);

                weightedProduct *= Math.Pow(value, weight);
                totalWeight += weight;
            }

            return Math.Pow(weightedProduct, 1 / totalWeight);
                       
        }

        [ExcelFunction(Name = "HARMEAN.W", Description = "Spočítá vážený harmonický průměr.")]
        public static double HarmeanWeighted(
            [ExcelArgument(Name = "hodnoty", Description = "oblast buněk s hodnotami")]
            object[] values,
            [ExcelArgument(Name = "váhy", Description = "oblast buněk s váhami")]
            object[] weights
        )
        {
            if (values.Length != weights.Length)
            {
                throw new ArgumentException(unequalLengthError);
            }

            double weightedSum = 0;
            double totalWeight = 0;

            for (int i = 0; i < values.Length; i++)
            {
                double value = Convert.ToDouble(values[i]);
                double weight = Convert.ToDouble(weights[i]);

                weightedSum += (weight/value);
                totalWeight += weight;
            }

            if (totalWeight == 0)
            {
                throw new DivideByZeroException("Total weight cannot be zero.");
            }

            return totalWeight / weightedSum;
        }

        [ExcelFunction(Name = "VAR.P.W", Description = "Spočítá vážený rozptyl v populaci.")]
        public static double VariancePopulationWeighted(
            [ExcelArgument(Name = "hodnoty", Description = "oblast buněk s hodnotami")]
                    object[] values,
            [ExcelArgument(Name = "váhy", Description = "oblast buněk s váhami")]
                    object[] weights
        )
        {
            if (values.Length != weights.Length)
            {
                throw new ArgumentException(unequalLengthError);
            }

            double n = 0;
            double xn = 0;
            double x2n = 0;                        

            for (int i = 0; i < values.Length; i++)
            {                
                double value = Convert.ToDouble(values[i]);
                double weight = Convert.ToDouble(weights[i]);

                n += weight;
                xn += value * weight;
                x2n += value * value * weight;
            }

            if (n == 0)
            {
                throw new DivideByZeroException("Total weight cannot be zero.");
            }

            return x2n/n - Math.Pow(xn/n,2);
        }

        [ExcelFunction(Name = "SMODCH.P.W", Description = "Spočítá váženou směrodatnou odchylku v populaci.")]
        public static double StdevPopulationWeighted(
            [ExcelArgument(Name = "hodnoty", Description = "oblast buněk s hodnotami")]
                    object[] values,
            [ExcelArgument(Name = "váhy", Description = "oblast buněk s váhami")]
                    object[] weights
        )
        {
            return Math.Sqrt(VariancePopulationWeighted(values, weights));
        }

        [ExcelFunction(Name = "VAR.KOEF.P", Description = "Spočítá variační koeficient populace.")]
        public static double VariationCoefficient(
            [ExcelArgument(Name = "hodnoty", Description = "oblast buněk s hodnotami")]
            object[] values)
        {
            if (values == null || values.Length == 0)
            {
                throw new ArgumentException("Hodnoty nemohou být prázdné.");
            }

            double sum = 0;
            double sumOfSquares = 0;
            int count = 0;

            foreach (var value in values)
            {
                double number = Convert.ToDouble(value);
                sum += number;
                sumOfSquares += number * number;
                count++;
            }

            if (count == 0)
            {
                throw new DivideByZeroException("Žádné platné hodnoty pro výpočet.");
            }

            double mean = sum / count;
            double variance = (sumOfSquares / count) - (mean * mean);
            double standardDeviation = Math.Sqrt(variance);

            if (mean == 0)
            {
                throw new DivideByZeroException("Průměr nemůže být nula pro výpočet variačního koeficientu.");
            }

            double variationCoefficient = (standardDeviation / mean);
            return variationCoefficient;
        }

        [ExcelFunction(Name = "VAR.RANGE", Description = "Spočítá variační rozpětí.")]
        public static double VariationRange(
            [ExcelArgument(Name = "hodnoty", Description = "oblast buněk s hodnotami")]
            object[] values)
        {
            if (values == null || values.Length == 0)
            {
                throw new ArgumentException("Hodnoty nemohou být prázdné.");
            }

            double min = double.MaxValue;
            double max = double.MinValue;

            foreach (var value in values)
            {
                double number = Convert.ToDouble(value);
                if (number < min) min = number;
                if (number > max) max = number;
            }

            if (min == double.MaxValue || max == double.MinValue)
            {
                throw new DivideByZeroException("Žádné platné hodnoty pro výpočet.");
            }

            return max - min; // Variační rozpětí
        }


        // variační koeficient - vzorek
        // variační koeficient - vzorek - vážený        
        // variační koeficient - populace - vážený        
        // Kontingece: G
        // Kontigence: p-value
        // Kontingence: Cramer
        // Kontingence: Pearson        
        // Spearman (p-value)

        #region Correlations

        [ExcelFunction(Name = "SPEARMAN", Description = "Spočítá Spearmanův korelační koeficient s ošetřením opakovaných pořadí (ties).")]
        public static double SpearmanCorrelationR(
            [ExcelArgument(Name = "hodnoty X", Description = "oblast buněk s hodnotami X")] object[] valuesX,
            [ExcelArgument(Name = "hodnoty Y", Description = "oblast buněk s hodnotami Y")] object[] valuesY)
        {
            if (valuesX.Length != valuesY.Length)
            {
                throw new ArgumentException("Obě pole musí mít stejný počet prvků.");
            }

            // Převod hodnot z Excelu na double
            double[] numericX = valuesX.Select(Convert.ToDouble).ToArray();
            double[] numericY = valuesY.Select(Convert.ToDouble).ToArray();
            double[] ranksX = GetRanks(numericX);
            double[] ranksY = GetRanks(numericY);
            int n = numericX.Length;
            double sumOfSquaredDiffs = 0;
            for (int i = 0; i < n; i++)
            {
                double d = ranksX[i] - ranksY[i];
                sumOfSquaredDiffs += d * d;
            }
            double correctionX = CorrectionForTies(numericX);
            double correctionY = CorrectionForTies(numericY);
            double denominator = Math.Sqrt((n * (n * n - 1) - correctionX) * (n * (n * n - 1) - correctionY));
            return 1 - (6 * sumOfSquaredDiffs) / denominator;
        }

        [ExcelFunction(Name = "SPEARMAN.T", Description = "Spočítá T statistiku pro Spearmanův korelační koeficient.")]
        public static double SpearmanCorrelationT(
            [ExcelArgument(Name = "hodnoty X", Description = "oblast buněk s hodnotami X")] object[] valuesX,
            [ExcelArgument(Name = "hodnoty Y", Description = "oblast buněk s hodnotami Y")] object[] valuesY
            ) {
            int n = valuesX.Length;
            double r = SpearmanCorrelationR(valuesX, valuesY);
            double T = r * Math.Sqrt(n - 2) / Math.Sqrt(1 - Math.Pow(r, 2));
            return T;
        }

        [ExcelFunction(Name = "SPEARMAN.PV", Description = "Spočítá p-hodnotu pro Spearmanův korelační koeficient.", Category = "SP201")]
        public static double SpearmanCorrelationPValue(
            [ExcelArgument(Name = "hodnoty X", Description = "oblast buněk s hodnotami X")] object[] valuesX,
            [ExcelArgument(Name = "hodnoty Y", Description = "oblast buněk s hodnotami Y")] object[] valuesY
            )
        {            
            int n = valuesX.Length;
            double t = SpearmanCorrelationT(valuesX, valuesY);
            double p = 2 * (1 - StudentT.CDF(0,1,n-1,Math.Abs(t)));
            return p;           
            
        }

        private static double[] GetRanks(double[] values)
        {
            int n = values.Length;
            double[] ranks = new double[n];
            int[] indices = Enumerable.Range(0, n).ToArray();

            // Seřadíme hodnoty a sledujeme původní indexy
            Array.Sort(values, indices);

            int i = 0;
            while (i < n)
            {
                int start = i;
                double sumRanks = 0;

                // Zjistíme, kolik hodnot je stejných (ties)
                while (i < n && values[start] == values[i])
                {
                    sumRanks += i + 1; // Pořadí je index + 1
                    i++;
                }

                // Průměrné pořadí pro stejné hodnoty
                double averageRank = sumRanks / (i - start);
                for (int j = start; j < i; j++)
                {
                    ranks[indices[j]] = averageRank;
                }
            }

            return ranks;
        }

        // Korekce pro opakovaná pořadí (ties) podle odkazu
        private static double CorrectionForTies(double[] values)
        {
            var groups = values.GroupBy(v => v).Where(g => g.Count() > 1);
            double correction = 0;

            foreach (var group in groups)
            {
                int count = group.Count();
                correction += count * (count * count - 1);
            }

            return correction;
        }

        #endregion

        private static double Sum(object[,] values)
        {
            double sum = 0;
            foreach (object v in values)
            {
                if (v is double || v is int)
                {
                    sum += Convert.ToDouble(v);
                }
                else if (v is string str && double.TryParse(str, out double num))
                {
                    sum += num;
                }
            }
            return sum;
        }



    }
}




