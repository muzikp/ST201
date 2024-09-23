using System.Collections.ObjectModel;
using ExcelDna.Integration;
using System;
using System.Linq;
using MathNet.Numerics.Distributions;
using MathNet.Numerics.Statistics;
using System.Diagnostics;
using ExcelDna.IntelliSense;

namespace ST201
{
    public class IntelliSenseAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Install();
        }
        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }
    public static class MyFunctions
    {
        private const string unequalLengthError = "Všechny sloupce musí mít stejný počet buněk";        

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

        [ExcelFunction(Name = "ROZKLAD.ROZPTYLU", Description = "Spočítá větu o rozkladu rotpylu.")]
        public static double VariancePopulationWeighted(
            [ExcelArgument(Name = "rozptyly", Description = "oblast buněk s hodnotami dílčích rozptylů")]
                    object[] vars,
            [ExcelArgument(Name = "průměry", Description = "oblast buněk s hodnotami dílčích průměrů")]
                    object[] means,
            [ExcelArgument(Name = "váhy", Description = "oblast buněk s váhami")]
                    object[] weights
        )
        {
            double n = 0;
            double s2n = 0;
            double xi_x2n = 0;
            double x = MeanWeighted(means, weights);
            for (int i = 0; i < vars.Length; i++) {
                double wi = Convert.ToDouble(weights[i]);
                double s2i = Convert.ToDouble(vars[i]);
                double xi = Convert.ToDouble(means[i]);                
                n += wi;
                s2n += s2i * wi;
                xi_x2n += Math.Pow(xi - x, 2) * wi;
            }
            return s2n / n + xi_x2n / n;
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

        [ExcelFunction(Name = "MAD", Description = "Spočítá absolutní mediánovou odchylku.")]
        public static double MedianAbsoluteDifference(
            [ExcelArgument(Name = "hodnoty", Description = "oblast buněk s hodnotami")]
            object[] values)
        {
            double[] vals = ObjArrToDoubleArr(values);
            double[] scores = new double[vals.Length];
            double median = Statistics.Median(vals);
            for (int i = 0; i < vals.Length; i++) {
                scores[i] += Math.Abs(vals[i] - median);
            }
            return Statistics.Median(scores);
        }


        #region Correlations

        [ExcelFunction(Name = "SPEARMAN", Description = "Spočítá Spearmanův korelační koeficient s ošetřením opakovaných pořadí (ties).",Category ="Statistical", HelpTopic = "https://www.seznam.cz")]
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

        [ExcelFunction(Name = "NORM.DIST.RANGE", Description = "Spočítá pravděpodobnost jevu mezi dvěma referenčními body u veličiny s normálním rozdělením.", Category = "SP201")]
        public static double NormalDistributionRange(
            [ExcelArgument(Name = "x", Description = "střední hodnota rozdělení")] double x,
            [ExcelArgument(Name = "s", Description = "směrodatná odchylka rozdělení")] double s,
            [ExcelArgument(Name = "x1", Description = "spodní hranice")] double x1,
            [ExcelArgument(Name = "x2", Description = "horní hranice")] double x2
            )
        {
            double p_upper = Normal.CDF(x, s, x2);
            double p_lower = Normal.CDF(x, s, x1);
            return p_upper - p_lower;
        }

        [ExcelFunction(Name = "KONTINGENCE.G", Description = "Spočítá testovou statistiku G pro kontingenční tabulku.", Category = "SP201")]
        public static double PivotTableGStatistic(double[,] observed)
        {            
            int rows = observed.GetLength(0);
            int cols = observed.GetLength(1);
            double[] rowSums = new double[rows];
            double[] colSums = new double[cols];         
            double totalObserved = 0;            
            double[] eMatrix = new double[rows * cols];
            // totals of rows
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    observed[i, j] = Convert.ToDouble(observed[i, j]);
                    rowSums[i] += observed[i, j];
                }
            }
            // totals of columns
            for (int j = 0; j < cols; j++)
            {
                for (int i = 0; i < cols; i++)
                {
                    colSums[i] += observed[j, i];
                }
            }            
            totalObserved = rowSums.Sum();
            // expected matrix
            int index = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    double e = rowSums[i] * colSums[j] / totalObserved;
                    eMatrix[index] = (Math.Pow(e - observed[i, j], 2) / e);
                    index++;
                }                
            }
            return eMatrix.Sum();            
        }


        [ExcelFunction(Name = "KONTINGENCE.C", Description = "Spočítá Pearsonův koeficient C z kontingenční tabulky.", Category = "SP201")]
        public static double PivotTableCStatistic(double[,] observed)
        {
            int rows = observed.GetLength(0);
            int cols = observed.GetLength(1);
            double[] rowSums = new double[rows];
            double[] colSums = new double[cols];
            double totalObserved = 0;
            double[] eMatrix = new double[rows * cols];
            // totals of rows
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    observed[i, j] = Convert.ToDouble(observed[i, j]);
                    rowSums[i] += observed[i, j];
                }
            }
            // totals of columns
            for (int j = 0; j < cols; j++)
            {
                for (int i = 0; i < cols; i++)
                {
                    colSums[i] += observed[j, i];
                }
            }
            totalObserved = rowSums.Sum();
            // expected matrix
            double g = PivotTableGStatistic(observed);
            return Math.Sqrt(g / (g + totalObserved));
        }

        [ExcelFunction(Name = "KONTINGENCE.V", Description = "Spočítá Cramérův koeficient V z kontingenční tabulky.", Category = "SP201")]
        public static double PivotTableVStatistic(double[,] observed)
        {
            int rows = observed.GetLength(0);
            int cols = observed.GetLength(1);
            double[] rowSums = new double[rows];
            double[] colSums = new double[cols];
            double totalObserved = 0;
            double[] eMatrix = new double[rows * cols];
            // totals of rows
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    observed[i, j] = Convert.ToDouble(observed[i, j]);
                    rowSums[i] += observed[i, j];
                }
            }            
            totalObserved = rowSums.Sum();
            // expected matrix
            double g = PivotTableGStatistic(observed);
            int[] minmax = new int[2];
            minmax[0]=rows;
            minmax[1] = cols;            
            return Math.Sqrt(g / (totalObserved * (minmax.Min()-1)));
        }

        [ExcelFunction(Name = "KONTINGENCE.PV", Description = "Spočítá p-hodnotu pro kontengenční tabulku.", Category = "SP201")]
        public static double PivotTablePValueStatistic(double[,] observed)
        {
            int rows = observed.GetLength(0);
            int cols = observed.GetLength(1);
            double[] rowSums = new double[rows];
            double[] colSums = new double[cols];
            double totalObserved = 0;
            double[] eMatrix = new double[rows * cols];
            // totals of rows
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    observed[i, j] = Convert.ToDouble(observed[i, j]);
                    rowSums[i] += observed[i, j];
                }
            }
            totalObserved = rowSums.Sum();
            double g = PivotTableGStatistic(observed);
            return 1-ChiSquared.CDF((rows - 1) * (cols - 1), g);            
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

        private static double Sum(object[] values)
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

        private static double Mean(object[] values)
        {
            double sum = 0;
            int n = 0;
            foreach (object v in values)
            {
                if (v is double || v is int)
                {
                    sum += Convert.ToDouble(v);
                    n++;
                }
                else if (v is string str && double.TryParse(str, out double num))
                {
                    sum += num;
                    n++;
                }
            }
            return sum/n;
        }
        private static double[] ObjArrToDoubleArr(object[] values)
        {
            double[] arr = new double[values.Length];
            double number;
            for (int i = 0; i < values.Length; i++) {
                var value = Convert.ToString(values[i]);
                if (Double.TryParse(value, out number)) {
                    arr[i] = Convert.ToDouble(values[i]);
                }
            }
            return arr;
        }


    }
}




