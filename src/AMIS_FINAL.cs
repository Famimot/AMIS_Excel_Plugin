using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AMIS_32bit
{
    public static class AmisCalculator
    {
        private static (double[] x, double[] y) RemoveDuplicatesAndSort(double[] x, double[] y)
        {
            var sortedIndices = Enumerable.Range(0, x.Length).OrderBy(i => x[i]).ToArray();
            var xSorted = sortedIndices.Select(i => x[i]).ToArray();
            var ySorted = sortedIndices.Select(i => y[i]).ToArray();

            var uniqueDict = new Dictionary<double, double>();
            for (int i = 0; i < xSorted.Length; i++)
            {
                uniqueDict[xSorted[i]] = ySorted[i];
            }

            var uniqueX = uniqueDict.Keys.OrderBy(val => val).ToArray();
            var uniqueY = uniqueX.Select(val => uniqueDict[val]).ToArray();

            return (uniqueX, uniqueY);
        }

        private static double? SafeNanMean(IEnumerable<double> values)
        {
            var validValues = values.Where(v => !double.IsNaN(v)).ToArray();
            if (validValues.Length == 0) return null;
            return validValues.Average();
        }

        public static Dictionary<string, string> GetAvailableModels(int n)
        {
            var availableModels = new Dictionary<string, string>();

            if (n >= 100)
            {
                availableModels.Add("17_points", "17 points (recommended)");
                availableModels.Add("9_points", "9 points");
                availableModels.Add("5_points", "5 points");
                availableModels.Add("3_points", "3 points");
                availableModels.Add("linear", "Linear");
            }
            else if (n >= 50)
            {
                availableModels.Add("9_points", "9 points (recommended)");
                availableModels.Add("5_points", "5 points");
                availableModels.Add("3_points", "3 points");
                availableModels.Add("linear", "Linear");
            }
            else if (n >= 20)
            {
                availableModels.Add("5_points", "5 points (recommended)");
                availableModels.Add("3_points", "3 points");
                availableModels.Add("linear", "Linear");
            }
            else if (n >= 10)
            {
                availableModels.Add("3_points", "3 points (recommended)");
                availableModels.Add("linear", "Linear");
            }
            else
            {
                availableModels.Add("linear", "Linear (recommended)");
            }

            return availableModels;
        }

        private static Func<double, double> CreateInterpolator(double[] x, double[] y)
        {
            var (xClean, yClean) = RemoveDuplicatesAndSort(x, y);

            if (xClean.Length < 2)
            {
                double minX = xClean.Length > 0 ? xClean[0] : 0;
                double maxX = xClean.Length > 0 ? xClean[xClean.Length - 1] : 100;
                return value => 100 * (value - minX) / (maxX - minX);
            }

            return value =>
            {
                if (value <= xClean[0]) return yClean[0];
                if (value >= xClean[xClean.Length - 1]) return yClean[yClean.Length - 1];

                for (int i = 0; i < xClean.Length - 1; i++)
                {
                    if (value >= xClean[i] && value <= xClean[i + 1])
                    {
                        double t = (value - xClean[i]) / (xClean[i + 1] - xClean[i]);
                        return yClean[i] + t * (yClean[i + 1] - yClean[i]);
                    }
                }

                return yClean[yClean.Length - 1];
            };
        }

        public static (Dictionary<string, double[]> converted,
                      Dictionary<string, double[]> pointsX,
                      Dictionary<string, double[]> pointsY,
                      Dictionary<string, string> availableModels)
            AmisConversion(double[] data, double? fixedMin = null, double? fixedMax = null)
        {
            var cleanData = data.Where(v => !double.IsNaN(v)).ToArray();
            int n = cleanData.Length;

            if (n < 10)
                throw new ArgumentException($"Need minimum 10 values, got {n}");

            double dataMin = fixedMin ?? cleanData.Min();
            double dataMax = fixedMax ?? cleanData.Max();

            if (dataMin >= dataMax)
                throw new ArgumentException($"Minimum ({dataMin}) must be less than maximum ({dataMax})");

            double? x5Nullable = SafeNanMean(cleanData);
            double x5 = x5Nullable ?? (dataMin + dataMax) / 2;

            var lowerData = cleanData.Where(v => v >= dataMin && v <= x5).ToArray();
            var upperData = cleanData.Where(v => v >= x5 && v <= dataMax).ToArray();

            double? x3Nullable = SafeNanMean(lowerData);
            double? x7Nullable = SafeNanMean(upperData);

            double x3 = x3Nullable ?? (dataMin + x5) / 2;
            double x7 = x7Nullable ?? (x5 + dataMax) / 2;

            double x2 = 0, x4 = 0, x6 = 0, x8 = 0;
            if (n >= 20)
            {
                var mask1 = cleanData.Where(v => v >= dataMin && v <= x3).ToArray();
                var mask2 = cleanData.Where(v => v >= x3 && v <= x5).ToArray();
                var mask3 = cleanData.Where(v => v >= x5 && v <= x7).ToArray();
                var mask4 = cleanData.Where(v => v >= x7 && v <= dataMax).ToArray();

                x2 = SafeNanMean(mask1) ?? (dataMin + x3) / 2;
                x4 = SafeNanMean(mask2) ?? (x3 + x5) / 2;
                x6 = SafeNanMean(mask3) ?? (x5 + x7) / 2;
                x8 = SafeNanMean(mask4) ?? (x7 + dataMax) / 2;
            }

            double x25 = 0, x35 = 0, x45 = 0, x55 = 0, x65 = 0, x75 = 0, x85 = 0, x95 = 0;
            if (n >= 50)
            {
                var mask25 = cleanData.Where(v => v >= dataMin && v <= x2).ToArray();
                var mask35 = cleanData.Where(v => v >= x2 && v <= x3).ToArray();
                var mask45 = cleanData.Where(v => v >= x3 && v <= x4).ToArray();
                var mask55 = cleanData.Where(v => v >= x4 && v <= x5).ToArray();
                var mask65 = cleanData.Where(v => v >= x5 && v <= x6).ToArray();
                var mask75 = cleanData.Where(v => v >= x6 && v <= x7).ToArray();
                var mask85 = cleanData.Where(v => v >= x7 && v <= x8).ToArray();
                var mask95 = cleanData.Where(v => v >= x8 && v <= dataMax).ToArray();

                x25 = SafeNanMean(mask25) ?? (dataMin + x2) / 2;
                x35 = SafeNanMean(mask35) ?? (x2 + x3) / 2;
                x45 = SafeNanMean(mask45) ?? (x3 + x4) / 2;
                x55 = SafeNanMean(mask55) ?? (x4 + x5) / 2;
                x65 = SafeNanMean(mask65) ?? (x5 + x6) / 2;
                x75 = SafeNanMean(mask75) ?? (x6 + x7) / 2;
                x85 = SafeNanMean(mask85) ?? (x7 + x8) / 2;
                x95 = SafeNanMean(mask95) ?? (x8 + dataMax) / 2;

                if (Math.Abs(x95 - dataMax) < 1e-10)
                    x95 = (x8 + dataMax) / 2;
            }

            var converted = new Dictionary<string, double[]>();
            var pointsX = new Dictionary<string, double[]>();
            var pointsY = new Dictionary<string, double[]>();

            // Линейная модель (всегда доступна)
            var pointsLine = new double[] { dataMin, dataMax };
            var yLine = new double[] { 0, 100 };
            var interpLine = CreateInterpolator(pointsLine, yLine);
            converted.Add("linear", cleanData.Select(interpLine).ToArray());
            pointsX.Add("linear", pointsLine);
            pointsY.Add("linear", yLine);

            // 3-точечная модель (n >= 10)
            if (n >= 10)
            {
                var points3 = new double[] { dataMin, x5, dataMax };
                var y3 = new double[] { 0, 50, 100 };
                var interp3 = CreateInterpolator(points3, y3);
                converted.Add("3_points", cleanData.Select(interp3).ToArray());
                pointsX.Add("3_points", points3);
                pointsY.Add("3_points", y3);
            }

            // 5-точечная модель (n >= 20)
            if (n >= 20)
            {
                var points5 = new double[] { dataMin, x3, x5, x7, dataMax };
                var y5 = new double[] { 0, 25, 50, 75, 100 };
                var interp5 = CreateInterpolator(points5, y5);
                converted.Add("5_points", cleanData.Select(interp5).ToArray());
                pointsX.Add("5_points", points5);
                pointsY.Add("5_points", y5);
            }

            // 9-точечная модель (n >= 50)
            if (n >= 50)
            {
                var points9 = new double[] { dataMin, x2, x3, x4, x5, x6, x7, x8, dataMax };
                var y9 = new double[] { 0, 12.5, 25, 37.5, 50, 62.5, 75, 87.5, 100 };
                var interp9 = CreateInterpolator(points9, y9);
                converted.Add("9_points", cleanData.Select(interp9).ToArray());
                pointsX.Add("9_points", points9);
                pointsY.Add("9_points", y9);
            }

            // 17-точечная модель (n >= 100)
            if (n >= 100)
            {
                var points17 = new double[] { dataMin, x25, x2, x35, x3, x45, x4, x55, x5, x65, x6, x75, x7, x85, x8, x95, dataMax };
                var y17 = new double[] { 0, 6.25, 12.5, 18.75, 25, 31.25, 37.5, 43.75, 50, 56.25, 62.5, 68.75, 75, 81.25, 87.5, 93.75, 100 };
                var interp17 = CreateInterpolator(points17, y17);
                converted.Add("17_points", cleanData.Select(interp17).ToArray());
                pointsX.Add("17_points", points17);
                pointsY.Add("17_points", y17);
            }

            var availableModels = GetAvailableModels(n);

            return (converted, pointsX, pointsY, availableModels);
        }

        public static double[] Normalize(double[] data, string model = "5_points",
                                        double? fixedMin = null, double? fixedMax = null)
        {
            var (converted, _, _, availableModels) = AmisConversion(data, fixedMin, fixedMax);

            if (!availableModels.ContainsKey(model))
            {
                var available = availableModels.Keys.ToList();
                string availableStr = string.Join(", ", available);
                string recommended = available.Count > 0 ? available[0] : "linear";

                throw new ArgumentException(
                    $"Модель '{model}' недоступна для {data.Length} значений. " +
                    $"Доступно: {availableStr}. " +
                    $"Рекомендуется: {recommended}");
            }

            return converted[model];
        }
    }

    public static class ExcelFunctions
    {
        private static bool TryConvertToDouble(object obj, out double result)
        {
            result = 0;
            if (obj == null) return false;

            try
            {
                if (obj is double d)
                {
                    result = d;
                    return !double.IsNaN(d);
                }
                if (obj is int i)
                {
                    result = i;
                    return true;
                }
                if (obj is string str)
                {
                    return double.TryParse(str, out result);
                }

                result = Convert.ToDouble(obj);
                return true;
            }
            catch
            {
                return false;
            }
        }

        [ExcelFunction(Name = "AMIS_TEST")]
        public static string AMIS_TEST()
        {
            return "AMIS работает " + DateTime.Now.ToString("HH:mm:ss");
        }

        [ExcelFunction(Name = "AMIS", Description = "Нормализация AMIS")]
        public static object[,] AMIS(
            object[,] range,
            [ExcelArgument(Description = "Модель: 1=linear, 3=3_points, 5=5_points, 9=9_points, 17=17_points")]
            double modelNumber = 5)
        {
            try
            {
                // 1. Собираем ВСЕ значения из ВСЕГО диапазона
                List<double> allValues = new List<double>();
                int rows = range.GetLength(0);
                int cols = range.GetLength(1);

                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < cols; col++)
                    {
                        if (TryConvertToDouble(range[row, col], out double val))
                        {
                            allValues.Add(val);
                        }
                    }
                }

                // 2. Проверяем минимальное количество
                if (allValues.Count < 10)
                {
                    object[,] error = new object[1, 1];
                    error[0, 0] = $"Нужно минимум 10 значений. Найдено: {allValues.Count}";
                    return error;
                }

                // 3. Преобразуем номер модели в название
                string model;
                switch (modelNumber)
                {
                    case 1: model = "linear"; break;
                    case 3: model = "3_points"; break;
                    case 5: model = "5_points"; break;
                    case 9: model = "9_points"; break;
                    case 17: model = "17_points"; break;
                    default: model = "5_points"; break;
                }

                // 4. Нормализуем ВЕСЬ набор данных
                double[] normalizedAll = AmisCalculator.Normalize(allValues.ToArray(), model);

                // 5. Возвращаем нормализованные значения для ВСЕХ ячеек исходного диапазона
                object[,] result = new object[rows, cols];

                int valueIndex = 0;
                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < cols; col++)
                    {
                        if (TryConvertToDouble(range[row, col], out double _))
                        {
                            // Для ячеек с числами - нормализованное значение
                            if (valueIndex < normalizedAll.Length)
                            {
                                result[row, col] = normalizedAll[valueIndex];
                                valueIndex++;
                            }
                            else
                            {
                                result[row, col] = ExcelError.ExcelErrorNA;
                            }
                        }
                        else
                        {
                            // Для нечисловых ячеек - ошибка
                            result[row, col] = ExcelError.ExcelErrorNA;
                        }
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                object[,] error = new object[1, 1];
                error[0, 0] = "Ошибка: " + ex.Message;
                return error;
            }
        }

        [ExcelFunction(Name = "AMIS_AVAILABLE_MODELS", Description = "Получить доступные модели для данных")]
        public static object[,] AMIS_AVAILABLE_MODELS(object[,] range)
        {
            try
            {
                List<double> values = new List<double>();
                int rows = range.GetLength(0);
                int cols = range.GetLength(1);

                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < cols; col++)
                    {
                        if (TryConvertToDouble(range[row, col], out double val))
                            values.Add(val);
                    }
                }

                var models = AmisCalculator.GetAvailableModels(values.Count);

                object[,] result = new object[models.Count, 2];
                int i = 0;
                foreach (var kvp in models)
                {
                    result[i, 0] = kvp.Key;
                    result[i, 1] = kvp.Value;
                    i++;
                }

                return result;
            }
            catch (Exception ex)
            {
                object[,] error = new object[1, 1];
                error[0, 0] = "Ошибка: " + ex.Message;
                return error;
            }
        }
    }
}