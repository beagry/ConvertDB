using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Formater
{
    /// <summary>
    /// Класс для сравнения описания, рабоатет хорошо
    ///  Source: http://www.catalysoft.com/articles/StrikeAMatch.html
    /// </summary>
    public static class SimilarityTool
    {
        /// <summary>
        /// Compares the two strings based on letter pair matches
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <returns>The percentage match from 0.0 to 1.0 where 1.0 is 100%</returns>
        public static double CompareStrings(string str1, string str2)
        {
            List<string> pairs1 = WordLetterPairs(str1.ToUpper());
            List<string> pairs2 = WordLetterPairs(str2.ToUpper());

            int intersection = 0;
            int union = pairs1.Count + pairs2.Count;

            for (int i = 0; i < pairs1.Count; i++)
            {
                for (int j = 0; j < pairs2.Count; j++)
                {
                    if (pairs1[i] == pairs2[j])
                    {
                        intersection++;
                        pairs2.RemoveAt(j);//Must remove the match to prevent "GGGG" from appearing to match "GG" with 100% success

                        break;
                    }
                }
            }

            return (2.0 * intersection) / union;
        }

        /// <summary>
        /// Gets all letter pairs for each
        /// individual word in the string
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private static List<string> WordLetterPairs(string str)
        {
            List<string> AllPairs = new List<string>();

            // Tokenize the string and put the tokens/words into an array
            string[] Words = Regex.Split(str, @"\s");

            // For each word
            for (int w = 0; w < Words.Length; w++)
            {
                if (!string.IsNullOrEmpty(Words[w]))
                {
                    // Find the pairs of characters
                    String[] PairsInWord = LetterPairs(Words[w]);

                    for (int p = 0; p < PairsInWord.Length; p++)
                    {
                        AllPairs.Add(PairsInWord[p]);
                    }
                }
            }

            return AllPairs;
        }

        /// <summary>
        /// Generates an array containing every 
        /// two consecutive letters in the input string
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private static string[] LetterPairs(string str)
        {
            int numPairs = str.Length - 1;

            string[] pairs = new string[numPairs];

            for (int i = 0; i < numPairs; i++)
            {
                pairs[i] = str.Substring(i, 2);
            }

            return pairs;
        }

        public static decimal ComapareSimilarLists(List<string> list1, List<string> list2)
        {
            //Сравниваем два массива строк и выводим десятичное число = процент их схожести

            //Убираем дубли
            List<string> uniqSubjsList = list1.Distinct().ToList();

            //Вес объекта
            Dictionary<string, decimal> weigthOfObjects =
                list1.GroupBy(x => x).ToDictionary(x => x.Key, x => Decimal.Round((decimal)x.Count() / (decimal)list1.Count, 4));

            //Console.WriteLine(weigthOfObjects.Values.Sum(x => x));
            //Словарь где записывается результат  - процент максимального совпадения с одним из пунктов LandCategoriesList
            //Мы получим ответ, на сколько каждая строка из уникальных в выгрузке соответствует максимально похожему критерию Категория земли
            Dictionary<string, decimal> similaritiesDictionary = new Dictionary<string, decimal>();
            foreach (string s1 in uniqSubjsList)
            {
                List<decimal> similaritiesCellWithLandListList = list2.Select(s2 => (decimal)CompareStrings(s1, s2)).ToList();
                similaritiesDictionary.Add(s1, similaritiesCellWithLandListList.Max());
            }

            //decimal maxSimilarity = similaritiesDictionary.Values.Max();
            decimal averageSimilarity = weigthOfObjects.Select(x => similaritiesDictionary[x.Key] * x.Value).Sum();
            return averageSimilarity;
        }

        public static string GetMostSimilarityValue(string s, List<string> list, double minimalPercentSimilarity = 0)
        {
            if (list == null) return null;
            if (String.IsNullOrEmpty(s)) return null;

            var resultsDictionary = new Dictionary<string, double>();

            foreach (var s2 in list)
            {
                if (resultsDictionary.ContainsKey(s2)) continue;

                resultsDictionary.Add(s2, CompareStrings(s, s2));
            }
            var res = resultsDictionary.OrderBy(x => x.Value).First();

            if (minimalPercentSimilarity == 0)
                return res.Key;

            return res.Value >= minimalPercentSimilarity ? res.Key : null;

        }
    }
    
}
