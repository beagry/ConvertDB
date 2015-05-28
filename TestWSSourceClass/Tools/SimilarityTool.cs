using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Converter.Tools
{
    /// <summary>
    ///     This class implements string comparison algorithm
    ///     based on character pair similarity
    ///     Source: http://www.catalysoft.com/articles/StrikeAMatch.html
    /// </summary>
    public static class SimilarityTool
    {
        /// <summary>
        ///     Compares the two strings based on letter pair matches
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <returns>The percentage match from 0.0 to 1.0 where 1.0 is 100%</returns>
        public static double CompareStrings(string str1, string str2)
        {
            var pairs1 = WordLetterPairs(str1.ToUpper());
            var pairs2 = WordLetterPairs(str2.ToUpper());

            var intersection = 0;
            var union = pairs1.Count + pairs2.Count;

            for (var i = 0; i < pairs1.Count; i++)
            {
                for (var j = 0; j < pairs2.Count; j++)
                {
                    if (pairs1[i] == pairs2[j])
                    {
                        intersection++;
                        pairs2.RemoveAt(j);
                            //Must remove the match to prevent "GGGG" from appearing to match "GG" with 100% success

                        break;
                    }
                }
            }

            return (2.0*intersection)/union;
        }

        /// <summary>
        ///     Gets all letter pairs for each
        ///     individual word in the string
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private static List<string> WordLetterPairs(string str)
        {
            var AllPairs = new List<string>();

            // Tokenize the string and put the tokens/words into an array
            var Words = Regex.Split(str, @"\s");

            // For each word
            for (var w = 0; w < Words.Length; w++)
            {
                if (!string.IsNullOrEmpty(Words[w]))
                {
                    // Find the pairs of characters
                    var PairsInWord = LetterPairs(Words[w]);

                    for (var p = 0; p < PairsInWord.Length; p++)
                    {
                        AllPairs.Add(PairsInWord[p]);
                    }
                }
            }

            return AllPairs;
        }

        /// <summary>
        ///     Generates an array containing every
        ///     two consecutive letters in the input string
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private static string[] LetterPairs(string str)
        {
            var numPairs = str.Length - 1;

            var pairs = new string[numPairs];

            for (var i = 0; i < numPairs; i++)
            {
                pairs[i] = str.Substring(i, 2);
            }

            return pairs;
        }

        public static decimal ComapareSimilarLists(List<string> list1, List<string> list2)
        {
            //Сравниваем два массива строк и выводим десятичное число = процент их схожести

            //Убираем дубли
            var uniqSubjsList = list1.Distinct().ToList();

            //Вес объекта
            var weigthOfObjects =
                list1.GroupBy(x => x).ToDictionary(x => x.Key, x => decimal.Round(x.Count()/(decimal) list1.Count, 4));

            //Console.WriteLine(weigthOfObjects.Values.Sum(x => x));
            //Словарь где записывается результат  - процент максимального совпадения с одним из пунктов LandCategoriesList
            //Мы получим ответ, на сколько каждая строка из уникальных в выгрузке соответствует максимально похожему критерию Категория земли
            var similaritiesDictionary = new Dictionary<string, decimal>();
            foreach (var s1 in uniqSubjsList)
            {
                var similaritiesCellWithLandListList = list2.Select(s2 => (decimal) CompareStrings(s1, s2)).ToList();
                similaritiesDictionary.Add(s1, similaritiesCellWithLandListList.Max());
            }

            //decimal maxSimilarity = similaritiesDictionary.Values.Max();
            var averageSimilarity = weigthOfObjects.Select(x => similaritiesDictionary[x.Key]*x.Value).Sum();
            return averageSimilarity;
        }
    }
}