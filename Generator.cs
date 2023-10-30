using System;
using System.Collections.Generic;
using System.Linq;

namespace GeneratorPozdravleni
{
    internal class Generator
    {
        private static List<List<int>> GetCombs(List<int> list, int length) // составление уникальных комбинаций тем
        {
            if (length == 1) return list.Select(t => new List<int> { t }).ToList();
            return GetCombs(list, length - 1).SelectMany(t => list.Where(o => o.CompareTo(t.Last()) > 0), (t1, t2) => t1.Concat(new List<int> { t2 }).ToList()).ToList();
        }
        private static bool UniqueCongratulations(string[,] Phrases, List<string> People, int countBookmarks) // проверка на возможность составления уникальных поздравлений
        {
            bool flag = false;

            int[] countCongrInThemes = new int[Phrases.GetLength(1)]; // кол-во фраз по темам
            for (int j = 0; j < Phrases.GetLength(1); j++)
                for (int i = 0; i < Phrases.GetLength(0); i++)
                    if (Phrases[i, j] != "") countCongrInThemes[j]++;

            int countUniqueCongr = 0; // количество уникальных наборов фраз
            List<int> indexTheme = new List<int>(Phrases.GetLength(1));
            for (int i = 0; i < Phrases.GetLength(1); i++) indexTheme.Add(i);

            List<List<int>> UniqueProizved = GetCombs(indexTheme, countBookmarks); // составляем уникальные комбинации
            foreach (List<int> proizved in UniqueProizved)
            {
                int proiz = 1;
                foreach (int item in proizved) proiz *= countCongrInThemes[item]; // считаем кол-во вариантов одной комбинации
                countUniqueCongr += proiz; // суммируем все комбинации
            }

            if (countUniqueCongr >= People.Count) flag = true;
            return flag;
        }
        public static List<List<int>> CreateCongratulations(string[,] Phrases, List<string> People, int countBookmarks, ref List<List<int>> Themes) // составление всех поздравлений из фраз-пожеланий
        {
            if (!UniqueCongratulations(Phrases, People, countBookmarks))
            {
                Console.WriteLine("Создать уникальные поздравления для необходимого количества людей невозможно.\nДля решения проблемы добавьте возможные фразы-пожелания в таблицу.");
                System.Environment.Exit(1);
            }

            Console.WriteLine("Идет создание поздравлений для " + People.Count + " людей . . .");

            List<List<int>> Congratulations= new List<List<int>>();
            int[,] DataUsage = new int[Phrases.GetLength(0), Phrases.GetLength(1)];

            List<int> ThemeUsage = new List<int>(Phrases.GetLength(1));
            for (int i = 0; i < Phrases.GetLength(1); i++) ThemeUsage.Add(0);

            for (int i = 0; i < DataUsage.GetLength(0); i++)
                for (int j = 0; j < DataUsage.GetLength(1); j++)
                    if (Phrases[i, j] == "") DataUsage[i, j] = int.MaxValue;

            Random random = new Random();
            
            for (int i = 0; i < People.Count; i++)
            {
                List<int> congratulation = new List<int>();
                List<int> randomThemes = new List<int>();
                
                int count = 0;
                while (count < countBookmarks) // рандомно определяем темы поздравлений с учетом их использования
                {
                    int indexTheme;
                    do
                    {
                        indexTheme = random.Next(0, Phrases.GetLength(1));
                    } while (randomThemes.Contains(indexTheme) || ThemeUsage[indexTheme] != ThemeUsage.Min());
                    randomThemes.Add(indexTheme);
                    ThemeUsage[indexTheme]++; // увеличиваем счетчик использования темы
                    count++;
                }
                randomThemes.Sort();
                Themes.Add(randomThemes);

                count = 0;
                while (count < countBookmarks) // проходимся по всем закладкам
                {
                    List<int> stolb = new List<int>();

                    for (int k = 0; k < DataUsage.GetLength(0); k++)
                        stolb.Add(DataUsage[k, randomThemes[count]]);     // сохраяем кол-во использований фраз текущей темы
                    int indexrandom;
                    do
                    {
                        indexrandom = random.Next(0, DataUsage.GetLength(0)); //рандомно выбираем фразу из темы
                    } while (stolb[indexrandom] != stolb.Min()); // пока не выпадет фраза с наименьшим числом использований

                    congratulation.Add(indexrandom);
                    count++;

                }
                if (Congratulations.Contains(congratulation)) { i--; continue; } // если такая комбинация уже составлена, то ее использовать нельзя
                Congratulations.Add(congratulation);

                count = 0;
                foreach (int index in congratulation)
                    DataUsage[index, randomThemes[count++]]++; // увеличиваем счетчик использований фраз
            }

            Console.WriteLine("Поздравления составлены.\n");
            return Congratulations;
        }
    }
}
