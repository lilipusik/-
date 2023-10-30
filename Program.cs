using System;
using System.Collections.Generic;

namespace GeneratorPozdravleni
{
    internal class Program
    {
        static void Main(string[] args)
        {
            List<string> People = ReadExcelFile.ReadName();                            // список имен людей
            string[,] Phrases = ReadExcelFile.ReadCongratulations();                   // массив с фразами-пожеланиями
            CreateWordFile.CreateFileWord(People, Phrases);                            // создание выходного файла
            People.Clear();
        }
    }
}
