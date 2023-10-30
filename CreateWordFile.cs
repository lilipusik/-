using System;
using System.Collections.Generic;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace GeneratorPozdravleni
{
    internal class CreateWordFile
    {
        public static void CreateFileWord(List<string> People, string[,] Phrases) // создание выходного файла Word
        {
            Console.WriteLine("Началось создание выходного файла . . .");

            DirectoryInfo dirInfo = new DirectoryInfo(@"C:\Users\Юлия\source\repos\GeneratorPozdravleni\GeneratorPozdravleni\SaveFiles");
            if (!dirInfo.Exists) dirInfo.Create();

            string FontText = "", PathTemplate = ""; int SizeText = 0;
            ReadExcelFile.ReadSettings(ref FontText, ref SizeText, ref PathTemplate);

            Word.Application AppWord = new Word.Application();
            object pathTemplate = @"" + PathTemplate;
            object missing = System.Reflection.Missing.Value;

            Word.Document newDoc = AppWord.Documents.Add(pathTemplate, ref missing, ref missing, ref missing);
            Word.Selection selection = AppWord.Selection;
            var tableToUse = selection.Tables[1];
            Word.Range range = tableToUse.Range;
            range.Copy();

            int numerBookMarks = 1, countBookMarks = AppWord.ActiveDocument.Bookmarks.Count;

            if (countBookMarks - 1 > Phrases.GetLength(1))
            {
                Console.WriteLine("Шаблон открытки не подходит для считанного количества тем.\nДобавьте кол-во тем во входной файл.");
                System.Environment.Exit(1);
            }

            List<List<int>> Themes = new List<List<int>>();
            List<List<int>> Congratulations = Generator.CreateCongratulations(Phrases, People, countBookMarks - 1, ref Themes);

            int numHumen = 1;
            foreach (string human in People)
            {
                if (numHumen > 1)
                {
                    selection.EndKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);
                    object breakType = Word.WdBreakType.wdPageBreak;
                    selection.InsertBreak(ref breakType);
                    Word.Table tableCopy = newDoc.Tables.Add(selection.Range, 1, 1, ref missing, ref missing);
                    tableCopy.Range.Paste();
                    tableToUse = tableCopy;
                }

                foreach (Word.Bookmark bm in AppWord.ActiveDocument.Bookmarks)
                {
                    if (bm.Name == "Имя") bm.Range.Text = human;
                    else if (bm.Name == $"Поздравление{numerBookMarks}")
                    {
                        string congratulation = "";
                        List<int> congr = Congratulations[numHumen - 1];
                        List<int> Theme = Themes[numHumen - 1];
                        congratulation += Phrases[congr[--numerBookMarks], Theme[numerBookMarks]] + "\n";
                        numerBookMarks += 2;
                        bm.Range.Font.Size = SizeText;
                        bm.Range.Font.Name = FontText;
                        bm.Range.Text = congratulation;
                    }
                }
                numHumen++;
                numerBookMarks = 1;
            }
            string[] FilesInDirectory = Directory.GetFiles(@"C:\Users\Юлия\source\repos\GeneratorPozdravleni\GeneratorPozdravleni\SaveFiles");
            object pathSave = $@"C:\Users\Юлия\source\repos\GeneratorPozdravleni\GeneratorPozdravleni\SaveFiles\NewFile №{FilesInDirectory.Length + 1}";
            newDoc.SaveAs(ref pathSave, Word.WdSaveFormat.wdFormatDocument, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            newDoc.Close(); object falseObj = false;
            AppWord.Quit(falseObj);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(AppWord);
            ReadExcelFile.CloseProcess("WINWORD");
            Console.WriteLine("Файл с поздравлениями готов!\nЕго можно найти по пути:\n" + pathSave + ".docx\n\n");
        }
    }
}
