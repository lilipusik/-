using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace GeneratorPozdravleni
{
    internal class ReadExcelFile
    {
        public static void CloseProcess(string NameProcess) //закрытие процессов Excel и Word
        {
            Process[] List = Process.GetProcessesByName(NameProcess);
            foreach (Process p in List) p.Kill();
        }

		public static List<string> ReadName() // чтение списка имен
        {
            Console.WriteLine("Началось чтение входных данных . . .");

            Excel.Application AppExcel = new Excel.Application();
            Excel.Workbook WorkbookExcel = AppExcel.Workbooks.Open(@"C:\Users\Юлия\source\repos\GeneratorPozdravleni\GeneratorPozdravleni\bin\Debug\Генератор поздравлений.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Excel.Worksheet WorksheetExcel = (Excel.Worksheet)WorkbookExcel.Sheets["Имена"];

            List<string> People = new List<string>();

            for (int i = 1; i <= WorksheetExcel.UsedRange.Rows.Count; i++)
                People.Add(WorksheetExcel.UsedRange.Cells[i].Value.ToString());

            AppExcel.Workbooks.Close();
            AppExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(AppExcel);
            CloseProcess("EXCEL");
            return People;
        }
        public static string[,] ReadCongratulations() // чтение фраз-пожеланий
        {
            Excel.Application AppExcel = new Excel.Application();
            Excel.Workbook WorkbookExcel = AppExcel.Workbooks.Open(@"C:\Users\Юлия\source\repos\GeneratorPozdravleni\GeneratorPozdravleni\bin\Debug\Генератор поздравлений.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Excel.Worksheet WorksheetExcel = (Excel.Worksheet)WorkbookExcel.Sheets["Пожелания"];

            string[,] Phrases = new string[WorksheetExcel.UsedRange.Rows.Count - 1, WorksheetExcel.UsedRange.Columns.Count];

            for (int i = 2; i <= WorksheetExcel.UsedRange.Rows.Count; i++)
                for (int j = 1; j <= WorksheetExcel.UsedRange.Columns.Count; j++)
                    if (WorksheetExcel.UsedRange.Cells[i, j] != null && WorksheetExcel.UsedRange.Cells[i, j].Value2 != null)
                        Phrases[i - 2, j - 1] = WorksheetExcel.UsedRange.Cells[i, j].Value2.ToString();

            for (int i = 0; i < Phrases.GetLength(0); i++)
                for (int j = 0; j < Phrases.GetLength(1); j++)
                    if (Phrases[i, j] == null) Phrases[i, j] = "";

            AppExcel.Workbooks.Close();
            AppExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(AppExcel);
            CloseProcess("EXCEL");

            Console.WriteLine("Чтение данных завершено.\n");

            return Phrases;
        }
        public static void ReadSettings(ref string FontText, ref int SizeText, ref string PathTemplate) // чтение настроечных данных
        {
            Excel.Application AppExcel = new Excel.Application();
            Excel.Workbook WorkbookExcel = AppExcel.Workbooks.Open(@"C:\Users\Юлия\source\repos\GeneratorPozdravleni\GeneratorPozdravleni\bin\Debug\Генератор поздравлений.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Excel.Worksheet WorksheetExcel = (Excel.Worksheet)WorkbookExcel.Sheets["Настройки"];

            for (int i = 1; i <= WorksheetExcel.UsedRange.Rows.Count; i++)
                switch (i)
                {
                    case 1: FontText = WorksheetExcel.UsedRange.Cells[i].Value.ToString(); break;
                    case 2: SizeText = int.Parse(WorksheetExcel.UsedRange.Cells[i].Value.ToString()); break;
                    case 3: PathTemplate = WorksheetExcel.UsedRange.Cells[i].Value.ToString(); break;
                    default: break;
                }

            AppExcel.Workbooks.Close();
            AppExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(AppExcel);
            CloseProcess("EXCEL");
        }
    }
}
