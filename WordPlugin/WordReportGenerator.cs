// Файл: WordPlugin/WordReportGenerator.cs
using System;
using System.Runtime.InteropServices;
using PluginBase;
using Word = Microsoft.Office.Interop.Word;

namespace WordPlugin
{
    public class WordReportGenerator : IReportGenerator
    {
        public string Name => "MS Word Exporter (COM)";

        public void GenerateReport(string[] items, string outputPath)
        {
            var wordApp = new Word.Application();
            Word.Document doc = null;

            try
            {
                // Створюємо новий документ
                doc = wordApp.Documents.Add();

                // Додаємо заголовок
                Word.Paragraph heading = doc.Paragraphs.Add();
                heading.Range.Text = "Звіт у Microsoft Word";
                heading.Range.set_Style(Word.WdBuiltinStyle.wdStyleHeading1);
                heading.Range.InsertParagraphAfter();

                // Готуємо діапазон для таблиці — останній параграф документа
                Word.Range tableRange = doc.Paragraphs[doc.Paragraphs.Count].Range;

                // Додаємо таблицю: (кількість рядків = items.Length + 1, 1 стовпець)
                Word.Table table = doc.Tables.Add(tableRange, items.Length + 1, 1);
                table.Cell(1, 1).Range.Text = "Елемент";
                table.Rows[1].Range.set_Style(Word.WdBuiltinStyle.wdStyleHeading2);

                // Заповнюємо таблицю даними
                for (int i = 0; i < items.Length; i++)
                {
                    table.Cell(i + 2, 1).Range.Text = items[i];
                }

                // Зберігаємо документ
                doc.SaveAs2(outputPath);
            }
            finally
            {
                // Правильне закриття та звільнення COM-об'єктів
                if (doc != null)
                {
                    doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    Marshal.ReleaseComObject(doc);
                }

                wordApp.Quit();
                Marshal.ReleaseComObject(wordApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
