using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace EasyOffice
{

    static public class ContentMethods
    {
        static public Color ForeColor = Color.FromArgb(140, 140, 140);
        static public Color BackColor = Color.FromArgb(228, 228, 228);
        static public Color TextColor = Color.FromArgb(114, 114, 114);
        static public void ConcatPDF(string[] inFiles, string outDirectory)
        {
            using PdfDocument outputDocument = new PdfDocument();
            var lastFileName = inFiles.Last().Split('\\').Last();
            foreach (string file in inFiles)
            {
                PdfDocument inputDocument = PdfReader.Open(file, PdfDocumentOpenMode.Import);
                int count = inputDocument.PageCount;
                for (int idx = 0; idx < count; idx++)
                {
                    PdfPage page = inputDocument.Pages[idx];
                    outputDocument.AddPage(page);
                }
            }
            outputDocument.Save(outDirectory + $@"\{lastFileName}");
            outputDocument.Close();
        }

        public static void ConvertDocx(Word.Document doc, Word.Application app, string path, string exportName, int from, int to)
        {
            doc.ExportAsFixedFormat(exportName, Word.WdExportFormat.wdExportFormatPDF, Range:Word.WdExportRange.wdExportFromTo,From:from,To:to);
        }

        public static Dictionary<int, string> ParseIdsFromFilePathsAsDictionary(List<string> inFiles)
        {
            return inFiles.ToDictionary(path =>
            {
                var splitPath = path.Split(new[]
                {
                    ' ', '_'
                }, StringSplitOptions.RemoveEmptyEntries);
                return int.Parse(splitPath.Last());
            });
        }
        public static ILookup<int, string> ParseIdsFromFilePathsAsLookup(List<string> inFiles)
        {
            return inFiles.ToLookup(path =>
            {
                var splitPath = path.Split(new[]
                {
                    ' ', '_'
                }, StringSplitOptions.RemoveEmptyEntries);
                return int.Parse(splitPath.Last());
            });
        }
        static public List<string> GetFilesPath(string dir, string format = ".pdf")
        {
            if (Directory.Exists(dir))
            {
                List<string> maskFormat = Directory.GetFiles(dir).OrderBy(file => file).ToList();
                for (int k = maskFormat.Count - 1; k >= 0; k--)
                {
                    if (!maskFormat[k].EndsWith(format))
                    {
                        maskFormat.RemoveAt(k);
                    }
                    else
                    {
                        maskFormat[k] = maskFormat[k].Remove(maskFormat[k].Length - 4, 4);
                    }
                }
                return maskFormat;
            }
            return new List<string>();
        }

        public static void StartNumbering(int start, int end, bool title, bool instruction, TextBox process, ProgressBar bar, Form1.Doc docTitle, Form1.Doc docContent)
        {
            Word.Application app = null;
            try
            {
                Action action = () =>
                {
                    process.Text += "Запуск Word..." + Environment.NewLine;
                    process.Focus();
                    process.Select(process.Text.Length, 0);
                };
                process.Invoke(action);
                app = new Word.Application();
                if (title)
                    app.GenerateDocs(start, end, "Обложка", process, bar, docTitle);

                if (instruction)
                    app.GenerateDocs(start, end, "Инструкция", process, bar, docContent);

                bar.Invoke(new Action(() => { MessageBox.Show("Успешно завершено!"); }));

            }
            catch (Exception ex)
            {
                Action action = () =>
                {
                    process.Text += ex.Message + Environment.NewLine + ex.StackTrace + Environment.NewLine;
                    process.Focus();
                    process.Select(process.Text.Length, 0);
                };
                process.Invoke(action);
            }
            finally
            {
                app?.Quit();
            }
        }

        private static void GenerateDocs(this Word.Application app, int start, int end, string fileName, TextBox process, ProgressBar bar,
            Form1.Doc docContent)
        {
            {
                Action action = () =>
                {
                    process.Text += $"Открытие документа '{fileName}'" + Environment.NewLine;
                    process.Focus();
                    process.Select(process.Text.Length, 0);
                };
                process.Invoke(action);
            }
            var docFilePath = Directory.EnumerateFiles($@"{Environment.CurrentDirectory}\Word\", $"*{fileName}*.doc?")
                .First();
            for (int i = start; i <= end; i++)
            {
                var doc = app.Documents.Open(docFilePath);
                var bookmarks = doc.Bookmarks;
                foreach (var bookmark in bookmarks.Cast<Word.Bookmark>().ToList())
                    bookmark.Range.Text = i.ToString();

                var message = $"{fileName} {i}";
                Action action = () =>
                {
                    process.Text += message + Environment.NewLine;
                    process.Focus();
                    process.Select(process.Text.Length, 0);
                };
                process.Invoke(action);
                var wordFormat = docContent.IsConvert
                    ? Word.WdSaveFormat.wdFormatPDF
                    : Word.WdSaveFormat.wdFormatDocumentDefault;
                doc.SaveAs($@"{docContent.Path}\{docContent.FileName}{i}", FileFormat: wordFormat);

                Action action3 = () => { bar.Value++; };
                bar.Invoke(action3);
                
                doc.Close(SaveChanges: Word.WdSaveOptions.wdDoNotSaveChanges, OriginalFormat :Word.WdOriginalFormat.wdOriginalDocumentFormat);
            }
        }
    }
}
