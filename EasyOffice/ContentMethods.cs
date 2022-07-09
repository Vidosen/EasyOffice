using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace EasyOffice
{
    public enum FindMethod
    {
        Folder,
        Name
    }

    static public class ContentMethods
    {
        static public Color ForeColor = Color.FromArgb(140, 140, 140);
        static public Color BackColor = Color.FromArgb(228, 228, 228);
        static public Color TextColor = Color.FromArgb(114, 114, 114);
        static public Task ConcatPDF(string[] inFiles, string outDirectory)
        {

            PdfDocument outputDocument = new PdfDocument();
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
            outputDocument.Save(outDirectory);
            outputDocument.Close();
            return Task.Run(() => outputDocument.Dispose());
        }

        public static void ConvertDocx(Word.Document doc, Word.Application app, string path, string exportName, int from, int to)
        {
            doc.ExportAsFixedFormat(exportName, Word.WdExportFormat.wdExportFormatPDF,Range:Word.WdExportRange.wdExportFromTo,From:from,To:to);
        }

        public static Dictionary<int, string> ParseIdsFromFilePaths(List<string> inFiles)
        {
            return inFiles.ToDictionary(path =>
            {
                var splitPath = path.Split(new[]
                {
                    ' ', '_'
                }, StringSplitOptions.RemoveEmptyEntries);
                return int.Parse(splitPath.Last());
            }, path => path);
        }
        static public List<List<string>> GetFilesPath(FindMethod method, List<string> dirIn = null, string dirName = null, List<string> nameIn = null, string format = ".pdf")
        {
            if (method == FindMethod.Folder)
            {
                List<List<string>> result = new List<List<string>>();
                foreach (var dir in dirIn)
                {
                    if (Directory.Exists(dir))
                    {
                        List<string> maskFormat = Directory.GetFiles(dir).OrderBy(file=> file).ToList();
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
                        result.Add(maskFormat);
                    }
                }
                return result;
            }
            else if (method == FindMethod.Name)
            {
                if (Directory.Exists(dirName))
                {
                    List<List<string>> result = new List<List<string>>();
                    foreach(var name in nameIn)
                    {

                        string[] tmpDirectory = Directory.GetFiles(dirName);
                        List<string> tmpFiles = new List<string>();
                        foreach (string item in tmpDirectory)
                        {
                            if (item.Contains(name) && item.EndsWith(format))
                            {

                                tmpFiles.Add(item.Remove(item.Length - 4, 4));
                            }
                        }
                        result.Add(tmpFiles);
                    }
                    return result;
                }
            }
            throw new Exception("Ошибка выполнения функции \'GetFilesPath\'");
        }
        public static bool CheckHasWord()
        {
            
            return Directory.Exists(@"C:\Program Files (x86)\Microsoft Office");
        }

        public static void StartNumbering(int start, int end, bool title, bool instruction,TextBox process, ProgressBar bar, Form1.Doc docTitle, Form1.Doc docContent)
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
            
            Word.Document doc = app.Documents.Open(Directory.EnumerateFiles($@"{Environment.CurrentDirectory}\Word\", $"{fileName}.*")
                .First());
            for (int i = start; i <= end; i++)
            {
                var bookmarks = doc.Bookmarks;
                foreach (var bookmark in bookmarks.Cast<Word.Bookmark>().ToList())
                {
                    var bookmarkName = bookmark.Name;
                    var newRange = doc.Range(bookmark.Range.Start, bookmark.Range.End);
                    newRange.Text = i.ToString();
                    if (!doc.Bookmarks.Exists(bookmarkName))
                    {
                        doc.Bookmarks.Add(bookmarkName, newRange);
                    }
                }

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
            }
            doc.Close(SaveChanges: Word.WdSaveOptions.wdDoNotSaveChanges, OriginalFormat :Word.WdOriginalFormat.wdOriginalDocumentFormat);
        }
    }
}
