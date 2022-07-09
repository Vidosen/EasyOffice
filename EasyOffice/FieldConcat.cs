using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using EasyOffice.Properties;

namespace EasyOffice
{
    public class FieldConcat
    {
        Panel InternalPanelConcat;
        Label pathConcat;
        public List<FolderBrowserDialog> FileDirectory { get; set; }
        public List<TextBox> ShowValue { get; set; }
        public List<Button> Open { get; set; }
        TextBox ShowValueName;
        Button OpenName;
        FolderBrowserDialog FileDirectoryName;
        public int Count;
        public FieldConcat(int count, Panel parent)
        {
            Count = count;
            InternalPanelConcat = new Panel
            {
                Parent = parent,
                Location = new Point(0, 160),
                Size = new Size(552, 164),
                AutoScroll = true
            };
            FileDirectory = new List<FolderBrowserDialog>();
            Open = new List<Button>();
            ShowValue = new List<TextBox>();
            pathConcat = new Label
            {
                Parent = InternalPanelConcat,
                Location = new Point(0, 0),
            };

            pathConcat.Size = new Size(150, 20);
            pathConcat.Text = "Выберете папки:";
            for (int i = 0; i < count; i++)
            {
                FileDirectory.Add(new FolderBrowserDialog()
                {
                    RootFolder = Environment.SpecialFolder.Desktop,
                    ShowNewFolderButton = true
                });
                Open.Add(new Button()
                {
                    Parent = InternalPanelConcat,
                    Text = "Выбрать папку",
                    Location = new Point(252, 40 * i + 34),
                    Size = new Size(152, 30),
                    TextAlign = ContentAlignment.MiddleCenter,
                    FlatStyle = FlatStyle.Flat

                });
                ShowValue.Add(new TextBox()
                {
                    Parent = InternalPanelConcat,
                    Text = $"Папка {i + 1} не указана",
                    Location = new Point(0, 40 * i + 36),
                    Size = new Size(242, 40),
                    ReadOnly = true,
                    ForeColor = ContentMethods.ForeColor

                });
                Open[i].Click += (sender, e) =>
                {
                    for (int k = 0; k < Open.Count; k++)
                    {
                        if ((Button)sender == Open[k])
                        {
                            if (FileDirectory[k].ShowDialog() == DialogResult.OK)
                            {
                                string currPath = FileDirectory[k].SelectedPath;
                                ShowValue[k].Text = currPath.Contains($@"C:\Users\{Environment.UserName}\")
                                    ? "..." + currPath.Remove(
                                        FileDirectory[k].SelectedPath.IndexOf($@"C:\Users\{Environment.UserName}\"),
                                        $@"C:\Users\{Environment.UserName}".Length)
                                    : currPath;
                            }
                        }
                    }
                };
            }
        }
        public void Refresh(int count)
        {
            pathConcat.Size = new Size(150, 20);
            pathConcat.Text = "Выберете папки:";
            for (int i = 0; i < count; i++)
            {
                FileDirectory.Add(new FolderBrowserDialog()
                {
                    RootFolder = Environment.SpecialFolder.Desktop,
                    ShowNewFolderButton = true
                });
                Open.Add(new Button()
                {
                    Parent = InternalPanelConcat,
                    Text = "Выбрать папку",
                    Location = new Point(252, 40 * i + 34),
                    Size = new Size(152, 30),
                    TextAlign = ContentAlignment.MiddleCenter,
                    FlatStyle = FlatStyle.Flat

                });
                ShowValue.Add(new TextBox()
                {
                    Parent = InternalPanelConcat,
                    ReadOnly = true,
                    Text = $"Папка {i + 1} не указана",
                    Location = new Point(0, 40 * i + 36),
                    Size = new Size(242, 40),
                    ForeColor = ContentMethods.ForeColor

                });
                Open[i].Click += (sender, e) =>
                {
                    for (int k = 0; k < Open.Count; k++)
                    {
                        if ((Button)sender == Open[k])
                        {
                            if (FileDirectory[k].ShowDialog() == DialogResult.OK)
                            {
                                string currPath = FileDirectory[k].SelectedPath;
                                ShowValue[k].Text = currPath.Contains($@"C:\Users\{Environment.UserName}\")
                                    ? "..." + currPath.Remove(
                                        FileDirectory[k].SelectedPath.IndexOf($@"C:\Users\{Environment.UserName}\"),
                                        $@"C:\Users\{Environment.UserName}".Length)
                                    : currPath;
                            }
                        }
                    }
                };
            }
        }

        public void Destroy(int delAt = 0)
        {
            int OpenCount = Open.Count;
            int ShowValueCount = ShowValue.Count;
            int FileDirectoryCount = FileDirectory.Count;
            for (int i = --OpenCount; i >= delAt; i--)
            {
                InternalPanelConcat.Controls.Remove(Open[i]);
                Open[i].Dispose();
                Open.RemoveAt(i);
            }

            for (int i = --ShowValueCount; i >= delAt; i--)
            {
                InternalPanelConcat.Controls.Remove(ShowValue[i]);
                ShowValue[i].Dispose();
                ShowValue.RemoveAt(i);
            }

            for (int i = --FileDirectoryCount; i >= delAt; i--)
            {
                FileDirectory[i].Dispose();
                FileDirectory.RemoveAt(i);
            }
        }
        public void ChangeNumeric(int newCount)
        {
            if (newCount > Count)
            {
                for (int i = Open.Count; i < newCount; i++)
                {
                    FileDirectory.Add(new FolderBrowserDialog()
                    {
                        RootFolder = Environment.SpecialFolder.Desktop,
                        ShowNewFolderButton = true
                    });
                    Open.Add(new Button()
                    {
                        Parent = InternalPanelConcat,
                        Text = "Выбрать папку",
                        Location = new Point(252, 40 * i + 34),
                        Size = new Size(152, 30),
                        TextAlign = ContentAlignment.MiddleCenter,
                        FlatStyle = FlatStyle.Flat

                    });
                    ShowValue.Add(new TextBox()
                    {
                        Parent = InternalPanelConcat,
                        ReadOnly = true,
                        Text = $"Папка {i + 1} не указана",
                        Location = new Point(0, 40 * i + 36),
                        Size = new Size(242, 40),
                        ForeColor = ContentMethods.ForeColor

                    });
                    Open[i].Click += (sender, e) =>
                    {
                        for (int k = 0; k < Open.Count; k++)
                        {
                            if ((Button)sender == Open[k])
                            {
                                if (FileDirectory[k].ShowDialog() == DialogResult.OK)
                                {
                                    string currPath = FileDirectory[k].SelectedPath;
                                    ShowValue[k].Text = currPath.Contains($@"C:\Users\{Environment.UserName}\")
                                        ? "..." + currPath.Remove(
                                            FileDirectory[k].SelectedPath.IndexOf($@"C:\Users\{Environment.UserName}\"),
                                            $@"C:\Users\{Environment.UserName}".Length)
                                        : currPath;
                                }
                            }
                        }
                    };
                }
            }
            if (newCount < Count)
            {
                Destroy(newCount);
            }
            Count = newCount;
        }
        public void StartProcess(TextBox process, string outDialog, ProgressBar progress)
        {
            List<string> concatDierctories = new List<string>();
            foreach (var dir in FileDirectory)
            {
                if (dir.SelectedPath != "")
                    concatDierctories.Add(dir.SelectedPath);
            }

            List<Dictionary<int, string>> indexedDirFiles = new List<Dictionary<int, string>>();
            foreach (var filesDir in concatDierctories)
            {
                var filesList = ContentMethods.GetFilesPath(filesDir);
                try
                {
                    indexedDirFiles.Add(ContentMethods.ParseIdsFromFilePathsAsDictionary(filesList));
                }
                catch (ArgumentException)
                {
                    var duplicateFilePaths = ContentMethods.ParseIdsFromFilePathsAsLookup(filesList)
                        .Where(group=> group.Count() > 1)
                        .SelectMany(group=> group);
                    process.Text += Resources.ErrorTag + "Папка содержит дублирующиеся файлы!" + Environment.NewLine;
                    foreach (var path in duplicateFilePaths)
                        process.Text += path + Environment.NewLine;
                    return;
                }
            }
            var indexedFiles = indexedDirFiles.SelectMany(list => list)
                .GroupBy(pair => pair.Key)
                .ToList();
            
            progress.Maximum = indexedFiles.Count;
            foreach (var group in indexedFiles)
            {
                ContentMethods.ConcatPDF(group.Select(pair=> pair.Value + ".pdf").ToArray(), outDialog);
                
                process.Text += $"Документ {group.Key} собран." + Environment.NewLine;
                progress.Value++;
            }
        }
    }
}
