using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EasyOffice
{
    public class FieldConcat
    {
        public FindMethod Method { get; set; }
        Panel InternalPanelConcat;
        Label pathConcat;
        public List<FolderBrowserDialog> FileDirectory { get; set; }
        public List<TextBox> ShowValue { get; set; }
        public List<Button> Open { get; set; }
        TextBox ShowValueName;
        Button OpenName;
        FolderBrowserDialog FileDirectoryName;
        public int Count;
        public FieldConcat(RadioButton check, int count, Panel parent)
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
            if (check.Name == "byFolders")
            {
                pathConcat.Size = new Size(150, 20);
                pathConcat.Text = "Выберете папки:";
                Method = FindMethod.Folder;
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
                    Open[i].Click += new EventHandler((object sender, EventArgs e) =>
                    {
                        for (int k = 0; k < Open.Count; k++)
                        {
                            if ((Button)sender == Open[k])
                            {
                                if (FileDirectory[k].ShowDialog() == DialogResult.OK)
                                {
                                    string currPath = FileDirectory[k].SelectedPath;
                                    ShowValue[k].Text = currPath.Contains($@"C:\Users\{Environment.UserName}\") ? "..." + currPath.Remove(FileDirectory[k].SelectedPath.IndexOf($@"C:\Users\{Environment.UserName}\"), $@"C:\Users\{Environment.UserName}".Length) : currPath;
                                }
                            }
                        }
                    });
                }
            }
            if (check.Name == "byNames")
            {
                pathConcat.Size = new Size(233, 38);
                pathConcat.Text = "Выберете папку и названия\n каждого типа файлов";
                Method = FindMethod.Name;
                FileDirectoryName = new FolderBrowserDialog()
                {
                    RootFolder = Environment.SpecialFolder.Desktop,
                    ShowNewFolderButton = true
                };
                OpenName = new Button()
                {
                    Parent = InternalPanelConcat,
                    Text = "Выбрать папку",
                    Location = new Point(252, 40),
                    Size = new Size(152, 30),
                    TextAlign = ContentAlignment.MiddleCenter,
                    FlatStyle = FlatStyle.Flat

                };
                ShowValueName = new TextBox()
                {
                    Parent = InternalPanelConcat,
                    ReadOnly = true,
                    Text = "Папка не указана",
                    Location = new Point(0, 42),
                    Size = new Size(242, 40),
                    ForeColor = ContentMethods.ForeColor
                };
                OpenName.Click += new EventHandler((object sender, EventArgs e) =>
                {
                    if (FileDirectoryName.ShowDialog() == DialogResult.OK)
                    {
                        string currPath = FileDirectoryName.SelectedPath;
                        ShowValueName.Text = currPath.Contains($@"C:\Users\{Environment.UserName}\") ? "..." + currPath.Remove(FileDirectoryName.SelectedPath.IndexOf($@"C:\Users\{Environment.UserName}\"), $@"C:\Users\{Environment.UserName}".Length) : currPath;
                    }
                });
                for (int i = 0; i < count; i++)
                {
                    ShowValue.Add(new TextBox()
                    {
                        Parent = InternalPanelConcat,
                        Text = $"Имя {i + 1}",
                        Location = new Point(0, i * 40 + 82),
                        Size = new Size(242, 40),
                        ForeColor = ContentMethods.ForeColor,
                        BorderStyle = BorderStyle.None
                    });
                    ShowValue[i].Click += new EventHandler((object sender, EventArgs e) =>
                    {
                        for (int k = 0; k < ShowValue.Count; k++)
                        {
                            if ((TextBox)sender == ShowValue[k])
                            {
                                ShowValue[k].Text = "";
                            }
                        }
                    });
                    ShowValue[i].Leave += new EventHandler((object sender, EventArgs e) =>
                    {
                        for (int k = 0; k < ShowValue.Count; k++)
                        {
                            if ((TextBox)sender == ShowValue[k] && ShowValue[k].Text == "")
                            {
                                ShowValue[k].Text = $"Имя {k + 1}";
                            }
                        }
                    });
                }
            }
        }
        public void Refresh(RadioButton check, int count)
        {
            if (check.Name == "byFolders")
            {
                pathConcat.Size = new Size(150, 20);
                pathConcat.Text = "Выберете папки:";
                Method = FindMethod.Folder;
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
                    Open[i].Click += new EventHandler((object sender, EventArgs e) =>
                    {
                        for (int k = 0; k < Open.Count; k++)
                        {
                            if ((Button)sender == Open[k])
                            {
                                if (FileDirectory[k].ShowDialog() == DialogResult.OK)
                                {
                                    string currPath = FileDirectory[k].SelectedPath;
                                    ShowValue[k].Text = currPath.Contains($@"C:\Users\{Environment.UserName}\") ? "..." + currPath.Remove(FileDirectory[k].SelectedPath.IndexOf($@"C:\Users\{Environment.UserName}\"), $@"C:\Users\{Environment.UserName}".Length) : currPath;
                                }
                            }
                        }
                    });
                }
            }
            if (check.Name == "byNames")
            {
                pathConcat.Size = new Size(233, 38);
                pathConcat.Text = "Выберете папку и названия\n каждого типа файлов";
                Method = FindMethod.Name;
                FileDirectoryName = new FolderBrowserDialog()
                {
                    RootFolder = Environment.SpecialFolder.Desktop,
                    ShowNewFolderButton = true
                };
                OpenName = new Button()
                {
                    Parent = InternalPanelConcat,
                    Text = "Выбрать папку",
                    Location = new Point(252, 40),
                    Size = new Size(152, 30),
                    TextAlign = ContentAlignment.MiddleCenter,
                    FlatStyle = FlatStyle.Flat

                };
                ShowValueName = new TextBox()
                {
                    Parent = InternalPanelConcat,
                    ReadOnly = true,
                    Text = "Папка не указана",
                    Location = new Point(0, 42),
                    Size = new Size(242, 40),
                    ForeColor = ContentMethods.ForeColor
                };
                OpenName.Click += new EventHandler((object sender, EventArgs e) =>
                {
                        if (FileDirectoryName.ShowDialog() == DialogResult.OK)
                        {
                            string currPath = FileDirectoryName.SelectedPath;
                            ShowValueName.Text = currPath.Contains($@"C:\Users\{Environment.UserName}\") ? "..." + currPath.Remove(FileDirectoryName.SelectedPath.IndexOf($@"C:\Users\{Environment.UserName}\"), $@"C:\Users\{Environment.UserName}".Length) : currPath;
                        }
                });
                for (int i = 0; i < count; i++)
                {
                    ShowValue.Add(new TextBox()
                    {
                        Parent = InternalPanelConcat,
                        Text = $"Имя {i + 1}",
                        Location = new Point(0, i * 40 + 82),
                        Size = new Size(242, 40),
                        BorderStyle = BorderStyle.None,
                        ForeColor = ContentMethods.ForeColor
                    });
                    ShowValue[i].Click += new EventHandler((object sender, EventArgs e) =>
                    {
                        for (int k = 0; k < ShowValue.Count; k++)
                        {
                            if ((TextBox)sender == ShowValue[k])
                            {
                                ShowValue[k].Text = "";
                            }
                        }
                    });
                    ShowValue[i].Leave += new EventHandler((object sender, EventArgs e) =>
                    {
                        for (int k = 0; k < ShowValue.Count; k++)
                        {
                            if ((TextBox)sender == ShowValue[k] && ShowValue[k].Text == "")
                            {
                                ShowValue[k].Text = $"Имя {k + 1}";
                            }
                        }
                    });
                }
            }
        }
        public bool CheckChanges(RadioButton check)
        {
            if (check.Name == "byFolders" && Method != FindMethod.Folder)
            {
                return true;
            }
            if (check.Name == "byNames" && Method != FindMethod.Name)
            {
                return true;
            }
            return false;
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
            if (delAt == 0 & Method == FindMethod.Name)
            {
                InternalPanelConcat.Controls.Remove(OpenName);
                InternalPanelConcat.Controls.Remove(ShowValueName);
                OpenName.Dispose();
                ShowValueName.Dispose();
                FileDirectoryName.Dispose();

            }

        }
        public void ChangeNumeric(int newCount)
        {
            if (newCount > Count)
            {
                if (FindMethod.Folder == Method)
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
                        Open[i].Click += new EventHandler((object sender, EventArgs e) =>
                        {
                            for (int k = 0; k < Open.Count; k++)
                            {
                                if ((Button)sender == Open[k])
                                {
                                    if (FileDirectory[k].ShowDialog() == DialogResult.OK)
                                    {
                                        string currPath = FileDirectory[k].SelectedPath;
                                        ShowValue[k].Text = currPath.Contains($@"C:\Users\{Environment.UserName}\") ? "..." + currPath.Remove(FileDirectory[k].SelectedPath.IndexOf($@"C:\Users\{Environment.UserName}\"), $@"C:\Users\{Environment.UserName}".Length) : currPath;
                                    }
                                }
                            }
                        });
                    }
                }

                if (FindMethod.Name == Method)
                {
                    for (int i = ShowValue.Count; i < newCount; i++)
                    {
                        ShowValue.Add(new TextBox()
                        {
                            Parent = InternalPanelConcat,
                            Text = $"Имя {i + 1}",
                            Location = new Point(0, i * 40 + 82),
                            Size = new Size(242, 40),
                            ForeColor = ContentMethods.ForeColor,
                            BorderStyle = BorderStyle.None
                        });
                        ShowValue[i].Click += new EventHandler((object sender, EventArgs e) =>
                        {
                            for (int k = 0; k < ShowValue.Count; k++)
                            {
                                if ((TextBox)sender == ShowValue[k])
                                {
                                    ShowValue[k].Text = "";
                                }
                            }
                        });
                        ShowValue[i].Leave += new EventHandler((object sender, EventArgs e) =>
                        {
                            for (int k = 0; k < ShowValue.Count; k++)
                            {
                                if ((TextBox)sender == ShowValue[k] && ShowValue[k].Text == "")
                                {
                                    ShowValue[k].Text = $"Имя {k}";
                                }
                            }
                        });
                    }
                }
            }
            if (newCount < Count)
            {
                if (FindMethod.Folder == Method)
                {
                    Destroy(newCount);
                }
                if (FindMethod.Name == Method)
                {
                    Destroy(newCount);
                }
            }
            Count = newCount;
        }
        public void StartProcess(TextBox process, string outDialog, ProgressBar progress, string FileName)
        {
            List<string> dirIn = new List<string>();
            foreach (var dir in FileDirectory)
            {
                if (dir.SelectedPath != "")
                {
                    dirIn.Add(dir.SelectedPath);
                }
            }
            List<string> NameIn = new List<string>();
            foreach (var name in ShowValue)
            {
                if (!name.Text.Contains("Имя " + ShowValue.IndexOf(name) + 1))
                {
                    NameIn.Add(name.Text);
                }
            }
            string dirName = null;
            if (FileDirectoryName != null)
            {
                dirName = FileDirectoryName.SelectedPath != "" ? FileDirectoryName.SelectedPath : null;
            }
            List<List<string>> filesList = ContentMethods.GetFilesPath(Method, dirIn, dirName, NameIn);
            var getIdexedFiles = filesList.Select(ContentMethods.ParseIdsFromFilePaths).SelectMany(list=> list).GroupBy(pair=> pair.Key);
            int maxCount = 0;
            int maxCountId = 0;
            for (int i = 0; i < filesList.Count - 1;i++)
            {
                if (filesList[i].Count > maxCount)
                {
                    maxCount = filesList[i].Count;
                    maxCountId = i;
                }
            }
            Action SetUp = () => { progress.Maximum = maxCount; };
            progress.Invoke(SetUp);
            var idexedFiles = getIdexedFiles.ToList();
            foreach (var group in idexedFiles)
            {
                ContentMethods.ConcatPDF(group.Select(pair=> pair.Value + ".pdf").ToArray(), outDialog + $"//{FileName}{group.Key}.pdf");
                Action print = () => { process.Text += $"Документ {group.Key} собран." + Environment.NewLine; };
                process.Invoke(print);
                
                Action increment = () => { progress.Value++; };
                progress.Invoke(increment);
            }
            string ConcatCurrent(int id)
            {
                List<string> inFiles = new List<string>();

                for (int i = 0; i < filesList.Count; i++)
                {
                    for (int k = 0; k < filesList[i].Count; k++)
                    {

                        if (filesList[i][k].Contains(id.ToString()))
                        {

                            inFiles.Add(filesList[i][k] + ".pdf");
                            break;
                        }
                        else if(k == filesList[i].Count - 1)
                        {
                            return $"Файл #{i + 1} для {id} документа не найден.";
                        }
                    }
                }
                ContentMethods.ConcatPDF(inFiles.ToArray(), outDialog + $"//{FileName}{id}.pdf");
                return $"Документ {id} собран.";
            }



        }



    }
}
