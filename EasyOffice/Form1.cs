using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using EasyOffice.Properties;

namespace EasyOffice
{
    public partial class Form1 : Form
    {
        #region Для боковых кнопок

        Dictionary<Panel, Appearance> buttonsLeft;
        FieldConcat Field;

        class Appearance
        {
            public Appearance(Panel contentPanel, PictureBox iconButton, Label textButton)
            {
                ContentPanel = contentPanel;
                IconButton = iconButton;
                TextButton = textButton;
            }

            public Panel ContentPanel { get; }
            public PictureBox IconButton { get; }
            public Label TextButton { get; }
        }
        #endregion

        public Form1()
        {
            InitializeComponent();
            buttonsLeft = new Dictionary<Panel, Appearance>()
            {
                { Button1, new Appearance(ContentNumbering, Numb, label1) },
                { Button2, new Appearance(ContentConcat, Concat, label2) }
            };
            LeftMenuClick(Button1, new EventArgs());
            Field = new FieldConcat((int)numericUpDown1.Value, ContentConcat);

        }
        private void ExitClick(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Collapse_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }
        private void LeftMenuClick(object sender, EventArgs e)
        {
            var currentButton = (Panel)sender;
            foreach (var pair in buttonsLeft)
            {
                var apperance = pair.Value;
                var isSelectedPanel = currentButton == pair.Key;
                pair.Key.BackColor = isSelectedPanel ? Color.FromArgb(228, 228, 228) : Color.FromArgb(140, 140, 140);
                apperance.TextButton.ForeColor =
                    isSelectedPanel ? Color.FromArgb(140, 140, 140) : Color.FromArgb(228, 228, 228);
                apperance.IconButton.Image = Image.FromFile(Environment.CurrentDirectory + @"\Resources\" +
                                                            pair.Value.IconButton.Name +
                                                            (isSelectedPanel ? "Clicked.png" : ".png"));
                apperance.ContentPanel.Visible = isSelectedPanel;
                apperance.ContentPanel.Enabled = isSelectedPanel;
            }
        }

        private void LabelClick(object sender, EventArgs e)
        {
            LeftMenuClick(((Label)sender).Parent, new EventArgs());
        }

        private void OnCheckedChanged()
        {
            Field.Destroy();
            Field.Refresh((int)numericUpDown1.Value);
        }

        private void NumericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            OnCheckedChanged();
            Field.ChangeNumeric((int)numericUpDown1.Value);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            TextBox Process = new TextBox()
            {
                Parent = ContentConcat,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Size = new Size(ContentConcat.Size.Width, ContentConcat.Size.Height - 40),
                Location = new Point(0, 0),
                ForeColor = Color.FromArgb(114, 114, 114),
                ReadOnly = true

            };
            Process.BringToFront();
            while (outDialog2.SelectedPath == "")
            {
                DialogResult d = MessageBox.Show("Выберете папку для сохранения!");
                if (d == DialogResult.OK)
                {
                    outDialog2.ShowDialog();
                }
            }
            button5.Enabled = false;
            button5.Visible = false;
            Button CloseButton = new Button()
            {
                Parent = ContentConcat,
                Location = new Point(4, 328),
                Size = new Size(98, 30),
                Text = Resources.CloseText,
                BackColor = Color.FromArgb(114, 114, 114),
                Enabled = false,
                Margin = new Padding(0, 0, 0, 5),
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.FromArgb(228, 228, 228)
            };
            Field.StartProcess(Process, outDialog2.SelectedPath, progressBar1);

            CloseButton.Enabled = true;
            CloseButton.Click += (senderOne, eOne) =>
            {
                OnCheckedChanged();
                Process.Dispose();
                button5.Visible = true;
                progressBar1.Value = 0;
                CloseButton.Dispose();
            };
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CheckFiles_Tick(new object(), new EventArgs());
        }
        public struct Doc
        {
            public string Path { get; }
            public string FileName { get; }
            public bool IsConvert { get; }
            public Doc(string path, string fileName, bool isConvert)
            {
                Path = path;
                FileName = fileName;
                IsConvert = isConvert;
            }
        }
        private void button6_Click_1(object sender, EventArgs e)
        {
            progressBar2.Maximum = 2*((int)numericUpDown3.Value - (int)numericUpDown2.Value + 1);
            TextBox Process = new TextBox()
            {
                Parent = ContentNumbering,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Size = new Size(ContentConcat.Size.Width, ContentConcat.Size.Height - 40),
                Location = new Point(0, 0),
                ForeColor = Color.FromArgb(114, 114, 114),
                ReadOnly = true
            

            };
            Process.BringToFront();
            while (tittleOutDialog.SelectedPath == "")
            {
                DialogResult d = MessageBox.Show("Выберете папку сохранения обложек!");
                if (d == DialogResult.OK)
                {
                    tittleOutDialog.ShowDialog();
                }
            }
            while (contentOutDialog.SelectedPath == "")
            {
                DialogResult d = MessageBox.Show("Выберете папку для сохранения инструкций!");
                if (d == DialogResult.OK)
                {
                    contentOutDialog.ShowDialog();
                }
            }
            EnterNumb.Enabled = false;
            EnterNumb.Visible = false;
            Button CloseButton = new Button()
            {
                Parent = ContentNumbering,
                Location = new Point(4, 328),
                Size = new Size(98, 30),
                Text = Resources.CloseText,
                BackColor = Color.FromArgb(114, 114, 114),
                Enabled = false,
                Margin = new Padding(0, 0, 0, 5),
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.FromArgb(228, 228, 228)
            };
            Task.Factory.StartNew(() => ContentMethods.StartNumbering((int)numericUpDown2.Value, (int)numericUpDown3.Value, BoxTitle.Checked, BoxInstruction.Checked, Process, progressBar2, new Doc(tittleOutDialog.SelectedPath,textBox3.Text,checkBox2.Checked), new Doc(contentOutDialog.SelectedPath, textBox4.Text, checkBox3.Checked)));
            CloseButton.Enabled = true;
            CloseButton.Click += (senderOne, eOne) =>
            {
                OnCheckedChanged();
                Process.Dispose();
                EnterNumb.Visible = true;
                progressBar2.Value = 0;
                CloseButton.Dispose();
            };
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string argument = $"/open, \"{Environment.CurrentDirectory}\\Word\\\"";
            System.Diagnostics.Process.Start("explorer.exe", argument);
        }

        private void CheckFiles_Tick(object sender, EventArgs e)
        {
            if (!Directory.Exists($@"{Environment.CurrentDirectory}\Word\"))
                Directory.CreateDirectory($@"{Environment.CurrentDirectory}\Word\");
            
            EnterNumb.Enabled = true;
            ErrorNumb.Text = "";

            NotifyIfDocWithNameNotFound(Resources.InstructionDocText);
            NotifyIfDocWithNameNotFound(Resources.TitleDocText);
        }

        private void NotifyIfDocWithNameNotFound(string title)
        {
            if (HasDocContainingName(title))
            {
                EnterNumb.Enabled = false;
                ErrorNumb.Text += string.Format(Resources.FileWithNameNotFoundText, title);
            }
        }
        private static bool HasDocContainingName(string docFileNameSubstring) => !Directory
            .EnumerateFiles($@"{Environment.CurrentDirectory}\Word\", $"*{docFileNameSubstring}*.doc?").Any();

        private void button9_Click(object sender, EventArgs e)
        {
            tittleOutDialog.ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            contentOutDialog.ShowDialog();
        }
    }
}
