using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
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
                { Button2, new Appearance(ContentConcat, Concat, label2) },
                { Button3, new Appearance(ContentConvert, Convert, label3) },
                { Button4, new Appearance(ContentSettings, Settings, label4) }
            };
            LeftMenuClick(Button1, new EventArgs());
            Field = new FieldConcat(byFolders, (int)numericUpDown1.Value, ContentConcat);

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
            Panel currentButton = (Panel)sender;
            foreach (KeyValuePair<Panel, Appearance> item in buttonsLeft)
            {
                if (currentButton.Equals(item.Key))
                {
                    currentButton.BackColor = Color.FromArgb(228, 228, 228);
                    item.Value.TextButton.ForeColor = Color.FromArgb(140, 140, 140);
                    item.Value.IconButton.Image = Image.FromFile(Environment.CurrentDirectory + @"\Resources\" + item.Value.IconButton.Name + "Clicked.png");
                    item.Value.ContentPanel.Visible = true;
                    item.Value.ContentPanel.Enabled = true;
                    continue;
                }
                item.Key.BackColor = Color.FromArgb(140, 140, 140);
                item.Value.TextButton.ForeColor = Color.FromArgb(228, 228, 228);
                item.Value.IconButton.Image = Image.FromFile(Environment.CurrentDirectory + @"\Resources\" + item.Value.IconButton.Name + ".png");
                item.Value.ContentPanel.Visible = false;
                item.Value.ContentPanel.Enabled = false;
            }


        }

        private void LabelClick(object sender, EventArgs e)
        {
            LeftMenuClick(((Label)sender).Parent, new EventArgs());
        }

        private void CheckedChanged(object sender, EventArgs e)
        {
            FindMethod method = Field.Method;
            if (Field.CheckChanges((RadioButton)sender))
            {
                Field.Destroy();
                if (method != FindMethod.Folder)
                {
                    Field.Refresh(byFolders, (int)numericUpDown1.Value);
                }
                if (method != FindMethod.Name)
                {
                    Field.Refresh(byNames, (int)numericUpDown1.Value);
                }
            }
        }

        private void NumericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if (byNames.Checked)
            {
                CheckedChanged(byNames, new EventArgs());
            }
            else if (byFolders.Checked)
            {
                CheckedChanged(byFolders, new EventArgs());
            }
            Field.ChangeNumeric((int)numericUpDown1.Value);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            outDialog2.ShowDialog();
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
                Text = "Закрыть",
                BackColor = Color.FromArgb(114, 114, 114),
                Enabled = false,
                Margin = new Padding(0, 0, 0, 5),
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.FromArgb(228, 228, 228)
            };
            Task  task = Task.Factory.StartNew(()=> Field.StartProcess(Process, outDialog2.SelectedPath,progressBar1, textBox1.Text));

            CloseButton.Enabled = true;
            CloseButton.Click += new EventHandler((object senderOne, EventArgs eOne) =>
            {
                CheckedChanged(byNames, new EventArgs());
                CheckedChanged(byFolders, new EventArgs());
                Process.Dispose();
                button5.Visible = true;
                progressBar1.Value = 0;
                CloseButton.Dispose();
            });
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox2.Text = textBox1.Text + "1.pdf , " + textBox1.Text + "2.pdf и т.д.";
            ValidEnter(button5);
        }
        private void textBox1_TextEnter(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }

        void textBox1_TextLeave(object sender, EventArgs e)
        {

            if (textBox1.Text == "")
            {
                textBox1.Text = "Например: поз.9, 6 секц, 9 эт, кв.";
            }

        }

        private void ValidEnter(object sender)
        {
            if (textBox1.Text != "Например: поз.9, 6 секц, 9 эт, кв." && textBox1.Text != "")
            {
                button5.Enabled = true;
            }
            else
            {
                button5.Enabled = false;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (ContentMethods.CheckHasWord())
            {
                ContentConvert.BackColor = Color.LawnGreen;
            }
            else
            {
                ContentConvert.BackColor = Color.OrangeRed;
            }
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
                Text = "Закрыть",
                BackColor = Color.FromArgb(114, 114, 114),
                Enabled = false,
                Margin = new Padding(0, 0, 0, 5),
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.FromArgb(228, 228, 228)
            };
            Task task = Task.Factory.StartNew(() => ContentMethods.StartNumbering((int)numericUpDown2.Value, (int)numericUpDown3.Value, BoxTitle.Checked, BoxInstruction.Checked, Process, progressBar2, new Doc(tittleOutDialog.SelectedPath,textBox3.Text,checkBox2.Checked), new Doc(contentOutDialog.SelectedPath, textBox4.Text, checkBox3.Checked)));
            CloseButton.Enabled = true;
            CloseButton.Click += new EventHandler((object senderOne, EventArgs eOne) =>
            {
                CheckedChanged(byNames, new EventArgs());
                CheckedChanged(byFolders, new EventArgs());
                Process.Dispose();
                EnterNumb.Visible = true;
                progressBar2.Value = 0;
                CloseButton.Dispose();
            });
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string argument = $"/open, \"{Environment.CurrentDirectory}\\Word\\\"";
            System.Diagnostics.Process.Start("explorer.exe", argument);
        }

        private void CheckFiles_Tick(object sender, EventArgs e)
        {
            EnterNumb.Enabled = true;
            ErrorNumb.Text = "";
            if (!Directory.EnumerateFiles($@"{Environment.CurrentDirectory}\Word\", "Инструкция.*").Any())
            {
                EnterNumb.Enabled = false;
                ErrorNumb.Text = "Файл 'Инструкция.doc' не найден! ";
            }
            if (!Directory.EnumerateFiles($@"{Environment.CurrentDirectory}\Word\", "Обложка.*").Any())
            {
                EnterNumb.Enabled = false;
                ErrorNumb.Text += "Файл 'Обложка.docx' не найден! ";
            }
        }

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
