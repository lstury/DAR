using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Data.SQLite;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;



namespace Daily_Achievement_Report
{
    public partial class DAR : Form
    {
        private Timer timer;
        private bool fadeIn;
        private readonly FontFamily customFontFamily = new FontFamily("Sylfaen");
        List<System.Windows.Forms.RichTextBox> textBoxes = new List<System.Windows.Forms.RichTextBox>();
        Dictionary<string, string> textBoxContents = new Dictionary<string, string>();
        private System.Windows.Forms.RichTextBox activeTextBox;
        private FontStyle currentFontStyle;

        public DAR()
        {
            InitializeComponent();
            InitializeForm();
            flowLayoutPanel1.FlowDirection = FlowDirection.TopDown;
            LoadCheckboxes();
            this.button36.Click += new System.EventHandler(this.Button36_Click);
            this.button37.Click += new System.EventHandler(this.Button37_Click);
            SubscribeCommonEventHandlers();
            SubscribeSecondCommonEventHandlers();
            this.Load += new System.EventHandler(this.Form2_Load);
            currentFontStyle = FontStyle.Regular;
            textBox1.KeyDown += new KeyEventHandler(textBox1_KeyDown);
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            panel9.Hide();
            panel15.Visible = false;
            panel2.Visible = true;
            SuccessSave.Visible = false;
            ClearText.Visible = false;
            WorkdayText.Visible = false;
            FileSaveText.Visible = false;
            LoadCheckboxes();
            panel10.Hide();
            jan228.Hide();
            jan227.Hide();
            jan9.Hide();
            jan16.Hide();
            jan13.Hide();
            jan14.Hide();
            jan30.Hide();
            jan31.Hide();
            feb3.Hide();
            feb4.Hide();
            feb6.Hide();
            Toolbox.Hide();
            this.ActiveControl = label2;
            LoadContent();
            SetTextBoxNames();
            italicOn.Image = ResizeImage(Properties.Resources.italic_button, 33, 27);
            boldOn.Image = ResizeImage(Properties.Resources.bold_button, 33, 27);
            foreach (Control ctrl in panel10.Controls)
            {
                if (ctrl is RichTextBox richTextBox)
                {
                    richTextBox.GotFocus += RichTextBox_GotFocus;
                }

            }

        }
        //
        //FADE IN
        //
        private void InitializeForm()
        {
            this.Opacity = 0;
            timer = new Timer
            {
                Interval = 10
            };
            timer.Tick += Timer_Tick;
            fadeIn = true;
            timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            AdjustOpacity();
        }

        private void AdjustOpacity()
        {
            if (fadeIn)
            {
                if (this.Opacity < 1)
                {
                    this.Opacity += 0.05;
                }
                else
                {
                    timer.Stop();
                }
            }
            else
            {
                if (this.Opacity > 0)
                {
                    this.Opacity -= 0.05;
                }
                else
                {
                    timer.Stop();
                    this.Close();
                }
            }
        }
        //
        //DRAG FORM
        //
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();

        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        private void PanelTitleBar_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.Style |= 0x20000;
                return cp;
            }
        }

        private void Form2_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }
        //
        //DATABASE
        //
        public class DatabaseHelper
        {
            private string connectionString = @"Data Source=C:\Users\LT-49\Desktop\immersion\lucky\MiniDB.db;Version=3;";

            public void CreateDatabase()
            {
                string dbFilePath = @"C:\Users\LT-49\Desktop\immersion\lucky\MiniDB.db";

                if (!File.Exists(dbFilePath)) 
                {
                    SQLiteConnection.CreateFile(dbFilePath);
                }

                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();

                    string createTableQuery = @"
                CREATE TABLE IF NOT EXISTS TextDocuments (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Title TEXT,
                    RtfContent TEXT,
                    CreatedAt DATETIME,
                    UpdatedAt DATETIME
                );";

                    using (var command = new SQLiteCommand(createTableQuery, connection))
                    {
                        command.ExecuteNonQuery();
                    }
                }
            }

            public void SaveRtfToDatabase(string title, string rtfContent)
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();

                    string insertQuery = @"
                INSERT INTO TextDocuments (Title, RtfContent, CreatedAt, UpdatedAt) 
                VALUES (@Title, @RtfContent, @CreatedAt, @UpdatedAt)";

                    using (var command = new SQLiteCommand(insertQuery, connection))
                    {
                        command.Parameters.AddWithValue("@Title", title);
                        command.Parameters.AddWithValue("@RtfContent", rtfContent);
                        command.Parameters.AddWithValue("@CreatedAt", DateTime.Now);
                        command.Parameters.AddWithValue("@UpdatedAt", DateTime.Now);

                        command.ExecuteNonQuery();
                    }
                }
            }
        }
        //
        //SAVING & COMPILING LOGIC
        //
        private void SetTextBoxNames()
        {
            jan228.Name = "January 28";
            jan227.Name = "January 27";
            jan16.Name = "January 16";
            jan13.Name = "January 13";
            jan14.Name = "January 14";
            jan30.Name = "January 30";
            jan31.Name = "January 31";
            feb3.Name = "February 3";
            feb4.Name = "February 4";
            feb6.Name = "February 6";

            textBoxes.AddRange(new[] { jan13, jan14, jan16, jan228, jan227, jan30, jan31, feb3, feb4, feb6 });
        }

        private void LoadContent()
        {
            if (activeTextBox != null)
            {
                string connectionString = @"Data Source=C:\Users\LT-49\Desktop\immersion\lucky\MiniDB.db;Version=3;";
                try
                {
                    using (var connection = new SQLiteConnection(connectionString))
                    {
                        connection.Open();
                        string selectQuery = "SELECT RtfContent FROM TextDocuments WHERE Title = @Title";

                        using (var command = new SQLiteCommand(selectQuery, connection))
                        {
                            command.Parameters.AddWithValue("@Title", activeTextBox.Name);
                            using (SQLiteDataReader reader = command.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    string rtfContent = reader["RtfContent"].ToString();

                                    if (!string.IsNullOrEmpty(rtfContent))
                                    {
                                        activeTextBox.Rtf = rtfContent;
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error loading content from database: {ex.Message}");
                }
            }
        }

        private void Save_Click(object sender, EventArgs e)
        {
            if (activeTextBox != null && panel10.Controls.Contains(activeTextBox))
            {
                var dbHelper = new DatabaseHelper();
                dbHelper.CreateDatabase();

                string title = activeTextBox.Name;
                string rtfContent = activeTextBox.Rtf;

                if (DocumentExists(title))
                {
                    UpdateDocumentContent(title, rtfContent);
                    SuccessSave.Visible = true;
                    SuccessSave.Location = new System.Drawing.Point(48, 190);
                }
                else
                {
                    dbHelper.SaveRtfToDatabase(title, rtfContent);
                    SuccessSave.Visible = true;
                    SuccessSave.Location = new System.Drawing.Point(48, 190);
                }
            }
            else
            {
                MessageBox.Show("No active RichTextBox found or it's not inside panel10.");
            }
        }

        private void CloseSave_Click(object sender, EventArgs e)
        {
            SuccessSave.Visible = false;
        }

        private bool DocumentExists(string title)
        {
            string connectionString = @"Data Source=C:\Users\LT-49\Desktop\immersion\lucky\MiniDB.db;Version=3;";
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();
                    string selectQuery = "SELECT COUNT(*) FROM TextDocuments WHERE Title = @Title";

                    using (var command = new SQLiteCommand(selectQuery, connection))
                    {
                        command.Parameters.AddWithValue("@Title", title);
                        int count = Convert.ToInt32(command.ExecuteScalar());
                        return count > 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error checking document existence: {ex.Message}");
                return false;
            }
        }

        private void UpdateDocumentContent(string title, string rtfContent)
        {
            string connectionString = @"Data Source=C:\Users\LT-49\Desktop\immersion\lucky\MiniDB.db;Version=3;";
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();
                    string updateQuery = "UPDATE TextDocuments SET RtfContent = @RtfContent WHERE Title = @Title";

                    using (var command = new SQLiteCommand(updateQuery, connection))
                    {
                        command.Parameters.AddWithValue("@Title", title);
                        command.Parameters.AddWithValue("@RtfContent", rtfContent);
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating document content: {ex.Message}");
            }
        }

        private void button34_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false;
            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Add();


            foreach (var textBox in textBoxes)
            {
                if (textBox != null && !string.IsNullOrWhiteSpace(textBox.Text))
                {
                    string rtfContent = textBox.Rtf;

                    Range range = doc.Content;
                    range.InsertAfter("\n");
                    range.InsertAfter("Date: " + textBox.Name + "\n\n");

                    RichTextBox tempRtfBox = new RichTextBox();
                    tempRtfBox.Rtf = rtfContent;

                    tempRtfBox.SelectAll();
                    tempRtfBox.Copy();

                    range = doc.Content;
                    range.Collapse(WdCollapseDirection.wdCollapseEnd); 
                    range.Paste();

                    range.Collapse(WdCollapseDirection.wdCollapseEnd);
                }
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Word Files (*.docx)|*.docx",
                Title = "Save Word Document As",
                FileName = "Daily Achievement Report.docx"
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string docxFilePath = saveFileDialog.FileName;

                Debug.WriteLine($"File selected for saving: {docxFilePath}");

                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(docxFilePath);
                string fileExtension = Path.GetExtension(docxFilePath);
                string directory = Path.GetDirectoryName(docxFilePath);

                int fileCounter = 1;
                while (File.Exists(docxFilePath))
                {
                    docxFilePath = Path.Combine(directory, $"{fileNameWithoutExtension} ({fileCounter}){fileExtension}");
                    fileCounter++;
                }

                Debug.WriteLine($"Saving file as: {docxFilePath}");

                doc.SaveAs2(docxFilePath);
                doc.Close();
                wordApp.Quit();

                Marshal.ReleaseComObject(doc);
                Marshal.ReleaseComObject(wordApp);

                FileSaveText.Visible = true;
                FileSaveText.Location = new System.Drawing.Point(48, 190);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FileSaveText.Visible = false;
            
        }
        //
        //TO DO LIST
        //
        private void Button35_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.CheckBox checkBox = new System.Windows.Forms.CheckBox();
            checkBox.Text = textBox1.Text;
            checkBox.Font = new System.Drawing.Font(customFontFamily, 12);
            flowLayoutPanel1.Controls.Add(checkBox);
            checkBox.AutoSize = true;
            SaveCheckboxes();
            checkBox.CheckedChanged += new EventHandler(DynamicCheckBox_CheckedChanged);
            ApplyStrikethrough(checkBox);
            textBox1.Clear();
        }

        private void DynamicCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.CheckBox checkBox = sender as System.Windows.Forms.CheckBox;
            ApplyStrikethrough(checkBox);
            SaveCheckboxes();
        }

        private void ApplyStrikethrough(System.Windows.Forms.CheckBox checkBox)
        {
            if (checkBox.Checked)
            {
                checkBox.Font = new System.Drawing.Font(customFontFamily, 12, FontStyle.Strikeout);
            }
            else
            {
                checkBox.Font = new System.Drawing.Font(customFontFamily, 12, FontStyle.Regular);
            }
        }

        private void SaveCheckboxes()
        {
            List<string> checkboxStates = new List<string>();
            foreach (Control control in flowLayoutPanel1.Controls)
            {
                if (control is System.Windows.Forms.CheckBox checkBox)
                {
                    checkboxStates.Add(checkBox.Checked + "," + checkBox.Text + "," + checkBox.Font.Size + "," + checkBox.Font.FontFamily.Name);
                }
            }
            Properties.Settings.Default.CheckBox = string.Join(";", checkboxStates);
            Properties.Settings.Default.Save();
        }

        private void LoadCheckboxes()
        {
            flowLayoutPanel1.Controls.Clear();
            string savedStates = Properties.Settings.Default.CheckBox;
            if (!string.IsNullOrEmpty(savedStates))
            {
                string[] checkboxStates = savedStates.Split(';');
                foreach (string state in checkboxStates)
                {
                    string[] parts = state.Split(',');
                    if (parts.Length == 4)
                    {
                        System.Windows.Forms.CheckBox checkBox = new System.Windows.Forms.CheckBox();
                        checkBox.Checked = bool.Parse(parts[0]);
                        checkBox.Text = parts[1];
                        FontFamily fontFamily = new FontFamily(parts[3]);
                        checkBox.Font = new System.Drawing.Font(fontFamily, 12, checkBox.Checked ? FontStyle.Strikeout : FontStyle.Regular);
                        checkBox.AutoSize = true;
                        checkBox.CheckedChanged += new EventHandler(DynamicCheckBox_CheckedChanged);
                        flowLayoutPanel1.Controls.Add(checkBox);
                    }
                }
            }
        }

        private void Button36_Click(object sender, EventArgs e)
        {
            ClearCheckboxes();
        }

        private void ClearCheckboxes()
        {
            flowLayoutPanel1.Controls.Clear();
            Properties.Settings.Default.CheckBox = string.Empty;
            Properties.Settings.Default.Save();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Button35_Click(sender, e);

                e.SuppressKeyPress = true;
            }
        }
        //
        //CALENDAR
        //
        private void SubscribeCommonEventHandlers()
        {
            Button[] commonButtons = new Button[]
            {
                jan2, jan1, jan8, jan5, jan15, jan12, jan29, jan26, jan22, jan20,
                jan19, jan23, jan3, jan10, jan17, jan24, jan18, jan11, jan4,
                button9, button8, jan27, jan28, button11, button15, button44, button41,
                button19, button12, button39, button38, button33, button32, button40,
                button18, button10, button25, button24, button23, button22, button26,
                button17, button6, button30, button29, button28, button27, button31, button7
            };

            foreach (Button btn in commonButtons)
            {
                btn.Click += CommonButton_Click;
            }
        }

        private void SubscribeSecondCommonEventHandlers()
        {
            Button[] secondCommonButtons = new Button[]
            {
        button11, button14, button13, button16, button21, button5, jan20,
        jan21, button42, button43, button45
            };

            foreach (Button btn in secondCommonButtons)
            {
                btn.Click += SecondCommonButton_Click;
            }
        }

        private void CommonButton_Click(object sender, EventArgs e)
        {
            TogglePanels(true);
        }

        private void SecondCommonButton_Click(object sender, EventArgs e)
        {
            TogglePanels(false);
        }

        private void TogglePanels(bool isCommonButton)
        {
            label7.Show();
            button37.Show();
            panel10.Show();
            panel9.Hide();
            Toolbox.Hide();
            italicOn.Image = ResizeImage(Properties.Resources.italic_button, 33, 27);
            boldOn.Image = ResizeImage(Properties.Resources.bold_button, 33, 27);

            if (isCommonButton)
            {
                if (activeTextBox != null)
                {
                    activeTextBox.Hide();
                }
            }
            else
            {
                panel9.Show();
                label7.Hide();
                button37.Hide();
                italicOn.Image = ResizeImage(Properties.Resources.italic_button, 33, 27);
                boldOn.Image = ResizeImage(Properties.Resources.bold_button, 33, 27);
                Toolbox.Show();
                if (activeTextBox != null)
                {
                    activeTextBox.Show();
                }
            }
        }

        private void Button43_Click(object sender, EventArgs e)
        {
            HideTextBoxesExcept(activeTextBox);
            activeTextBox = feb3;
            LoadContent();
        }

        private void Button42_Click(object sender, EventArgs e)
        {
            HideTextBoxesExcept(activeTextBox);
            activeTextBox = feb4;
            LoadContent();
        }

        private void Button45_Click(object sender, EventArgs e)
        {
            HideTextBoxesExcept(activeTextBox);
            activeTextBox = feb6;
            LoadContent();
        }

        private void Jan21_Click(object sender, EventArgs e)
        {
            HideTextBoxesExcept(activeTextBox);
            activeTextBox = jan227;
            LoadContent();
        }

        private void Jan20_Click(object sender, EventArgs e)
        {
            HideTextBoxesExcept(activeTextBox);
            activeTextBox = jan9;
            LoadContent();
        }

        private void Button14_Click(object sender, EventArgs e)
        {
            HideTextBoxesExcept(activeTextBox);
            activeTextBox = jan13;
            LoadContent();
        }

        private void Button13_Click(object sender, EventArgs e)
        {
            HideTextBoxesExcept(activeTextBox);
            activeTextBox = jan14;
            LoadContent();
        }

        private void Button16_Click(object sender, EventArgs e)
        {
            HideTextBoxesExcept(activeTextBox);
            activeTextBox = jan16;
            LoadContent();
        }

        private void Button21_Click(object sender, EventArgs e)
        {
            HideTextBoxesExcept(activeTextBox);
            activeTextBox = jan30;
            LoadContent();
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            HideTextBoxesExcept(activeTextBox);
            activeTextBox = jan31;
            LoadContent();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            panel2.Visible = false;
            panel15.Visible = true;
            panel15.Location = new System.Drawing.Point(12, 115);
            panel9.Hide();
            panel10.Hide();
            Toolbox.Visible = false;

            if (activeTextBox != null)
            {
                activeTextBox.Hide();
            }
        }

        private void button46_Click(object sender, EventArgs e)
        {
            panel2.Visible = true;
            panel15.Visible = false;
            panel10.Hide();
            Toolbox.Visible= false;

            if (activeTextBox != null)
            {
                activeTextBox.Hide();
            }
        }

        private void HideTextBoxesExcept(System.Windows.Forms.RichTextBox activeTextBox)
        {
            foreach (System.Windows.Forms.RichTextBox textBox in textBoxes)
            {
                if (textBox != activeTextBox)
                {
                    textBox.Visible = false;
                }
            }
        }
        //
        // WORKDAY
        //
        private void Button37_Click(object sender, EventArgs e)
        {

            WorkdayText.Visible = true;
            WorkdayText.Location = new System.Drawing.Point(48, 193);
        }

        private void WorkDayBtn_Click(object sender, EventArgs e)
        {
            WorkdayText.Visible = false;
        }
        //
        //SIGNOUT
        //
        private void Button2_Click(object sender, EventArgs e)
        {
            LogIn form1 = new LogIn();
            form1.Show();
            this.Hide();
        }
        //
        //TOOLBOX
        //
        public System.Drawing.Image ResizeImage(System.Drawing.Image image, int width, int height)
        {
            Bitmap resizedImage = new Bitmap(image, new Size(width, height));
            return resizedImage;
        }

        private void RichTextBox_GotFocus(object sender, EventArgs e)
        {
            activeTextBox = sender as RichTextBox;
        }

        private void ItalicOn_Click(object sender, EventArgs e)
        {
            if (activeTextBox != null)
            {
                ApplyFontStyle(FontStyle.Italic);
                activeTextBox.Focus();
                if ((currentFontStyle & FontStyle.Italic) == FontStyle.Italic)
                {
                    italicOn.Image = ResizeImage(Properties.Resources.italic_on, 33, 27);
                }
                else
                {
                    italicOn.Image = ResizeImage(Properties.Resources.italic_button, 33, 27);
                }
            }
        }

        private void BoldOn_Click(object sender, EventArgs e)
        {
            if (activeTextBox != null)
            {
                ApplyFontStyle(FontStyle.Bold);
                activeTextBox.Focus();

                if ((currentFontStyle & FontStyle.Bold) == FontStyle.Bold)
                {
                    boldOn.Image = ResizeImage(Properties.Resources.bold_on, 33, 27);
                }
                else
                {
                    boldOn.Image = ResizeImage(Properties.Resources.bold_button, 33, 27);
                }
            }
        }

        public void ApplyFontStyle(FontStyle style)
        {
            if (activeTextBox != null)
            {
                currentFontStyle ^= style;
                activeTextBox.SelectionFont = new System.Drawing.Font(customFontFamily, activeTextBox.SelectionFont.Size, currentFontStyle);
            }
        }

        private void Clear_Click(object sender, EventArgs e)
        {
            if (activeTextBox != null)
            {
                ClearText.Visible = true;
                ClearText.Location = new System.Drawing.Point(48, 190);
                ClearText.BringToFront();
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (activeTextBox != null)
            {
                activeTextBox.Clear();
                activeTextBox.Focus();
            }

            ClearText.Visible = false;
        }

        private void button47_Click(object sender, EventArgs e)
        {
            ClearText.Visible = false;
        }
    }
}

