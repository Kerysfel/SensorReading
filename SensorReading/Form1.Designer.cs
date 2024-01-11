
namespace SensorReading
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panel1 = new System.Windows.Forms.Panel();
            this.OpenFormFull = new System.Windows.Forms.Button();
            this.ConnectRedMini = new System.Windows.Forms.Panel();
            this.ConnectYellowMini = new System.Windows.Forms.Panel();
            this.ConnectGrinMini = new System.Windows.Forms.Panel();
            this.SignedText = new System.Windows.Forms.Label();
            this.Signed = new System.Windows.Forms.PictureBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.ConnectComPort = new System.Windows.Forms.Button();
            this.StatusPanel = new System.Windows.Forms.Label();
            this.ConnectGrin = new System.Windows.Forms.Panel();
            this.ConnectYellow = new System.Windows.Forms.Panel();
            this.ConnectRed = new System.Windows.Forms.Panel();
            this.ConnectGreenText = new System.Windows.Forms.Label();
            this.ConnectYellowText = new System.Windows.Forms.Label();
            this.ConnectRedText = new System.Windows.Forms.Label();
            this.ComConnectorsList = new System.Windows.Forms.ListBox();
            this.StatusListing = new System.Windows.Forms.PictureBox();
            this.CloseComPort = new System.Windows.Forms.Button();
            this.SensorGridView = new System.Windows.Forms.DataGridView();
            this.DataPeriodBox = new System.Windows.Forms.ComboBox();
            this.StopReading = new System.Windows.Forms.Button();
            this.ContinueReading = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SaveDataToFile = new System.Windows.Forms.Button();
            this.ClearButton = new System.Windows.Forms.Button();
            this.selectTableBox = new System.Windows.Forms.ComboBox();
            this.TableType = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Signed)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.StatusListing)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.SensorGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.panel1.Controls.Add(this.OpenFormFull);
            this.panel1.Controls.Add(this.ConnectRedMini);
            this.panel1.Controls.Add(this.ConnectYellowMini);
            this.panel1.Controls.Add(this.ConnectGrinMini);
            this.panel1.Controls.Add(this.SignedText);
            this.panel1.Controls.Add(this.Signed);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Location = new System.Drawing.Point(-1, -1);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(801, 36);
            this.panel1.TabIndex = 0;
            this.panel1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseDown);
            // 
            // OpenFormFull
            // 
            this.OpenFormFull.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.OpenFormFull.FlatAppearance.BorderColor = System.Drawing.SystemColors.WindowFrame;
            this.OpenFormFull.FlatAppearance.BorderSize = 0;
            this.OpenFormFull.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.OpenFormFull.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.OpenFormFull.Location = new System.Drawing.Point(731, -1);
            this.OpenFormFull.Margin = new System.Windows.Forms.Padding(0);
            this.OpenFormFull.Name = "OpenFormFull";
            this.OpenFormFull.Size = new System.Drawing.Size(35, 37);
            this.OpenFormFull.TabIndex = 8;
            this.OpenFormFull.Text = "O";
            this.OpenFormFull.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.OpenFormFull.UseVisualStyleBackColor = false;
            this.OpenFormFull.Click += new System.EventHandler(this.OpenFormFull_Click);
            this.OpenFormFull.MouseEnter += new System.EventHandler(this.OpenFormFull_MouseEnter);
            this.OpenFormFull.MouseLeave += new System.EventHandler(this.OpenFormFull_MouseLeave);
            // 
            // ConnectRedMini
            // 
            this.ConnectRedMini.BackColor = System.Drawing.Color.Red;
            this.ConnectRedMini.Location = new System.Drawing.Point(316, 8);
            this.ConnectRedMini.Name = "ConnectRedMini";
            this.ConnectRedMini.Size = new System.Drawing.Size(20, 20);
            this.ConnectRedMini.TabIndex = 7;
            // 
            // ConnectYellowMini
            // 
            this.ConnectYellowMini.BackColor = System.Drawing.Color.Yellow;
            this.ConnectYellowMini.Location = new System.Drawing.Point(285, 8);
            this.ConnectYellowMini.Name = "ConnectYellowMini";
            this.ConnectYellowMini.Size = new System.Drawing.Size(20, 20);
            this.ConnectYellowMini.TabIndex = 6;
            // 
            // ConnectGrinMini
            // 
            this.ConnectGrinMini.BackColor = System.Drawing.Color.Lime;
            this.ConnectGrinMini.Location = new System.Drawing.Point(254, 8);
            this.ConnectGrinMini.Name = "ConnectGrinMini";
            this.ConnectGrinMini.Size = new System.Drawing.Size(20, 20);
            this.ConnectGrinMini.TabIndex = 5;
            // 
            // SignedText
            // 
            this.SignedText.AutoSize = true;
            this.SignedText.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.SignedText.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.SignedText.Location = new System.Drawing.Point(45, 8);
            this.SignedText.Name = "SignedText";
            this.SignedText.Size = new System.Drawing.Size(203, 20);
            this.SignedText.TabIndex = 3;
            this.SignedText.Text = "Считыватель COM-порта";
            // 
            // Signed
            // 
            this.Signed.Image = ((System.Drawing.Image)(resources.GetObject("Signed.Image")));
            this.Signed.Location = new System.Drawing.Point(3, 0);
            this.Signed.Name = "Signed";
            this.Signed.Size = new System.Drawing.Size(36, 38);
            this.Signed.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.Signed.TabIndex = 1;
            this.Signed.TabStop = false;
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.button2.FlatAppearance.BorderColor = System.Drawing.SystemColors.WindowFrame;
            this.button2.FlatAppearance.BorderSize = 0;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button2.Location = new System.Drawing.Point(696, 1);
            this.button2.Margin = new System.Windows.Forms.Padding(0);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(35, 35);
            this.button2.TabIndex = 2;
            this.button2.Text = "_";
            this.button2.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            this.button2.MouseEnter += new System.EventHandler(this.button2_MouseEnter);
            this.button2.MouseLeave += new System.EventHandler(this.button2_MouseLeave);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.button1.FlatAppearance.BorderColor = System.Drawing.SystemColors.WindowFrame;
            this.button1.FlatAppearance.BorderSize = 0;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.Location = new System.Drawing.Point(766, 0);
            this.button1.Margin = new System.Windows.Forms.Padding(0);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(35, 38);
            this.button1.TabIndex = 1;
            this.button1.Text = "X";
            this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.button1.MouseEnter += new System.EventHandler(this.button1_MouseEnter);
            this.button1.MouseLeave += new System.EventHandler(this.button1_MouseLeave);
            // 
            // ConnectComPort
            // 
            this.ConnectComPort.BackColor = System.Drawing.Color.LightSteelBlue;
            this.ConnectComPort.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ConnectComPort.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ConnectComPort.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ConnectComPort.ForeColor = System.Drawing.Color.Black;
            this.ConnectComPort.Location = new System.Drawing.Point(335, 139);
            this.ConnectComPort.Name = "ConnectComPort";
            this.ConnectComPort.Size = new System.Drawing.Size(119, 29);
            this.ConnectComPort.TabIndex = 1;
            this.ConnectComPort.Text = "Подключиться";
            this.ConnectComPort.UseVisualStyleBackColor = false;
            this.ConnectComPort.Click += new System.EventHandler(this.ConnectComPort_Click);
            // 
            // StatusPanel
            // 
            this.StatusPanel.AutoSize = true;
            this.StatusPanel.Font = new System.Drawing.Font("Microsoft Sans Serif", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.StatusPanel.Location = new System.Drawing.Point(52, 47);
            this.StatusPanel.Name = "StatusPanel";
            this.StatusPanel.Size = new System.Drawing.Size(218, 22);
            this.StatusPanel.TabIndex = 2;
            this.StatusPanel.Text = "Состояние подключения";
            // 
            // ConnectGrin
            // 
            this.ConnectGrin.BackColor = System.Drawing.Color.Lime;
            this.ConnectGrin.Location = new System.Drawing.Point(21, 81);
            this.ConnectGrin.Name = "ConnectGrin";
            this.ConnectGrin.Size = new System.Drawing.Size(25, 25);
            this.ConnectGrin.TabIndex = 4;
            // 
            // ConnectYellow
            // 
            this.ConnectYellow.BackColor = System.Drawing.Color.Gold;
            this.ConnectYellow.Location = new System.Drawing.Point(21, 112);
            this.ConnectYellow.Name = "ConnectYellow";
            this.ConnectYellow.Size = new System.Drawing.Size(25, 25);
            this.ConnectYellow.TabIndex = 5;
            // 
            // ConnectRed
            // 
            this.ConnectRed.BackColor = System.Drawing.Color.Red;
            this.ConnectRed.Location = new System.Drawing.Point(21, 143);
            this.ConnectRed.Name = "ConnectRed";
            this.ConnectRed.Size = new System.Drawing.Size(25, 25);
            this.ConnectRed.TabIndex = 6;
            // 
            // ConnectGreenText
            // 
            this.ConnectGreenText.AutoSize = true;
            this.ConnectGreenText.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ConnectGreenText.Location = new System.Drawing.Point(52, 81);
            this.ConnectGreenText.Name = "ConnectGreenText";
            this.ConnectGreenText.Size = new System.Drawing.Size(107, 20);
            this.ConnectGreenText.TabIndex = 7;
            this.ConnectGreenText.Text = "Подключено";
            // 
            // ConnectYellowText
            // 
            this.ConnectYellowText.AutoSize = true;
            this.ConnectYellowText.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ConnectYellowText.Location = new System.Drawing.Point(52, 112);
            this.ConnectYellowText.Name = "ConnectYellowText";
            this.ConnectYellowText.Size = new System.Drawing.Size(96, 20);
            this.ConnectYellowText.TabIndex = 8;
            this.ConnectYellowText.Text = "Отключено";
            // 
            // ConnectRedText
            // 
            this.ConnectRedText.AutoSize = true;
            this.ConnectRedText.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ConnectRedText.Location = new System.Drawing.Point(52, 143);
            this.ConnectRedText.Name = "ConnectRedText";
            this.ConnectRedText.Size = new System.Drawing.Size(69, 20);
            this.ConnectRedText.TabIndex = 9;
            this.ConnectRedText.Text = "Ошибка";
            // 
            // ComConnectorsList
            // 
            this.ComConnectorsList.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ComConnectorsList.FormattingEnabled = true;
            this.ComConnectorsList.ItemHeight = 16;
            this.ComConnectorsList.Location = new System.Drawing.Point(335, 67);
            this.ComConnectorsList.Name = "ComConnectorsList";
            this.ComConnectorsList.Size = new System.Drawing.Size(92, 68);
            this.ComConnectorsList.TabIndex = 10;
            // 
            // StatusListing
            // 
            this.StatusListing.Cursor = System.Windows.Forms.Cursors.Hand;
            this.StatusListing.Image = ((System.Drawing.Image)(resources.GetObject("StatusListing.Image")));
            this.StatusListing.Location = new System.Drawing.Point(16, 43);
            this.StatusListing.Name = "StatusListing";
            this.StatusListing.Size = new System.Drawing.Size(30, 30);
            this.StatusListing.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.StatusListing.TabIndex = 11;
            this.StatusListing.TabStop = false;
            this.StatusListing.Click += new System.EventHandler(this.StatusListing_Click);
            // 
            // CloseComPort
            // 
            this.CloseComPort.BackColor = System.Drawing.Color.LightSteelBlue;
            this.CloseComPort.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.CloseComPort.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.CloseComPort.Location = new System.Drawing.Point(460, 139);
            this.CloseComPort.Name = "CloseComPort";
            this.CloseComPort.Size = new System.Drawing.Size(119, 29);
            this.CloseComPort.TabIndex = 12;
            this.CloseComPort.Text = "Отключиться";
            this.CloseComPort.UseVisualStyleBackColor = false;
            this.CloseComPort.Click += new System.EventHandler(this.CloseComPort_Click);
            // 
            // SensorGridView
            // 
            this.SensorGridView.BackgroundColor = System.Drawing.SystemColors.Menu;
            this.SensorGridView.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.SensorGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.SensorGridView.ColumnHeadersVisible = false;
            this.SensorGridView.Location = new System.Drawing.Point(16, 223);
            this.SensorGridView.Name = "SensorGridView";
            this.SensorGridView.RowHeadersVisible = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Menu;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.LightSteelBlue;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.White;
            this.SensorGridView.RowsDefaultCellStyle = dataGridViewCellStyle1;
            this.SensorGridView.Size = new System.Drawing.Size(772, 317);
            this.SensorGridView.TabIndex = 13;
            this.SensorGridView.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.SensorGridView_CellPainting_1);
            this.SensorGridView.MouseClick += new System.Windows.Forms.MouseEventHandler(this.SensorGridView_MouseClick);
            // 
            // DataPeriodBox
            // 
            this.DataPeriodBox.FormattingEnabled = true;
            this.DataPeriodBox.Items.AddRange(new object[] {
            "0100",
            "0500",
            "1000",
            "3000"});
            this.DataPeriodBox.Location = new System.Drawing.Point(443, 67);
            this.DataPeriodBox.Name = "DataPeriodBox";
            this.DataPeriodBox.Size = new System.Drawing.Size(121, 21);
            this.DataPeriodBox.TabIndex = 15;
            // 
            // StopReading
            // 
            this.StopReading.BackColor = System.Drawing.Color.LightSteelBlue;
            this.StopReading.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.StopReading.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.StopReading.Location = new System.Drawing.Point(335, 174);
            this.StopReading.Name = "StopReading";
            this.StopReading.Size = new System.Drawing.Size(119, 29);
            this.StopReading.TabIndex = 16;
            this.StopReading.Text = "Стоп";
            this.StopReading.UseVisualStyleBackColor = false;
            this.StopReading.Click += new System.EventHandler(this.StopReading_Click);
            // 
            // ContinueReading
            // 
            this.ContinueReading.BackColor = System.Drawing.Color.LightSteelBlue;
            this.ContinueReading.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ContinueReading.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ContinueReading.Location = new System.Drawing.Point(460, 174);
            this.ContinueReading.Name = "ContinueReading";
            this.ContinueReading.Size = new System.Drawing.Size(119, 29);
            this.ContinueReading.TabIndex = 17;
            this.ContinueReading.Text = "Продолжить";
            this.ContinueReading.UseVisualStyleBackColor = false;
            this.ContinueReading.Click += new System.EventHandler(this.ContinueReading_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(335, 48);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 17);
            this.label1.TabIndex = 18;
            this.label1.Text = "COM-порты";
            this.label1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.label1_MouseDoubleClick);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(440, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(193, 17);
            this.label2.TabIndex = 19;
            this.label2.Text = "Период накопления данных";
            // 
            // SaveDataToFile
            // 
            this.SaveDataToFile.BackColor = System.Drawing.Color.LightSteelBlue;
            this.SaveDataToFile.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.SaveDataToFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.SaveDataToFile.Location = new System.Drawing.Point(585, 139);
            this.SaveDataToFile.Name = "SaveDataToFile";
            this.SaveDataToFile.Size = new System.Drawing.Size(119, 29);
            this.SaveDataToFile.TabIndex = 20;
            this.SaveDataToFile.Text = "Сохранить";
            this.SaveDataToFile.UseVisualStyleBackColor = false;
            this.SaveDataToFile.Click += new System.EventHandler(this.SaveDataToFile_Click);
            // 
            // ClearButton
            // 
            this.ClearButton.BackColor = System.Drawing.Color.LightSteelBlue;
            this.ClearButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ClearButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ClearButton.Location = new System.Drawing.Point(585, 174);
            this.ClearButton.Name = "ClearButton";
            this.ClearButton.Size = new System.Drawing.Size(119, 29);
            this.ClearButton.TabIndex = 22;
            this.ClearButton.Text = "Очистить";
            this.ClearButton.UseVisualStyleBackColor = false;
            this.ClearButton.Click += new System.EventHandler(this.ClearButton_Click);
            // 
            // selectTableBox
            // 
            this.selectTableBox.FormattingEnabled = true;
            this.selectTableBox.Items.AddRange(new object[] {
            "Все данные",
            "Магнитные азимуты",
            "Акселерометры",
            "Магнитометры",
            "Круги Эйлера",
            "Маджвик"});
            this.selectTableBox.Location = new System.Drawing.Point(443, 114);
            this.selectTableBox.Name = "selectTableBox";
            this.selectTableBox.Size = new System.Drawing.Size(121, 21);
            this.selectTableBox.TabIndex = 26;
            this.selectTableBox.SelectedIndexChanged += new System.EventHandler(this.selectTableBox_SelectedIndexChanged);
            // 
            // TableType
            // 
            this.TableType.AutoSize = true;
            this.TableType.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.TableType.Location = new System.Drawing.Point(440, 94);
            this.TableType.Name = "TableType";
            this.TableType.Size = new System.Drawing.Size(94, 17);
            this.TableType.TabIndex = 27;
            this.TableType.Text = "Тип таблицы";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(800, 552);
            this.Controls.Add(this.TableType);
            this.Controls.Add(this.selectTableBox);
            this.Controls.Add(this.ClearButton);
            this.Controls.Add(this.SaveDataToFile);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ContinueReading);
            this.Controls.Add(this.ConnectRed);
            this.Controls.Add(this.ConnectYellow);
            this.Controls.Add(this.ConnectGrin);
            this.Controls.Add(this.StopReading);
            this.Controls.Add(this.DataPeriodBox);
            this.Controls.Add(this.SensorGridView);
            this.Controls.Add(this.CloseComPort);
            this.Controls.Add(this.StatusListing);
            this.Controls.Add(this.ComConnectorsList);
            this.Controls.Add(this.ConnectRedText);
            this.Controls.Add(this.ConnectYellowText);
            this.Controls.Add(this.ConnectGreenText);
            this.Controls.Add(this.StatusPanel);
            this.Controls.Add(this.ConnectComPort);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SensorReader";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainForm_FormClosed);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Signed)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.StatusListing)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.SensorGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.PictureBox Signed;
        private System.Windows.Forms.Label SignedText;
        private System.Windows.Forms.Button ConnectComPort;
        private System.Windows.Forms.Label StatusPanel;
        private System.Windows.Forms.Panel ConnectGrin;
        private System.Windows.Forms.Panel ConnectYellow;
        private System.Windows.Forms.Panel ConnectRed;
        private System.Windows.Forms.Label ConnectGreenText;
        private System.Windows.Forms.Label ConnectYellowText;
        private System.Windows.Forms.Label ConnectRedText;
        private System.Windows.Forms.ListBox ComConnectorsList;
        private System.Windows.Forms.PictureBox StatusListing;
        private System.Windows.Forms.Button CloseComPort;
        private System.Windows.Forms.DataGridView SensorGridView;
        private System.Windows.Forms.ComboBox DataPeriodBox;
        private System.Windows.Forms.Button StopReading;
        private System.Windows.Forms.Button ContinueReading;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button SaveDataToFile;
        private System.Windows.Forms.Panel ConnectRedMini;
        private System.Windows.Forms.Panel ConnectYellowMini;
        private System.Windows.Forms.Panel ConnectGrinMini;
        private System.Windows.Forms.Button OpenFormFull;
        private System.Windows.Forms.Button ClearButton;
        private System.Windows.Forms.ComboBox selectTableBox;
        private System.Windows.Forms.Label TableType;
    }
}