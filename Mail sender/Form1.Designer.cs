namespace Mail_sender
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.excelFileSelectBtn = new System.Windows.Forms.Button();
            this.excelFileSelectLabel = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.htmlFileSelectBtn = new System.Windows.Forms.Button();
            this.htmlFileSelectLabel = new System.Windows.Forms.Label();
            this.loadDataFromFileBtn = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.filesPathLabel = new System.Windows.Forms.Label();
            this.folderWithFilesSelectBtn = new System.Windows.Forms.Button();
            this.openMultiFileForAttachmentDialog = new System.Windows.Forms.OpenFileDialog();
            this.attachFilesSelectBtn = new System.Windows.Forms.Button();
            this.attachFilesSelectLabel = new System.Windows.Forms.Label();
            this.subjectTextBox = new System.Windows.Forms.TextBox();
            this.subjectLabel = new System.Windows.Forms.Label();
            this.emailRowNumberBox = new System.Windows.Forms.TextBox();
            this.emailRowNumberLabel = new System.Windows.Forms.Label();
            this.fioRowNumberBox = new System.Windows.Forms.TextBox();
            this.fioRowNumberLabel = new System.Windows.Forms.Label();
            this.importantBox = new System.Windows.Forms.CheckBox();
            this.replacePtoDIV = new System.Windows.Forms.CheckBox();
            this.filesExtentionLabel = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.filesExtentionSelectBox = new System.Windows.Forms.ComboBox();
            this.progressLabel = new System.Windows.Forms.Label();
            this.wordToHtmlLink = new System.Windows.Forms.LinkLabel();
            this.addAppointment = new System.Windows.Forms.CheckBox();
            this.startlabel = new System.Windows.Forms.Label();
            this.start = new System.Windows.Forms.DateTimePicker();
            this.end = new System.Windows.Forms.DateTimePicker();
            this.endLabel = new System.Windows.Forms.Label();
            this.appointmentSubject = new System.Windows.Forms.TextBox();
            this.appointmentSubjectLabel = new System.Windows.Forms.Label();
            this.appointmentLocation = new System.Windows.Forms.TextBox();
            this.appointmentLocationLabel = new System.Windows.Forms.Label();
            this.appointmentBodyLabel = new System.Windows.Forms.Label();
            this.appointmentBody = new System.Windows.Forms.TextBox();
            this.appointmentPanel = new System.Windows.Forms.Panel();
            this.setReminder = new System.Windows.Forms.CheckBox();
            this.reminderTimeLabel = new System.Windows.Forms.Label();
            this.reminderMinutes = new System.Windows.Forms.TextBox();
            this.appointmentPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // excelFileSelectBtn
            // 
            this.excelFileSelectBtn.Location = new System.Drawing.Point(21, 22);
            this.excelFileSelectBtn.Name = "excelFileSelectBtn";
            this.excelFileSelectBtn.Size = new System.Drawing.Size(123, 23);
            this.excelFileSelectBtn.TabIndex = 12;
            this.excelFileSelectBtn.Text = "Выбери файл Excel";
            this.excelFileSelectBtn.UseVisualStyleBackColor = true;
            this.excelFileSelectBtn.Click += new System.EventHandler(this.excelFileSelectBtn_Click);
            // 
            // excelFileSelectLabel
            // 
            this.excelFileSelectLabel.AutoSize = true;
            this.excelFileSelectLabel.Location = new System.Drawing.Point(21, 48);
            this.excelFileSelectLabel.Name = "excelFileSelectLabel";
            this.excelFileSelectLabel.Size = new System.Drawing.Size(0, 13);
            this.excelFileSelectLabel.TabIndex = 11;
            // 
            // htmlFileSelectBtn
            // 
            this.htmlFileSelectBtn.Location = new System.Drawing.Point(21, 64);
            this.htmlFileSelectBtn.Name = "htmlFileSelectBtn";
            this.htmlFileSelectBtn.Size = new System.Drawing.Size(123, 23);
            this.htmlFileSelectBtn.TabIndex = 16;
            this.htmlFileSelectBtn.Text = "Выбери файл html";
            this.htmlFileSelectBtn.UseVisualStyleBackColor = true;
            this.htmlFileSelectBtn.Click += new System.EventHandler(this.htmlFileSelectBtn_Click);
            // 
            // htmlFileSelectLabel
            // 
            this.htmlFileSelectLabel.AutoSize = true;
            this.htmlFileSelectLabel.Location = new System.Drawing.Point(21, 90);
            this.htmlFileSelectLabel.Name = "htmlFileSelectLabel";
            this.htmlFileSelectLabel.Size = new System.Drawing.Size(0, 13);
            this.htmlFileSelectLabel.TabIndex = 15;
            // 
            // loadDataFromFileBtn
            // 
            this.loadDataFromFileBtn.Location = new System.Drawing.Point(656, 22);
            this.loadDataFromFileBtn.Name = "loadDataFromFileBtn";
            this.loadDataFromFileBtn.Size = new System.Drawing.Size(133, 23);
            this.loadDataFromFileBtn.TabIndex = 14;
            this.loadDataFromFileBtn.Text = "Загрузить и отправить";
            this.loadDataFromFileBtn.UseVisualStyleBackColor = true;
            this.loadDataFromFileBtn.Click += new System.EventHandler(this.loadDataFromFileBtn_Click);
            // 
            // filesPathLabel
            // 
            this.filesPathLabel.AutoSize = true;
            this.filesPathLabel.Location = new System.Drawing.Point(21, 132);
            this.filesPathLabel.Name = "filesPathLabel";
            this.filesPathLabel.Size = new System.Drawing.Size(0, 13);
            this.filesPathLabel.TabIndex = 17;
            // 
            // folderWithFilesSelectBtn
            // 
            this.folderWithFilesSelectBtn.Location = new System.Drawing.Point(21, 106);
            this.folderWithFilesSelectBtn.Name = "folderWithFilesSelectBtn";
            this.folderWithFilesSelectBtn.Size = new System.Drawing.Size(273, 23);
            this.folderWithFilesSelectBtn.TabIndex = 18;
            this.folderWithFilesSelectBtn.Text = "Выбери папку с файлами по именам для отправки";
            this.folderWithFilesSelectBtn.UseVisualStyleBackColor = true;
            this.folderWithFilesSelectBtn.Click += new System.EventHandler(this.folderWithFilesSelectBtn_Click);
            // 
            // openMultiFileForAttachmentDialog
            // 
            this.openMultiFileForAttachmentDialog.FileName = "openFileDialog2";
            this.openMultiFileForAttachmentDialog.Multiselect = true;
            // 
            // attachFilesSelectBtn
            // 
            this.attachFilesSelectBtn.Location = new System.Drawing.Point(21, 374);
            this.attachFilesSelectBtn.Name = "attachFilesSelectBtn";
            this.attachFilesSelectBtn.Size = new System.Drawing.Size(273, 23);
            this.attachFilesSelectBtn.TabIndex = 20;
            this.attachFilesSelectBtn.Text = "Выбери файлы для вложения во все письма ";
            this.attachFilesSelectBtn.UseVisualStyleBackColor = true;
            this.attachFilesSelectBtn.Click += new System.EventHandler(this.attachFilesSelectBtn_Click);
            // 
            // attachFilesSelectLabel
            // 
            this.attachFilesSelectLabel.AutoSize = true;
            this.attachFilesSelectLabel.Location = new System.Drawing.Point(21, 404);
            this.attachFilesSelectLabel.MaximumSize = new System.Drawing.Size(750, 500);
            this.attachFilesSelectLabel.Name = "attachFilesSelectLabel";
            this.attachFilesSelectLabel.Size = new System.Drawing.Size(0, 13);
            this.attachFilesSelectLabel.TabIndex = 19;
            // 
            // subjectTextBox
            // 
            this.subjectTextBox.Location = new System.Drawing.Point(21, 161);
            this.subjectTextBox.Name = "subjectTextBox";
            this.subjectTextBox.Size = new System.Drawing.Size(466, 20);
            this.subjectTextBox.TabIndex = 24;
            // 
            // subjectLabel
            // 
            this.subjectLabel.AutoSize = true;
            this.subjectLabel.Location = new System.Drawing.Point(21, 144);
            this.subjectLabel.Name = "subjectLabel";
            this.subjectLabel.Size = new System.Drawing.Size(75, 13);
            this.subjectLabel.TabIndex = 23;
            this.subjectLabel.Text = "Тема письма";
            // 
            // emailRowNumberBox
            // 
            this.emailRowNumberBox.Location = new System.Drawing.Point(389, 22);
            this.emailRowNumberBox.Name = "emailRowNumberBox";
            this.emailRowNumberBox.Size = new System.Drawing.Size(46, 20);
            this.emailRowNumberBox.TabIndex = 26;
            this.emailRowNumberBox.Text = "8";
            // 
            // emailRowNumberLabel
            // 
            this.emailRowNumberLabel.AutoSize = true;
            this.emailRowNumberLabel.Location = new System.Drawing.Point(389, 5);
            this.emailRowNumberLabel.Name = "emailRowNumberLabel";
            this.emailRowNumberLabel.Size = new System.Drawing.Size(195, 13);
            this.emailRowNumberLabel.TabIndex = 25;
            this.emailRowNumberLabel.Text = "Номер столбца с Email, где A-1, B-2...";
            // 
            // fioRowNumberBox
            // 
            this.fioRowNumberBox.Location = new System.Drawing.Point(177, 22);
            this.fioRowNumberBox.Name = "fioRowNumberBox";
            this.fioRowNumberBox.Size = new System.Drawing.Size(46, 20);
            this.fioRowNumberBox.TabIndex = 28;
            this.fioRowNumberBox.Text = "1";
            // 
            // fioRowNumberLabel
            // 
            this.fioRowNumberLabel.AutoSize = true;
            this.fioRowNumberLabel.Location = new System.Drawing.Point(177, 5);
            this.fioRowNumberLabel.Name = "fioRowNumberLabel";
            this.fioRowNumberLabel.Size = new System.Drawing.Size(197, 13);
            this.fioRowNumberLabel.TabIndex = 27;
            this.fioRowNumberLabel.Text = "Номер столбца с ФИО, где A-1, B-2...";
            // 
            // importantBox
            // 
            this.importantBox.AutoSize = true;
            this.importantBox.Location = new System.Drawing.Point(505, 164);
            this.importantBox.Name = "importantBox";
            this.importantBox.Size = new System.Drawing.Size(106, 17);
            this.importantBox.TabIndex = 30;
            this.importantBox.Text = "Важное письмо";
            this.importantBox.UseVisualStyleBackColor = true;
            // 
            // replacePtoDIV
            // 
            this.replacePtoDIV.AutoSize = true;
            this.replacePtoDIV.Checked = true;
            this.replacePtoDIV.CheckState = System.Windows.Forms.CheckState.Checked;
            this.replacePtoDIV.Location = new System.Drawing.Point(177, 68);
            this.replacePtoDIV.Name = "replacePtoDIV";
            this.replacePtoDIV.Size = new System.Drawing.Size(180, 17);
            this.replacePtoDIV.TabIndex = 31;
            this.replacePtoDIV.Text = "Убрать межстрочные отступы";
            this.replacePtoDIV.UseVisualStyleBackColor = true;
            // 
            // filesExtentionLabel
            // 
            this.filesExtentionLabel.AutoSize = true;
            this.filesExtentionLabel.Location = new System.Drawing.Point(303, 115);
            this.filesExtentionLabel.Name = "filesExtentionLabel";
            this.filesExtentionLabel.Size = new System.Drawing.Size(37, 13);
            this.filesExtentionLabel.TabIndex = 32;
            this.filesExtentionLabel.Text = "ФИО.";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(795, 22);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(100, 23);
            this.progressBar.TabIndex = 33;
            // 
            // filesExtentionSelectBox
            // 
            this.filesExtentionSelectBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.filesExtentionSelectBox.FormattingEnabled = true;
            this.filesExtentionSelectBox.Items.AddRange(new object[] {
            "doc",
            "docx",
            "pdf",
            "xls",
            "xlsx",
            "jpeg",
            "jpg",
            "png"});
            this.filesExtentionSelectBox.Location = new System.Drawing.Point(346, 108);
            this.filesExtentionSelectBox.Name = "filesExtentionSelectBox";
            this.filesExtentionSelectBox.Size = new System.Drawing.Size(47, 21);
            this.filesExtentionSelectBox.TabIndex = 34;
            this.filesExtentionSelectBox.SelectedIndexChanged += new System.EventHandler(this.filesExtentionSelectBox_SelectedIndexChanged);
            // 
            // progressLabel
            // 
            this.progressLabel.AutoSize = true;
            this.progressLabel.Location = new System.Drawing.Point(901, 26);
            this.progressLabel.Name = "progressLabel";
            this.progressLabel.Size = new System.Drawing.Size(0, 13);
            this.progressLabel.TabIndex = 35;
            // 
            // wordToHtmlLink
            // 
            this.wordToHtmlLink.AutoSize = true;
            this.wordToHtmlLink.Location = new System.Drawing.Point(21, 3);
            this.wordToHtmlLink.Name = "wordToHtmlLink";
            this.wordToHtmlLink.Size = new System.Drawing.Size(146, 13);
            this.wordToHtmlLink.TabIndex = 36;
            this.wordToHtmlLink.TabStop = true;
            this.wordToHtmlLink.Text = "Преобразовать Word в html";
            this.wordToHtmlLink.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.wordToHtmlLink_LinkClicked);
            // 
            // addAppointment
            // 
            this.addAppointment.AutoSize = true;
            this.addAppointment.Location = new System.Drawing.Point(3, 3);
            this.addAppointment.Name = "addAppointment";
            this.addAppointment.Size = new System.Drawing.Size(122, 17);
            this.addAppointment.TabIndex = 37;
            this.addAppointment.Text = "Назначить встречу";
            this.addAppointment.UseVisualStyleBackColor = true;
            // 
            // startlabel
            // 
            this.startlabel.AutoSize = true;
            this.startlabel.Location = new System.Drawing.Point(3, 27);
            this.startlabel.Name = "startlabel";
            this.startlabel.Size = new System.Drawing.Size(14, 13);
            this.startlabel.TabIndex = 38;
            this.startlabel.Text = "C";
            // 
            // start
            // 
            this.start.Location = new System.Drawing.Point(23, 21);
            this.start.Name = "start";
            this.start.Size = new System.Drawing.Size(200, 20);
            this.start.TabIndex = 39;
            // 
            // end
            // 
            this.end.Location = new System.Drawing.Point(265, 20);
            this.end.Name = "end";
            this.end.Size = new System.Drawing.Size(200, 20);
            this.end.TabIndex = 41;
            // 
            // endLabel
            // 
            this.endLabel.AutoSize = true;
            this.endLabel.Location = new System.Drawing.Point(238, 26);
            this.endLabel.Name = "endLabel";
            this.endLabel.Size = new System.Drawing.Size(21, 13);
            this.endLabel.TabIndex = 40;
            this.endLabel.Text = "По";
            // 
            // appointmentSubject
            // 
            this.appointmentSubject.Location = new System.Drawing.Point(483, 21);
            this.appointmentSubject.Name = "appointmentSubject";
            this.appointmentSubject.Size = new System.Drawing.Size(513, 20);
            this.appointmentSubject.TabIndex = 43;
            // 
            // appointmentSubjectLabel
            // 
            this.appointmentSubjectLabel.AutoSize = true;
            this.appointmentSubjectLabel.Location = new System.Drawing.Point(483, 4);
            this.appointmentSubjectLabel.Name = "appointmentSubjectLabel";
            this.appointmentSubjectLabel.Size = new System.Drawing.Size(77, 13);
            this.appointmentSubjectLabel.TabIndex = 42;
            this.appointmentSubjectLabel.Text = "Тема встречи";
            // 
            // appointmentLocation
            // 
            this.appointmentLocation.Location = new System.Drawing.Point(483, 67);
            this.appointmentLocation.Name = "appointmentLocation";
            this.appointmentLocation.Size = new System.Drawing.Size(513, 20);
            this.appointmentLocation.TabIndex = 45;
            // 
            // appointmentLocationLabel
            // 
            this.appointmentLocationLabel.AutoSize = true;
            this.appointmentLocationLabel.Location = new System.Drawing.Point(483, 50);
            this.appointmentLocationLabel.Name = "appointmentLocationLabel";
            this.appointmentLocationLabel.Size = new System.Drawing.Size(82, 13);
            this.appointmentLocationLabel.TabIndex = 44;
            this.appointmentLocationLabel.Text = "Место встречи";
            // 
            // appointmentBodyLabel
            // 
            this.appointmentBodyLabel.AutoSize = true;
            this.appointmentBodyLabel.Location = new System.Drawing.Point(3, 49);
            this.appointmentBodyLabel.Name = "appointmentBodyLabel";
            this.appointmentBodyLabel.Size = new System.Drawing.Size(80, 13);
            this.appointmentBodyLabel.TabIndex = 46;
            this.appointmentBodyLabel.Text = "Текст встречи";
            // 
            // appointmentBody
            // 
            this.appointmentBody.Location = new System.Drawing.Point(3, 65);
            this.appointmentBody.Multiline = true;
            this.appointmentBody.Name = "appointmentBody";
            this.appointmentBody.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.appointmentBody.Size = new System.Drawing.Size(462, 100);
            this.appointmentBody.TabIndex = 47;
            // 
            // appointmentPanel
            // 
            this.appointmentPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.appointmentPanel.Controls.Add(this.reminderMinutes);
            this.appointmentPanel.Controls.Add(this.reminderTimeLabel);
            this.appointmentPanel.Controls.Add(this.setReminder);
            this.appointmentPanel.Controls.Add(this.addAppointment);
            this.appointmentPanel.Controls.Add(this.appointmentBody);
            this.appointmentPanel.Controls.Add(this.startlabel);
            this.appointmentPanel.Controls.Add(this.appointmentBodyLabel);
            this.appointmentPanel.Controls.Add(this.start);
            this.appointmentPanel.Controls.Add(this.appointmentLocation);
            this.appointmentPanel.Controls.Add(this.endLabel);
            this.appointmentPanel.Controls.Add(this.appointmentLocationLabel);
            this.appointmentPanel.Controls.Add(this.end);
            this.appointmentPanel.Controls.Add(this.appointmentSubject);
            this.appointmentPanel.Controls.Add(this.appointmentSubjectLabel);
            this.appointmentPanel.Location = new System.Drawing.Point(21, 187);
            this.appointmentPanel.Name = "appointmentPanel";
            this.appointmentPanel.Size = new System.Drawing.Size(1001, 181);
            this.appointmentPanel.TabIndex = 48;
            // 
            // setReminder
            // 
            this.setReminder.AutoSize = true;
            this.setReminder.Location = new System.Drawing.Point(483, 102);
            this.setReminder.Name = "setReminder";
            this.setReminder.Size = new System.Drawing.Size(172, 17);
            this.setReminder.TabIndex = 48;
            this.setReminder.Text = "Установить напоминание за";
            this.setReminder.UseVisualStyleBackColor = true;
            // 
            // reminderTimeLabel
            // 
            this.reminderTimeLabel.AutoSize = true;
            this.reminderTimeLabel.Location = new System.Drawing.Point(713, 103);
            this.reminderTimeLabel.Name = "reminderTimeLabel";
            this.reminderTimeLabel.Size = new System.Drawing.Size(90, 13);
            this.reminderTimeLabel.TabIndex = 49;
            this.reminderTimeLabel.Text = "минут до начала";
            // 
            // reminderMinutes
            // 
            this.reminderMinutes.Location = new System.Drawing.Point(661, 100);
            this.reminderMinutes.Name = "reminderMinutes";
            this.reminderMinutes.Size = new System.Drawing.Size(46, 20);
            this.reminderMinutes.TabIndex = 50;
            this.reminderMinutes.Text = "30";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1046, 543);
            this.Controls.Add(this.appointmentPanel);
            this.Controls.Add(this.wordToHtmlLink);
            this.Controls.Add(this.progressLabel);
            this.Controls.Add(this.filesExtentionSelectBox);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.filesExtentionLabel);
            this.Controls.Add(this.replacePtoDIV);
            this.Controls.Add(this.importantBox);
            this.Controls.Add(this.fioRowNumberBox);
            this.Controls.Add(this.fioRowNumberLabel);
            this.Controls.Add(this.emailRowNumberBox);
            this.Controls.Add(this.emailRowNumberLabel);
            this.Controls.Add(this.subjectTextBox);
            this.Controls.Add(this.subjectLabel);
            this.Controls.Add(this.attachFilesSelectBtn);
            this.Controls.Add(this.attachFilesSelectLabel);
            this.Controls.Add(this.folderWithFilesSelectBtn);
            this.Controls.Add(this.filesPathLabel);
            this.Controls.Add(this.excelFileSelectBtn);
            this.Controls.Add(this.excelFileSelectLabel);
            this.Controls.Add(this.htmlFileSelectBtn);
            this.Controls.Add(this.htmlFileSelectLabel);
            this.Controls.Add(this.loadDataFromFileBtn);
            this.Name = "Form1";
            this.Text = "Mail sender";
            this.appointmentPanel.ResumeLayout(false);
            this.appointmentPanel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button excelFileSelectBtn;
        private System.Windows.Forms.Label excelFileSelectLabel;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button htmlFileSelectBtn;
        private System.Windows.Forms.Label htmlFileSelectLabel;
        private System.Windows.Forms.Button loadDataFromFileBtn;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label filesPathLabel;
        private System.Windows.Forms.Button folderWithFilesSelectBtn;
        private System.Windows.Forms.OpenFileDialog openMultiFileForAttachmentDialog;
        private System.Windows.Forms.Button attachFilesSelectBtn;
        private System.Windows.Forms.Label attachFilesSelectLabel;
        private System.Windows.Forms.TextBox subjectTextBox;
        private System.Windows.Forms.Label subjectLabel;
        private System.Windows.Forms.TextBox emailRowNumberBox;
        private System.Windows.Forms.Label emailRowNumberLabel;
        private System.Windows.Forms.TextBox fioRowNumberBox;
        private System.Windows.Forms.Label fioRowNumberLabel;
        private System.Windows.Forms.CheckBox importantBox;
        private System.Windows.Forms.CheckBox replacePtoDIV;
        private System.Windows.Forms.Label filesExtentionLabel;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.ComboBox filesExtentionSelectBox;
        private System.Windows.Forms.Label progressLabel;
        private System.Windows.Forms.LinkLabel wordToHtmlLink;
        private System.Windows.Forms.CheckBox addAppointment;
        private System.Windows.Forms.Label startlabel;
        private System.Windows.Forms.DateTimePicker start;
        private System.Windows.Forms.DateTimePicker end;
        private System.Windows.Forms.Label endLabel;
        private System.Windows.Forms.TextBox appointmentSubject;
        private System.Windows.Forms.Label appointmentSubjectLabel;
        private System.Windows.Forms.TextBox appointmentLocation;
        private System.Windows.Forms.Label appointmentLocationLabel;
        private System.Windows.Forms.Label appointmentBodyLabel;
        private System.Windows.Forms.TextBox appointmentBody;
        private System.Windows.Forms.Panel appointmentPanel;
        private System.Windows.Forms.TextBox reminderMinutes;
        private System.Windows.Forms.Label reminderTimeLabel;
        private System.Windows.Forms.CheckBox setReminder;
    }
}

