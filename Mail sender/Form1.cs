using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Excel;
using HtmlAgilityPack;

namespace Mail_sender
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            start.Format = DateTimePickerFormat.Custom;
            start.CustomFormat = "dd.MM.yyyy HH:mm:ss";
            end.Format = DateTimePickerFormat.Custom;
            end.CustomFormat = "dd.MM.yyyy HH:mm:ss";
        }

        private string excelFile;
        private string htmlFile;
        private string filesPath;
        private string filesToAttach;
        private string selectedFilesExtention;

        private string[] passportFio;
        private string[] passportEmail;
        private List<string> filesToAttachArray;
        private Microsoft.Office.Interop.Outlook.Application outlook;

        //Выбор файла Excel
        private void excelFileSelectBtn_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "xls files (*.xlsx;*.xls)|*.xlsx;*.xls";
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                excelFile = openFileDialog1.FileName;
                excelFileSelectLabel.Text = excelFile;
            }
        }

        //Выбор файла html с письмом
        private void htmlFileSelectBtn_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "html files (*.html)|*.html";
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                htmlFile = openFileDialog1.FileName;
                htmlFileSelectLabel.Text = htmlFile;
            }
        }

        //Выбор пути к файлам для вложения
        private void folderWithFilesSelectBtn_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                filesPath = folderBrowserDialog1.SelectedPath;
                filesPathLabel.Text = filesPath;
            }
        }

        //Выбор пути к файлам для вложения
        private void attachFilesSelectBtn_Click(object sender, EventArgs e)
        {
            filesToAttachArray = new List<string>();

            DialogResult result = openMultiFileForAttachmentDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                foreach (String file in openMultiFileForAttachmentDialog.FileNames)
                {
                    filesToAttach += file;
                    attachFilesSelectLabel.Text += file + ";\n";
                    filesToAttachArray.Add(file.ToString());
                }
            }
        }

        private void LockInterface(bool lockInt)
        {
            if (lockInt)
            {
                excelFileSelectBtn.Enabled = false;
                fioRowNumberBox.Enabled = false;
                emailRowNumberBox.Enabled = false;
                htmlFileSelectBtn.Enabled = false;
                replacePtoDIV.Enabled = false;
                folderWithFilesSelectBtn.Enabled = false;
                filesExtentionSelectBox.Enabled = false;
                subjectTextBox.Enabled = false;
                importantBox.Enabled = false;
                loadDataFromFileBtn.Enabled = false;
                addAppointment.Enabled = false;
                start.Enabled = false;
                end.Enabled = false;
                appointmentSubject.Enabled = false;
                appointmentLocation.Enabled = false;
                appointmentBody.Enabled = false;
                setReminder.Enabled = false;
                reminderMinutes.Enabled = false;
                attachFilesSelectBtn.Enabled = false;
            }
            else
            {
                excelFileSelectBtn.Enabled = true;
                fioRowNumberBox.Enabled = true;
                emailRowNumberBox.Enabled = true;
                htmlFileSelectBtn.Enabled = true;
                replacePtoDIV.Enabled = true;
                folderWithFilesSelectBtn.Enabled = true;
                filesExtentionSelectBox.Enabled = true;
                subjectTextBox.Enabled = true;
                importantBox.Enabled = true;
                loadDataFromFileBtn.Enabled = true;
                addAppointment.Enabled = true;
                start.Enabled = true;
                end.Enabled = true;
                appointmentSubject.Enabled = true;
                appointmentLocation.Enabled = true;
                appointmentBody.Enabled = true;
                setReminder.Enabled = true;
                reminderMinutes.Enabled = true;
                attachFilesSelectBtn.Enabled = true;
            }
        }

        //Загрузка и отправка
        private async void loadDataFromFileBtn_Click(object sender, EventArgs e)
        {
            LockInterface(true);
            try
            {
                //Проверяем заполнение полей
                string errorText = "";
                bool goNext = true;
                if (excelFile == null)
                {
                    goNext = false;
                    errorText += "Не выбран файл Excel\n";
                }
                if (htmlFile == null)
                {
                    goNext = false;
                    errorText += "Не выбран файл html\n";
                }
                if (filesPath != null)
                {
                    if (filesExtentionSelectBox.SelectedItem == null)
                    {
                        goNext = false;
                        errorText += "Не выбрано расширение файлов по именам для отправки\n";
                    }
                }
                
                if (subjectTextBox.Text == string.Empty)
                {
                    goNext = false;
                    errorText += "Не указана тема письма\n";
                }
                if (fioRowNumberBox.Text == string.Empty)
                {
                    goNext = false;
                    errorText += "Не указан номер столбца с ФИО\n";
                }
                if (emailRowNumberBox.Text == string.Empty)
                {
                    goNext = false;
                    errorText += "Не указан номер столбца с Email";
                }

                if (goNext)
                {
                    outlook = new Microsoft.Office.Interop.Outlook.Application();
                    //Создаём приложение.
                    Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                    //Открываем книгу.
                    Workbook ObjWorkBook = ObjExcel.Workbooks.Open(excelFile, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //Выбираем таблицу(лист).
                    Worksheet ObjWorkSheet;
                    ObjWorkSheet = (Worksheet) ObjWorkBook.Sheets[1];


                    //ФИО работника
                    Range passportFioColumn = ObjWorkSheet.UsedRange.Columns[Convert.ToInt32(fioRowNumberBox.Text)];
                    Array passportFioValues = (Array) passportFioColumn.Cells.Value;
                    passportFio = passportFioValues.OfType<object>().Select(o => o.ToString()).ToArray();
                    //Email
                    Range passportEmailColumn = ObjWorkSheet.UsedRange.Columns[Convert.ToInt32(emailRowNumberBox.Text)];
                    Array passportEmailValues = (Array) passportEmailColumn.Cells.Value;
                    passportEmail = passportEmailValues.OfType<object>().Select(o => o.ToString()).ToArray();

                    //Удаляем приложение (выходим из экселя) - ато будет висеть в процессах!
                    ObjWorkBook.Close();
                    ObjExcel.Quit();
                    killExcel();

                    for (int i = 1; i < passportFio.Length; i++)
                    {
                        //Открываем шаблон HTML в правильной кодировке
                        var doc = new HtmlAgilityPack.HtmlDocument();
                        StreamReader reader = new StreamReader(WebRequest.Create(htmlFile).GetResponse().GetResponseStream(), Encoding.UTF8);
                        doc.Load(reader);

                        //Если убрать межстрочные отступы
                        if (replacePtoDIV.Checked)
                        {
                            foreach (var er in doc.DocumentNode.Descendants("p"))
                            {
                                er.Name = "div";
                            }
                        }

                        //Ищем ноды с ФИО и меняем в них значения
                        HtmlNodeCollection fioValues = doc.DocumentNode.SelectNodes("//span[contains(@class, 'fio')]");
                        for (int j = 0; j < fioValues.Count; j++)
                        {
                            fioValues[j].InnerHtml = passportFio[i] + ",";
                        }
                        

                        //Сохраняем html в новом файле
                        FileStream sw = new FileStream(passportFio[i] + ".html", FileMode.Create);
                        doc.Save(sw, Encoding.UTF8);
                        sw.Dispose();

                        //Читаем новый HTML
                        string html = File.ReadAllText(passportFio[i] + ".html");

                        //Отправляем письмо
                        SendMail(html, i);
                        
                        //Назначаем встречу
                        if (addAppointment.Checked)
                        {
                            SendAppointment(passportEmail[i]);
                        }

                        //Прогресс
                        progressLabel.Text = i + "/" + (passportFio.Length-1);

                        progressBar.Maximum = passportFio.Length - 1;
                        progressBar.Step = 1;
                        

                        var progress = new Progress<int>(v =>
                        {
                            progressBar.Value = i;
                        });
                        await Task.Run(() => DoWork(progress));
                    }

                    LockInterface(false);
                    var result = MessageBox.Show("Готово", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    LockInterface(false);
                    var result = MessageBox.Show(errorText, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (System.Exception exception)
            {
                LockInterface(false);
                var result = MessageBox.Show(exception.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        private void SendMail(string template, int index)
        {
            try
            {
                Microsoft.Office.Interop.Outlook.MailItem mailItem = (Microsoft.Office.Interop.Outlook.MailItem)outlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                mailItem.Subject = subjectTextBox.Text;
                mailItem.To = passportEmail[index];
                mailItem.HTMLBody = template;
                mailItem.Importance = importantBox.Checked ? Microsoft.Office.Interop.Outlook.OlImportance.olImportanceHigh : Microsoft.Office.Interop.Outlook.OlImportance.olImportanceNormal;
                mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML;
                mailItem.Display(false);
                //Добавляем вложение уникальное
                if (filesPath != null && filesExtentionSelectBox.SelectedIndex != 0)
                {
                    string uniqueAttachmentFileName = filesPath + "\\" + passportFio[index] +"."+ selectedFilesExtention;
                    mailItem.Attachments.Add(@"" + uniqueAttachmentFileName, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                }
                //Добавляем общие вложения
                if (filesToAttachArray != null)
                {
                    foreach (string fileName in filesToAttachArray)
                    {
                        mailItem.Attachments.Add(@"" + fileName, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                    }
                }

                mailItem.Send();

                //Удаляем HTML временный файл
                string fn = passportFio[index] + ".html";
                File.Delete(@"" + fn);
            }
            catch (System.Exception e)
            {
                var result = MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        private void SendAppointment(string recipientEmail)
        {
            try
            {
                Microsoft.Office.Interop.Outlook.AppointmentItem appointment = null;
                Microsoft.Office.Interop.Outlook.Recipients recipients = null;
                Microsoft.Office.Interop.Outlook.Recipient recipient = null;
                try
                {
                    appointment = (Microsoft.Office.Interop.Outlook.AppointmentItem)outlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
                    appointment.Start = start.Value;
                    appointment.End = end.Value;
                    appointment.Subject = appointmentSubject.Text;
                    appointment.Location = appointmentLocation.Text;
                    appointment.Body = appointmentBody.Text;
                    appointment.BusyStatus = Microsoft.Office.Interop.Outlook.OlBusyStatus.olBusy;
                    if (setReminder.Checked)
                    {
                        appointment.ReminderSet = true;
                        appointment.ReminderMinutesBeforeStart = Convert.ToInt32(reminderMinutes.Text);
                    }
                    recipients = appointment.Recipients;
                    recipient = recipients.Add(recipientEmail);
                    recipient.Type = (int)Microsoft.Office.Interop.Outlook.OlMeetingRecipientType.olRequired;
                    //appointment.Save();
                    if (recipient.Resolve())
                    {
                        appointment.Send();
                    }
                    
                }
                finally
                {
                    if (recipient != null) Marshal.ReleaseComObject(recipient);
                    if (recipients != null) Marshal.ReleaseComObject(recipients);
                    if (appointment != null) Marshal.ReleaseComObject(appointment);
                }
            }
            catch (System.Exception e)
            {
                var result = MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        private void killExcel()
        {
            System.Diagnostics.Process[] PROC = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process PK in PROC)
            {
                if (PK.MainWindowTitle.Length == 0)
                {
                    PK.Kill();
                }
            }
        }

        //Выбор расширения для файлов
        private void filesExtentionSelectBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedFilesExtention = filesExtentionSelectBox.SelectedItem.ToString();
        }

        private void Calculate(int i)
        {
            double pow = Math.Pow(i, i);
        }

        //Рассчет прогресса
        public void DoWork(IProgress<int> progress)
        {
            for (int j = 0; j < 100000; j++)
            {
                Calculate(j);
                if (progress != null)
                {
                    progress.Report((j + 1) * 100 / 100000);
                }
            }
        }

        private void wordToHtmlLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://wordtohtml.net/");
        }
    }
}
