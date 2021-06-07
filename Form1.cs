using System;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using MailKit;
using MailKit.Net.Pop3;
using MailKit.Net.Imap;
using System.IO;
//using SautinSoft;
using RtfPipe;
using System.Drawing;
using System.Text;
using MimeKit;
using System.Collections.Generic;
using MarkupConverter;
using System.Linq;
using MahApps.Metro.Converters;
using System.Diagnostics;
using System.Configuration;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Data.SqlClient;
using System.Data;

namespace Lab2
{
    public partial class Form1 : Form
    {


        public Form1()
        {
            InitializeComponent();

            imapClient.Connect("imap.yandex.ru", 993, true);
            imapClient.Authenticate(textBox7.Text, textBox6.Text);
            comboBox2.Items.Clear();
            comboBox2.Items.AddRange(getImapFolders(imapClient).ToArray());
            comboBox2.SelectedIndex = 0;
        }
        string filename = "";
        Pop3Client popClient = new Pop3Client();                //Клиент POP3
        ImapClient imapClient = new ImapClient();               //Клиент IMAP
        RSACryptoServiceProvider rsa;                   // Объявление переменной асимметричного шифрования
        CspParameters cspp = new CspParameters();       // Создание экземпляра параметров асимметричного шифрования
        string publicKeyImported = "<RSAKeyValue><Modulus>21wEnTU+mcD2w0Lfo1Gv4rtcSWsQJQTNa6gio05AOkV/Er9w3Y13Ddo5wGtjJ19402S71HUeN0vbKILLJdRSES5MHSdJPSVrOqdrll/vLXxDxWs/U0UT1c8u6k/Ogx9hTtZxYwoeYqdhDblof3E75d9n2F0Zvf6iTb4cI7j6fMs=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";
        //string publicKeyImported = "";

        private void button1_Click(object sender, EventArgs e)
        {
            string[] recievers = textBox2.Text.Split(',');
            //string messageDir = messageDir = DateTime.Now.ToString().Replace(':', '_');
            //messageDir = "smtp\\" +
            //        comboBox1.Items[comboBox1.SelectedIndex]
            //        + "\\Sent\\" + messageDir;
            //if (!Directory.Exists(messageDir))
            //{
            //    Directory.CreateDirectory(messageDir);
            //}
            string encryptedKeyFile = "";
            var text = richTextBox1.Rtf;
            if (checkBox1.Checked)                                      // Шифрование тела письма если включено в настройках
            {
                AppSettingsReader settingsReader =
                                    new AppSettingsReader();
                string key = (string)settingsReader.GetValue("SecurityKey",
                                                 typeof(String));
                bool useHash = false;
                if (checkBox2.Checked)
                {
                    useHash = true;
                }
                text = TripleDES_Encrypt(richTextBox1.Rtf, useHash, key);
                string encryptedKey = RSAEncryption(key);   // Импорт открытого ключа асимметричного шифрования RSA
                encryptedKeyFile = "TripleDESEncryptedKey.txt"; // Шифрование ключа симметричного шифрования 
                File.WriteAllText(encryptedKeyFile, encryptedKey);      // 

            }
            string signatureFile = "";
            string pubKeyFile = "";
            if (checkBox3.Checked)                                      // Создание цифровой подписи методом Диффи-Хеллмана, 
            {                                                           // если включено в настройках
                var messageKeySignature = CngKey.Create(CngAlgorithm.ECDsaP256);
                var messagepubKey = Convert.ToBase64String(messageKeySignature.Export(CngKeyBlobFormat.GenericPublicBlob));
                string messageData = text;
                string messageSignature = CreateSignature(messageData, messageKeySignature);
                signatureFile = "signature.txt";
                File.WriteAllText(signatureFile, messageSignature);
                string pubKey = messagepubKey;
                pubKeyFile = "pubKey.txt";
                File.WriteAllText(pubKeyFile, pubKey);

            }
            string rtfString = text;
            MailAddress senderer = new MailAddress(textBox1.Text);
            foreach (var val in recievers)
            {
                MailAddress receiver = new MailAddress(val);
                MailMessage message = new MailMessage(senderer, receiver);
                message.IsBodyHtml = true;
                message.Subject = textBox5.Text;
                message.Body = text;
                if (filename != "")
                    message.Attachments.Add(new Attachment(filename));
                if (checkBox1.Checked)
                {
                    message.Attachments.Add(new Attachment(encryptedKeyFile));
                }
                if (checkBox3.Checked)
                {
                    message.Attachments.Add(new Attachment(signatureFile));
                    message.Attachments.Add(new Attachment(pubKeyFile));
                }
                SmtpClient smtp = new SmtpClient("smtp.yandex.ru", 587);
                smtp.Credentials = new NetworkCredential(textBox7.Text, textBox6.Text);
                smtp.EnableSsl = true;
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                try
                {
                    smtp.Send(message);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Возникла ошибка при отправке письма! Описание ошибки:\n {ex.Message}");
                }
            }
        }



        public string CreateSignature(string data, CngKey key)          // Метод создания цифровой подписи
        {
            var signingAlg = new ECDsaCng(key);
            var byteData = signingAlg.SignData(Convert.FromBase64String(data));
            string signature = Convert.ToBase64String(byteData);
            signingAlg.Clear();
            return signature;
        }

        public bool VerifySignature(string data, string signature, string pubKey)       // Метод проверки цифровой подписи
        {
            byte[] byteData = Convert.FromBase64String(data);
            byte[] signatureArr = Convert.FromBase64String(signature);
            byte[] pubKeyArr = Convert.FromBase64String(pubKey);
            bool retValue = false;
            using (CngKey key = CngKey.Import(pubKeyArr, CngKeyBlobFormat.GenericPublicBlob))
            {
                var signingAlg = new ECDsaCng(key);
                retValue = signingAlg.VerifyData(byteData, signatureArr);
                signingAlg.Clear();
            }
            return retValue;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            filename = openFileDialog1.FileName;
            label3.Text = filename;
            openFileDialog1.Dispose();
            openFileDialog1.Reset();
        }


        private void button3_Click(object sender, EventArgs ev)
        {
            try
            {
                // Синхронизация сообщений сервера IMAP с ПК
                if (radioButton1.Checked)
                {
                    listBox1.Text = "";
                    listBox1.Items.Clear();
                    //comboBox2.Items.Clear();
                    RecievedMessages.Clear();
                    label13.Text = "Всего сообщений: ";
                    imapClient.Disconnect(true);                            // Переподключение IMAP клиента
                    imapClient.Connect("imap.yandex.ru", 993, true);
                    imapClient.Authenticate(textBox7.Text, textBox6.Text);

                    var inbox = imapClient.GetFolder(                       // Получение выбранной папки почты
                        comboBox2.Text);
                    inbox.Open(FolderAccess.ReadOnly);                      // Открытие папки только для чтения
                    label13.Text = "Всего сообщений:" + inbox.Count + "\r\n";
                    imapClient.Timeout = 999999999;
                    listBox1.Items.Clear();
                    for (int i = inbox.Count - 1; i >= 0; i--)
                    {
                        var message = inbox.GetMessage(i);                  // Получение сообщения по индексу
                        string messageFrom = message.From.ToString().Replace("<", "").Replace(">", "").Replace("\"", "");

                        Console.WriteLine("Письмо: {0}", messageFrom);
                        string msg = "\r\n" + "" + message.Date.DateTime + " От: " + messageFrom + " - ";
                        listBox1.Items.Add(msg + "\r\n" + message.Subject + "\r\n\r\n");
                        string messageDir = message.Date.DateTime.ToString().Replace(":", "_") + " " + messageFrom;
                        //string temp = messageDir;
                        messageDir = "imap\\" + comboBox1.Text + "\\" + comboBox2.Items[comboBox2.SelectedIndex].ToString() + "\\" + messageDir;
                        if (!Directory.Exists(messageDir))
                        {
                            Directory.CreateDirectory(messageDir);
                        }

                        string htmlFile = messageDir + "\\" + "messageBody.html";

                        File.WriteAllText(htmlFile, message.HtmlBody);      // Сохранение тела письма в формате html

                        //Скачивание вложений IMAP в каталог отправителя в корневом каталоге программы
                        foreach (MimeEntity attachment in message.Attachments)
                        {
                            var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;

                            using (var stream = File.Create(messageDir + "//" + fileName))
                            {
                                if (attachment is MessagePart)
                                {
                                    var rfc822 = (MessagePart)attachment;

                                    rfc822.Message.WriteTo(stream);
                                }
                                else
                                {
                                    var part = (MimePart)attachment;

                                    part.ContentObject.DecodeTo(stream);
                                }
                            }
                        }
                    }
                }
                else
                {
                    // Синхронизация сообщений сервера POP3 с ПК
                    listBox1.Text = "";
                    listBox1.Items.Clear();
                    //comboBox2.Items.Clear();
                    RecievedMessages.Clear();
                    label13.Text = "Всего сообщений: ";

                    popClient.Disconnect(true);                             // Переподключение POP3 клиента
                    popClient.Connect("pop.yandex.ru", 995, true);
                    popClient.Authenticate(textBox7.Text, textBox6.Text);

                    listBox1.Text = "Всего сообщений:" + popClient.Count + "\r\n";
                    for (int i = 0; i < popClient.Count; i++)
                    {
                        var message = popClient.GetMessage(i);              // Получение сообщения по индексу
                        Console.WriteLine("Письмо: {0}", message.Subject);
                        string messageFrom = message.From.ToString().Replace("<", "").Replace(">", "").Replace("\"", "");

                        string msg = "\r\n" + "" + message.Date.DateTime + " От: " + messageFrom + " - ";
                        listBox1.Items.Add(msg + "\r\n" + message.Subject + "\r\n\r\n");

                        // Создание каталога сообщения
                        string messageDir = message.Date.DateTime.ToString().Replace(":", "_") + " " + messageFrom;
                        messageDir = "pop3\\" + comboBox1.Items[comboBox1.SelectedIndex] + "\\INBOX\\" + messageDir;
                        if (!Directory.Exists(messageDir))
                        {
                            Directory.CreateDirectory(messageDir);
                        }

                        string htmlFile = messageDir + "\\" + "messageBody.html";

                        File.WriteAllText(htmlFile, message.HtmlBody);       // Сохранение тела письма в формате html

                        //Скачивание вложений POP3 в каталог отправителя в корневом каталоге программы
                        foreach (MimeEntity attachment in message.Attachments)
                        {
                            var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;

                            using (var stream = File.Create(messageDir + "//" + fileName))
                            {
                                if (attachment is MessagePart)
                                {
                                    var rfc822 = (MessagePart)attachment;

                                    rfc822.Message.WriteTo(stream);
                                }
                                else
                                {
                                    var part = (MimePart)attachment;

                                    part.ContentObject.DecodeTo(stream);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                listBox1.Text = "";
            }

        }

        private void button4_Click(object sender, EventArgs ev)
        {

        }


        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            richTextBox1.Undo();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            richTextBox1.Redo();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            if (((ToolStripButton)toolStrip.Items[2]).Checked)
            {
                ((ToolStripButton)toolStrip.Items[2]).Checked = false;
            }
            else
            {
                ((ToolStripButton)toolStrip.Items[2]).Checked = true;
            }
            SetSelectionFont();
        }

        bool locker;
        private void MailMessage_SelectionChanged(object sender, EventArgs e)   // Метод обработки форматирования
        {
            if (richTextBox1.SelectionFont == null)
                return;
            else
            {
                locker = true;

                ((ToolStripComboBox)toolStrip.Items[0]).Text = richTextBox1.SelectionFont.Name;
                ((ToolStripComboBox)toolStrip.Items[1]).Text = richTextBox1.SelectionFont.Size.ToString();
                ((ToolStripButton)toolStrip.Items[2]).Checked = richTextBox1.SelectionFont.Bold;
                ((ToolStripButton)toolStrip.Items[3]).Checked = richTextBox1.SelectionFont.Italic;
                ((ToolStripButton)toolStrip.Items[4]).Checked = richTextBox1.SelectionFont.Underline;

                locker = false;
            }
        }

        private void SetSelectionFont()                             // Метод обновления панели форматирования
        {
            if (locker == true)
                return;
            else
            {
                FontStyle style = FontStyle.Regular;
                if (((ToolStripButton)toolStrip.Items[2]).Checked)
                {
                    style |= FontStyle.Bold;
                }
                if (((ToolStripButton)toolStrip.Items[3]).Checked)
                {
                    style |= FontStyle.Italic;
                }
                if (((ToolStripButton)toolStrip.Items[4]).Checked)
                {
                    style |= FontStyle.Underline;
                }

                richTextBox1.SelectionFont = new System.Drawing.Font(((ToolStripComboBox)toolStrip.Items[7]).Text, Convert.ToSingle(((ToolStripComboBox)toolStrip.Items[8]).Text), style);
                richTextBox1.Focus();
            }
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetSelectionFont();
        }

        private void toolStripComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetSelectionFont();
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            if (((ToolStripButton)this.toolStrip.Items[3]).Checked)
            {
                ((ToolStripButton)this.toolStrip.Items[3]).Checked = false;
            }
            else
            {
                ((ToolStripButton)this.toolStrip.Items[3]).Checked = true;
            }

            SetSelectionFont();

        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            if (((ToolStripButton)this.toolStrip.Items[4]).Checked)
            {
                ((ToolStripButton)this.toolStrip.Items[4]).Checked = false;
            }
            else
            {
                ((ToolStripButton)this.toolStrip.Items[4]).Checked = true;
            }

            SetSelectionFont();
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            ColorDialog dialog = new ColorDialog();
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                richTextBox1.SelectionColor = dialog.Color;
            }

        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            ColorDialog dialog = new ColorDialog();
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                this.richTextBox1.SelectionBackColor = dialog.Color;
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                attachedFile.Text = "";
                if (radioButton1.Checked)                                        // Для IMAP
                {

                    RecievedMessages.Rtf = "";
                    imapClient.Disconnect(true);
                    imapClient.Connect("imap.yandex.ru", 993, true);
                    imapClient.Authenticate(textBox7.Text, textBox6.Text);
                    var inbox = imapClient.GetFolder(comboBox2.Text);
                    inbox.Open(FolderAccess.ReadOnly);


                    var message = inbox.GetMessage(listBox1.Items.Count - 1 - listBox1.SelectedIndex);
                    //var msg = HtmlToRtfConverter.ConvertHtmlToRtf(message.HtmlBody);
                    string msg = message.HtmlBody;
                    string key = "";

                    string messageFrom = message.From.ToString().Replace("<", "").Replace(">", "").Replace("\"", "");
                    string messageDir = message.Date.DateTime.ToString().Replace(":", "_") + " " + messageFrom;
                    messageDir = "imap\\" + comboBox1.Items[comboBox1.SelectedIndex] + "\\" + comboBox2.Items[comboBox2.SelectedIndex].ToString() + "\\" + messageDir.Replace("<", "").Replace(">", "");
                    if (checkBox1.Checked)      // Дешифровка, если присутствует шифрование
                    {
                        //var keyFile = message.Attachments.Last();
                        //var fileName = keyFile.ContentDisposition?.FileName ?? keyFile.ContentType.Name;
                        string file = OpenKey();
                        //Получаем значение ключа
                        string sipher = File.ReadAllText(file);
                        key = RSADecryption(sipher);

                        bool useHash = false;
                        if (checkBox2.Checked)
                        {
                            useHash = true;
                        }
                        msg = TripleDES_Decrypt(message.HtmlBody, useHash, key);
                    }
                    if (checkBox3.Checked)      // Проверка цифровой подписи, если включено в настройках
                    {
                        string pubKeyFile = messageDir + "\\pubKey.txt";
                        string signatureFile = messageDir + "\\signature.txt";
                        string messageData = RecievedMessages.Text;
                        string pubKeyStr = File.ReadAllText(pubKeyFile);
                        string signatureStr = File.ReadAllText(signatureFile);
                        var messageKeySignature = CngKey.Create(CngAlgorithm.ECDsaP256);
                        string messagepubKey = Convert.ToBase64String(messageKeySignature.Export(CngKeyBlobFormat.GenericPublicBlob));

                        if (VerifySignature(message.HtmlBody, signatureStr, pubKeyStr))
                            MessageBox.Show("Проверка цифровой подписи пройдена.");
                        else MessageBox.Show("Цифровая подпись не верна.");
                    }


                    attachedFile.Text = "";
                    var data = message.Attachments.First();
                    if (data != null)
                    {
                        var fileName = data.ContentDisposition?.FileName ?? data.ContentType.Name;
                        //Добавление вложения в listbox
                        attachedFile.Text = fileName;
                    }
                    RecievedMessages.Rtf = msg;


                }
                else                                    // Для POP3
                {
                    popClient.Disconnect(true);
                    popClient.Connect("pop.yandex.ru", 995, true);
                    popClient.Authenticate(textBox7.Text, textBox6.Text);
                    var message = popClient.GetMessage(listBox1.SelectedIndex);
                    //var msg = HtmlToRtfConverter.ConvertHtmlToRtf(message.HtmlBody);
                    string msg = message.HtmlBody;
                    string key = "";
                 

                    attachedFile.Text = "";

                    string messageFrom = message.From.ToString().Replace("<", "").Replace(">", "").Replace("\"", "");
                    string messageDir = message.Date.DateTime.ToString().Replace(":", "_") + " " + messageFrom;
                    messageDir = "pop3\\" + comboBox1.Items[comboBox1.SelectedIndex] + "\\INBOX" + "\\" + messageDir.Replace("<", "").Replace(">", "");

                    if (checkBox1.Checked)      // Дешифровка, если присутствует шифрование
                    {
                        //var keyFile = message.Attachments.Last();
                        //var fileName = keyFile.ContentDisposition?.FileName ?? keyFile.ContentType.Name;
                        string file = OpenKey();
                        //Получаем значение ключа
                        string sipher = File.ReadAllText(file);
                        key = RSADecryption(sipher);

                        bool useHash = false;
                        if (checkBox2.Checked)
                        {
                            useHash = true;
                        }
                        msg = TripleDES_Decrypt(message.HtmlBody, useHash, key);
                    }
                    if (checkBox3.Checked)      // Проверка цифровой подписи, если включено в настройках
                    {
                        string pubKeyFile = messageDir + "\\pubKey.txt";
                        string signatureFile = messageDir + "\\signature.txt";
                        string messageData = RecievedMessages.Text;
                        string pubKeyStr = File.ReadAllText(pubKeyFile);
                        string signatureStr = File.ReadAllText(signatureFile);
                        var messageKeySignature = CngKey.Create(CngAlgorithm.ECDsaP256);
                        string messagepubKey = Convert.ToBase64String(messageKeySignature.Export(CngKeyBlobFormat.GenericPublicBlob));

                        if (VerifySignature(messageData, signatureStr, pubKeyStr))
                            MessageBox.Show("Цифровая подпись верна!");
                    }

                    attachedFile.Text = "";
                    var data = message.Attachments.First();
                    if (data != null)
                    {
                        var fileName = data.ContentDisposition?.FileName ?? data.ContentType.Name;
                        //Добавление вложения в listbox
                        attachedFile.Text = fileName;
                    }
                    RecievedMessages.Rtf = msg;

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        List<string> getImapFolders(ImapClient client)                  // Метод получения папок почты
        {
            IList<IMailFolder> folders = client.GetFolders(client.PersonalNamespaces.First());
            List<string> names = folders.Select(t => t.Name).ToList();
            return names;
        }

        private string OpenKey()
        {
            try
            {
                attachedFile.Text = "";
                if (radioButton1.Checked)                                // Для IMAP
                {
                    var inbox = imapClient.GetFolder(comboBox2.Items[comboBox2.SelectedIndex].ToString());

                    var message = inbox.GetMessage(listBox1.Items.Count - 1 - listBox1.SelectedIndex);

                    var attachment = message.Attachments.ToArray();

                    var fileName = "TripleDESEncryptedKey.txt";
                    //var fileName = attachment[listBox2.SelectedIndex].ContentDisposition?.FileName ?? attachment[listBox2.SelectedIndex].ContentType.Name;

                    string messageFrom = message.From.ToString();

                    string messageDir = message.Date.DateTime.ToString().Replace(":","_") + " " + messageFrom;
                    //string temp = messageDir;
                    messageDir = "imap\\" + comboBox1.Items[comboBox1.SelectedIndex] + "\\" + comboBox2.Items[comboBox2.SelectedIndex].ToString() + "\\" + messageDir;
                    if (!Directory.Exists(messageDir))
                    {
                        Directory.CreateDirectory(messageDir);
                    }

                    //using (var stream = File.Create(messageDir + @"\" + fileName))
                    //{
                    //    if (attachment is MessagePart)
                    //    {
                    //        var rfc822 = (MessagePart)attachment[1];

                    //        rfc822.Message.WriteTo(stream);
                    //    }
                    //    else
                    //    {
                    //        var part = (MimePart)attachment[1];

                    //        part.ContentObject.DecodeTo(stream);
                    //    }
                    //}
                    return messageDir + @"\" + fileName;
                }
                else                               // Для POP3
                {
                    var message = popClient.GetMessage(listBox1.SelectedIndex);

                    var attachment = message.Attachments.ToArray();

                    var fileName = "TripleDESEncryptedKey.txt";
                    //var fileName = attachment[listBox2.SelectedIndex].ContentDisposition?.FileName ?? attachment[listBox2.SelectedIndex].ContentType.Name;

                    string messageFrom = message.From.ToString();

                    string messageDir = message.Date.DateTime.ToString().Replace(":", "_") + " " + messageFrom;
                    //string temp = messageDir;
                    messageDir = "pop3\\" + comboBox1.Items[comboBox1.SelectedIndex] + "\\INBOX\\" + messageDir;
                    if (!Directory.Exists(messageDir))
                    {
                        Directory.CreateDirectory(messageDir);
                    }

                    //using (var stream = File.Create(messageDir + @"\" + fileName))
                    //{
                    //    if (attachment is MessagePart)
                    //    {
                    //        var rfc822 = (MessagePart)attachment[1];

                    //        rfc822.Message.WriteTo(stream);
                    //    }
                    //    else
                    //    {
                    //        var part = (MimePart)attachment[1];

                    //        part.ContentObject.DecodeTo(stream);
                    //    }
                    //}
                    return messageDir + @"\" + fileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return "";
            }

        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            try
            {
                //attachedFile.Text = "";
                if (radioButton1.Checked)                                // Для IMAP
                {
                    var inbox = imapClient.GetFolder(comboBox2.Items[comboBox2.SelectedIndex].ToString());

                    var message = inbox.GetMessage(listBox1.Items.Count - 1 - listBox1.SelectedIndex);

                    var attachment = message.Attachments.ToArray();

                    var fileName = attachedFile.Text;
                    //var fileName = attachment[listBox2.SelectedIndex].ContentDisposition?.FileName ?? attachment[listBox2.SelectedIndex].ContentType.Name;

                    string messageFrom = message.From.ToString();

                    string messageDir = message.Date.DateTime.ToString().Replace(":","_") + " " + messageFrom;
                    //string temp = messageDir;
                    messageDir = "imap\\" + comboBox1.Items[comboBox1.SelectedIndex] + "\\" + comboBox2.Items[comboBox2.SelectedIndex].ToString() + "\\" + messageDir;
                    if (!Directory.Exists(messageDir))
                    {
                        Directory.CreateDirectory(messageDir);
                    }

                    using (var stream = File.Create(messageDir + @"\" + fileName))
                    {
                        if (attachment is MessagePart)
                        {
                            var rfc822 = (MessagePart)attachment[0];

                            rfc822.Message.WriteTo(stream);
                        }
                        else
                        {
                            var part = (MimePart)attachment[0];

                            part.ContentObject.DecodeTo(stream);
                        }
                    }
                    Process.Start(messageDir + "\\" + fileName);
                }
                else                               // Для POP3
                {
                    var message = popClient.GetMessage(listBox1.SelectedIndex);

                    var attachment = message.Attachments.ToArray();

                    var fileName = attachedFile.Text;
                    //var fileName = attachment[listBox2.SelectedIndex].ContentDisposition?.FileName ?? attachment[listBox2.SelectedIndex].ContentType.Name;

                    string messageFrom = message.From.ToString();

                    string messageDir = message.Date.DateTime.ToString().Replace(":","_") + " " + messageFrom;
                    //string temp = messageDir;
                    messageDir = "pop3\\" + comboBox1.Items[comboBox1.SelectedIndex] + "\\INBOX\\" + messageDir;
                    if (!Directory.Exists(messageDir))
                    {
                        Directory.CreateDirectory(messageDir);
                    }

                    using (var stream = File.Create(messageDir + "\\" + fileName))
                    {
                        if (attachment is MessagePart)
                        {
                            var rfc822 = (MessagePart)attachment[0];

                            rfc822.Message.WriteTo(stream);
                        }
                        else
                        {
                            var part = (MimePart)attachment[0];

                            part.ContentObject.DecodeTo(stream);
                        }
                    }
                    Process.Start(messageDir + @"\" + fileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        //3-DES шифрование
        public static string TripleDES_Encrypt(string toEncrypt, bool useHashing, string key)
        {
            byte[] keyArray;
            byte[] toEncryptArray = UTF8Encoding.UTF8.GetBytes(toEncrypt);

            AppSettingsReader settingsReader =
                                                new AppSettingsReader();
            // Get the key from config file

            //string key = (string)settingsReader.GetValue("SecurityKey",
            //                                                 typeof(String));
            //System.Windows.Forms.MessageBox.Show(key);
            //If hashing use get hashcode regards to your key
            if (useHashing)
            {
                //SHA1CryptoServiceProvider hashSHA1 = new SHA1CryptoServiceProvider();
                //keyArray = hashSHA1.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
                //SHA256CryptoServiceProvider hashSHA2 = new SHA256CryptoServiceProvider();
                //keyArray = hashSHA2.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
                //Always release the resources and flush data
                // of the Cryptographic service provide. Best Practice
                hashmd5.Clear();
            }
            else
                keyArray = UTF8Encoding.UTF8.GetBytes(key);

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            //set the secret key for the tripleDES algorithm
            tdes.Key = keyArray;
            //mode of operation. there are other 4 modes.
            //We choose ECB(Electronic code Book)
            tdes.Mode = CipherMode.ECB;
            //padding mode(if any extra byte added)

            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateEncryptor();
            //transform the specified region of bytes array to resultArray
            byte[] resultArray =
              cTransform.TransformFinalBlock(toEncryptArray, 0,
              toEncryptArray.Length);
            //Release resources held by TripleDes Encryptor
            tdes.Clear();
            //Return the encrypted data into unreadable string format
            return Convert.ToBase64String(resultArray, 0, resultArray.Length);
        }

        // 3-DES расшифрование
        public static string TripleDES_Decrypt(string cipherString, bool useHashing,string key)
        {
            byte[] keyArray;
            //get the byte code of the string

            byte[] toEncryptArray = Convert.FromBase64String(cipherString);

            //AppSettingsReader settingsReader =
            //                                    new AppSettingsReader();
            ////Get your key from config file to open the lock!
            //string key = (string)settingsReader.GetValue("SecurityKey",
            //                                             typeof(String));

            if (useHashing)
            {
                //if hashing was used get the hash code with regards to your key
                //SHA256CryptoServiceProvider hashSHA2 = new SHA256CryptoServiceProvider();
                //keyArray = hashSHA2.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
                //release any resource held by the MD5CryptoServiceProvider

                hashmd5.Clear();
            }
            else
            {
                //if hashing was not implemented get the byte code of the key
                keyArray = UTF8Encoding.UTF8.GetBytes(key);
            }

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            //set the secret key for the tripleDES algorithm
            tdes.Key = keyArray;
            //mode of operation. there are other 4 modes. 
            //We choose ECB(Electronic code Book)

            tdes.Mode = CipherMode.ECB;
            //padding mode(if any extra byte added)
            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateDecryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(
                                 toEncryptArray, 0, toEncryptArray.Length);
            //Release resources held by TripleDes Encryptor                
            tdes.Clear();
            //return the Clear decrypted TEXT
            return UTF8Encoding.UTF8.GetString(resultArray);
        }


        //public string Encrypt_RSA(string encryptSrc)     // Метод асимметричного шифрования
        //{
        //    var publicKey = "<RSAKeyValue><Modulus>21wEnTU+mcD2w0Lfo1Gv4rtcSWsQJQTNa6gio05AOkV/Er9w3Y13Ddo5wGtjJ19402S71HUeN0vbKILLJdRSES5MHSdJPSVrOqdrll/vLXxDxWs/U0UT1c8u6k/Ogx9hTtZxYwoeYqdhDblof3E75d9n2F0Zvf6iTb4cI7j6fMs=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";

        //    cspp.KeyContainerName = "Key01";
        //    rsa = new RSACryptoServiceProvider(cspp);
        //    rsa.FromXmlString(publicKey);
        //    rsa.PersistKeyInCsp = true;

        //    // Создание закрытого ключа
        //    cspp.KeyContainerName = "Key01";

        //    rsa = new RSACryptoServiceProvider(cspp);
        //    rsa.PersistKeyInCsp = true;

        //    byte[] encryptSrcByte = Convert.FromBase64String(encryptSrc);
        //    byte[] encryptedSrc;
        //    encryptedSrc = rsa.Encrypt(encryptSrcByte, false);


        //    return Convert.ToBase64String(encryptedSrc);
        //}
        public string RSAEncryption(string strText)
        {
            var publicKey = publicKeyImported;

            var testData = Encoding.UTF8.GetBytes(strText);

            using (var rsa = new RSACryptoServiceProvider(1024))
            {
                try
                {
                    // client encrypting data with public key issued by server                    
                    rsa.FromXmlString(publicKey.ToString());

                    var encryptedData = rsa.Encrypt(testData, true);

                    var base64Encrypted = Convert.ToBase64String(encryptedData);

                    return base64Encrypted;
                }
                finally
                {
                    rsa.PersistKeyInCsp = false;
                }
            }
        }

        public string RSADecryption(string strText)
        {
            var privateKey = "<RSAKeyValue><Modulus>21wEnTU+mcD2w0Lfo1Gv4rtcSWsQJQTNa6gio05AOkV/Er9w3Y13Ddo5wGtjJ19402S71HUeN0vbKILLJdRSES5MHSdJPSVrOqdrll/vLXxDxWs/U0UT1c8u6k/Ogx9hTtZxYwoeYqdhDblof3E75d9n2F0Zvf6iTb4cI7j6fMs=</Modulus><Exponent>AQAB</Exponent><P>/aULPE6jd5IkwtWXmReyMUhmI/nfwfkQSyl7tsg2PKdpcxk4mpPZUdEQhHQLvE84w2DhTyYkPHCtq/mMKE3MHw==</P><Q>3WV46X9Arg2l9cxb67KVlNVXyCqc/w+LWt/tbhLJvV2xCF/0rWKPsBJ9MC6cquaqNPxWWEav8RAVbmmGrJt51Q==</Q><DP>8TuZFgBMpBoQcGUoS2goB4st6aVq1FcG0hVgHhUI0GMAfYFNPmbDV3cY2IBt8Oj/uYJYhyhlaj5YTqmGTYbATQ==</DP><DQ>FIoVbZQgrAUYIHWVEYi/187zFd7eMct/Yi7kGBImJStMATrluDAspGkStCWe4zwDDmdam1XzfKnBUzz3AYxrAQ==</DQ><InverseQ>QPU3Tmt8nznSgYZ+5jUo9E0SfjiTu435ihANiHqqjasaUNvOHKumqzuBZ8NRtkUhS6dsOEb8A2ODvy7KswUxyA==</InverseQ><D>cgoRoAUpSVfHMdYXW9nA3dfX75dIamZnwPtFHq80ttagbIe4ToYYCcyUz5NElhiNQSESgS5uCgNWqWXt5PnPu4XmCXx6utco1UVH8HGLahzbAnSy6Cj3iUIQ7Gj+9gQ7PkC434HTtHazmxVgIR5l56ZjoQ8yGNCPZnsdYEmhJWk=</D></RSAKeyValue>";
            //var privateKey = publicKeyImported;

            var testData = Encoding.UTF8.GetBytes(strText);

            using (var rsa = new RSACryptoServiceProvider(1024))
            {
                try
                {
                    var base64Encrypted = strText;

                    // server decrypting data with private key                    
                    rsa.FromXmlString(privateKey);

                    var resultBytes = Convert.FromBase64String(base64Encrypted);
                    var decryptedBytes = rsa.Decrypt(resultBytes, true);
                    var decryptedData = Encoding.UTF8.GetString(decryptedBytes);
                    return decryptedData.ToString();
                }
                finally
                {
                    rsa.PersistKeyInCsp = false;
                }
            }
        }

        class Sender
        {
            public static byte[] senderPublicKey;

            public static byte[] Do(string message)
            {
                using (ECDiffieHellmanCng alice = new ECDiffieHellmanCng())
                {

                    alice.KeyDerivationFunction = ECDiffieHellmanKeyDerivationFunction.Hash;
                    alice.HashAlgorithm = CngAlgorithm.Sha256;
                    senderPublicKey = alice.PublicKey.ToByteArray();
                    Reciever bob = new Reciever();
                    CngKey k = CngKey.Import(bob.recieverPublicKey, CngKeyBlobFormat.EccPublicBlob);
                    byte[] aliceKey = alice.DeriveKeyMaterial(CngKey.Import(bob.recieverPublicKey, CngKeyBlobFormat.EccPublicBlob));
                    byte[] encryptedMessage = null;
                    byte[] iv = null;
                    Send(aliceKey, message, out encryptedMessage, out iv);
                    return encryptedMessage;
                    //bob.Receive(encryptedMessage, iv);
                }
            }

            private static void Send(byte[] key, string secretMessage, out byte[] encryptedMessage, out byte[] iv)
            {
                using (Aes aes = new AesCryptoServiceProvider())
                {
                    aes.Key = key;
                    iv = aes.IV;

                    // Encrypt the message
                    using (MemoryStream ciphertext = new MemoryStream())
                    using (CryptoStream cs = new CryptoStream(ciphertext, aes.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        byte[] plaintextMessage = Encoding.UTF8.GetBytes(secretMessage);
                        cs.Write(plaintextMessage, 0, plaintextMessage.Length);
                        cs.Close();
                        encryptedMessage = ciphertext.ToArray();
                    }
                }
            }
        }
        public class Reciever
        {
            public byte[] recieverPublicKey;
            private byte[] recieverKey;
            public Reciever()
            {
                using (ECDiffieHellmanCng bob = new ECDiffieHellmanCng())
                {

                    bob.KeyDerivationFunction = ECDiffieHellmanKeyDerivationFunction.Hash;
                    bob.HashAlgorithm = CngAlgorithm.Sha256;
                    recieverPublicKey = bob.PublicKey.ToByteArray();
                    recieverKey = bob.DeriveKeyMaterial(CngKey.Import(Sender.senderPublicKey, CngKeyBlobFormat.EccPublicBlob));
                }
            }

            public void Receive(byte[] encryptedMessage, byte[] iv)
            {

                using (Aes aes = new AesCryptoServiceProvider())
                {
                    aes.Key = recieverKey;
                    aes.IV = iv;
                    // Decrypt the message
                    using (MemoryStream plaintext = new MemoryStream())
                    {
                        using (CryptoStream cs = new CryptoStream(plaintext, aes.CreateDecryptor(), CryptoStreamMode.Write))
                        {
                            cs.Write(encryptedMessage, 0, encryptedMessage.Length);
                            cs.Close();
                            string message = Encoding.UTF8.GetString(plaintext.ToArray());
                            Console.WriteLine(message);
                        }
                    }
                }
            }
        }

        public async void Form1_Load(object sender = null, EventArgs e = null)
        {
            string sqlCon = @"Data Source=(LocalDB)\MSSQLLocalDB;" +
                            @"AttachDbFilename=|DataDirectory|\Database.mdf;
                Integrated Security=True;
                Connect Timeout=30;";
            SqlConnection Con = new SqlConnection(sqlCon);
            Con.Open();
            SqlDataReader sqlReader = null;

            SqlCommand command = new SqlCommand("SELECT * FROM [Users]", Con);

            try
            {
                sqlReader = await command.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    comboBox1.Items.Add(Convert.ToString(sqlReader["Users"]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }
            Con.Close();
            comboBox1.SelectedIndex = 0;
        }

        private async void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sqlCon = @"Data Source=(LocalDB)\MSSQLLocalDB;" +
                @"AttachDbFilename=|DataDirectory|\Database.mdf;
                Integrated Security=True;
                Connect Timeout=30;";
            SqlConnection Con = new SqlConnection(sqlCon);
            Con.Open();
            SqlDataReader sqlReader = null;

            SqlCommand command = new SqlCommand("SELECT * FROM [Users] where Users = '" + comboBox1.Text + "'", Con);

            try
            {
                sqlReader = await command.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    textBox7.Text = Convert.ToString(sqlReader["Login"]);
                    textBox6.Text = Convert.ToString(sqlReader["Password"]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }
            Con.Close();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            NewUser newUser = new NewUser();
            newUser.Owner = this;
            newUser.ShowDialog();
        }

        private void button8_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем открытый ключ RSA
            StreamReader sr = new StreamReader(openFileDialog1.FileName);
            string keytxt = sr.ReadToEnd();
            publicKeyImported = keytxt;
            label14.Text = openFileDialog1.FileName;
            openFileDialog1.Dispose();
            openFileDialog1.Reset();

        }

        private void button6_Click(object sender, EventArgs e)
        {
            // Экспорт открытого ключа, созданного RSA шифрованием
            // в файл.
            rsa = new RSACryptoServiceProvider(cspp);
            StreamWriter sw = new StreamWriter("importedKey.txt");
            sw.Write(rsa.ToXmlString(false));
            sw.Close();

        }
    }
}
