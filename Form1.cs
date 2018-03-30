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
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;

namespace RegAccYandex
{
    public partial class Form1 : Form
    {

        List<User> users;
        SeleniumModul sm;
        ExcelFile excel;

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
            users = new List<User>();
            excel = new ExcelFile();

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
                var retData = this.sm.takeToken(textBox1.Text, textBox5.Text, textBox3.Text, textBox4.Text);
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
                File.Open("accounts.txt", FileMode.Append));
            sw.WriteLine(data);
            sw.Close();
            Task.Run(() =>
            {

                SendMail("smtp.yandex.ru", "azyexxxar@yandex.ru", "93svxrpj", "exxxar@gmail.com", "Interpromotion yandex accounts", data);
                MessageBox.Show("Тоукен сгенерирован и сохранен в файл!");


            });
            button1.Enabled = true;
            textBox1.Text = textBox3.Text = textBox4.Text = textBox5.Text = textBox6.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                listBox1.Enabled = false;
                listBox1.Items.Clear();
                users.Clear();
                string filePath = openFileDialog1.FileName;
                excel.readExecel(filePath);
                users.ForEach(u => listBox1.Items.Add(u.ToString()));
                listBox1.Enabled = true;
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            listBox1.Enabled = false;
            listBox1.Items.Clear();


            if (!File.Exists(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\accounts.xls"))
            {
                excel.createExcel(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\accounts.xls");

                MessageBox.Show("Excel file created , you can find the file " + System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\accounts.xls");
            }
            else
            {
                MessageBox.Show("Файл уже есть! Проверьте путь " + System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
            }
        }



        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            User u = users[listBox1.SelectedIndex];
            textBox3.Text = u.login;
            textBox4.Text = u.pass;
            textBox1.Text = u.prog_id;
            textBox5.Text = u.prog_pass;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            User u = sm.regYandexAcc();
            try
            {
                if (!File.Exists(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\accounts.csv"))
                    File.WriteAllText(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\accounts.csv", u.ToString() + "\n", Encoding.UTF8);
                else
                    File.AppendAllText(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\accounts.csv", u.ToString() + "\n", Encoding.UTF8);
            }catch
            {

            }

        }
    }
}
