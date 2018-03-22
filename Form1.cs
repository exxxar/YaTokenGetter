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

namespace RegAccYandex
{
    public partial class Form1 : Form
    {
        public class user
        {
            public string login { get; set; }
            public string pass { get; set; }
            public string prog_id { get; set; }
            public string prog_pass { get; set; }

            public string toString()
            {
                return String.Format("{0};{1};{2};{3}",
                    this.login,
                    this.pass,
                    this.prog_id,
                    this.prog_pass
                    );
            }
        }

        List<user> users;
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
            users = new List<user>();

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

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                listBox1.Items.Clear();
                users.Clear();
                string filePath = openFileDialog1.FileName;
                var task = Task.Run(() => this.readExecel(filePath));
                Task.WaitAll(new Task[] { task });
                users.ForEach(u => listBox1.Items.Add(u.toString()));
            }
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }

        public async void readExecel(string filePath)
        {
            
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            if (rowCount < 2)
                return;

            Console.WriteLine("row=>{0} col=>{0}", rowCount, colCount);


            for (int i = 2; i <= rowCount; i++)
            {

                user u = new user();
                u.login = ""+xlRange.Cells[i, 3].Value2.ToString();
                u.pass = "" + xlRange.Cells[i, 4].Value2.ToString();
                u.prog_id = "" + xlRange.Cells[i, 8].Value2.ToString();
                u.prog_pass = "" + xlRange.Cells[i, 9].Value2.ToString();

                users.Add(u);
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        

            

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            user u = users[listBox1.SelectedIndex];
            textBox3.Text = u.login;
            textBox4.Text = u.pass;
            textBox1.Text = u.prog_id;
            textBox5.Text = u.prog_pass;

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
