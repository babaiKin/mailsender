using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.Net.Mail;
using Microsoft.VisualBasic.FileIO;

namespace mailsender
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        async private void button1_Click(object sender, EventArgs e)
        {
            string mail="";
            int count=0;


            for (int i = 0; i < (dataGridView1.RowCount - 1); i++)
            {
                try
                {
                    mail = (dataGridView1.Rows[i].Cells[1].Value).ToString();
                    //string logpass = (dataGridView1.Rows[i].Cells[2].Value).ToString();

                    if (mail != "")
                    {
                        Sender(mail);
                        if (i != 0 && i % 100 == 0)
                        {
                            await Task.Delay(1000 * 60 * 10); //каждые 100 записей остановка на 10 минут (1000мс*60с*10мин)
                        }
                        await Task.Delay(1000 * 1); //10 секунд задержки
                        System.IO.File.AppendAllText(@".\Files\send_log.txt", mail + " Отправлено " + DateTime.Now + "\n");
                    }
                    count++;
                }
                
                catch (Exception ex)
                {
                    System.IO.File.AppendAllText(@".\Files\err_log.txt", DateTime.Now + "\n");
                    System.IO.File.AppendAllText(@".\Files\err_log.txt", "Проблемы с отправкой на адрес:" + mail + "\n");
                    System.IO.File.AppendAllText(@".\Files\err_log.txt", ex + "\n");
                }                
            }
            MessageBox.Show("РАССЫЛКА ОКОНЧЕНА \nПисем отправлено: " + count.ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string csvFile;
            OpenFileDialog sfd = new OpenFileDialog();
            sfd.Filter = "CSV Files (*.csv)|*.csv";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                csvFile= sfd.FileName;
                BindData(csvFile);
                //MessageBox.Show(csvFile);
            }
            
        }


        private void Sender(string mail)
        {
            /////////////////////////////////
            //РАБОТАЕТ!!!! НЕ ТРОЖЬ!!!!!!!!!!
            //дописать цикл отправки сообщений по разным адресам
            /////////////////////////////////
            
            string file1 = @".\Files\Письмо предложение.pdf";
            string file2 = @".\Files\Письмо предприятиям январь приложение 1.pdf";
            string file3 = @".\Files\Письмо предприятиям январь приложение 2.pdf";

            SmtpClient Smtp = new SmtpClient("mail.nncsm.ru", 25);
            Smtp.EnableSsl = true;
            Smtp.Credentials = new NetworkCredential("nikolaevn", "Bdkbtd2021");    // Логин и пароль почты отправителя            
            MailMessage MyMessage = new MailMessage();
            MyMessage.From = new MailAddress("no-reply@nncsm.ru"); // От кого отправляем почту
            MyMessage.To.Add(mail);    // Кому отправляем почту
            //MyMessage.To.Add("nikolaevn@nncsm.ru");   // Кому отправляем почту

            MyMessage.Attachments.Add(new Attachment(file1));
            MyMessage.Attachments.Add(new Attachment(file2));
            MyMessage.Attachments.Add(new Attachment(file3));

            MyMessage.Subject = "Письмо предложение от ФБУ Нижегородский ЦСМ";   // Тема письма
            

            MyMessage.Body = "<div style=\"color: #003366;\">Уважаемый руководитель!</div>" +

                "<br>" +
                "<div style=\"color: #003366;\">Федеральное агентство «Росстандарт» в лице ФБУ «Нижегородский ЦСМ» направляет ежегодное уведомление о необходимости своевременной поверки и регулировки средств измерений, оборудования и инструмента.</div>" +
                "<br>" +
                "<div style=\"color: #003366;\">Получите предварительный расчет стоимости поверки и других метрологических услуг для вашей организации по тел. 8 (800) 200-22-14, mail@nncsm.ru.</ div > " +
                "<br>" +
                "<br>" +
                "<img src=https://nncsm.ru/assets/templates/csm/images/rst_top.png>" +
                "<div style=\"color: #003366;\"><a href=http://nncsm.ru>www.nncsm.ru</a></div>" +
                "<div style=\"color: #003366;\">603950, г.Нижний Новгород</div>" +
                "<div style=\"color: #003366;\">ул.Республиканская, д. 1</div>" +
                "<img src =https://nncsm.ru/assets/templates/csm/images/rst_bottom.png>";
            MyMessage.IsBodyHtml = true;

            Smtp.Send(MyMessage);
            //MessageBox.Show("Сообщение успешно отправлено");

            /////////////////////////////////
            //РАБОТАЕТ!!!! НЕ ТРОЖЬ!!!!!!!!!!
            /////////////////////////////////
        }

        private void BindData(string filePath)
        {
            DataTable dt = new DataTable();
            string[] lines = System.IO.File.ReadAllLines(filePath);
            if (lines.Length > 0)
            {
                //first line to create header
                string firstLine = lines[0];
                string[] headerLabels = firstLine.Split(';');
                foreach (string headerWord in headerLabels)
                {
                    dt.Columns.Add(new DataColumn(headerWord));
                }
                //For Data
                for (int i = 1; i < lines.Length; i++)
                {
                    string[] dataWords = lines[i].Split(';');
                    DataRow dr = dt.NewRow();
                    int columnIndex = 0;
                    foreach (string headerWord in headerLabels)
                    {
                        dr[headerWord] = dataWords[columnIndex++];
                    }
                    dt.Rows.Add(dr);
                }
            }
            if (dt.Rows.Count > 0)
            {
                dataGridView1.DataSource = dt;
            }

        }
    }
}
