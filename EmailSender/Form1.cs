using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using System.Speech;
using System.Speech.Synthesis;

namespace EmailSender
{
    public partial class Form1 : Form
    {
        SpeechSynthesizer obj = new SpeechSynthesizer();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            obj.SpeakAsync("Welcome to Email Sender Application");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox4.Text.Contains("@gmail.com"))
                {
                    //obj.SpeakAsync("You can only use gmail account to send email to any other account");
                    //MessageBox.Show("You can only use gmail account to send email to any other account", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    MailMessage msg = new MailMessage();
                    msg.From = new MailAddress(textBox4.Text);
                    msg.Subject = textBox2.Text;
                    msg.Body = textBox3.Text;
                    foreach (string s in textBox1.Text.Split(';'))
                        msg.To.Add(s);
                    SmtpClient client = new SmtpClient();
                    client.Credentials = new NetworkCredential(textBox4.Text, textBox5.Text);
                    client.Host = "smtp.gmail.com";
                    client.Port = 587;
                    client.EnableSsl = true;
                    client.Send(msg);
                    obj.SpeakAsync("Message Sent Successfully.");
                    client.SendCompleted += new SendCompletedEventHandler(client_SendCompleted);
                }
                else if (textBox4.Text.Contains("@yahoo.com"))
                {
                    MailMessage msg = new MailMessage();
                    msg.From = new MailAddress(textBox4.Text);
                    msg.Subject = textBox2.Text;
                    msg.Body = textBox3.Text;
                    foreach (string s in textBox1.Text.Split(';'))
                        msg.To.Add(s);
                    SmtpClient client = new SmtpClient("smtp.mail.yahoo.com");
                    client.Credentials = new NetworkCredential(textBox4.Text, textBox5.Text);
                    client.Host = "smtp.mail.yahoo.com";
                    client.Port = 587;
                    client.EnableSsl = true;
                    client.Send(msg);
                    obj.SpeakAsync("Message Sent Successfully.");
                }
            }
            catch 
            {
                obj.SpeakAsync("There was an error sending the message. Make sure you typed in your credentials correctly and you have an internet connection.");
                MessageBox.Show("There was an error sending the message. Make sure you typed in\r\nyour credentials(Sender Details correctly and you have an internet connection."
                                    ,"Error",MessageBoxButtons.OK,MessageBoxIcon.Error); 
            }
            finally { button1.Enabled = true; }
        }

        void client_SendCompleted(object sender, AsyncCompletedEventArgs e)
        {
            throw new NotImplementedException();
            if (e.Cancelled)
            {
                obj.SpeakAsync("Message not sent. Try Again");            
            }
            if (e.Error != null)
            {
                obj.SpeakAsync("Unexpected Error occures.Please try again after some time.");
            }
            else
            {
                obj.SpeakAsync("Message Sent Successfully.");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                obj.SpeakAsync("Recipient's are ");
                obj.SpeakAsync(textBox1.Text);
            }
            else
            {
                obj.SpeakAsync("No Recipient");
            }
            if (textBox2.Text != "")
            {
                obj.SpeakAsync("Subject of email is ");
                obj.SpeakAsync(textBox2.Text);
            }
            else
            {
                obj.SpeakAsync("No Subject");
            }
            if (textBox3.Text != "")
            {
                obj.SpeakAsync("Body of Email is");
                obj.SpeakAsync(textBox3.Text);
            }
            else
            {
                obj.SpeakAsync("No Body ");
            }
            if (textBox4.Text != "")
            {
                obj.SpeakAsync("your Email is ");
                obj.SpeakAsync(textBox4.Text);
            }
            else
            {
                obj.SpeakAsync("No Sender Address");
            }
        }
    }
}
