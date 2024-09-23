using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Media;


namespace COVERALLS
{
    public partial class Form6 : Form
    {
        public Form6()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (richTextBox1.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не введен текст сообщения!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            try
            {
                // Create the Outlook application by using inline initialization.
                Outlook.Application oApp = new Outlook.Application();

                //Create the new message by using the simplest approach.
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);

                //Add a recipient.
                // TODO: Change the following recipient where appropriate.
                Outlook.Recipient oRecip = (Outlook.Recipient)oMsg.Recipients.Add("skvortcovoi@sibgenco.ru");
                oRecip.Resolve();

                //Set the basic properties.
                oMsg.Subject = "Проблема с программой \"Учет спецодежды\"";
                oMsg.Body = richTextBox1.Text;

                //Add an attachment.
                // TODO: change file path where appropriate
                /*String sSource = "C:\\setupxlg.txt";
                String sDisplayName = "MyFirstAttachment";
                int iPosition = (int)oMsg.Body.Length + 1;
                int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                Outlook.Attachment oAttach = oMsg.Attachments.Add(sSource, iAttachType, iPosition, sDisplayName);
                */
                // If you want to, display the message.
                // oMsg.Display(true);  //modal

                //Send the message.
                oMsg.Save();
                oMsg.Send();

                //Explicitly release objects.
                oRecip = null;
                //oAttach = null;
                oMsg = null;
                oApp = null;

                SystemSounds.Beep.Play();
                MessageBox.Show("Сообщение отправлено удачно. В ближайщее время ваш вопрос будет рассмотрен.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                richTextBox1.Clear();
                this.Close();
            }

                   // Simple error handler.
            catch (Exception)
            {
                //Console.WriteLine("{0} Exception caught: ", e);
                SystemSounds.Beep.Play();
                MessageBox.Show("Не удалось отправить сообщение!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}
