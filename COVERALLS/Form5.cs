using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Media;
using System.Data.SqlClient;
using System.Security.Cryptography; 


namespace COVERALLS
{
    public partial class Form5 : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
        
        Form2 form2 = new Form2();
        
        public Form5()
        {
            InitializeComponent();
        }

        //Функция шифрования с помощью алгоритма MD5
        string GetHashString(string s)
        {
            //переводим строку в байт-массим  
            byte[] bytes = Encoding.Unicode.GetBytes(s);

            //создаем объект для получения средст шифрования  
            MD5CryptoServiceProvider CSP = new MD5CryptoServiceProvider();

            //вычисляем хеш-представление в байтах  
            byte[] byteHash = CSP.ComputeHash(bytes);

            string hash = string.Empty;

            //формируем одну цельную строку из массива  
            foreach (byte b in byteHash)
                hash += string.Format("{0:x2}", b);

            return hash;
        }
        /////////////////////////////////////////


        private void button1_Click(object sender, EventArgs e)
        {
            if ((maskedTextBox1.Text.Length == 0) || (maskedTextBox2.Text.Length == 0))
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Заполнены не все поля!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (maskedTextBox1.Text != maskedTextBox2.Text)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Введенные пароли не совпадают", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            /////////Обновление пароля
            SqlCommand scmd2 = conn.CreateCommand();
            scmd2.CommandText = "UPDATE USERS SET PASSW=" + "'" + GetHashString(maskedTextBox2.Text) + "' WHERE ID=" + Form2.val;
            try
            {
                conn.Open();
            }
            catch
            {
                MessageBox.Show("Ошибка соединения с базой данных");
            }
            SqlDataReader reader2;
            reader2 = scmd2.ExecuteReader();
            conn.Close();
            //////////////////

            SystemSounds.Beep.Play();
            MessageBox.Show("Пароль изменен успешно", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            maskedTextBox1.Clear();
            maskedTextBox2.Clear();
            this.Close();
        }

       
    }
}
