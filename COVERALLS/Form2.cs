using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Media;
using System.Security.Cryptography;
using System.Diagnostics;
using System.Configuration;
using System.IO;
using System.Globalization;

namespace COVERALLS
{
    public partial class Form2 : Form
    {

        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");

        SqlDataAdapter da;
        public static int val;
        public static string name_user;
        public static string administration;
        public static string branch;
        public static string psw;


        public Form2()
        {
            InitializeComponent();
            this.AcceptButton = button1; //Задает кнопку, которая нажимается при нажатии на ENTER
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

        private void button1_Click(object sender, EventArgs e)
        {
            if ((textBox1.Text.Length == 0) || (maskedTextBox1.Text.Length == 0))
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Заполнены не все поля!", "Внимание", MessageBoxButtons.OK);
                return;
            }

            //Удаляем файл существующий, перед записью в него введенного логина
            if (File.Exists("login.txt")) //Создание файла для хранения последнего введеного логина, чтобы при входе он его показывал
            {
                File.Delete("login.txt");
            }
            //Запись введенного логина в файл
            StreamWriter ff = new StreamWriter(Environment.CurrentDirectory + @"\login.txt", true);
            ff.Write(textBox1.Text);
            ff.Close();
            //

            string pass = GetHashString(maskedTextBox1.Text);

            conn.Open();
            SqlCommand mycommand = new SqlCommand("select * from users where user_name=" + "'" + textBox1.Text + "' and passw='" + pass + "'", conn);

            da = new SqlDataAdapter(mycommand);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            da.Fill(ds, "USERS");
            if (ds.Tables[0].Rows.Count != 0)
            {
                object value = ds.Tables[0].Rows[0][0].ToString();
                if (value != null)
                {

                    val = Convert.ToInt16(value);
                    name_user = "Пользователь: " + ds.Tables[0].Rows[0][2].ToString();
                    administration = ds.Tables[0].Rows[0][7].ToString();
                    branch = ds.Tables[0].Rows[0][8].ToString();
                    psw = ds.Tables[0].Rows[0][3].ToString();

                    /*
                    //Проверка не работает ли уже пользователь с таким именем в системе
                    conn.Open();
                    SqlCommand mycommand1 = new SqlCommand("select * from active_users where id_user=" + val, conn);
                    SqlDataAdapter da1 = new SqlDataAdapter(mycommand1);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                    DataSet ds1 = new DataSet();
                    conn.Close();
                    da1.Fill(ds1, "ACTIVE_USERS");
                    if (ds1.Tables[0].Rows.Count != 0)
                    {
                        SystemSounds.Beep.Play();
                        MessageBox.Show("Пользователь "+name_user+" уже находится в системе. Параллельное нахождение 2-х и более пользователей под одним логином в системе запрещено!", "Внимание", MessageBoxButtons.OK);
                        return;
                    } */
                    
                    Form1 form1 = new Form1();
                    form1.Show();
                    this.Visible = false;

                    ////////////////////////////////////        
                    
                }

            }
            else
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не верно введено либо имя либо пароль!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

        }

        private void Form2_Load(object sender, EventArgs e)
        {
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(new CultureInfo("en-US")); //Переключение раскладки клавы
            //InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(new CultureInfo("ru-RU")); //Переключение раскладки клавы

            if (!File.Exists("login.txt")) //Создание файла для хранения последнего введеного логина, чтобы при входе он его показывал
            {
                FileStream fs = File.Create("login.txt");
                fs.Close();
            }
            else
            {
                StreamReader rr = File.OpenText("login.txt");
                textBox1.Text = rr.ReadLine();
                rr.Close();
            }
            //Запуска обновляльщика
             Process p = new Process();
            p.StartInfo.FileName = Environment.CurrentDirectory + @"\updater.exe";
            p.Start(); 

        }
        /////////////////////////////////////////

    }
}
