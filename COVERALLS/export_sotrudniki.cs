using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Media;
using System.Drawing.Printing;
using System.IO;

namespace COVERALLS
{
    public partial class export_sotrudniki : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
     
        private Microsoft.Office.Interop.Excel.Application ObjExcel;
        private Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
        private Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;

        public export_sotrudniki()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label2.Visible = true;
                label2.Text = openFileDialog1.FileName;
            }
                       
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            if (openFileDialog1.FileName.Length > 0)
            {
                button2.Visible = true;
                comboBox1.Visible = true;
                label1.Visible = true;
            }
            else
            {
                button2.Visible = false;
                comboBox1.Visible = false;
                label1.Visible = false;
            }         

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не выбран филиал!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
                        
            /////Создание объекта Экспортируемый файл
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            //Присваивание переменной кода филиала
            int br = 0;
            if (comboBox1.Text == "Красноярский филиал ОАО \"СибЭР\"") { br = 4; }
            if (comboBox1.Text == "Абаканский филиал ОАО \"СибЭР\"") { br = 1; }
            if (comboBox1.Text == "Кемеровский филиал ОАО \"СибЭР\"") { br = 3; }
            if (comboBox1.Text == "Барнаульский филиал ОАО \"СибЭР\"") { br = 2; }
            if (comboBox1.Text == "Исполнительный аппарат") { br = 5; }

            //Задание массива для хранения импортируемых ФИО
            String[] array_fio = new String[ObjWorkSheet.UsedRange.Rows.Count-4]; //Задание кол-ва элементов массива - это кол-во записей для импорта
            
            for (int j = 5; j <= ObjWorkSheet.UsedRange.Rows.Count; j++)
            {
                string fio=Convert.ToString(ObjWorkSheet.Cells[j, 4].value);
                string prof = Convert.ToString(ObjWorkSheet.Cells[j, 6].value);
                string dep = Convert.ToString(ObjWorkSheet.Cells[j, 7].value);
                string size_shoes = Convert.ToString(ObjWorkSheet.Cells[j, 14].value);
                string size_clothes = Convert.ToString(ObjWorkSheet.Cells[j, 15].value);
                string height = Convert.ToString(ObjWorkSheet.Cells[j, 16].value);


                if (fio != null)
                {
                    //Проверка на существование в БД импортируемого сотрудника. Если существует, то показываем ФИО и переходим к следующей записи
                    SqlCommand command = new SqlCommand("select FIO from SOTRUDNIKI where FIO=@ff", conn);
                    SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    command.Parameters.AddWithValue("ff", fio);
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    if (ds.Tables[0].Rows.Count != 0)
                    {
                        SystemSounds.Beep.Play();
                        MessageBox.Show("Сотрудник " + fio + " уже существует в базе! Экспорт данного сотрудника будет пропущен.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        continue;
                    }
                }
                
                /////Вставка данных в таблицу SOTRUDNIKI
                SqlCommand scmd4 = conn.CreateCommand();
                scmd4.CommandText = "insert into SOTRUDNIKI (FIO,PROFESSION,BRANCH,DEPARTMENT,USER_ID,USERS_BRANCH_ID,NEED_SIZ,SIZE_SHOES,SIZE_CLOTHES,HEIGHT) values (@f, @pr, @br, @dp, @us_id, @br_id, @nd_siz, @size_sh, @size_cl, @he)";
                scmd4.Parameters.AddWithValue("f",fio);
                scmd4.Parameters.AddWithValue("pr", prof);
                scmd4.Parameters.AddWithValue("br", comboBox1.Text);
                scmd4.Parameters.AddWithValue("dp", dep);
                scmd4.Parameters.AddWithValue("us_id", Form2.val);
                scmd4.Parameters.AddWithValue("br_id", br);
                scmd4.Parameters.AddWithValue("nd_siz", 0);
                if (size_shoes == null) { scmd4.Parameters.AddWithValue("size_sh", DBNull.Value); } else { scmd4.Parameters.AddWithValue("size_sh", size_shoes); } //Если значение в параметре пустое, то нужно применить свойство DBNull.Value, иначе выводится ошибка
                if (size_clothes == null) { scmd4.Parameters.AddWithValue("size_cl", DBNull.Value); } else { scmd4.Parameters.AddWithValue("size_cl", size_clothes); }
                if (height == null) { scmd4.Parameters.AddWithValue("he", DBNull.Value); } else { scmd4.Parameters.AddWithValue("he", height); }
                try
                {
                    conn.Open();
                }
                catch
                {
                    MessageBox.Show("Ошибка соединения с базой данных");
                }
                SqlDataReader reader4;
                reader4 = scmd4.ExecuteReader();
                conn.Close();
                //////////////////
                

                array_fio.SetValue(fio, j - 5);

            }

            SystemSounds.Beep.Play();
            MessageBox.Show("Импорт сотрудников завершен удачно!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            for (int t = 0; t <= array_fio.Length-1; t++)
            {
                if (array_fio.GetValue(t)!=null) //Если элемент массива не пустой, тогда отображаем его, иначе пропускаем
                {
                    richTextBox1.AppendText(array_fio.GetValue(t).ToString());
                    richTextBox1.AppendText("\n");
                }
                else continue;
            }
            
            ObjExcel.Quit();
                    
        }

        private void export_sotrudniki_Load(object sender, EventArgs e)
        {
            //Заполнение поля Филиал, данными из БД
            SqlCommand command3 = conn.CreateCommand();
            command3.CommandText = "select BRANCH from BRANCHES";
            try
            {
                conn.Open();
            }
            catch { }
            SqlDataReader reader;
            reader = command3.ExecuteReader();
            while (reader.Read())
            {
                try
                {
                    string result = reader.GetString(0);
                    comboBox1.Items.Add(result);
                }
                catch { }

            }
            conn.Close();
            ////////////////////////
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            //Проверка на изменение поля если текст введен руками и такого нет в БД, то обнуляем поле
            SqlCommand command = new SqlCommand("select BRANCH from BRANCHES where BRANCH='" + comboBox1.Text + "'", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "BRANCHES");

            if (ds.Tables[0].Rows.Count == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Выбор значения возможен только из списка!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                comboBox1.Text = "";
                return;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (richTextBox1.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Печатать нечего!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (printDialog1.ShowDialog() == DialogResult.OK)
            {
                StringReader reader = new StringReader(richTextBox1.Text);
                printDocument1.PrintPage += new PrintPageEventHandler(DocumentToPrint_PrintPage);
                printDocument1.Print();
            }
        }

        private void DocumentToPrint_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            StringReader reader = new StringReader(richTextBox1.Text);
            float LinesPerPage = 0;
            float YPosition = 0;
            int Count = 0;
            float LeftMargin = e.MarginBounds.Left;
            float TopMargin = e.MarginBounds.Top;
            string Line = null;
            Font PrintFont = this.richTextBox1.Font;
            SolidBrush PrintBrush = new SolidBrush(Color.Black);

            LinesPerPage = e.MarginBounds.Height / PrintFont.GetHeight(e.Graphics);

            while (Count < LinesPerPage && ((Line = reader.ReadLine()) != null))
            {
                YPosition = TopMargin + (Count * PrintFont.GetHeight(e.Graphics));
                e.Graphics.DrawString(Line, PrintFont, PrintBrush, LeftMargin, YPosition, new StringFormat());
                Count++;
            }

            if (Line != null)
            {
                e.HasMorePages = true;
            }
            else
            {
                e.HasMorePages = false;
            }
            PrintBrush.Dispose();
        }
    }
}
