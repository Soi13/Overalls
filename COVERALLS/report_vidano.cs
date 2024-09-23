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
using Microsoft.Office.Interop.Excel;

namespace COVERALLS
{
    public partial class report_vidano : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");

        SqlDataAdapter da;

        public report_vidano()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["FIO"].HeaderText = "ФИО";
            dataGridView1.Columns["FIO"].Width = 150;
            dataGridView1.Columns["NAIMEN_OVERALLS"].HeaderText = "Наименование";
            dataGridView1.Columns["NAIMEN_OVERALLS"].Width = 200;
            dataGridView1.Columns["ED_IZM"].HeaderText = "Ед.изм";
            dataGridView1.Columns["ED_IZM"].Width = 60;
            dataGridView1.Columns["KOLVO"].HeaderText = "Кол-во";
            dataGridView1.Columns["KOLVO"].Width = 60;
            dataGridView1.Columns["PERIOD_ISPOLZOVAN"].HeaderText = "Период исп-я";
            dataGridView1.Columns["PERIOD_ISPOLZOVAN"].Width = 100;
            dataGridView1.Columns["DATE_VIDACHI"].HeaderText = "Дата выдачи";
            dataGridView1.Columns["DATE_VIDACHI"].Width = 100;
            dataGridView1.Columns["DATE_SLED_VIDACHI"].HeaderText = "Дата след. выдачи";
            dataGridView1.Columns["DATE_SLED_VIDACHI"].Width = 100;
            dataGridView1.Columns["SIZE"].HeaderText = "Размер";
            dataGridView1.Columns["SIZE"].Width = 50;
            dataGridView1.Columns["IAIS_NAIMEN_OVERALLS"].HeaderText = "Наименование ИАИС";
            dataGridView1.Columns["IAIS_NAIMEN_OVERALLS"].Width = 200;
            
        }
        //////////////////////////////////////////////////////

        private void report_vidano_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;
            statusStrip1.Items[0].Text = "Всего записей: 0";

            //Заполнение поля профессия, данными из БД
            SqlCommand command2 = conn.CreateCommand();
            command2.CommandText = "select distinct PROFESSION from SOTRUDNIKI order by PROFESSION";
            try
            {
                conn.Open();
            }
            catch { }
            SqlDataReader reader1;
            reader1 = command2.ExecuteReader();
            while (reader1.Read())
            {
                try
                {
                    string result1 = reader1.GetString(0);
                    comboBox3.Items.Add(result1);                    
                }
                catch { }

            }
            conn.Close();
            ////////////////////////


            //Заполнение поля Отдел, данными из БД
            SqlCommand command3 = conn.CreateCommand();
            command3.CommandText = "select distinct DEPARTMENT from SOTRUDNIKI order by DEPARTMENT";
            try
            {
                conn.Open();
            }
            catch
            {
                MessageBox.Show("Ошибка соединения с базой данных");
            }
            SqlDataReader reader2;
            reader2 = command3.ExecuteReader();
            while (reader2.Read())
            {
                try
                {
                    string result2 = reader2.GetString(0);
                    comboBox2.Items.Add(result2);                    
                }
                catch { }

            }
            conn.Close();
            ////////////////////////


            //Заполнение поля по сотруднику, данными из БД
            if ((Form2.val == 115) || (Form2.val == 3)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ
            {
                SqlCommand command4 = conn.CreateCommand();
                command4.CommandText = "select distinct FIO from SOTRUDNIKI";
                try
                {
                    conn.Open();
                }
                catch { }
                SqlDataReader reader4;
                reader4 = command4.ExecuteReader();
                while (reader4.Read())
                {
                    try
                    {
                        string result4 = reader4.GetString(0);
                        comboBox1.Items.Add(result4);
                    }
                    catch { }

                }
                conn.Close();
                ////////////////////////
            }
            else
            {
                SqlCommand command4 = conn.CreateCommand();
                command4.CommandText = "select distinct FIO from SOTRUDNIKI where USERS_BRANCH_ID='" + Form2.branch + "' order by FIO";
                try
                {
                    conn.Open();
                }
                catch { }
                SqlDataReader reader4;
                reader4 = command4.ExecuteReader();
                while (reader4.Read())
                {
                    try
                    {
                        string result4 = reader4.GetString(0);
                        comboBox1.Items.Add(result4);
                    }
                    catch { }

                }
                conn.Close();
                ////////////////////////
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox4.Checked = false;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if ((checkBox1.Checked == false) && (checkBox2.Checked == false) && (checkBox3.Checked == false) && (checkBox4.Checked == false))
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не выбраны параметры поиска! Отображать нечего.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            
            if (checkBox1.Checked == true)
            {
                SqlCommand command = new SqlCommand("select SOTRUDNIKI.FIO, VIDANO.NAIMEN_OVERALLS, VIDANO.ED_IZM, VIDANO.KOLVO, VIDANO.PERIOD_ISPOLZOVAN,VIDANO.DATE_VIDACHI,VIDANO.DATE_SLED_VIDACHI,VIDANO.SIZE, VIDANO.IAIS_NAIMEN_OVERALLS from SOTRUDNIKI,VIDANO where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and SOTRUDNIKI.FIO='" + comboBox1.Text + "' and SOTRUDNIKI.USERS_BRANCH_ID='"+Form2.branch +"'", conn);
                da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "SOTRUDNIKI");
                dataGridView1.DataSource = ds.Tables[0];
                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview();
            }

            if (checkBox2.Checked == true)
            {
                SqlCommand command = new SqlCommand("select SOTRUDNIKI.FIO, VIDANO.NAIMEN_OVERALLS, VIDANO.ED_IZM, VIDANO.KOLVO, VIDANO.PERIOD_ISPOLZOVAN,VIDANO.DATE_VIDACHI,VIDANO.DATE_SLED_VIDACHI,VIDANO.SIZE, VIDANO.IAIS_NAIMEN_OVERALLS from SOTRUDNIKI,VIDANO where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and SOTRUDNIKI.DEPARTMENT='" + comboBox2.Text + "' and SOTRUDNIKI.USERS_BRANCH_ID='" + Form2.branch + "'", conn);
                da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "SOTRUDNIKI");
                dataGridView1.DataSource = ds.Tables[0];
                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview();
            }

            if (checkBox3.Checked == true)
            {
                SqlCommand command = new SqlCommand("select SOTRUDNIKI.FIO, VIDANO.NAIMEN_OVERALLS, VIDANO.ED_IZM, VIDANO.KOLVO, VIDANO.PERIOD_ISPOLZOVAN,VIDANO.DATE_VIDACHI,VIDANO.DATE_SLED_VIDACHI,VIDANO.SIZE, VIDANO.IAIS_NAIMEN_OVERALLS from SOTRUDNIKI,VIDANO where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and SOTRUDNIKI.PROFESSION='" + comboBox3.Text + "' and SOTRUDNIKI.USERS_BRANCH_ID='" + Form2.branch + "'", conn);
                da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "SOTRUDNIKI");
                dataGridView1.DataSource = ds.Tables[0];
                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview();
            }

            if (checkBox4.Checked == true)
            {
                SqlCommand command = new SqlCommand("select SOTRUDNIKI.FIO, VIDANO.NAIMEN_OVERALLS, VIDANO.ED_IZM, VIDANO.KOLVO, VIDANO.PERIOD_ISPOLZOVAN,VIDANO.DATE_VIDACHI,VIDANO.DATE_SLED_VIDACHI,VIDANO.SIZE, VIDANO.IAIS_NAIMEN_OVERALLS from SOTRUDNIKI,VIDANO where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and SOTRUDNIKI.USERS_BRANCH_ID='" + Form2.branch + "'", conn);
                da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "SOTRUDNIKI");
                dataGridView1.DataSource = ds.Tables[0];
                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не выбраны позиции для экспорта!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);

            //делаем временно неактивным документ
            ExcelApp.Interactive = false;
            ExcelApp.EnableEvents = false;

            ExcelApp.Columns[1].ColumnWidth = 50;
            ExcelApp.Columns[2].ColumnWidth = 38;
            ExcelApp.Columns[3].ColumnWidth = 14;
            ExcelApp.Columns[4].ColumnWidth = 15;
            ExcelApp.Columns[5].ColumnWidth = 18;
            ExcelApp.Columns[6].ColumnWidth = 19;
            ExcelApp.Columns[7].ColumnWidth = 25;
            ExcelApp.Columns[8].ColumnWidth = 7;
            ExcelApp.Columns[9].ColumnWidth = 25;

            ExcelApp.Rows[1].Font.Bold = true; //Установка жирного шрифта на первой строчке
            ExcelApp.Rows[1].Font.Size = 20;
            ExcelApp.Rows[4].Font.Bold = true;


            ExcelApp.Cells[1, 3] = "Движение/Выдача СИЗ";
            ExcelApp.Cells[1, 3].Font.Color = Color.Blue;
            ExcelApp.Cells[2, 1] = Form1.sotrudn_fio;
            ExcelApp.Cells[2, 1].Font.Bold = true;


            ExcelApp.Cells[4, 1] = "ФИО";
            ExcelApp.Cells[4, 2] = "Наименование";
            ExcelApp.Cells[4, 3] = "Ед. измерения";
            ExcelApp.Cells[4, 4] = "Кол-во";
            ExcelApp.Cells[4, 5] = "Период использования";
            ExcelApp.Cells[4, 6] = "Дата выдачи";
            ExcelApp.Cells[4, 7] = "Дата след. выдачи";
            ExcelApp.Cells[4, 8] = "Размер";
            ExcelApp.Cells[4, 9] = "Наименование ИАИС";

            for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
            {
                ExcelApp.Cells[i + 6, 1] = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value);
                ExcelApp.Cells[i + 6, 2] = Convert.ToString(dataGridView1.Rows[i].Cells[1].Value);
                ExcelApp.Cells[i + 6, 3] = Convert.ToString(dataGridView1.Rows[i].Cells[2].Value);
                ExcelApp.Cells[i + 6, 4] = Convert.ToString(dataGridView1.Rows[i].Cells[3].Value);
                ExcelApp.Cells[i + 6, 5] = Convert.ToString(dataGridView1.Rows[i].Cells[4].Value);
                ExcelApp.Cells[i + 6, 6] = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);
                ExcelApp.Cells[i + 6, 7] = Convert.ToString(dataGridView1.Rows[i].Cells[6].Value);
                ExcelApp.Cells[i + 6, 8] = Convert.ToString(dataGridView1.Rows[i].Cells[7].Value);
                ExcelApp.Cells[i + 6, 9] = Convert.ToString(dataGridView1.Rows[i].Cells[8].Value);
            }

            //Показываем ексель
            ExcelApp.Visible = true;

            ExcelApp.Interactive = true;
            ExcelApp.ScreenUpdating = true;
            ExcelApp.UserControl = true;
            
        }

        
    }
}
