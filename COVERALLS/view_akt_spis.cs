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

namespace COVERALLS
{
    public partial class view_akt_spis : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");

        SqlDataAdapter da;
        SqlDataAdapter da1;
  
        public view_akt_spis()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["SOTRUDNIK_ID"].HeaderText = "SOTRUDNIK_ID";
            dataGridView1.Columns["SOTRUDNIK_ID"].Width = 40;
            dataGridView1.Columns["SOTRUDNIK_ID"].Visible = false;
            dataGridView1.Columns["NUMBER"].HeaderText = "№п/п";
            dataGridView1.Columns["NUMBER"].Width = 50;
            dataGridView1.Columns["FIO"].HeaderText = "ФИО";
            dataGridView1.Columns["FIO"].Width = 300;
            dataGridView1.Columns["DATE_SPISAN"].HeaderText = "Дата акта";
            dataGridView1.Columns["DATE_SPISAN"].Width = 70;
            dataGridView1.Columns["NOTES"].HeaderText = "Заметки";
            dataGridView1.Columns["NOTES"].Width = 400;
            dataGridView1.Columns["DATE_CREATE"].HeaderText = "DATE_CREATE";
            dataGridView1.Columns["DATE_CREATE"].Width = 20;
            dataGridView1.Columns["DATE_CREATE"].Visible = false;
            dataGridView1.Columns["USER_ID"].HeaderText = "USER_ID";
            dataGridView1.Columns["USER_ID"].Width = 20;
            dataGridView1.Columns["USER_ID"].Visible = false;

        }
        //////////////////////////////////////////////////////

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview1()
        {
            dataGridView2.Columns["ID"].HeaderText = "ID";
            dataGridView2.Columns["ID"].Width = 40;
            dataGridView2.Columns["ID"].Visible = false;
            dataGridView2.Columns["AKT_SPISAN_ID"].HeaderText = "AKT_SPISAN_ID";
            dataGridView2.Columns["AKT_SPISAN_ID"].Width = 40;
            dataGridView2.Columns["AKT_SPISAN_ID"].Visible = false;
            dataGridView2.Columns["SOTRUDNIK_ID"].HeaderText = "SOTRUDNIK_ID";
            dataGridView2.Columns["SOTRUDNIK_ID"].Width = 40;
            dataGridView2.Columns["SOTRUDNIK_ID"].Visible = false;
            dataGridView2.Columns["NORMS_LIST_ID"].HeaderText = "NORMS_LIST_ID";
            dataGridView2.Columns["NORMS_LIST_ID"].Width = 50;
            dataGridView2.Columns["NORMS_LIST_ID"].Visible = false;
            dataGridView2.Columns["NAIMEN_OVERALLS"].HeaderText = "Наименование";
            dataGridView2.Columns["NAIMEN_OVERALLS"].Width = 250;
            dataGridView2.Columns["KOLVO"].HeaderText = "Кол-во";
            dataGridView2.Columns["KOLVO"].Width = 50;
            dataGridView2.Columns["ED_IZM"].HeaderText = "Ед.изм.";
            dataGridView2.Columns["ED_IZM"].Width = 55;
            dataGridView2.Columns["PERIOD_ISPOLZOVAN"].HeaderText = "Период испол-я";
            dataGridView2.Columns["PERIOD_ISPOLZOVAN"].Width = 60;
            dataGridView2.Columns["DATE_VIDACHI"].HeaderText = "Дата выдачи";
            dataGridView2.Columns["DATE_VIDACHI"].Width = 100;
            dataGridView2.Columns["DATE_SLED_VIDACHI"].HeaderText = "Дата след.выдачи";
            dataGridView2.Columns["DATE_SLED_VIDACHI"].Width = 100;
            dataGridView2.Columns["SIZE"].HeaderText = "Размер";
            dataGridView2.Columns["SIZE"].Width = 60;
            dataGridView2.Columns["USER_ID"].HeaderText = "USER_ID";
            dataGridView2.Columns["USER_ID"].Width = 40;
            dataGridView2.Columns["USER_ID"].Visible = false;
            dataGridView2.Columns["DATETIME_CREATE"].HeaderText = "DATETIME_CREATE";
            dataGridView2.Columns["DATETIME_CREATE"].Width = 20;
            dataGridView2.Columns["DATETIME_CREATE"].Visible = false;
            dataGridView2.Columns["IAIS_NAIMEN_OVERALLS"].HeaderText = "Наименование ИАИС";
            dataGridView2.Columns["IAIS_NAIMEN_OVERALLS"].Width = 250;
            dataGridView2.Columns["ID_IAIS"].HeaderText = "ID_IAIS";
            dataGridView2.Columns["ID_IAIS"].Width = 40;
            dataGridView2.Columns["ID_IAIS"].Visible = false;

        }
        //////////////////////////////////////////////////////

        private void view_akt_spis_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;//Запретить пользователю добавлять строки из грида напрямую
            dataGridView2.AllowUserToAddRows = false;//Запретить пользователю добавлять строки из грида напрямую

            SqlCommand command = new SqlCommand("select AKT_SPISAN.ID, AKT_SPISAN.SOTRUDNIK_ID, AKT_SPISAN.NUMBER, SOTRUDNIKI.FIO,AKT_SPISAN.DATE_SPISAN, AKT_SPISAN.NOTES, AKT_SPISAN.DATE_CREATE, AKT_SPISAN.USER_ID from AKT_SPISAN, SOTRUDNIKI where AKT_SPISAN.SOTRUDNIK_ID=SOTRUDNIKI.ID", conn);
            da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "AKT_SPISAN");
            dataGridView1.DataSource = ds.Tables[0];

            fill_gridview();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow.Cells[0].Value != DBNull.Value)
            {
                SqlCommand command = new SqlCommand("select * from AKT_SPISAN_NOMENKL where AKT_SPISAN_ID=" + dataGridView1.CurrentRow.Cells[0].Value + " order by NAIMEN_OVERALLS", conn);
                da1 = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da1);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da1.Fill(ds, "AKT_SPISAN_NOMENKL");
                dataGridView2.DataSource = ds.Tables[0];

                fill_gridview1();
            }
            else
            {
                dataGridView2.DataSource = null;
                dataGridView2.Rows.Clear();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {

                if ((Form2.val == 115) || (Form2.val == 3)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ
                {
                    SqlCommand command = new SqlCommand("select AKT_SPISAN.ID, AKT_SPISAN.SOTRUDNIK_ID, AKT_SPISAN.NUMBER, SOTRUDNIKI.FIO,AKT_SPISAN.DATE_SPISAN, AKT_SPISAN.NOTES, AKT_SPISAN.DATE_CREATE, AKT_SPISAN.USER_ID from AKT_SPISAN, SOTRUDNIKI where AKT_SPISAN.SOTRUDNIK_ID=SOTRUDNIKI.ID and SOTRUDNIKI.FIO like '%"+ textBox2.Text +"%'", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "AKT_SPISAN");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
                else  //Если не Петушков и Скворцов, то показываем исходя из того из каког филиала пользователь производит поиск
                {
                    SqlCommand command = new SqlCommand("select AKT_SPISAN.ID, AKT_SPISAN.SOTRUDNIK_ID, AKT_SPISAN.NUMBER, SOTRUDNIKI.FIO,AKT_SPISAN.DATE_SPISAN, AKT_SPISAN.NOTES, AKT_SPISAN.DATE_CREATE, AKT_SPISAN.USER_ID from AKT_SPISAN, SOTRUDNIKI where AKT_SPISAN.SOTRUDNIK_ID=SOTRUDNIKI.ID and SOTRUDNIKI.FIO like '%" + textBox2.Text + "%' and SOTRUDNIKI.USERS_BRANCH_ID="+Form2.branch, conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "AKT_SPISAN");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }

            }
            else
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Поиск не включен! Включите.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                textBox2.Clear();
                return;
            }
        }
    }
}
