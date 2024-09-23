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
    public partial class spr_norns : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");

        public spr_norns()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["PROFESSION"].HeaderText = "Профессия/должность";
            dataGridView1.Columns["PROFESSION"].Width = 500;
            dataGridView1.Columns["KEY_N"].HeaderText = "KEY_N";
            dataGridView1.Columns["KEY_N"].Width = 20;
            dataGridView1.Columns["KEY_N"].Visible = false;

        }
        //////////////////////////////////////////////////////

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview1()
        {
            dataGridView2.Columns["ID"].HeaderText = "ID";
            dataGridView2.Columns["ID"].Width = 40;
            dataGridView2.Columns["ID"].Visible = false;
            dataGridView2.Columns["NORMS_ID"].HeaderText = "NORMS_ID";
            dataGridView2.Columns["NORMS_ID"].Width = 40;
            dataGridView2.Columns["NORMS_ID"].Visible = false;
            dataGridView2.Columns["NAIMEN_OVERALLS"].HeaderText = "Наименование";
            dataGridView2.Columns["NAIMEN_OVERALLS"].Width = 200;
            dataGridView2.Columns["CHARACTERISTIC_OVERALLS"].HeaderText = "Характеристики";
            dataGridView2.Columns["CHARACTERISTIC_OVERALLS"].Width = 200;
            dataGridView2.Columns["ED_IZM"].HeaderText = "Ед. изм.";
            dataGridView2.Columns["ED_IZM"].Width = 50;
            dataGridView2.Columns["KOLVO"].HeaderText = "Кол-во";
            dataGridView2.Columns["KOLVO"].Width = 50;
            dataGridView2.Columns["PERIOD_ISPOLZOVAN"].HeaderText = "Период испол-я";
            dataGridView2.Columns["PERIOD_ISPOLZOVAN"].Width = 100;
            dataGridView2.Columns["PRICE_EDENIC_PLAN"].HeaderText = "Стоим. за ед. план.";
            dataGridView2.Columns["PRICE_EDENIC_PLAN"].Width = 80;
            dataGridView2.Columns["PRICE_KOMPLEKT_PLAN"].HeaderText = "Стоим. за комплект план.";
            dataGridView2.Columns["PRICE_KOMPLEKT_PLAN"].Width = 80;
            dataGridView2.Columns["IAIS_NAIMEN_OVERALLS"].HeaderText = "ИАИС Наименование";
            dataGridView2.Columns["IAIS_NAIMEN_OVERALLS"].Width = 230;
            dataGridView2.Columns["CODE_GROUP_IAIS"].HeaderText = "CODE_GROUP_IAIS";
            dataGridView2.Columns["CODE_GROUP_IAIS"].Width = 40;
            dataGridView2.Columns["CODE_GROUP_IAIS"].Visible = false;
            dataGridView2.Columns["KONTRAGENT_ID"].HeaderText = "KONTRAGENT_ID";
            dataGridView2.Columns["KONTRAGENT_ID"].Width = 10;
            dataGridView2.Columns["KONTRAGENT_ID"].Visible = false;
            dataGridView2.Columns["CHANGED"].HeaderText = "CHANGED";
            dataGridView2.Columns["CHANGED"].Width = 10;
            dataGridView2.Columns["CHANGED"].Visible = false;

        }
        //////////////////////////////////////////////////////

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void spr_norns_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;//Запретить пользователю добавлять строки из грида напрямую
            dataGridView2.AllowUserToAddRows = false;//Запретить пользователю добавлять строки из грида напрямую

            SqlCommand command = new SqlCommand("select * from NORMS order by PROFESSION", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            da.Fill(ds, "NORMS");
            dataGridView1.DataSource = ds.Tables[0];

            fill_gridview();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow.Cells[0].Value != DBNull.Value)
            {
                SqlCommand command = new SqlCommand("select * from NORMS_LIST where NORMS_ID=" + dataGridView1.CurrentRow.Cells[0].Value + " order by NAIMEN_OVERALLS", conn);
                SqlDataAdapter da1 = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da1);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da1.Fill(ds, "NORMS_LIST");
                dataGridView2.DataSource = ds.Tables[0];

                fill_gridview1();
            }
            else
            {
                dataGridView2.DataSource = null;
                dataGridView2.Rows.Clear();
            }
        }

        private void экспортВExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Open();
            }
            catch { }

            SqlCommand command = new SqlCommand("select * from NORMS_LIST where NORMS_ID=" + dataGridView1.CurrentRow.Cells[0].Value + " order by NAIMEN_OVERALLS", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "NORMS_LIST");



            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);

            //делаем временно неактивным документ
            ExcelApp.Interactive = false;
            ExcelApp.EnableEvents = false;

            ExcelApp.Columns[1].ColumnWidth = 50;
            ExcelApp.Columns[2].ColumnWidth = 29;
            ExcelApp.Columns[3].ColumnWidth = 25;
            ExcelApp.Columns[4].ColumnWidth = 15;
            ExcelApp.Columns[5].ColumnWidth = 18;
            ExcelApp.Columns[6].ColumnWidth = 19;
            ExcelApp.Columns[7].ColumnWidth = 25;
            ExcelApp.Columns[8].ColumnWidth = 25;


            ExcelApp.Rows[1].Font.Bold = true; //Установка жирного шрифта на первой строчке
            ExcelApp.Rows[1].Font.Size = 20; //Установка жирного шрифта на первой строчке
            ExcelApp.Rows[4].Font.Bold = true; //Установка жирного шрифта на первой строчке


            ExcelApp.Cells[1, 5] = "Нормы - " + dataGridView1.CurrentRow.Cells[1].Value;
            ExcelApp.Cells[1, 5].Font.Color = Color.Blue;
            ExcelApp.Cells[2, 1] = Form1.sotrudn_fio;
            ExcelApp.Cells[2, 1].Font.Bold = true;


            ExcelApp.Cells[4, 1] = "Наименование";
            ExcelApp.Cells[4, 2] = "Характеристики";
            ExcelApp.Cells[4, 3] = "Ед.изм.";
            ExcelApp.Cells[4, 4] = "Кол-во";
            ExcelApp.Cells[4, 5] = "Период испол-я";
            ExcelApp.Cells[4, 6] = "Стоим. за ед. план.";
            ExcelApp.Cells[4, 7] = "Стоим. за комплект план.";
            ExcelApp.Cells[4, 8] = "Наименование ИАИС";


            for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
            {
                ExcelApp.Cells[i + 6, 1] = Convert.ToString(ds.Tables[0].Rows[i][2]);
                ExcelApp.Cells[i + 6, 2] = Convert.ToString(ds.Tables[0].Rows[i][3]);
                ExcelApp.Cells[i + 6, 3] = Convert.ToString(ds.Tables[0].Rows[i][4]);
                ExcelApp.Cells[i + 6, 4] = Convert.ToString(ds.Tables[0].Rows[i][5]);
                ExcelApp.Cells[i + 6, 5] = Convert.ToString(ds.Tables[0].Rows[i][6]);
                ExcelApp.Cells[i + 6, 6] = Convert.ToString(ds.Tables[0].Rows[i][7]);
                ExcelApp.Cells[i + 6, 7] = Convert.ToString(ds.Tables[0].Rows[i][8]);
                ExcelApp.Cells[i + 6, 8] = Convert.ToString(ds.Tables[0].Rows[i][9]);


            }

            //Показываем ексель
            ExcelApp.Visible = true;

            ExcelApp.Interactive = true;
            ExcelApp.ScreenUpdating = true;
            ExcelApp.UserControl = true;
            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                    SqlCommand command = new SqlCommand("select * from NORMS where PROFESSION like '%"+ textBox2.Text +"%'  order by PROFESSION", conn);
                    SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "NORMS");
                    dataGridView1.DataSource = ds.Tables[0];

                    fill_gridview();             

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
