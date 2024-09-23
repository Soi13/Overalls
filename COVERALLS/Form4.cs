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
using Microsoft.Office.Interop.Excel;

namespace COVERALLS
{
    public partial class Form4 : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
        SqlDataAdapter da;

        public Form4()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["NORMS_ID"].HeaderText = "NORMS_ID";
            dataGridView1.Columns["NORMS_ID"].Width = 40;
            dataGridView1.Columns["NORMS_ID"].Visible = false;
            dataGridView1.Columns["NAIMEN_OVERALLS"].HeaderText = "Наименование";
            dataGridView1.Columns["NAIMEN_OVERALLS"].Width = 200;
            dataGridView1.Columns["CHARACTERISTIC_OVERALLS"].HeaderText = "Характеристики";
            dataGridView1.Columns["CHARACTERISTIC_OVERALLS"].Width = 200;
            dataGridView1.Columns["ED_IZM"].HeaderText = "Ед. изм.";
            dataGridView1.Columns["ED_IZM"].Width = 50;
            dataGridView1.Columns["KOLVO"].HeaderText = "Кол-во";
            dataGridView1.Columns["KOLVO"].Width = 50;
            dataGridView1.Columns["PERIOD_ISPOLZOVAN"].HeaderText = "Период испол-я";
            dataGridView1.Columns["PERIOD_ISPOLZOVAN"].Width = 100;
            dataGridView1.Columns["PRICE_EDENIC_PLAN"].HeaderText = "Стоим. за ед. план.";
            dataGridView1.Columns["PRICE_EDENIC_PLAN"].Width = 100;
            dataGridView1.Columns["PRICE_KOMPLEKT_PLAN"].HeaderText = "Стоим. за комплект план.";
            dataGridView1.Columns["PRICE_KOMPLEKT_PLAN"].Width = 100;
            dataGridView1.Columns["IAIS_NAIMEN_OVERALLS"].HeaderText = "Наименование ИАИС";
            dataGridView1.Columns["IAIS_NAIMEN_OVERALLS"].Width = 200;
            dataGridView1.Columns["CODE_GROUP_IAIS"].HeaderText = "CODE_GROUP_IAIS";
            dataGridView1.Columns["CODE_GROUP_IAIS"].Width = 40;
            dataGridView1.Columns["CODE_GROUP_IAIS"].Visible = false;
            dataGridView1.Columns["KONTRAGENT_ID"].HeaderText = "KONTRAGENT_ID";
            dataGridView1.Columns["KONTRAGENT_ID"].Width = 20;
            dataGridView1.Columns["KONTRAGENT_ID"].Visible = false;

        }

        private void Form4_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;
            label2.Text = "Сотрудник: "+Form1.sotrudn_fio;

            SqlCommand command = new SqlCommand("select norms_list.* from norms_personal,norms_list where norms_personal.norms_id=norms_list.norms_id and norms_personal.sotrudnik_id=" + Form1.sotrudn_id, conn);
            da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "NORMS_LIST");
            dataGridView1.DataSource = ds.Tables[0];

            fill_gridview();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Open();
            }
            catch { }

            SqlCommand command = new SqlCommand("select norms_list.* from norms_personal,norms_list where norms_personal.norms_id=norms_list.norms_id and norms_personal.sotrudnik_id=" + Form1.sotrudn_id, conn);
            da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
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

            ExcelApp.Rows[1].Font.Bold = true; //Установка жирного шрифта на первой строчке
            ExcelApp.Rows[1].Font.Size = 20; //Установка жирного шрифта на первой строчке
            ExcelApp.Rows[4].Font.Bold = true; //Установка жирного шрифта на первой строчке


            ExcelApp.Cells[1, 5] = "Персональные нормы";
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
    }
}
