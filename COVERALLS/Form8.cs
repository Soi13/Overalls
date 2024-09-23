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
    public partial class Form8 : Form
    {
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
        
        SqlDataAdapter da;
        public string NORMS_LIST_ID;

        public Form8()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["SOTRUDNIK_ID"].HeaderText = "NORMS_ID";
            dataGridView1.Columns["SOTRUDNIK_ID"].Width = 40;
            dataGridView1.Columns["SOTRUDNIK_ID"].Visible = false;
            dataGridView1.Columns["NORMS_LIST_ID"].HeaderText = "NORMS_LIST_ID";
            dataGridView1.Columns["NORMS_LIST_ID"].Width = 40;
            dataGridView1.Columns["NORMS_LIST_ID"].Visible = false;
            dataGridView1.Columns["NAIMEN_OVERALLS"].HeaderText = "Наименование";
            dataGridView1.Columns["NAIMEN_OVERALLS"].Width = 200;
            dataGridView1.Columns["ED_IZM"].HeaderText = "Ед. изм.";
            dataGridView1.Columns["ED_IZM"].Width = 50;
            dataGridView1.Columns["KOLVO"].HeaderText = "Кол-во";
            dataGridView1.Columns["KOLVO"].Width = 50;
            dataGridView1.Columns["PERIOD_ISPOLZOVAN"].HeaderText = "Период испол-я";
            dataGridView1.Columns["PERIOD_ISPOLZOVAN"].Width = 100;
            dataGridView1.Columns["DATE_VIDACHI"].HeaderText = "Дата выдачи";
            dataGridView1.Columns["DATE_VIDACHI"].Width = 100;
            dataGridView1.Columns["DATE_SLED_VIDACHI"].HeaderText = "Дата след. выдачи";
            dataGridView1.Columns["DATE_SLED_VIDACHI"].Width = 100;
            dataGridView1.Columns["SIZE"].HeaderText = "Размер";
            dataGridView1.Columns["SIZE"].Width = 60;
            dataGridView1.Columns["USER_ID"].HeaderText = "USER_ID";
            dataGridView1.Columns["USER_ID"].Width = 40;
            dataGridView1.Columns["USER_ID"].Visible = false;
            dataGridView1.Columns["DATETIME_CREATE"].HeaderText = "DATETIME_CREATE";
            dataGridView1.Columns["DATETIME_CREATE"].Width = 40;
            dataGridView1.Columns["DATETIME_CREATE"].Visible = false;
            dataGridView1.Columns["IAIS_NAIMEN_OVERALLS"].HeaderText = "IAIS_NAIMEN_OVERALLS";
            dataGridView1.Columns["IAIS_NAIMEN_OVERALLS"].Width = 40;
            dataGridView1.Columns["IAIS_NAIMEN_OVERALLS"].Visible = false;
            dataGridView1.Columns["ID_IAIS"].HeaderText = "ID_IAIS";
            dataGridView1.Columns["ID_IAIS"].Width = 40;
            dataGridView1.Columns["ID_IAIS"].Visible = false;
            dataGridView1.Columns["CARD_UIN"].HeaderText = "CARD_UIN";
            dataGridView1.Columns["CARD_UIN"].Width = 20;
            dataGridView1.Columns["CARD_UIN"].Visible = false;

        }

        public void fill_gridview1()
        {
            dataGridView2.Columns["ID"].HeaderText = "ID";
            dataGridView2.Columns["ID"].Width = 40;
            dataGridView2.Columns["ID"].Visible = false;
            dataGridView2.Columns["NORMS_ID"].HeaderText = "NORMS_ID";
            dataGridView2.Columns["NORMS_ID"].Width = 40;
            dataGridView2.Columns["NORMS_ID"].Visible = false;
            dataGridView2.Columns["NAIMEN_OVERALLS"].HeaderText = "Наименование";
            dataGridView2.Columns["NAIMEN_OVERALLS"].Width = 400;
            dataGridView2.Columns["ED_IZM"].HeaderText = "Ед. изм.";
            dataGridView2.Columns["ED_IZM"].Width = 50;
            dataGridView2.Columns["ED_IZM"].ReadOnly = true;
            dataGridView2.Columns["KOLVO"].HeaderText = "Кол-во";
            dataGridView2.Columns["KOLVO"].Width = 50;
            dataGridView2.Columns["KOLVO"].ToolTipText = "Это кол-во по норме";
            dataGridView2.Columns["PERIOD_ISPOLZOVAN"].HeaderText = "Период испол-я";
            dataGridView2.Columns["PERIOD_ISPOLZOVAN"].Width = 100;
        }


        private void Form8_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;
            label2.Text = "Сотрудник: " + Form1.sotrudn_fio;
          
            SqlCommand command = new SqlCommand("select * from vidano where SOTRUDNIK_ID=" + Form1.sotrudn_id, conn);
            da = new SqlDataAdapter(command);
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            da.Fill(ds, "VIDANO");
            dataGridView1.DataSource = ds.Tables[0];


            ///////////////Вывод остатка к выдаче по норме

            //Определение нормы сотрудника
            SqlCommand cmd = new SqlCommand("select norms_list.* from norms_personal,norms_list where norms_personal.norms_id=norms_list.norms_id and norms_personal.sotrudnik_id=" + Form1.sotrudn_id, conn);
            da = new SqlDataAdapter(cmd);//Переменная объявлена как глобальная
            SqlCommandBuilder cb_ = new SqlCommandBuilder(da);
            DataSet ds_ = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds_, "NORMS_LIST");
            ///////////////////////////////

            //Определение что уже выдано сотруднику
            SqlCommand cmd1 = new SqlCommand("select sotrudnik_id,norms_list_id,naimen_overalls,ed_izm,SUM(kolvo) as kolvo,period_ispolzovan from vidano where SOTRUDNIK_ID=" + Form1.sotrudn_id + " group by SOTRUDNIK_ID,NORMS_LIST_ID,NAIMEN_OVERALLS,ED_IZM,PERIOD_ISPOLZOVAN", conn);
            da = new SqlDataAdapter(cmd1);
            SqlCommandBuilder cb1 = new SqlCommandBuilder(da);
            DataSet ds1 = new DataSet();
            conn.Close();
            da.Fill(ds1, "VIDANO");
            ///////////////////////////////

            ///////////////////////////////////////////////////
            for (int nr = 0; nr <= ds_.Tables[0].Rows.Count - 1; nr++)
            {
                int aa = 0;
                string ID = Convert.ToString(ds_.Tables[0].Rows[nr][0]);
                string NORMS_ID = Convert.ToString(ds_.Tables[0].Rows[nr][1]);
                string NAIMEN_OVERALLS = Convert.ToString(ds_.Tables[0].Rows[nr][2]);
                string ED_IZM = Convert.ToString(ds_.Tables[0].Rows[nr][4]);
                string KOLVO = Convert.ToString(ds_.Tables[0].Rows[nr][5]);
                string PERIOD_ISPOLZOVAN = Convert.ToString(ds_.Tables[0].Rows[nr][6]);

                for (int vd = 0; vd <= ds1.Tables[0].Rows.Count - 1; vd++)
                {
                    string SOTRUDNIK_ID = Convert.ToString(ds1.Tables[0].Rows[vd][0]);
                    NORMS_LIST_ID = Convert.ToString(ds1.Tables[0].Rows[vd][1]);
                    string NAIMEN_OVERALLS1 = Convert.ToString(ds1.Tables[0].Rows[vd][2]);
                    string ED_IZM1 = Convert.ToString(ds1.Tables[0].Rows[vd][3]);
                    string KOLVO1 = Convert.ToString(ds1.Tables[0].Rows[vd][4]);
                    string PERIOD_ISPOLZOVAN1 = Convert.ToString(ds1.Tables[0].Rows[vd][5]);

                    //Поиск в нормах того, что уже отпущено для вычисления кол-ва остатка по норме (то, что еще можно выдать)
                    if (ID == NORMS_LIST_ID)
                    {
                        int ostatok = Convert.ToInt16(KOLVO) - Convert.ToInt16(KOLVO1); //Вычисление остатка от нормы, возможного к отпуску
                        if (ostatok != 0)//Если остаток по позиции не 0, то выводим его в datagridview
                        {
                            dataGridView2.Rows.Add(ID, NORMS_ID, NAIMEN_OVERALLS, ED_IZM, ostatok, PERIOD_ISPOLZOVAN);
                        }
                        aa++;

                    }
                }

                if (aa == 0)
                {
                    dataGridView2.Rows.Add(ID, NORMS_ID, NAIMEN_OVERALLS, ED_IZM, KOLVO, PERIOD_ISPOLZOVAN);
                }

            }
            /////////////////////////////////////////////////////

            fill_gridview();
            fill_gridview1();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Open();
            }
            catch{}

            SqlCommand command = new SqlCommand("select * from vidano where SOTRUDNIK_ID=" + Form1.sotrudn_id, conn);
            da = new SqlDataAdapter(command);
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            da.Fill(ds, "VIDANO");



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
            ExcelApp.Rows[1].Font.Size = 20; 
            ExcelApp.Rows[4].Font.Bold = true; 


            ExcelApp.Cells[1, 5] = "Личная карточка";
            ExcelApp.Cells[1, 5].Font.Color = Color.Blue;
            ExcelApp.Cells[2, 1] = Form1.sotrudn_fio;
            ExcelApp.Cells[2, 1].Font.Bold = true;


            ExcelApp.Cells[4, 1] = "Наименование";
            ExcelApp.Cells[4, 2] = "Ед.изм.";
            ExcelApp.Cells[4, 3] = "Кол-во";
            ExcelApp.Cells[4, 4] = "Период испол-я";
            ExcelApp.Cells[4, 5] = "Дата выдачи";
            ExcelApp.Cells[4, 6] = "Дата след. выдачи";
            ExcelApp.Cells[4, 7] = "Размер";
            
            for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
            {
                ExcelApp.Cells[i + 6, 1] = Convert.ToString(ds.Tables[0].Rows[i][3]);
                ExcelApp.Cells[i + 6, 2] = Convert.ToString(ds.Tables[0].Rows[i][4]);
                ExcelApp.Cells[i + 6, 3] = Convert.ToString(ds.Tables[0].Rows[i][5]);
                ExcelApp.Cells[i + 6, 4] = Convert.ToString(ds.Tables[0].Rows[i][6]);
                ExcelApp.Cells[i + 6, 5] = Convert.ToString(ds.Tables[0].Rows[i][7]);
                ExcelApp.Cells[i + 6, 6] = Convert.ToString(ds.Tables[0].Rows[i][8]);
                ExcelApp.Cells[i + 6, 7] = Convert.ToString(ds.Tables[0].Rows[i][9]);
            

            }

            //Показываем ексель
            ExcelApp.Visible = true;

            ExcelApp.Interactive = true;
            ExcelApp.ScreenUpdating = true;
            ExcelApp.UserControl = true;
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);

            //делаем временно неактивным документ
            ExcelApp.Interactive = false;
            ExcelApp.EnableEvents = false;

            ExcelApp.Columns[1].ColumnWidth = 50;
            ExcelApp.Columns[2].ColumnWidth = 29;
            ExcelApp.Columns[3].ColumnWidth = 25;
            ExcelApp.Columns[4].ColumnWidth = 15;
           

            ExcelApp.Rows[1].Font.Bold = true; //Установка жирного шрифта на первой строчке
            ExcelApp.Rows[1].Font.Size = 20; //Установка жирного шрифта на первой строчке
            ExcelApp.Rows[4].Font.Bold = true; //Установка жирного шрифта на первой строчке


            ExcelApp.Cells[1, 5] = "Остаток СИЗ к выдаче";
            ExcelApp.Cells[1, 5].Font.Color = Color.Blue;
            ExcelApp.Cells[2, 1] = Form1.sotrudn_fio;
            ExcelApp.Cells[2, 1].Font.Bold = true;


            ExcelApp.Cells[4, 1] = "Наименование";
            ExcelApp.Cells[4, 2] = "Ед.изм.";
            ExcelApp.Cells[4, 3] = "Кол-во";
            ExcelApp.Cells[4, 4] = "Период испол-я";
            
            for (int i = 0; i <= dataGridView2.Rows.Count-1; i++)
            {
                ExcelApp.Cells[i + 6, 1] = Convert.ToString(dataGridView2.Rows[i].Cells[2].Value);
                ExcelApp.Cells[i + 6, 2] = Convert.ToString(dataGridView2.Rows[i].Cells[3].Value);
                ExcelApp.Cells[i + 6, 3] = Convert.ToString(dataGridView2.Rows[i].Cells[4].Value);
                ExcelApp.Cells[i + 6, 4] = Convert.ToString(dataGridView2.Rows[i].Cells[5].Value);
                
            }

            //Показываем ексель
            ExcelApp.Visible = true;

            ExcelApp.Interactive = true;
            ExcelApp.ScreenUpdating = true;
            ExcelApp.UserControl = true;
            
        }

        private void Form8_FormClosed(object sender, FormClosedEventArgs e)
        {
            dataGridView2.Rows.Clear();
        }

    }
}
