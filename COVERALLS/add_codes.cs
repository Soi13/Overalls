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
    public partial class add_codes : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
      
        public add_codes()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["CODE_IAIS"].HeaderText = "Код ИАИС";
            dataGridView1.Columns["CODE_IAIS"].Width = 80;
            dataGridView1.Columns["NAIMEN_IAIS"].HeaderText = "Наименование";
            dataGridView1.Columns["NAIMEN_IAIS"].Width = 200;
            dataGridView1.Columns["ED_IZM_IAIS"].HeaderText = "Ед.изм";
            dataGridView1.Columns["ED_IZM_IAIS"].Width = 50;
            dataGridView1.Columns["CODE_GROUP_IAIS"].HeaderText = "Код группы";
            dataGridView1.Columns["CODE_GROUP_IAIS"].Width = 60;
            dataGridView1.Columns["GROUP_IAIS"].HeaderText = "Наимен.группы";
            dataGridView1.Columns["GROUP_IAIS"].Width = 200;
            dataGridView1.Columns["CODE_SAP"].HeaderText = "CODE_SAP";
            dataGridView1.Columns["CODE_SAP"].Width = 40;
            dataGridView1.Columns["CODE_SAP"].Visible = false;
            dataGridView1.Columns["NAIMEN_SAP"].HeaderText = "NAIMEN_SAP";
            dataGridView1.Columns["NAIMEN_SAP"].Width = 20;
            dataGridView1.Columns["NAIMEN_SAP"].Visible = false;
            dataGridView1.Columns["ED_IZM_SAP"].HeaderText = "ED_IZM_SAP";
            dataGridView1.Columns["ED_IZM_SAP"].Width = 20;
            dataGridView1.Columns["ED_IZM_SAP"].Visible = false;
            dataGridView1.Columns["CODE_GROUP_SAP"].HeaderText = "CODE_GROUP_SAP";
            dataGridView1.Columns["CODE_GROUP_SAP"].Width = 20;
            dataGridView1.Columns["CODE_GROUP_SAP"].Visible = false;
            dataGridView1.Columns["GROUP_SAP"].HeaderText = "GROUP_SAP";
            dataGridView1.Columns["GROUP_SAP"].Width = 20;
            dataGridView1.Columns["GROUP_SAP"].Visible = false;
        }
        //////////////////////////////////////////////////////

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview1()
        {
            dataGridView2.Columns["ID"].HeaderText = "ID";
            dataGridView2.Columns["ID"].Width = 40;
            dataGridView2.Columns["ID"].Visible = false;
            dataGridView2.Columns["PROFESSION"].HeaderText = "Профессия/должность";
            dataGridView2.Columns["PROFESSION"].Width = 500;
            dataGridView2.Columns["KEY_N"].HeaderText = "KEY_N";
            dataGridView2.Columns["KEY_N"].Width = 20;
            dataGridView2.Columns["KEY_N"].Visible = false;

        }
        //////////////////////////////////////////////////////

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview2()
        {
            dataGridView3.Columns["ID"].HeaderText = "ID";
            dataGridView3.Columns["ID"].Width = 40;
            dataGridView3.Columns["ID"].Visible = false;
            dataGridView3.Columns["NORMS_ID"].HeaderText = "NORMS_ID";
            dataGridView3.Columns["NORMS_ID"].Width = 40;
            dataGridView3.Columns["NORMS_ID"].Visible = false;
            dataGridView3.Columns["NAIMEN_OVERALLS"].HeaderText = "Наименование";
            dataGridView3.Columns["NAIMEN_OVERALLS"].Width = 200;
            dataGridView3.Columns["CHARACTERISTIC_OVERALLS"].HeaderText = "Характеристики";
            dataGridView3.Columns["CHARACTERISTIC_OVERALLS"].Width = 200;
            dataGridView3.Columns["ED_IZM"].HeaderText = "Ед. изм.";
            dataGridView3.Columns["ED_IZM"].Width = 50;
            dataGridView3.Columns["KOLVO"].HeaderText = "Кол-во";
            dataGridView3.Columns["KOLVO"].Width = 50;
            dataGridView3.Columns["KOLVO"].Visible = false;                       
            dataGridView3.Columns["PERIOD_ISPOLZOVAN"].HeaderText = "Период испол-я";
            dataGridView3.Columns["PERIOD_ISPOLZOVAN"].Width = 100;
            dataGridView3.Columns["PERIOD_ISPOLZOVAN"].Visible = false;                                   
            dataGridView3.Columns["PRICE_EDENIC_PLAN"].HeaderText = "Стоим. за ед. план.";
            dataGridView3.Columns["PRICE_EDENIC_PLAN"].Width = 80;
            dataGridView3.Columns["PRICE_EDENIC_PLAN"].Visible = false;                                              
            dataGridView3.Columns["PRICE_KOMPLEKT_PLAN"].HeaderText = "Стоим. за комплект план.";
            dataGridView3.Columns["PRICE_KOMPLEKT_PLAN"].Width = 80;
            dataGridView3.Columns["PRICE_KOMPLEKT_PLAN"].Visible = false;                                                          
            dataGridView3.Columns["IAIS_NAIMEN_OVERALLS"].HeaderText = "ИАИС Наименование";
            dataGridView3.Columns["IAIS_NAIMEN_OVERALLS"].Width = 230;
            dataGridView3.Columns["IAIS_NAIMEN_OVERALLS"].Visible = false;                                                                      
            dataGridView3.Columns["CODE_GROUP_IAIS"].HeaderText = "CODE_4_IAIS";
            dataGridView3.Columns["CODE_GROUP_IAIS"].Width = 50;
            
        }
        //////////////////////////////////////////////////////

        private void add_codes_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;//Запретить пользователю добавлять строки из грида напрямую
            dataGridView2.AllowUserToAddRows = false;//Запретить пользователю добавлять строки из грида напрямую
            dataGridView3.AllowUserToAddRows = false;//Запретить пользователю добавлять строки из грида напрямую

            SqlCommand command = new SqlCommand("select * from NEW_ADDED_POSITION_TO_IAIS", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "NEW_ADDED_POSITION_TO_IAIS");
            dataGridView1.DataSource = ds.Tables[0];
                    
            fill_gridview();

            SqlCommand command1 = new SqlCommand("select * from NORMS order by PROFESSION", conn);
            SqlDataAdapter da1 = new SqlDataAdapter(command1);//Переменная объявлена как глобальная
            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
            DataSet ds1 = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da1.Fill(ds1, "NORMS");
            dataGridView2.DataSource = ds1.Tables[0];

            fill_gridview1();

            statusStrip1.Items[0].Text = "Всего вновь поступивших записей в ИАИС: " + Convert.ToString(ds.Tables[0].Rows.Count);
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow.Cells[0].Value != DBNull.Value)
            {
                SqlCommand command = new SqlCommand("select * from NORMS_LIST where NORMS_ID=" + dataGridView2.CurrentRow.Cells[0].Value + " order by NAIMEN_OVERALLS", conn);
                SqlDataAdapter da1 = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da1);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da1.Fill(ds, "NORMS_LIST");
                dataGridView3.DataSource = ds.Tables[0];

                fill_gridview2();
            }
            else
            {
                dataGridView3.DataSource = null;
                dataGridView3.Rows.Clear();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите установить соответствие?", "Вопрос", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                //////////Выбираем строку в таблице соответствия и временно сохраняем набор кодов для последующего добавления к ним нового кода
                SqlCommand command = new SqlCommand("select * from IAIS_CODES where CODE_GROUP_IAIS=" + dataGridView3.CurrentRow.Cells[10].Value, conn);
                SqlDataAdapter da1 = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da1);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da1.Fill(ds, "IAIS_CODES");
                string codes=ds.Tables[0].Rows[0][2].ToString();
                ///////////////////

                string res_codes = codes + ", " + dataGridView1.CurrentRow.Cells[1].Value;// Добавление выбранного кода к существующим


                /////Update данных в таблице соответствия
                SqlCommand scmd4 = conn.CreateCommand();
                scmd4.CommandText = "update IAIS_CODES set CODE_IAIS='" + res_codes + "' where CODE_GROUP_IAIS=" + dataGridView3.CurrentRow.Cells[10].Value;
                try
                {
                    conn.Open();
                }
                catch { }
                SqlDataReader reader4;
                reader4 = scmd4.ExecuteReader();
                conn.Close();
                //////////////////

                SystemSounds.Beep.Play();
                MessageBox.Show("Соответствие установлено удачно!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);                                            
                
            }
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SqlCommand cmd = conn.CreateCommand();
            cmd.CommandText = "delete from NEW_ADDED_POSITION_TO_IAIS where ID=" + dataGridView1.CurrentRow.Cells[0].Value;
            try
            {
                conn.Open();
            }
            catch { }
            SqlDataReader reader;
            reader = cmd.ExecuteReader();
            conn.Close();         
            ////////////////////////////

            ///////////////////////////
            SqlCommand command = new SqlCommand("select * from NEW_ADDED_POSITION_TO_IAIS", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "NEW_ADDED_POSITION_TO_IAIS");
            dataGridView1.DataSource = ds.Tables[0];

            fill_gridview();
            ///////////////////////////

            statusStrip1.Items[0].Text = "Всего вновь поступивших записей в ИАИС: " + Convert.ToString(ds.Tables[0].Rows.Count);
        }
    }
}
