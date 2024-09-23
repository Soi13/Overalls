using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace COVERALLS
{
    public partial class zayav_detail : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");

        Boolean filter_state;
        public static string idd;
        public static string naim;
        public static string kol_vo;
        public static string size_cl;
        public static string size_sh;
        public static string height_;

        public zayav_detail()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["GUID"].HeaderText = "GUID";
            dataGridView1.Columns["GUID"].Width = 20;
            dataGridView1.Columns["GUID"].Visible = false;
            dataGridView1.Columns["NAIMENOVAN_COVERALLS"].HeaderText = "Наименование";
            dataGridView1.Columns["NAIMENOVAN_COVERALLS"].Width = 250;
            dataGridView1.Columns["CHARACTERISTIC_OVERALLS"].HeaderText = "Характеристики";
            dataGridView1.Columns["CHARACTERISTIC_OVERALLS"].Width = 250;
            dataGridView1.Columns["DEPARTMENT"].HeaderText = "Подразделение";
            dataGridView1.Columns["DEPARTMENT"].Width = 100;            
            dataGridView1.Columns["BRANCH"].HeaderText = "Филиал";
            dataGridView1.Columns["BRANCH"].Width = 100;
            dataGridView1.Columns["ED_IZM"].HeaderText = "Ед. изм.";
            dataGridView1.Columns["ED_IZM"].Width = 50;
            dataGridView1.Columns["KOL_VO"].HeaderText = "Кол-во";
            dataGridView1.Columns["KOL_VO"].Width = 50;
            dataGridView1.Columns["PRICE_EDINIC_PLAN"].HeaderText = "Цена";
            dataGridView1.Columns["PRICE_EDINIC_PLAN"].Width = 100;
            dataGridView1.Columns["PRICE_KOMPLEKT_PLAN"].HeaderText = "Стоимость";
            dataGridView1.Columns["PRICE_KOMPLEKT_PLAN"].Width = 100;
            dataGridView1.Columns["JAN"].HeaderText = "Январь";
            dataGridView1.Columns["JAN"].Width = 80;
            dataGridView1.Columns["FEB"].HeaderText = "Февраль";
            dataGridView1.Columns["FEB"].Width = 80;
            dataGridView1.Columns["MAR"].HeaderText = "Март";
            dataGridView1.Columns["MAR"].Width = 80;
            dataGridView1.Columns["APR"].HeaderText = "Апрель";
            dataGridView1.Columns["APR"].Width = 80;
            dataGridView1.Columns["MAY"].HeaderText = "Май";
            dataGridView1.Columns["MAY"].Width = 80;
            dataGridView1.Columns["JUN"].HeaderText = "Июнь";
            dataGridView1.Columns["JUN"].Width = 80;
            dataGridView1.Columns["JUL"].HeaderText = "Июль";
            dataGridView1.Columns["JUL"].Width = 80;
            dataGridView1.Columns["AUG"].HeaderText = "Август";
            dataGridView1.Columns["AUG"].Width = 80;
            dataGridView1.Columns["SEP"].HeaderText = "Сентябрь";
            dataGridView1.Columns["SEP"].Width = 80;
            dataGridView1.Columns["OKT"].HeaderText = "Октябрь";
            dataGridView1.Columns["OKT"].Width = 80;
            dataGridView1.Columns["NOV"].HeaderText = "Ноябрь";
            dataGridView1.Columns["NOV"].Width = 80;
            dataGridView1.Columns["DEC"].HeaderText = "Декабрь";
            dataGridView1.Columns["DEC"].Width = 80;
            dataGridView1.Columns["DATETIME_CREATE"].HeaderText = "DATETIME_CREATE";
            dataGridView1.Columns["DATETIME_CREATE"].Width = 20;
            dataGridView1.Columns["DATETIME_CREATE"].Visible = false;
            dataGridView1.Columns["KONTRAGENT"].HeaderText = "Контрагент";
            dataGridView1.Columns["KONTRAGENT"].Width = 200;
            dataGridView1.Columns["FIO"].HeaderText = "ФИО";
            dataGridView1.Columns["FIO"].Width = 200;
            dataGridView1.Columns["SIZE_CLOTHES"].HeaderText = "Размер одежды";
            dataGridView1.Columns["SIZE_CLOTHES"].Width = 50;
            dataGridView1.Columns["SIZE_SHOES"].HeaderText = "Размер обуви";
            dataGridView1.Columns["SIZE_SHOES"].Width = 50;
            dataGridView1.Columns["HEIGHT"].HeaderText = "Рост";
            dataGridView1.Columns["HEIGHT"].Width = 50;
        }
        //////////////////////////////////////////////////////

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview_gr()
        {                     
            dataGridView1.Columns["GUID"].HeaderText = "GUID";
            dataGridView1.Columns["GUID"].Width = 20;
            dataGridView1.Columns["GUID"].Visible = false;
            dataGridView1.Columns["NAIMENOVAN_COVERALLS"].HeaderText = "Наименование";
            dataGridView1.Columns["NAIMENOVAN_COVERALLS"].Width = 400;        
            dataGridView1.Columns["KOLVO"].HeaderText = "Кол-во";
            dataGridView1.Columns["KOLVO"].Width = 50;
            dataGridView1.Columns["PRICE_EDINIC_PLAN"].HeaderText = "Цена";
            dataGridView1.Columns["PRICE_EDINIC_PLAN"].Width = 50;
            dataGridView1.Columns["PRICE_KOMPLEKT_PLAN"].HeaderText = "Стоимость";
            dataGridView1.Columns["PRICE_KOMPLEKT_PLAN"].Width = 50;
            dataGridView1.Columns["JAN"].HeaderText = "Янв";
            dataGridView1.Columns["JAN"].Width = 50;
            dataGridView1.Columns["FEB"].HeaderText = "Фев";
            dataGridView1.Columns["FEB"].Width = 50;
            dataGridView1.Columns["MAR"].HeaderText = "Мар";
            dataGridView1.Columns["MAR"].Width = 50;
            dataGridView1.Columns["APR"].HeaderText = "Апр";
            dataGridView1.Columns["APR"].Width = 50;
            dataGridView1.Columns["MAY"].HeaderText = "Май";
            dataGridView1.Columns["MAY"].Width = 50;
            dataGridView1.Columns["JUN"].HeaderText = "Июн";
            dataGridView1.Columns["JUN"].Width = 50;
            dataGridView1.Columns["JUL"].HeaderText = "Июл";
            dataGridView1.Columns["JUL"].Width = 50;
            dataGridView1.Columns["AUG"].HeaderText = "Авг";
            dataGridView1.Columns["AUG"].Width = 50;
            dataGridView1.Columns["SEP"].HeaderText = "Сен";
            dataGridView1.Columns["SEP"].Width = 50;
            dataGridView1.Columns["OKT"].HeaderText = "Окт";
            dataGridView1.Columns["OKT"].Width = 50;
            dataGridView1.Columns["NOV"].HeaderText = "Ноя";
            dataGridView1.Columns["NOV"].Width = 50;
            dataGridView1.Columns["DEC"].HeaderText = "Дек";
            dataGridView1.Columns["DEC"].Width = 50;         
            dataGridView1.Columns["SIZE_CLOTHES"].HeaderText = "Размер одежды";
            dataGridView1.Columns["SIZE_CLOTHES"].Width = 50;
            dataGridView1.Columns["SIZE_SHOES"].HeaderText = "Размер обуви";
            dataGridView1.Columns["SIZE_SHOES"].Width = 50;
            dataGridView1.Columns["HEIGHT"].HeaderText = "Рост";
            dataGridView1.Columns["HEIGHT"].Width = 50;
        }
        //////////////////////////////////////////////////////

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview_group_naimenovan()
        {            
            dataGridView1.Columns["GUID"].HeaderText = "GUID";
            dataGridView1.Columns["GUID"].Width = 20;
            dataGridView1.Columns["GUID"].Visible = false;
            dataGridView1.Columns["NAIMENOVAN_COVERALLS"].HeaderText = "Наименование";
            dataGridView1.Columns["NAIMENOVAN_COVERALLS"].Width = 350;
            dataGridView1.Columns["COUNT_SIZ"].HeaderText = "Кол-во по заявке";
            dataGridView1.Columns["COUNT_SIZ"].Width = 100;
            
        }
        //////////////////////////////////////////////////////

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview_group_department()
        {
            dataGridView1.Columns["GUID"].HeaderText = "GUID";
            dataGridView1.Columns["GUID"].Width = 20;
            dataGridView1.Columns["GUID"].Visible = false;
            dataGridView1.Columns["DEPARTMENT"].HeaderText = "Подразделение/отдел/цех";
            dataGridView1.Columns["DEPARTMENT"].Width = 350;
            dataGridView1.Columns["COUNT_SIZ"].HeaderText = "Кол-во по заявке";
            dataGridView1.Columns["COUNT_SIZ"].Width = 100;

        }
        //////////////////////////////////////////////////////

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview_group_branch()
        {
            dataGridView1.Columns["GUID"].HeaderText = "GUID";
            dataGridView1.Columns["GUID"].Width = 20;
            dataGridView1.Columns["GUID"].Visible = false;
            dataGridView1.Columns["BRANCH"].HeaderText = "Филиал";
            dataGridView1.Columns["BRANCH"].Width = 350;
            dataGridView1.Columns["COUNT_SIZ"].HeaderText = "Кол-во по заявке";
            dataGridView1.Columns["COUNT_SIZ"].Width = 100;

        }
        //////////////////////////////////////////////////////

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview_group_fio()
        {
            dataGridView1.Columns["GUID"].HeaderText = "GUID";
            dataGridView1.Columns["GUID"].Width = 20;
            dataGridView1.Columns["GUID"].Visible = false;
            dataGridView1.Columns["FIO"].HeaderText = "ФИО";
            dataGridView1.Columns["FIO"].Width = 350;
            dataGridView1.Columns["COUNT_SIZ"].HeaderText = "Кол-во по заявке";
            dataGridView1.Columns["COUNT_SIZ"].Width = 100;

        }
        //////////////////////////////////////////////////////

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview_group_ka()
        {
            dataGridView1.Columns["GUID"].HeaderText = "GUID";
            dataGridView1.Columns["GUID"].Width = 20;
            dataGridView1.Columns["GUID"].Visible = false;
            dataGridView1.Columns["KONTRAGENT"].HeaderText = "Контрагент";
            dataGridView1.Columns["KONTRAGENT"].Width = 350;
            dataGridView1.Columns["COUNT_SIZ"].HeaderText = "Кол-во по заявке";
            dataGridView1.Columns["COUNT_SIZ"].Width = 100;

        }
        //////////////////////////////////////////////////////

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview_group_height()
        {
            dataGridView1.Columns["GUID"].HeaderText = "GUID";
            dataGridView1.Columns["GUID"].Width = 20;
            dataGridView1.Columns["GUID"].Visible = false;
            dataGridView1.Columns["HEIGHT"].HeaderText = "Рост";
            dataGridView1.Columns["HEIGHT"].Width = 350;
            dataGridView1.Columns["COUNT_SIZ"].HeaderText = "Кол-во по заявке";
            dataGridView1.Columns["COUNT_SIZ"].Width = 100;

        }
        //////////////////////////////////////////////////////

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview_group_size_shoes()
        {
            dataGridView1.Columns["GUID"].HeaderText = "GUID";
            dataGridView1.Columns["GUID"].Width = 20;
            dataGridView1.Columns["GUID"].Visible = false;
            dataGridView1.Columns["SIZE_SHOES"].HeaderText = "Размер обуви";
            dataGridView1.Columns["SIZE_SHOES"].Width = 350;
            dataGridView1.Columns["COUNT_SIZ"].HeaderText = "Кол-во по заявке";
            dataGridView1.Columns["COUNT_SIZ"].Width = 100;

        }
        //////////////////////////////////////////////////////

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview_group_size_clothes()
        {
            dataGridView1.Columns["GUID"].HeaderText = "GUID";
            dataGridView1.Columns["GUID"].Width = 20;
            dataGridView1.Columns["GUID"].Visible = false;
            dataGridView1.Columns["SIZE_CLOTHES"].HeaderText = "Размер одежды";
            dataGridView1.Columns["SIZE_CLOTHES"].Width = 350;
            dataGridView1.Columns["COUNT_SIZ"].HeaderText = "Кол-во по заявке";
            dataGridView1.Columns["COUNT_SIZ"].Width = 100;

        }
        //////////////////////////////////////////////////////

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview_group_torgi()
        {
            dataGridView1.Columns["GUID"].HeaderText = "GUID";
            dataGridView1.Columns["GUID"].Width = 20;
            dataGridView1.Columns["GUID"].Visible = false;
            dataGridView1.Columns["NAIMENOVAN_COVERALLS"].HeaderText = "Наименование";
            dataGridView1.Columns["NAIMENOVAN_COVERALLS"].Width = 350;
            dataGridView1.Columns["COUNT_SIZ"].HeaderText = "Кол-во по заявке";
            dataGridView1.Columns["COUNT_SIZ"].Width = 100;

        }
        //////////////////////////////////////////////////////

        private void zayav_detail_Load(object sender, EventArgs e)
        {
            filter_state = false;
            label3.Text = "Режим фильтра отключен";
            
            //Запрос на вывод содержания заявки
            SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, SUM(KOL_VO) as KOLVO, PRICE_EDINIC_PLAN, sum(PRICE_KOMPLEKT_PLAN) as PRICE_KOMPLEKT_PLAN, sum(JAN) as JAN, sum(FEB) as FEB, sum(MAR) as MAR, sum(APR) as APR, sum(MAY) as MAY, sum(JUN) as JUN, sum(JUL) as JUL, sum(AUG) as AUG, sum(SEP) as SEP, sum(OKT) as OKT, sum(NOV) as NOV, sum(DEC) as DEC, SIZE_CLOTHES, SIZE_SHOES, HEIGHT from ZAYAVKA_BODY where GUID=@gd group by GUID, NAIMENOVAN_COVERALLS, PRICE_EDINIC_PLAN, SIZE_CLOTHES, SIZE_SHOES, HEIGHT order by NAIMENOVAN_COVERALLS", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            command.Parameters.AddWithValue("gd",view_zayavki.gd);
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "ZAYAVKA_BODY");
            dataGridView1.DataSource = ds.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);
                      
            fill_gridview_gr();                  
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            //Фильтр по ИА
            if (radioButton1.Checked == true)
            {
                filter_state = true;
                label3.Text = "Режим фильтра включен";

                 //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, SUM(KOL_VO) as KOLVO, PRICE_EDINIC_PLAN, sum(PRICE_KOMPLEKT_PLAN) as PRICE_KOMPLEKT_PLAN, sum(JAN) as JAN, sum(FEB) as FEB, sum(MAR) as MAR, sum(APR) as APR, sum(MAY) as MAY, sum(JUN) as JUN, sum(JUL) as JUL, sum(AUG) as AUG, sum(SEP) as SEP, sum(OKT) as OKT, sum(NOV) as NOV, sum(DEC) as DEC, SIZE_CLOTHES, SIZE_SHOES,HEIGHT from ZAYAVKA_BODY where GUID=@gd and BRANCH=@br group by GUID, NAIMENOVAN_COVERALLS, PRICE_EDINIC_PLAN, SIZE_CLOTHES, SIZE_SHOES, HEIGHT order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", view_zayavki.gd);
                command.Parameters.AddWithValue("br", "Исполнительный аппарат");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "ZAYAVKA_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_gr();     
            }            
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            filter_state = true;
            label3.Text = "Режим фильтра включен";

            //Фильтр по Абакану
            if (radioButton2.Checked == true)
            {            
                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, SUM(KOL_VO) as KOLVO, PRICE_EDINIC_PLAN, sum(PRICE_KOMPLEKT_PLAN) as PRICE_KOMPLEKT_PLAN, sum(JAN) as JAN, sum(FEB) as FEB, sum(MAR) as MAR, sum(APR) as APR, sum(MAY) as MAY, sum(JUN) as JUN, sum(JUL) as JUL, sum(AUG) as AUG, sum(SEP) as SEP, sum(OKT) as OKT, sum(NOV) as NOV, sum(DEC) as DEC, SIZE_CLOTHES, SIZE_SHOES,HEIGHT from ZAYAVKA_BODY where GUID=@gd and BRANCH=@br group by GUID, NAIMENOVAN_COVERALLS, PRICE_EDINIC_PLAN, SIZE_CLOTHES, SIZE_SHOES, HEIGHT order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", view_zayavki.gd);
                command.Parameters.AddWithValue("br", "Абаканский филиал ОАО \"СибЭР\"");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "ZAYAVKA_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_gr();                
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            filter_state = true;
            label3.Text = "Режим фильтра включен";

            //Фильтр по Барнаулу
            if (radioButton3.Checked == true)
            {
                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, SUM(KOL_VO) as KOLVO, PRICE_EDINIC_PLAN, sum(PRICE_KOMPLEKT_PLAN) as PRICE_KOMPLEKT_PLAN, sum(JAN) as JAN, sum(FEB) as FEB, sum(MAR) as MAR, sum(APR) as APR, sum(MAY) as MAY, sum(JUN) as JUN, sum(JUL) as JUL, sum(AUG) as AUG, sum(SEP) as SEP, sum(OKT) as OKT, sum(NOV) as NOV, sum(DEC) as DEC, SIZE_CLOTHES, SIZE_SHOES,HEIGHT from ZAYAVKA_BODY where GUID=@gd and BRANCH=@br group by GUID, NAIMENOVAN_COVERALLS, PRICE_EDINIC_PLAN, SIZE_CLOTHES, SIZE_SHOES, HEIGHT order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", view_zayavki.gd);
                command.Parameters.AddWithValue("br", "Барнаульский филиал ОАО \"СибЭР\"");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "ZAYAVKA_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_gr();         
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            filter_state = true;
            label3.Text = "Режим фильтра включен";

            //Фильтр по Кемерово
            if (radioButton4.Checked == true)
            {
                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, SUM(KOL_VO) as KOLVO, PRICE_EDINIC_PLAN, sum(PRICE_KOMPLEKT_PLAN) as PRICE_KOMPLEKT_PLAN, sum(JAN) as JAN, sum(FEB) as FEB, sum(MAR) as MAR, sum(APR) as APR, sum(MAY) as MAY, sum(JUN) as JUN, sum(JUL) as JUL, sum(AUG) as AUG, sum(SEP) as SEP, sum(OKT) as OKT, sum(NOV) as NOV, sum(DEC) as DEC, SIZE_CLOTHES, SIZE_SHOES,HEIGHT from ZAYAVKA_BODY where GUID=@gd and BRANCH=@br group by GUID, NAIMENOVAN_COVERALLS, PRICE_EDINIC_PLAN, SIZE_CLOTHES, SIZE_SHOES, HEIGHT order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", view_zayavki.gd);
                command.Parameters.AddWithValue("br", "Кемеровский филиал ОАО \"СибЭР\"");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "ZAYAVKA_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_gr();         
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            filter_state = true;
            label3.Text = "Режим фильтра включен";

            //Фильтр по Красноярску
            if (radioButton5.Checked == true)
            {
                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, SUM(KOL_VO) as KOLVO, PRICE_EDINIC_PLAN, sum(PRICE_KOMPLEKT_PLAN) as PRICE_KOMPLEKT_PLAN, sum(JAN) as JAN, sum(FEB) as FEB, sum(MAR) as MAR, sum(APR) as APR, sum(MAY) as MAY, sum(JUN) as JUN, sum(JUL) as JUL, sum(AUG) as AUG, sum(SEP) as SEP, sum(OKT) as OKT, sum(NOV) as NOV, sum(DEC) as DEC, SIZE_CLOTHES, SIZE_SHOES,HEIGHT from ZAYAVKA_BODY where GUID=@gd and BRANCH=@br group by GUID, NAIMENOVAN_COVERALLS, PRICE_EDINIC_PLAN, SIZE_CLOTHES, SIZE_SHOES, HEIGHT order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", view_zayavki.gd);
                command.Parameters.AddWithValue("br", "Красноярский филиал ОАО \"СибЭР\"");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "ZAYAVKA_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_gr();         
            }
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            filter_state = false;
            label3.Text = "Режим фильтра отключен";

            //Запрос на вывод содержания заявки
            SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, SUM(KOL_VO) as KOLVO, PRICE_EDINIC_PLAN, sum(PRICE_KOMPLEKT_PLAN) as PRICE_KOMPLEKT_PLAN, sum(JAN) as JAN, sum(FEB) as FEB, sum(MAR) as MAR, sum(APR) as APR, sum(MAY) as MAY, sum(JUN) as JUN, sum(JUL) as JUL, sum(AUG) as AUG, sum(SEP) as SEP, sum(OKT) as OKT, sum(NOV) as NOV, sum(DEC) as DEC, SIZE_CLOTHES, SIZE_SHOES,HEIGHT from ZAYAVKA_BODY where GUID=@gd group by GUID, NAIMENOVAN_COVERALLS, PRICE_EDINIC_PLAN, SIZE_CLOTHES, SIZE_SHOES, HEIGHT order by NAIMENOVAN_COVERALLS", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            command.Parameters.AddWithValue("gd", view_zayavki.gd);
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "ZAYAVKA_BODY");
            dataGridView1.DataSource = ds.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

            fill_gridview_gr();              
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //Группировка по наименованию
            if (checkBox1.Checked == true) 
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;                

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, COUNT(*) as COUNT_SIZ from ZAYAVKA_BODY where GUID=@gd group by GUID, NAIMENOVAN_COVERALLS order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", view_zayavki.gd);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "ZAYAVKA_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_naimenovan();
                
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            //Группировка по подразделению
            if (checkBox2.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;                

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, DEPARTMENT, COUNT(*) as COUNT_SIZ from ZAYAVKA_BODY where GUID=@gd group by GUID, DEPARTMENT order by DEPARTMENT", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", view_zayavki.gd);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "ZAYAVKA_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_department();

            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            //Группировка по филиалу
            if (checkBox3.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;                

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, BRANCH, COUNT(*) as COUNT_SIZ from ZAYAVKA_BODY where GUID=@gd group by GUID, BRANCH order by BRANCH", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", view_zayavki.gd);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "ZAYAVKA_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_branch();

            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            //Группировка по ФИО
            if (checkBox4.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;                

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, FIO, COUNT(*) as COUNT_SIZ from ZAYAVKA_BODY where GUID=@gd group by GUID, FIO order by FIO", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", view_zayavki.gd);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "ZAYAVKA_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_fio();
                              
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            //Группировка по КА
            if (checkBox5.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;                

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, KONTRAGENT, COUNT(*) as COUNT_SIZ from ZAYAVKA_BODY where GUID=@gd group by GUID, KONTRAGENT order by KONTRAGENT", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", view_zayavki.gd);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "ZAYAVKA_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_ka();

            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            //Группировка по росту
            if (checkBox6.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;                

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, HEIGHT, COUNT(*) as COUNT_SIZ from ZAYAVKA_BODY where GUID=@gd group by GUID, HEIGHT order by HEIGHT", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", view_zayavki.gd);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "ZAYAVKA_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_height();

            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            //Группировка по размеру обуви
            if (checkBox7.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;                

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, SIZE_SHOES, COUNT(*) as COUNT_SIZ from ZAYAVKA_BODY where GUID=@gd group by GUID, SIZE_SHOES order by SIZE_SHOES", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", view_zayavki.gd);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "ZAYAVKA_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_size_shoes();
                            
            }
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            //Группировка по размеру одежды
            if (checkBox8.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox9.Checked = false;
                

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, SIZE_CLOTHES, COUNT(*) as COUNT_SIZ from ZAYAVKA_BODY where GUID=@gd group by GUID, SIZE_CLOTHES order by SIZE_CLOTHES", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", view_zayavki.gd);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "ZAYAVKA_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_size_clothes();
                              

            }
        }

        private void показатьОстаткиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            skl_ostatki skl_ostatki = new skl_ostatki();
            skl_ostatki.Show();
        }

        private void привязатьОстаткиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            naim = dataGridView1.CurrentRow.Cells[2].ToString();
            kol_vo = dataGridView1.CurrentRow.Cells[3].ToString();
            size_cl = dataGridView1.CurrentRow.Cells[4].ToString();
            size_sh = dataGridView1.CurrentRow.Cells[5].ToString();
            height_ = dataGridView1.CurrentRow.Cells[6].ToString();
            
            select_skl_ostatki select_skl_ostatki = new select_skl_ostatki();
            select_skl_ostatki.ShowDialog();
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            //Группировка по размеру одежды
            if (checkBox9.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;


                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, COUNT(*) as COUNT_SIZ from ZAYAVKA_BODY where GUID=@gd group by GUID, NAIMENOVAN_COVERALLS order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", view_zayavki.gd);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "ZAYAVKA_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_torgi();

            }
        }
    }
}
