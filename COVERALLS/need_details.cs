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
    public partial class need_details : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");

        Boolean filter_state;
        public static string idd;
        public static string naim;
        public static string kol_vo;
        public static string size_cl;
        public static string size_sh;
        public static string height_;

        public need_details()
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
            dataGridView1.Columns["PRICE_EDINIC_PLAN"].HeaderText = "Цена плановая";
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
            dataGridView1.Columns["CHARACTERISTIC_OVERALLS"].HeaderText = "Характеристики";
            dataGridView1.Columns["CHARACTERISTIC_OVERALLS"].Width = 150;
            dataGridView1.Columns["DEPARTMENT"].HeaderText = "Подразделение";
            dataGridView1.Columns["DEPARTMENT"].Width = 150;
            dataGridView1.Columns["BRANCH"].HeaderText = "Филиал";
            dataGridView1.Columns["BRANCH"].Width = 150;
            dataGridView1.Columns["1st_kvartal"].HeaderText = "1 квартал";
            dataGridView1.Columns["1st_kvartal"].Width = 100;
            dataGridView1.Columns["2st_kvartal"].HeaderText = "2 квартал";
            dataGridView1.Columns["2st_kvartal"].Width = 100;
            dataGridView1.Columns["3st_kvartal"].HeaderText = "3 квартал";
            dataGridView1.Columns["3st_kvartal"].Width = 100;
            dataGridView1.Columns["4st_kvartal"].HeaderText = "4 квартал";
            dataGridView1.Columns["4st_kvartal"].Width = 100;
            dataGridView1.Columns["OFFERED_NAIM_COVERALL_KA"].HeaderText = "Предлагаемое наименование";
            dataGridView1.Columns["OFFERED_NAIM_COVERALL_KA"].Width = 100;
            dataGridView1.Columns["NAIM_KA"].HeaderText = "Наименование КА";
            dataGridView1.Columns["NAIM_KA"].Width = 100;
            dataGridView1.Columns["PRICE_WITHOUT_NDS"].HeaderText = "Цена, без НДС";
            dataGridView1.Columns["PRICE_WITHOUT_NDS"].Width = 100;
            dataGridView1.Columns["VALUE_WITHOUT_NDS"].HeaderText = "Стоимость";
            dataGridView1.Columns["VALUE_WITHOUT_NDS"].Width = 100;
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
            dataGridView1.Columns["SUM_PR"].HeaderText = "Сумма";
            dataGridView1.Columns["SUM_PR"].Width = 100;

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
            dataGridView1.Columns["SUM_PR"].HeaderText = "Сумма";
            dataGridView1.Columns["SUM_PR"].Width = 100;

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
            dataGridView1.Columns["SUM_PR"].HeaderText = "Стоимость по заявке";
            dataGridView1.Columns["SUM_PR"].Width = 100;

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
            dataGridView1.Columns["SUM_PR"].HeaderText = "Стоимость по заявке";
            dataGridView1.Columns["SUM_PR"].Width = 100;

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
            dataGridView1.Columns["SUM_PR"].HeaderText = "Кол-во по заявке";
            dataGridView1.Columns["SUM_PR"].Width = 100;

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
            dataGridView1.Columns["NAIMENOVAN_COVERALLS"].Width = 400;
            dataGridView1.Columns["KOLVO"].HeaderText = "Кол-во";
            dataGridView1.Columns["KOLVO"].Width = 50;
            dataGridView1.Columns["PRICE_EDINIC_PLAN"].HeaderText = "Цена плановая";
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
            dataGridView1.Columns["CHARACTERISTIC_OVERALLS"].HeaderText = "Характеристики";
            dataGridView1.Columns["CHARACTERISTIC_OVERALLS"].Width = 150;
            dataGridView1.Columns["DEPARTMENT"].HeaderText = "Подразделение";
            dataGridView1.Columns["DEPARTMENT"].Width = 150;
            dataGridView1.Columns["BRANCH"].HeaderText = "Филиал";
            dataGridView1.Columns["BRANCH"].Width = 150;
            dataGridView1.Columns["1st_kvartal"].HeaderText = "1 квартал";
            dataGridView1.Columns["1st_kvartal"].Width = 100;
            dataGridView1.Columns["2st_kvartal"].HeaderText = "2 квартал";
            dataGridView1.Columns["2st_kvartal"].Width = 100;
            dataGridView1.Columns["3st_kvartal"].HeaderText = "3 квартал";
            dataGridView1.Columns["3st_kvartal"].Width = 100;
            dataGridView1.Columns["4st_kvartal"].HeaderText = "4 квартал";
            dataGridView1.Columns["4st_kvartal"].Width = 100;

        }
        //////////////////////////////////////////////////////

        private void need_details_Load(object sender, EventArgs e)
        {
            filter_state = false;
            label3.Text = "Режим фильтра отключен";

            //Запрос на вывод содержания заявки
            SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, SIZE_CLOTHES, HEIGHT, SIZE_SHOES, DEPARTMENT, BRANCH, SUM(KOL_VO) as KOLVO, PRICE_EDINIC_PLAN, sum(PRICE_KOMPLEKT_PLAN) as PRICE_KOMPLEKT_PLAN, sum(JAN) as JAN, sum(FEB) as FEB, sum(MAR) as MAR, sum(JAN+FEB+MAR) as '1st_kvartal', sum(APR) as APR, sum(MAY) as MAY, sum(JUN) as JUN, sum(APR+MAY+JUN) as '2st_kvartal', sum(JUL) as JUL, sum(AUG) as AUG, sum(SEP) as SEP, sum(JUL+AUG+SEP) as '3st_kvartal', sum(OKT) as OKT, sum(NOV) as NOV, sum(DEC) as DEC, sum(OKT+NOV+DEC) as '4st_kvartal', OFFERED_NAIM_COVERALL_KA, NAIM_KA, PRICE_WITHOUT_NDS, VALUE_WITHOUT_NDS from NEED_BODY where GUID=@gd group by GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, PRICE_EDINIC_PLAN, SIZE_CLOTHES, SIZE_SHOES, HEIGHT, OFFERED_NAIM_COVERALL_KA, NAIM_KA, PRICE_WITHOUT_NDS, VALUE_WITHOUT_NDS order by NAIMENOVAN_COVERALLS", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            command.Parameters.AddWithValue("gd", need.gd);
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "NEED_BODY");
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
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, SIZE_CLOTHES, HEIGHT, SIZE_SHOES, DEPARTMENT, BRANCH, SUM(KOL_VO) as KOLVO, PRICE_EDINIC_PLAN, sum(PRICE_KOMPLEKT_PLAN) as PRICE_KOMPLEKT_PLAN, sum(JAN) as JAN, sum(FEB) as FEB, sum(MAR) as MAR, sum(JAN+FEB+MAR) as '1st_kvartal', sum(APR) as APR, sum(MAY) as MAY, sum(JUN) as JUN, sum(APR+MAY+JUN) as '2st_kvartal', sum(JUL) as JUL, sum(AUG) as AUG, sum(SEP) as SEP, sum(JUL+AUG+SEP) as '3st_kvartal', sum(OKT) as OKT, sum(NOV) as NOV, sum(DEC) as DEC, sum(OKT+NOV+DEC) as '4st_kvartal' from NEED_BODY where GUID=@gd and BRANCH=@br group by GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, PRICE_EDINIC_PLAN, SIZE_CLOTHES, SIZE_SHOES, HEIGHT order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                command.Parameters.AddWithValue("br", "Исполнительный аппарат");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
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
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, SIZE_CLOTHES, HEIGHT, SIZE_SHOES, DEPARTMENT, BRANCH, SUM(KOL_VO) as KOLVO, PRICE_EDINIC_PLAN, sum(PRICE_KOMPLEKT_PLAN) as PRICE_KOMPLEKT_PLAN, sum(JAN) as JAN, sum(FEB) as FEB, sum(MAR) as MAR, sum(JAN+FEB+MAR) as '1st_kvartal', sum(APR) as APR, sum(MAY) as MAY, sum(JUN) as JUN, sum(APR+MAY+JUN) as '2st_kvartal', sum(JUL) as JUL, sum(AUG) as AUG, sum(SEP) as SEP, sum(JUL+AUG+SEP) as '3st_kvartal', sum(OKT) as OKT, sum(NOV) as NOV, sum(DEC) as DEC, sum(OKT+NOV+DEC) as '4st_kvartal' from NEED_BODY where GUID=@gd and BRANCH=@br group by GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, PRICE_EDINIC_PLAN, SIZE_CLOTHES, SIZE_SHOES, HEIGHT order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                command.Parameters.AddWithValue("br", "Абаканский филиал ОАО \"СибЭР\"");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
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
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, SIZE_CLOTHES, HEIGHT, SIZE_SHOES, DEPARTMENT, BRANCH, SUM(KOL_VO) as KOLVO, PRICE_EDINIC_PLAN, sum(PRICE_KOMPLEKT_PLAN) as PRICE_KOMPLEKT_PLAN, sum(JAN) as JAN, sum(FEB) as FEB, sum(MAR) as MAR, sum(JAN+FEB+MAR) as '1st_kvartal', sum(APR) as APR, sum(MAY) as MAY, sum(JUN) as JUN, sum(APR+MAY+JUN) as '2st_kvartal', sum(JUL) as JUL, sum(AUG) as AUG, sum(SEP) as SEP, sum(JUL+AUG+SEP) as '3st_kvartal', sum(OKT) as OKT, sum(NOV) as NOV, sum(DEC) as DEC, sum(OKT+NOV+DEC) as '4st_kvartal' from NEED_BODY where GUID=@gd and BRANCH=@br group by GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, SIZE_CLOTHES, SIZE_SHOES, HEIGHT, DEPARTMENT, BRANCH, PRICE_EDINIC_PLAN, SIZE_CLOTHES, SIZE_SHOES, HEIGHT order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                command.Parameters.AddWithValue("br", "Барнаульский филиал ОАО \"СибЭР\"");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
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
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, SIZE_CLOTHES, HEIGHT, SIZE_SHOES, DEPARTMENT, BRANCH, SUM(KOL_VO) as KOLVO, PRICE_EDINIC_PLAN, sum(PRICE_KOMPLEKT_PLAN) as PRICE_KOMPLEKT_PLAN, sum(JAN) as JAN, sum(FEB) as FEB, sum(MAR) as MAR, sum(JAN+FEB+MAR) as '1st_kvartal', sum(APR) as APR, sum(MAY) as MAY, sum(JUN) as JUN, sum(APR+MAY+JUN) as '2st_kvartal', sum(JUL) as JUL, sum(AUG) as AUG, sum(SEP) as SEP, sum(JUL+AUG+SEP) as '3st_kvartal', sum(OKT) as OKT, sum(NOV) as NOV, sum(DEC) as DEC, sum(OKT+NOV+DEC) as '4st_kvartal' from NEED_BODY where GUID=@gd and BRANCH=@br group by GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, SIZE_CLOTHES, SIZE_SHOES, HEIGHT, DEPARTMENT, BRANCH, PRICE_EDINIC_PLAN, SIZE_CLOTHES, SIZE_SHOES, HEIGHT order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                command.Parameters.AddWithValue("br", "Кемеровский филиал ОАО \"СибЭР\"");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
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
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, SIZE_CLOTHES, HEIGHT, SIZE_SHOES, DEPARTMENT, BRANCH, SUM(KOL_VO) as KOLVO, PRICE_EDINIC_PLAN, sum(PRICE_KOMPLEKT_PLAN) as PRICE_KOMPLEKT_PLAN, sum(JAN) as JAN, sum(FEB) as FEB, sum(MAR) as MAR, sum(JAN+FEB+MAR) as '1st_kvartal', sum(APR) as APR, sum(MAY) as MAY, sum(JUN) as JUN, sum(APR+MAY+JUN) as '2st_kvartal', sum(JUL) as JUL, sum(AUG) as AUG, sum(SEP) as SEP, sum(JUL+AUG+SEP) as '3st_kvartal', sum(OKT) as OKT, sum(NOV) as NOV, sum(DEC) as DEC, sum(OKT+NOV+DEC) as '4st_kvartal' from NEED_BODY where GUID=@gd and BRANCH=@br group by GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, SIZE_CLOTHES, SIZE_SHOES, HEIGHT, DEPARTMENT, BRANCH, PRICE_EDINIC_PLAN, SIZE_CLOTHES, SIZE_SHOES, HEIGHT order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                command.Parameters.AddWithValue("br", "Красноярский филиал ОАО \"СибЭР\"");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
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
            SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, SIZE_CLOTHES, HEIGHT, SIZE_SHOES, DEPARTMENT, BRANCH, SUM(KOL_VO) as KOLVO, PRICE_EDINIC_PLAN, sum(PRICE_KOMPLEKT_PLAN) as PRICE_KOMPLEKT_PLAN, sum(JAN) as JAN, sum(FEB) as FEB, sum(MAR) as MAR, sum(JAN+FEB+MAR) as '1st_kvartal', sum(APR) as APR, sum(MAY) as MAY, sum(JUN) as JUN, sum(APR+MAY+JUN) as '2st_kvartal', sum(JUL) as JUL, sum(AUG) as AUG, sum(SEP) as SEP, sum(JUL+AUG+SEP) as '3st_kvartal', sum(OKT) as OKT, sum(NOV) as NOV, sum(DEC) as DEC, sum(OKT+NOV+DEC) as '4st_kvartal' from NEED_BODY where GUID=@gd group by GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, PRICE_EDINIC_PLAN, SIZE_CLOTHES, SIZE_SHOES, HEIGHT order by NAIMENOVAN_COVERALLS", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            command.Parameters.AddWithValue("gd", need.gd);
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "NEED_BODY");
            dataGridView1.DataSource = ds.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

            fill_gridview_gr();            
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //Группировка по наименованию и Испол аппарату
            if ((checkBox1.Checked == true) && (radioButton1.Checked==true))
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox9.Checked = false;

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, ROUND(SUM(PRICE_KOMPLEKT_PLAN), 2) as SUM_PR from NEED_BODY where GUID=@gd and BRANCH=@br group by GUID, NAIMENOVAN_COVERALLS order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                command.Parameters.AddWithValue("br", "Исполнительный аппарат");                              
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_naimenovan();

            }

            //Группировка по наименованию и Абаканскому филиалу
            if ((checkBox1.Checked == true) && (radioButton2.Checked == true))
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox9.Checked = false;

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, ROUND(SUM(PRICE_KOMPLEKT_PLAN), 2) as SUM_PR from NEED_BODY where GUID=@gd and BRANCH=@br group by GUID, NAIMENOVAN_COVERALLS order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                command.Parameters.AddWithValue("br", "Абаканский филиал ОАО \"СибЭР\"");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_naimenovan();

            }

            //Группировка по наименованию и Барнаульскому филиалу
            if ((checkBox1.Checked == true) && (radioButton3.Checked == true))
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox9.Checked = false;

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, ROUND(SUM(PRICE_KOMPLEKT_PLAN), 2) as SUM_PR from NEED_BODY where GUID=@gd and BRANCH=@br group by GUID, NAIMENOVAN_COVERALLS order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                command.Parameters.AddWithValue("br", "Барнаульский филиал ОАО \"СибЭР\"");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_naimenovan();

            }

            //Группировка по наименованию и Кемеровскому филиалу
            if ((checkBox1.Checked == true) && (radioButton4.Checked == true))
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox9.Checked = false;

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, ROUND(SUM(PRICE_KOMPLEKT_PLAN), 2) as SUM_PR from NEED_BODY where GUID=@gd and BRANCH=@br group by GUID, NAIMENOVAN_COVERALLS order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                command.Parameters.AddWithValue("br", "Кемеровский филиал ОАО \"СибЭР\"");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_naimenovan();

            }

            //Группировка по наименованию и Красноярскому филиалу
            if ((checkBox1.Checked == true) && (radioButton5.Checked == true))
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox9.Checked = false;

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, ROUND(SUM(PRICE_KOMPLEKT_PLAN), 2) as SUM_PR from NEED_BODY where GUID=@gd and BRANCH=@br group by GUID, NAIMENOVAN_COVERALLS order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                command.Parameters.AddWithValue("br", "Красноярский филиал ОАО \"СибЭР\"");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_naimenovan();

            }

            //Группировка по наименованию и отключенной фильтрацией
            if ((checkBox1.Checked == true) && (radioButton6.Checked == true))
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox9.Checked = false;

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, ROUND(SUM(PRICE_KOMPLEKT_PLAN), 2) as SUM_PR from NEED_BODY where GUID=@gd group by GUID, NAIMENOVAN_COVERALLS order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_naimenovan();

            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            //Группировка по подразделению и исполнительному аппарату
            if ((checkBox2.Checked == true) && (radioButton1.Checked==true))
            {
                checkBox1.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox9.Checked = false;

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, DEPARTMENT, ROUND(SUM(PRICE_KOMPLEKT_PLAN), 2) as SUM_PR from NEED_BODY where GUID=@gd and BRANCH=@br group by GUID, DEPARTMENT order by DEPARTMENT", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                command.Parameters.AddWithValue("br", "Исполнительный аппарат");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_department();

            }

            //Группировка по подразделению и Абаканскому филиалу
            if ((checkBox2.Checked == true) && (radioButton2.Checked == true))
            {
                checkBox1.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox9.Checked = false;

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, DEPARTMENT, ROUND(SUM(PRICE_KOMPLEKT_PLAN), 2) as SUM_PR from NEED_BODY where GUID=@gd and BRANCH=@br group by GUID, DEPARTMENT order by DEPARTMENT", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                command.Parameters.AddWithValue("br", "Абаканский филиал ОАО \"СибЭР\"");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_department();

            }

            //Группировка по подразделению и Барнаульскому филиалу
            if ((checkBox2.Checked == true) && (radioButton3.Checked == true))
            {
                checkBox1.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox9.Checked = false;

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, DEPARTMENT, ROUND(SUM(PRICE_KOMPLEKT_PLAN), 2) as SUM_PR from NEED_BODY where GUID=@gd and BRANCH=@br group by GUID, DEPARTMENT order by DEPARTMENT", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                command.Parameters.AddWithValue("br", "Барнаульский филиал ОАО \"СибЭР\"");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_department();

            }

            //Группировка по подразделению и Кемеровскому филиалу
            if ((checkBox2.Checked == true) && (radioButton4.Checked == true))
            {
                checkBox1.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox9.Checked = false;

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, DEPARTMENT, ROUND(SUM(PRICE_KOMPLEKT_PLAN), 2) as SUM_PR from NEED_BODY where GUID=@gd and BRANCH=@br group by GUID, DEPARTMENT order by DEPARTMENT", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                command.Parameters.AddWithValue("br", "Кемеровский филиал ОАО \"СибЭР\"");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_department();

            }

            //Группировка по подразделению и Красноярскому филиалу
            if ((checkBox2.Checked == true) && (radioButton5.Checked == true))
            {
                checkBox1.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox9.Checked = false;

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, DEPARTMENT, ROUND(SUM(PRICE_KOMPLEKT_PLAN), 2) as SUM_PR from NEED_BODY where GUID=@gd and BRANCH=@br group by GUID, DEPARTMENT order by DEPARTMENT", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                command.Parameters.AddWithValue("br", "Красноярский филиал ОАО \"СибЭР\"");
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_department();

            }

            //Группировка по подразделению и без фильтрации
            if ((checkBox2.Checked == true) && (radioButton6.Checked == true))
            {
                checkBox1.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox9.Checked = false;

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, DEPARTMENT, ROUND(SUM(PRICE_KOMPLEKT_PLAN), 2) as SUM_PR from NEED_BODY where GUID=@gd group by GUID, DEPARTMENT order by DEPARTMENT", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
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
                checkBox9.Checked = false;

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, BRANCH, ROUND(SUM(PRICE_KOMPLEKT_PLAN), 2) as SUM_PR from NEED_BODY where GUID=@gd group by GUID, BRANCH order by BRANCH", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
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
                checkBox9.Checked = false;

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, FIO, ROUND(SUM(PRICE_KOMPLEKT_PLAN), 2) as SUM_PR from NEED_BODY where GUID=@gd group by GUID, FIO order by FIO", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
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
                checkBox9.Checked = false;

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, KONTRAGENT, ROUND(SUM(PRICE_KOMPLEKT_PLAN), 2) as SUM_PR from NEED_BODY where GUID=@gd group by GUID, KONTRAGENT order by KONTRAGENT", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_ka();

            }
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

                filter_state = false;
                label3.Text = "Режим фильтра отключен";

                //Запрос на вывод содержания заявки
                SqlCommand command = new SqlCommand("select GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, SUM(KOL_VO) as KOLVO, PRICE_EDINIC_PLAN, sum(PRICE_KOMPLEKT_PLAN) as PRICE_KOMPLEKT_PLAN, sum(JAN) as JAN, sum(FEB) as FEB, sum(MAR) as MAR, sum(JAN+FEB+MAR) as '1st_kvartal', sum(APR) as APR, sum(MAY) as MAY, sum(JUN) as JUN, sum(APR+MAY+JUN) as '2st_kvartal', sum(JUL) as JUL, sum(AUG) as AUG, sum(SEP) as SEP, sum(JUL+AUG+SEP) as '3st_kvartal', sum(OKT) as OKT, sum(NOV) as NOV, sum(DEC) as DEC, sum(OKT+NOV+DEC) as '4st_kvartal' from NEED_BODY where GUID=@gd group by GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, PRICE_EDINIC_PLAN order by NAIMENOVAN_COVERALLS", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", need.gd);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_BODY");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview_group_torgi();       

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            //Книга.
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица.
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            for (int i = 0; i < dataGridView1.Rows.Count-1; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount-1; j++)
                {
                    ExcelApp.Cells[i + 1, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            //Вызываем нашу созданную эксельку.
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;  
        }

      

        
    }
}
