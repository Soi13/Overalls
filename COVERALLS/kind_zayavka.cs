using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Media;
using System.Data.SqlClient;

namespace COVERALLS
{
    public partial class kind_zayavka : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
   
        string beg_m, end_m, month, year;
        public string NORMS_LIST_ID;
        public string DEPARTMENT;
        public string BRANCH;

        public kind_zayavka()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if ((checkBox3.Checked == false) && (comboBox7.Text.Length == 0) || (comboBox1.Text.Length == 0))
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не выбран месяц или подразделение!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("Сформировать заявку?", "Вопрос", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                //Присвоение значений переменным взависимости от выбранного квартала
                if (comboBox1.Text == "Январь")
                {
                    month = "1";
                }
                if (comboBox1.Text == "Февраль")
                {
                    month = "2";
                }
                if (comboBox1.Text == "Март")
                {
                    month = "3";
                }
                if (comboBox1.Text == "Апрель")
                {
                    month = "4";
                }
                if (comboBox1.Text == "Май")
                {
                    month = "5";
                }
                if (comboBox1.Text == "Июнь")
                {
                    month = "6";
                }
                if (comboBox1.Text == "Июль")
                {
                    month = "7";
                }
                if (comboBox1.Text == "Август")
                {
                    month = "8";
                }
                if (comboBox1.Text == "Сентябрь")
                {
                    month = "9";
                }
                if (comboBox1.Text == "Октябрь")
                {
                    month = "10";
                }
                if (comboBox1.Text == "Ноябрь")
                {
                    month = "11";
                }
                if (comboBox1.Text == "Декабрь")
                {
                    month = "12";
                }

                year = Convert.ToString(DateTime.Today.Year);

                progressBar1.Value = 0;


                if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2) || (Form2.val == 1)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ
                {
                    //Если выбрано по всем подразделениям
                    if (checkBox3.Checked == true)
                    {
                        conn.Open();
                        //Получаем список всех сотрудников. 
                        SqlCommand command = new SqlCommand("select * from sotrudniki", conn);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        SqlCommandBuilder cb = new SqlCommandBuilder(da);
                        DataSet ds = new DataSet();
                        //conn.Close();
                        da.Fill(ds, "SOTRUDNIKI");
                        ///////////////////////////////////////////////////////////////

                        //Генерируем ключа GUID для связи заголовка заявки с его телом
                        SqlCommand command6 = new SqlCommand("select NEWID()", conn);
                        SqlDataAdapter da6 = new SqlDataAdapter(command6);//Переменная объявлена как глобальная
                        SqlCommandBuilder cb6 = new SqlCommandBuilder(da6);
                        DataSet ds6 = new DataSet();
                        //Заполнение DataGridView наименованиями полей 
                        da6.Fill(ds6, "NEWID");
                        /////////////////////////////

                        //Вставка данных в таблицу Заголовок Заявки
                        SqlCommand scmd4 = conn.CreateCommand();
                        scmd4.CommandText = "declare @ID uniqueidentifier " +
                                            "set @ID=NEWID() " +
                                            "INSERT into ZAYAVKA_MAIN (GUID, USER_ID, KIND_ZAYAV, PODR, PERIOD, DATETIME_CREATE) VALUES (@guid, @us_id, @kd_zayav, @podr, @per, GETDATE())";

                        scmd4.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                        scmd4.Parameters.AddWithValue("us_id", Form2.val);
                        scmd4.Parameters.AddWithValue("kd_zayav", "Месячная");
                        scmd4.Parameters.AddWithValue("podr", comboBox7.Text);
                        scmd4.Parameters.AddWithValue("per", comboBox1.Text);
                        SqlDataReader reader4;
                        reader4 = scmd4.ExecuteReader();
                        /////////////////////////////////

                        progressBar1.Maximum = ds.Tables[0].Rows.Count;
                        progressBar1.Value = 0;

                        //Идем циклом по сформированному списку сотрудников
                        for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                        {
                            //Получаем то, что выдано сотруднику
                            SqlCommand command1 = new SqlCommand("select NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and DATEPART(year,VIDANO.DATE_SLED_VIDACHI)=@year and DATEPART(month,VIDANO.DATE_SLED_VIDACHI)=@month and SOTRUDNIKI.ID=@id_sotr group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            command1.Parameters.AddWithValue("id_sotr", ds.Tables[0].Rows[i][0].ToString());
                            command1.Parameters.AddWithValue("year", year);
                            command1.Parameters.AddWithValue("month", month);
                            SqlDataAdapter da1 = new SqlDataAdapter(command1);
                            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                            DataSet ds1 = new DataSet();
                            conn.Close();
                            da1.Fill(ds1, "VIDANO");
                            ////////////////////////////////////////////////////////////////

                            //Проверяем выдачу, если сотруднику ничего не выдано, т.е. кол-во записей 0, то этот блок не выполняется, т.к. разбивать по месяцам нечего. Сразу переходим к следующему блоку, в котором проверяется, что не было выдано сотруднику.
                            if (ds1.Tables[0].Rows.Count != 0)
                            {
                                //Разбиваем по месяцам кол-во , которое попало в заявку
                                for (int g = 0; g <= ds1.Tables[0].Rows.Count - 1; g++)
                                {
                                    conn.Open();

                                    //Разбивка по месяцам потребности СИЗ по каждой позиции
                                    SqlCommand command5 = new SqlCommand("select VIDANO.NAIMEN_OVERALLS, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 1 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as JAN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 2 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as FEB, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 3 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as MAR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 4 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as APR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 5 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as MAY, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 6 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as JUN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 7 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as JUL, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 8 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as AUG, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 9 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as SEP, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 10 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as OKT, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 11 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as NOV, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 12 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as DEC " +
                                                                         "from VIDANO " +
                                                                         "where VIDANO.SOTRUDNIK_ID=@id_sot and VIDANO.NAIMEN_OVERALLS=@naim_cover " +
                                                                         "group by VIDANO.NAIMEN_OVERALLS", conn);

                                    command5.Parameters.AddWithValue("id_sot", ds.Tables[0].Rows[i][0].ToString());
                                    command5.Parameters.AddWithValue("naim_cover", ds1.Tables[0].Rows[g][0].ToString());
                                    SqlDataAdapter da5 = new SqlDataAdapter(command5);//Переменная объявлена как глобальная
                                    SqlCommandBuilder cb5 = new SqlCommandBuilder(da5);
                                    DataSet ds5 = new DataSet();
                                    da5.Fill(ds5, "VIDANO");

                                    //Если записи при выборе нет, то пропускаем этот шаг цикла
                                    if (ds5.Tables[0].Rows.Count > 0)
                                    {
                                        //Вставка данных в таблицу Заявки
                                        SqlCommand scmd5 = conn.CreateCommand();
                                        scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover1, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                        scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                        scmd5.Parameters.AddWithValue("naim_cover1", ds1.Tables[0].Rows[g][0]);
                                        scmd5.Parameters.AddWithValue("characterist", ds1.Tables[0].Rows[g][1]);
                                        scmd5.Parameters.AddWithValue("dep", ds1.Tables[0].Rows[g][4]);
                                        scmd5.Parameters.AddWithValue("br", ds1.Tables[0].Rows[g][6]);
                                        scmd5.Parameters.AddWithValue("ed_iz", ds1.Tables[0].Rows[g][5]);
                                        scmd5.Parameters.AddWithValue("kolvo", ds1.Tables[0].Rows[g][2]);
                                        scmd5.Parameters.AddWithValue("pr_ed_pl", ds1.Tables[0].Rows[g][3]);
                                        scmd5.Parameters.AddWithValue("jan", ds5.Tables[0].Rows[0][1]);
                                        scmd5.Parameters.AddWithValue("feb", ds5.Tables[0].Rows[0][2]);
                                        scmd5.Parameters.AddWithValue("mar", ds5.Tables[0].Rows[0][3]);
                                        scmd5.Parameters.AddWithValue("apr", ds5.Tables[0].Rows[0][4]);
                                        scmd5.Parameters.AddWithValue("may", ds5.Tables[0].Rows[0][5]);
                                        scmd5.Parameters.AddWithValue("jun", ds5.Tables[0].Rows[0][6]);
                                        scmd5.Parameters.AddWithValue("jul", ds5.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("aug", ds5.Tables[0].Rows[0][8]);
                                        scmd5.Parameters.AddWithValue("sep", ds5.Tables[0].Rows[0][9]);
                                        scmd5.Parameters.AddWithValue("okt", ds5.Tables[0].Rows[0][10]);
                                        scmd5.Parameters.AddWithValue("nov", ds5.Tables[0].Rows[0][11]);
                                        scmd5.Parameters.AddWithValue("dec", ds5.Tables[0].Rows[0][12]);
                                        scmd5.Parameters.AddWithValue("kontr", ds1.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                        scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                        scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                        scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                        SqlDataReader reader5;
                                        reader5 = scmd5.ExecuteReader();
                                        conn.Close();
                                    }
                                    else
                                    {
                                        conn.Close();
                                        continue;
                                    }

                                }
                            }


                            //Получаем то, что осталось к выдаче, но пока не выдано
                            //Определение нормы сотрудника
                            SqlCommand cmd = new SqlCommand("select norms_list.* from norms_personal,norms_list where norms_personal.norms_id=norms_list.norms_id and norms_personal.sotrudnik_id=@id_sotr1", conn);
                            cmd.Parameters.AddWithValue("id_sotr1", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da2 = new SqlDataAdapter(cmd);//Переменная объявлена как глобальная
                            SqlCommandBuilder cb2 = new SqlCommandBuilder(da2);
                            DataSet ds2 = new DataSet();
                            da2.Fill(ds2, "NORMS_LIST");
                            ///////////////////////////////

                            //Определение что уже выдано сотруднику
                            SqlCommand cmd1 = new SqlCommand("select VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID, NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and SOTRUDNIKI.ID=@id_sotr2 group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN, VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            cmd1.Parameters.AddWithValue("id_sotr2", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da3 = new SqlDataAdapter(cmd1);
                            SqlCommandBuilder cb3 = new SqlCommandBuilder(da3);
                            DataSet ds3 = new DataSet();
                            da3.Fill(ds3, "VIDANO");
                            ///////////////////////////////

                            ///////////////////////////////////////////////////
                            for (int nr = 0; nr <= ds2.Tables[0].Rows.Count - 1; nr++)
                            {
                                int aa = 0;
                                string ID = Convert.ToString(ds2.Tables[0].Rows[nr][0]);
                                string NORMS_ID = Convert.ToString(ds2.Tables[0].Rows[nr][1]);
                                string NAIMEN_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string CHARACTERISTIC_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string ED_IZM = Convert.ToString(ds2.Tables[0].Rows[nr][4]);
                                string KOLVO = Convert.ToString(ds2.Tables[0].Rows[nr][5]);

                                for (int vd = 0; vd <= ds3.Tables[0].Rows.Count - 1; vd++)
                                {
                                    string SOTRUDNIK_ID = Convert.ToString(ds3.Tables[0].Rows[vd][0]);
                                    NORMS_LIST_ID = Convert.ToString(ds3.Tables[0].Rows[vd][1]);
                                    string NAIMEN_OVERALLS1 = Convert.ToString(ds3.Tables[0].Rows[vd][2]);
                                    string DEPARTMENT = Convert.ToString(ds3.Tables[0].Rows[vd][6]);
                                    string BRANCH = Convert.ToString(ds3.Tables[0].Rows[vd][8]);
                                    string ED_IZM1 = Convert.ToString(ds3.Tables[0].Rows[vd][3]);
                                    string KOLVO1 = Convert.ToString(ds3.Tables[0].Rows[vd][4]);
                                    string PERIOD_ISPOLZOVAN1 = Convert.ToString(ds3.Tables[0].Rows[vd][5]);

                                    //Поиск в нормах того, что уже отпущено для вычисления кол-ва остатка по норме (то, что еще можно выдать)
                                    if (ID == NORMS_LIST_ID)
                                    {
                                        int ostatok = Convert.ToInt16(KOLVO) - Convert.ToInt16(KOLVO1); //Вычисление остатка от нормы, возможного к отпуску
                                        if (ostatok != 0)//Если остаток по позиции не 0, то выводим его в datagridview
                                        {
                                            //Вставка данных в таблицу Заявки
                                            conn.Open();
                                            SqlCommand scmd5 = conn.CreateCommand();
                                            scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                            scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                            scmd5.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                            scmd5.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                            scmd5.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                            scmd5.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                            scmd5.Parameters.AddWithValue("ed_iz", ED_IZM);
                                            scmd5.Parameters.AddWithValue("kolvo", ostatok);
                                            scmd5.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                            scmd5.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                            scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                            scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                            scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                            scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                            //Проверяем какой месяц выбран и исходя из этого кол-во невыданных СИЗов помещаем на выбранный месяц 
                                            if (comboBox1.Text == "Январь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", ostatok); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Февраль")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", ostatok);
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Март")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", ostatok);
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Апрель")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", ostatok);
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Май")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", ostatok);
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Июнь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", ostatok);
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Июль")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", ostatok);
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Август")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", ostatok);
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Сентябрь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", ostatok);
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Октябрь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", ostatok);
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Ноябрь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", ostatok);
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Декабрь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", ostatok);
                                            }
                                            SqlDataReader reader5;
                                            reader5 = scmd5.ExecuteReader();
                                            conn.Close();
                                        }
                                        aa++;
                                    }
                                }

                                if (aa == 0)
                                {
                                    //Вставка данных в таблицу Заявки
                                    conn.Open();
                                    SqlCommand scmd6 = conn.CreateCommand();
                                    scmd6.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                    scmd6.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                    scmd6.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                    scmd6.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                    scmd6.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                    scmd6.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                    scmd6.Parameters.AddWithValue("ed_iz", ED_IZM);
                                    scmd6.Parameters.AddWithValue("kolvo", KOLVO);
                                    scmd6.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                    scmd6.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                    scmd6.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                    scmd6.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                    scmd6.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                    scmd6.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                    //Проверяем какой месяц выбран и исходя из этого кол-во невыданных СИЗов помещаем на выбранный месяц 
                                    if (comboBox1.Text == "Январь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", KOLVO); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Февраль")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", KOLVO);
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Март")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", KOLVO);
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Апрель")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", KOLVO);
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Май")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", KOLVO);
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Июнь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", KOLVO);
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Июль")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", KOLVO);
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Август")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", KOLVO);
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Сентябрь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", KOLVO);
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Октябрь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", KOLVO);
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Ноябрь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", KOLVO);
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Декабрь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", KOLVO);
                                    }
                                    SqlDataReader reader6;
                                    reader6 = scmd6.ExecuteReader();
                                    conn.Close();
                                }

                            }
                            ////////////////////////////////////////////////////////////////

                            progressBar1.Value++;
                        }

                        SystemSounds.Beep.Play();
                        MessageBox.Show("Заявка сформирована. Просмотреть ее возможно в списке заявок.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                    }
                    else
                    {
                        conn.Open();
                        //Получаем список всех сотрудников. 
                        SqlCommand command = new SqlCommand("select * from sotrudniki where department=@dp", conn);
                        command.Parameters.AddWithValue("dp", comboBox7.Text);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        SqlCommandBuilder cb = new SqlCommandBuilder(da);
                        DataSet ds = new DataSet();
                        //conn.Close();
                        da.Fill(ds, "SOTRUDNIKI");
                        ///////////////////////////////////////////////////////////////

                        //Генерируем ключа GUID для связи заголовка заявки с его телом
                        SqlCommand command6 = new SqlCommand("select NEWID()", conn);
                        SqlDataAdapter da6 = new SqlDataAdapter(command6);//Переменная объявлена как глобальная
                        SqlCommandBuilder cb6 = new SqlCommandBuilder(da6);
                        DataSet ds6 = new DataSet();
                        //Заполнение DataGridView наименованиями полей 
                        da6.Fill(ds6, "NEWID");
                        /////////////////////////////

                        //Вставка данных в таблицу Заголовок Заявки
                        SqlCommand scmd4 = conn.CreateCommand();
                        scmd4.CommandText = "declare @ID uniqueidentifier " +
                                            "set @ID=NEWID() " +
                                            "INSERT into ZAYAVKA_MAIN (GUID, USER_ID, KIND_ZAYAV, PODR, PERIOD, DATETIME_CREATE) VALUES (@guid, @us_id, @kd_zayav, @podr, @per, GETDATE())";

                        scmd4.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                        scmd4.Parameters.AddWithValue("us_id", Form2.val);
                        scmd4.Parameters.AddWithValue("kd_zayav", "Месячная");
                        scmd4.Parameters.AddWithValue("podr", comboBox7.Text);
                        scmd4.Parameters.AddWithValue("per", comboBox1.Text);
                        SqlDataReader reader4;
                        reader4 = scmd4.ExecuteReader();
                        /////////////////////////////////

                        progressBar1.Maximum = ds.Tables[0].Rows.Count;
                        progressBar1.Value = 0;

                        //Идем циклом по сформированному списку сотрудников
                        for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                        {
                            //Получаем то, что выдано сотруднику
                            SqlCommand command1 = new SqlCommand("select NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and DATEPART(year,VIDANO.DATE_SLED_VIDACHI)=@year and DATEPART(month,VIDANO.DATE_SLED_VIDACHI)=@month and SOTRUDNIKI.ID=@id_sotr group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            command1.Parameters.AddWithValue("id_sotr", ds.Tables[0].Rows[i][0].ToString());
                            command1.Parameters.AddWithValue("year", year);
                            command1.Parameters.AddWithValue("month", month);
                            SqlDataAdapter da1 = new SqlDataAdapter(command1);
                            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                            DataSet ds1 = new DataSet();
                            conn.Close();
                            da1.Fill(ds1, "VIDANO");
                            ////////////////////////////////////////////////////////////////

                            //Проверяем выдачу, если сотруднику ничего не выдано, т.е. кол-во записей 0, то этот блок не выполняется, т.к. разбивать по месяцам нечего. Сразу переходим к следующему блоку, в котором проверяется, что не было выдано сотруднику.
                            if (ds1.Tables[0].Rows.Count != 0)
                            {
                                //Разбиваем по месяцам кол-во , которое попало в заявку
                                for (int g = 0; g <= ds1.Tables[0].Rows.Count - 1; g++)
                                {
                                    conn.Open();

                                    //Разбивка по месяцам потребности СИЗ по каждой позиции
                                    SqlCommand command5 = new SqlCommand("select VIDANO.NAIMEN_OVERALLS, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 1 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as JAN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 2 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as FEB, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 3 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as MAR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 4 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as APR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 5 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as MAY, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 6 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as JUN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 7 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as JUL, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 8 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as AUG, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 9 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as SEP, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 10 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as OKT, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 11 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as NOV, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 12 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as DEC " +
                                                                         "from VIDANO " +
                                                                         "where VIDANO.SOTRUDNIK_ID=@id_sot and VIDANO.NAIMEN_OVERALLS=@naim_cover " +
                                                                         "group by VIDANO.NAIMEN_OVERALLS", conn);

                                    command5.Parameters.AddWithValue("id_sot", ds.Tables[0].Rows[i][0].ToString());
                                    command5.Parameters.AddWithValue("naim_cover", ds1.Tables[0].Rows[g][0].ToString());
                                    SqlDataAdapter da5 = new SqlDataAdapter(command5);//Переменная объявлена как глобальная
                                    SqlCommandBuilder cb5 = new SqlCommandBuilder(da5);
                                    DataSet ds5 = new DataSet();
                                    da5.Fill(ds5, "VIDANO");

                                    //Если записи при выборе нет, то пропускаем этот шаг цикла
                                    if (ds5.Tables[0].Rows.Count > 0)
                                    {
                                        //Вставка данных в таблицу Заявки
                                        SqlCommand scmd5 = conn.CreateCommand();
                                        scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover1, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                        scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                        scmd5.Parameters.AddWithValue("naim_cover1", ds1.Tables[0].Rows[g][0]);
                                        scmd5.Parameters.AddWithValue("characterist", ds1.Tables[0].Rows[g][1]);
                                        scmd5.Parameters.AddWithValue("dep", ds1.Tables[0].Rows[g][4]);
                                        scmd5.Parameters.AddWithValue("br", ds1.Tables[0].Rows[g][6]);
                                        scmd5.Parameters.AddWithValue("ed_iz", ds1.Tables[0].Rows[g][5]);
                                        scmd5.Parameters.AddWithValue("kolvo", ds1.Tables[0].Rows[g][2]);
                                        scmd5.Parameters.AddWithValue("pr_ed_pl", ds1.Tables[0].Rows[g][3]);
                                        scmd5.Parameters.AddWithValue("jan", ds5.Tables[0].Rows[0][1]);
                                        scmd5.Parameters.AddWithValue("feb", ds5.Tables[0].Rows[0][2]);
                                        scmd5.Parameters.AddWithValue("mar", ds5.Tables[0].Rows[0][3]);
                                        scmd5.Parameters.AddWithValue("apr", ds5.Tables[0].Rows[0][4]);
                                        scmd5.Parameters.AddWithValue("may", ds5.Tables[0].Rows[0][5]);
                                        scmd5.Parameters.AddWithValue("jun", ds5.Tables[0].Rows[0][6]);
                                        scmd5.Parameters.AddWithValue("jul", ds5.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("aug", ds5.Tables[0].Rows[0][8]);
                                        scmd5.Parameters.AddWithValue("sep", ds5.Tables[0].Rows[0][9]);
                                        scmd5.Parameters.AddWithValue("okt", ds5.Tables[0].Rows[0][10]);
                                        scmd5.Parameters.AddWithValue("nov", ds5.Tables[0].Rows[0][11]);
                                        scmd5.Parameters.AddWithValue("dec", ds5.Tables[0].Rows[0][12]);
                                        scmd5.Parameters.AddWithValue("kontr", ds1.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                        scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                        scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                        scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                        SqlDataReader reader5;
                                        reader5 = scmd5.ExecuteReader();
                                        conn.Close();
                                    }
                                    else
                                    {
                                        conn.Close();
                                        continue;
                                    }

                                }
                            }


                            //Получаем то, что осталось к выдаче, но пока не выдано
                            //Определение нормы сотрудника
                            SqlCommand cmd = new SqlCommand("select norms_list.* from norms_personal,norms_list where norms_personal.norms_id=norms_list.norms_id and norms_personal.sotrudnik_id=@id_sotr1", conn);
                            cmd.Parameters.AddWithValue("id_sotr1", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da2 = new SqlDataAdapter(cmd);//Переменная объявлена как глобальная
                            SqlCommandBuilder cb2 = new SqlCommandBuilder(da2);
                            DataSet ds2 = new DataSet();
                            da2.Fill(ds2, "NORMS_LIST");
                            ///////////////////////////////

                            //Определение что уже выдано сотруднику
                            SqlCommand cmd1 = new SqlCommand("select VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID, NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and SOTRUDNIKI.ID=@id_sotr2 group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN, VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            cmd1.Parameters.AddWithValue("id_sotr2", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da3 = new SqlDataAdapter(cmd1);
                            SqlCommandBuilder cb3 = new SqlCommandBuilder(da3);
                            DataSet ds3 = new DataSet();
                            da3.Fill(ds3, "VIDANO");
                            ///////////////////////////////

                            ///////////////////////////////////////////////////
                            for (int nr = 0; nr <= ds2.Tables[0].Rows.Count - 1; nr++)
                            {
                                int aa = 0;
                                string ID = Convert.ToString(ds2.Tables[0].Rows[nr][0]);
                                string NORMS_ID = Convert.ToString(ds2.Tables[0].Rows[nr][1]);
                                string NAIMEN_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string CHARACTERISTIC_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string ED_IZM = Convert.ToString(ds2.Tables[0].Rows[nr][4]);
                                string KOLVO = Convert.ToString(ds2.Tables[0].Rows[nr][5]);

                                for (int vd = 0; vd <= ds3.Tables[0].Rows.Count - 1; vd++)
                                {
                                    string SOTRUDNIK_ID = Convert.ToString(ds3.Tables[0].Rows[vd][0]);
                                    NORMS_LIST_ID = Convert.ToString(ds3.Tables[0].Rows[vd][1]);
                                    string NAIMEN_OVERALLS1 = Convert.ToString(ds3.Tables[0].Rows[vd][2]);
                                    string DEPARTMENT = Convert.ToString(ds3.Tables[0].Rows[vd][6]);
                                    string BRANCH = Convert.ToString(ds3.Tables[0].Rows[vd][8]);
                                    string ED_IZM1 = Convert.ToString(ds3.Tables[0].Rows[vd][3]);
                                    string KOLVO1 = Convert.ToString(ds3.Tables[0].Rows[vd][4]);
                                    string PERIOD_ISPOLZOVAN1 = Convert.ToString(ds3.Tables[0].Rows[vd][5]);

                                    //Поиск в нормах того, что уже отпущено для вычисления кол-ва остатка по норме (то, что еще можно выдать)
                                    if (ID == NORMS_LIST_ID)
                                    {
                                        int ostatok = Convert.ToInt16(KOLVO) - Convert.ToInt16(KOLVO1); //Вычисление остатка от нормы, возможного к отпуску
                                        if (ostatok != 0)//Если остаток по позиции не 0, то выводим его в datagridview
                                        {
                                            //Вставка данных в таблицу Заявки
                                            conn.Open();
                                            SqlCommand scmd5 = conn.CreateCommand();
                                            scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                            scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                            scmd5.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                            scmd5.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                            scmd5.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                            scmd5.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                            scmd5.Parameters.AddWithValue("ed_iz", ED_IZM);
                                            scmd5.Parameters.AddWithValue("kolvo", ostatok);
                                            scmd5.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                            scmd5.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                            scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                            scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                            scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                            scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                            //Проверяем какой месяц выбран и исходя из этого кол-во невыданных СИЗов помещаем на выбранный месяц 
                                            if (comboBox1.Text == "Январь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", ostatok); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Февраль")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", ostatok);
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Март")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", ostatok);
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Апрель")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", ostatok);
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Май")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", ostatok);
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Июнь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", ostatok);
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Июль")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", ostatok);
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Август")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", ostatok);
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Сентябрь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", ostatok);
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Октябрь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", ostatok);
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Ноябрь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", ostatok);
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Декабрь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", ostatok);
                                            }
                                            SqlDataReader reader5;
                                            reader5 = scmd5.ExecuteReader();
                                            conn.Close();
                                        }
                                        aa++;
                                    }
                                }

                                if (aa == 0)
                                {
                                    //Вставка данных в таблицу Заявки
                                    conn.Open();
                                    SqlCommand scmd6 = conn.CreateCommand();
                                    scmd6.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                    scmd6.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                    scmd6.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                    scmd6.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                    scmd6.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                    scmd6.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                    scmd6.Parameters.AddWithValue("ed_iz", ED_IZM);
                                    scmd6.Parameters.AddWithValue("kolvo", KOLVO);
                                    scmd6.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                    scmd6.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                    scmd6.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                    scmd6.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                    scmd6.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                    scmd6.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                    //Проверяем какой месяц выбран и исходя из этого кол-во невыданных СИЗов помещаем на выбранный месяц 
                                    if (comboBox1.Text == "Январь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", KOLVO); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Февраль")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", KOLVO);
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Март")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", KOLVO);
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Апрель")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", KOLVO);
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Май")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", KOLVO);
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Июнь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", KOLVO);
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Июль")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", KOLVO);
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Август")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", KOLVO);
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Сентябрь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", KOLVO);
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Октябрь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", KOLVO);
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Ноябрь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", KOLVO);
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Декабрь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", KOLVO);
                                    }
                                    SqlDataReader reader6;
                                    reader6 = scmd6.ExecuteReader();
                                    conn.Close();
                                }

                            }
                            ////////////////////////////////////////////////////////////////

                            progressBar1.Value++;
                        }

                        SystemSounds.Beep.Play();
                        MessageBox.Show("Заявка сформирована. Просмотреть ее возможно в списке заявок.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information); 

                    }
                }
                else //Если заходят остальные. Все кроме Петушкова и Скворцова
                {
                    //Если выбрано по всем подразделениям
                    if (checkBox3.Checked == true)
                    {
                        conn.Open();
                        //Получаем список всех сотрудников. 
                        SqlCommand command = new SqlCommand("select * from sotrudniki where users_branch_id=@br", conn);
                        command.Parameters.AddWithValue("br",Form2.branch);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        SqlCommandBuilder cb = new SqlCommandBuilder(da);
                        DataSet ds = new DataSet();
                        //conn.Close();
                        da.Fill(ds, "SOTRUDNIKI");
                        ///////////////////////////////////////////////////////////////

                        //Генерируем ключа GUID для связи заголовка заявки с его телом
                        SqlCommand command6 = new SqlCommand("select NEWID()", conn);
                        SqlDataAdapter da6 = new SqlDataAdapter(command6);//Переменная объявлена как глобальная
                        SqlCommandBuilder cb6 = new SqlCommandBuilder(da6);
                        DataSet ds6 = new DataSet();
                        //Заполнение DataGridView наименованиями полей 
                        da6.Fill(ds6, "NEWID");
                        /////////////////////////////

                        //Вставка данных в таблицу Заголовок Заявки
                        SqlCommand scmd4 = conn.CreateCommand();
                        scmd4.CommandText = "declare @ID uniqueidentifier " +
                                            "set @ID=NEWID() " +
                                            "INSERT into ZAYAVKA_MAIN (GUID, USER_ID, KIND_ZAYAV, PODR, PERIOD, DATETIME_CREATE) VALUES (@guid, @us_id, @kd_zayav, @podr, @per, GETDATE())";

                        scmd4.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                        scmd4.Parameters.AddWithValue("us_id", Form2.val);
                        scmd4.Parameters.AddWithValue("kd_zayav", "Месячная");
                        scmd4.Parameters.AddWithValue("podr", comboBox7.Text);
                        scmd4.Parameters.AddWithValue("per", comboBox1.Text);
                        SqlDataReader reader4;
                        reader4 = scmd4.ExecuteReader();
                        /////////////////////////////////

                        progressBar1.Maximum = ds.Tables[0].Rows.Count;
                        progressBar1.Value = 0;

                        //Идем циклом по сформированному списку сотрудников
                        for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                        {
                            //Получаем то, что выдано сотруднику
                            SqlCommand command1 = new SqlCommand("select NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and DATEPART(year,VIDANO.DATE_SLED_VIDACHI)=@year and DATEPART(month,VIDANO.DATE_SLED_VIDACHI)=@month and SOTRUDNIKI.ID=@id_sotr group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            command1.Parameters.AddWithValue("id_sotr", ds.Tables[0].Rows[i][0].ToString());
                            command1.Parameters.AddWithValue("year", year);
                            command1.Parameters.AddWithValue("month", month);
                            SqlDataAdapter da1 = new SqlDataAdapter(command1);
                            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                            DataSet ds1 = new DataSet();
                            conn.Close();
                            da1.Fill(ds1, "VIDANO");
                            ////////////////////////////////////////////////////////////////

                            //Проверяем выдачу, если сотруднику ничего не выдано, т.е. кол-во записей 0, то этот блок не выполняется, т.к. разбивать по месяцам нечего. Сразу переходим к следующему блоку, в котором проверяется, что не было выдано сотруднику.
                            if (ds1.Tables[0].Rows.Count != 0)
                            {
                                //Разбиваем по месяцам кол-во , которое попало в заявку
                                for (int g = 0; g <= ds1.Tables[0].Rows.Count - 1; g++)
                                {
                                    conn.Open();

                                    //Разбивка по месяцам потребности СИЗ по каждой позиции
                                    SqlCommand command5 = new SqlCommand("select VIDANO.NAIMEN_OVERALLS, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 1 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as JAN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 2 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as FEB, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 3 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as MAR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 4 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as APR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 5 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as MAY, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 6 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as JUN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 7 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as JUL, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 8 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as AUG, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 9 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as SEP, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 10 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as OKT, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 11 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as NOV, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 12 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as DEC " +
                                                                         "from VIDANO " +
                                                                         "where VIDANO.SOTRUDNIK_ID=@id_sot and VIDANO.NAIMEN_OVERALLS=@naim_cover " +
                                                                         "group by VIDANO.NAIMEN_OVERALLS", conn);

                                    command5.Parameters.AddWithValue("id_sot", ds.Tables[0].Rows[i][0].ToString());
                                    command5.Parameters.AddWithValue("naim_cover", ds1.Tables[0].Rows[g][0].ToString());
                                    SqlDataAdapter da5 = new SqlDataAdapter(command5);//Переменная объявлена как глобальная
                                    SqlCommandBuilder cb5 = new SqlCommandBuilder(da5);
                                    DataSet ds5 = new DataSet();
                                    da5.Fill(ds5, "VIDANO");

                                    //Если записи при выборе нет, то пропускаем этот шаг цикла
                                    if (ds5.Tables[0].Rows.Count > 0)
                                    {
                                        //Вставка данных в таблицу Заявки
                                        SqlCommand scmd5 = conn.CreateCommand();
                                        scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover1, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                        scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                        scmd5.Parameters.AddWithValue("naim_cover1", ds1.Tables[0].Rows[g][0]);
                                        scmd5.Parameters.AddWithValue("characterist", ds1.Tables[0].Rows[g][1]);
                                        scmd5.Parameters.AddWithValue("dep", ds1.Tables[0].Rows[g][4]);
                                        scmd5.Parameters.AddWithValue("br", ds1.Tables[0].Rows[g][6]);
                                        scmd5.Parameters.AddWithValue("ed_iz", ds1.Tables[0].Rows[g][5]);
                                        scmd5.Parameters.AddWithValue("kolvo", ds1.Tables[0].Rows[g][2]);
                                        scmd5.Parameters.AddWithValue("pr_ed_pl", ds1.Tables[0].Rows[g][3]);
                                        scmd5.Parameters.AddWithValue("jan", ds5.Tables[0].Rows[0][1]);
                                        scmd5.Parameters.AddWithValue("feb", ds5.Tables[0].Rows[0][2]);
                                        scmd5.Parameters.AddWithValue("mar", ds5.Tables[0].Rows[0][3]);
                                        scmd5.Parameters.AddWithValue("apr", ds5.Tables[0].Rows[0][4]);
                                        scmd5.Parameters.AddWithValue("may", ds5.Tables[0].Rows[0][5]);
                                        scmd5.Parameters.AddWithValue("jun", ds5.Tables[0].Rows[0][6]);
                                        scmd5.Parameters.AddWithValue("jul", ds5.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("aug", ds5.Tables[0].Rows[0][8]);
                                        scmd5.Parameters.AddWithValue("sep", ds5.Tables[0].Rows[0][9]);
                                        scmd5.Parameters.AddWithValue("okt", ds5.Tables[0].Rows[0][10]);
                                        scmd5.Parameters.AddWithValue("nov", ds5.Tables[0].Rows[0][11]);
                                        scmd5.Parameters.AddWithValue("dec", ds5.Tables[0].Rows[0][12]);
                                        scmd5.Parameters.AddWithValue("kontr", ds1.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                        scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                        scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                        scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                        SqlDataReader reader5;
                                        reader5 = scmd5.ExecuteReader();
                                        conn.Close();
                                    }
                                    else
                                    {
                                        conn.Close();
                                        continue;
                                    }

                                }
                            }


                            //Получаем то, что осталось к выдаче, но пока не выдано
                            //Определение нормы сотрудника
                            SqlCommand cmd = new SqlCommand("select norms_list.* from norms_personal,norms_list where norms_personal.norms_id=norms_list.norms_id and norms_personal.sotrudnik_id=@id_sotr1", conn);
                            cmd.Parameters.AddWithValue("id_sotr1", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da2 = new SqlDataAdapter(cmd);//Переменная объявлена как глобальная
                            SqlCommandBuilder cb2 = new SqlCommandBuilder(da2);
                            DataSet ds2 = new DataSet();
                            da2.Fill(ds2, "NORMS_LIST");
                            ///////////////////////////////

                            //Определение что уже выдано сотруднику
                            SqlCommand cmd1 = new SqlCommand("select VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID, NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and SOTRUDNIKI.ID=@id_sotr2 group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN, VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            cmd1.Parameters.AddWithValue("id_sotr2", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da3 = new SqlDataAdapter(cmd1);
                            SqlCommandBuilder cb3 = new SqlCommandBuilder(da3);
                            DataSet ds3 = new DataSet();
                            da3.Fill(ds3, "VIDANO");
                            ///////////////////////////////

                            ///////////////////////////////////////////////////
                            for (int nr = 0; nr <= ds2.Tables[0].Rows.Count - 1; nr++)
                            {
                                int aa = 0;
                                string ID = Convert.ToString(ds2.Tables[0].Rows[nr][0]);
                                string NORMS_ID = Convert.ToString(ds2.Tables[0].Rows[nr][1]);
                                string NAIMEN_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string CHARACTERISTIC_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string ED_IZM = Convert.ToString(ds2.Tables[0].Rows[nr][4]);
                                string KOLVO = Convert.ToString(ds2.Tables[0].Rows[nr][5]);

                                for (int vd = 0; vd <= ds3.Tables[0].Rows.Count - 1; vd++)
                                {
                                    string SOTRUDNIK_ID = Convert.ToString(ds3.Tables[0].Rows[vd][0]);
                                    NORMS_LIST_ID = Convert.ToString(ds3.Tables[0].Rows[vd][1]);
                                    string NAIMEN_OVERALLS1 = Convert.ToString(ds3.Tables[0].Rows[vd][2]);
                                    string DEPARTMENT = Convert.ToString(ds3.Tables[0].Rows[vd][6]);
                                    string BRANCH = Convert.ToString(ds3.Tables[0].Rows[vd][8]);
                                    string ED_IZM1 = Convert.ToString(ds3.Tables[0].Rows[vd][3]);
                                    string KOLVO1 = Convert.ToString(ds3.Tables[0].Rows[vd][4]);
                                    string PERIOD_ISPOLZOVAN1 = Convert.ToString(ds3.Tables[0].Rows[vd][5]);

                                    //Поиск в нормах того, что уже отпущено для вычисления кол-ва остатка по норме (то, что еще можно выдать)
                                    if (ID == NORMS_LIST_ID)
                                    {
                                        int ostatok = Convert.ToInt16(KOLVO) - Convert.ToInt16(KOLVO1); //Вычисление остатка от нормы, возможного к отпуску
                                        if (ostatok != 0)//Если остаток по позиции не 0, то выводим его в datagridview
                                        {
                                            //Вставка данных в таблицу Заявки
                                            conn.Open();
                                            SqlCommand scmd5 = conn.CreateCommand();
                                            scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                            scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                            scmd5.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                            scmd5.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                            scmd5.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                            scmd5.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                            scmd5.Parameters.AddWithValue("ed_iz", ED_IZM);
                                            scmd5.Parameters.AddWithValue("kolvo", ostatok);
                                            scmd5.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                            scmd5.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                            scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                            scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                            scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                            scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                            //Проверяем какой месяц выбран и исходя из этого кол-во невыданных СИЗов помещаем на выбранный месяц 
                                            if (comboBox1.Text == "Январь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", ostatok); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Февраль")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", ostatok);
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Март")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", ostatok);
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Апрель")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", ostatok);
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Май")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", ostatok);
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Июнь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", ostatok);
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Июль")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", ostatok);
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Август")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", ostatok);
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Сентябрь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", ostatok);
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Октябрь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", ostatok);
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Ноябрь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", ostatok);
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Декабрь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", ostatok);
                                            }
                                            SqlDataReader reader5;
                                            reader5 = scmd5.ExecuteReader();
                                            conn.Close();
                                        }
                                        aa++;
                                    }
                                }

                                if (aa == 0)
                                {
                                    //Вставка данных в таблицу Заявки
                                    conn.Open();
                                    SqlCommand scmd6 = conn.CreateCommand();
                                    scmd6.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                    scmd6.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                    scmd6.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                    scmd6.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                    scmd6.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                    scmd6.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                    scmd6.Parameters.AddWithValue("ed_iz", ED_IZM);
                                    scmd6.Parameters.AddWithValue("kolvo", KOLVO);
                                    scmd6.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                    scmd6.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                    scmd6.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                    scmd6.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                    scmd6.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                    scmd6.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                    //Проверяем какой месяц выбран и исходя из этого кол-во невыданных СИЗов помещаем на выбранный месяц 
                                    if (comboBox1.Text == "Январь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", KOLVO); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Февраль")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", KOLVO);
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Март")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", KOLVO);
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Апрель")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", KOLVO);
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Май")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", KOLVO);
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Июнь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", KOLVO);
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Июль")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", KOLVO);
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Август")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", KOLVO);
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Сентябрь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", KOLVO);
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Октябрь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", KOLVO);
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Ноябрь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", KOLVO);
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Декабрь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", KOLVO);
                                    }
                                    SqlDataReader reader6;
                                    reader6 = scmd6.ExecuteReader();
                                    conn.Close();
                                }

                            }
                            ////////////////////////////////////////////////////////////////

                            progressBar1.Value++;
                        }

                        SystemSounds.Beep.Play();
                        MessageBox.Show("Заявка сформирована. Просмотреть ее возможно в списке заявок.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information); 

                    }
                    else
                    {
                        conn.Open();
                        //Получаем список всех сотрудников. 
                        SqlCommand command = new SqlCommand("select * from sotrudniki where users_branch_id=@br and department=@dp", conn);
                        command.Parameters.AddWithValue("br", Form2.branch);
                        command.Parameters.AddWithValue("dp", comboBox7.Text);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        SqlCommandBuilder cb = new SqlCommandBuilder(da);
                        DataSet ds = new DataSet();
                        //conn.Close();
                        da.Fill(ds, "SOTRUDNIKI");
                        ///////////////////////////////////////////////////////////////

                        //Генерируем ключа GUID для связи заголовка заявки с его телом
                        SqlCommand command6 = new SqlCommand("select NEWID()", conn);
                        SqlDataAdapter da6 = new SqlDataAdapter(command6);//Переменная объявлена как глобальная
                        SqlCommandBuilder cb6 = new SqlCommandBuilder(da6);
                        DataSet ds6 = new DataSet();
                        //Заполнение DataGridView наименованиями полей 
                        da6.Fill(ds6, "NEWID");
                        /////////////////////////////

                        //Вставка данных в таблицу Заголовок Заявки
                        SqlCommand scmd4 = conn.CreateCommand();
                        scmd4.CommandText = "declare @ID uniqueidentifier " +
                                            "set @ID=NEWID() " +
                                            "INSERT into ZAYAVKA_MAIN (GUID, USER_ID, KIND_ZAYAV, PODR, PERIOD, DATETIME_CREATE) VALUES (@guid, @us_id, @kd_zayav, @podr, @per, GETDATE())";

                        scmd4.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                        scmd4.Parameters.AddWithValue("us_id", Form2.val);
                        scmd4.Parameters.AddWithValue("kd_zayav", "Месячная");
                        scmd4.Parameters.AddWithValue("podr", comboBox7.Text);
                        scmd4.Parameters.AddWithValue("per", comboBox1.Text);
                        SqlDataReader reader4;
                        reader4 = scmd4.ExecuteReader();
                        /////////////////////////////////

                        progressBar1.Maximum = ds.Tables[0].Rows.Count;
                        progressBar1.Value = 0;

                        //Идем циклом по сформированному списку сотрудников
                        for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                        {
                            //Получаем то, что выдано сотруднику
                            SqlCommand command1 = new SqlCommand("select NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and DATEPART(year,VIDANO.DATE_SLED_VIDACHI)=@year and DATEPART(month,VIDANO.DATE_SLED_VIDACHI)=@month and SOTRUDNIKI.ID=@id_sotr group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            command1.Parameters.AddWithValue("id_sotr", ds.Tables[0].Rows[i][0].ToString());
                            command1.Parameters.AddWithValue("year", year);
                            command1.Parameters.AddWithValue("month", month);
                            SqlDataAdapter da1 = new SqlDataAdapter(command1);
                            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                            DataSet ds1 = new DataSet();
                            conn.Close();
                            da1.Fill(ds1, "VIDANO");
                            ////////////////////////////////////////////////////////////////

                            //Проверяем выдачу, если сотруднику ничего не выдано, т.е. кол-во записей 0, то этот блок не выполняется, т.к. разбивать по месяцам нечего. Сразу переходим к следующему блоку, в котором проверяется, что не было выдано сотруднику.
                            if (ds1.Tables[0].Rows.Count != 0)
                            {
                                //Разбиваем по месяцам кол-во , которое попало в заявку
                                for (int g = 0; g <= ds1.Tables[0].Rows.Count - 1; g++)
                                {
                                    conn.Open();

                                    //Разбивка по месяцам потребности СИЗ по каждой позиции
                                    SqlCommand command5 = new SqlCommand("select VIDANO.NAIMEN_OVERALLS, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 1 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as JAN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 2 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as FEB, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 3 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as MAR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 4 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as APR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 5 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as MAY, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 6 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as JUN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 7 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as JUL, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 8 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as AUG, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 9 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as SEP, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 10 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as OKT, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 11 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as NOV, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 12 and YEAR(DATE_SLED_VIDACHI)=" + year + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI)=" + month + " then KOLVO else 0 end) as DEC " +
                                                                         "from VIDANO " +
                                                                         "where VIDANO.SOTRUDNIK_ID=@id_sot and VIDANO.NAIMEN_OVERALLS=@naim_cover " +
                                                                         "group by VIDANO.NAIMEN_OVERALLS", conn);

                                    command5.Parameters.AddWithValue("id_sot", ds.Tables[0].Rows[i][0].ToString());
                                    command5.Parameters.AddWithValue("naim_cover", ds1.Tables[0].Rows[g][0].ToString());
                                    SqlDataAdapter da5 = new SqlDataAdapter(command5);//Переменная объявлена как глобальная
                                    SqlCommandBuilder cb5 = new SqlCommandBuilder(da5);
                                    DataSet ds5 = new DataSet();
                                    da5.Fill(ds5, "VIDANO");

                                    //Если записи при выборе нет, то пропускаем этот шаг цикла
                                    if (ds5.Tables[0].Rows.Count > 0)
                                    {
                                        //Вставка данных в таблицу Заявки
                                        SqlCommand scmd5 = conn.CreateCommand();
                                        scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover1, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                        scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                        scmd5.Parameters.AddWithValue("naim_cover1", ds1.Tables[0].Rows[g][0]);
                                        scmd5.Parameters.AddWithValue("characterist", ds1.Tables[0].Rows[g][1]);
                                        scmd5.Parameters.AddWithValue("dep", ds1.Tables[0].Rows[g][4]);
                                        scmd5.Parameters.AddWithValue("br", ds1.Tables[0].Rows[g][6]);
                                        scmd5.Parameters.AddWithValue("ed_iz", ds1.Tables[0].Rows[g][5]);
                                        scmd5.Parameters.AddWithValue("kolvo", ds1.Tables[0].Rows[g][2]);
                                        scmd5.Parameters.AddWithValue("pr_ed_pl", ds1.Tables[0].Rows[g][3]);
                                        scmd5.Parameters.AddWithValue("jan", ds5.Tables[0].Rows[0][1]);
                                        scmd5.Parameters.AddWithValue("feb", ds5.Tables[0].Rows[0][2]);
                                        scmd5.Parameters.AddWithValue("mar", ds5.Tables[0].Rows[0][3]);
                                        scmd5.Parameters.AddWithValue("apr", ds5.Tables[0].Rows[0][4]);
                                        scmd5.Parameters.AddWithValue("may", ds5.Tables[0].Rows[0][5]);
                                        scmd5.Parameters.AddWithValue("jun", ds5.Tables[0].Rows[0][6]);
                                        scmd5.Parameters.AddWithValue("jul", ds5.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("aug", ds5.Tables[0].Rows[0][8]);
                                        scmd5.Parameters.AddWithValue("sep", ds5.Tables[0].Rows[0][9]);
                                        scmd5.Parameters.AddWithValue("okt", ds5.Tables[0].Rows[0][10]);
                                        scmd5.Parameters.AddWithValue("nov", ds5.Tables[0].Rows[0][11]);
                                        scmd5.Parameters.AddWithValue("dec", ds5.Tables[0].Rows[0][12]);
                                        scmd5.Parameters.AddWithValue("kontr", ds1.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                        scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                        scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                        scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                        SqlDataReader reader5;
                                        reader5 = scmd5.ExecuteReader();
                                        conn.Close();
                                    }
                                    else
                                    {
                                        conn.Close();
                                        continue;
                                    }

                                }
                            }


                            //Получаем то, что осталось к выдаче, но пока не выдано
                            //Определение нормы сотрудника
                            SqlCommand cmd = new SqlCommand("select norms_list.* from norms_personal,norms_list where norms_personal.norms_id=norms_list.norms_id and norms_personal.sotrudnik_id=@id_sotr1", conn);
                            cmd.Parameters.AddWithValue("id_sotr1", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da2 = new SqlDataAdapter(cmd);//Переменная объявлена как глобальная
                            SqlCommandBuilder cb2 = new SqlCommandBuilder(da2);
                            DataSet ds2 = new DataSet();
                            da2.Fill(ds2, "NORMS_LIST");
                            ///////////////////////////////

                            //Определение что уже выдано сотруднику
                            SqlCommand cmd1 = new SqlCommand("select VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID, NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and SOTRUDNIKI.ID=@id_sotr2 group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN, VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            cmd1.Parameters.AddWithValue("id_sotr2", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da3 = new SqlDataAdapter(cmd1);
                            SqlCommandBuilder cb3 = new SqlCommandBuilder(da3);
                            DataSet ds3 = new DataSet();
                            da3.Fill(ds3, "VIDANO");
                            ///////////////////////////////

                            ///////////////////////////////////////////////////
                            for (int nr = 0; nr <= ds2.Tables[0].Rows.Count - 1; nr++)
                            {
                                int aa = 0;
                                string ID = Convert.ToString(ds2.Tables[0].Rows[nr][0]);
                                string NORMS_ID = Convert.ToString(ds2.Tables[0].Rows[nr][1]);
                                string NAIMEN_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string CHARACTERISTIC_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string ED_IZM = Convert.ToString(ds2.Tables[0].Rows[nr][4]);
                                string KOLVO = Convert.ToString(ds2.Tables[0].Rows[nr][5]);

                                for (int vd = 0; vd <= ds3.Tables[0].Rows.Count - 1; vd++)
                                {
                                    string SOTRUDNIK_ID = Convert.ToString(ds3.Tables[0].Rows[vd][0]);
                                    NORMS_LIST_ID = Convert.ToString(ds3.Tables[0].Rows[vd][1]);
                                    string NAIMEN_OVERALLS1 = Convert.ToString(ds3.Tables[0].Rows[vd][2]);
                                    string DEPARTMENT = Convert.ToString(ds3.Tables[0].Rows[vd][6]);
                                    string BRANCH = Convert.ToString(ds3.Tables[0].Rows[vd][8]);
                                    string ED_IZM1 = Convert.ToString(ds3.Tables[0].Rows[vd][3]);
                                    string KOLVO1 = Convert.ToString(ds3.Tables[0].Rows[vd][4]);
                                    string PERIOD_ISPOLZOVAN1 = Convert.ToString(ds3.Tables[0].Rows[vd][5]);

                                    //Поиск в нормах того, что уже отпущено для вычисления кол-ва остатка по норме (то, что еще можно выдать)
                                    if (ID == NORMS_LIST_ID)
                                    {
                                        int ostatok = Convert.ToInt16(KOLVO) - Convert.ToInt16(KOLVO1); //Вычисление остатка от нормы, возможного к отпуску
                                        if (ostatok != 0)//Если остаток по позиции не 0, то выводим его в datagridview
                                        {
                                            //Вставка данных в таблицу Заявки
                                            conn.Open();
                                            SqlCommand scmd5 = conn.CreateCommand();
                                            scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                            scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                            scmd5.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                            scmd5.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                            scmd5.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                            scmd5.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                            scmd5.Parameters.AddWithValue("ed_iz", ED_IZM);
                                            scmd5.Parameters.AddWithValue("kolvo", ostatok);
                                            scmd5.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                            scmd5.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                            scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                            scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                            scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                            scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                            //Проверяем какой месяц выбран и исходя из этого кол-во невыданных СИЗов помещаем на выбранный месяц 
                                            if (comboBox1.Text == "Январь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", ostatok); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Февраль")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", ostatok);
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Март")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", ostatok);
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Апрель")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", ostatok);
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Май")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", ostatok);
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Июнь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", ostatok);
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Июль")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", ostatok);
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Август")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", ostatok);
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Сентябрь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", ostatok);
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Октябрь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", ostatok);
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Ноябрь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", ostatok);
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox1.Text == "Декабрь")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", ostatok);
                                            }
                                            SqlDataReader reader5;
                                            reader5 = scmd5.ExecuteReader();
                                            conn.Close();
                                        }
                                        aa++;
                                    }
                                }

                                if (aa == 0)
                                {
                                    //Вставка данных в таблицу Заявки
                                    conn.Open();
                                    SqlCommand scmd6 = conn.CreateCommand();
                                    scmd6.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                    scmd6.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                    scmd6.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                    scmd6.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                    scmd6.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                    scmd6.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                    scmd6.Parameters.AddWithValue("ed_iz", ED_IZM);
                                    scmd6.Parameters.AddWithValue("kolvo", KOLVO);
                                    scmd6.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                    scmd6.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                    scmd6.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                    scmd6.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                    scmd6.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                    scmd6.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                    //Проверяем какой месяц выбран и исходя из этого кол-во невыданных СИЗов помещаем на выбранный месяц 
                                    if (comboBox1.Text == "Январь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", KOLVO); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Февраль")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", KOLVO);
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Март")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", KOLVO);
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Апрель")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", KOLVO);
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Май")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", KOLVO);
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Июнь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", KOLVO);
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Июль")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", KOLVO);
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Август")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", KOLVO);
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Сентябрь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", KOLVO);
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Октябрь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", KOLVO);
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Ноябрь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", KOLVO);
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox1.Text == "Декабрь")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В месячной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в месяц формируемого месяца (на текущий год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", KOLVO);
                                    }
                                    SqlDataReader reader6;
                                    reader6 = scmd6.ExecuteReader();
                                    conn.Close();
                                }

                            }
                            ////////////////////////////////////////////////////////////////

                            progressBar1.Value++;
                        }

                        SystemSounds.Beep.Play();
                        MessageBox.Show("Заявка сформирована. Просмотреть ее возможно в списке заявок.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                        
                    }
                }
            }
             
             
        }

        private void kind_zayavka_Load(object sender, EventArgs e)
        {
            if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2) || (Form2.val == 1)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ или шепилева
            {
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
                        comboBox4.Items.Add(result2);
                        comboBox5.Items.Add(result2);
                        comboBox7.Items.Add(result2);
                    }
                    catch { }

                }
                conn.Close();
                ////////////////////////
            }
            else
            {
                //Заполнение поля Отдел, данными из БД
                SqlCommand command3 = conn.CreateCommand();
                command3.CommandText = "select distinct DEPARTMENT from SOTRUDNIKI where SOTRUDNIKI.USERS_BRANCH_ID=" + Form2.branch + " order by DEPARTMENT";
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
                        comboBox4.Items.Add(result2);
                        comboBox5.Items.Add(result2);
                        comboBox7.Items.Add(result2);
                    }
                    catch { }

                }
                conn.Close();
                ////////////////////////
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if ((checkBox1.Checked==false) && (comboBox4.Text.Length == 0) || (comboBox3.Text.Length == 0))
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не выбрано подразделение или год!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            int dt_y = DateTime.Today.Year;

            if (Convert.ToInt16(comboBox3.Text) < dt_y)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Формирование потребности прошлым периодом запрещено!!!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            
            if (MessageBox.Show("Сформировать потребность?", "Вопрос", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {               
                progressBar1.Value = 0;

                if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2) || (Form2.val == 1)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ или Шепилева
                {
                    //Если выбрано по всем подразделениям
                    if (checkBox1.Checked == true)
                    {
                        conn.Open();
                        //Получаем список всех сотрудников. 
                        SqlCommand command = new SqlCommand("select * from sotrudniki", conn);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        SqlCommandBuilder cb = new SqlCommandBuilder(da);
                        DataSet ds = new DataSet();
                        //conn.Close();
                        da.Fill(ds, "SOTRUDNIKI");
                        ///////////////////////////////////////////////////////////////

                        //Генерируем ключа GUID для связи заголовка заявки с его телом
                        SqlCommand command6 = new SqlCommand("select NEWID()", conn);
                        SqlDataAdapter da6 = new SqlDataAdapter(command6);//Переменная объявлена как глобальная
                        SqlCommandBuilder cb6 = new SqlCommandBuilder(da6);
                        DataSet ds6 = new DataSet();
                        //Заполнение DataGridView наименованиями полей 
                        da6.Fill(ds6, "NEWID");
                        /////////////////////////////

                        //Вставка данных в таблицу Заголовок Заявки
                        SqlCommand scmd4 = conn.CreateCommand();
                        scmd4.CommandText = "declare @ID uniqueidentifier " +
                                            "set @ID=NEWID() " +
                                            "INSERT into ZAYAVKA_MAIN (GUID, USER_ID, KIND_ZAYAV, PODR, PERIOD, DATETIME_CREATE) VALUES (@guid, @us_id, @kd_zayav, @podr, @per, GETDATE())";

                        scmd4.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                        scmd4.Parameters.AddWithValue("us_id", Form2.val);
                        scmd4.Parameters.AddWithValue("kd_zayav", "Годовая");
                        scmd4.Parameters.AddWithValue("podr", comboBox4.Text);
                        scmd4.Parameters.AddWithValue("per", comboBox3.Text);
                        SqlDataReader reader4;
                        reader4 = scmd4.ExecuteReader();
                        //conn.Close();
                        /////////////////////////////////

                        progressBar1.Maximum = ds.Tables[0].Rows.Count;
                        progressBar1.Value = 0;

                        //Идем циклом по сформированному списку сотрудников
                        for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                        {
                            //Получаем то, что выдано сотруднику
                            SqlCommand command1 = new SqlCommand("select NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and DATEPART(year,VIDANO.DATE_SLED_VIDACHI)=@year and SOTRUDNIKI.ID=@id_sotr group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            command1.Parameters.AddWithValue("id_sotr", ds.Tables[0].Rows[i][0].ToString());
                            command1.Parameters.AddWithValue("year", comboBox3.Text);
                            SqlDataAdapter da1 = new SqlDataAdapter(command1);
                            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                            DataSet ds1 = new DataSet();
                            conn.Close();
                            da1.Fill(ds1, "VIDANO");
                            ////////////////////////////////////////////////////////////////

                            //Проверяем выдачу, если сотруднику ничего не выдано, т.е. кол-во записей 0, то этот блок не выполняется, т.к. разбивать по месяцам нечего. Сразу переходим к следующему блоку, в котором проверяется, что не было выдано сотруднику.
                            if (ds1.Tables[0].Rows.Count != 0)
                            {
                                //Разбиваем по месяцам кол-во , которое попало в заявку
                                for (int g = 0; g <= ds1.Tables[0].Rows.Count - 1; g++)
                                {
                                    conn.Open();

                                    //Разбивка по месяцам потребности СИЗ по каждой позиции
                                    SqlCommand command5 = new SqlCommand("select VIDANO.NAIMEN_OVERALLS, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 1 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as JAN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 2 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as FEB, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 3 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as MAR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 4 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as APR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 5 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as MAY, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 6 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as JUN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 7 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as JUL, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 8 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as AUG, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 9 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as SEP, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 10 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as OKT, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 11 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as NOV, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 12 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as DEC " +
                                                                         "from VIDANO " +
                                                                         "where VIDANO.SOTRUDNIK_ID=@id_sot and VIDANO.NAIMEN_OVERALLS=@naim_cover " +
                                                                         "group by VIDANO.NAIMEN_OVERALLS", conn);

                                    command5.Parameters.AddWithValue("id_sot", ds.Tables[0].Rows[i][0].ToString());
                                    command5.Parameters.AddWithValue("naim_cover", ds1.Tables[0].Rows[g][0].ToString());
                                    SqlDataAdapter da5 = new SqlDataAdapter(command5);//Переменная объявлена как глобальная
                                    SqlCommandBuilder cb5 = new SqlCommandBuilder(da5);
                                    DataSet ds5 = new DataSet();
                                    da5.Fill(ds5, "VIDANO");

                                    //Если записи при выборе нет, то пропускаем этот шаг цикла
                                    if (ds5.Tables[0].Rows.Count > 0)
                                    {
                                        //Вставка данных в таблицу Заявки
                                        SqlCommand scmd5 = conn.CreateCommand();
                                        scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover1, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                        scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                        scmd5.Parameters.AddWithValue("naim_cover1", ds1.Tables[0].Rows[g][0]);
                                        scmd5.Parameters.AddWithValue("characterist", ds1.Tables[0].Rows[g][1]);
                                        scmd5.Parameters.AddWithValue("dep", ds1.Tables[0].Rows[g][4]);
                                        scmd5.Parameters.AddWithValue("br", ds1.Tables[0].Rows[g][6]);
                                        scmd5.Parameters.AddWithValue("ed_iz", ds1.Tables[0].Rows[g][5]);
                                        scmd5.Parameters.AddWithValue("kolvo", ds1.Tables[0].Rows[g][2]);
                                        scmd5.Parameters.AddWithValue("pr_ed_pl", ds1.Tables[0].Rows[g][3]);
                                        scmd5.Parameters.AddWithValue("jan", ds5.Tables[0].Rows[0][1]);
                                        scmd5.Parameters.AddWithValue("feb", ds5.Tables[0].Rows[0][2]);
                                        scmd5.Parameters.AddWithValue("mar", ds5.Tables[0].Rows[0][3]);
                                        scmd5.Parameters.AddWithValue("apr", ds5.Tables[0].Rows[0][4]);
                                        scmd5.Parameters.AddWithValue("may", ds5.Tables[0].Rows[0][5]);
                                        scmd5.Parameters.AddWithValue("jun", ds5.Tables[0].Rows[0][6]);
                                        scmd5.Parameters.AddWithValue("jul", ds5.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("aug", ds5.Tables[0].Rows[0][8]);
                                        scmd5.Parameters.AddWithValue("sep", ds5.Tables[0].Rows[0][9]);
                                        scmd5.Parameters.AddWithValue("okt", ds5.Tables[0].Rows[0][10]);
                                        scmd5.Parameters.AddWithValue("nov", ds5.Tables[0].Rows[0][11]);
                                        scmd5.Parameters.AddWithValue("dec", ds5.Tables[0].Rows[0][12]);
                                        scmd5.Parameters.AddWithValue("kontr", ds1.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                        scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                        scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                        scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);                                        
                                        SqlDataReader reader5;
                                        reader5 = scmd5.ExecuteReader();
                                        conn.Close();
                                    }
                                    else
                                    {
                                        conn.Close();
                                        continue;
                                    }

                                }
                            }


                            //Получаем то, что осталось к выдаче, но пока не выдано
                            //Определение нормы сотрудника
                            SqlCommand cmd = new SqlCommand("select norms_list.* from norms_personal,norms_list where norms_personal.norms_id=norms_list.norms_id and norms_personal.sotrudnik_id=@id_sotr1", conn);
                            cmd.Parameters.AddWithValue("id_sotr1", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da2 = new SqlDataAdapter(cmd);//Переменная объявлена как глобальная
                            SqlCommandBuilder cb2 = new SqlCommandBuilder(da2);
                            DataSet ds2 = new DataSet();
                            da2.Fill(ds2, "NORMS_LIST");
                            ///////////////////////////////

                            //Определение что уже выдано сотруднику
                            SqlCommand cmd1 = new SqlCommand("select VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID, NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and SOTRUDNIKI.ID=@id_sotr2 group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN, VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            cmd1.Parameters.AddWithValue("id_sotr2", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da3 = new SqlDataAdapter(cmd1);
                            SqlCommandBuilder cb3 = new SqlCommandBuilder(da3);
                            DataSet ds3 = new DataSet();
                            da3.Fill(ds3, "VIDANO");
                            ///////////////////////////////

                            ///////////////////////////////////////////////////
                            for (int nr = 0; nr <= ds2.Tables[0].Rows.Count - 1; nr++)
                            {
                                int aa = 0;
                                string ID = Convert.ToString(ds2.Tables[0].Rows[nr][0]);
                                string NORMS_ID = Convert.ToString(ds2.Tables[0].Rows[nr][1]);
                                string NAIMEN_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string CHARACTERISTIC_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string ED_IZM = Convert.ToString(ds2.Tables[0].Rows[nr][4]);
                                string KOLVO = Convert.ToString(ds2.Tables[0].Rows[nr][5]);


                                for (int vd = 0; vd <= ds3.Tables[0].Rows.Count - 1; vd++)
                                {
                                    string SOTRUDNIK_ID = Convert.ToString(ds3.Tables[0].Rows[vd][0]);
                                    NORMS_LIST_ID = Convert.ToString(ds3.Tables[0].Rows[vd][1]);
                                    string NAIMEN_OVERALLS1 = Convert.ToString(ds3.Tables[0].Rows[vd][2]);
                                    string DEPARTMENT = Convert.ToString(ds3.Tables[0].Rows[vd][6]);
                                    string BRANCH = Convert.ToString(ds3.Tables[0].Rows[vd][8]);
                                    string ED_IZM1 = Convert.ToString(ds3.Tables[0].Rows[vd][3]);
                                    string KOLVO1 = Convert.ToString(ds3.Tables[0].Rows[vd][4]);
                                    string PERIOD_ISPOLZOVAN1 = Convert.ToString(ds3.Tables[0].Rows[vd][5]);

                                    //Поиск в нормах того, что уже отпущено для вычисления кол-ва остатка по норме (то, что еще можно выдать)
                                    if (ID == NORMS_LIST_ID)
                                    {
                                        int ostatok = Convert.ToInt16(KOLVO) - Convert.ToInt16(KOLVO1); //Вычисление остатка от нормы, возможного к отпуску
                                        if (ostatok != 0)//Если остаток по позиции не 0, то выводим его в datagridview
                                        {
                                            //Вставка данных в таблицу Заявки
                                            conn.Open();
                                            SqlCommand scmd5 = conn.CreateCommand();
                                            scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                            scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                            scmd5.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                            scmd5.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                            scmd5.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                            scmd5.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                            scmd5.Parameters.AddWithValue("ed_iz", ED_IZM);
                                            scmd5.Parameters.AddWithValue("kolvo", ostatok);
                                            scmd5.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                            scmd5.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                            scmd5.Parameters.AddWithValue("jan", ostatok); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                            scmd5.Parameters.AddWithValue("feb", "0");
                                            scmd5.Parameters.AddWithValue("mar", "0");
                                            scmd5.Parameters.AddWithValue("apr", "0");
                                            scmd5.Parameters.AddWithValue("may", "0");
                                            scmd5.Parameters.AddWithValue("jun", "0");
                                            scmd5.Parameters.AddWithValue("jul", "0");
                                            scmd5.Parameters.AddWithValue("aug", "0");
                                            scmd5.Parameters.AddWithValue("sep", "0");
                                            scmd5.Parameters.AddWithValue("okt", "0");
                                            scmd5.Parameters.AddWithValue("nov", "0");
                                            scmd5.Parameters.AddWithValue("dec", "0");
                                            scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                            scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                            scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                            scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                            SqlDataReader reader5;
                                            reader5 = scmd5.ExecuteReader();
                                            conn.Close();
                                        }
                                        aa++;
                                    }
                                }

                                if (aa == 0)
                                {
                                    //Вставка данных в таблицу Заявки
                                    conn.Open();
                                    SqlCommand scmd6 = conn.CreateCommand();
                                    scmd6.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                    scmd6.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                    scmd6.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                    scmd6.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                    scmd6.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                    scmd6.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                    scmd6.Parameters.AddWithValue("ed_iz", ED_IZM);
                                    scmd6.Parameters.AddWithValue("kolvo", KOLVO);
                                    scmd6.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                    scmd6.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                    scmd6.Parameters.AddWithValue("jan", KOLVO); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                    scmd6.Parameters.AddWithValue("feb", "0");
                                    scmd6.Parameters.AddWithValue("mar", "0");
                                    scmd6.Parameters.AddWithValue("apr", "0");
                                    scmd6.Parameters.AddWithValue("may", "0");
                                    scmd6.Parameters.AddWithValue("jun", "0");
                                    scmd6.Parameters.AddWithValue("jul", "0");
                                    scmd6.Parameters.AddWithValue("aug", "0");
                                    scmd6.Parameters.AddWithValue("sep", "0");
                                    scmd6.Parameters.AddWithValue("okt", "0");
                                    scmd6.Parameters.AddWithValue("nov", "0");
                                    scmd6.Parameters.AddWithValue("dec", "0");
                                    scmd6.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                    scmd6.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                    scmd6.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                    scmd6.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                    SqlDataReader reader6;
                                    reader6 = scmd6.ExecuteReader();
                                    conn.Close();
                                }

                            }
                            ////////////////////////////////////////////////////////////////

                            progressBar1.Value++;
                        }

                        SystemSounds.Beep.Play();
                        MessageBox.Show("Заявка сформирована. Просмотреть ее возможно в списке заявок.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);                        
                    }
                    else //здесь выборка если не по всем подразделениям, а по выбраному 
                    {
                        conn.Open();
                        //Получаем список всех сотрудников. 
                        SqlCommand command = new SqlCommand("select * from sotrudniki where department=@dp", conn);
                        command.Parameters.AddWithValue("dp", comboBox4.Text);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        SqlCommandBuilder cb = new SqlCommandBuilder(da);
                        DataSet ds = new DataSet();
                        da.Fill(ds, "SOTRUDNIKI");
                        ///////////////////////////////////////////////////////////////

                        //Генерируем ключа GUID для связи заголовка заявки с его телом
                        SqlCommand command6 = new SqlCommand("select NEWID()", conn);
                        SqlDataAdapter da6 = new SqlDataAdapter(command6);//Переменная объявлена как глобальная
                        SqlCommandBuilder cb6 = new SqlCommandBuilder(da6);
                        DataSet ds6 = new DataSet();
                        //Заполнение DataGridView наименованиями полей 
                        da6.Fill(ds6, "NEWID");
                        /////////////////////////////

                        //Вставка данных в таблицу Заголовок Заявки
                        SqlCommand scmd4 = conn.CreateCommand();
                        scmd4.CommandText = "declare @ID uniqueidentifier " +
                                            "set @ID=NEWID() " +
                                            "INSERT into ZAYAVKA_MAIN (GUID, USER_ID, KIND_ZAYAV, PODR, PERIOD, DATETIME_CREATE) VALUES (@guid, @us_id, @kd_zayav, @podr, @per, GETDATE())";

                        scmd4.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                        scmd4.Parameters.AddWithValue("us_id", Form2.val);
                        scmd4.Parameters.AddWithValue("kd_zayav", "Годовая");
                        scmd4.Parameters.AddWithValue("podr", comboBox4.Text);
                        scmd4.Parameters.AddWithValue("per", comboBox3.Text);
                        SqlDataReader reader4;
                        reader4 = scmd4.ExecuteReader();
                        /////////////////////////////////

                        progressBar1.Maximum = ds.Tables[0].Rows.Count;
                        progressBar1.Value = 0;

                        //Идем циклом по сформированному списку сотрудников
                        for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                        {
                            //Получаем то, что выдано сотруднику
                            SqlCommand command1 = new SqlCommand("select NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and DATEPART(year,VIDANO.DATE_SLED_VIDACHI)=@year and SOTRUDNIKI.ID=@id_sotr group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            command1.Parameters.AddWithValue("id_sotr", ds.Tables[0].Rows[i][0].ToString());
                            command1.Parameters.AddWithValue("year", comboBox3.Text);
                            SqlDataAdapter da1 = new SqlDataAdapter(command1);
                            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                            DataSet ds1 = new DataSet();
                            conn.Close();
                            da1.Fill(ds1, "VIDANO");
                            ////////////////////////////////////////////////////////////////

                            //Проверяем выдачу, если сотруднику ничего не выдано, т.е. кол-во записей 0, то этот блок не выполняется, т.к. разбивать по месяцам нечего. Сразу переходим к следующему блоку, в котором проверяется, что не было выдано сотруднику.
                            if (ds1.Tables[0].Rows.Count != 0)
                            {
                                //Разбиваем по месяцам кол-во , которое попало в заявку
                                for (int g = 0; g <= ds1.Tables[0].Rows.Count - 1; g++)
                                {
                                    conn.Open();

                                    //Разбивка по месяцам потребности СИЗ по каждой позиции
                                    SqlCommand command5 = new SqlCommand("select VIDANO.NAIMEN_OVERALLS, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 1 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as JAN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 2 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as FEB, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 3 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as MAR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 4 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as APR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 5 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as MAY, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 6 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as JUN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 7 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as JUL, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 8 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as AUG, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 9 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as SEP, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 10 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as OKT, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 11 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as NOV, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 12 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as DEC " +
                                                                         "from VIDANO " +
                                                                         "where VIDANO.SOTRUDNIK_ID=@id_sot and VIDANO.NAIMEN_OVERALLS=@naim_cover " +
                                                                         "group by VIDANO.NAIMEN_OVERALLS", conn);

                                    command5.Parameters.AddWithValue("id_sot", ds.Tables[0].Rows[i][0].ToString());
                                    command5.Parameters.AddWithValue("naim_cover", ds1.Tables[0].Rows[g][0].ToString());
                                    SqlDataAdapter da5 = new SqlDataAdapter(command5);//Переменная объявлена как глобальная
                                    SqlCommandBuilder cb5 = new SqlCommandBuilder(da5);
                                    DataSet ds5 = new DataSet();
                                    da5.Fill(ds5, "VIDANO");

                                    //Если записи при выборе нет, то пропускаем этот шаг цикла
                                    if (ds5.Tables[0].Rows.Count > 0)
                                    {
                                        //Вставка данных в таблицу Заявки
                                        SqlCommand scmd5 = conn.CreateCommand();
                                        scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover1, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                        scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                        scmd5.Parameters.AddWithValue("naim_cover1", ds1.Tables[0].Rows[g][0]);
                                        scmd5.Parameters.AddWithValue("characterist", ds1.Tables[0].Rows[g][1]);
                                        scmd5.Parameters.AddWithValue("dep", ds1.Tables[0].Rows[g][4]);
                                        scmd5.Parameters.AddWithValue("br", ds1.Tables[0].Rows[g][6]);
                                        scmd5.Parameters.AddWithValue("ed_iz", ds1.Tables[0].Rows[g][5]);
                                        scmd5.Parameters.AddWithValue("kolvo", ds1.Tables[0].Rows[g][2]);
                                        scmd5.Parameters.AddWithValue("pr_ed_pl", ds1.Tables[0].Rows[g][3]);
                                        scmd5.Parameters.AddWithValue("jan", ds5.Tables[0].Rows[0][1]);
                                        scmd5.Parameters.AddWithValue("feb", ds5.Tables[0].Rows[0][2]);
                                        scmd5.Parameters.AddWithValue("mar", ds5.Tables[0].Rows[0][3]);
                                        scmd5.Parameters.AddWithValue("apr", ds5.Tables[0].Rows[0][4]);
                                        scmd5.Parameters.AddWithValue("may", ds5.Tables[0].Rows[0][5]);
                                        scmd5.Parameters.AddWithValue("jun", ds5.Tables[0].Rows[0][6]);
                                        scmd5.Parameters.AddWithValue("jul", ds5.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("aug", ds5.Tables[0].Rows[0][8]);
                                        scmd5.Parameters.AddWithValue("sep", ds5.Tables[0].Rows[0][9]);
                                        scmd5.Parameters.AddWithValue("okt", ds5.Tables[0].Rows[0][10]);
                                        scmd5.Parameters.AddWithValue("nov", ds5.Tables[0].Rows[0][11]);
                                        scmd5.Parameters.AddWithValue("dec", ds5.Tables[0].Rows[0][12]);
                                        scmd5.Parameters.AddWithValue("kontr", ds1.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                        scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                        scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                        scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                        SqlDataReader reader5;
                                        reader5 = scmd5.ExecuteReader();
                                        conn.Close();
                                    }
                                    else
                                    {
                                        conn.Close();
                                        continue;
                                    }

                                }
                            }


                            //Получаем то, что осталось к выдаче, но пока не выдано
                            //Определение нормы сотрудника
                            SqlCommand cmd = new SqlCommand("select norms_list.* from norms_personal,norms_list where norms_personal.norms_id=norms_list.norms_id and norms_personal.sotrudnik_id=@id_sotr1", conn);
                            cmd.Parameters.AddWithValue("id_sotr1", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da2 = new SqlDataAdapter(cmd);//Переменная объявлена как глобальная
                            SqlCommandBuilder cb2 = new SqlCommandBuilder(da2);
                            DataSet ds2 = new DataSet();
                            da2.Fill(ds2, "NORMS_LIST");
                            ///////////////////////////////

                            //Определение что уже выдано сотруднику
                            SqlCommand cmd1 = new SqlCommand("select VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID, NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and SOTRUDNIKI.ID=@id_sotr2 group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN, VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            cmd1.Parameters.AddWithValue("id_sotr2", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da3 = new SqlDataAdapter(cmd1);
                            SqlCommandBuilder cb3 = new SqlCommandBuilder(da3);
                            DataSet ds3 = new DataSet();
                            da3.Fill(ds3, "VIDANO");
                            ///////////////////////////////

                            ///////////////////////////////////////////////////
                            for (int nr = 0; nr <= ds2.Tables[0].Rows.Count - 1; nr++)
                            {
                                int aa = 0;
                                string ID = Convert.ToString(ds2.Tables[0].Rows[nr][0]);
                                string NORMS_ID = Convert.ToString(ds2.Tables[0].Rows[nr][1]);
                                string NAIMEN_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string CHARACTERISTIC_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string ED_IZM = Convert.ToString(ds2.Tables[0].Rows[nr][4]);
                                string KOLVO = Convert.ToString(ds2.Tables[0].Rows[nr][5]);

                                for (int vd = 0; vd <= ds3.Tables[0].Rows.Count - 1; vd++)
                                {
                                    string SOTRUDNIK_ID = Convert.ToString(ds3.Tables[0].Rows[vd][0]);
                                    NORMS_LIST_ID = Convert.ToString(ds3.Tables[0].Rows[vd][1]);
                                    string NAIMEN_OVERALLS1 = Convert.ToString(ds3.Tables[0].Rows[vd][2]);
                                    string DEPARTMENT = Convert.ToString(ds3.Tables[0].Rows[vd][6]);
                                    string BRANCH = Convert.ToString(ds3.Tables[0].Rows[vd][8]);
                                    string ED_IZM1 = Convert.ToString(ds3.Tables[0].Rows[vd][3]);
                                    string KOLVO1 = Convert.ToString(ds3.Tables[0].Rows[vd][4]);
                                    string PERIOD_ISPOLZOVAN1 = Convert.ToString(ds3.Tables[0].Rows[vd][5]);

                                    //Поиск в нормах того, что уже отпущено для вычисления кол-ва остатка по норме (то, что еще можно выдать)
                                    if (ID == NORMS_LIST_ID)
                                    {
                                        int ostatok = Convert.ToInt16(KOLVO) - Convert.ToInt16(KOLVO1); //Вычисление остатка от нормы, возможного к отпуску
                                        if (ostatok != 0)//Если остаток по позиции не 0, то выводим его в datagridview
                                        {
                                            //Вставка данных в таблицу Заявки
                                            conn.Open();
                                            SqlCommand scmd5 = conn.CreateCommand();
                                            scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                            scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                            scmd5.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                            scmd5.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                            scmd5.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                            scmd5.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                            scmd5.Parameters.AddWithValue("ed_iz", ED_IZM);
                                            scmd5.Parameters.AddWithValue("kolvo", ostatok);
                                            scmd5.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                            scmd5.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                            scmd5.Parameters.AddWithValue("jan", ostatok); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                            scmd5.Parameters.AddWithValue("feb", "0");
                                            scmd5.Parameters.AddWithValue("mar", "0");
                                            scmd5.Parameters.AddWithValue("apr", "0");
                                            scmd5.Parameters.AddWithValue("may", "0");
                                            scmd5.Parameters.AddWithValue("jun", "0");
                                            scmd5.Parameters.AddWithValue("jul", "0");
                                            scmd5.Parameters.AddWithValue("aug", "0");
                                            scmd5.Parameters.AddWithValue("sep", "0");
                                            scmd5.Parameters.AddWithValue("okt", "0");
                                            scmd5.Parameters.AddWithValue("nov", "0");
                                            scmd5.Parameters.AddWithValue("dec", "0");
                                            scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                            scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                            scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                            scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                            SqlDataReader reader5;
                                            reader5 = scmd5.ExecuteReader();
                                            conn.Close();
                                        }
                                        aa++;
                                    }
                                }

                                if (aa == 0)
                                {
                                    //Вставка данных в таблицу Заявки
                                    conn.Open();
                                    SqlCommand scmd6 = conn.CreateCommand();
                                    scmd6.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                    scmd6.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                    scmd6.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                    scmd6.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                    scmd6.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                    scmd6.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                    scmd6.Parameters.AddWithValue("ed_iz", ED_IZM);
                                    scmd6.Parameters.AddWithValue("kolvo", KOLVO);
                                    scmd6.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                    scmd6.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                    scmd6.Parameters.AddWithValue("jan", KOLVO); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                    scmd6.Parameters.AddWithValue("feb", "0");
                                    scmd6.Parameters.AddWithValue("mar", "0");
                                    scmd6.Parameters.AddWithValue("apr", "0");
                                    scmd6.Parameters.AddWithValue("may", "0");
                                    scmd6.Parameters.AddWithValue("jun", "0");
                                    scmd6.Parameters.AddWithValue("jul", "0");
                                    scmd6.Parameters.AddWithValue("aug", "0");
                                    scmd6.Parameters.AddWithValue("sep", "0");
                                    scmd6.Parameters.AddWithValue("okt", "0");
                                    scmd6.Parameters.AddWithValue("nov", "0");
                                    scmd6.Parameters.AddWithValue("dec", "0");
                                    scmd6.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                    scmd6.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                    scmd6.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                    scmd6.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                    SqlDataReader reader6;
                                    reader6 = scmd6.ExecuteReader();
                                    conn.Close();
                                }

                            }
                            ////////////////////////////////////////////////////////////////

                            progressBar1.Value++;
                        }

                        SystemSounds.Beep.Play();
                        MessageBox.Show("Заявка сформирована. Просмотреть ее возможно в списке заявок.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);  
                    }
                }
                else //Если заходят остальные. Все кроме Петушкова и Скворцова и Шепилевой
                {
                    //Если выбрано по всем подразделениям
                    if (checkBox1.Checked == true)
                    {
                        conn.Open();
                        //Получаем список всех сотрудников. 
                        SqlCommand command = new SqlCommand("select * from sotrudniki where users_branch_id=@br", conn);
                        command.Parameters.AddWithValue("br", Form2.branch);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        SqlCommandBuilder cb = new SqlCommandBuilder(da);
                        DataSet ds = new DataSet();
                        da.Fill(ds, "SOTRUDNIKI");
                        ///////////////////////////////////////////////////////////////

                        //Генерируем ключа GUID для связи заголовка заявки с его телом
                        SqlCommand command6 = new SqlCommand("select NEWID()", conn);
                        SqlDataAdapter da6 = new SqlDataAdapter(command6);//Переменная объявлена как глобальная
                        SqlCommandBuilder cb6 = new SqlCommandBuilder(da6);
                        DataSet ds6 = new DataSet();
                        //Заполнение DataGridView наименованиями полей 
                        da6.Fill(ds6, "NEWID");
                        /////////////////////////////

                        //Вставка данных в таблицу Заголовок Заявки
                        SqlCommand scmd4 = conn.CreateCommand();
                        scmd4.CommandText = "declare @ID uniqueidentifier " +
                                            "set @ID=NEWID() " +
                                            "INSERT into ZAYAVKA_MAIN (GUID, USER_ID, KIND_ZAYAV, PODR, PERIOD, DATETIME_CREATE) VALUES (@guid, @us_id, @kd_zayav, @podr, @per, GETDATE())";

                        scmd4.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                        scmd4.Parameters.AddWithValue("us_id", Form2.val);
                        scmd4.Parameters.AddWithValue("kd_zayav", "Годовая");
                        scmd4.Parameters.AddWithValue("podr", comboBox4.Text);
                        scmd4.Parameters.AddWithValue("per", comboBox3.Text);
                        SqlDataReader reader4;
                        reader4 = scmd4.ExecuteReader();
                        /////////////////////////////////

                        progressBar1.Maximum = ds.Tables[0].Rows.Count;
                        progressBar1.Value = 0;

                        //Идем циклом по сформированному списку сотрудников
                        for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                        {
                            //Получаем то, что выдано сотруднику
                            SqlCommand command1 = new SqlCommand("select NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and DATEPART(year,VIDANO.DATE_SLED_VIDACHI)=@year and SOTRUDNIKI.ID=@id_sotr group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            command1.Parameters.AddWithValue("id_sotr", ds.Tables[0].Rows[i][0].ToString());
                            command1.Parameters.AddWithValue("year", comboBox3.Text);
                            SqlDataAdapter da1 = new SqlDataAdapter(command1);
                            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                            DataSet ds1 = new DataSet();
                            conn.Close();
                            da1.Fill(ds1, "VIDANO");
                            ////////////////////////////////////////////////////////////////

                            //Проверяем выдачу, если сотруднику ничего не выдано, т.е. кол-во записей 0, то этот блок не выполняется, т.к. разбивать по месяцам нечего. Сразу переходим к следующему блоку, в котором проверяется, что не было выдано сотруднику.
                            if (ds1.Tables[0].Rows.Count != 0)
                            {
                                //Разбиваем по месяцам кол-во , которое попало в заявку
                                for (int g = 0; g <= ds1.Tables[0].Rows.Count - 1; g++)
                                {
                                    conn.Open();

                                    //Разбивка по месяцам потребности СИЗ по каждой позиции
                                    SqlCommand command5 = new SqlCommand("select VIDANO.NAIMEN_OVERALLS, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 1 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as JAN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 2 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as FEB, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 3 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as MAR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 4 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as APR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 5 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as MAY, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 6 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as JUN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 7 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as JUL, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 8 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as AUG, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 9 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as SEP, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 10 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as OKT, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 11 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as NOV, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 12 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as DEC " +
                                                                         "from VIDANO " +
                                                                         "where VIDANO.SOTRUDNIK_ID=@id_sot and VIDANO.NAIMEN_OVERALLS=@naim_cover " +
                                                                         "group by VIDANO.NAIMEN_OVERALLS", conn);

                                    command5.Parameters.AddWithValue("id_sot", ds.Tables[0].Rows[i][0].ToString());
                                    command5.Parameters.AddWithValue("naim_cover", ds1.Tables[0].Rows[g][0].ToString());
                                    SqlDataAdapter da5 = new SqlDataAdapter(command5);//Переменная объявлена как глобальная
                                    SqlCommandBuilder cb5 = new SqlCommandBuilder(da5);
                                    DataSet ds5 = new DataSet();
                                    da5.Fill(ds5, "VIDANO");

                                    //Если записи при выборе нет, то пропускаем этот шаг цикла
                                    if (ds5.Tables[0].Rows.Count > 0)
                                    {
                                        //Вставка данных в таблицу Заявки
                                        SqlCommand scmd5 = conn.CreateCommand();
                                        scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover1, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                        scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                        scmd5.Parameters.AddWithValue("naim_cover1", ds1.Tables[0].Rows[g][0]);
                                        scmd5.Parameters.AddWithValue("characterist", ds1.Tables[0].Rows[g][1]);
                                        scmd5.Parameters.AddWithValue("dep", ds1.Tables[0].Rows[g][4]);
                                        scmd5.Parameters.AddWithValue("br", ds1.Tables[0].Rows[g][6]);
                                        scmd5.Parameters.AddWithValue("ed_iz", ds1.Tables[0].Rows[g][5]);
                                        scmd5.Parameters.AddWithValue("kolvo", ds1.Tables[0].Rows[g][2]);
                                        scmd5.Parameters.AddWithValue("pr_ed_pl", ds1.Tables[0].Rows[g][3]);
                                        scmd5.Parameters.AddWithValue("jan", ds5.Tables[0].Rows[0][1]);
                                        scmd5.Parameters.AddWithValue("feb", ds5.Tables[0].Rows[0][2]);
                                        scmd5.Parameters.AddWithValue("mar", ds5.Tables[0].Rows[0][3]);
                                        scmd5.Parameters.AddWithValue("apr", ds5.Tables[0].Rows[0][4]);
                                        scmd5.Parameters.AddWithValue("may", ds5.Tables[0].Rows[0][5]);
                                        scmd5.Parameters.AddWithValue("jun", ds5.Tables[0].Rows[0][6]);
                                        scmd5.Parameters.AddWithValue("jul", ds5.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("aug", ds5.Tables[0].Rows[0][8]);
                                        scmd5.Parameters.AddWithValue("sep", ds5.Tables[0].Rows[0][9]);
                                        scmd5.Parameters.AddWithValue("okt", ds5.Tables[0].Rows[0][10]);
                                        scmd5.Parameters.AddWithValue("nov", ds5.Tables[0].Rows[0][11]);
                                        scmd5.Parameters.AddWithValue("dec", ds5.Tables[0].Rows[0][12]);
                                        scmd5.Parameters.AddWithValue("kontr", ds1.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                        scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                        scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                        scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                        SqlDataReader reader5;
                                        reader5 = scmd5.ExecuteReader();
                                        conn.Close();
                                    }
                                    else
                                    {
                                        conn.Close();
                                        continue;
                                    }

                                }
                            }


                            //Получаем то, что осталось к выдаче, но пока не выдано
                            //Определение нормы сотрудника
                            SqlCommand cmd = new SqlCommand("select norms_list.* from norms_personal,norms_list where norms_personal.norms_id=norms_list.norms_id and norms_personal.sotrudnik_id=@id_sotr1", conn);
                            cmd.Parameters.AddWithValue("id_sotr1", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da2 = new SqlDataAdapter(cmd);//Переменная объявлена как глобальная
                            SqlCommandBuilder cb2 = new SqlCommandBuilder(da2);
                            DataSet ds2 = new DataSet();
                            da2.Fill(ds2, "NORMS_LIST");
                            ///////////////////////////////

                            //Определение что уже выдано сотруднику
                            SqlCommand cmd1 = new SqlCommand("select VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID, NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and SOTRUDNIKI.ID=@id_sotr2 group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN, VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            cmd1.Parameters.AddWithValue("id_sotr2", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da3 = new SqlDataAdapter(cmd1);
                            SqlCommandBuilder cb3 = new SqlCommandBuilder(da3);
                            DataSet ds3 = new DataSet();
                            da3.Fill(ds3, "VIDANO");
                            ///////////////////////////////

                            ///////////////////////////////////////////////////
                            for (int nr = 0; nr <= ds2.Tables[0].Rows.Count - 1; nr++)
                            {
                                int aa = 0;
                                string ID = Convert.ToString(ds2.Tables[0].Rows[nr][0]);
                                string NORMS_ID = Convert.ToString(ds2.Tables[0].Rows[nr][1]);
                                string NAIMEN_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string CHARACTERISTIC_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string ED_IZM = Convert.ToString(ds2.Tables[0].Rows[nr][4]);
                                string KOLVO = Convert.ToString(ds2.Tables[0].Rows[nr][5]);

                                for (int vd = 0; vd <= ds3.Tables[0].Rows.Count - 1; vd++)
                                {
                                    string SOTRUDNIK_ID = Convert.ToString(ds3.Tables[0].Rows[vd][0]);
                                    NORMS_LIST_ID = Convert.ToString(ds3.Tables[0].Rows[vd][1]);
                                    string NAIMEN_OVERALLS1 = Convert.ToString(ds3.Tables[0].Rows[vd][2]);
                                    string DEPARTMENT = Convert.ToString(ds3.Tables[0].Rows[vd][6]);
                                    string BRANCH = Convert.ToString(ds3.Tables[0].Rows[vd][8]);
                                    string ED_IZM1 = Convert.ToString(ds3.Tables[0].Rows[vd][3]);
                                    string KOLVO1 = Convert.ToString(ds3.Tables[0].Rows[vd][4]);
                                    string PERIOD_ISPOLZOVAN1 = Convert.ToString(ds3.Tables[0].Rows[vd][5]);

                                    //Поиск в нормах того, что уже отпущено для вычисления кол-ва остатка по норме (то, что еще можно выдать)
                                    if (ID == NORMS_LIST_ID)
                                    {
                                        int ostatok = Convert.ToInt16(KOLVO) - Convert.ToInt16(KOLVO1); //Вычисление остатка от нормы, возможного к отпуску
                                        if (ostatok != 0)//Если остаток по позиции не 0, то выводим его в datagridview
                                        {
                                            //Вставка данных в таблицу Заявки
                                            conn.Open();
                                            SqlCommand scmd5 = conn.CreateCommand();
                                            scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                            scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                            scmd5.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                            scmd5.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                            scmd5.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                            scmd5.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                            scmd5.Parameters.AddWithValue("ed_iz", ED_IZM);
                                            scmd5.Parameters.AddWithValue("kolvo", ostatok);
                                            scmd5.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                            scmd5.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                            scmd5.Parameters.AddWithValue("jan", ostatok); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                            scmd5.Parameters.AddWithValue("feb", "0");
                                            scmd5.Parameters.AddWithValue("mar", "0");
                                            scmd5.Parameters.AddWithValue("apr", "0");
                                            scmd5.Parameters.AddWithValue("may", "0");
                                            scmd5.Parameters.AddWithValue("jun", "0");
                                            scmd5.Parameters.AddWithValue("jul", "0");
                                            scmd5.Parameters.AddWithValue("aug", "0");
                                            scmd5.Parameters.AddWithValue("sep", "0");
                                            scmd5.Parameters.AddWithValue("okt", "0");
                                            scmd5.Parameters.AddWithValue("nov", "0");
                                            scmd5.Parameters.AddWithValue("dec", "0");
                                            scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                            scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                            scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                            scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                            SqlDataReader reader5;
                                            reader5 = scmd5.ExecuteReader();
                                            conn.Close();
                                        }
                                        aa++;
                                    }
                                }

                                if (aa == 0)
                                {
                                    //Вставка данных в таблицу Заявки
                                    conn.Open();
                                    SqlCommand scmd6 = conn.CreateCommand();
                                    scmd6.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                    scmd6.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                    scmd6.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                    scmd6.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                    scmd6.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                    scmd6.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                    scmd6.Parameters.AddWithValue("ed_iz", ED_IZM);
                                    scmd6.Parameters.AddWithValue("kolvo", KOLVO);
                                    scmd6.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                    scmd6.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                    scmd6.Parameters.AddWithValue("jan", KOLVO); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                    scmd6.Parameters.AddWithValue("feb", "0");
                                    scmd6.Parameters.AddWithValue("mar", "0");
                                    scmd6.Parameters.AddWithValue("apr", "0");
                                    scmd6.Parameters.AddWithValue("may", "0");
                                    scmd6.Parameters.AddWithValue("jun", "0");
                                    scmd6.Parameters.AddWithValue("jul", "0");
                                    scmd6.Parameters.AddWithValue("aug", "0");
                                    scmd6.Parameters.AddWithValue("sep", "0");
                                    scmd6.Parameters.AddWithValue("okt", "0");
                                    scmd6.Parameters.AddWithValue("nov", "0");
                                    scmd6.Parameters.AddWithValue("dec", "0");
                                    scmd6.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                    scmd6.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                    scmd6.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                    scmd6.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                    SqlDataReader reader6;
                                    reader6 = scmd6.ExecuteReader();
                                    conn.Close();
                                }

                            }
                            ////////////////////////////////////////////////////////////////

                            progressBar1.Value++;
                        }

                        SystemSounds.Beep.Play();
                        MessageBox.Show("Заявка сформирована. Просмотреть ее возможно в списке заявок.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);     

                    }
                    else
                    {
                        conn.Open();
                        //Получаем список всех сотрудников. 
                        SqlCommand command = new SqlCommand("select * from sotrudniki where users_branch_id=@br and department=@dp", conn);
                        command.Parameters.AddWithValue("br", Form2.branch);
                        command.Parameters.AddWithValue("dp", comboBox4.Text);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        SqlCommandBuilder cb = new SqlCommandBuilder(da);
                        DataSet ds = new DataSet();
                        da.Fill(ds, "SOTRUDNIKI");
                        ///////////////////////////////////////////////////////////////

                        //Генерируем ключа GUID для связи заголовка заявки с его телом
                        SqlCommand command6 = new SqlCommand("select NEWID()", conn);
                        SqlDataAdapter da6 = new SqlDataAdapter(command6);//Переменная объявлена как глобальная
                        SqlCommandBuilder cb6 = new SqlCommandBuilder(da6);
                        DataSet ds6 = new DataSet();
                        //Заполнение DataGridView наименованиями полей 
                        da6.Fill(ds6, "NEWID");
                        /////////////////////////////

                        //Вставка данных в таблицу Заголовок Заявки
                        SqlCommand scmd4 = conn.CreateCommand();
                        scmd4.CommandText = "declare @ID uniqueidentifier " +
                                            "set @ID=NEWID() " +
                                            "INSERT into ZAYAVKA_MAIN (GUID, USER_ID, KIND_ZAYAV, PODR, PERIOD, DATETIME_CREATE) VALUES (@guid, @us_id, @kd_zayav, @podr, @per, GETDATE())";

                        scmd4.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                        scmd4.Parameters.AddWithValue("us_id", Form2.val);
                        scmd4.Parameters.AddWithValue("kd_zayav", "Годовая");
                        scmd4.Parameters.AddWithValue("podr", comboBox4.Text);
                        scmd4.Parameters.AddWithValue("per", comboBox3.Text);
                        SqlDataReader reader4;
                        reader4 = scmd4.ExecuteReader();
                        /////////////////////////////////

                        progressBar1.Maximum = ds.Tables[0].Rows.Count;
                        progressBar1.Value = 0;

                        //Идем циклом по сформированному списку сотрудников
                        for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                        {
                            //Получаем то, что выдано сотруднику
                            SqlCommand command1 = new SqlCommand("select NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and DATEPART(year,VIDANO.DATE_SLED_VIDACHI)=@year and SOTRUDNIKI.ID=@id_sotr group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            command1.Parameters.AddWithValue("id_sotr", ds.Tables[0].Rows[i][0].ToString());
                            command1.Parameters.AddWithValue("year", comboBox3.Text);
                            SqlDataAdapter da1 = new SqlDataAdapter(command1);
                            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                            DataSet ds1 = new DataSet();
                            conn.Close();
                            da1.Fill(ds1, "VIDANO");
                            ////////////////////////////////////////////////////////////////

                            //Проверяем выдачу, если сотруднику ничего не выдано, т.е. кол-во записей 0, то этот блок не выполняется, т.к. разбивать по месяцам нечего. Сразу переходим к следующему блоку, в котором проверяется, что не было выдано сотруднику.
                            if (ds1.Tables[0].Rows.Count != 0)
                            {
                                //Разбиваем по месяцам кол-во , которое попало в заявку
                                for (int g = 0; g <= ds1.Tables[0].Rows.Count - 1; g++)
                                {
                                    conn.Open();

                                    //Разбивка по месяцам потребности СИЗ по каждой позиции
                                    SqlCommand command5 = new SqlCommand("select VIDANO.NAIMEN_OVERALLS, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 1 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as JAN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 2 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as FEB, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 3 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as MAR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 4 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as APR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 5 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as MAY, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 6 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as JUN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 7 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as JUL, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 8 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as AUG, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 9 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as SEP, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 10 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as OKT, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 11 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as NOV, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 12 and YEAR(DATE_SLED_VIDACHI)=" + comboBox3.Text + " then KOLVO else 0 end) as DEC " +
                                                                         "from VIDANO " +
                                                                         "where VIDANO.SOTRUDNIK_ID=@id_sot and VIDANO.NAIMEN_OVERALLS=@naim_cover " +
                                                                         "group by VIDANO.NAIMEN_OVERALLS", conn);

                                    command5.Parameters.AddWithValue("id_sot", ds.Tables[0].Rows[i][0].ToString());
                                    command5.Parameters.AddWithValue("naim_cover", ds1.Tables[0].Rows[g][0].ToString());
                                    SqlDataAdapter da5 = new SqlDataAdapter(command5);//Переменная объявлена как глобальная
                                    SqlCommandBuilder cb5 = new SqlCommandBuilder(da5);
                                    DataSet ds5 = new DataSet();
                                    da5.Fill(ds5, "VIDANO");

                                    //Если записи при выборе нет, то пропускаем этот шаг цикла
                                    if (ds5.Tables[0].Rows.Count > 0)
                                    {
                                        //Вставка данных в таблицу Заявки
                                        SqlCommand scmd5 = conn.CreateCommand();
                                        scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover1, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                        scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                        scmd5.Parameters.AddWithValue("naim_cover1", ds1.Tables[0].Rows[g][0]);
                                        scmd5.Parameters.AddWithValue("characterist", ds1.Tables[0].Rows[g][1]);
                                        scmd5.Parameters.AddWithValue("dep", ds1.Tables[0].Rows[g][4]);
                                        scmd5.Parameters.AddWithValue("br", ds1.Tables[0].Rows[g][6]);
                                        scmd5.Parameters.AddWithValue("ed_iz", ds1.Tables[0].Rows[g][5]);
                                        scmd5.Parameters.AddWithValue("kolvo", ds1.Tables[0].Rows[g][2]);
                                        scmd5.Parameters.AddWithValue("pr_ed_pl", ds1.Tables[0].Rows[g][3]);
                                        scmd5.Parameters.AddWithValue("jan", ds5.Tables[0].Rows[0][1]);
                                        scmd5.Parameters.AddWithValue("feb", ds5.Tables[0].Rows[0][2]);
                                        scmd5.Parameters.AddWithValue("mar", ds5.Tables[0].Rows[0][3]);
                                        scmd5.Parameters.AddWithValue("apr", ds5.Tables[0].Rows[0][4]);
                                        scmd5.Parameters.AddWithValue("may", ds5.Tables[0].Rows[0][5]);
                                        scmd5.Parameters.AddWithValue("jun", ds5.Tables[0].Rows[0][6]);
                                        scmd5.Parameters.AddWithValue("jul", ds5.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("aug", ds5.Tables[0].Rows[0][8]);
                                        scmd5.Parameters.AddWithValue("sep", ds5.Tables[0].Rows[0][9]);
                                        scmd5.Parameters.AddWithValue("okt", ds5.Tables[0].Rows[0][10]);
                                        scmd5.Parameters.AddWithValue("nov", ds5.Tables[0].Rows[0][11]);
                                        scmd5.Parameters.AddWithValue("dec", ds5.Tables[0].Rows[0][12]);
                                        scmd5.Parameters.AddWithValue("kontr", ds1.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                        scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                        scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                        scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                        SqlDataReader reader5;
                                        reader5 = scmd5.ExecuteReader();
                                        conn.Close();
                                    }
                                    else
                                    {
                                        conn.Close();
                                        continue;
                                    }

                                }
                            }


                            //Получаем то, что осталось к выдаче, но пока не выдано
                            //Определение нормы сотрудника
                            SqlCommand cmd = new SqlCommand("select norms_list.* from norms_personal,norms_list where norms_personal.norms_id=norms_list.norms_id and norms_personal.sotrudnik_id=@id_sotr1", conn);
                            cmd.Parameters.AddWithValue("id_sotr1", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da2 = new SqlDataAdapter(cmd);//Переменная объявлена как глобальная
                            SqlCommandBuilder cb2 = new SqlCommandBuilder(da2);
                            DataSet ds2 = new DataSet();
                            da2.Fill(ds2, "NORMS_LIST");
                            ///////////////////////////////

                            //Определение что уже выдано сотруднику
                            SqlCommand cmd1 = new SqlCommand("select VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID, NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and SOTRUDNIKI.ID=@id_sotr2 group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN, VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            cmd1.Parameters.AddWithValue("id_sotr2", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da3 = new SqlDataAdapter(cmd1);
                            SqlCommandBuilder cb3 = new SqlCommandBuilder(da3);
                            DataSet ds3 = new DataSet();
                            da3.Fill(ds3, "VIDANO");
                            ///////////////////////////////

                            ///////////////////////////////////////////////////
                            for (int nr = 0; nr <= ds2.Tables[0].Rows.Count - 1; nr++)
                            {
                                int aa = 0;
                                string ID = Convert.ToString(ds2.Tables[0].Rows[nr][0]);
                                string NORMS_ID = Convert.ToString(ds2.Tables[0].Rows[nr][1]);
                                string NAIMEN_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string CHARACTERISTIC_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string ED_IZM = Convert.ToString(ds2.Tables[0].Rows[nr][4]);
                                string KOLVO = Convert.ToString(ds2.Tables[0].Rows[nr][5]);

                                for (int vd = 0; vd <= ds3.Tables[0].Rows.Count - 1; vd++)
                                {
                                    string SOTRUDNIK_ID = Convert.ToString(ds3.Tables[0].Rows[vd][0]);
                                    NORMS_LIST_ID = Convert.ToString(ds3.Tables[0].Rows[vd][1]);
                                    string NAIMEN_OVERALLS1 = Convert.ToString(ds3.Tables[0].Rows[vd][2]);
                                    string DEPARTMENT = Convert.ToString(ds3.Tables[0].Rows[vd][6]);
                                    string BRANCH = Convert.ToString(ds3.Tables[0].Rows[vd][8]);
                                    string ED_IZM1 = Convert.ToString(ds3.Tables[0].Rows[vd][3]);
                                    string KOLVO1 = Convert.ToString(ds3.Tables[0].Rows[vd][4]);
                                    string PERIOD_ISPOLZOVAN1 = Convert.ToString(ds3.Tables[0].Rows[vd][5]);

                                    //Поиск в нормах того, что уже отпущено для вычисления кол-ва остатка по норме (то, что еще можно выдать)
                                    if (ID == NORMS_LIST_ID)
                                    {
                                        int ostatok = Convert.ToInt16(KOLVO) - Convert.ToInt16(KOLVO1); //Вычисление остатка от нормы, возможного к отпуску
                                        if (ostatok != 0)//Если остаток по позиции не 0, то выводим его в datagridview
                                        {
                                            //Вставка данных в таблицу Заявки
                                            conn.Open();
                                            SqlCommand scmd5 = conn.CreateCommand();
                                            scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                            scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                            scmd5.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                            scmd5.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                            scmd5.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                            scmd5.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                            scmd5.Parameters.AddWithValue("ed_iz", ED_IZM);
                                            scmd5.Parameters.AddWithValue("kolvo", ostatok);
                                            scmd5.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                            scmd5.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                            scmd5.Parameters.AddWithValue("jan", ostatok); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                            scmd5.Parameters.AddWithValue("feb", "0");
                                            scmd5.Parameters.AddWithValue("mar", "0");
                                            scmd5.Parameters.AddWithValue("apr", "0");
                                            scmd5.Parameters.AddWithValue("may", "0");
                                            scmd5.Parameters.AddWithValue("jun", "0");
                                            scmd5.Parameters.AddWithValue("jul", "0");
                                            scmd5.Parameters.AddWithValue("aug", "0");
                                            scmd5.Parameters.AddWithValue("sep", "0");
                                            scmd5.Parameters.AddWithValue("okt", "0");
                                            scmd5.Parameters.AddWithValue("nov", "0");
                                            scmd5.Parameters.AddWithValue("dec", "0");
                                            scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                            scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                            scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                            scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                            SqlDataReader reader5;
                                            reader5 = scmd5.ExecuteReader();
                                            conn.Close();
                                        }
                                        aa++;
                                    }
                                }

                                if (aa == 0)
                                {
                                    //Вставка данных в таблицу Заявки
                                    conn.Open();
                                    SqlCommand scmd6 = conn.CreateCommand();
                                    scmd6.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                    scmd6.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                    scmd6.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                    scmd6.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                    scmd6.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                    scmd6.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                    scmd6.Parameters.AddWithValue("ed_iz", ED_IZM);
                                    scmd6.Parameters.AddWithValue("kolvo", KOLVO);
                                    scmd6.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                    scmd6.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                    scmd6.Parameters.AddWithValue("jan", KOLVO); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                    scmd6.Parameters.AddWithValue("feb", "0");
                                    scmd6.Parameters.AddWithValue("mar", "0");
                                    scmd6.Parameters.AddWithValue("apr", "0");
                                    scmd6.Parameters.AddWithValue("may", "0");
                                    scmd6.Parameters.AddWithValue("jun", "0");
                                    scmd6.Parameters.AddWithValue("jul", "0");
                                    scmd6.Parameters.AddWithValue("aug", "0");
                                    scmd6.Parameters.AddWithValue("sep", "0");
                                    scmd6.Parameters.AddWithValue("okt", "0");
                                    scmd6.Parameters.AddWithValue("nov", "0");
                                    scmd6.Parameters.AddWithValue("dec", "0");
                                    scmd6.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                    scmd6.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                    scmd6.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                    scmd6.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                    SqlDataReader reader6;
                                    reader6 = scmd6.ExecuteReader();
                                    conn.Close();
                                }

                            }
                            ////////////////////////////////////////////////////////////////

                            progressBar1.Value++;
                        }

                        SystemSounds.Beep.Play();
                        MessageBox.Show("Заявка сформирована. Просмотреть ее возможно в списке заявок.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);   
                                                
                    }
                }
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                comboBox4.Enabled = false;
            }
            else
            {
                comboBox4.Enabled = true;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                comboBox5.Enabled = false;
            }
            else
            {
                comboBox5.Enabled = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if ((checkBox2.Checked == false) && (comboBox5.Text.Length == 0) || (comboBox2.Text.Length == 0) || (comboBox6.Text.Length == 0))
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не выбрано подразделение, месяц или год!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            int dt_y = DateTime.Today.Year;

            if (Convert.ToInt16(comboBox6.Text) < dt_y)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Формирование потребности прошлым периодом запрещено!!!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            int dt_m=DateTime.Today.Month;
            int dt_d=DateTime.Today.Day;

            if ((comboBox2.Text == "1 квартал") && (dt_m == 12) && (dt_d > 5))
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Потребность на 1-й квартал формировать уже поздно. С данного момента имеется возможность формировать потребность только на 2-й квартал.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }


            if ((comboBox2.Text == "2 квартал") && (dt_m == 3) && (dt_d > 5))
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Потребность на 2-й квартал формировать уже поздно. С данного момента имеется возможность формировать потребность только на 3-й квартал.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if ((comboBox2.Text == "3 квартал") && (dt_m == 6) && (dt_d > 5))
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Потребность на 3-й квартал формировать уже поздно. С данного момента имеется возможность формировать потребность только на 4-й квартал.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if ((comboBox2.Text == "4 квартал") && (dt_m == 9) && (dt_d > 5))
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Потребность на 4-й квартал формировать уже поздно. С данного момента имеется возможность формировать потребность только на 1-й квартал следующего года.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("Сформировать потребность?", "Вопрос", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                //Присвоение значений переменным взависимости от выбранного квартала
                if (comboBox2.Text == "1 квартал")
                {
                    beg_m = "1";
                    end_m = "3";
                    
                }
                if (comboBox2.Text == "2 квартал")
                {
                    beg_m = "4";
                    end_m = "6";
                }
                if (comboBox2.Text == "3 квартал")
                {
                    beg_m = "7";
                    end_m = "9";
                }
                if (comboBox2.Text == "4 квартал")
                {
                    beg_m = "10";
                    end_m = "12";
                }
             
                progressBar1.Value = 0;


                if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2) || (Form2.val == 1)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ
                {
                    //Если выбрано по всем подразделениям
                    if (checkBox2.Checked == true)
                    {                        
                        conn.Open();
                        //Получаем список всех сотрудников. 
                        SqlCommand command = new SqlCommand("select * from sotrudniki", conn);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        SqlCommandBuilder cb = new SqlCommandBuilder(da);
                        DataSet ds = new DataSet();
                        //conn.Close();
                        da.Fill(ds, "SOTRUDNIKI");
                        ///////////////////////////////////////////////////////////////

                        //Генерируем ключа GUID для связи заголовка заявки с его телом
                        SqlCommand command6 = new SqlCommand("select NEWID()", conn);
                        SqlDataAdapter da6 = new SqlDataAdapter(command6);//Переменная объявлена как глобальная
                        SqlCommandBuilder cb6 = new SqlCommandBuilder(da6);
                        DataSet ds6 = new DataSet();
                        //Заполнение DataGridView наименованиями полей 
                        da6.Fill(ds6, "NEWID");
                        /////////////////////////////

                        //Вставка данных в таблицу Заголовок Заявки
                        SqlCommand scmd4 = conn.CreateCommand();
                        scmd4.CommandText = "declare @ID uniqueidentifier " +
                                            "set @ID=NEWID() " +
                                            "INSERT into ZAYAVKA_MAIN (GUID, USER_ID, KIND_ZAYAV, PODR, PERIOD, DATETIME_CREATE) VALUES (@guid, @us_id, @kd_zayav, @podr, @per, GETDATE())";

                        scmd4.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                        scmd4.Parameters.AddWithValue("us_id", Form2.val);
                        scmd4.Parameters.AddWithValue("kd_zayav", "Квартальная");
                        scmd4.Parameters.AddWithValue("podr", comboBox5.Text);
                        scmd4.Parameters.AddWithValue("per", comboBox2.Text+" "+comboBox6.Text);
                        SqlDataReader reader4;
                        reader4 = scmd4.ExecuteReader();
                        /////////////////////////////////

                        progressBar1.Maximum = ds.Tables[0].Rows.Count;
                        progressBar1.Value = 0;

                        //Идем циклом по сформированному списку сотрудников
                        for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                        {
                            //Получаем то, что выдано сотруднику
                            SqlCommand command1 = new SqlCommand("select NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and DATEPART(year,VIDANO.DATE_SLED_VIDACHI)=@year and DATEPART(month,VIDANO.DATE_SLED_VIDACHI) between @bm and @em and SOTRUDNIKI.ID=@id_sotr group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            command1.Parameters.AddWithValue("id_sotr", ds.Tables[0].Rows[i][0].ToString());
                            command1.Parameters.AddWithValue("year", comboBox6.Text);
                            command1.Parameters.AddWithValue("bm", beg_m);
                            command1.Parameters.AddWithValue("em", end_m);

                            SqlDataAdapter da1 = new SqlDataAdapter(command1);
                            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                            DataSet ds1 = new DataSet();
                            conn.Close();
                            da1.Fill(ds1, "VIDANO");
                            ////////////////////////////////////////////////////////////////

                            //Проверяем выдачу, если сотруднику ничего не выдано, т.е. кол-во записей 0, то этот блок не выполняется, т.к. разбивать по месяцам нечего. Сразу переходим к следующему блоку, в котором проверяется, что не было выдано сотруднику.
                            if (ds1.Tables[0].Rows.Count != 0)
                            {
                                //Разбиваем по месяцам кол-во , которое попало в заявку
                                for (int g = 0; g <= ds1.Tables[0].Rows.Count - 1; g++)
                                {
                                    conn.Open();

                                    //Разбивка по месяцам потребности СИЗ по каждой позиции
                                    SqlCommand command5 = new SqlCommand("select VIDANO.NAIMEN_OVERALLS, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 1 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as JAN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 2 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as FEB, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 3 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as MAR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 4 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as APR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 5 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as MAY, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 6 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as JUN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 7 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as JUL, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 8 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as AUG, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 9 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as SEP, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 10 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as OKT, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 11 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as NOV, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 12 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as DEC " +
                                                                         "from VIDANO " +
                                                                         "where VIDANO.SOTRUDNIK_ID=@id_sot and VIDANO.NAIMEN_OVERALLS=@naim_cover " +
                                                                         "group by VIDANO.NAIMEN_OVERALLS", conn);

                                    command5.Parameters.AddWithValue("id_sot", ds.Tables[0].Rows[i][0].ToString());
                                    command5.Parameters.AddWithValue("naim_cover", ds1.Tables[0].Rows[g][0].ToString());
                                    SqlDataAdapter da5 = new SqlDataAdapter(command5);//Переменная объявлена как глобальная
                                    SqlCommandBuilder cb5 = new SqlCommandBuilder(da5);
                                    DataSet ds5 = new DataSet();
                                    da5.Fill(ds5, "VIDANO");

                                    //Если записи при выборе нет, то пропускаем этот шаг цикла
                                    if (ds5.Tables[0].Rows.Count > 0)
                                    {
                                        //Вставка данных в таблицу Заявки
                                        SqlCommand scmd5 = conn.CreateCommand();
                                        scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover1, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                        scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                        scmd5.Parameters.AddWithValue("naim_cover1", ds1.Tables[0].Rows[g][0]);
                                        scmd5.Parameters.AddWithValue("characterist", ds1.Tables[0].Rows[g][1]);
                                        scmd5.Parameters.AddWithValue("dep", ds1.Tables[0].Rows[g][4]);
                                        scmd5.Parameters.AddWithValue("br", ds1.Tables[0].Rows[g][6]);
                                        scmd5.Parameters.AddWithValue("ed_iz", ds1.Tables[0].Rows[g][5]);
                                        scmd5.Parameters.AddWithValue("kolvo", ds1.Tables[0].Rows[g][2]);
                                        scmd5.Parameters.AddWithValue("pr_ed_pl", ds1.Tables[0].Rows[g][3]);
                                        scmd5.Parameters.AddWithValue("jan", ds5.Tables[0].Rows[0][1]);
                                        scmd5.Parameters.AddWithValue("feb", ds5.Tables[0].Rows[0][2]);
                                        scmd5.Parameters.AddWithValue("mar", ds5.Tables[0].Rows[0][3]);
                                        scmd5.Parameters.AddWithValue("apr", ds5.Tables[0].Rows[0][4]);
                                        scmd5.Parameters.AddWithValue("may", ds5.Tables[0].Rows[0][5]);
                                        scmd5.Parameters.AddWithValue("jun", ds5.Tables[0].Rows[0][6]);
                                        scmd5.Parameters.AddWithValue("jul", ds5.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("aug", ds5.Tables[0].Rows[0][8]);
                                        scmd5.Parameters.AddWithValue("sep", ds5.Tables[0].Rows[0][9]);
                                        scmd5.Parameters.AddWithValue("okt", ds5.Tables[0].Rows[0][10]);
                                        scmd5.Parameters.AddWithValue("nov", ds5.Tables[0].Rows[0][11]);
                                        scmd5.Parameters.AddWithValue("dec", ds5.Tables[0].Rows[0][12]);
                                        scmd5.Parameters.AddWithValue("kontr", ds1.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                        scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                        scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                        scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                        SqlDataReader reader5;
                                        reader5 = scmd5.ExecuteReader();
                                        conn.Close();
                                    }
                                    else
                                    {
                                        conn.Close();
                                        continue;
                                    }

                                }
                            }


                            //Получаем то, что осталось к выдаче, но пока не выдано
                            //Определение нормы сотрудника
                            SqlCommand cmd = new SqlCommand("select norms_list.* from norms_personal,norms_list where norms_personal.norms_id=norms_list.norms_id and norms_personal.sotrudnik_id=@id_sotr1", conn);
                            cmd.Parameters.AddWithValue("id_sotr1", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da2 = new SqlDataAdapter(cmd);//Переменная объявлена как глобальная
                            SqlCommandBuilder cb2 = new SqlCommandBuilder(da2);
                            DataSet ds2 = new DataSet();
                            da2.Fill(ds2, "NORMS_LIST");
                            ///////////////////////////////

                            //Определение что уже выдано сотруднику
                            SqlCommand cmd1 = new SqlCommand("select VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID, NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and SOTRUDNIKI.ID=@id_sotr2 group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN, VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            cmd1.Parameters.AddWithValue("id_sotr2", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da3 = new SqlDataAdapter(cmd1);
                            SqlCommandBuilder cb3 = new SqlCommandBuilder(da3);
                            DataSet ds3 = new DataSet();
                            da3.Fill(ds3, "VIDANO");
                            ///////////////////////////////

                            ///////////////////////////////////////////////////
                            for (int nr = 0; nr <= ds2.Tables[0].Rows.Count - 1; nr++)
                            {
                                int aa = 0;
                                string ID = Convert.ToString(ds2.Tables[0].Rows[nr][0]);
                                string NORMS_ID = Convert.ToString(ds2.Tables[0].Rows[nr][1]);
                                string NAIMEN_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string CHARACTERISTIC_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string ED_IZM = Convert.ToString(ds2.Tables[0].Rows[nr][4]);
                                string KOLVO = Convert.ToString(ds2.Tables[0].Rows[nr][5]);

                                for (int vd = 0; vd <= ds3.Tables[0].Rows.Count - 1; vd++)
                                {
                                    string SOTRUDNIK_ID = Convert.ToString(ds3.Tables[0].Rows[vd][0]);
                                    NORMS_LIST_ID = Convert.ToString(ds3.Tables[0].Rows[vd][1]);
                                    string NAIMEN_OVERALLS1 = Convert.ToString(ds3.Tables[0].Rows[vd][2]);
                                    string DEPARTMENT = Convert.ToString(ds3.Tables[0].Rows[vd][6]);
                                    string BRANCH = Convert.ToString(ds3.Tables[0].Rows[vd][8]);
                                    string ED_IZM1 = Convert.ToString(ds3.Tables[0].Rows[vd][3]);
                                    string KOLVO1 = Convert.ToString(ds3.Tables[0].Rows[vd][4]);
                                    string PERIOD_ISPOLZOVAN1 = Convert.ToString(ds3.Tables[0].Rows[vd][5]);

                                    //Поиск в нормах того, что уже отпущено для вычисления кол-ва остатка по норме (то, что еще можно выдать)
                                    if (ID == NORMS_LIST_ID)
                                    {
                                        int ostatok = Convert.ToInt16(KOLVO) - Convert.ToInt16(KOLVO1); //Вычисление остатка от нормы, возможного к отпуску
                                        if (ostatok != 0)//Если остаток по позиции не 0, то выводим его в datagridview
                                        {
                                            //Вставка данных в таблицу Заявки
                                            conn.Open();
                                            SqlCommand scmd5 = conn.CreateCommand();
                                            scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                            scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                            scmd5.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                            scmd5.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                            scmd5.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                            scmd5.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                            scmd5.Parameters.AddWithValue("ed_iz", ED_IZM);
                                            scmd5.Parameters.AddWithValue("kolvo", ostatok);
                                            scmd5.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                            scmd5.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                            scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                            scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                            scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                            scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                            //Проверяем какой квартал выбран и исходя из этого кол-во невыданных СИЗов помещаем на первый месяц выбранного квартала 
                                            if (comboBox2.Text == "1 квартал")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", ostatok); //В квартальной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого квартала (на тот год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox2.Text == "2 квартал")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В квартальной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого квартала (на тот год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", ostatok);
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox2.Text == "3 квартал")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В квартальной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого квартала (на тот год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", ostatok);
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox2.Text == "4 квартал")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В квартальной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого квартала (на тот год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", ostatok);
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            SqlDataReader reader5;
                                            reader5 = scmd5.ExecuteReader();
                                            conn.Close();
                                        }
                                        aa++;
                                    }
                                }

                                if (aa == 0)
                                {
                                    //Вставка данных в таблицу Заявки
                                    conn.Open();
                                    SqlCommand scmd6 = conn.CreateCommand();
                                    scmd6.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                    scmd6.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                    scmd6.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                    scmd6.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                    scmd6.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                    scmd6.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                    scmd6.Parameters.AddWithValue("ed_iz", ED_IZM);
                                    scmd6.Parameters.AddWithValue("kolvo", KOLVO);
                                    scmd6.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                    scmd6.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                    scmd6.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                    scmd6.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                    scmd6.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                    scmd6.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                    //Проверяем какой квартал выбран и исходя из этого кол-во невыданных СИЗов помещаем на первый месяц выбранного квартала 
                                    if (comboBox2.Text == "1 квартал")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", KOLVO); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox2.Text == "2 квартал")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", KOLVO);
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox2.Text == "3 квартал")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", KOLVO);
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox2.Text == "4 квартал")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", KOLVO);
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    SqlDataReader reader6;
                                    reader6 = scmd6.ExecuteReader();
                                    conn.Close();
                                }

                            }
                            ////////////////////////////////////////////////////////////////

                            progressBar1.Value++;
                        }

                        SystemSounds.Beep.Play();
                        MessageBox.Show("Заявка сформирована. Просмотреть ее возможно в списке заявок.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information); 
               
                    }
                    else
                    {
                        conn.Open();
                        //Получаем список всех сотрудников. 
                        SqlCommand command = new SqlCommand("select * from sotrudniki where department=@dp", conn);
                        command.Parameters.AddWithValue("dp", comboBox5.Text);
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        SqlCommandBuilder cb = new SqlCommandBuilder(da);
                        DataSet ds = new DataSet();
                        //conn.Close();
                        da.Fill(ds, "SOTRUDNIKI");
                        ///////////////////////////////////////////////////////////////

                        //Генерируем ключа GUID для связи заголовка заявки с его телом
                        SqlCommand command6 = new SqlCommand("select NEWID()", conn);
                        SqlDataAdapter da6 = new SqlDataAdapter(command6);//Переменная объявлена как глобальная
                        SqlCommandBuilder cb6 = new SqlCommandBuilder(da6);
                        DataSet ds6 = new DataSet();
                        //Заполнение DataGridView наименованиями полей 
                        da6.Fill(ds6, "NEWID");
                        /////////////////////////////

                        //Вставка данных в таблицу Заголовок Заявки
                        SqlCommand scmd4 = conn.CreateCommand();
                        scmd4.CommandText = "declare @ID uniqueidentifier " +
                                            "set @ID=NEWID() " +
                                            "INSERT into ZAYAVKA_MAIN (GUID, USER_ID, KIND_ZAYAV, PODR, PERIOD, DATETIME_CREATE) VALUES (@guid, @us_id, @kd_zayav, @podr, @per, GETDATE())";

                        scmd4.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                        scmd4.Parameters.AddWithValue("us_id", Form2.val);
                        scmd4.Parameters.AddWithValue("kd_zayav", "Квартальная");
                        scmd4.Parameters.AddWithValue("podr", comboBox5.Text);
                        scmd4.Parameters.AddWithValue("per", comboBox2.Text + " " + comboBox6.Text);
                        SqlDataReader reader4;
                        reader4 = scmd4.ExecuteReader();
                        //conn.Close();
                        /////////////////////////////////

                        progressBar1.Maximum = ds.Tables[0].Rows.Count;
                        progressBar1.Value = 0;

                        //Идем циклом по сформированному списку сотрудников
                        for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                        {
                            //Получаем то, что выдано сотруднику
                            SqlCommand command1 = new SqlCommand("select NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and DATEPART(year,VIDANO.DATE_SLED_VIDACHI)=@year and DATEPART(month,VIDANO.DATE_SLED_VIDACHI) between @bm and @em and SOTRUDNIKI.ID=@id_sotr group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            command1.Parameters.AddWithValue("id_sotr", ds.Tables[0].Rows[i][0].ToString());
                            command1.Parameters.AddWithValue("year", comboBox6.Text);
                            command1.Parameters.AddWithValue("bm", beg_m);
                            command1.Parameters.AddWithValue("em", end_m);

                            SqlDataAdapter da1 = new SqlDataAdapter(command1);
                            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                            DataSet ds1 = new DataSet();
                            conn.Close();
                            da1.Fill(ds1, "VIDANO");
                            ////////////////////////////////////////////////////////////////

                            //Проверяем выдачу, если сотруднику ничего не выдано, т.е. кол-во записей 0, то этот блок не выполняется, т.к. разбивать по месяцам нечего. Сразу переходим к следующему блоку, в котором проверяется, что не было выдано сотруднику.
                            if (ds1.Tables[0].Rows.Count != 0)
                            {
                                //Разбиваем по месяцам кол-во , которое попало в заявку
                                for (int g = 0; g <= ds1.Tables[0].Rows.Count - 1; g++)
                                {
                                    conn.Open();

                                    //Разбивка по месяцам потребности СИЗ по каждой позиции
                                    SqlCommand command5 = new SqlCommand("select VIDANO.NAIMEN_OVERALLS, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 1 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as JAN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 2 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as FEB, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 3 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as MAR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 4 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as APR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 5 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as MAY, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 6 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as JUN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 7 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as JUL, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 8 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as AUG, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 9 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as SEP, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 10 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as OKT, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 11 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as NOV, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 12 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as DEC " +
                                                                         "from VIDANO " +
                                                                         "where VIDANO.SOTRUDNIK_ID=@id_sot and VIDANO.NAIMEN_OVERALLS=@naim_cover " +
                                                                         "group by VIDANO.NAIMEN_OVERALLS", conn);

                                    command5.Parameters.AddWithValue("id_sot", ds.Tables[0].Rows[i][0].ToString());
                                    command5.Parameters.AddWithValue("naim_cover", ds1.Tables[0].Rows[g][0].ToString());
                                    SqlDataAdapter da5 = new SqlDataAdapter(command5);//Переменная объявлена как глобальная
                                    SqlCommandBuilder cb5 = new SqlCommandBuilder(da5);
                                    DataSet ds5 = new DataSet();
                                    da5.Fill(ds5, "VIDANO");

                                    //Если записи при выборе нет, то пропускаем этот шаг цикла
                                    if (ds5.Tables[0].Rows.Count > 0)
                                    {
                                        //Вставка данных в таблицу Заявки
                                        SqlCommand scmd5 = conn.CreateCommand();
                                        scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover1, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                        scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                        scmd5.Parameters.AddWithValue("naim_cover1", ds1.Tables[0].Rows[g][0]);
                                        scmd5.Parameters.AddWithValue("characterist", ds1.Tables[0].Rows[g][1]);
                                        scmd5.Parameters.AddWithValue("dep", ds1.Tables[0].Rows[g][4]);
                                        scmd5.Parameters.AddWithValue("br", ds1.Tables[0].Rows[g][6]);
                                        scmd5.Parameters.AddWithValue("ed_iz", ds1.Tables[0].Rows[g][5]);
                                        scmd5.Parameters.AddWithValue("kolvo", ds1.Tables[0].Rows[g][2]);
                                        scmd5.Parameters.AddWithValue("pr_ed_pl", ds1.Tables[0].Rows[g][3]);
                                        scmd5.Parameters.AddWithValue("jan", ds5.Tables[0].Rows[0][1]);
                                        scmd5.Parameters.AddWithValue("feb", ds5.Tables[0].Rows[0][2]);
                                        scmd5.Parameters.AddWithValue("mar", ds5.Tables[0].Rows[0][3]);
                                        scmd5.Parameters.AddWithValue("apr", ds5.Tables[0].Rows[0][4]);
                                        scmd5.Parameters.AddWithValue("may", ds5.Tables[0].Rows[0][5]);
                                        scmd5.Parameters.AddWithValue("jun", ds5.Tables[0].Rows[0][6]);
                                        scmd5.Parameters.AddWithValue("jul", ds5.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("aug", ds5.Tables[0].Rows[0][8]);
                                        scmd5.Parameters.AddWithValue("sep", ds5.Tables[0].Rows[0][9]);
                                        scmd5.Parameters.AddWithValue("okt", ds5.Tables[0].Rows[0][10]);
                                        scmd5.Parameters.AddWithValue("nov", ds5.Tables[0].Rows[0][11]);
                                        scmd5.Parameters.AddWithValue("dec", ds5.Tables[0].Rows[0][12]);
                                        scmd5.Parameters.AddWithValue("kontr", ds1.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                        scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                        scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                        scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                        SqlDataReader reader5;
                                        reader5 = scmd5.ExecuteReader();
                                        conn.Close();
                                    }
                                    else
                                    {
                                        conn.Close();
                                        continue;
                                    }

                                }
                            }


                            //Получаем то, что осталось к выдаче, но пока не выдано
                            //Определение нормы сотрудника
                            SqlCommand cmd = new SqlCommand("select norms_list.* from norms_personal,norms_list where norms_personal.norms_id=norms_list.norms_id and norms_personal.sotrudnik_id=@id_sotr1", conn);
                            cmd.Parameters.AddWithValue("id_sotr1", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da2 = new SqlDataAdapter(cmd);//Переменная объявлена как глобальная
                            SqlCommandBuilder cb2 = new SqlCommandBuilder(da2);
                            DataSet ds2 = new DataSet();
                            da2.Fill(ds2, "NORMS_LIST");
                            ///////////////////////////////

                            //Определение что уже выдано сотруднику
                            SqlCommand cmd1 = new SqlCommand("select VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID, NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and SOTRUDNIKI.ID=@id_sotr2 group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN, VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            cmd1.Parameters.AddWithValue("id_sotr2", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da3 = new SqlDataAdapter(cmd1);
                            SqlCommandBuilder cb3 = new SqlCommandBuilder(da3);
                            DataSet ds3 = new DataSet();
                            da3.Fill(ds3, "VIDANO");
                            ///////////////////////////////

                            ///////////////////////////////////////////////////
                            for (int nr = 0; nr <= ds2.Tables[0].Rows.Count - 1; nr++)
                            {
                                int aa = 0;
                                string ID = Convert.ToString(ds2.Tables[0].Rows[nr][0]);
                                string NORMS_ID = Convert.ToString(ds2.Tables[0].Rows[nr][1]);
                                string NAIMEN_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string CHARACTERISTIC_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string ED_IZM = Convert.ToString(ds2.Tables[0].Rows[nr][4]);
                                string KOLVO = Convert.ToString(ds2.Tables[0].Rows[nr][5]);

                                for (int vd = 0; vd <= ds3.Tables[0].Rows.Count - 1; vd++)
                                {
                                    string SOTRUDNIK_ID = Convert.ToString(ds3.Tables[0].Rows[vd][0]);
                                    NORMS_LIST_ID = Convert.ToString(ds3.Tables[0].Rows[vd][1]);
                                    string NAIMEN_OVERALLS1 = Convert.ToString(ds3.Tables[0].Rows[vd][2]);
                                    string DEPARTMENT = Convert.ToString(ds3.Tables[0].Rows[vd][6]);
                                    string BRANCH = Convert.ToString(ds3.Tables[0].Rows[vd][8]);
                                    string ED_IZM1 = Convert.ToString(ds3.Tables[0].Rows[vd][3]);
                                    string KOLVO1 = Convert.ToString(ds3.Tables[0].Rows[vd][4]);
                                    string PERIOD_ISPOLZOVAN1 = Convert.ToString(ds3.Tables[0].Rows[vd][5]);

                                    //Поиск в нормах того, что уже отпущено для вычисления кол-ва остатка по норме (то, что еще можно выдать)
                                    if (ID == NORMS_LIST_ID)
                                    {
                                        int ostatok = Convert.ToInt16(KOLVO) - Convert.ToInt16(KOLVO1); //Вычисление остатка от нормы, возможного к отпуску
                                        if (ostatok != 0)//Если остаток по позиции не 0, то выводим его в datagridview
                                        {
                                            //Вставка данных в таблицу Заявки
                                            conn.Open();
                                            SqlCommand scmd5 = conn.CreateCommand();
                                            scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                            scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                            scmd5.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                            scmd5.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                            scmd5.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                            scmd5.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                            scmd5.Parameters.AddWithValue("ed_iz", ED_IZM);
                                            scmd5.Parameters.AddWithValue("kolvo", ostatok);
                                            scmd5.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                            scmd5.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                            scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                            scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                            scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                            scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                            //Проверяем какой квартал выбран и исходя из этого кол-во невыданных СИЗов помещаем на первый месяц выбранного квартала 
                                            if (comboBox2.Text == "1 квартал")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", ostatok); //В квартальной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого квартала (на тот год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox2.Text == "2 квартал")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В квартальной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого квартала (на тот год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", ostatok);
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox2.Text == "3 квартал")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В квартальной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого квартала (на тот год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", ostatok);
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox2.Text == "4 квартал")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В квартальной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого квартала (на тот год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", ostatok);
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            SqlDataReader reader5;
                                            reader5 = scmd5.ExecuteReader();
                                            conn.Close();
                                        }
                                        aa++;
                                    }
                                }

                                if (aa == 0)
                                {
                                    //Вставка данных в таблицу Заявки
                                    conn.Open();
                                    SqlCommand scmd6 = conn.CreateCommand();
                                    scmd6.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                    scmd6.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                    scmd6.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                    scmd6.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                    scmd6.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                    scmd6.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                    scmd6.Parameters.AddWithValue("ed_iz", ED_IZM);
                                    scmd6.Parameters.AddWithValue("kolvo", KOLVO);
                                    scmd6.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                    scmd6.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                    scmd6.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                    scmd6.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                    scmd6.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                    scmd6.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                    //Проверяем какой квартал выбран и исходя из этого кол-во невыданных СИЗов помещаем на первый месяц выбранного квартала 
                                    if (comboBox2.Text == "1 квартал")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", KOLVO); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox2.Text == "2 квартал")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", KOLVO);
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox2.Text == "3 квартал")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", KOLVO);
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox2.Text == "4 квартал")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", KOLVO);
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    SqlDataReader reader6;
                                    reader6 = scmd6.ExecuteReader();
                                    conn.Close();
                                }

                            }
                            ////////////////////////////////////////////////////////////////

                            progressBar1.Value++;
                        }

                        SystemSounds.Beep.Play();
                        MessageBox.Show("Заявка сформирована. Просмотреть ее возможно в списке заявок.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                                                            
                    }
                }
                else //Если заходят остальные. Все кроме Петушкова и Скворцова и Шепилевой
                {
                    //Если выбрано по всем подразделениям
                    if (checkBox2.Checked == true)
                    {
                        conn.Open();
                        //Получаем список всех сотрудников. 
                        SqlCommand command = new SqlCommand("select * from sotrudniki where users_branch_id=@br", conn);
                        command.Parameters.AddWithValue("br", Form2.branch);                            
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        SqlCommandBuilder cb = new SqlCommandBuilder(da);
                        DataSet ds = new DataSet();
                        //conn.Close();
                        da.Fill(ds, "SOTRUDNIKI");
                        ///////////////////////////////////////////////////////////////

                        //Генерируем ключа GUID для связи заголовка заявки с его телом
                        SqlCommand command6 = new SqlCommand("select NEWID()", conn);
                        SqlDataAdapter da6 = new SqlDataAdapter(command6);//Переменная объявлена как глобальная
                        SqlCommandBuilder cb6 = new SqlCommandBuilder(da6);
                        DataSet ds6 = new DataSet();
                        //Заполнение DataGridView наименованиями полей 
                        da6.Fill(ds6, "NEWID");
                        /////////////////////////////

                        //Вставка данных в таблицу Заголовок Заявки
                        SqlCommand scmd4 = conn.CreateCommand();
                        scmd4.CommandText = "declare @ID uniqueidentifier " +
                                            "set @ID=NEWID() " +
                                            "INSERT into ZAYAVKA_MAIN (GUID, USER_ID, KIND_ZAYAV, PODR, PERIOD, DATETIME_CREATE) VALUES (@guid, @us_id, @kd_zayav, @podr, @per, GETDATE())";

                        scmd4.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                        scmd4.Parameters.AddWithValue("us_id", Form2.val);
                        scmd4.Parameters.AddWithValue("kd_zayav", "Квартальная");
                        scmd4.Parameters.AddWithValue("podr", comboBox5.Text);
                        scmd4.Parameters.AddWithValue("per", comboBox2.Text + " " + comboBox6.Text);
                        SqlDataReader reader4;
                        reader4 = scmd4.ExecuteReader();
                        /////////////////////////////////

                        progressBar1.Maximum = ds.Tables[0].Rows.Count;
                        progressBar1.Value = 0;

                        //Идем циклом по сформированному списку сотрудников
                        for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                        {
                            //Получаем то, что выдано сотруднику
                            SqlCommand command1 = new SqlCommand("select NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and DATEPART(year,VIDANO.DATE_SLED_VIDACHI)=@year and DATEPART(month,VIDANO.DATE_SLED_VIDACHI) between @bm and @em and SOTRUDNIKI.ID=@id_sotr group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            command1.Parameters.AddWithValue("id_sotr", ds.Tables[0].Rows[i][0].ToString());
                            command1.Parameters.AddWithValue("year", comboBox6.Text);
                            command1.Parameters.AddWithValue("bm", beg_m);
                            command1.Parameters.AddWithValue("em", end_m);
                            SqlDataAdapter da1 = new SqlDataAdapter(command1);
                            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                            DataSet ds1 = new DataSet();
                            conn.Close();
                            da1.Fill(ds1, "VIDANO");
                            ////////////////////////////////////////////////////////////////

                            //Проверяем выдачу, если сотруднику ничего не выдано, т.е. кол-во записей 0, то этот блок не выполняется, т.к. разбивать по месяцам нечего. Сразу переходим к следующему блоку, в котором проверяется, что не было выдано сотруднику.
                            if (ds1.Tables[0].Rows.Count != 0)
                            {
                                //Разбиваем по месяцам кол-во , которое попало в заявку
                                for (int g = 0; g <= ds1.Tables[0].Rows.Count - 1; g++)
                                {
                                    conn.Open();

                                    //Разбивка по месяцам потребности СИЗ по каждой позиции
                                    SqlCommand command5 = new SqlCommand("select VIDANO.NAIMEN_OVERALLS, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 1 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as JAN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 2 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as FEB, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 3 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as MAR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 4 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as APR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 5 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as MAY, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 6 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as JUN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 7 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as JUL, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 8 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as AUG, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 9 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as SEP, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 10 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as OKT, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 11 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as NOV, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 12 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as DEC " +
                                                                         "from VIDANO " +
                                                                         "where VIDANO.SOTRUDNIK_ID=@id_sot and VIDANO.NAIMEN_OVERALLS=@naim_cover " +
                                                                         "group by VIDANO.NAIMEN_OVERALLS", conn);

                                    command5.Parameters.AddWithValue("id_sot", ds.Tables[0].Rows[i][0].ToString());
                                    command5.Parameters.AddWithValue("naim_cover", ds1.Tables[0].Rows[g][0].ToString());
                                    SqlDataAdapter da5 = new SqlDataAdapter(command5);//Переменная объявлена как глобальная
                                    SqlCommandBuilder cb5 = new SqlCommandBuilder(da5);
                                    DataSet ds5 = new DataSet();
                                    da5.Fill(ds5, "VIDANO");

                                    //Если записи при выборе нет, то пропускаем этот шаг цикла
                                    if (ds5.Tables[0].Rows.Count > 0)
                                    {
                                        //Вставка данных в таблицу Заявки
                                        SqlCommand scmd5 = conn.CreateCommand();
                                        scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover1, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                        scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                        scmd5.Parameters.AddWithValue("naim_cover1", ds1.Tables[0].Rows[g][0]);
                                        scmd5.Parameters.AddWithValue("characterist", ds1.Tables[0].Rows[g][1]);
                                        scmd5.Parameters.AddWithValue("dep", ds1.Tables[0].Rows[g][4]);
                                        scmd5.Parameters.AddWithValue("br", ds1.Tables[0].Rows[g][6]);
                                        scmd5.Parameters.AddWithValue("ed_iz", ds1.Tables[0].Rows[g][5]);
                                        scmd5.Parameters.AddWithValue("kolvo", ds1.Tables[0].Rows[g][2]);
                                        scmd5.Parameters.AddWithValue("pr_ed_pl", ds1.Tables[0].Rows[g][3]);
                                        scmd5.Parameters.AddWithValue("jan", ds5.Tables[0].Rows[0][1]);
                                        scmd5.Parameters.AddWithValue("feb", ds5.Tables[0].Rows[0][2]);
                                        scmd5.Parameters.AddWithValue("mar", ds5.Tables[0].Rows[0][3]);
                                        scmd5.Parameters.AddWithValue("apr", ds5.Tables[0].Rows[0][4]);
                                        scmd5.Parameters.AddWithValue("may", ds5.Tables[0].Rows[0][5]);
                                        scmd5.Parameters.AddWithValue("jun", ds5.Tables[0].Rows[0][6]);
                                        scmd5.Parameters.AddWithValue("jul", ds5.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("aug", ds5.Tables[0].Rows[0][8]);
                                        scmd5.Parameters.AddWithValue("sep", ds5.Tables[0].Rows[0][9]);
                                        scmd5.Parameters.AddWithValue("okt", ds5.Tables[0].Rows[0][10]);
                                        scmd5.Parameters.AddWithValue("nov", ds5.Tables[0].Rows[0][11]);
                                        scmd5.Parameters.AddWithValue("dec", ds5.Tables[0].Rows[0][12]);
                                        scmd5.Parameters.AddWithValue("kontr", ds1.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                        scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                        scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                        scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                        SqlDataReader reader5;
                                        reader5 = scmd5.ExecuteReader();
                                        conn.Close();
                                    }
                                    else
                                    {
                                        conn.Close();
                                        continue;
                                    }

                                }
                            }


                            //Получаем то, что осталось к выдаче, но пока не выдано
                            //Определение нормы сотрудника
                            SqlCommand cmd = new SqlCommand("select norms_list.* from norms_personal,norms_list where norms_personal.norms_id=norms_list.norms_id and norms_personal.sotrudnik_id=@id_sotr1", conn);
                            cmd.Parameters.AddWithValue("id_sotr1", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da2 = new SqlDataAdapter(cmd);//Переменная объявлена как глобальная
                            SqlCommandBuilder cb2 = new SqlCommandBuilder(da2);
                            DataSet ds2 = new DataSet();
                            da2.Fill(ds2, "NORMS_LIST");
                            ///////////////////////////////

                            //Определение что уже выдано сотруднику
                            SqlCommand cmd1 = new SqlCommand("select VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID, NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and SOTRUDNIKI.ID=@id_sotr2 group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN, VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            cmd1.Parameters.AddWithValue("id_sotr2", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da3 = new SqlDataAdapter(cmd1);
                            SqlCommandBuilder cb3 = new SqlCommandBuilder(da3);
                            DataSet ds3 = new DataSet();
                            da3.Fill(ds3, "VIDANO");
                            ///////////////////////////////

                            ///////////////////////////////////////////////////
                            for (int nr = 0; nr <= ds2.Tables[0].Rows.Count - 1; nr++)
                            {
                                int aa = 0;
                                string ID = Convert.ToString(ds2.Tables[0].Rows[nr][0]);
                                string NORMS_ID = Convert.ToString(ds2.Tables[0].Rows[nr][1]);
                                string NAIMEN_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string CHARACTERISTIC_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string ED_IZM = Convert.ToString(ds2.Tables[0].Rows[nr][4]);
                                string KOLVO = Convert.ToString(ds2.Tables[0].Rows[nr][5]);

                                for (int vd = 0; vd <= ds3.Tables[0].Rows.Count - 1; vd++)
                                {
                                    string SOTRUDNIK_ID = Convert.ToString(ds3.Tables[0].Rows[vd][0]);
                                    NORMS_LIST_ID = Convert.ToString(ds3.Tables[0].Rows[vd][1]);
                                    string NAIMEN_OVERALLS1 = Convert.ToString(ds3.Tables[0].Rows[vd][2]);
                                    string DEPARTMENT = Convert.ToString(ds3.Tables[0].Rows[vd][6]);
                                    string BRANCH = Convert.ToString(ds3.Tables[0].Rows[vd][8]);
                                    string ED_IZM1 = Convert.ToString(ds3.Tables[0].Rows[vd][3]);
                                    string KOLVO1 = Convert.ToString(ds3.Tables[0].Rows[vd][4]);
                                    string PERIOD_ISPOLZOVAN1 = Convert.ToString(ds3.Tables[0].Rows[vd][5]);

                                    //Поиск в нормах того, что уже отпущено для вычисления кол-ва остатка по норме (то, что еще можно выдать)
                                    if (ID == NORMS_LIST_ID)
                                    {
                                        int ostatok = Convert.ToInt16(KOLVO) - Convert.ToInt16(KOLVO1); //Вычисление остатка от нормы, возможного к отпуску
                                        if (ostatok != 0)//Если остаток по позиции не 0, то выводим его в datagridview
                                        {
                                            //Вставка данных в таблицу Заявки
                                            conn.Open();
                                            SqlCommand scmd5 = conn.CreateCommand();
                                            scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                            scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                            scmd5.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                            scmd5.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                            scmd5.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                            scmd5.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                            scmd5.Parameters.AddWithValue("ed_iz", ED_IZM);
                                            scmd5.Parameters.AddWithValue("kolvo", ostatok);
                                            scmd5.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                            scmd5.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                            scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                            scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                            scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                            scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                            //Проверяем какой квартал выбран и исходя из этого кол-во невыданных СИЗов помещаем на первый месяц выбранного квартала 
                                            if (comboBox2.Text == "1 квартал")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", ostatok); //В квартальной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого квартала (на тот год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox2.Text == "2 квартал")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В квартальной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого квартала (на тот год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", ostatok);
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox2.Text == "3 квартал")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В квартальной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого квартала (на тот год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", ostatok);
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox2.Text == "4 квартал")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В квартальной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого квартала (на тот год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", ostatok);
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            SqlDataReader reader5;
                                            reader5 = scmd5.ExecuteReader();
                                            conn.Close();
                                        }
                                        aa++;
                                    }
                                }

                                if (aa == 0)
                                {
                                    //Вставка данных в таблицу Заявки
                                    conn.Open();
                                    SqlCommand scmd6 = conn.CreateCommand();
                                    scmd6.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                    scmd6.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                    scmd6.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                    scmd6.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                    scmd6.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                    scmd6.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                    scmd6.Parameters.AddWithValue("ed_iz", ED_IZM);
                                    scmd6.Parameters.AddWithValue("kolvo", KOLVO);
                                    scmd6.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                    scmd6.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                    scmd6.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                    scmd6.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                    scmd6.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                    scmd6.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                    //Проверяем какой квартал выбран и исходя из этого кол-во невыданных СИЗов помещаем на первый месяц выбранного квартала 
                                    if (comboBox2.Text == "1 квартал")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", KOLVO); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox2.Text == "2 квартал")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", KOLVO);
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox2.Text == "3 квартал")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", KOLVO);
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox2.Text == "4 квартал")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", KOLVO);
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    SqlDataReader reader6;
                                    reader6 = scmd6.ExecuteReader();
                                    conn.Close();
                                }

                            }
                            ////////////////////////////////////////////////////////////////

                            progressBar1.Value++;
                        }

                        SystemSounds.Beep.Play();
                        MessageBox.Show("Заявка сформирована. Просмотреть ее возможно в списке заявок.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                        
                    }
                    else
                    {
                        conn.Open();
                        //Получаем список всех сотрудников. 
                        SqlCommand command = new SqlCommand("select * from sotrudniki where users_branch_id=@br and department=@dp", conn);
                        command.Parameters.AddWithValue("br", Form2.branch);
                        command.Parameters.AddWithValue("dp", comboBox5.Text);                        
                        SqlDataAdapter da = new SqlDataAdapter(command);
                        SqlCommandBuilder cb = new SqlCommandBuilder(da);
                        DataSet ds = new DataSet();
                        //conn.Close();
                        da.Fill(ds, "SOTRUDNIKI");
                        ///////////////////////////////////////////////////////////////

                        //Генерируем ключа GUID для связи заголовка заявки с его телом
                        SqlCommand command6 = new SqlCommand("select NEWID()", conn);
                        SqlDataAdapter da6 = new SqlDataAdapter(command6);//Переменная объявлена как глобальная
                        SqlCommandBuilder cb6 = new SqlCommandBuilder(da6);
                        DataSet ds6 = new DataSet();
                        //Заполнение DataGridView наименованиями полей 
                        da6.Fill(ds6, "NEWID");
                        /////////////////////////////

                        //Вставка данных в таблицу Заголовок Заявки
                        SqlCommand scmd4 = conn.CreateCommand();
                        scmd4.CommandText = "declare @ID uniqueidentifier " +
                                            "set @ID=NEWID() " +
                                            "INSERT into ZAYAVKA_MAIN (GUID, USER_ID, KIND_ZAYAV, PODR, PERIOD, DATETIME_CREATE) VALUES (@guid, @us_id, @kd_zayav, @podr, @per, GETDATE())";

                        scmd4.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                        scmd4.Parameters.AddWithValue("us_id", Form2.val);
                        scmd4.Parameters.AddWithValue("kd_zayav", "Квартальная");
                        scmd4.Parameters.AddWithValue("podr", comboBox5.Text);
                        scmd4.Parameters.AddWithValue("per", comboBox2.Text + " " + comboBox6.Text);
                        SqlDataReader reader4;
                        reader4 = scmd4.ExecuteReader();
                        /////////////////////////////////

                        progressBar1.Maximum = ds.Tables[0].Rows.Count;
                        progressBar1.Value = 0;

                        //Идем циклом по сформированному списку сотрудников
                        for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                        {
                            //Получаем то, что выдано сотруднику
                            SqlCommand command1 = new SqlCommand("select NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and DATEPART(year,VIDANO.DATE_SLED_VIDACHI)=@year and DATEPART(month,VIDANO.DATE_SLED_VIDACHI) between @bm and @em and SOTRUDNIKI.ID=@id_sotr group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            command1.Parameters.AddWithValue("id_sotr", ds.Tables[0].Rows[i][0].ToString());
                            command1.Parameters.AddWithValue("year", comboBox6.Text);
                            command1.Parameters.AddWithValue("bm", beg_m);
                            command1.Parameters.AddWithValue("em", end_m);
                            SqlDataAdapter da1 = new SqlDataAdapter(command1);
                            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                            DataSet ds1 = new DataSet();
                            conn.Close();
                            da1.Fill(ds1, "VIDANO");
                            ////////////////////////////////////////////////////////////////

                            //Проверяем выдачу, если сотруднику ничего не выдано, т.е. кол-во записей 0, то этот блок не выполняется, т.к. разбивать по месяцам нечего. Сразу переходим к следующему блоку, в котором проверяется, что не было выдано сотруднику.
                            if (ds1.Tables[0].Rows.Count != 0)
                            {
                                //Разбиваем по месяцам кол-во , которое попало в заявку
                                for (int g = 0; g <= ds1.Tables[0].Rows.Count - 1; g++)
                                {
                                    conn.Open();

                                    //Разбивка по месяцам потребности СИЗ по каждой позиции
                                    SqlCommand command5 = new SqlCommand("select VIDANO.NAIMEN_OVERALLS, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 1 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as JAN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 2 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as FEB, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 3 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as MAR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 4 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as APR, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 5 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as MAY, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 6 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as JUN, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 7 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as JUL, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 8 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as AUG, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 9 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as SEP, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 10 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as OKT, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 11 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as NOV, " +
                                                                         "sum(case when MONTH (DATE_SLED_VIDACHI) = 12 and YEAR(DATE_SLED_VIDACHI)=" + comboBox6.Text + " and DATEPART(MONTH,VIDANO.DATE_SLED_VIDACHI) between " + beg_m + " and " + end_m + " then KOLVO else 0 end) as DEC " +
                                                                         "from VIDANO " +
                                                                         "where VIDANO.SOTRUDNIK_ID=@id_sot and VIDANO.NAIMEN_OVERALLS=@naim_cover " +
                                                                         "group by VIDANO.NAIMEN_OVERALLS", conn);

                                    command5.Parameters.AddWithValue("id_sot", ds.Tables[0].Rows[i][0].ToString());
                                    command5.Parameters.AddWithValue("naim_cover", ds1.Tables[0].Rows[g][0].ToString());
                                    SqlDataAdapter da5 = new SqlDataAdapter(command5);//Переменная объявлена как глобальная
                                    SqlCommandBuilder cb5 = new SqlCommandBuilder(da5);
                                    DataSet ds5 = new DataSet();
                                    da5.Fill(ds5, "VIDANO");

                                    //Если записи при выборе нет, то пропускаем этот шаг цикла
                                    if (ds5.Tables[0].Rows.Count > 0)
                                    {
                                        //Вставка данных в таблицу Заявки
                                        SqlCommand scmd5 = conn.CreateCommand();
                                        scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover1, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                        scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                        scmd5.Parameters.AddWithValue("naim_cover1", ds1.Tables[0].Rows[g][0]);
                                        scmd5.Parameters.AddWithValue("characterist", ds1.Tables[0].Rows[g][1]);
                                        scmd5.Parameters.AddWithValue("dep", ds1.Tables[0].Rows[g][4]);
                                        scmd5.Parameters.AddWithValue("br", ds1.Tables[0].Rows[g][6]);
                                        scmd5.Parameters.AddWithValue("ed_iz", ds1.Tables[0].Rows[g][5]);
                                        scmd5.Parameters.AddWithValue("kolvo", ds1.Tables[0].Rows[g][2]);
                                        scmd5.Parameters.AddWithValue("pr_ed_pl", ds1.Tables[0].Rows[g][3]);
                                        scmd5.Parameters.AddWithValue("jan", ds5.Tables[0].Rows[0][1]);
                                        scmd5.Parameters.AddWithValue("feb", ds5.Tables[0].Rows[0][2]);
                                        scmd5.Parameters.AddWithValue("mar", ds5.Tables[0].Rows[0][3]);
                                        scmd5.Parameters.AddWithValue("apr", ds5.Tables[0].Rows[0][4]);
                                        scmd5.Parameters.AddWithValue("may", ds5.Tables[0].Rows[0][5]);
                                        scmd5.Parameters.AddWithValue("jun", ds5.Tables[0].Rows[0][6]);
                                        scmd5.Parameters.AddWithValue("jul", ds5.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("aug", ds5.Tables[0].Rows[0][8]);
                                        scmd5.Parameters.AddWithValue("sep", ds5.Tables[0].Rows[0][9]);
                                        scmd5.Parameters.AddWithValue("okt", ds5.Tables[0].Rows[0][10]);
                                        scmd5.Parameters.AddWithValue("nov", ds5.Tables[0].Rows[0][11]);
                                        scmd5.Parameters.AddWithValue("dec", ds5.Tables[0].Rows[0][12]);
                                        scmd5.Parameters.AddWithValue("kontr", ds1.Tables[0].Rows[0][7]);
                                        scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                        scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                        scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                        scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                        SqlDataReader reader5;
                                        reader5 = scmd5.ExecuteReader();
                                        conn.Close();
                                    }
                                    else
                                    {
                                        conn.Close();
                                        continue;
                                    }

                                }
                            }


                            //Получаем то, что осталось к выдаче, но пока не выдано
                            //Определение нормы сотрудника
                            SqlCommand cmd = new SqlCommand("select norms_list.* from norms_personal,norms_list where norms_personal.norms_id=norms_list.norms_id and norms_personal.sotrudnik_id=@id_sotr1", conn);
                            cmd.Parameters.AddWithValue("id_sotr1", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da2 = new SqlDataAdapter(cmd);//Переменная объявлена как глобальная
                            SqlCommandBuilder cb2 = new SqlCommandBuilder(da2);
                            DataSet ds2 = new DataSet();
                            da2.Fill(ds2, "NORMS_LIST");
                            ///////////////////////////////

                            //Определение что уже выдано сотруднику
                            SqlCommand cmd1 = new SqlCommand("select VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID, NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SUM(VIDANO.KOLVO) as KOLVO, NORMS_LIST.PRICE_EDENIC_PLAN, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN from VIDANO, SOTRUDNIKI, NORMS_LIST left join KONTRAGENT on KONTRAGENT.ID=NORMS_LIST.KONTRAGENT_ID where SOTRUDNIKI.ID=VIDANO.SOTRUDNIK_ID and VIDANO.NORMS_LIST_ID=NORMS_LIST.ID and SOTRUDNIKI.ID=@id_sotr2 group by NORMS_LIST.NAIMEN_OVERALLS, NORMS_LIST.CHARACTERISTIC_OVERALLS, SOTRUDNIKI.DEPARTMENT, VIDANO.ED_IZM, SOTRUDNIKI.BRANCH, KONTRAGENT.NAIMENOVAN, NORMS_LIST.PRICE_EDENIC_PLAN, VIDANO.SOTRUDNIK_ID, VIDANO.NORMS_LIST_ID order by NORMS_LIST.NAIMEN_OVERALLS", conn);
                            cmd1.Parameters.AddWithValue("id_sotr2", ds.Tables[0].Rows[i][0].ToString());
                            SqlDataAdapter da3 = new SqlDataAdapter(cmd1);
                            SqlCommandBuilder cb3 = new SqlCommandBuilder(da3);
                            DataSet ds3 = new DataSet();
                            da3.Fill(ds3, "VIDANO");
                            ///////////////////////////////

                            ///////////////////////////////////////////////////
                            for (int nr = 0; nr <= ds2.Tables[0].Rows.Count - 1; nr++)
                            {
                                int aa = 0;
                                string ID = Convert.ToString(ds2.Tables[0].Rows[nr][0]);
                                string NORMS_ID = Convert.ToString(ds2.Tables[0].Rows[nr][1]);
                                string NAIMEN_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string CHARACTERISTIC_OVERALLS = Convert.ToString(ds2.Tables[0].Rows[nr][2]);
                                string ED_IZM = Convert.ToString(ds2.Tables[0].Rows[nr][4]);
                                string KOLVO = Convert.ToString(ds2.Tables[0].Rows[nr][5]);

                                for (int vd = 0; vd <= ds3.Tables[0].Rows.Count - 1; vd++)
                                {
                                    string SOTRUDNIK_ID = Convert.ToString(ds3.Tables[0].Rows[vd][0]);
                                    NORMS_LIST_ID = Convert.ToString(ds3.Tables[0].Rows[vd][1]);
                                    string NAIMEN_OVERALLS1 = Convert.ToString(ds3.Tables[0].Rows[vd][2]);
                                    string DEPARTMENT = Convert.ToString(ds3.Tables[0].Rows[vd][6]);
                                    string BRANCH = Convert.ToString(ds3.Tables[0].Rows[vd][8]);
                                    string ED_IZM1 = Convert.ToString(ds3.Tables[0].Rows[vd][3]);
                                    string KOLVO1 = Convert.ToString(ds3.Tables[0].Rows[vd][4]);
                                    string PERIOD_ISPOLZOVAN1 = Convert.ToString(ds3.Tables[0].Rows[vd][5]);

                                    //Поиск в нормах того, что уже отпущено для вычисления кол-ва остатка по норме (то, что еще можно выдать)
                                    if (ID == NORMS_LIST_ID)
                                    {
                                        int ostatok = Convert.ToInt16(KOLVO) - Convert.ToInt16(KOLVO1); //Вычисление остатка от нормы, возможного к отпуску
                                        if (ostatok != 0)//Если остаток по позиции не 0, то выводим его в datagridview
                                        {
                                            //Вставка данных в таблицу Заявки
                                            conn.Open();
                                            SqlCommand scmd5 = conn.CreateCommand();
                                            scmd5.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                            scmd5.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                            scmd5.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                            scmd5.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                            scmd5.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                            scmd5.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                            scmd5.Parameters.AddWithValue("ed_iz", ED_IZM);
                                            scmd5.Parameters.AddWithValue("kolvo", ostatok);
                                            scmd5.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                            scmd5.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                            scmd5.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                            scmd5.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                            scmd5.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                            scmd5.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                            //Проверяем какой квартал выбран и исходя из этого кол-во невыданных СИЗов помещаем на первый месяц выбранного квартала 
                                            if (comboBox2.Text == "1 квартал")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", ostatok); //В квартальной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого квартала (на тот год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox2.Text == "2 квартал")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В квартальной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого квартала (на тот год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", ostatok);
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox2.Text == "3 квартал")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В квартальной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого квартала (на тот год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", ostatok);
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", "0");
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            if (comboBox2.Text == "4 квартал")
                                            {
                                                scmd5.Parameters.AddWithValue("jan", "0"); //В квартальной заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого квартала (на тот год, на который формируется заявка)
                                                scmd5.Parameters.AddWithValue("feb", "0");
                                                scmd5.Parameters.AddWithValue("mar", "0");
                                                scmd5.Parameters.AddWithValue("apr", "0");
                                                scmd5.Parameters.AddWithValue("may", "0");
                                                scmd5.Parameters.AddWithValue("jun", "0");
                                                scmd5.Parameters.AddWithValue("jul", "0");
                                                scmd5.Parameters.AddWithValue("aug", "0");
                                                scmd5.Parameters.AddWithValue("sep", "0");
                                                scmd5.Parameters.AddWithValue("okt", ostatok);
                                                scmd5.Parameters.AddWithValue("nov", "0");
                                                scmd5.Parameters.AddWithValue("dec", "0");
                                            }
                                            SqlDataReader reader5;
                                            reader5 = scmd5.ExecuteReader();
                                            conn.Close();
                                        }
                                        aa++;
                                    }
                                }

                                if (aa == 0)
                                {
                                    //Вставка данных в таблицу Заявки
                                    conn.Open();
                                    SqlCommand scmd6 = conn.CreateCommand();
                                    scmd6.CommandText = "INSERT into ZAYAVKA_BODY (GUID, NAIMENOVAN_COVERALLS, CHARACTERISTIC_OVERALLS, DEPARTMENT, BRANCH, ED_IZM, KOL_VO, PRICE_EDINIC_PLAN, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OKT, NOV, DEC, DATETIME_CREATE, KONTRAGENT, FIO, SIZE_CLOTHES, SIZE_SHOES, HEIGHT) VALUES (@guid, @naim_cover2, @characterist, @dep, @br, @ed_iz, @kolvo, @pr_ed_pl, @jan, @feb, @mar, @apr, @may, @jun, @jul, @aug, @sep, @okt, @nov, @dec, GETDATE(), @kontr, @fio, @size_clo, @size_shoe, @height)";
                                    scmd6.Parameters.AddWithValue("guid", ds6.Tables[0].Rows[0][0]);
                                    scmd6.Parameters.AddWithValue("naim_cover2", NAIMEN_OVERALLS);
                                    scmd6.Parameters.AddWithValue("characterist", CHARACTERISTIC_OVERALLS);
                                    scmd6.Parameters.AddWithValue("dep", ds.Tables[0].Rows[i][4]);
                                    scmd6.Parameters.AddWithValue("br", ds.Tables[0].Rows[i][3]);
                                    scmd6.Parameters.AddWithValue("ed_iz", ED_IZM);
                                    scmd6.Parameters.AddWithValue("kolvo", KOLVO);
                                    scmd6.Parameters.AddWithValue("pr_ed_pl", ds2.Tables[0].Rows[nr][7]);
                                    scmd6.Parameters.AddWithValue("kontr", ds2.Tables[0].Rows[nr][11]);
                                    scmd6.Parameters.AddWithValue("fio", ds.Tables[0].Rows[i][1]);
                                    scmd6.Parameters.AddWithValue("size_clo", ds.Tables[0].Rows[i][11]);
                                    scmd6.Parameters.AddWithValue("size_shoe", ds.Tables[0].Rows[i][10]);
                                    scmd6.Parameters.AddWithValue("height", ds.Tables[0].Rows[i][12]);
                                    //Проверяем какой квартал выбран и исходя из этого кол-во невыданных СИЗов помещаем на первый месяц выбранного квартала 
                                    if (comboBox2.Text == "1 квартал")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", KOLVO); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox2.Text == "2 квартал")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", KOLVO);
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox2.Text == "3 квартал")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", KOLVO);
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", "0");
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    if (comboBox2.Text == "4 квартал")
                                    {
                                        scmd6.Parameters.AddWithValue("jan", "0"); //В годовой заявке кол-во СИЗ, которое не было выдано, попадает в заявку в первый месяц формируемого года (на тот год, на который формируется заявка)
                                        scmd6.Parameters.AddWithValue("feb", "0");
                                        scmd6.Parameters.AddWithValue("mar", "0");
                                        scmd6.Parameters.AddWithValue("apr", "0");
                                        scmd6.Parameters.AddWithValue("may", "0");
                                        scmd6.Parameters.AddWithValue("jun", "0");
                                        scmd6.Parameters.AddWithValue("jul", "0");
                                        scmd6.Parameters.AddWithValue("aug", "0");
                                        scmd6.Parameters.AddWithValue("sep", "0");
                                        scmd6.Parameters.AddWithValue("okt", KOLVO);
                                        scmd6.Parameters.AddWithValue("nov", "0");
                                        scmd6.Parameters.AddWithValue("dec", "0");
                                    }
                                    SqlDataReader reader6;
                                    reader6 = scmd6.ExecuteReader();
                                    conn.Close();
                                }

                            }
                            ////////////////////////////////////////////////////////////////

                            progressBar1.Value++;
                        }

                        SystemSounds.Beep.Play();
                        MessageBox.Show("Заявка сформирована. Просмотреть ее возможно в списке заявок.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                                            
                    }
                }
            }

        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                comboBox7.Enabled = false;
            }
            else
            {
                comboBox7.Enabled = true;
            }
        }

        
    }
}
