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
    public partial class Form9 : Form
    {
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");

        SqlDataAdapter da;
        DataSet ds;
        public static DataSet ds3;
        public Form7 f7;
        public static string date_sled_vidachi;
        public static string query;
        public static string iais_naimen;
        public static string id_iais;
    

        public Form9(Form7 form7)
        {
            InitializeComponent();
            f7 = form7;
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
            dataGridView1.Columns["ED_IZM"].ReadOnly = true;
            dataGridView1.Columns["KOLVO"].HeaderText = "Кол-во";
            dataGridView1.Columns["KOLVO"].Width = 50;
            dataGridView1.Columns["KOLVO"].ToolTipText = "Это кол-во по норме";
            dataGridView1.Columns["PERIOD_ISPOLZOVAN"].HeaderText = "Период испол-я";
            dataGridView1.Columns["PERIOD_ISPOLZOVAN"].Width = 100;
            dataGridView1.Columns["PRICE_EDENIC_PLAN"].HeaderText = "Стоим. за ед. план.";
            dataGridView1.Columns["PRICE_EDENIC_PLAN"].Width = 100;
            dataGridView1.Columns["PRICE_EDENIC_PLAN"].Visible = false;
            dataGridView1.Columns["PRICE_KOMPLEKT_PLAN"].HeaderText = "Стоим. за комплект план.";
            dataGridView1.Columns["PRICE_KOMPLEKT_PLAN"].Width = 100;
            dataGridView1.Columns["PRICE_KOMPLEKT_PLAN"].Visible = false;
            dataGridView1.Columns["IAIS_NAIMEN_OVERALLS"].HeaderText = "Наимен. из ИАИС";
            dataGridView1.Columns["IAIS_NAIMEN_OVERALLS"].Width = 150;
            dataGridView1.Columns["IAIS_NAIMEN_OVERALLS"].Visible=false;            
            dataGridView1.Columns["CODE_GROUP_IAIS"].HeaderText = "CODE_GROUP_IAIS";
            dataGridView1.Columns["CODE_GROUP_IAIS"].Width = 40;
            dataGridView1.Columns["CODE_GROUP_IAIS"].Visible = false;
            dataGridView1.Columns["KONTRAGENT_ID"].HeaderText = "KONTRAGENT_ID";
            dataGridView1.Columns["KONTRAGENT_ID"].Width = 40;
            dataGridView1.Columns["KONTRAGENT_ID"].Visible = false;            
            
        }


        private void Form9_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;
            
            SqlCommand command = new SqlCommand("select norms_list.* from norms_personal,norms_list where norms_personal.norms_id=norms_list.norms_id and norms_personal.sotrudnik_id=" + Form1.sotrudn_id, conn);
            da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "NORMS_LIST");
            dataGridView1.DataSource = ds.Tables[0];

            fill_gridview();
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            id_iais = "";
            iais_naimen = "";
            string id_position_coverall = Convert.ToString(dataGridView1.CurrentRow.Cells[0].Value);
            string naimen = Convert.ToString(dataGridView1.CurrentRow.Cells[2].Value); //глобальная переменная
            string charact = Convert.ToString(dataGridView1.CurrentRow.Cells[3].Value);
            string edizm = Convert.ToString(dataGridView1.CurrentRow.Cells[4].Value);
            string kol_vo = "";
            string per_isp = Convert.ToString(dataGridView1.CurrentRow.Cells[6].Value);
            string date_posled_vidachi = Convert.ToString(DateTime.Now);

            //Определение срока использования носки, если он меньше месяца, например 0.7, то используется функция AddDays, если более месяца (целое число), то AddMonth
            string srok =dataGridView1.CurrentRow.Cells[6].Value.ToString();
            string decimal_sep = System.Globalization.NumberFormatInfo.CurrentInfo.CurrencyDecimalSeparator;
            srok = srok.Replace(".", decimal_sep);
            srok = srok.Replace(",", decimal_sep);
            double srok_p = Convert.ToDouble(srok);
                
            if (srok_p < 1)
            {
                double rs = srok_p * Convert.ToDouble(DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month));
                date_sled_vidachi = Convert.ToString(DateTime.Now.Date.AddDays(rs));//Вычисление даты следующей выдачи спецодежы            
            }
            else
            {
                date_sled_vidachi = Convert.ToString(DateTime.Now.Date.AddMonths(Convert.ToInt16(dataGridView1.CurrentRow.Cells[6].Value)));//Вычисление даты следующей выдачи спецодежы
            }
            iais_naimen = Convert.ToString(dataGridView1.CurrentRow.Cells[9].Value);
            string code_iais = Convert.ToString(dataGridView1.CurrentRow.Cells[10].Value);
            ////////////////////////////////////////
         /* 
            //Поиск строки с кодами IAIS в справочнике соответсвия
            SqlCommand command = new SqlCommand("select * from iais_codes where code_group_iais=" + "'" + code_iais + "'", conn);
            da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            ds = new DataSet();
            conn.Close();
            da.Fill(ds, "IAIS_CODES");
            if (ds.Tables[0].Rows.Count > 0) //Если строка находится в справочнике соответствия, то выполняются условия во вложении ниже
            {
                string str = Convert.ToString(ds.Tables[0].Rows[0][2]);
                String[] words = str.Split(new char[] {','}, StringSplitOptions.RemoveEmptyEntries);
                
                //Строительство запроса исходя из кол-ва кодов в поле/результате запроса
               //Проверка длинны полякодов ИАИС. Если в поле один код, то запрос строится элементарно, если 2 и более, то применяем конструктор запроса
                if (words.Length==1)
                {
                    SqlCommand command1 = new SqlCommand("select id, code_iais, naimen_iais from iais where code_iais='" + words[0] + "'", conn);
                    da = new SqlDataAdapter(command1);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb1 = new SqlCommandBuilder(da);
                    ds3 = new DataSet();
                    conn.Close();
                    da.Fill(ds3, "IAIS");
                    iais iais = new iais();
                    iais.ShowDialog(); 
                }
                else if (words.Length > 1)
                {
                    string part = "";
                    foreach (string a in words)
                    {
                        part = part + " CODE_IAIS='" + a+"' or ";
                    }
                    string res = "select id, code_iais, naimen_iais from iais where " + part;
                    string rr = res.Remove(res.Length - 4, 4); //Удаляем с конца строки последнего кода значение or и пробелы
                    
                    SqlCommand command2 = new SqlCommand(rr, conn); //Подставляем сюда переменную с текстом построенного запроса
                    da = new SqlDataAdapter(command2);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb2 = new SqlCommandBuilder(da);
                    ds3 = new DataSet();
                    conn.Close();
                    da.Fill(ds3, "IAIS");
                    iais iais = new iais();
                    iais.ShowDialog();                   
                }
            }*/
              
            if (f7.dataGridView2.Rows.Count > 0)
            {
                for (int s = 0; s <= f7.dataGridView2.Rows.Count-1; s++)
                {
                    string iidd = Convert.ToString(f7.dataGridView2.Rows[s].Cells[0].Value);
                    string naim = Convert.ToString(f7.dataGridView2.Rows[s].Cells[1].Value);
                    if ((iidd == id_position_coverall) || (naim == naimen))
                    {
                        SystemSounds.Beep.Play();
                        MessageBox.Show("Данная позиция уже присутствует в списке выбранных!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }
            }
            //f7.dataGridView2.Rows.Add(id_position_coverall, naimen, charact, edizm, kol_vo, per_isp, date_posled_vidachi, date_sled_vidachi, "", iais_naimen, id_iais);
            f7.dataGridView2.Rows.Add(id_position_coverall, naimen, charact, edizm, kol_vo, per_isp, date_posled_vidachi, date_sled_vidachi, "");            
            this.Close();
        }

    

        
    }
}
