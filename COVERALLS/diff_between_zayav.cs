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
    public partial class diff_between_zayav : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");

        public diff_between_zayav()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["NAIMENOVAN_COVERALLS"].HeaderText = "Наименование";
            dataGridView1.Columns["NAIMENOVAN_COVERALLS"].Width = 350;
            dataGridView1.Columns["CHARACTERISTIC_OVERALLS"].HeaderText = "Характеристики";
            dataGridView1.Columns["CHARACTERISTIC_OVERALLS"].Width = 200;
            dataGridView1.Columns["BRANCH"].HeaderText = "Филиал";
            dataGridView1.Columns["BRANCH"].Width = 150;
            dataGridView1.Columns["ED_IZM"].HeaderText = "Ед. изм";
            dataGridView1.Columns["ED_IZM"].Width = 40;
            dataGridView1.Columns["kolvo"].HeaderText = "Кол-во";
            dataGridView1.Columns["kolvo"].Width = 60;
            
        }
        //////////////////////////////////////////////////////

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void diff_between_zayav_Load(object sender, EventArgs e)
        {
            view_zayavki view_zayavki = (view_zayavki)this.Owner;

            //Строим запрос для сравнения 2-х таблиц. Для Этого создаем временные таблицы на осонованни GUIDов и их сравниваем. После сравнения удаляем.
            SqlCommand command1 = new SqlCommand("select * into #zayav1 from ZAYAVKA_BODY where GUID=@g1 " +
                                                 "select * into #zayav2 from ZAYAVKA_BODY where GUID=@g2 " +
                                                 "select #zayav1.NAIMENOVAN_COVERALLS, #zayav1.CHARACTERISTIC_OVERALLS, #zayav1.BRANCH, #zayav1.ED_IZM,  case when (#zayav2.kol_vo-#zayav1.kol_vo)=0 then #zayav1.kol_vo else (#zayav2.kol_vo-#zayav1.kol_vo) end as kolvo from #zayav1, #zayav2 where #zayav1.NAIMENOVAN_COVERALLS=#zayav2.NAIMENOVAN_COVERALLS " +
                                                 "union " +
                                                 "SELECT #zayav1.NAIMENOVAN_COVERALLS, #zayav1.CHARACTERISTIC_OVERALLS, #zayav1.BRANCH, #zayav1.ED_IZM, #zayav1.kol_vo " +
                                                 "FROM #zayav1 " +
                                                 "WHERE NOT EXISTS(" +
                                                 "SELECT #zayav2.NAIMENOVAN_COVERALLS, #zayav2.CHARACTERISTIC_OVERALLS, #zayav2.BRANCH, #zayav2.ED_IZM, #zayav2.kol_vo " +
                                                 "FROM #zayav2 " +
                                                 "WHERE #zayav2.NAIMENOVAN_COVERALLS=#zayav1.NAIMENOVAN_COVERALLS) " +
                                                 "union " +
                                                 "SELECT #zayav2.NAIMENOVAN_COVERALLS, #zayav2.CHARACTERISTIC_OVERALLS, #zayav2.BRANCH, #zayav2.ED_IZM, #zayav2.kol_vo " +
                                                 "FROM #zayav2 " +
                                                 "WHERE NOT EXISTS(" +
                                                 "SELECT #zayav1.NAIMENOVAN_COVERALLS, #zayav1.CHARACTERISTIC_OVERALLS, #zayav1.BRANCH, #zayav1.ED_IZM, #zayav1.kol_vo " +
                                                 "FROM #zayav1 " +
                                                 "WHERE #zayav1.NAIMENOVAN_COVERALLS=#zayav2.NAIMENOVAN_COVERALLS) " +
                                                 "drop table #zayav1 " +
                                                 "drop table #zayav2 ", conn);
            SqlDataAdapter da1 = new SqlDataAdapter(command1);//Переменная объявлена как глобальная
            command1.Parameters.AddWithValue("g1", view_zayavki.arr_guid[0]);
            command1.Parameters.AddWithValue("g2", view_zayavki.arr_guid[1]);
            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
            DataSet ds1 = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da1.Fill(ds1, "zayav1");
            dataGridView1.DataSource = ds1.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds1.Tables[0].Rows.Count);

            fill_gridview();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            help_diff_zayav help_diff_zayav = new help_diff_zayav();
            help_diff_zayav.ShowDialog();
        }
    }
}
