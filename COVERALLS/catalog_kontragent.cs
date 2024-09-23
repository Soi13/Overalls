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
    public partial class catalog_kontragent : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");

        public static int idd;        

        public catalog_kontragent()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;            
            dataGridView1.Columns["INN_KPP"].HeaderText = "ИНН/КПП";
            dataGridView1.Columns["INN_KPP"].Width = 80;
            dataGridView1.Columns["NAIMENOVAN"].HeaderText = "Наименование";
            dataGridView1.Columns["NAIMENOVAN"].Width = 200;
            dataGridView1.Columns["UR_ADDRESS"].HeaderText = "Юр. адрес";
            dataGridView1.Columns["UR_ADDRESS"].Width = 300;
            dataGridView1.Columns["POST_ADDRESS"].HeaderText = "Почтовый адрес";
            dataGridView1.Columns["POST_ADDRESS"].Width = 300;
            dataGridView1.Columns["CONTACTS"].HeaderText = "Контакты";
            dataGridView1.Columns["CONTACTS"].Width = 200;
            dataGridView1.Columns["EMAIL"].HeaderText = "E-Mail";
            dataGridView1.Columns["EMAIL"].Width = 100;
        }

        //Обновление данных в гриде после ввода новой записи
        public void refill()
        {
            SqlCommand command = new SqlCommand("select * from KONTRAGENT", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "KONTRAGENT");
            dataGridView1.DataSource = ds.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

            fill_gridview();
        }

        private void catalog_kontragent_Load(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand("select * from KONTRAGENT", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "KONTRAGENT");
            dataGridView1.DataSource = ds.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

            fill_gridview();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            catalog_kontragent_add catalog_kontragent_add = new catalog_kontragent_add();
            catalog_kontragent_add.Owner = this;
            catalog_kontragent_add.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            idd = Convert.ToInt16(dataGridView1.CurrentRow.Cells[0].Value);

            catalog_kontragent_edit catalog_kontragent_edit = new catalog_kontragent_edit();
            catalog_kontragent_edit.Owner = this;
            catalog_kontragent_edit.ShowDialog();
        }

        private void привязатьНоменклатуруToolStripMenuItem_Click(object sender, EventArgs e)
        {
            idd = Convert.ToInt16(dataGridView1.CurrentRow.Cells[0].Value);

            choose_nomenkl choose_nomenkl = new choose_nomenkl();
            choose_nomenkl.Owner = this;
            choose_nomenkl.ShowDialog();
        }

        private void просмотрПривязаннойНоменклатурыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            idd = Convert.ToInt16(dataGridView1.CurrentRow.Cells[0].Value);

            SqlCommand command = new SqlCommand("select * from NORMS_LIST where KONTRAGENT_ID=@id order by NAIMEN_OVERALLS", conn);
            SqlDataAdapter da1 = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da1);
            DataSet ds = new DataSet();
            command.Parameters.AddWithValue("id", idd);
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da1.Fill(ds, "NORMS_LIST");

            if (ds.Tables[0].Rows.Count == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("К данному контрагенту номенклатура не привязана!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            show_tie_nomenkl show_tie_nomenkl = new show_tie_nomenkl();
            show_tie_nomenkl.Owner = this;
            show_tie_nomenkl.ShowDialog();
        }
    }
}
