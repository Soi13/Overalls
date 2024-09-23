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
    public partial class cancel_vidano_add_positions : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");

        SqlDataAdapter da;
        DataSet ds;
        public cancel_vidano CV;

        public cancel_vidano_add_positions(cancel_vidano cancel_vidano)
        {
            InitializeComponent();
            CV = cancel_vidano;
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["SOTRUDNIK_ID"].HeaderText = "SOTRUDNIK_ID";
            dataGridView1.Columns["SOTRUDNIK_ID"].Width = 40;
            dataGridView1.Columns["SOTRUDNIK_ID"].Visible = false;
            dataGridView1.Columns["NORMS_LIST_ID"].HeaderText = "NORMS_LIST_ID";
            dataGridView1.Columns["NORMS_LIST_ID"].Width = 200;
            dataGridView1.Columns["NORMS_LIST_ID"].Visible = false;
            dataGridView1.Columns["NAIMEN_OVERALLS"].HeaderText = "Наименование";
            dataGridView1.Columns["NAIMEN_OVERALLS"].Width = 200;
            dataGridView1.Columns["ED_IZM"].HeaderText = "Ед. изм.";
            dataGridView1.Columns["ED_IZM"].Width = 50;
            dataGridView1.Columns["KOLVO"].HeaderText = "Кол-во";
            dataGridView1.Columns["KOLVO"].Width = 50;
            dataGridView1.Columns["KOLVO"].ToolTipText = "Это кол-во для отмены выдачи";
            dataGridView1.Columns["PERIOD_ISPOLZOVAN"].HeaderText = "Период испол-я";
            dataGridView1.Columns["PERIOD_ISPOLZOVAN"].Width = 100;
            dataGridView1.Columns["DATE_VIDACHI"].HeaderText = "Дата послед. выдачи";
            dataGridView1.Columns["DATE_VIDACHI"].Width = 100;
            dataGridView1.Columns["DATE_SLED_VIDACHI"].HeaderText = "Дата след. выдачи";
            dataGridView1.Columns["DATE_SLED_VIDACHI"].Width = 100;
            dataGridView1.Columns["SIZE"].HeaderText = "Размер";
            dataGridView1.Columns["SIZE"].Width = 70;
            dataGridView1.Columns["USER_ID"].HeaderText = "USER_ID";
            dataGridView1.Columns["USER_ID"].Width = 40;
            dataGridView1.Columns["USER_ID"].Visible = false;
            dataGridView1.Columns["DATETIME_CREATE"].HeaderText = "DATETIME_CREATE";
            dataGridView1.Columns["DATETIME_CREATE"].Width = 40;
            dataGridView1.Columns["DATETIME_CREATE"].Visible = false;
            dataGridView1.Columns["IAIS_NAIMEN_OVERALLS"].HeaderText = "Наимен. из ИАИС";
            dataGridView1.Columns["IAIS_NAIMEN_OVERALLS"].Width = 150;
            dataGridView1.Columns["ID_IAIS"].HeaderText = "Наимен. из ИАИС";
            dataGridView1.Columns["ID_IAIS"].Width = 50;
            dataGridView1.Columns["ID_IAIS"].Visible = false;
            
        }

        private void cancel_vidano_add_positions_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;

            SqlCommand command = new SqlCommand("select * from vidano where SOTRUDNIK_ID=" + Form1.sotrudn_id, conn);
            da = new SqlDataAdapter(command);
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            da.Fill(ds, "VIDANO");
            dataGridView1.DataSource = ds.Tables[0];

            fill_gridview();
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            string id_position_coverall = Convert.ToString(dataGridView1.CurrentRow.Cells[0].Value);
            string id_sotrudnik = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);
            string naimen = Convert.ToString(dataGridView1.CurrentRow.Cells[3].Value); 
            string edizm = Convert.ToString(dataGridView1.CurrentRow.Cells[4].Value);
            string kol_vo = Convert.ToString(dataGridView1.CurrentRow.Cells[5].Value);
            string per_isp = Convert.ToString(dataGridView1.CurrentRow.Cells[6].Value);
            string date_posled_vidachi = Convert.ToString(dataGridView1.CurrentRow.Cells[7].Value);
            string date_sled_vidachi = Convert.ToString(dataGridView1.CurrentRow.Cells[8].Value);
            string size = Convert.ToString(dataGridView1.CurrentRow.Cells[9].Value);
            string iais_naimen = Convert.ToString(dataGridView1.CurrentRow.Cells[12].Value);

            if (CV.dataGridView2.Rows.Count > 0)
            {
                for (int s = 0; s <= CV.dataGridView2.Rows.Count - 1; s++)
                {
                    string iidd = Convert.ToString(CV.dataGridView2.Rows[s].Cells[0].Value);
                    string naim = Convert.ToString(CV.dataGridView2.Rows[s].Cells[2].Value);
                    if ((iidd == id_position_coverall) || (naim == naimen))
                    {
                        SystemSounds.Beep.Play();
                        MessageBox.Show("Данная позиция уже присутствует в списке выбранных!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }
            }
            CV.dataGridView2.Rows.Add(id_position_coverall, id_sotrudnik, naimen, edizm, kol_vo, per_isp, date_posled_vidachi, date_sled_vidachi, "", iais_naimen);
            this.Close();
        }

        
    }
}
