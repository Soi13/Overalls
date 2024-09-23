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
    public partial class alert_norms_changed : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
      
        public alert_norms_changed()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {           
            dataGridView1.Columns["PROFESSION"].HeaderText = "Норма-Профессия/должность";
            dataGridView1.Columns["PROFESSION"].Width = 300;
            dataGridView1.Columns["NAIMEN_OVERALLS"].HeaderText = "Наименование";
            dataGridView1.Columns["NAIMEN_OVERALLS"].Width = 300;            
        }
        //////////////////////////////////////////////////////

        private void alert_norms_changed_Load(object sender, EventArgs e)
        {
            SystemSounds.Beep.Play();

            SqlCommand command = new SqlCommand("select NORMS.PROFESSION, NORMS_LIST.NAIMEN_OVERALLS from NORMS, NORMS_LIST where NORMS.ID=NORMS_LIST.NORMS_ID and NORMS_LIST.CHANGED=1", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "NORMS_LIST");
            dataGridView1.DataSource = ds.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

            fill_gridview();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Установка статуса Изменено в 0
            SqlCommand scmd4 = conn.CreateCommand();
            scmd4.CommandText = "update NORMS_LIST set CHANGED=0 where CHANGED=1";
            try
            {
                conn.Open();
            }
            catch
            {
                MessageBox.Show("Ошибка соединения с базой данных");
            }
            SqlDataReader reader4;
            reader4 = scmd4.ExecuteReader();
            conn.Close();
            ////////////////////////

            this.Close();
        }
    }
}
