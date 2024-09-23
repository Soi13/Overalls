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
    public partial class movement_select_profession : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
      
        public movement_select_profession()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["PROFESSION"].HeaderText = "Профессия/должность";
            dataGridView1.Columns["PROFESSION"].Width = 400;            
        }
        //////////////////////////////////////////////////////

        private void movement_select_profession_Load(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand("select distinct PROFESSION from SOTRUDNIKI order by PROFESSION", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "SOTRUDNIKI");
            dataGridView1.DataSource = ds.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

            fill_gridview();
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            movement_param movement_param = (movement_param)this.Owner;

            if (movement_param.dataGridView2.Rows.Count > 0)
            {
                if (movement_param.dataGridView2.CurrentRow.Cells[0].Value.ToString().Length != 0)
                {
                    SystemSounds.Beep.Play();
                    MessageBox.Show("Профессия уже выбрана!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }

            string profession = dataGridView1.CurrentRow.Cells[0].Value.ToString();

            if (movement_param.dataGridView2.Rows.Count == 0)
            {
                movement_param.dataGridView2.Rows.Add("", "");
                movement_param.dataGridView2.Rows[0].Cells[0].Value = profession;
                this.Close();
            }
            else
            {
                movement_param.dataGridView2.Rows[0].Cells[0].Value = profession;
                this.Close();
            }
        }
    }
}
