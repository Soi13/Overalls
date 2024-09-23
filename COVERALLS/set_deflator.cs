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
    public partial class set_deflator : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");

        public set_deflator()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не введено значение в поле!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            
            if (MessageBox.Show("Внимание! При установке дефлятора, все цены в нормах будут пересчитаны в соответствии с заданным процентом. Вы уверены, что необходимо выполнить установку дефлятора?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                SqlCommand scmd4 = conn.CreateCommand();
                scmd4.CommandText = "update NORMS_LIST set PRICE_EDENIC_PLAN=PRICE_EDENIC_PLAN*(100-@proc)/100";
                scmd4.Parameters.AddWithValue("proc", textBox1.Text);
                try
                {
                    conn.Open();
                }
                catch { }
                SqlDataReader reader4;
                reader4 = scmd4.ExecuteReader();
                conn.Close();

                SystemSounds.Beep.Play();
                MessageBox.Show("Дефлятор установлен.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
            }
        }
    }
}
