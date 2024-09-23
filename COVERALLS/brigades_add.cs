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
    public partial class brigades_add : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
       
        public brigades_add()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Бригадир уже выбран!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            add_brigadir add_brigadir = new add_brigadir();
            add_brigadir.Owner = this;
            add_brigadir.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не выбран бригадир!", "Внимание", MessageBoxButtons.OK);
                return;
            }

            add_consist_brigade add_consist_brigade = new add_consist_brigade();
            add_consist_brigade.Owner = this;
            add_consist_brigade.ShowDialog();
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count > 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Удалить бригадира невозможно, т.к. уже выбраны члены бригады!", "Внимание", MessageBoxButtons.OK);
                return;
            }

            if (MessageBox.Show("Вы уверены, что хотите удалить текущую позицию?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                dataGridView1.Rows.Remove(dataGridView1.CurrentRow);
            }
        }

        private void удалитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите удалить текущую позицию?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                dataGridView2.Rows.Remove(dataGridView2.CurrentRow);
            }
        }
    }
}
