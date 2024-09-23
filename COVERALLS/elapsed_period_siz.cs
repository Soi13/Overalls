using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace COVERALLS
{
    public partial class elapsed_period_siz : Form
    {
        public elapsed_period_siz()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["FIO"].HeaderText = "ФИО";
            dataGridView1.Columns["FIO"].Width = 150;
            dataGridView1.Columns["DEPARTMENT"].HeaderText = "Отдел/участок/цех";
            dataGridView1.Columns["DEPARTMENT"].Width = 200;
            dataGridView1.Columns["NAIMEN_OVERALLS"].HeaderText = "Отдел/участок/цех";
            dataGridView1.Columns["NAIMEN_OVERALLS"].Width = 300;

        }


        private void elapsed_period_siz_Load(object sender, EventArgs e)
        {
            fill_gridview();
        }
    }
}
