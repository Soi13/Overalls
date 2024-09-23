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
    public partial class brigades : Form
    {
        public brigades()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            brigades_add brigades_add = new brigades_add();
            brigades_add.Owner = this;
            brigades_add.ShowDialog();
        }
    }
}
