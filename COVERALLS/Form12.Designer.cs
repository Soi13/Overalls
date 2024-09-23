namespace COVERALLS
{
    partial class Form12
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SOTRUDNIK_ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NORMS_LIST_ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NAIMEN_OVERALLS = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ED_IZM = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.KOLVO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PERIOD_ISPOLZOVAN = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PRICE_EDENIC_PLAN = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PRICE_KOMPLEKT_PLAN = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SIZE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.USER_ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NAIMEN_IAIS = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ID_IAIS = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.label4 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Times New Roman", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.ForeColor = System.Drawing.Color.Blue;
            this.label1.Location = new System.Drawing.Point(505, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(193, 26);
            this.label1.TabIndex = 1;
            this.label1.Text = "Акт на списание";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.ForeColor = System.Drawing.Color.Red;
            this.label2.Location = new System.Drawing.Point(12, 44);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(49, 19);
            this.label2.TabIndex = 3;
            this.label2.Text = "label2";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.Location = new System.Drawing.Point(12, 419);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(141, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "Добавить позицию";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button2.Location = new System.Drawing.Point(171, 419);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(141, 23);
            this.button2.TabIndex = 5;
            this.button2.Text = "Удалить позицию";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button3.Location = new System.Drawing.Point(884, 419);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(129, 44);
            this.button3.TabIndex = 6;
            this.button3.Text = "Сформировать акт";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button4.Location = new System.Drawing.Point(1032, 419);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(129, 44);
            this.button4.TabIndex = 7;
            this.button4.Text = "Отмена формирования";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToAddRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView2.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ID,
            this.SOTRUDNIK_ID,
            this.NORMS_LIST_ID,
            this.NAIMEN_OVERALLS,
            this.ED_IZM,
            this.KOLVO,
            this.PERIOD_ISPOLZOVAN,
            this.PRICE_EDENIC_PLAN,
            this.PRICE_KOMPLEKT_PLAN,
            this.SIZE,
            this.USER_ID,
            this.NAIMEN_IAIS,
            this.ID_IAIS});
            this.dataGridView2.Location = new System.Drawing.Point(11, 66);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new System.Drawing.Size(1151, 345);
            this.dataGridView2.TabIndex = 9;
            this.dataGridView2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dataGridView2_KeyPress);
            // 
            // ID
            // 
            this.ID.HeaderText = "ID";
            this.ID.Name = "ID";
            this.ID.ReadOnly = true;
            this.ID.Visible = false;
            this.ID.Width = 50;
            // 
            // SOTRUDNIK_ID
            // 
            this.SOTRUDNIK_ID.HeaderText = "SOTRUDNIK_ID";
            this.SOTRUDNIK_ID.Name = "SOTRUDNIK_ID";
            this.SOTRUDNIK_ID.ReadOnly = true;
            this.SOTRUDNIK_ID.Visible = false;
            this.SOTRUDNIK_ID.Width = 50;
            // 
            // NORMS_LIST_ID
            // 
            this.NORMS_LIST_ID.HeaderText = "NORMS_LIST_ID";
            this.NORMS_LIST_ID.Name = "NORMS_LIST_ID";
            this.NORMS_LIST_ID.ReadOnly = true;
            this.NORMS_LIST_ID.Visible = false;
            // 
            // NAIMEN_OVERALLS
            // 
            this.NAIMEN_OVERALLS.HeaderText = "Наименование";
            this.NAIMEN_OVERALLS.Name = "NAIMEN_OVERALLS";
            this.NAIMEN_OVERALLS.ReadOnly = true;
            this.NAIMEN_OVERALLS.Width = 250;
            // 
            // ED_IZM
            // 
            this.ED_IZM.HeaderText = "Ед. изм.";
            this.ED_IZM.Name = "ED_IZM";
            this.ED_IZM.ReadOnly = true;
            this.ED_IZM.Width = 50;
            // 
            // KOLVO
            // 
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.KOLVO.DefaultCellStyle = dataGridViewCellStyle2;
            this.KOLVO.HeaderText = "Кол-во";
            this.KOLVO.Name = "KOLVO";
            // 
            // PERIOD_ISPOLZOVAN
            // 
            this.PERIOD_ISPOLZOVAN.HeaderText = "Период испол-я (мес.)";
            this.PERIOD_ISPOLZOVAN.Name = "PERIOD_ISPOLZOVAN";
            this.PERIOD_ISPOLZOVAN.ReadOnly = true;
            // 
            // PRICE_EDENIC_PLAN
            // 
            this.PRICE_EDENIC_PLAN.HeaderText = "Дата последн. выдачи";
            this.PRICE_EDENIC_PLAN.Name = "PRICE_EDENIC_PLAN";
            this.PRICE_EDENIC_PLAN.ReadOnly = true;
            // 
            // PRICE_KOMPLEKT_PLAN
            // 
            this.PRICE_KOMPLEKT_PLAN.HeaderText = "Дата след. выдачи";
            this.PRICE_KOMPLEKT_PLAN.Name = "PRICE_KOMPLEKT_PLAN";
            this.PRICE_KOMPLEKT_PLAN.ReadOnly = true;
            // 
            // SIZE
            // 
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.White;
            this.SIZE.DefaultCellStyle = dataGridViewCellStyle3;
            this.SIZE.HeaderText = "Размер";
            this.SIZE.Name = "SIZE";
            this.SIZE.ReadOnly = true;
            this.SIZE.Width = 60;
            // 
            // USER_ID
            // 
            this.USER_ID.HeaderText = "USER_ID";
            this.USER_ID.Name = "USER_ID";
            this.USER_ID.ReadOnly = true;
            this.USER_ID.Visible = false;
            // 
            // NAIMEN_IAIS
            // 
            this.NAIMEN_IAIS.HeaderText = "Наименование ИАИС";
            this.NAIMEN_IAIS.Name = "NAIMEN_IAIS";
            this.NAIMEN_IAIS.ReadOnly = true;
            this.NAIMEN_IAIS.Width = 200;
            // 
            // ID_IAIS
            // 
            this.ID_IAIS.HeaderText = "ID_IAIS";
            this.ID_IAIS.Name = "ID_IAIS";
            this.ID_IAIS.ReadOnly = true;
            this.ID_IAIS.Visible = false;
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(1021, 40);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(140, 20);
            this.dateTimePicker1.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Times New Roman", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.ForeColor = System.Drawing.Color.Green;
            this.label3.Location = new System.Drawing.Point(898, 41);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(120, 19);
            this.label3.TabIndex = 11;
            this.label3.Text = "Дата списания:";
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(358, 432);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(484, 62);
            this.richTextBox1.TabIndex = 12;
            this.richTextBox1.Text = "";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label4.ForeColor = System.Drawing.Color.Red;
            this.label4.Location = new System.Drawing.Point(355, 413);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(121, 16);
            this.label4.TabIndex = 13;
            this.label4.Text = "Заметки к акту";
            // 
            // Form12
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1173, 506);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form12";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Формирование акта на списание";
            this.Load += new System.EventHandler(this.Form12_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        public System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DataGridViewTextBoxColumn ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn SOTRUDNIK_ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn NORMS_LIST_ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn NAIMEN_OVERALLS;
        private System.Windows.Forms.DataGridViewTextBoxColumn ED_IZM;
        private System.Windows.Forms.DataGridViewTextBoxColumn KOLVO;
        private System.Windows.Forms.DataGridViewTextBoxColumn PERIOD_ISPOLZOVAN;
        private System.Windows.Forms.DataGridViewTextBoxColumn PRICE_EDENIC_PLAN;
        private System.Windows.Forms.DataGridViewTextBoxColumn PRICE_KOMPLEKT_PLAN;
        private System.Windows.Forms.DataGridViewTextBoxColumn SIZE;
        private System.Windows.Forms.DataGridViewTextBoxColumn USER_ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn NAIMEN_IAIS;
        private System.Windows.Forms.DataGridViewTextBoxColumn ID_IAIS;
    }
}