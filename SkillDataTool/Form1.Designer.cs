﻿namespace SkillDataTool
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            button1 = new Button();
            button2 = new Button();
            dataGridView1 = new DataGridView();
            listBox1 = new ListBox();
            textBox1 = new TextBox();
            button3 = new Button();
            comboBox1 = new ComboBox();
            label1 = new Label();
            textBox2 = new TextBox();
            label2 = new Label();
            textBox3 = new TextBox();
            label3 = new Label();
            dataGridView2 = new DataGridView();
            label4 = new Label();
            label6 = new Label();
            metroLabel1 = new MetroFramework.Controls.MetroLabel();
            saveFileDialog1 = new SaveFileDialog();
            metroButton1 = new MetroFramework.Controls.MetroButton();
            metroProgressBar1 = new MetroFramework.Controls.MetroProgressBar();
            backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            textBox4 = new TextBox();
            label7 = new Label();
            label8 = new Label();
            textBox5 = new TextBox();
            label9 = new Label();
            textBox6 = new TextBox();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView2).BeginInit();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(14, 93);
            button1.Name = "button1";
            button1.Size = new Size(76, 36);
            button1.TabIndex = 0;
            button1.Text = "Open";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button2
            // 
            button2.Location = new Point(14, 135);
            button2.Name = "button2";
            button2.Size = new Size(76, 39);
            button2.TabIndex = 1;
            button2.Text = "Load";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // dataGridView1
            // 
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(14, 342);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowTemplate.Height = 25;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.Size = new Size(773, 193);
            dataGridView1.TabIndex = 2;
            // 
            // listBox1
            // 
            listBox1.FormattingEnabled = true;
            listBox1.ItemHeight = 15;
            listBox1.Location = new Point(104, 95);
            listBox1.Name = "listBox1";
            listBox1.Size = new Size(318, 79);
            listBox1.TabIndex = 3;
            // 
            // textBox1
            // 
            textBox1.ForeColor = Color.Black;
            textBox1.Location = new Point(559, 95);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(142, 23);
            textBox1.TabIndex = 20;
            textBox1.KeyDown += searchTextBox_KeyDown;
            // 
            // button3
            // 
            button3.Location = new Point(718, 95);
            button3.Name = "button3";
            button3.Size = new Size(69, 23);
            button3.TabIndex = 5;
            button3.Text = "Find";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // comboBox1
            // 
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox1.FormattingEnabled = true;
            comboBox1.Location = new Point(718, 144);
            comboBox1.Name = "comboBox1";
            comboBox1.Size = new Size(69, 23);
            comboBox1.TabIndex = 6;
            comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(14, 324);
            label1.Name = "label1";
            label1.Size = new Size(131, 15);
            label1.TabIndex = 7;
            label1.Text = "Skill Effect Level Group";
            // 
            // textBox2
            // 
            textBox2.Location = new Point(15, 220);
            textBox2.Name = "textBox2";
            textBox2.Size = new Size(408, 23);
            textBox2.TabIndex = 8;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(15, 202);
            label2.Name = "label2";
            label2.Size = new Size(65, 15);
            label2.TabIndex = 9;
            label2.Text = "Skill Name";
            // 
            // textBox3
            // 
            textBox3.Location = new Point(163, 284);
            textBox3.Name = "textBox3";
            textBox3.Size = new Size(144, 23);
            textBox3.TabIndex = 10;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(163, 266);
            label3.Name = "label3";
            label3.Size = new Size(56, 15);
            label3.TabIndex = 11;
            label3.Text = "Cooltime";
            // 
            // dataGridView2
            // 
            dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView2.Location = new Point(14, 565);
            dataGridView2.Name = "dataGridView2";
            dataGridView2.RowTemplate.Height = 25;
            dataGridView2.Size = new Size(773, 139);
            dataGridView2.TabIndex = 12;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(14, 547);
            label4.Name = "label4";
            label4.Size = new Size(120, 15);
            label4.TabIndex = 13;
            label4.Text = "Skill Effect Operation";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Font = new Font("맑은 고딕", 10F, FontStyle.Regular, GraphicsUnit.Point);
            label6.Location = new Point(604, 146);
            label6.Name = "label6";
            label6.Size = new Size(108, 19);
            label6.TabIndex = 15;
            label6.Text = "Skill Effect Level";
            // 
            // metroLabel1
            // 
            metroLabel1.AutoSize = true;
            metroLabel1.Font = new Font("맑은 고딕", 10F, FontStyle.Regular, GraphicsUnit.Point);
            metroLabel1.Location = new Point(513, 95);
            metroLabel1.Name = "metroLabel1";
            metroLabel1.Size = new Size(40, 19);
            metroLabel1.TabIndex = 16;
            metroLabel1.Text = "Index";
            // 
            // metroButton1
            // 
            metroButton1.Font = new Font("맑은 고딕", 11F, FontStyle.Regular, GraphicsUnit.Point);
            metroButton1.Location = new Point(654, 284);
            metroButton1.Name = "metroButton1";
            metroButton1.Size = new Size(133, 23);
            metroButton1.TabIndex = 17;
            metroButton1.Text = "SaveFile";
            metroButton1.UseSelectable = true;
            metroButton1.Click += metroButton1_Click;
            // 
            // metroProgressBar1
            // 
            metroProgressBar1.Location = new Point(15, 180);
            metroProgressBar1.Name = "metroProgressBar1";
            metroProgressBar1.Size = new Size(408, 10);
            metroProgressBar1.TabIndex = 18;
            metroProgressBar1.Click += metroProgressBar1_Click;
            // 
            // backgroundWorker1
            // 
            backgroundWorker1.DoWork += backgroundWorker1_DoWork;
            // 
            // textBox4
            // 
            textBox4.Location = new Point(313, 284);
            textBox4.Name = "textBox4";
            textBox4.Size = new Size(163, 23);
            textBox4.TabIndex = 22;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(313, 266);
            label7.Name = "label7";
            label7.Size = new Size(131, 15);
            label7.TabIndex = 23;
            label7.Text = "Combo Element Count";
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.Location = new Point(482, 266);
            label8.Name = "label8";
            label8.Size = new Size(40, 15);
            label8.TabIndex = 25;
            label8.Text = "Target";
            // 
            // textBox5
            // 
            textBox5.Location = new Point(482, 284);
            textBox5.Name = "textBox5";
            textBox5.Size = new Size(163, 23);
            textBox5.TabIndex = 24;
            // 
            // label9
            // 
            label9.AutoSize = true;
            label9.Location = new Point(15, 266);
            label9.Name = "label9";
            label9.Size = new Size(97, 15);
            label9.TabIndex = 27;
            label9.Text = "Skill Show Order";
            // 
            // textBox6
            // 
            textBox6.Location = new Point(15, 284);
            textBox6.Name = "textBox6";
            textBox6.Size = new Size(142, 23);
            textBox6.TabIndex = 26;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoScroll = true;
            ClientSize = new Size(797, 713);
            Controls.Add(label9);
            Controls.Add(textBox6);
            Controls.Add(label8);
            Controls.Add(textBox5);
            Controls.Add(label7);
            Controls.Add(textBox4);
            Controls.Add(metroProgressBar1);
            Controls.Add(metroButton1);
            Controls.Add(metroLabel1);
            Controls.Add(label6);
            Controls.Add(label4);
            Controls.Add(dataGridView2);
            Controls.Add(label3);
            Controls.Add(textBox3);
            Controls.Add(label2);
            Controls.Add(textBox2);
            Controls.Add(label1);
            Controls.Add(comboBox1);
            Controls.Add(button3);
            Controls.Add(textBox1);
            Controls.Add(listBox1);
            Controls.Add(dataGridView1);
            Controls.Add(button2);
            Controls.Add(button1);
            Name = "Form1";
            TransparencyKey = Color.Empty;
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView2).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private Button button2;
        private DataGridView dataGridView1;
        private ListBox listBox1;
        private TextBox textBox1;
        private Button button3;
        private ComboBox comboBox1;
        private Label label1;
        private TextBox textBox2;
        private Label label2;
        private TextBox textBox3;
        private Label label3;
        private DataGridView dataGridView2;
        private Label label4;
        private Label label6;
        private MetroFramework.Controls.MetroLabel metroLabel1;
        private SaveFileDialog saveFileDialog1;
        private MetroFramework.Controls.MetroButton metroButton1;
        private MetroFramework.Controls.MetroProgressBar metroProgressBar1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private TextBox textBox4;
        private Label label7;
        private Label label8;
        private TextBox textBox5;
        private Label label9;
        private TextBox textBox6;
    }
}