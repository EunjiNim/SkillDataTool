namespace SkillDataTool
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
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView2).BeginInit();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(12, 12);
            button1.Name = "button1";
            button1.Size = new Size(76, 36);
            button1.TabIndex = 0;
            button1.Text = "Open";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button2
            // 
            button2.Location = new Point(12, 54);
            button2.Name = "button2";
            button2.Size = new Size(76, 39);
            button2.TabIndex = 1;
            button2.Text = "Load";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(15, 171);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowTemplate.Height = 25;
            dataGridView1.Size = new Size(773, 193);
            dataGridView1.TabIndex = 2;
            // 
            // listBox1
            // 
            listBox1.FormattingEnabled = true;
            listBox1.ItemHeight = 15;
            listBox1.Location = new Point(94, 12);
            listBox1.Name = "listBox1";
            listBox1.Size = new Size(318, 79);
            listBox1.TabIndex = 3;
            // 
            // textBox1
            // 
            textBox1.Location = new Point(560, 12);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(142, 23);
            textBox1.TabIndex = 4;
            // 
            // button3
            // 
            button3.Location = new Point(719, 12);
            button3.Name = "button3";
            button3.Size = new Size(69, 23);
            button3.TabIndex = 5;
            button3.Text = "Find";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // comboBox1
            // 
            comboBox1.FormattingEnabled = true;
            comboBox1.Location = new Point(719, 41);
            comboBox1.Name = "comboBox1";
            comboBox1.Size = new Size(69, 23);
            comboBox1.TabIndex = 6;
            comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(12, 153);
            label1.Name = "label1";
            label1.Size = new Size(131, 15);
            label1.TabIndex = 7;
            label1.Text = "Skill Effect Level Group";
            // 
            // textBox2
            // 
            textBox2.Location = new Point(17, 117);
            textBox2.Name = "textBox2";
            textBox2.Size = new Size(163, 23);
            textBox2.TabIndex = 8;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(17, 99);
            label2.Name = "label2";
            label2.Size = new Size(65, 15);
            label2.TabIndex = 9;
            label2.Text = "Skill Name";
            // 
            // textBox3
            // 
            textBox3.Location = new Point(213, 117);
            textBox3.Name = "textBox3";
            textBox3.Size = new Size(144, 23);
            textBox3.TabIndex = 10;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(213, 99);
            label3.Name = "label3";
            label3.Size = new Size(56, 15);
            label3.TabIndex = 11;
            label3.Text = "Cooltime";
            // 
            // dataGridView2
            // 
            dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView2.Location = new Point(15, 417);
            dataGridView2.Name = "dataGridView2";
            dataGridView2.RowTemplate.Height = 25;
            dataGridView2.Size = new Size(773, 86);
            dataGridView2.TabIndex = 12;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(15, 399);
            label4.Name = "label4";
            label4.Size = new Size(120, 15);
            label4.TabIndex = 13;
            label4.Text = "Skill Effect Operation";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 523);
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
            Text = "Form1";
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
    }
}