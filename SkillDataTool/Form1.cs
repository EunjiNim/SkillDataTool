using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Drawing.Text;


using Application = System.Windows.Forms.Application;
using DataTable = System.Data.DataTable;
using DocumentFormat.OpenXml;

namespace SkillDataTool
{
    public partial class Form1 : Form
    {
        private string Excel07Constring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=NO;IMEX=1;'";


        // �̹� �����Ͱ� ���� �ִ��� �Ǻ��� �� ���
        private string Skill = string.Empty;
        private string SkillEffect = string.Empty;
        private string SkillEffectLevelGroup = string.Empty;

        private int Level_Index = 0;

        // �޾ƿ� ���� ��Ʈ �̸�
        private string Sheet_Name = "Table$";

        // �������� ������ �����͸� ��ųʸ� ������ �������. Value�� ������Ʈ �迭 �������� ����, �˻� �� ������ �÷� ���� �����ϱ� ������ �ϱ� ����
        private Dictionary<string, object[]> SkillData = new Dictionary<string, object[]>();
        private Dictionary<string, object[]> SkillEffectData = new Dictionary<string, object[]>();

        // SkillEffectLevelGroup�� ������ Ű ���� �����Ƿ� ���߰� ��ųʸ��� �����ϱ� ���� Value�� List �������� �־���
        private Dictionary<string, List<object[]>> SkillEffectLevelGroupData = new Dictionary<string, List<object[]>>();

        // �ʿ��� �����͸� ���� ��� �׸��� �信 ����ֱ� ���� ���
        private DataTable GridViewInData = new DataTable();


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        // ���� ���� ���� ��ư 
        // �ߺ��Ǵ� ������ ������ ���� �����͸� ���� �Ͱ� ������ ���ε��� �и���Ŵ
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                // ���� ���ϸ� ���� �� �ֵ��� ��
                DefaultExt = "xlsx",
                Multiselect = true,
                Filter = "TextFile(*.xls, *.xlsx) |*.xlsx;*.xls",
                FileName = string.Empty
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (openFileDialog.FileName.Length > 0)
                    {
                        string[] filenames = openFileDialog.FileNames;

                        for (int i = 0; i < (int)filenames.Length; i++)
                        {
                            string str = filenames[i];
                            openFileDialog.FileNames[i] = str;

                            string[] listboxs = openFileDialog.SafeFileNames;
                            string listbox = listboxs[i];
                            openFileDialog.SafeFileNames[i] = listbox;

                            // �ҷ��� ���� ���� �̸��� �����
                            this.listBox1.Items.Add(listbox);

                            // ���� ������ ���� ���� �ٸ��� �ּҸ� ������ ��
                            if (str.Contains("Skill.xlsx"))
                            {
                                if (this.Skill.Length != 0)
                                {
                                    Console.WriteLine("�̹� ������ �����Ͱ� �ֽ��ϴ�.");
                                }
                                else
                                {
                                    this.Skill = string.Format(this.Excel07Constring, str, 0);
                                }
                            }
                            else if (str.Contains("SkillEffect.xlsx"))
                            {
                                if (this.SkillEffect.Length != 0)
                                {
                                    Console.WriteLine("�̹� ������ �����Ͱ� �ֽ��ϴ�.");

                                }
                                else
                                {
                                    this.SkillEffect = string.Format(this.Excel07Constring, str, 0);
                                }
                            }
                            else if (str.Contains("SkillEffectLevelGroup.xlsx"))
                            {
                                if (this.SkillEffectLevelGroup.Length != 0)
                                {
                                    Console.WriteLine("�̹� ������ �����Ͱ� �ֽ��ϴ�.");

                                }
                                else
                                {
                                    this.SkillEffectLevelGroup = string.Format(this.Excel07Constring, str, 0);
                                }
                            }
                            else
                            {
                                MessageBox.Show("����� �� ���� �����Դϴ�. ������ �ٽ� Ȯ���� �ּ���.");
                                Application.Restart();
                            }


                        }
                    }

                }
                catch (Exception ex)
                {
                    // �̹� �ҷ��� �����Ͱ� ���� ��� ������ ���������
                    Exception exception = ex;
                    MessageBox.Show(ex.InnerException != null ? ex.InnerException.Message : "�̹� ����� �����Ͱ� �־� ���α׷��� �ٽ� �����մϴ�.");
                    Application.Restart();

                }
            }

        }

        // ������ ���ε��ϱ�
        private void button2_Click(object sender, EventArgs e)
        {
            // ����ó�� 
            if (this.Skill.Length == 0)
            {
                // ���� ��Ʈ ��ΰ� �� ���� ����
                MessageBox.Show("�����Ͱ� �������� �ʽ��ϴ�.");
                return;
            }
            if (this.SkillData.Values.Count != 0)
            {
                // �̹� ó���� �������� ���
                MessageBox.Show("�̹� �ε尡 �Ϸ�� �������Դϴ�.");
                return;
            }

            using (OleDbConnection conn = new OleDbConnection(this.Skill))
            {
                using (OleDbCommand comm = new OleDbCommand())
                {
                    using (OleDbDataAdapter adap = new OleDbDataAdapter())
                    {
                        // Skill ������ ���
                        DataTable datatable1 = new DataTable();
                        comm.CommandText = string.Concat("SELECT * From [", this.Sheet_Name, "]");
                        comm.Connection = conn;
                        conn.Open();
                        adap.SelectCommand = comm;
                        adap.Fill(datatable1);

                        foreach (DataRow row in datatable1.Rows)
                        {
                            // ������ ���̺� ��ܿ� ������� �ʴ� �����Ͱ� 7�� ���� ���ԵǾ� ����
                            // ������ ID�� Ű ������ ����ؾ� �ϴµ� ������ ������ ���� �־� �� ��� ������ �κк��͸� �����͸� ������� �� �ֵ��� ��
                            if (row.Table.Rows.IndexOf(row) > 7)
                            {
                                if (!SkillData.ContainsKey(row.ItemArray[0].ToString()))
                                {
                                    // 0���� ID�� Ű������ ������ ��� �ش� array�� ��� ������. ���� ID�� �ε����� �Ͽ� �˻��ϱ� ����
                                    SkillData.Add(row.ItemArray[0].ToString(), row.ItemArray);
                                }
                                else
                                {
                                    // �ߺ��Ǵ� ���� ������ ��� ������ ����Ƿ� Ȯ���� ���� ��Ƽ�� ��
                                    MessageBox.Show("�ߺ��Ǵ� Ű ���� �ֽ��ϴ�. " + row.ItemArray[0].ToString() + " �����͸� Ȯ���� �ּ���.");
                                }
                            }
                            conn.Close();
                        }
                    }
                }
            }

            using (OleDbConnection conn = new OleDbConnection(this.SkillEffect))
            {
                using (OleDbCommand comm = conn.CreateCommand())
                {
                    using (OleDbDataAdapter adap = new OleDbDataAdapter())
                    {
                        // SkillEffect ������ ���
                        DataTable dataTable2 = new DataTable();
                        comm.CommandText = string.Concat("SELECT * From [", this.Sheet_Name, "]");
                        comm.Connection = conn;
                        conn.Open();
                        adap.SelectCommand = comm;
                        adap.Fill(dataTable2);

                        foreach (DataRow row in dataTable2.Rows)
                        {
                            if (row.Table.Rows.IndexOf(row) > 7)
                            {
                                if (!SkillEffectData.ContainsKey(row.ItemArray[0].ToString()))
                                {
                                    SkillEffectData.Add(row.ItemArray[0].ToString(), row.ItemArray);
                                }
                                else
                                {
                                    MessageBox.Show("�ߺ��Ǵ� Ű ���� �ֽ��ϴ�. " + row.ItemArray[0].ToString() + " �����͸� Ȯ���� �ּ���.");
                                }
                            }
                            conn.Close();
                        }
                    }
                }
            }

            using (OleDbConnection conn = new OleDbConnection(this.SkillEffectLevelGroup))
            {
                using (OleDbCommand comm = conn.CreateCommand())
                {
                    using (OleDbDataAdapter adap = new OleDbDataAdapter())
                    {
                        // SkillEffectLevelGroup ������ ���
                        DataTable dataTable3 = new DataTable();
                        comm.CommandText = string.Concat("SELECT * From [", this.Sheet_Name, "]");
                        comm.Connection = conn;
                        conn.Open();
                        adap.SelectCommand = comm;
                        adap.Fill(dataTable3);

                        foreach (DataRow row in dataTable3.Rows)
                        {
                            if (row.Table.Rows.IndexOf(row) > 7)
                            {
                                // SkillEffectLevelGroupData�� ID ���� ������ �׷����� ����ϹǷ� �ش� Ű ���� �����Ǿ� ���� ��� ����Ʈ �������� ���� �־���
                                // Dictonary�� Ű ���� �ߺ����� ����� �� ���� ������ Value�� ����Ʈ �������� �����Ͽ� ����� �� �ֵ��� ����� �����
                                if (SkillEffectLevelGroupData.ContainsKey(row.ItemArray[0].ToString()))
                                {
                                    SkillEffectLevelGroupData[row.ItemArray[0].ToString()].Add(row.ItemArray);
                                }
                                else
                                {
                                    // Ű ���� ���� ��� ���Ӱ� ����Ʈ�� ������
                                    SkillEffectLevelGroupData[row.ItemArray[0].ToString()] = new List<object[]> { row.ItemArray };
                                }
                            }
                            conn.Close();
                        }
                    }
                }
            }
        }


        // �ε��� �˻� ���
        private void button3_Click(object sender, EventArgs e)
        {
            // �÷� ��ġ�� ����� �� �����Ƿ� ����Ʈ�� �־� �̸��� ã�� �� �÷��� �ε����� ��ȯ�� �� �ֵ��� ��
            /*object[] ItemDataIndex = new object[1];
            SkillData.TryGetValue("skill_id", out ItemDataIndex);*/


            // 10 << ��ų ����Ʈ ID �̹Ƿ� �ش� �ε��� ������ ã�Ƶ��� ��

            // 100100111

            // SkillEffectLevelGroupData�� ã���� �ϴ� ������ �˾ƾ� �����͸� ������ �� ���� ����
            List<object[]> ItemDataIndex = new List<object[]>();
            SkillEffectLevelGroupData.TryGetValue("skill_effect_level_group_id", out ItemDataIndex);

            List<object[]> Search_SkillEffectLevelData = new List<object[]>();
            SkillEffectLevelGroupData.TryGetValue("100100111", out Search_SkillEffectLevelData);




            // SkillEffectLevelGroup�� �÷��� ���������� �׸��� �信 �־���
            for (int i = 0; i < ItemDataIndex[0].ToArray().Length; ++i)
            {
                GridViewInData.Columns.Add(ItemDataIndex[0].ToArray()[i].ToString());
            }

            List<string> Level_List = new List<string>();

            // SkillEffectLevelGroup�� ���� ����Ʈ�� �޺��ڽ��� �־���
            foreach (var item in Search_SkillEffectLevelData)
            {
                Level_List.Add(item[4].ToString());
            }
            this.comboBox1.Items.AddRange(Level_List.ToArray());



            // �˻��� �ε����� row �����͸� ��� �׸��� �信 �־���
            //GridViewInData.Rows.Add(testdata[1]);
            foreach (var test in Search_SkillEffectLevelData)
            {
                GridViewInData.Rows.Add(test);
            }


            // ������ �׸��� �信 ��Ƶ� �����͸� ��� ��
            this.dataGridView1.DataSource = GridViewInData;

        }

        // �޺��ڽ� �׽�Ʈ
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int Skill_Level = 0;
            
            // �޺��ڽ��� string ������ int�� ��ȯ�� �ְ� �ε����� �����ֱ� ���� 1�� ��
            Skill_Level = Convert.ToInt16(comboBox1.Text) - 1;      

            // ����Ʈ �ʱ�ȭ
            GridViewInData.Clear();

            List<object[]> Search_SkillEffectLevelData = new List<object[]>();
            SkillEffectLevelGroupData.TryGetValue("100100111", out Search_SkillEffectLevelData);

            // Ư�� ������ ������ ����� ��
            GridViewInData.Rows.Add(Search_SkillEffectLevelData[Skill_Level]);
            this.dataGridView1.DataSource = GridViewInData;

        }
    }
}