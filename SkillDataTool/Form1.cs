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

using MetroFramework.Forms;
using DocumentFormat.OpenXml.Presentation;
using System.Runtime.InteropServices;
using OfficeOpenXml;
using Range = Microsoft.Office.Interop.Excel.Range;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2016.Drawing.Charts;
using System.Collections.Immutable;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace SkillDataTool
{
    public partial class Form1 : MetroForm
    {
        private string Excel07Constring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=NO;IMEX=1;'";


        // �̹� �����Ͱ� ���� �ִ��� �Ǻ��� �� ���
        private string Skill = string.Empty;
        private string SkillEffect = string.Empty;
        private string SkillEffectLevelGroup = string.Empty;
        private string SkillEffectOperation = string.Empty;

        // ����ڰ� �Է��� �ε����� ����
        private string Index_Num = string.Empty;

        // ����ڰ� �Է��� �ε����� �������� ����� �ε����� �޾ƿ�
        private string SkillEffect_Num = string.Empty;
        private string SkillEffectOperation_Num = string.Empty;

        private int Level_Index = 0;
        private int AllFile_Count = 0;

        // �޾ƿ� ���� ��Ʈ �̸�
        private string Sheet_Name = "Table$";

        // �������� ������ �����͸� ��ųʸ� ������ �������. Value�� ������Ʈ �迭 �������� ����, �˻� �� ������ �÷� ���� �����ϱ� ������ �ϱ� ����
        private Dictionary<string, object[]>? SkillData = new Dictionary<string, object[]>();
        private Dictionary<string, object[]>? SkillEffectData = new Dictionary<string, object[]>();
        private Dictionary<string, object[]>? SkillEffectOperationData = new Dictionary<string, object[]>();

        // SkillEffectLevelGroup�� ������ Ű ���� �����Ƿ� ���߰� ��ųʸ��� �����ϱ� ���� Value�� List �������� �־���
        private Dictionary<string, List<object[]>>? SkillEffectLevelGroupData = new Dictionary<string, List<object[]>>();

        // �ʿ��� �����͸� ���� ��� �׸��� �信 ����ֱ� ���� ���
        private DataTable? GridViewInData = new DataTable();
        private DataTable? GridViewOperationData = new DataTable();


        private DataTable? ConvertGridViewInData = new DataTable();




        public Form1()
        {
            InitializeComponent();

            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
            backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = "Skill Search Tool";
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
                                    //Workbook wb = new Workbook(this.Skill.ToString());
                                }
                            }
                            else if (str.Contains("SkillEffectGroup.xlsx"))
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
                            else if (str.Contains("SkillEffectOperation.xlsx"))
                            {
                                if (this.SkillEffectOperation.Length != 0)
                                {
                                    Console.WriteLine("�̹� ������ �����Ͱ� �ֽ��ϴ�.");

                                }
                                else
                                {
                                    this.SkillEffectOperation = string.Format(this.Excel07Constring, str, 0);
                                }
                            }
                            else
                            {
                                MetroFramework.MetroMessageBox.Show(this, "����� �� ���� �����Դϴ�. ������ �ٽ� Ȯ���� �ּ���.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning, 100);
                            }
                        }
                    }

                }
                catch (Exception ex)
                {
                    // �̹� �ҷ��� �����Ͱ� ���� ��� ������ ���������
                    Exception exception = ex;
                    MessageBox.Show(ex.InnerException != null ? ex.InnerException.Message : "������ �߻��Ͽ� ���α׷��� �ٽ� �����մϴ�.");

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
                this.listBox1.Items.Clear();
                return;
            }
            if (this.SkillData.Values.Count != 0)
            {
                // �̹� ó���� �������� ���
                MessageBox.Show("�̹� �ε尡 �Ϸ�� �������Դϴ�.");
                return;
            }

            int z = 0;


            // backgroundWorker�� �̿��ؼ� progressbar�� ���������
            backgroundWorker1.RunWorkerAsync();

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
                        }

                        comm.Dispose();
                        conn.Close();
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
                        }

                        comm.Dispose();
                        conn.Close();
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


                        // ������� ������ �÷� �ƿ� ����. �����͸� �����鼭 �Ǽ��� �����Ͽ� �� �÷��� �ڲ� ����
                        for (int i = dataTable3.Columns.Count - 1; i >= 0; i--)
                        {
                            for (int j = 0; j < dataTable3.Rows.Count; j++)
                            {
                                if (dataTable3.Rows[j][i].ToString() == "")
                                {
                                    if (j == dataTable3.Rows.Count - 1)
                                    {
                                        dataTable3.Columns.RemoveAt(i);
                                    }
                                    continue;
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }



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
                        }

                        comm.Dispose();
                        conn.Close();
                    }
                }
            }

            using (OleDbConnection conn = new OleDbConnection(this.SkillEffectOperation))
            {
                using (OleDbCommand comm = conn.CreateCommand())
                {
                    using (OleDbDataAdapter adap = new OleDbDataAdapter())
                    {
                        // SkillEffectOperation ������ ���
                        DataTable dataTable4 = new DataTable();
                        comm.CommandText = string.Concat("SELECT * From [", this.Sheet_Name, "]");
                        comm.Connection = conn;
                        conn.Open();
                        adap.SelectCommand = comm;
                        adap.Fill(dataTable4);

                        // ������� ������ �÷� �ƿ� ����. �����͸� �����鼭 �Ǽ��� �����Ͽ� �� �÷��� �ڲ� ����
                        for (int i = dataTable4.Columns.Count - 1; i >= 0; i--)
                        {
                            for (int j = 0; j < dataTable4.Rows.Count; j++)
                            {
                                if (dataTable4.Rows[j][i].ToString() == "")
                                {
                                    if (j == dataTable4.Rows.Count - 1)
                                    {
                                        dataTable4.Columns.RemoveAt(i);
                                    }
                                    continue;
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }

                        foreach (DataRow row in dataTable4.Rows)
                        {
                            if (row.Table.Rows.IndexOf(row) > 7)
                            {
                                if (!SkillEffectOperationData.ContainsKey(row.ItemArray[0].ToString()))
                                {
                                    SkillEffectOperationData.Add(row.ItemArray[0].ToString(), row.ItemArray);
                                }
                                else
                                {
                                    MessageBox.Show("�ߺ��Ǵ� Ű ���� �ֽ��ϴ�. " + row.ItemArray[0].ToString() + " �����͸� Ȯ���� �ּ���.");
                                }
                            }
                        }

                        comm.Dispose();
                        conn.Close();
                    }
                }
            }

            // ������ �ε��� �Ϸ�Ǹ� �ؽ�Ʈ�� ������ ��
            if (SkillData.Count != null && SkillEffectData.Count != null && SkillEffectLevelGroupData.Count != null && SkillEffectOperationData.Count != null)
            {
                this.button2.Text = "Complate";
                backgroundWorker1.CancelAsync();

            }

        }


        // �ε��� �˻� ���
        private void button3_Click(object sender, EventArgs e)
        {
            // ������ �ؽ�Ʈ �ڽ��� �Է��� ��ġ�� �־���
            this.Index_Num = this.textBox1.Text;

            // �ε��� �ѹ��� Skill Data�� ã�ƿ�
            object[]? SkillDatasRow = new object[1];
            SkillData.TryGetValue(Index_Num, out SkillDatasRow);

            // �˻��� Ű�� ���� ��쿡�� �ؽ�Ʈ �ڽ��� �����, ���� �˾��� �����
            if (!SkillData.ContainsKey(Index_Num))
            {
                MessageBox.Show("�������� �ʴ� �ε����Դϴ�.");
                this.textBox1.Clear();
                return;
            }

            // ��˻��ø��� ����Ʈ�� ���̹Ƿ� ��ư�� Ŭ���Ҷ����� �ʱ�ȭ �� ��
            this.comboBox1.Items.Clear();
            GridViewInData.Clear();
            GridViewOperationData.Clear();
            // �޺��ڽ� �ؽ�Ʈ ����
            this.comboBox1.ResetText();

            // �÷� ��ġ�� ����� �� �����Ƿ� ����Ʈ�� �־� �̸��� ã�� �� �÷��� �ε����� ��ȯ�� �� �ֵ��� ��
            object[]? SkillDataIndex = new object[1];
            SkillData.TryGetValue("skill_id", out SkillDataIndex);

            // �ε����� �´� ��ų ������ ã����
            object[]? Search_SkillData = new object[1];
            SkillData.TryGetValue(Index_Num, out Search_SkillData);

            // ��ų �����Ϳ��� ��ų ����Ʈ �ѹ� �޾ƿ�
            SkillEffect_Num = SkillDatasRow.ToArray()[SkillDataIndex.ToList().IndexOf("link_skill_effect_id")].ToString();

            // ��ų ����Ʈ ������ �ε����� ���۷��̼� �ε����� ã�Ƴ��� ���� SkillEffectGroup ��Ʈ ����
            object[]? search_SkillEffectData = new object[1];
            SkillEffectData.TryGetValue(SkillEffect_Num, out search_SkillEffectData);

            if (search_SkillEffectData != null)
            {
                // Skill Effect Operation ������ �̾Ƴ�
                SkillEffectOperation_Num = search_SkillEffectData.ToArray()[SkillEffectData.ToArray()[1].Value.ToList().IndexOf("link_skill_effect_operation_id")].ToString();

            }

            // �ε����� �̱����� ����Ʈ�� �־���
            List<string>? SkillIndexList = new List<string>();
            foreach (Object list in SkillDataIndex)
            {
                SkillIndexList.Add(list.ToString());
            }

            // SkillEffectLevelGroupData�� ã���� �ϴ� ������ �˾ƾ� �����͸� ������ �� ���� ����
            List<object[]>? ItemDataIndex = new List<object[]>();
            SkillEffectLevelGroupData.TryGetValue("��ų ȿ�� ���� �׷�", out ItemDataIndex);


            // skill ��Ʈ���� ã�Ƴ� ��ų ����Ʈ�� �ε����� �̾Ƴ��� �� �ε����� SkillEffectLevelGroup ��Ʈ���� �����͸� ã��
            List<object[]>? Search_SkillEffectLevelData = new List<object[]>();
            SkillEffectLevelGroupData.TryGetValue(SkillEffect_Num, out Search_SkillEffectLevelData);


            // �׸��� �信 �÷� ���� �������� ������ �־���
            if (GridViewInData.Columns.Count == 0)
            {
                // SkillEffectLevelGroup�� �÷��� ���������� �׸��� �信 �־���
                for (int i = 0; i < ItemDataIndex[0].ToArray().Length; ++i)
                {
                    GridViewInData.Columns.Add(ItemDataIndex[0].ToArray()[i].ToString());
                    ConvertGridViewInData.Columns.Add(ItemDataIndex[0].ToArray()[i].ToString());
                }
            }

            // Skill Effect Operation ������ ������ �׸��� �信 �־��ֱ� ���� �ε��� �˻�
            object[]? SkillEffectOperation_Result = new object[1];
            SkillEffectOperationData.TryGetValue(SkillEffectOperation_Num, out SkillEffectOperation_Result);

            object[]? SkillEffectOperation_Column = new object[1];
            SkillEffectOperationData.TryGetValue("��ų ȿ�� �۵� ���̵�", out SkillEffectOperation_Column);

            // ���� ��ų�� operation ������ �����Ƿ� �˻��� �� ���� ���������� �׸��� �信 �־� �� 
            if (SkillEffectOperation_Result != null)
            {
                this.dataGridView2.Show();
                // �׸��� �信 �÷��� �������� ���� �־���
                if (GridViewOperationData.Columns.Count == 0)
                {
                    for (int i = 0; i < SkillEffectOperation_Column.Length; ++i)
                    {
                        GridViewOperationData.Columns.Add(SkillEffectOperation_Column.ToArray()[i].ToString());
                    }
                }
                // �˻��� ���۷��̼� ������ �־���
                GridViewOperationData.Rows.Add(SkillEffectOperation_Result);
                this.dataGridView2.DataSource = GridViewOperationData;
            }
            else
            {
                // �����Ͱ� ������ �׸��� �並 �ƿ� �������
                this.dataGridView2.Hide();
            }

            List<string>? Level_List = new List<string>();

            // ���� �ܰ谡 �������� �����Ƿ� ����ó�� ��
            if (Search_SkillEffectLevelData == null)
            {
                this.dataGridView1.Hide();
            }
            else
            {
                foreach (var item in Search_SkillEffectLevelData)
                {
                    Level_List.Add(item[ItemDataIndex[0].ToList().IndexOf("����")].ToString());
                }

                // �޺� �ڽ��� ���� ������ ���� ����Ʈ�� �־���
                this.comboBox1.Items.AddRange(Level_List.ToArray());

                // �˻��� �ε����� row �����͸� ��� �׸��� �信 �־���
                foreach (var SkillEffectResualtData in Search_SkillEffectLevelData)
                {
                    GridViewInData.Rows.Add(SkillEffectResualtData);
                }

                this.dataGridView1.Show();
            }

            // ���� �ܰ谡 �������� �����Ƿ� ����ó�� ��
            /*if (Search_SkillEffectLevelData.Count != 0)
             {
                 // SkillEffectLevelGroup�� ���� ����Ʈ�� �޺��ڽ��� �־���
                 foreach (var item in Search_SkillEffectLevelData)
                 {
                     Level_List.Add(item[ItemDataIndex[0].ToList().IndexOf("����")].ToString());
                 }

             }*/


            // �޺� �ڽ��� ���� ������ ���� ����Ʈ�� �־���
            //this.comboBox1.Items.AddRange(Level_List.ToArray());

            // �˻��� �ε����� row �����͸� ��� �׸��� �信 �־���
            /*foreach (var SkillEffectResualtData in Search_SkillEffectLevelData)
            {
                GridViewInData.Rows.Add(SkillEffectResualtData);
            }*/

            // �ؽ�Ʈ �ڽ��� Skill.xlsx �� ������ ����� 
            // �������� ���� �Ҵ��� ���̰� ����Ʈ���� �ٷ� �ε����� �̾Ƽ� �־���. �÷� ������ �ε����� ������ ���� �÷� ��ġ�� ����Ǿ ���ϴ� ���� ������ �� �ֱ� ����
            this.textBox2.Text = Search_SkillData[SkillIndexList.IndexOf("skill_name")].ToString();
            this.textBox3.Text = Search_SkillData[SkillIndexList.IndexOf("skill_cooltime")].ToString();
            this.textBox4.Text = Search_SkillData[SkillIndexList.IndexOf("combo_attribute_element_count")].ToString();
            this.textBox5.Text = Search_SkillData[SkillIndexList.IndexOf("target")].ToString();
            this.textBox6.Text = Search_SkillData[SkillIndexList.IndexOf("skill_show_order")].ToString();


            // ������ �׸��� �信 ��Ƶ� �����͸� ��� ��
            this.dataGridView1.DataSource = GridViewInData;
            this.dataGridView1.AllowUserToOrderColumns = false;

            // �������� ������ ������ ROW�� ����
            this.dataGridView2.AllowUserToAddRows = false;

        }

        // Skill Effect Leve Group ���� �� ���� ���
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int Skill_Level = 0;

            // �޺��ڽ��� string ������ int�� ��ȯ�� �ְ� �ε����� �����ֱ� ���� 1�� ��
            Skill_Level = Convert.ToInt16(comboBox1.Text) - 1;

            // ����Ʈ �ʱ�ȭ
            GridViewInData.Clear();

            List<object[]> Search_SkillEffectLevelData = new List<object[]>();
            SkillEffectLevelGroupData.TryGetValue(SkillEffect_Num, out Search_SkillEffectLevelData);

            // Ư�� ������ ������ ����� ��
            GridViewInData.Rows.Add(Search_SkillEffectLevelData[Skill_Level]);
            this.dataGridView1.DataSource = GridViewInData;
        }

        // ���� ���� ��ư
        private void metroButton1_Click(object sender, EventArgs e)
        {
            /*string filePath = string.Empty;
            //string fileName = @"������������"  + DateHe

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "���� ��θ� �����ϼ���.";
            saveFileDialog.OverwritePrompt = true;
            saveFileDialog.Filter = "Excel ���� ����|*.xlsx";
            saveFileDialog.InitialDirectory = @"D:\";
            
            if(saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = saveFileDialog.FileName;
                // ���������̺� �ִ� �����͸� �������� �־��ְ� ���Ϸ� ��������
 

                 Excel.Application app = new Excel.Application();
                 Excel.Workbook workbook = app.Workbooks.Open(filePath, 0 ,false, 5,"","", false, Excel.XlPlatform.xlWindows, "",true, false, 0, true, false, false);
                 Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[0];
                 Excel.Range range = worksheet.UsedRange;
                 app.Visible = true;

                dataGridview_ExportExcel(saveFileDialog.FileName, dataGridView1, dataGridView2);

            }*/


            string fileName = @"SkillSearchData";
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (ExcelPackage excelPackage = new ExcelPackage())
            {

                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(fileName);

                /* for (int i = 0; i <= GridViewInData.Rows.Count; i++)
                 {
                     for(int j = 0; j <=  GridViewInData.Columns.Count; j++) 
                     {
                         if (int.TryParse(GridViewInData.Rows[i].ItemArray[j].ToString(), out int result))
                         {
                             //ConvertGridViewInData.Rows.Add(result);

                             int data = result;
                             ConvertGridViewInData.Rows[i].ItemArray[j] = data;

                             ConvertGridViewInData.Rows.Add((int)data);
                         }
                         else
                         {
                             ConvertGridViewInData.Rows[i].ItemArray[j] = GridViewInData.Rows[i].ItemArray[j].ToString();
                             //ConvertGridViewInData.Rows.Add(GridViewInData.Rows[i].ItemArray[j].ToString());
                         }
                     }
                 }*/

                // GridViewInData.Columns[0].DataType = typeof(int);


                // Skill Effect Level Group �����͸� �־���
                worksheet.Cells["A1"].LoadFromDataTable(GridViewInData, true, OfficeOpenXml.Table.TableStyles.Light8);

                // Skill Effect Operation  �����͸� �־���
                string index = "A" + (GridViewInData.Rows.Count + 3).ToString();
                worksheet.Cells[index].LoadFromDataTable(GridViewOperationData, true, OfficeOpenXml.Table.TableStyles.Light8);

                // ǥ �������
                worksheet.Columns.AutoFit();


                // Title ���� ��
                /*worksheet.Cells[1, 1, 1, GridViewInData.Columns.Count].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[1, 1, 1, GridViewInData.Columns.Count].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Rows[1].Range.Style.Font.Bold = true;
                // Right ��
                worksheet.Cells[1, GridViewInData.Columns.Count, GridViewInData.Rows.Count +1, GridViewInData.Columns.Count].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                // Bottom ��
                worksheet.Cells[GridViewInData.Rows.Count +1, 1, GridViewInData.Rows.Count +1, GridViewInData.Columns.Count].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                // Left ��
                worksheet.Cells[1, 1, GridViewInData.Rows.Count + 1, GridViewInData.Columns.Count].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;*/

                // Save File Dialog �⺻ ����
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Title = "���� ��θ� �����ϼ���.";
                saveFileDialog.OverwritePrompt = true;
                saveFileDialog.Filter = "Excel ���� ����|*.xlsx";
                saveFileDialog.InitialDirectory = @"D:\";
                saveFileDialog.FileName = fileName + ".xlsx";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    FileInfo fileInfo = new FileInfo(saveFileDialog.FileName);
                    excelPackage.SaveAs(fileInfo);
                }

                excelPackage.Dispose();
            }
        }

        private void dataGridview_ExportExcel(string fileName, DataGridView dgv, DataGridView dgv2)
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Add(true);
            Excel._Worksheet worksheet = workbook.Worksheets.get_Item(1) as Excel._Worksheet;
            worksheet.Name = "SaveData";

            // �׸��� �信 �����Ͱ� ���� ��� ���� �˾� ���
            if (dgv.Rows.Count == 0 && dgv2.Rows.Count == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "����� �����Ͱ� �����ϴ�.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning, 100);
            }

            // ������ �׸��� �� 1 ����� 2 ������ ���ʴ�� �־���
            for (int datarows = 0; datarows < dgv.Rows.Count; datarows++)
            {
                worksheet.Rows.Cells[datarows] = dgv.Rows[datarows].Cells;
            }

            worksheet.Columns.AutoFit();

            string filetype = fileName.Split('.')[1];

            if (filetype == "xls" || filetype == "xlsx")
            {
                workbook.SaveAs(fileName, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }

            workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            app.Quit();


            ReleaseExcelObject(app);
            ReleaseExcelObject(workbook);
            ReleaseExcelObject(worksheet);

        }

        // Marshal.ReleaseComObject ���� �Լ�
        static void ReleaseExcelObject(Object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }

            }
            catch (Exception e)
            {
                obj = null;
                throw e;
            }
            finally
            {
                GC.Collect();
            }

        }

        private void metroProgressBar1_Click(object sender, EventArgs e)
        {
            this.metroProgressBar1.Minimum = 0;
            this.metroProgressBar1.Maximum = 100;
            this.metroProgressBar1.Step = 1;
            this.metroProgressBar1.Value = 0;

        }


        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            //nt max = (int)e.Argument;

            for (int i = 0; i < 101; i++)
            {

                //nt progressPercentage = Convert.ToInt32(((double)i / max) * 100);
                if (backgroundWorker1.CancellationPending)
                {
                    e.Cancel = true;
                    break;
                }

                backgroundWorker1.ReportProgress(i);
                //System.Threading.Thread.Sleep(2);


            }

            e.Result = 0;
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            metroProgressBar1.Value = e.ProgressPercentage;

        }

        // nulltext�� ����ϰ� ����
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (this.textBox1.Text != "")
            {
                this.label5.Visible = false;
            }
            else
            {
                this.label5.Visible = true;
            }
        }

    }
}