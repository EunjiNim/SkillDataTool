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


        // 이미 데이터가 들어와 있는지 판별할 때 사용
        private string Skill = string.Empty;
        private string SkillEffect = string.Empty;
        private string SkillEffectLevelGroup = string.Empty;

        private int Level_Index = 0;

        // 받아올 엑셀 시트 이름
        private string Sheet_Name = "Table$";

        // 엑셀에서 오픈한 데이터를 딕셔너리 구조로 집어넣음. Value는 오브젝트 배열 형식으로 지정, 검색 시 각각의 컬럼 값에 접근하기 쉽도록 하기 위함
        private Dictionary<string, object[]> SkillData = new Dictionary<string, object[]>();
        private Dictionary<string, object[]> SkillEffectData = new Dictionary<string, object[]>();

        // SkillEffectLevelGroup은 동일한 키 값을 가지므로 다중값 딕셔너리를 구현하기 위해 Value를 List 형식으로 넣어줌
        private Dictionary<string, List<object[]>> SkillEffectLevelGroupData = new Dictionary<string, List<object[]>>();

        // 필요한 데이터만 따로 모아 그리드 뷰에 띄워주기 위해 사용
        private DataTable GridViewInData = new DataTable();


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        // 엑셀 파일 오픈 버튼 
        // 중복되는 데이터 문제로 엑셀 데이터를 여는 것과 데이터 바인딩을 분리시킴
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                // 엑셀 파일만 받을 수 있도록 함
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

                            // 불러온 엑셀 문서 이름을 띄워줌
                            this.listBox1.Items.Add(listbox);

                            // 엑셀 문서에 따라 각기 다르게 주소를 저장해 줌
                            if (str.Contains("Skill.xlsx"))
                            {
                                if (this.Skill.Length != 0)
                                {
                                    Console.WriteLine("이미 생성된 데이터가 있습니다.");
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
                                    Console.WriteLine("이미 생성된 데이터가 있습니다.");

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
                                    Console.WriteLine("이미 생성된 데이터가 있습니다.");

                                }
                                else
                                {
                                    this.SkillEffectLevelGroup = string.Format(this.Excel07Constring, str, 0);
                                }
                            }
                            else
                            {
                                MessageBox.Show("사용할 수 없는 문서입니다. 문서를 다시 확인해 주세요.");
                                Application.Restart();
                            }


                        }
                    }

                }
                catch (Exception ex)
                {
                    // 이미 불러온 데이터가 있을 경우 강제로 재시작해줌
                    Exception exception = ex;
                    MessageBox.Show(ex.InnerException != null ? ex.InnerException.Message : "이미 저장된 데이터가 있어 프로그램을 다시 시작합니다.");
                    Application.Restart();

                }
            }

        }

        // 데이터 바인딩하기
        private void button2_Click(object sender, EventArgs e)
        {
            // 예외처리 
            if (this.Skill.Length == 0)
            {
                // 엑셀 시트 경로가 들어가 있지 않음
                MessageBox.Show("데이터가 존재하지 않습니다.");
                return;
            }
            if (this.SkillData.Values.Count != 0)
            {
                // 이미 처리된 데이터일 경우
                MessageBox.Show("이미 로드가 완료된 데이터입니다.");
                return;
            }

            using (OleDbConnection conn = new OleDbConnection(this.Skill))
            {
                using (OleDbCommand comm = new OleDbCommand())
                {
                    using (OleDbDataAdapter adap = new OleDbDataAdapter())
                    {
                        // Skill 데이터 담기
                        DataTable datatable1 = new DataTable();
                        comm.CommandText = string.Concat("SELECT * From [", this.Sheet_Name, "]");
                        comm.Connection = conn;
                        conn.Open();
                        adap.SelectCommand = comm;
                        adap.Fill(datatable1);

                        foreach (DataRow row in datatable1.Rows)
                        {
                            // 데이터 테이블 상단에 사용하지 않는 데이터가 7줄 정도 기입되어 있음
                            // 데이터 ID를 키 값으로 사용해야 하는데 동일한 데이터 값이 있어 실 사용 데이터 부분부터만 데이터를 집어넣을 수 있도록 함
                            if (row.Table.Rows.IndexOf(row) > 7)
                            {
                                if (!SkillData.ContainsKey(row.ItemArray[0].ToString()))
                                {
                                    // 0번의 ID를 키값으로 기준을 잡고 해당 array를 모두 저장함. 추후 ID를 인덱스로 하여 검색하기 위함
                                    SkillData.Add(row.ItemArray[0].ToString(), row.ItemArray);
                                }
                                else
                                {
                                    // 중복되는 값이 생겼을 경우 문제가 생기므로 확인을 위해 노티해 줌
                                    MessageBox.Show("중복되는 키 값이 있습니다. " + row.ItemArray[0].ToString() + " 데이터를 확인해 주세요.");
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
                        // SkillEffect 데이터 담기
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
                                    MessageBox.Show("중복되는 키 값이 있습니다. " + row.ItemArray[0].ToString() + " 데이터를 확인해 주세요.");
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
                        // SkillEffectLevelGroup 데이터 담기
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
                                // SkillEffectLevelGroupData는 ID 값을 동일한 그룹으로 사용하므로 해당 키 값이 생성되어 있을 경우 리스트 형식으로 값을 넣어줌
                                // Dictonary는 키 값을 중복으로 사용할 수 없기 때문에 Value를 리스트 형식으로 구성하여 사용할 수 있도록 편법을 사용함
                                if (SkillEffectLevelGroupData.ContainsKey(row.ItemArray[0].ToString()))
                                {
                                    SkillEffectLevelGroupData[row.ItemArray[0].ToString()].Add(row.ItemArray);
                                }
                                else
                                {
                                    // 키 값이 없을 경우 새롭게 리스트를 구성함
                                    SkillEffectLevelGroupData[row.ItemArray[0].ToString()] = new List<object[]> { row.ItemArray };
                                }
                            }
                            conn.Close();
                        }
                    }
                }
            }
        }


        // 인덱스 검색 기능
        private void button3_Click(object sender, EventArgs e)
        {
            // 컬럼 위치가 변경될 수 있으므로 리스트에 넣어 이름을 찾고 그 컬럼의 인덱스를 반환할 수 있도록 함
            /*object[] ItemDataIndex = new object[1];
            SkillData.TryGetValue("skill_id", out ItemDataIndex);*/


            // 10 << 스킬 이펙트 ID 이므로 해당 인덱스 정보를 찾아들어가야 함

            // 100100111

            // SkillEffectLevelGroupData는 찾고자 하는 레벨을 알아야 데이터를 가지고 올 수가 있음
            List<object[]> ItemDataIndex = new List<object[]>();
            SkillEffectLevelGroupData.TryGetValue("skill_effect_level_group_id", out ItemDataIndex);

            List<object[]> Search_SkillEffectLevelData = new List<object[]>();
            SkillEffectLevelGroupData.TryGetValue("100100111", out Search_SkillEffectLevelData);




            // SkillEffectLevelGroup의 컬럼을 순차적으로 그리드 뷰에 넣어줌
            for (int i = 0; i < ItemDataIndex[0].ToArray().Length; ++i)
            {
                GridViewInData.Columns.Add(ItemDataIndex[0].ToArray()[i].ToString());
            }

            List<string> Level_List = new List<string>();

            // SkillEffectLevelGroup의 레벨 리스트를 콤보박스에 넣어줌
            foreach (var item in Search_SkillEffectLevelData)
            {
                Level_List.Add(item[4].ToString());
            }
            this.comboBox1.Items.AddRange(Level_List.ToArray());



            // 검색한 인덱스의 row 데이터를 모두 그리드 뷰에 넣어줌
            //GridViewInData.Rows.Add(testdata[1]);
            foreach (var test in Search_SkillEffectLevelData)
            {
                GridViewInData.Rows.Add(test);
            }


            // 데이터 그리드 뷰에 모아둔 데이터를 띄워 줌
            this.dataGridView1.DataSource = GridViewInData;

        }

        // 콤보박스 테스트
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int Skill_Level = 0;
            
            // 콤보박스의 string 형식을 int로 변환해 주고 인덱스를 맞춰주기 위해 1을 뺌
            Skill_Level = Convert.ToInt16(comboBox1.Text) - 1;      

            // 리스트 초기화
            GridViewInData.Clear();

            List<object[]> Search_SkillEffectLevelData = new List<object[]>();
            SkillEffectLevelGroupData.TryGetValue("100100111", out Search_SkillEffectLevelData);

            // 특정 레벨의 정보만 출력해 줌
            GridViewInData.Rows.Add(Search_SkillEffectLevelData[Skill_Level]);
            this.dataGridView1.DataSource = GridViewInData;

        }
    }
}