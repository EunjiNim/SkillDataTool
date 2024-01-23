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


        // 이미 데이터가 들어와 있는지 판별할 때 사용
        private string Skill = string.Empty;
        private string SkillEffect = string.Empty;
        private string SkillEffectLevelGroup = string.Empty;
        private string SkillEffectOperation = string.Empty;

        // 사용자가 입력한 인덱스를 저장
        private string Index_Num = string.Empty;

        // 사용자가 입력한 인덱스를 바탕으로 연결된 인덱스를 받아옴
        private string SkillEffect_Num = string.Empty;
        private string SkillEffectOperation_Num = string.Empty;

        private int Level_Index = 0;
        private int AllFile_Count = 0;

        // 받아올 엑셀 시트 이름
        private string Sheet_Name = "Table$";

        // 엑셀에서 오픈한 데이터를 딕셔너리 구조로 집어넣음. Value는 오브젝트 배열 형식으로 지정, 검색 시 각각의 컬럼 값에 접근하기 쉽도록 하기 위함
        private Dictionary<string, object[]>? SkillData = new Dictionary<string, object[]>();
        private Dictionary<string, object[]>? SkillEffectData = new Dictionary<string, object[]>();
        private Dictionary<string, object[]>? SkillEffectOperationData = new Dictionary<string, object[]>();

        // SkillEffectLevelGroup은 동일한 키 값을 가지므로 다중값 딕셔너리를 구현하기 위해 Value를 List 형식으로 넣어줌
        private Dictionary<string, List<object[]>>? SkillEffectLevelGroupData = new Dictionary<string, List<object[]>>();

        // 필요한 데이터만 따로 모아 그리드 뷰에 띄워주기 위해 사용
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
                                    //Workbook wb = new Workbook(this.Skill.ToString());
                                }
                            }
                            else if (str.Contains("SkillEffectGroup.xlsx"))
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
                            else if (str.Contains("SkillEffectOperation.xlsx"))
                            {
                                if (this.SkillEffectOperation.Length != 0)
                                {
                                    Console.WriteLine("이미 생성된 데이터가 있습니다.");

                                }
                                else
                                {
                                    this.SkillEffectOperation = string.Format(this.Excel07Constring, str, 0);
                                }
                            }
                            else
                            {
                                MetroFramework.MetroMessageBox.Show(this, "사용할 수 없는 문서입니다. 문서를 다시 확인해 주세요.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning, 100);
                            }
                        }
                    }

                }
                catch (Exception ex)
                {
                    // 이미 불러온 데이터가 있을 경우 강제로 재시작해줌
                    Exception exception = ex;
                    MessageBox.Show(ex.InnerException != null ? ex.InnerException.Message : "문제가 발생하여 프로그램을 다시 시작합니다.");

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
                this.listBox1.Items.Clear();
                return;
            }
            if (this.SkillData.Values.Count != 0)
            {
                // 이미 처리된 데이터일 경우
                MessageBox.Show("이미 로드가 완료된 데이터입니다.");
                return;
            }

            int z = 0;


            // backgroundWorker를 이용해서 progressbar를 실행시켜줌
            backgroundWorker1.RunWorkerAsync();

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
                        // SkillEffectLevelGroup 데이터 담기
                        DataTable dataTable3 = new DataTable();
                        comm.CommandText = string.Concat("SELECT * From [", this.Sheet_Name, "]");
                        comm.Connection = conn;
                        conn.Open();
                        adap.SelectCommand = comm;
                        adap.Fill(dataTable3);


                        // 쓸모없는 마지막 컬럼 아예 삭제. 데이터를 넣으면서 실수를 많이하여 빈 컬럼이 자꾸 생김
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
                        // SkillEffectOperation 데이터 담기
                        DataTable dataTable4 = new DataTable();
                        comm.CommandText = string.Concat("SELECT * From [", this.Sheet_Name, "]");
                        comm.Connection = conn;
                        conn.Open();
                        adap.SelectCommand = comm;
                        adap.Fill(dataTable4);

                        // 쓸모없는 마지막 컬럼 아예 삭제. 데이터를 넣으면서 실수를 많이하여 빈 컬럼이 자꾸 생김
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
                                    MessageBox.Show("중복되는 키 값이 있습니다. " + row.ItemArray[0].ToString() + " 데이터를 확인해 주세요.");
                                }
                            }
                        }

                        comm.Dispose();
                        conn.Close();
                    }
                }
            }

            // 데이터 로딩이 완료되면 텍스트를 변경해 줌
            if (SkillData.Count != null && SkillEffectData.Count != null && SkillEffectLevelGroupData.Count != null && SkillEffectOperationData.Count != null)
            {
                this.button2.Text = "Complate";
                backgroundWorker1.CancelAsync();

            }

        }


        // 인덱스 검색 기능
        private void button3_Click(object sender, EventArgs e)
        {
            // 변수에 텍스트 박스에 입력한 수치를 넣어줌
            this.Index_Num = this.textBox1.Text;

            // 인덱스 넘버로 Skill Data를 찾아옴
            object[]? SkillDatasRow = new object[1];
            SkillData.TryGetValue(Index_Num, out SkillDatasRow);

            // 검색한 키가 없는 경우에는 텍스트 박스를 지우고, 에러 팝업을 띄워줌
            if (!SkillData.ContainsKey(Index_Num))
            {
                MessageBox.Show("존재하지 않는 인덱스입니다.");
                this.textBox1.Clear();
                return;
            }

            // 재검색시마다 리스트가 쌓이므로 버튼을 클릭할때마다 초기화 해 줌
            this.comboBox1.Items.Clear();
            GridViewInData.Clear();
            GridViewOperationData.Clear();
            // 콤보박스 텍스트 리셋
            this.comboBox1.ResetText();

            // 컬럼 위치가 변경될 수 있으므로 리스트에 넣어 이름을 찾고 그 컬럼의 인덱스를 반환할 수 있도록 함
            object[]? SkillDataIndex = new object[1];
            SkillData.TryGetValue("skill_id", out SkillDataIndex);

            // 인덱스에 맞는 스킬 정보를 찾아줌
            object[]? Search_SkillData = new object[1];
            SkillData.TryGetValue(Index_Num, out Search_SkillData);

            // 스킬 데이터에서 스킬 이펙트 넘버 받아옴
            SkillEffect_Num = SkillDatasRow.ToArray()[SkillDataIndex.ToList().IndexOf("link_skill_effect_id")].ToString();

            // 스킬 이펙트 데이터 인덱스로 오퍼레이션 인덱스를 찾아내기 위해 SkillEffectGroup 시트 참고
            object[]? search_SkillEffectData = new object[1];
            SkillEffectData.TryGetValue(SkillEffect_Num, out search_SkillEffectData);

            if (search_SkillEffectData != null)
            {
                // Skill Effect Operation 데이터 뽑아냄
                SkillEffectOperation_Num = search_SkillEffectData.ToArray()[SkillEffectData.ToArray()[1].Value.ToList().IndexOf("link_skill_effect_operation_id")].ToString();

            }

            // 인덱스를 뽑기위해 리스트에 넣어줌
            List<string>? SkillIndexList = new List<string>();
            foreach (Object list in SkillDataIndex)
            {
                SkillIndexList.Add(list.ToString());
            }

            // SkillEffectLevelGroupData는 찾고자 하는 레벨을 알아야 데이터를 가지고 올 수가 있음
            List<object[]>? ItemDataIndex = new List<object[]>();
            SkillEffectLevelGroupData.TryGetValue("스킬 효과 레벨 그룹", out ItemDataIndex);


            // skill 시트에서 찾아낸 스킬 이펙트의 인덱스를 뽑아내고 이 인덱스로 SkillEffectLevelGroup 시트에서 데이터를 찾음
            List<object[]>? Search_SkillEffectLevelData = new List<object[]>();
            SkillEffectLevelGroupData.TryGetValue(SkillEffect_Num, out Search_SkillEffectLevelData);


            // 그리드 뷰에 컬럼 값이 없을때만 내용을 넣어줌
            if (GridViewInData.Columns.Count == 0)
            {
                // SkillEffectLevelGroup의 컬럼을 순차적으로 그리드 뷰에 넣어줌
                for (int i = 0; i < ItemDataIndex[0].ToArray().Length; ++i)
                {
                    GridViewInData.Columns.Add(ItemDataIndex[0].ToArray()[i].ToString());
                    ConvertGridViewInData.Columns.Add(ItemDataIndex[0].ToArray()[i].ToString());
                }
            }

            // Skill Effect Operation 정보를 데이터 그리드 뷰에 넣어주기 위해 인덱스 검색
            object[]? SkillEffectOperation_Result = new object[1];
            SkillEffectOperationData.TryGetValue(SkillEffectOperation_Num, out SkillEffectOperation_Result);

            object[]? SkillEffectOperation_Column = new object[1];
            SkillEffectOperationData.TryGetValue("스킬 효과 작동 아이디", out SkillEffectOperation_Column);

            // 몬스터 스킬은 operation 정보에 없으므로 검색한 후 값이 있을때에만 그리드 뷰에 넣어 줌 
            if (SkillEffectOperation_Result != null)
            {
                this.dataGridView2.Show();
                // 그리드 뷰에 컬럼이 없을때만 값을 넣어줌
                if (GridViewOperationData.Columns.Count == 0)
                {
                    for (int i = 0; i < SkillEffectOperation_Column.Length; ++i)
                    {
                        GridViewOperationData.Columns.Add(SkillEffectOperation_Column.ToArray()[i].ToString());
                    }
                }
                // 검색된 오퍼레이션 정보를 넣어줌
                GridViewOperationData.Rows.Add(SkillEffectOperation_Result);
                this.dataGridView2.DataSource = GridViewOperationData;
            }
            else
            {
                // 데이터가 없으면 그리드 뷰를 아예 감춰버림
                this.dataGridView2.Hide();
            }

            List<string>? Level_List = new List<string>();

            // 레벨 단계가 없을수도 있으므로 예외처리 함
            if (Search_SkillEffectLevelData == null)
            {
                this.dataGridView1.Hide();
            }
            else
            {
                foreach (var item in Search_SkillEffectLevelData)
                {
                    Level_List.Add(item[ItemDataIndex[0].ToList().IndexOf("레벨")].ToString());
                }

                // 콤보 박스에 레벨 추출한 레벨 리스트를 넣어줌
                this.comboBox1.Items.AddRange(Level_List.ToArray());

                // 검색한 인덱스의 row 데이터를 모두 그리드 뷰에 넣어줌
                foreach (var SkillEffectResualtData in Search_SkillEffectLevelData)
                {
                    GridViewInData.Rows.Add(SkillEffectResualtData);
                }

                this.dataGridView1.Show();
            }

            // 레벨 단계가 없을수도 있으므로 예외처리 함
            /*if (Search_SkillEffectLevelData.Count != 0)
             {
                 // SkillEffectLevelGroup의 레벨 리스트를 콤보박스에 넣어줌
                 foreach (var item in Search_SkillEffectLevelData)
                 {
                     Level_List.Add(item[ItemDataIndex[0].ToList().IndexOf("레벨")].ToString());
                 }

             }*/


            // 콤보 박스에 레벨 추출한 레벨 리스트를 넣어줌
            //this.comboBox1.Items.AddRange(Level_List.ToArray());

            // 검색한 인덱스의 row 데이터를 모두 그리드 뷰에 넣어줌
            /*foreach (var SkillEffectResualtData in Search_SkillEffectLevelData)
            {
                GridViewInData.Rows.Add(SkillEffectResualtData);
            }*/

            // 텍스트 박스에 Skill.xlsx 의 내용을 띄워줌 
            // 쓸데없는 변수 할당은 줄이고 리스트에서 바로 인덱스를 뽑아서 넣어줌. 컬럼 명으로 인덱스를 뽑으면 추후 컬럼 위치가 변경되어도 원하는 값을 가져올 수 있기 때문
            this.textBox2.Text = Search_SkillData[SkillIndexList.IndexOf("skill_name")].ToString();
            this.textBox3.Text = Search_SkillData[SkillIndexList.IndexOf("skill_cooltime")].ToString();
            this.textBox4.Text = Search_SkillData[SkillIndexList.IndexOf("combo_attribute_element_count")].ToString();
            this.textBox5.Text = Search_SkillData[SkillIndexList.IndexOf("target")].ToString();
            this.textBox6.Text = Search_SkillData[SkillIndexList.IndexOf("skill_show_order")].ToString();


            // 데이터 그리드 뷰에 모아둔 데이터를 띄워 줌
            this.dataGridView1.DataSource = GridViewInData;
            this.dataGridView1.AllowUserToOrderColumns = false;

            // 쓸데없이 나오는 마지막 ROW를 감춤
            this.dataGridView2.AllowUserToAddRows = false;

        }

        // Skill Effect Leve Group 레벨 별 정보 출력
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int Skill_Level = 0;

            // 콤보박스의 string 형식을 int로 변환해 주고 인덱스를 맞춰주기 위해 1을 뺌
            Skill_Level = Convert.ToInt16(comboBox1.Text) - 1;

            // 리스트 초기화
            GridViewInData.Clear();

            List<object[]> Search_SkillEffectLevelData = new List<object[]>();
            SkillEffectLevelGroupData.TryGetValue(SkillEffect_Num, out Search_SkillEffectLevelData);

            // 특정 레벨의 정보만 출력해 줌
            GridViewInData.Rows.Add(Search_SkillEffectLevelData[Skill_Level]);
            this.dataGridView1.DataSource = GridViewInData;
        }

        // 파일 저장 버튼
        private void metroButton1_Click(object sender, EventArgs e)
        {
            /*string filePath = string.Empty;
            //string fileName = @"엑셀파일저장"  + DateHe

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "저장 경로를 지정하세요.";
            saveFileDialog.OverwritePrompt = true;
            saveFileDialog.Filter = "Excel 통합 문서|*.xlsx";
            saveFileDialog.InitialDirectory = @"D:\";
            
            if(saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = saveFileDialog.FileName;
                // 데이터테이블에 있는 데이터를 엑셀에다 넣어주고 파일로 저장해줌
 

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


                // Skill Effect Level Group 데이터를 넣어줌
                worksheet.Cells["A1"].LoadFromDataTable(GridViewInData, true, OfficeOpenXml.Table.TableStyles.Light8);

                // Skill Effect Operation  데이터를 넣어줌
                string index = "A" + (GridViewInData.Rows.Count + 3).ToString();
                worksheet.Cells[index].LoadFromDataTable(GridViewOperationData, true, OfficeOpenXml.Table.TableStyles.Light8);

                // 표 선만들기
                worksheet.Columns.AutoFit();


                // Title 구분 선
                /*worksheet.Cells[1, 1, 1, GridViewInData.Columns.Count].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[1, 1, 1, GridViewInData.Columns.Count].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Rows[1].Range.Style.Font.Bold = true;
                // Right 선
                worksheet.Cells[1, GridViewInData.Columns.Count, GridViewInData.Rows.Count +1, GridViewInData.Columns.Count].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                // Bottom 선
                worksheet.Cells[GridViewInData.Rows.Count +1, 1, GridViewInData.Rows.Count +1, GridViewInData.Columns.Count].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                // Left 선
                worksheet.Cells[1, 1, GridViewInData.Rows.Count + 1, GridViewInData.Columns.Count].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;*/

                // Save File Dialog 기본 세팅
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Title = "저장 경로를 지정하세요.";
                saveFileDialog.OverwritePrompt = true;
                saveFileDialog.Filter = "Excel 통합 문서|*.xlsx";
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

            // 그리드 뷰에 데이터가 없을 경우 에러 팝업 출력
            if (dgv.Rows.Count == 0 && dgv2.Rows.Count == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "출력할 데이터가 없습니다.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning, 100);
            }

            // 데이터 그리드 뷰 1 내용과 2 내용을 차례대로 넣어줌
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

        // Marshal.ReleaseComObject 실행 함수
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

        // nulltext를 비슷하게 구현
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