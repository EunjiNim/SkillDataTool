using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace SkillDataTool
{
    public partial class Form1 : Form
    {
        private string Excel07Constring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=NO;IMEX=1;'";

        // �޾ƿ� ���� ��Ʈ
        private string Skill_Sheet = "Table$"; 

       


        // �������� ������ �����͸� ��ųʸ� ������ �������. Value�� ������Ʈ �迭 �������� ����, ������ �÷� ���� �����ϱ� ������ �ϱ� ����
        private Dictionary<string, object[]> SkillData = new Dictionary<string, object[]>();
        private Dictionary<string, object[]> SkillEffectData = new Dictionary<string, object[]>();


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}