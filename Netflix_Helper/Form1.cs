using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace Netflix_Helper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            prepare_comboBox();
        }

        private string filePath = "";

        public void prepare_comboBox()
        {
            // item 추가하기
            comboBox1.Items.Add("SF");
            comboBox1.Items.Add("드라마");
            comboBox1.Items.Add("스릴러");
            comboBox1.Items.Add("액션");
            comboBox1.Items.Add("범죄");
            comboBox1.Items.Add("코미디");
            comboBox1.Items.Add("다큐멘터리");
            comboBox1.Items.Add("판타지");
            comboBox1.Items.Add("스릴러");
            comboBox1.Items.Add("음악");
            comboBox1.Items.Add("스포츠");
            comboBox1.Items.Add("서부");
            comboBox1.Items.Add("애니메이션");
            comboBox1.Items.Add("역사");
            comboBox1.Items.Add("가족");
            comboBox1.Items.Add("전쟁");
            comboBox1.Items.Add("Reality TV");
            comboBox1.Items.Add("Made in Europe");
        }

        // 제목
        private void label1_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            richTextBox3.Clear();
            if (filePath != "")
            {
                Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = application.Workbooks.Open(Filename: @filePath);
                Worksheet worksheet1 = workbook.Worksheets.get_Item("Sheet1");
                application.Visible = false;
                Range range = worksheet1.UsedRange;
                String data_title = "";
                String data_genre = "";
                int num = 0;

                for (int i = 1; i <= range.Rows.Count; ++i)
                {
                    if (i == 1)
                    {
                        data_title += ((range.Cells[i, 1] as Range).Value2.ToString());
                    }
                    else
                    {
                        data_title += (num + ". " + (range.Cells[i, 1] as Range).Value2.ToString());
                    }
                    for (int j = 2; j <= range.Columns.Count; ++j)
                    {
                        data_genre += ((range.Cells[i, j] as Range).Value2.ToString() + " ");
                    }
                    data_title += "\n";
                    data_title += "--------------------\n";
                    data_genre += "\n";
                    data_genre += "-------------------------------------------------\n";
                    num++;
                }

                richTextBox1.Text = data_title;
                richTextBox3.Text = data_genre;

                DeleteObject(worksheet1);
                DeleteObject(workbook);
                application.Quit();
                DeleteObject(application);
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {

        }

        // 일반 검색
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            // textBox1.Text
        }

        // 장르 검색
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        // 제목 검색 버튼
        private void button1_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            richTextBox3.Clear();
            if (filePath != "")
            {
                Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = application.Workbooks.Open(Filename: @filePath);
                Worksheet worksheet1 = workbook.Worksheets.get_Item("Sheet1");
                application.Visible = false;
                Range range = worksheet1.UsedRange;
                String data_title = "";
                String data_genre = "";
                int num = 1;

                for (int i = 1; i <= range.Rows.Count; ++i)
                {
                    if (i == 1)
                    {
                        data_title += ((range.Cells[i, 1] as Range).Value2.ToString());
                        data_title += "\n";
                        data_title += "--------------------\n";
                        data_genre += ((range.Cells[i, 2] as Range).Value2.ToString());
                        data_genre += "\n";
                        data_genre += "-------------------------------------------------\n";
                    }
                    else
                    {
                        if (textBox1.Text == (range.Cells[i, 1] as Range).Value2.ToString())
                        {
                            data_title += (num + ". " + (range.Cells[i, 1] as Range).Value2.ToString());
                            for (int j = 2; j <= range.Columns.Count; ++j)
                            {
                                data_genre += ((range.Cells[i, j] as Range).Value2.ToString());
                                data_genre += "\n";
                                data_genre += "-------------------------------------------------\n";
                            }
                            data_title += "\n";
                            data_title += "--------------------\n";
                            num++;
                        }
                    }
                }

                richTextBox1.Text = data_title;
                richTextBox3.Text = data_genre;

                DeleteObject(worksheet1);
                DeleteObject(workbook);
                application.Quit();
                DeleteObject(application);
            }
        }

        // 엑셀 파일 위치 검색
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog OFD = new OpenFileDialog();

            if (OFD.ShowDialog() == DialogResult.OK) {
                richTextBox2.Clear();
                richTextBox2.Text = OFD.FileName;
                filePath = OFD.FileName;
            }
        }

        private void DeleteObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("메모리 할당을 해제하는 중 문제가 발생하였습니다." + ex.ToString(), "경고!");
            }
            finally
            {
                GC.Collect();
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        // 엑셀 내 데이터 읽기
        private void button3_Click(object sender, EventArgs e)
        {
            if (filePath != "") {
                Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = application.Workbooks.Open(Filename: @filePath);
                Worksheet worksheet1 = workbook.Worksheets.get_Item("Sheet1");
                application.Visible = false;
                Range range = worksheet1.UsedRange;
                String data_title = "";
                String data_genre = "";
                int num = 0;
                
                for (int i = 1; i <= range.Rows.Count; ++i) {
                    if( i == 1)
                    {
                        data_title += ((range.Cells[i, 1] as Range).Value2.ToString());
                    }
                    else
                    {
                        data_title += (num + ". " + (range.Cells[i, 1] as Range).Value2.ToString());
                    }
                    for (int j = 2; j <= range.Columns.Count; ++j) {
                        data_genre += ((range.Cells[i, j] as Range).Value2.ToString() + " ");
                    } 
                    data_title += "\n";
                    data_title += "--------------------\n";
                    data_genre += "\n";
                    data_genre += "-------------------------------------------------\n";
                    num++;
                }

                richTextBox1.Text = data_title;
                richTextBox3.Text = data_genre;

                DeleteObject(worksheet1);
                DeleteObject(workbook);
                application.Quit();
                DeleteObject(application);
            }
        }

        // 장르 출력
        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {

        }

        // 장르 검색 버튼
        private void button4_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            richTextBox3.Clear();
            if (filePath != "")
            {
                Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = application.Workbooks.Open(Filename: @filePath);
                Worksheet worksheet1 = workbook.Worksheets.get_Item("Sheet1");
                application.Visible = false;
                Range range = worksheet1.UsedRange;
                String data_title = "";
                String data_genre = "";
                int num = 0;

                for (int i = 1; i <= range.Rows.Count; ++i)
                {
                    if (i == 1)
                    {
                        data_title += ((range.Cells[i, 1] as Range).Value2.ToString());
                        data_title += "\n";
                        data_title += "--------------------\n";
                        data_genre += ((range.Cells[i, 2] as Range).Value2.ToString());
                        data_genre += "\n";
                        data_genre += "-------------------------------------------------\n";
                    }
                    else
                    {
                        if (comboBox1.SelectedItem.ToString() == (range.Cells[i, 2] as Range).Value2.ToString())
                        {
                            data_title += (num + ". " + (range.Cells[i, 1] as Range).Value2.ToString());
                            for (int j = 2; j <= range.Columns.Count; ++j)
                            {
                                data_genre += ((range.Cells[i, j] as Range).Value2.ToString());
                                data_genre += "\n";
                                data_genre += "-------------------------------------------------\n";
                            }
                            data_title += "\n";
                            data_title += "--------------------\n";
                            num++;
                        }
                    }
                }

                richTextBox1.Text = data_title;
                richTextBox3.Text = data_genre;

                DeleteObject(worksheet1);
                DeleteObject(workbook);
                application.Quit();
                DeleteObject(application);
            }
        }
    }
}
