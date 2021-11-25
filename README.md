# 푸른황소의 혼 - 7조
--------------------------------------------------
## :notebook_with_decorative_cover: 프로젝트 내용
###  [Netflix Helper]
#### - justwatch(https://www.justwatch.com/kr) 를 통해 한국 동영상 서비스 목록에서 크롤링 시도
#### - CSV 데이터 파일을 C#과 연동
#### - C# 디자인을 통해 프로그램 기능 구현

<br>

## :notebook_with_decorative_cover: 구현결과
<img src="https://user-images.githubusercontent.com/81347125/143462681-6395d376-9dca-4cc9-9802-2a87d7ab6b26.gif" width="60%">  

<br>

## :notebook_with_decorative_cover:설치방법 및 사용 방법
#### 1. git clone 생성
#### 2. visual studio C#, sln 파일 실행
#### 3. textBox 및 ComboBox를 통해 검색,장르 기능 사용 

<br>

## :notebook_with_decorative_cover:과제 코드리뷰
### :pushpin: 엑셀 파일 위치 검색

<pre>
<code>
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog OFD = new OpenFileDialog();

            if (OFD.ShowDialog() == DialogResult.OK) {
                richTextBox2.Clear();
                richTextBox2.Text = OFD.FileName;
                filePath = OFD.FileName;
            }
        }

</code>
</pre>

### :pushpin: 엑셀 내 데이터 호출 및 읽기

<pre>
<code>
        private void button3_Click(object sender, EventArgs e)
        {
            if (filePath != "") {
                Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = application.Workbooks.Open(Filename: @filePath);
                Worksheet worksheet1 = workbook.Worksheets.get_Item("Sheet1");
                application.Visible = false;
                Range range = worksheet1.UsedRange;
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
</code>
</pre>

### :pushpin: 오류 처리

<pre>
<code>
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
</code>
</pre>



<br>
<hr>

