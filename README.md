# 푸른황소의 혼 - 7조
--------------------------------------------------
## :notebook_with_decorative_cover: Netflix Helper 
크롤링 -> C#을 통해 csv파일 호출 및 읽기 -> 검색 및 장르선택 기능을 통해 검색 

## :notebook_with_decorative_cover: 구현결과
<img src="https://user-images.githubusercontent.com/81347125/143462681-6395d376-9dca-4cc9-9802-2a87d7ab6b26.gif" width="60%">  



<br>

## :notebook_with_decorative_cover:과제리뷰
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

