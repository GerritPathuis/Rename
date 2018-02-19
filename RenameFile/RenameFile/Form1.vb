Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.ColumnCount = 4
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.AllCells
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Clear()
        OpenFileDialog1.Title = "Please Select a File"
        OpenFileDialog1.InitialDirectory = TextBox1.Text
        OpenFileDialog1.FileName = TextBox1.Text
        OpenFileDialog1.ShowDialog()

        TextBox1.Text = OpenFileDialog1.FileName
        TextBox7.Clear()
        '==============  Read from file into the dataview grid ==============
        'Read_from_file()
        Read_excel()
    End Sub

    Private Sub Read_excel()
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        If File.Exists(TextBox1.Text) = True Then
            Try
                xlApp = New Excel.ApplicationClass
                xlWorkBook = xlApp.Workbooks.Open(TextBox1.Text)
                xlWorkSheet = xlWorkBook.Worksheets("Referenties")

                'Read the excel file
                ProgressBar1.Visible = True
                For row = 1 To 100
                    ProgressBar1.Value = 100 - row

                    'DataGridView1.Rows.Item(row).Cells(1).Value = xlWorkSheet.Cells(row, 2).ToString
                    DataGridView1.Rows.Add(New String() {xlWorkSheet.Cells(row, 1).value, xlWorkSheet.Cells(row, 2).value, xlWorkSheet.Cells(row, 3).value, xlWorkSheet.Cells(row, 4).value})
                Next
                ProgressBar1.Visible = False

                'MsgBox(xlWorkSheet.Cells(2, 2).value)
                'edit the cell with new value

                ' xlWorkSheet.Cells(2, 2) = "http://vb.net-informations.com"
                xlWorkBook.Close()
                xlApp.Quit()

                ReleaseObject(xlApp)
                ReleaseObject(xlWorkBook)
                ReleaseObject(xlWorkSheet)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                '---------------- now convert-----------------
            End Try
        Else
            MessageBox.Show("Tough shit, file does not exist ..")
        End If
    End Sub

    Private Sub Read_from_file()
        Dim coll, row_no As Integer

        DataGridView1.Rows.Clear()
        If File.Exists(TextBox1.Text) = True Then
            Try
                ProgressBar1.Value = 100
                row_no = -1
                For Each row As String In File.ReadAllLines(TextBox1.Text)
                    '-----------------------------------------------------
                    If ProgressBar1.Value > ProgressBar1.Minimum Then
                        ProgressBar1.Value -= 1
                    Else
                        ProgressBar1.Value = ProgressBar1.Maximum
                    End If
                    '-----------------------------------------------------

                    TextBox7.AppendText(row.ToString)

                    DataGridView1.Rows.Add()
                    row_no += 1
                    coll = 0
                    For Each column As String In row.Split(New String() {";"}, StringSplitOptions.None)
                        DataGridView1.Rows.Item(row_no).Cells(coll).Value = column.ToString
                        coll += 1
                    Next
                Next
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                '---------------- now convert-----------------
            End Try

            TabControl1.SelectedIndex = 1
            TextBox7.Clear()
        Else
            MessageBox.Show("Tough shit, file does not exist ..")
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Read_excel()
    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Class
