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
        TextBox2.Text = Path.GetDirectoryName(TextBox1.Text)

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

                xlApp.DisplayAlerts = False 'Suppress excel messages

                'Read the excel file
                ProgressBar1.Visible = True
                For row = 1 To 100
                    ProgressBar1.Value = 100 - row
                    DataGridView1.Rows.Add(New String() {xlWorkSheet.Cells(row, 1).value, xlWorkSheet.Cells(row, 2).value, xlWorkSheet.Cells(row, 3).value, xlWorkSheet.Cells(row, 4).value})
                Next
                ProgressBar1.Visible = False

                'MsgBox(xlWorkSheet.Cells(2, 2).value)
                'edit the cell with new value

                ' xlWorkSheet.Cells(2, 2) = "http://vb.net-informations.com"
                xlWorkBook.Close()  'Er zijn toch geen wijzigingen
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Check_file_exist()
    End Sub
    Private Sub Check_file_exist()
        Dim exist_name As String

        ProgressBar1.Visible = True
        Refresh()

        For row = 1 To 100
            ProgressBar1.Value = 100 - row

            exist_name = TextBox2.Text & "\" & DataGridView1.Rows.Item(row).Cells(0).Value & ".idw"
            If File.Exists(exist_name) Then
                DataGridView1.Rows.Item(row).Cells(0).Style.BackColor = Color.Green
            Else
                DataGridView1.Rows.Item(row).Cells(0).Style.BackColor = Color.Red
            End If
        Next
        ProgressBar1.Visible = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Rename_files()
    End Sub
    Private Sub Rename_files()
        Dim exist_name, new_name As String

        ProgressBar1.Visible = True
        Refresh()

        For row = 1 To 100
            ProgressBar1.Value = 100 - row

            exist_name = TextBox2.Text & "\" & DataGridView1.Rows.Item(row).Cells(0).Value & ".idw"
            new_name = TextBox2.Text & "\" & DataGridView1.Rows.Item(row).Cells(3).Value & ".idw"

            If File.Exists(exist_name) Then
                My.Computer.FileSystem.RenameFile(exist_name, new_name)
            End If
        Next
        ProgressBar1.Visible = False
    End Sub
End Class
