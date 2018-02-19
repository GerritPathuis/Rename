Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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
        Read_from_file()
    End Sub

    Private Sub Read_from_file()
        Dim coll, row_no As Integer

        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()

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


End Class
