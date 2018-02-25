Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form1
    Public _no_rowws As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.ColumnCount = 4
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.AllCells
        Label5.Text = ""
        Label6.Text = ""
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Clear()
        OpenFileDialog1.Title = "Please Select a File"
        OpenFileDialog1.InitialDirectory = TextBox1.Text
        OpenFileDialog1.FileName = TextBox1.Text
        OpenFileDialog1.ShowDialog()

        TextBox1.Text = OpenFileDialog1.FileName
        TextBox2.Text = Path.GetDirectoryName(TextBox1.Text)

        '==============  Read from file into the dataview grid ==============

        Read_excel()
        Label5.Text = ""
        Label6.Text = ""
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
                _no_rowws = xlWorkSheet.UsedRange.Rows.Count
                Button1.Text = "2) Read Excel File " & _no_rowws.ToString & " Elements"
                ProgressBar1.Maximum = _no_rowws
                ProgressBar1.Value = _no_rowws
                ProgressBar1.Visible = True
                Me.Refresh()

                DataGridView1.Rows.Clear()  'Make sure the grid is empty

                For row = 1 To _no_rowws
                    ProgressBar1.Value = _no_rowws - row
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
                MessageBox.Show("Read excel section " & ex.Message)
                '---------------- now convert-----------------
            End Try
        Else
            MessageBox.Show("Tough shit, file does not exist ..")
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Read_excel()
        DataGridView1.Sort(DataGridView1.Columns(3), ListSortDirection.Ascending)
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
        Dim exist_name, new_name As String
        Dim problem_counter1 As Integer
        Dim problem_counter2 As Integer

        ProgressBar1.Value = _no_rowws
        ProgressBar1.Visible = True
        Me.Refresh()

        For row = 1 To _no_rowws - 1
            ProgressBar1.Value = _no_rowws - row

            '--------------- Check exist name ----------------
            exist_name = TextBox2.Text & "\" & DataGridView1.Rows.Item(row).Cells(0).Value & ".idw"
            If File.Exists(exist_name) Then
                DataGridView1.Rows.Item(row).Cells(0).Style.BackColor = Color.White
            Else
                DataGridView1.Rows.Item(row).Cells(0).Style.BackColor = Color.Red
                problem_counter1 += 1
            End If

            '--------------- Check new name ----------------
            new_name = TextBox2.Text & "\" & DataGridView1.Rows.Item(row).Cells(3).Value & ".idw"
            If File.Exists(new_name) Then
                DataGridView1.Rows.Item(row).Cells(3).Style.BackColor = Color.Red
                problem_counter2 += 1
            Else
                DataGridView1.Rows.Item(row).Cells(3).Style.BackColor = Color.White
            End If

        Next
        ProgressBar1.Visible = False

        Label3.Text = "Problem " & problem_counter1.ToString & " Old IDW's Not found"
        Label4.Text = "Problem " & problem_counter2.ToString & " New IDW's already exist"
        Label3.Visible = IIf(problem_counter1 > 0, True, False)
        Label4.Visible = IIf(problem_counter2 > 0, True, False)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Rename_files()
    End Sub
    Private Sub Rename_files()
        Dim exist_name, new_name As String
        Dim succes_counter As Integer

        ProgressBar1.Value = _no_rowws
        ProgressBar1.Visible = True
        Me.Refresh()
        Application.DoEvents()

        Try
            For row = 1 To _no_rowws
                ProgressBar1.Value = _no_rowws - row

                exist_name = TextBox2.Text & "\" & DataGridView1.Rows.Item(row).Cells(0).Value & ".idw"
                new_name = DataGridView1.Rows.Item(row).Cells(3).Value & ".idw"

                'Conditions
                'Old file must exist
                'New file must be absent
                'Old file name length > 0
                'New file name length > 0
                If File.Exists(exist_name) And (Not File.Exists(TextBox2.Text & "\" & new_name)) And exist_name.Length > 0 And new_name.Length > 0 Then
                    My.Computer.FileSystem.RenameFile(exist_name, new_name)
                    succes_counter += 1
                End If
            Next
            ProgressBar1.Visible = False

        Catch ex As Exception
            MessageBox.Show("Renaming section" & ex.Message)
        End Try
        Label5.Visible = True
        If (succes_counter > 0) Then
            Label5.Text = succes_counter.ToString & " IDW files are renamed"
        Else
            Label5.Text = "None IDW files are renamed"
        End If
    End Sub
    Private Sub Identical_target_names()
        Dim target_name As String
        Dim target_name2 As String
        Dim double_file_cntr As Integer
        Dim rev_(26) As Char    'alphabet
        Dim a, b, c As Integer

        ProgressBar1.Value = _no_rowws
        ProgressBar1.Visible = True
        Me.Refresh()
        Application.DoEvents()

        ' Fill the array
        For i = 0 To 26
            rev_(i) = Chr(Asc("a") + i)
        Next

        Try
            For row = 1 To _no_rowws
                ProgressBar1.Value = _no_rowws - row
                target_name = DataGridView1.Rows.Item(row).Cells(3).Value
                double_file_cntr = -2
                For roww = 1 To _no_rowws
                    target_name2 = DataGridView1.Rows.Item(roww).Cells(3).Value
                    If String.Equals(target_name, target_name2) Then
                        double_file_cntr += 1

                        If (double_file_cntr > -1) Then
                            If double_file_cntr > 15645 Then MessageBox.Show("Error > 15624 identical files ")
                            a = Math.Floor(double_file_cntr / 625)
                            b = Math.Floor(double_file_cntr / 25)
                            c = double_file_cntr Mod 25 'Modulus

                            DataGridView1.Rows.Item(roww).Cells(3).Value = target_name2 & "_" & rev_(a) & rev_(b) & rev_(c)
                            DataGridView1.Rows.Item(roww).Cells(3).Style.BackColor = Color.Yellow
                        End If
                    End If

                Next
            Next
            ProgressBar1.Visible = False
            Label6.Text = "Done, see the datagrid look for yellow cells"
        Catch ex As Exception
            MessageBox.Show("Renaming section" & ex.Message)
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Identical_target_names()
    End Sub
End Class
