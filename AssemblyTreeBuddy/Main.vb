
Imports System.Data
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.MarshalByRefObject


Public Class frmMain


    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        On Error Resume Next

        If gTreeForm IsNot Nothing Then
            gTreeForm.Close()
        End If

        gMainForm = Nothing

        System.Windows.Forms.Application.Exit()

    End Sub


    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim i As Integer

        For i = 0 To 30
            gBitMask(i) = 2 ^ i
        Next

        gBitMask(31) = &H80000000

        gAppPath = GetSetting(System.Windows.Forms.Application.ProductName, "General", "Last Filepath", "")

        If Not System.IO.Directory.Exists(gAppPath) Then
            gAppPath = ""
        End If

    End Sub


    Private Sub cmdOpenFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOpenFile.Click

        Dim OpenFileDialog As New OpenFileDialog

        On Error GoTo ErrorHandler

        If gAppPath = "" Then
            OpenFileDialog.InitialDirectory = "c:\"
        Else
            OpenFileDialog.InitialDirectory = gAppPath
        End If

        OpenFileDialog.Filter = "Excel Files (*.xlsx)|*xlsx"
        OpenFileDialog.Title = "Select the Excel workbork for the assembly tree data."

        If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then

            gAppPath = System.IO.Path.GetDirectoryName(OpenFileDialog.FileName)
            SaveSetting(System.Windows.Forms.Application.ProductName, "General", "Last Filepath", gAppPath)

            gFilename = OpenFileDialog.FileName

            GetMachineWorkbookData(OpenFileDialog.FileName)

        End If

        Exit Sub

ErrorHandler:
        LogErrorMessage("sub cmdGetMachineData_Click", Err.Description)

    End Sub


    Public Sub GetMachineWorkbookData(ByVal Filename As String)

        Dim lvarExcel As New Excel.Application
        Dim lvarWorkBook As Workbook
        Dim lvarWorkSheet As Worksheet
        Dim Index As String
        Dim ValueString As String
        Dim Tokens() As String
        Dim row As Integer
        Dim ParentIndex As Integer = 0
        Dim Value As Integer
        Dim sw As New Stopwatch
        Dim i As Integer
        Dim j As Integer

        On Error GoTo ErrorHandler

        Me.Cursor = Cursors.WaitCursor

        pbStatus.Value = 0
        lblTime.Text = ""

        pbStatus.Visible = True
        lblTime.Visible = True

        sw.Start()

        ReDim gOutputValues(50000)

        lvarWorkBook = lvarExcel.Workbooks.Open(Filename, , False)

        lvarWorkSheet = lvarWorkBook.Worksheets.Item(1)

        With lvarWorkSheet

            Index = "A1"
            gOutputValues(0).PartNumber = lvarWorkSheet.Range(Index).Value

            Index = "C1"
            gOutputValues(0).Description = lvarWorkSheet.Range(Index).Value

            Index = "G1"
            gOutputValues(0).Revision = lvarWorkSheet.Range(Index).Value

            gOutputValues(0).Level = 0
            gOutputValues(0).FirstLevelIndex = 0
            gOutputValues(0).ParentIndex = 0
            gOutputValues(0).Index = 0

        End With

        For i = 1 To gOutputValues.Length

            row = i + 3

            ValueString = lvarWorkSheet.Cells(row, 4).value

            If ValueString = "" Then

                ReDim Preserve gOutputValues(i - 1)

                Exit For

            Else

                Tokens = Strings.Split(ValueString, "_")

                For j = 0 To Tokens.Length - 1

                    gOutputValues(i).PartNumber = lvarWorkSheet.Cells(row, 6).value
                    gOutputValues(i).Description = lvarWorkSheet.Cells(row, 9).value
                    gOutputValues(i).Revision = "" & lvarWorkSheet.Cells(row, 13).value
                    gOutputValues(i).Level = j + 1
                    gOutputValues(i).FirstLevelIndex = Tokens(0)

                    gOutputValues(i).Index = Tokens(j)

                Next

                If i > 0 Then

                    If gOutputValues(i).Level > gOutputValues(i - 1).Level Then

                        ParentIndex = i - 1

                    ElseIf gOutputValues(i).Level < gOutputValues(i - 1).Level Then

                        ParentIndex = 0

                        For j = gOutputValues(i - 1).ParentIndex To 1 Step -1

                            If gOutputValues(j).Level = gOutputValues(i).Level Then

                                ParentIndex = gOutputValues(j).ParentIndex

                                Exit For

                            End If
                        Next

                    End If

                End If

                gOutputValues(i).ParentIndex = ParentIndex

            End If

            If (i Mod 25) = 0 Then

                Value = (i / 10000) * 100

                If Value <= 100 Then
                    pbStatus.Value = Value
                    lblTime.Text = Format(sw.ElapsedMilliseconds / 1000, "0.0")
                End If

            End If

        Next

        pbStatus.Value = 50

        lvarWorkSheet = lvarWorkBook.Worksheets.Add(After:=lvarWorkBook.Worksheets(lvarWorkBook.Worksheets.Count))
        lvarWorkSheet.Name = "Assembly Listing" & lvarWorkBook.Worksheets.Count - 1

        lvarWorkSheet.Columns("A:C").ColumnWidth = 10
        lvarWorkSheet.Columns("D:D").ColumnWidth = 100

        lvarWorkSheet.Columns("F:F").ColumnWidth = 20
        lvarWorkSheet.Columns("G:G").ColumnWidth = 75
        lvarWorkSheet.Columns("H:H").ColumnWidth = 15

        lvarWorkSheet.Columns("A:C").HorizontalAlignment = Excel.Constants.xlCenter

        lvarWorkSheet.Columns("F:F").HorizontalAlignment = Excel.Constants.xlCenter
        lvarWorkSheet.Columns("H:H").HorizontalAlignment = Excel.Constants.xlCenter

        lvarWorkSheet.Cells(1, 1).value = "Item #"
        lvarWorkSheet.Cells(1, 2).value = "Parent Item Link"
        lvarWorkSheet.Cells(1, 3).value = "Parent Item"
        lvarWorkSheet.Cells(1, 4).value = "Full Description"

        lvarWorkSheet.Cells(1, 6).value = "Part Number"
        lvarWorkSheet.Cells(1, 7).value = "Description"
        lvarWorkSheet.Cells(1, 8).value = "Revision"

        For j = 1 To 8
            lvarWorkSheet.Cells(1, j).font.bold = True
        Next

        lvarWorkSheet.Cells(1, 4).HorizontalAlignment = Excel.Constants.xlCenter
        lvarWorkSheet.Cells(1, 7).HorizontalAlignment = Excel.Constants.xlCenter

        lvarWorkSheet.Range("A1:H1").Font.Size = 16
        lvarWorkSheet.Range("A1:H1").VerticalAlignment = Excel.Constants.xlCenter
        lvarWorkSheet.Range("A1:H1").WrapText = True

        lvarWorkSheet.Rows("1:1").RowHeight = 70

        For i = 0 To gOutputValues.Length - 1

            lvarWorkSheet.Cells(i + 2, 1).value = i

            lvarWorkSheet.Cells(i + 2, 2).value = gOutputValues(i).ParentIndex
            lvarWorkSheet.Cells(i + 2, 3).value = gOutputValues(i).ParentIndex

            If gOutputValues(i).Revision <> "" Then
                ValueString = gOutputValues(i).PartNumber & " " & gOutputValues(i).Description & " rev " & gOutputValues(i).Revision
            Else
                ValueString = gOutputValues(i).PartNumber & " " & gOutputValues(i).Description
            End If


            If i > 0 AndAlso gOutputValues(i).Level > gOutputValues(i - 1).Level Then
                lvarWorkSheet.Cells(i + 1, 4).Font.bold = True
            End If

            For j = 1 To gOutputValues(i).Level
                ValueString = "          " & ValueString
            Next

            lvarWorkSheet.Cells(i + 2, 4).value = ValueString
            lvarWorkSheet.Cells(i + 2, 6).value = gOutputValues(i).PartNumber
            lvarWorkSheet.Cells(i + 2, 7).value = gOutputValues(i).Description
            lvarWorkSheet.Cells(i + 2, 8).value = gOutputValues(i).Revision

            If i > 0 Then
                lvarWorkSheet.Hyperlinks.Add(Anchor:=lvarWorkSheet.Cells(i + 2, 2), Address:="", SubAddress:="A" & gOutputValues(i).ParentIndex + 2)
            End If

            If (i Mod 25) = 0 Then

                Value = 50 + (i / gOutputValues.Length) * 50

                If Value <= 100 Then
                    pbStatus.Value = Value
                    lblTime.Text = Format(sw.ElapsedMilliseconds / 1000, "0.0")
                End If

            End If

        Next

        With lvarWorkSheet.Cells(2, 5)

            .AddComment()
            .Comment.Visible = False
            .Comment.Text(Text:="created " & Format(Now, "yyyy-MM-dd HH:mm:ss") & vbCrLf & " by AssemblyTreeBuddy" & vbCrLf & "")
            .Comment.Visible = True
            .comment.shape.select(True)
            .comment.shape.ScaleWidth(1.5, 0, 0)
            .comment.shape.ScaleHeight(0.5, 0, 0)
            .comment.shape.IncrementLeft(-200)
            .comment.shape.IncrementTop(30)
            .Comment.Shape.TextFrame.Characters.Font.Color = Color.Red

        End With

        lvarWorkSheet.Cells(1, 5).Select()

        With lvarExcel.ActiveWindow
            .SplitColumn = 3
            .SplitRow = 1
            .FreezePanes = True
        End With

Done:
        On Error Resume Next

        sw.Stop()
        sw = Nothing

        lblTime.Visible = False
        pbStatus.Visible = False

        lvarWorkBook.Save()
        lvarExcel.Quit()

        'lvarExcel.Visible = True

        Marshal.ReleaseComObject(lvarWorkSheet)
        Marshal.ReleaseComObject(lvarWorkBook)
        Marshal.ReleaseComObject(lvarExcel)

        Me.Cursor = Cursors.Arrow

        cmdOpenAssemblyData.Visible = True

        cmdOpenAssemblyData.PerformClick()

        Exit Sub

ErrorHandler:
        Me.Cursor = Cursors.Arrow
        LogErrorMessage("sub GetMachineWorkbookData", Err.Description)
        GoTo Done

    End Sub


    Private Sub cmdTreeForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTreeForm.Click

        If gTreeForm Is Nothing Then
            gTreeForm = New frmTree
        End If

        gTreeForm.Show()

        gTreeForm.WindowState = FormWindowState.Normal

    End Sub


    Private Sub tmrStartup_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrStartup.Tick

        tmrStartup.Enabled = False


    End Sub


    Private Sub cmdOpenAssemblyData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOpenAssemblyData.Click

        If gAssemblyTreeDataForm Is Nothing Then
            gAssemblyTreeDataForm = New frmAssemblyTreeData
        End If

        gAssemblyTreeDataForm.Show()

        gAssemblyTreeDataForm.WindowState = FormWindowState.Normal

    End Sub


End Class
