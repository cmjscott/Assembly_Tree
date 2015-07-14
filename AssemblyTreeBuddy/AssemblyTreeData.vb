
Imports System.Data
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.MarshalByRefObject


Public Class frmAssemblyTreeData

    Private mvarLevel As Integer


    Private Sub frmAssemblyTreeData_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        gAssemblyTreeDataForm = Nothing

    End Sub


    Private Sub frmAssemblyTreeData_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim lvarCellView1 As New SourceGrid.Cells.Views.Cell
        Dim border = New DevAge.Drawing.RectangleBorder(New DevAge.Drawing.BorderLine(Color.Black), New DevAge.Drawing.BorderLine(Color.Black))
        Dim border2 = New DevAge.Drawing.RectangleBorder(New DevAge.Drawing.BorderLine(Color.Transparent))

        mvarLevel = 0

        With sgData

            .Rows(0).Height = 30

            .Columns(0).Width = 60
            .Columns(1).Width = 100
            .Columns(2).Width = 750

            .Columns(0).MinimalWidth = .Columns(0).Width
            .Columns(0).MaximalWidth = .Columns(0).Width

            .Columns(1).MinimalWidth = .Columns(1).Width
            .Columns(1).MaximalWidth = .Columns(1).Width

            .Columns(2).MinimalWidth = .Columns(2).Width
            .Columns(2).MaximalWidth = .Columns(2).Width

            lvarCellView1.BackColor = Color.PaleGoldenrod
            lvarCellView1.TextAlignment = DevAge.Drawing.ContentAlignment.MiddleCenter
            lvarCellView1.Font = New System.Drawing.Font(sgData.Font, FontStyle.Bold)
            lvarCellView1.Border = border

            sgData(0, 0) = New SourceGrid.Cells.ColumnHeader("Item #")
            sgData(0, 1) = New SourceGrid.Cells.ColumnHeader("Parent Item #")
            sgData(0, 2) = New SourceGrid.Cells.ColumnHeader("Description")

            sgData(0, 0).View = lvarCellView1
            sgData(0, 1).View = lvarCellView1
            sgData(0, 2).View = lvarCellView1

            .Selection.FocusStyle = SourceGrid.FocusStyle.None
            .Selection.EnableMultiSelection = False
            .Selection.BackColor = Color.Transparent
            .Selection.FocusBackColor = .Selection.BackColor
            .Selection.Border = border2

        End With

    End Sub


    Private Sub tmrStartup_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrStartup.Tick

        Dim lvarCellView1 As New SourceGrid.Cells.Views.Cell
        Dim lvarCellView2 As New SourceGrid.Cells.Views.Cell
        Dim lvarCellView3 As New SourceGrid.Cells.Views.Cell
        Dim border = New DevAge.Drawing.RectangleBorder(New DevAge.Drawing.BorderLine(Color.Black), New DevAge.Drawing.BorderLine(Color.Black))
        Dim Status As String
        Dim Desc As String
        Dim i As Integer
        Dim j As Integer
        Dim r As Integer = 1
        Dim p As SourceGrid.Position

        On Error GoTo ErrorHandler

        tmrStartup.Enabled = False

        lvarCellView1.BackColor = Color.White
        lvarCellView1.TextAlignment = DevAge.Drawing.ContentAlignment.MiddleCenter
        lvarCellView1.Border = border

        lvarCellView2.BackColor = Color.White
        lvarCellView2.TextAlignment = DevAge.Drawing.ContentAlignment.MiddleLeft
        lvarCellView2.Border = border

        lvarCellView3.BackColor = Color.White
        lvarCellView3.TextAlignment = DevAge.Drawing.ContentAlignment.MiddleLeft
        lvarCellView3.Font = New System.Drawing.Font(sgData.Font, FontStyle.Bold)
        lvarCellView3.Border = border

        r = 0

        With sgData

            For i = 0 To gOutputValues.Length - 1

                If gOutputValues(i).Level <= mvarLevel Or mvarLevel = 0 Or mvarLevel = 20 Then

                    If mvarLevel <> 20 Or (mvarLevel = 20 And Strings.Left(gOutputValues(i).PartNumber, 1) = "2") Then

                        .RowsCount = .RowsCount + 1

                        r = r + 1

                        sgData(r, 0) = New SourceGrid.Cells.Cell(i)
                        sgData(r, 0).View = lvarCellView1
                        sgData(r, 0).Tag = i + 1

                        sgData(r, 1) = New SourceGrid.Cells.Cell(gOutputValues(i).ParentIndex)
                        sgData(r, 1).View = lvarCellView1

                        If gOutputValues(i).Revision = "" Then
                            Desc = gOutputValues(i).PartNumber & " " & gOutputValues(i).Description
                        Else
                            Desc = gOutputValues(i).PartNumber & " " & gOutputValues(i).Description & " rev " & gOutputValues(i).Revision
                        End If

                        For j = 1 To gOutputValues(i).Level
                            Desc = "          " & Desc
                        Next

                        sgData(r, 2) = New SourceGrid.Cells.Cell(Desc)
                        sgData(r, 2).View = lvarCellView2

                        If mvarLevel = 20 Then
                            sgData(r, 2).View = lvarCellView3
                        Else

                            If r > 1 AndAlso gOutputValues(i).Level > gOutputValues(i - 1).Level Then
                                sgData(r - 1, 2).View = lvarCellView3
                            ElseIf r > 2 Then
                                sgData(r - 1, 2).View = lvarCellView2
                            End If

                        End If

                    End If

                End If

            Next

        End With

        sgData.Refresh()

        Select Case mvarLevel

            Case 0
                Me.Text = "Assembly Tree Data - All Items Shown"

            Case 1 To 3
                Me.Text = "Assembly Tree Data - Roll Up Level " & mvarLevel & " Items"

            Case 20
                Me.Text = "Assembly Tree Data - Assemblies Only"

        End Select


        Exit Sub

ErrorHandler:
        LogErrorMessage(Me.Name & ", sub tmrStartup_Tick", Err.Description)

    End Sub


    Private Sub sgData_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles sgData.MouseUp

        Dim row As Integer
        Dim col As Integer
        Dim ClickPoint As System.Drawing.Point
        Dim p As SourceGrid.Position
        Dim n As Integer
        Dim i As Integer

        On Error GoTo ErrorHandler

        With sgData

            ClickPoint.X = e.X
            ClickPoint.Y = e.Y

            p = .PositionAtPoint(ClickPoint)

            row = p.Row
            col = p.Column

            If row < 1 Or col < 0 Then
                Exit Sub
            End If

            If col = 1 Then

                If mvarLevel = 0 Then
                    n = CInt(sgData(row, 1).Value) + 1
                Else

                    For i = 1 To .RowsCount - 1

                        If CInt(sgData(i, 0).Value) = CInt(sgData(row, 1).Value) Then

                            n = i
                            Exit For

                        End If

                    Next

                End If

                .Selection.SelectRow(n, True)
                .Selection.FocusRow(n)

                .Selection.FocusFirstCell(True)

                .Selection.FocusStyle = SourceGrid.FocusStyle.RemoveSelectionOnLeave


            End If

        End With

        Exit Sub

ErrorHandler:
        LogErrorMessage(Me.Name & ", sub sgData_MouseDown", Err.Description)

    End Sub


    Private Sub RollUpLevels(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuShowAll.Click, mnuRollUpLevel1.Click, mnuRollUpLevel2.Click, mnuRollUpLevel3.Click, mnuShowAssemblies.Click

        mvarLevel = CInt(sender.tag)

        sgData.RowsCount = 1
        sgData.Refresh()

        tmrStartup.Enabled = True

    End Sub


    Private Sub mnuCopyToClipboard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCopyToClipboard.Click

        Dim i As Integer
        Dim j As Integer
        Dim Contents As String = ""

        On Error GoTo ErrorHandler

        Me.Cursor = Cursors.WaitCursor

        For i = 0 To sgData.RowsCount - 1

            For j = 0 To sgData.ColumnsCount - 1

                Contents = Contents & sgData(i, j).Value.ToString & ","

            Next

            Contents = Strings.Left(Contents, Len(Contents) - 1)

            Contents = Contents & vbCrLf

        Next

        Clipboard.Clear()

        Clipboard.SetText(Contents)

        Me.Cursor = Cursors.Arrow

        Exit Sub

ErrorHandler:
        Me.Cursor = Cursors.Arrow
        LogErrorMessage(Me.Name & ", sub mnuCopyToClipboard_Click", Err.Description)

    End Sub


    Private Sub mnuExportToWorksheet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExportToWorksheet.Click

        Dim lvarExcel As New Excel.Application
        Dim lvarWorkBook As Workbook
        Dim lvarWorkSheet As Worksheet
        Dim i As Integer
        Dim j As Integer

        On Error GoTo ErrorHandler

        Me.Cursor = Cursors.WaitCursor

        lvarWorkBook = lvarExcel.Workbooks.Open(gFilename, , False)

        lvarWorkSheet = lvarWorkBook.Worksheets.Add(After:=lvarWorkBook.Worksheets(lvarWorkBook.Worksheets.Count))

        Select Case mvarLevel

            Case 0
                lvarWorkSheet.Name = "All Items" & lvarWorkBook.Worksheets.Count - 1

            Case 1 To 3
                lvarWorkSheet.Name = "Level " & mvarLevel & " Roll Up" & lvarWorkBook.Worksheets.Count - 1

            Case 20
                lvarWorkSheet.Name = "Assemblies Only" & lvarWorkBook.Worksheets.Count - 1

        End Select

        lvarWorkSheet.Columns("A:B").ColumnWidth = 10
        lvarWorkSheet.Columns("C:C").ColumnWidth = 100

        lvarWorkSheet.Columns("A:B").HorizontalAlignment = Excel.Constants.xlCenter

        lvarWorkSheet.Cells(1, 1).value = "Item #"
        lvarWorkSheet.Cells(1, 2).value = "Parent Item"
        lvarWorkSheet.Cells(1, 3).value = "Full Description"

        For j = 1 To 8
            lvarWorkSheet.Cells(1, j).font.bold = True
        Next

        lvarWorkSheet.Cells(1, 3).HorizontalAlignment = Excel.Constants.xlCenter

        lvarWorkSheet.Range("A1:H1").Font.Size = 16
        lvarWorkSheet.Range("A1:H1").VerticalAlignment = Excel.Constants.xlCenter
        lvarWorkSheet.Range("A1:H1").WrapText = True

        lvarWorkSheet.Rows("1:1").RowHeight = 70

        For i = 0 To sgData.RowsCount - 1

            For j = 0 To sgData.ColumnsCount - 1

                lvarWorkSheet.Cells(i + 1, j + 1).value = sgData(i, j).Value.ToString

            Next

        Next

        With lvarExcel.ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
            .FreezePanes = True
        End With

Done:
        On Error Resume Next

        lvarWorkBook.Save()
        lvarExcel.Quit()

        'lvarExcel.Visible = True

        Marshal.ReleaseComObject(lvarWorkSheet)
        Marshal.ReleaseComObject(lvarWorkBook)
        Marshal.ReleaseComObject(lvarExcel)

        Me.Cursor = Cursors.Arrow

        Exit Sub

ErrorHandler:
        Me.Cursor = Cursors.Arrow
        LogErrorMessage("sub mnuExportToWorksheet_Click", Err.Description)
        GoTo Done

    End Sub


    Private Sub mnuOpenExcelfile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOpenExcelfile.Click

        Dim lvarExcel As New Excel.Application
        Dim lvarWorkBook As Workbook
        Dim lvarWorkSheet As Worksheet

        On Error GoTo ErrorHandler

        Me.Cursor = Cursors.WaitCursor

        lvarWorkBook = lvarExcel.Workbooks.Open(gFilename, , False)

Done:
        On Error Resume Next

        lvarExcel.Visible = True
        lvarExcel.WindowState = vbMaximizedFocus

        Marshal.ReleaseComObject(lvarWorkBook)
        Marshal.ReleaseComObject(lvarExcel)

        Me.Cursor = Cursors.Arrow

        Exit Sub

ErrorHandler:
        Me.Cursor = Cursors.Arrow
        LogErrorMessage("sub mnuOpenExcelfile_Click", Err.Description)
        GoTo Done

    End Sub


End Class