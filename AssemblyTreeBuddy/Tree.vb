
Imports System.Drawing.Printing
Imports System.Drawing.Imaging


Public Class frmTree


    Private Declare Auto Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As IntPtr, ByVal nXDest As Integer, ByVal _
                                                            nYDest As Integer, ByVal nWidth As Integer, ByVal _
                                                            nHeight As Integer, ByVal hdcSrc As IntPtr, ByVal nXSrc _
                                                            As Integer, ByVal nYSrc As Integer, ByVal dwRop As System.Int32) As Boolean

    Private Const SRCCOPY As Integer = &HCC0020

    Private mvarScrollBufferWidth As Integer
    Private mvarHscrollValue As Integer
    Private mvarVscrollValue As Integer

    ' Variables used to print.
    Private mvarPrintBitmap As Bitmap
    Private WithEvents mvarPrintDocument As PrintDocument


    Private Sub frmTree_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        gTreeForm = Nothing

    End Sub


    Private Sub frmTree_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        scbHscroll.Maximum = 2560   'My.Computer.Screen.Bounds.Width
        scbVscroll.Maximum = 2048   'My.Computer.Screen.Bounds.Height

        scbHscroll.LargeChange = 256
        scbVscroll.LargeChange = 256

        scbHscroll.SmallChange = GridX
        scbVscroll.SmallChange = GridY

    End Sub


    Private Sub frmTree_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown, Me.Activated

        tmrStartup.Enabled = True

    End Sub


    Private Sub frmTree_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown

        If e.Button = MouseButtons.Right Then

            mnuPopup.Show(Me.Left + e.X, Me.Top + e.Y)

        End If


    End Sub


    Private Sub frmTree_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseWheel

        If e.Y > scbHscroll.Top Then

            If e.Delta > 0 Then  ' move left

                If scbHscroll.Value - e.Delta < 0 Then
                    scbHscroll.Value = 0
                Else
                    scbHscroll.Value = scbHscroll.Value - e.Delta
                End If

            Else    ' move right

                If scbHscroll.Value - e.Delta > scbHscroll.Maximum Then
                    scbHscroll.Value = scbHscroll.Maximum
                Else
                    scbHscroll.Value = scbHscroll.Value - e.Delta
                End If

            End If

        Else

            If e.Delta > 0 Then  ' move up

                If scbVscroll.Value - e.Delta < 0 Then
                    scbVscroll.Value = 0
                Else
                    scbVscroll.Value = scbVscroll.Value - e.Delta
                End If

            Else    ' move down

                If scbVscroll.Value - e.Delta > scbVscroll.Maximum Then
                    scbVscroll.Value = scbVscroll.Maximum
                Else
                    scbVscroll.Value = scbVscroll.Value - e.Delta
                End If

            End If

        End If

    End Sub


    Private Function GetFormImage() As Bitmap

        ' Get this form's Graphics object.
        Dim me_gr As Graphics = Me.CreateGraphics

        ' Make a Bitmap to hold the image.
        Dim bm As New Bitmap(Me.ClientSize.Width, Me.ClientSize.Height, me_gr)
        Dim bm_gr As Graphics = Graphics.FromImage(bm)
        Dim bm_hdc As IntPtr = bm_gr.GetHdc

        ' Get the form's hDC. We must do this after 
        ' creating the new Bitmap, which uses me_gr.
        Dim me_hdc As IntPtr = me_gr.GetHdc

        ' BitBlt the form's image onto the Bitmap.
        BitBlt(bm_hdc, 0, 0, Me.ClientSize.Width, Me.ClientSize.Height, me_hdc, 0, 0, SRCCOPY)
        me_gr.ReleaseHdc(me_hdc)
        bm_gr.ReleaseHdc(bm_hdc)

        ' Return the result.
        Return bm

    End Function


    Private Function GetFormImage2(Optional ByVal with_decorations As Boolean = True) As Image

        Dim bm As New Bitmap(Me.Width, Me.Height)

        Me.DrawToBitmap(bm, New Rectangle(0, 0, Me.Width, Me.Height))

        If Not with_decorations Then

            ' Trim off the decorations.
            Dim client_origin As Point = Me.PointToScreen(New Point(0, 0))
            Dim deco_left As Integer = client_origin.X - Me.Left
            Dim deco_top As Integer = client_origin.Y - Me.Top
            Dim wid As Integer = Me.ClientRectangle.Width
            Dim hgt As Integer = Me.ClientRectangle.Height
            Dim new_bm As New Bitmap(wid, hgt)
            Dim gr As Graphics = Graphics.FromImage(new_bm)

            gr.DrawImage(bm, New Rectangle(0, 0, wid, hgt), New Rectangle(deco_left, deco_top, wid, hgt), GraphicsUnit.Pixel)
            bm = new_bm

        End If

        Return bm

    End Function


    Private Function GetCompleteFormImage(Optional ByVal PrintFlag As Boolean = False) As Bitmap

        Dim MaxWidth As Integer = 2560
        Dim MaxHeight As Integer = 2048
        Dim x1 As Integer
        Dim y1 As Integer
        Dim x2 As Integer
        Dim y2 As Integer
        Dim DwgText As String = ""

        Dim StringMeasure As System.Drawing.SizeF
        Dim Rect As System.Drawing.Rectangle
        Dim br As New System.Drawing.SolidBrush(Color.Black)
        Dim TextBrush As New System.Drawing.SolidBrush(Color.Black)
        Dim p As New System.Drawing.Pen(Brushes.Black)

        Dim bm As Bitmap
        Dim gr As Graphics = Me.CreateGraphics 'Graphics.FromImage(bm)

        If PrintFlag Then
            bm = New Bitmap(MaxWidth, MaxHeight)
            gr = Graphics.FromImage(bm)
        End If

        gr.Clear(Color.White)
        p.Width = 1

        StringMeasure = gr.MeasureString(vbTab, Me.Font)

        Dim TabLength As Integer = StringMeasure.Width

        For i = 1 To 5

            DwgText = "    Drawing " & vbCrLf & "    " & 300000 + i & vbTab

            StringMeasure = gr.MeasureString(DwgText, Me.Font)
            Rect.X = 50 - scbHscroll.Value
            Rect.Y = 100 * i - scbVscroll.Value
            Rect.Width = StringMeasure.Width - TabLength / 2
            Rect.Height = StringMeasure.Height * 2

            TextBrush.Color = Color.Black
            br.Color = Color.White
            gr.FillRectangle(br, Rect)
            'gr.FillEllipse(br, Rect)
            gr.DrawString(DwgText, Me.Font, TextBrush, Rect.X, Rect.Y + Rect.Height / 4)

            gr.DrawRectangle(Pens.Black, Rect)
            'gr.DrawEllipse(Pens.Black, Rect)

            x1 = 35 - scbHscroll.Value
            y1 = Rect.Y + (Rect.Height / 2)
            x2 = Rect.X
            y2 = y1

            gr.DrawLine(p, x1, y1, x2, y2)

        Next

        ' draw top horizontal line
        x1 = 20 - scbHscroll.Value
        y1 = 50 - scbVscroll.Value
        x2 = 2000 - scbHscroll.Value
        y2 = y1

        gr.DrawLine(p, x1, y1, x2, y2)

        ' draw vertical line
        x1 = 35 - scbHscroll.Value
        y1 = 50 - scbVscroll.Value
        x2 = x1
        y2 = 600 - scbVscroll.Value

        gr.DrawLine(p, x1, y1, x2, y2)

        Return bm

        br.Dispose()
        TextBrush.Dispose()
        p.Dispose()
        gr.Dispose()

        If PrintFlag Then
            bm.Dispose()
        End If

    End Function


    Private Sub PrintTabloidSize()

        Dim lvarPrintDialog As New PrintDialog
        Dim lvarPageSettings As System.Drawing.Printing.PageSettings = New System.Drawing.Printing.PageSettings

        On Error GoTo ErrorHandler

        With lvarPageSettings
            .Landscape = True
            .PaperSize = New System.Drawing.Printing.PaperSize("Tabloid", 1100, 1700)
            .Margins.Top = 0
            .Margins.Bottom = 0
            .Margins.Left = 0
            .Margins.Right = 0
        End With

        scbHscroll.Visible = False
        scbVscroll.Visible = False
        Sleep(250)
        Me.Refresh()

        mvarPrintBitmap = GetCompleteFormImage(True)

        scbHscroll.Visible = True
        scbVscroll.Visible = True

        mvarPrintDocument = New PrintDocument

        mvarPrintDocument.DefaultPageSettings = lvarPageSettings

        lvarPrintDialog.Document = mvarPrintDocument

        Dim r As DialogResult = lvarPrintDialog.ShowDialog

        If r = Windows.Forms.DialogResult.OK Then
            mvarPrintDocument.Print()
        End If

        Exit Sub

ErrorHandler:
        LogErrorMessage(Me.Name & ": sub mnuPrintTabloidSize_Click", Err.Description)

    End Sub


    Private Sub mvarPrintDocument_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles mvarPrintDocument.PrintPage

        '' Draw the image centered.
        'Dim x As Integer = e.MarginBounds.X + (e.MarginBounds.Width - mvarPrintBitmap.Width) \ 2
        'Dim y As Integer = e.MarginBounds.Y + (e.MarginBounds.Height - mvarPrintBitmap.Height) \ 2

        'e.Graphics.DrawImage(mvarPrintBitmap, x, y)

        '' There's only one page.
        'e.HasMorePages = False

        Dim PrintArea As System.Drawing.RectangleF = e.Graphics.VisibleClipBounds
        Dim NewImageSize As System.Drawing.RectangleF

        Dim SF As Double = CDbl(PrintArea.Width) / CDbl(mvarPrintBitmap.Width)

        NewImageSize.Width = CInt(mvarPrintBitmap.Width * SF)
        NewImageSize.Height = CInt(mvarPrintBitmap.Height * SF)

        'You can influence the quality of the resized image
        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        e.Graphics.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        e.Graphics.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality

        'Draw the image to the printer
        e.Graphics.DrawImage(mvarPrintBitmap, NewImageSize)

    End Sub


    Private Sub tmrStartup_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrStartup.Tick

        tmrStartup.Enabled = False

        GetCompleteFormImage()

    End Sub


    Private Sub scbHscroll_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles scbHscroll.Scroll

        Dim DeltaX As Integer
        Dim UpdateValid As Boolean = False

        On Error GoTo ErrorHandler

        Exit Sub

        If e Is Nothing Then
            UpdateValid = True
        ElseIf e.Type = ScrollEventType.EndScroll Then
            UpdateValid = True
        End If

        If UpdateValid Then

            LockWindowUpdate(Me.Handle.ToInt32)

            If scbHscroll.Value >= scbHscroll.Maximum - mvarScrollBufferWidth Then
                scbHscroll.Maximum = scbHscroll.Maximum + My.Computer.Screen.Bounds.Width / 2
            End If

            DeltaX = mvarHscrollValue - scbHscroll.Value
            mvarHscrollValue = scbHscroll.Value

            LockWindowUpdate(0)

        End If

        Exit Sub

ErrorHandler:
        LogErrorMessage(Me.Name & ": sub scbHscroll_Scroll, ", Err.Description)

    End Sub


    Private Sub scbVscroll_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles scbVscroll.Scroll

        Dim DeltaY As Integer
        Dim UpdateValid As Boolean = False

        On Error GoTo ErrorHandler

        Exit Sub

        If e Is Nothing Then
            UpdateValid = True
        ElseIf e.Type = ScrollEventType.EndScroll Then
            UpdateValid = True
        End If

        If UpdateValid Then

            LockWindowUpdate(Me.Handle.ToInt32)

            If scbVscroll.Value >= scbVscroll.Maximum - mvarScrollBufferWidth Then
                scbVscroll.Maximum = scbVscroll.Maximum + My.Computer.Screen.Bounds.Height / 2
            End If

            DeltaY = mvarVscrollValue - scbVscroll.Value
            mvarVscrollValue = scbVscroll.Value

            LockWindowUpdate(0)

        End If

        Exit Sub

ErrorHandler:
        LogErrorMessage(Me.Name & ": sub scbVscroll_Scroll, ", Err.Description)

    End Sub


    Private Sub scbVscroll_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles scbVscroll.ValueChanged

        Dim RemY As Integer
        Static UpdateGraphics As Boolean

        RemY = scbVscroll.Value Mod GridY

        If scbVscroll.Value > 0 Then

            If RemY > (GridY / 2) Then
                scbVscroll.Value = scbVscroll.Value + GridY - RemY
            ElseIf RemY > 0 And RemY <= (GridY / 2) Then
                scbVscroll.Value = scbVscroll.Value - RemY
            End If

        End If

        If UpdateGraphics Or scbVscroll.Value = 0 Then
            GetCompleteFormImage()
        End If

        UpdateGraphics = Not UpdateGraphics


    End Sub


    Private Sub scbHscroll_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles scbHscroll.ValueChanged

        Dim RemX As Integer
        Static UpdateGraphics As Boolean

        RemX = scbHscroll.Value Mod GridX

        If scbHscroll.Value > 0 Then

            If RemX > (GridX / 2) Then
                scbHscroll.Value = scbHscroll.Value + GridX - RemX
            ElseIf RemX > 0 And RemX <= (GridX / 2) Then
                scbHscroll.Value = scbHscroll.Value - RemX
            End If

        End If

        If UpdateGraphics Or scbHscroll.Value = 0 Then
            GetCompleteFormImage()
        End If

        UpdateGraphics = Not UpdateGraphics

    End Sub


    Private Sub mnuPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPrint.Click

        On Error GoTo ErrorHandler

        scbVscroll.Value = 0
        scbHscroll.Value = 0

        tmrPrint.Enabled = True

        Exit Sub

ErrorHandler:
        LogErrorMessage(Me.Name & ": sub mnuPrint_Click, ", Err.Description)

    End Sub


    Private Sub tmrPrint_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrPrint.Tick

        tmrPrint.Enabled = False

        PrintTabloidSize()

    End Sub



End Class