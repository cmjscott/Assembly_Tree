<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTree
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTree))
        Me.scbVscroll = New System.Windows.Forms.VScrollBar()
        Me.scbHscroll = New System.Windows.Forms.HScrollBar()
        Me.tmrStartup = New System.Windows.Forms.Timer(Me.components)
        Me.mnuPopup = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mnuPrint = New System.Windows.Forms.ToolStripMenuItem()
        Me.tmrPrint = New System.Windows.Forms.Timer(Me.components)
        Me.mnuPopup.SuspendLayout()
        Me.SuspendLayout()
        '
        'scbVscroll
        '
        Me.scbVscroll.Dock = System.Windows.Forms.DockStyle.Right
        Me.scbVscroll.LargeChange = 50
        Me.scbVscroll.Location = New System.Drawing.Point(996, 0)
        Me.scbVscroll.Name = "scbVscroll"
        Me.scbVscroll.Size = New System.Drawing.Size(20, 720)
        Me.scbVscroll.SmallChange = 50
        Me.scbVscroll.TabIndex = 16
        '
        'scbHscroll
        '
        Me.scbHscroll.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.scbHscroll.LargeChange = 50
        Me.scbHscroll.Location = New System.Drawing.Point(0, 720)
        Me.scbHscroll.Name = "scbHscroll"
        Me.scbHscroll.Size = New System.Drawing.Size(1016, 20)
        Me.scbHscroll.SmallChange = 50
        Me.scbHscroll.TabIndex = 15
        '
        'tmrStartup
        '
        Me.tmrStartup.Enabled = True
        '
        'mnuPopup
        '
        Me.mnuPopup.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuPrint})
        Me.mnuPopup.Name = "mnuPopup"
        Me.mnuPopup.Size = New System.Drawing.Size(97, 26)
        '
        'mnuPrint
        '
        Me.mnuPrint.Name = "mnuPrint"
        Me.mnuPrint.Size = New System.Drawing.Size(96, 22)
        Me.mnuPrint.Tag = "print"
        Me.mnuPrint.Text = "Print"
        '
        'tmrPrint
        '
        Me.tmrPrint.Interval = 250
        '
        'frmTree
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1016, 740)
        Me.Controls.Add(Me.scbVscroll)
        Me.Controls.Add(Me.scbHscroll)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmTree"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Tree"
        Me.mnuPopup.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents scbVscroll As System.Windows.Forms.VScrollBar
    Friend WithEvents scbHscroll As System.Windows.Forms.HScrollBar
    Friend WithEvents tmrStartup As System.Windows.Forms.Timer
    Friend WithEvents mnuPopup As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents mnuPrint As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tmrPrint As System.Windows.Forms.Timer
End Class
