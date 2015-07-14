<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAssemblyTreeData
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAssemblyTreeData))
        Me.tmrUpdate = New System.Windows.Forms.Timer(Me.components)
        Me.tmrStartup = New System.Windows.Forms.Timer(Me.components)
        Me.sgData = New SourceGrid.Grid()
        Me.muMain = New System.Windows.Forms.MenuStrip()
        Me.mnuView = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuShowAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuRollUpLevel1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuRollUpLevel2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuRollUpLevel3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.DataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuExportToWorksheet = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCopyToClipboard = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuOpenExcelfile = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuShowAssemblies = New System.Windows.Forms.ToolStripMenuItem()
        Me.muMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'tmrUpdate
        '
        Me.tmrUpdate.Enabled = True
        Me.tmrUpdate.Interval = 1000
        '
        'tmrStartup
        '
        Me.tmrStartup.Enabled = True
        '
        'sgData
        '
        Me.sgData.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.sgData.ColumnsCount = 3
        Me.sgData.EnableSort = True
        Me.sgData.FixedRows = 1
        Me.sgData.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sgData.Location = New System.Drawing.Point(12, 41)
        Me.sgData.Name = "sgData"
        Me.sgData.OptimizeMode = SourceGrid.CellOptimizeMode.ForRows
        Me.sgData.RowsCount = 1
        Me.sgData.SelectionMode = SourceGrid.GridSelectionMode.Row
        Me.sgData.Size = New System.Drawing.Size(994, 689)
        Me.sgData.TabIndex = 28
        Me.sgData.TabStop = True
        Me.sgData.ToolTipText = ""
        '
        'muMain
        '
        Me.muMain.BackColor = System.Drawing.Color.LightYellow
        Me.muMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuView, Me.DataToolStripMenuItem})
        Me.muMain.Location = New System.Drawing.Point(0, 0)
        Me.muMain.Name = "muMain"
        Me.muMain.Size = New System.Drawing.Size(1018, 24)
        Me.muMain.TabIndex = 29
        Me.muMain.Text = "MenuStrip1"
        '
        'mnuView
        '
        Me.mnuView.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuShowAll, Me.ToolStripSeparator2, Me.mnuRollUpLevel1, Me.mnuRollUpLevel2, Me.mnuRollUpLevel3, Me.ToolStripSeparator3, Me.mnuShowAssemblies})
        Me.mnuView.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mnuView.Name = "mnuView"
        Me.mnuView.Size = New System.Drawing.Size(51, 20)
        Me.mnuView.Text = "View"
        '
        'mnuShowAll
        '
        Me.mnuShowAll.Name = "mnuShowAll"
        Me.mnuShowAll.Size = New System.Drawing.Size(187, 22)
        Me.mnuShowAll.Tag = "0"
        Me.mnuShowAll.Text = "Show All Items"
        '
        'mnuRollUpLevel1
        '
        Me.mnuRollUpLevel1.Name = "mnuRollUpLevel1"
        Me.mnuRollUpLevel1.Size = New System.Drawing.Size(187, 22)
        Me.mnuRollUpLevel1.Tag = "1"
        Me.mnuRollUpLevel1.Text = "Roll Up Level 1"
        '
        'mnuRollUpLevel2
        '
        Me.mnuRollUpLevel2.Name = "mnuRollUpLevel2"
        Me.mnuRollUpLevel2.Size = New System.Drawing.Size(187, 22)
        Me.mnuRollUpLevel2.Tag = "2"
        Me.mnuRollUpLevel2.Text = "Roll Up Level 2"
        '
        'mnuRollUpLevel3
        '
        Me.mnuRollUpLevel3.Name = "mnuRollUpLevel3"
        Me.mnuRollUpLevel3.Size = New System.Drawing.Size(187, 22)
        Me.mnuRollUpLevel3.Tag = "3"
        Me.mnuRollUpLevel3.Text = "Roll Up Level 3"
        '
        'DataToolStripMenuItem
        '
        Me.DataToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuExportToWorksheet, Me.mnuCopyToClipboard, Me.ToolStripSeparator1, Me.mnuOpenExcelfile})
        Me.DataToolStripMenuItem.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataToolStripMenuItem.Name = "DataToolStripMenuItem"
        Me.DataToolStripMenuItem.Size = New System.Drawing.Size(51, 20)
        Me.DataToolStripMenuItem.Text = "Data"
        '
        'mnuExportToWorksheet
        '
        Me.mnuExportToWorksheet.Name = "mnuExportToWorksheet"
        Me.mnuExportToWorksheet.Size = New System.Drawing.Size(289, 22)
        Me.mnuExportToWorksheet.Text = "Export View To Worksheet"
        '
        'mnuCopyToClipboard
        '
        Me.mnuCopyToClipboard.Name = "mnuCopyToClipboard"
        Me.mnuCopyToClipboard.Size = New System.Drawing.Size(289, 22)
        Me.mnuCopyToClipboard.Text = "Copy View Contents To Clipboard"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(286, 6)
        '
        'mnuOpenExcelfile
        '
        Me.mnuOpenExcelfile.Name = "mnuOpenExcelfile"
        Me.mnuOpenExcelfile.Size = New System.Drawing.Size(289, 22)
        Me.mnuOpenExcelfile.Text = "Open Excel Data File"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(184, 6)
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(184, 6)
        '
        'mnuShowAssemblies
        '
        Me.mnuShowAssemblies.Name = "mnuShowAssemblies"
        Me.mnuShowAssemblies.Size = New System.Drawing.Size(187, 22)
        Me.mnuShowAssemblies.Tag = "20"
        Me.mnuShowAssemblies.Text = "Show Assemblies"
        '
        'frmAssemblyTreeData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1018, 742)
        Me.Controls.Add(Me.sgData)
        Me.Controls.Add(Me.muMain)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.muMain
        Me.MaximizeBox = False
        Me.Name = "frmAssemblyTreeData"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Assembly Tree Data"
        Me.muMain.ResumeLayout(False)
        Me.muMain.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tmrUpdate As System.Windows.Forms.Timer
    Friend WithEvents tmrStartup As System.Windows.Forms.Timer
    Friend WithEvents sgData As SourceGrid.Grid
    Friend WithEvents muMain As System.Windows.Forms.MenuStrip
    Friend WithEvents mnuView As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuRollUpLevel1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuRollUpLevel2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuRollUpLevel3 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuShowAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuExportToWorksheet As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCopyToClipboard As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuOpenExcelfile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuShowAssemblies As System.Windows.Forms.ToolStripMenuItem
End Class
