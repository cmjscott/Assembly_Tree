<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.cmdOpenFile = New System.Windows.Forms.Button()
        Me.cmdTreeForm = New System.Windows.Forms.Button()
        Me.tmrStartup = New System.Windows.Forms.Timer(Me.components)
        Me.pbStatus = New System.Windows.Forms.ProgressBar()
        Me.lblTime = New System.Windows.Forms.Label()
        Me.picMain = New System.Windows.Forms.PictureBox()
        Me.cmdOpenAssemblyData = New System.Windows.Forms.Button()
        CType(Me.picMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdOpenFile
        '
        Me.cmdOpenFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOpenFile.Location = New System.Drawing.Point(37, 309)
        Me.cmdOpenFile.Name = "cmdOpenFile"
        Me.cmdOpenFile.Size = New System.Drawing.Size(104, 40)
        Me.cmdOpenFile.TabIndex = 1
        Me.cmdOpenFile.Text = "Open Assembly Data File"
        Me.cmdOpenFile.UseVisualStyleBackColor = True
        '
        'cmdTreeForm
        '
        Me.cmdTreeForm.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdTreeForm.Location = New System.Drawing.Point(287, 309)
        Me.cmdTreeForm.Name = "cmdTreeForm"
        Me.cmdTreeForm.Size = New System.Drawing.Size(104, 40)
        Me.cmdTreeForm.TabIndex = 3
        Me.cmdTreeForm.Text = "Open Tree"
        Me.cmdTreeForm.UseVisualStyleBackColor = True
        Me.cmdTreeForm.Visible = False
        '
        'tmrStartup
        '
        Me.tmrStartup.Enabled = True
        '
        'pbStatus
        '
        Me.pbStatus.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pbStatus.ForeColor = System.Drawing.Color.LimeGreen
        Me.pbStatus.Location = New System.Drawing.Point(0, 361)
        Me.pbStatus.Name = "pbStatus"
        Me.pbStatus.Size = New System.Drawing.Size(434, 35)
        Me.pbStatus.TabIndex = 4
        Me.pbStatus.Visible = False
        '
        'lblTime
        '
        Me.lblTime.BackColor = System.Drawing.Color.White
        Me.lblTime.Location = New System.Drawing.Point(12, 9)
        Me.lblTime.Name = "lblTime"
        Me.lblTime.Size = New System.Drawing.Size(84, 25)
        Me.lblTime.TabIndex = 5
        Me.lblTime.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'picMain
        '
        Me.picMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.picMain.Image = Global.AssemblyTreeBuddy.My.Resources.Resources.Tree
        Me.picMain.Location = New System.Drawing.Point(0, 0)
        Me.picMain.Name = "picMain"
        Me.picMain.Size = New System.Drawing.Size(434, 396)
        Me.picMain.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picMain.TabIndex = 2
        Me.picMain.TabStop = False
        '
        'cmdOpenAssemblyData
        '
        Me.cmdOpenAssemblyData.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOpenAssemblyData.Location = New System.Drawing.Point(318, 12)
        Me.cmdOpenAssemblyData.Name = "cmdOpenAssemblyData"
        Me.cmdOpenAssemblyData.Size = New System.Drawing.Size(104, 40)
        Me.cmdOpenAssemblyData.TabIndex = 6
        Me.cmdOpenAssemblyData.Text = "Open Assembly Data"
        Me.cmdOpenAssemblyData.UseVisualStyleBackColor = True
        Me.cmdOpenAssemblyData.Visible = False
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(434, 396)
        Me.Controls.Add(Me.cmdOpenAssemblyData)
        Me.Controls.Add(Me.lblTime)
        Me.Controls.Add(Me.pbStatus)
        Me.Controls.Add(Me.cmdTreeForm)
        Me.Controls.Add(Me.cmdOpenFile)
        Me.Controls.Add(Me.picMain)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Assembly Tree Buddy"
        CType(Me.picMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdOpenFile As System.Windows.Forms.Button
    Friend WithEvents picMain As System.Windows.Forms.PictureBox
    Friend WithEvents cmdTreeForm As System.Windows.Forms.Button
    Friend WithEvents tmrStartup As System.Windows.Forms.Timer
    Friend WithEvents pbStatus As System.Windows.Forms.ProgressBar
    Friend WithEvents lblTime As System.Windows.Forms.Label
    Friend WithEvents cmdOpenAssemblyData As System.Windows.Forms.Button

End Class
