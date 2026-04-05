<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormAutoSave
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormAutoSave))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.gridFiles = New System.Windows.Forms.DataGridView()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnOpenFile = New System.Windows.Forms.Button()
        Me.btnSaveAs = New System.Windows.Forms.Button()
        Me.progress = New System.Windows.Forms.ProgressBar()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.chkbSearchBackup = New System.Windows.Forms.CheckBox()
        Me.chkbSearchXAR = New System.Windows.Forms.CheckBox()
        Me.txtCustomLocation = New System.Windows.Forms.TextBox()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.chkbSearchXLK = New System.Windows.Forms.CheckBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.lblCustomLocation = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.chkbTextSearch = New System.Windows.Forms.CheckBox()
        Me.txtTextSearch = New System.Windows.Forms.TextBox()
        Me.colFileName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colDateTime = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colFolder = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gridFiles, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(670, 11)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(31, 20)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 14
        Me.PictureBox1.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(634, 11)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(31, 20)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 15
        Me.PictureBox2.TabStop = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.White
        Me.Label11.Location = New System.Drawing.Point(256, 6)
        Me.Label11.Name = "Label11"
        Me.Label11.Padding = New System.Windows.Forms.Padding(3)
        Me.Label11.Size = New System.Drawing.Size(199, 32)
        Me.Label11.TabIndex = 0
        Me.Label11.Text = "Auto Saved Files"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'gridFiles
        '
        Me.gridFiles.AllowUserToAddRows = False
        Me.gridFiles.AllowUserToDeleteRows = False
        Me.gridFiles.AllowUserToResizeRows = False
        Me.gridFiles.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gridFiles.BackgroundColor = System.Drawing.Color.White
        Me.gridFiles.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.gridFiles.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridFiles.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colFileName, Me.colDateTime, Me.colFolder})
        Me.gridFiles.GridColor = System.Drawing.Color.Black
        Me.gridFiles.Location = New System.Drawing.Point(12, 222)
        Me.gridFiles.MultiSelect = False
        Me.gridFiles.Name = "gridFiles"
        Me.gridFiles.ReadOnly = True
        Me.gridFiles.RowHeadersVisible = False
        Me.gridFiles.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gridFiles.Size = New System.Drawing.Size(686, 183)
        Me.gridFiles.TabIndex = 6
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Location = New System.Drawing.Point(-52, 156)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(813, 2)
        Me.GroupBox1.TabIndex = 55
        Me.GroupBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Location = New System.Drawing.Point(-52, 42)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(813, 2)
        Me.GroupBox2.TabIndex = 57
        Me.GroupBox2.TabStop = False
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(598, 435)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(100, 28)
        Me.btnCancel.TabIndex = 11
        Me.btnCancel.Text = "Close"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOpenFile
        '
        Me.btnOpenFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnOpenFile.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnOpenFile.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOpenFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnOpenFile.Location = New System.Drawing.Point(11, 435)
        Me.btnOpenFile.Name = "btnOpenFile"
        Me.btnOpenFile.Size = New System.Drawing.Size(100, 28)
        Me.btnOpenFile.TabIndex = 7
        Me.btnOpenFile.Text = "Open File"
        Me.btnOpenFile.UseVisualStyleBackColor = True
        '
        'btnSaveAs
        '
        Me.btnSaveAs.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSaveAs.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSaveAs.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnSaveAs.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnSaveAs.Location = New System.Drawing.Point(117, 435)
        Me.btnSaveAs.Name = "btnSaveAs"
        Me.btnSaveAs.Size = New System.Drawing.Size(100, 28)
        Me.btnSaveAs.TabIndex = 8
        Me.btnSaveAs.Text = "Save As..."
        Me.btnSaveAs.UseVisualStyleBackColor = True
        '
        'progress
        '
        Me.progress.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.progress.Location = New System.Drawing.Point(223, 435)
        Me.progress.Name = "progress"
        Me.progress.Size = New System.Drawing.Size(263, 28)
        Me.progress.Step = 1
        Me.progress.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.progress.TabIndex = 9
        Me.progress.Visible = False
        '
        'btnStart
        '
        Me.btnStart.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnStart.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnStart.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnStart.Location = New System.Drawing.Point(492, 435)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(100, 28)
        Me.btnStart.TabIndex = 10
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'chkbSearchBackup
        '
        Me.chkbSearchBackup.AutoSize = True
        Me.chkbSearchBackup.BackColor = System.Drawing.Color.Transparent
        Me.chkbSearchBackup.Checked = True
        Me.chkbSearchBackup.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkbSearchBackup.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.chkbSearchBackup.ForeColor = System.Drawing.Color.White
        Me.chkbSearchBackup.Location = New System.Drawing.Point(13, 61)
        Me.chkbSearchBackup.Name = "chkbSearchBackup"
        Me.chkbSearchBackup.Size = New System.Drawing.Size(315, 17)
        Me.chkbSearchBackup.TabIndex = 0
        Me.chkbSearchBackup.Text = "Search within Excel backup locations for TMP files"
        Me.chkbSearchBackup.UseVisualStyleBackColor = False
        '
        'chkbSearchXAR
        '
        Me.chkbSearchXAR.AutoSize = True
        Me.chkbSearchXAR.BackColor = System.Drawing.Color.Transparent
        Me.chkbSearchXAR.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.chkbSearchXAR.ForeColor = System.Drawing.Color.White
        Me.chkbSearchXAR.Location = New System.Drawing.Point(13, 91)
        Me.chkbSearchXAR.Name = "chkbSearchXAR"
        Me.chkbSearchXAR.Size = New System.Drawing.Size(141, 17)
        Me.chkbSearchXAR.TabIndex = 1
        Me.chkbSearchXAR.Text = "Search for XAR files"
        Me.chkbSearchXAR.UseVisualStyleBackColor = False
        '
        'txtCustomLocation
        '
        Me.txtCustomLocation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCustomLocation.Enabled = False
        Me.txtCustomLocation.Location = New System.Drawing.Point(134, 121)
        Me.txtCustomLocation.Name = "txtCustomLocation"
        Me.txtCustomLocation.Size = New System.Drawing.Size(518, 20)
        Me.txtCustomLocation.TabIndex = 3
        '
        'btnBrowse
        '
        Me.btnBrowse.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowse.BackColor = System.Drawing.Color.Transparent
        Me.btnBrowse.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnBrowse.Enabled = False
        Me.btnBrowse.Location = New System.Drawing.Point(664, 116)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(34, 28)
        Me.btnBrowse.TabIndex = 4
        Me.btnBrowse.UseVisualStyleBackColor = False
        '
        'chkbSearchXLK
        '
        Me.chkbSearchXLK.AutoSize = True
        Me.chkbSearchXLK.BackColor = System.Drawing.Color.Transparent
        Me.chkbSearchXLK.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.chkbSearchXLK.ForeColor = System.Drawing.Color.White
        Me.chkbSearchXLK.Location = New System.Drawing.Point(173, 91)
        Me.chkbSearchXLK.Name = "chkbSearchXLK"
        Me.chkbSearchXLK.Size = New System.Drawing.Size(139, 17)
        Me.chkbSearchXLK.TabIndex = 2
        Me.chkbSearchXLK.Text = "Search for XLK files"
        Me.chkbSearchXLK.UseVisualStyleBackColor = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Location = New System.Drawing.Point(-51, 419)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(813, 2)
        Me.GroupBox3.TabIndex = 56
        Me.GroupBox3.TabStop = False
        '
        'lblCustomLocation
        '
        Me.lblCustomLocation.AutoSize = True
        Me.lblCustomLocation.BackColor = System.Drawing.Color.Transparent
        Me.lblCustomLocation.Enabled = False
        Me.lblCustomLocation.ForeColor = System.Drawing.Color.White
        Me.lblCustomLocation.Location = New System.Drawing.Point(48, 124)
        Me.lblCustomLocation.Name = "lblCustomLocation"
        Me.lblCustomLocation.Size = New System.Drawing.Size(80, 13)
        Me.lblCustomLocation.TabIndex = 58
        Me.lblCustomLocation.Text = "Select location:"
        '
        'GroupBox4
        '
        Me.GroupBox4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox4.Location = New System.Drawing.Point(-51, 207)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(813, 2)
        Me.GroupBox4.TabIndex = 56
        Me.GroupBox4.TabStop = False
        '
        'chkbTextSearch
        '
        Me.chkbTextSearch.AutoSize = True
        Me.chkbTextSearch.BackColor = System.Drawing.Color.Transparent
        Me.chkbTextSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.chkbTextSearch.ForeColor = System.Drawing.Color.White
        Me.chkbTextSearch.Location = New System.Drawing.Point(13, 175)
        Me.chkbTextSearch.Name = "chkbTextSearch"
        Me.chkbTextSearch.Size = New System.Drawing.Size(114, 17)
        Me.chkbTextSearch.TabIndex = 59
        Me.chkbTextSearch.Text = "Search for text:"
        Me.chkbTextSearch.UseVisualStyleBackColor = False
        '
        'txtTextSearch
        '
        Me.txtTextSearch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTextSearch.Enabled = False
        Me.txtTextSearch.Location = New System.Drawing.Point(133, 173)
        Me.txtTextSearch.Name = "txtTextSearch"
        Me.txtTextSearch.Size = New System.Drawing.Size(267, 20)
        Me.txtTextSearch.TabIndex = 5
        '
        'colFileName
        '
        Me.colFileName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.colFileName.HeaderText = "File Name"
        Me.colFileName.Name = "colFileName"
        Me.colFileName.ReadOnly = True
        Me.colFileName.Width = 79
        '
        'colDateTime
        '
        Me.colDateTime.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.colDateTime.HeaderText = "Last Changed"
        Me.colDateTime.Name = "colDateTime"
        Me.colDateTime.ReadOnly = True
        Me.colDateTime.Width = 98
        '
        'colFolder
        '
        Me.colFolder.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.colFolder.HeaderText = "Containing Folder"
        Me.colFolder.Name = "colFolder"
        Me.colFolder.ReadOnly = True
        '
        'FormAutoSave
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(710, 476)
        Me.Controls.Add(Me.txtTextSearch)
        Me.Controls.Add(Me.chkbTextSearch)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.lblCustomLocation)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.btnBrowse)
        Me.Controls.Add(Me.txtCustomLocation)
        Me.Controls.Add(Me.chkbSearchXLK)
        Me.Controls.Add(Me.chkbSearchXAR)
        Me.Controls.Add(Me.chkbSearchBackup)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.btnSaveAs)
        Me.Controls.Add(Me.btnOpenFile)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.gridFiles)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.progress)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FormAutoSave"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Excel Recovery"
        Me.TransparencyKey = System.Drawing.Color.Violet
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gridFiles, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents gridFiles As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOpenFile As System.Windows.Forms.Button
    Friend WithEvents btnSaveAs As System.Windows.Forms.Button
    Friend WithEvents progress As System.Windows.Forms.ProgressBar
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents chkbSearchBackup As System.Windows.Forms.CheckBox
    Friend WithEvents chkbSearchXAR As System.Windows.Forms.CheckBox
    Friend WithEvents txtCustomLocation As System.Windows.Forms.TextBox
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents chkbSearchXLK As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents lblCustomLocation As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents chkbTextSearch As System.Windows.Forms.CheckBox
    Friend WithEvents txtTextSearch As System.Windows.Forms.TextBox
    Friend WithEvents colFileName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colDateTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colFolder As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
