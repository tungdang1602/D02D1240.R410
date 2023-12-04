<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D02F3001
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D02F3001))
        Me.grp1 = New System.Windows.Forms.GroupBox()
        Me.txtAssetS1Name = New System.Windows.Forms.TextBox()
        Me.chkDisabled = New System.Windows.Forms.CheckBox()
        Me.txtAssetS1ID = New System.Windows.Forms.TextBox()
        Me.lblAssetS1ID = New System.Windows.Forms.Label()
        Me.lblAssetS1Name = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.grp1.SuspendLayout()
        Me.SuspendLayout()
        '
        'grp1
        '
        Me.grp1.Controls.Add(Me.txtAssetS1Name)
        Me.grp1.Controls.Add(Me.chkDisabled)
        Me.grp1.Controls.Add(Me.txtAssetS1ID)
        Me.grp1.Controls.Add(Me.lblAssetS1ID)
        Me.grp1.Controls.Add(Me.lblAssetS1Name)
        Me.grp1.Location = New System.Drawing.Point(5, 0)
        Me.grp1.Name = "grp1"
        Me.grp1.Size = New System.Drawing.Size(405, 84)
        Me.grp1.TabIndex = 0
        Me.grp1.TabStop = False
        '
        'txtAssetS1Name
        '
        Me.txtAssetS1Name.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtAssetS1Name.Location = New System.Drawing.Point(102, 46)
        Me.txtAssetS1Name.MaxLength = 250
        Me.txtAssetS1Name.Name = "txtAssetS1Name"
        Me.txtAssetS1Name.Size = New System.Drawing.Size(286, 22)
        Me.txtAssetS1Name.TabIndex = 4
        '
        'chkDisabled
        '
        Me.chkDisabled.AutoSize = True
        Me.chkDisabled.Location = New System.Drawing.Point(259, 20)
        Me.chkDisabled.Name = "chkDisabled"
        Me.chkDisabled.Size = New System.Drawing.Size(98, 17)
        Me.chkDisabled.TabIndex = 2
        Me.chkDisabled.Text = "Không sử dụng"
        Me.chkDisabled.UseVisualStyleBackColor = True
        '
        'txtAssetS1ID
        '
        Me.txtAssetS1ID.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtAssetS1ID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtAssetS1ID.Location = New System.Drawing.Point(103, 17)
        Me.txtAssetS1ID.MaxLength = 20
        Me.txtAssetS1ID.Name = "txtAssetS1ID"
        Me.txtAssetS1ID.Size = New System.Drawing.Size(128, 22)
        Me.txtAssetS1ID.TabIndex = 1
        '
        'lblAssetS1ID
        '
        Me.lblAssetS1ID.AutoSize = True
        Me.lblAssetS1ID.Location = New System.Drawing.Point(8, 22)
        Me.lblAssetS1ID.Name = "lblAssetS1ID"
        Me.lblAssetS1ID.Size = New System.Drawing.Size(68, 13)
        Me.lblAssetS1ID.TabIndex = 0
        Me.lblAssetS1ID.Text = "Mã phân loại"
        Me.lblAssetS1ID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblAssetS1Name
        '
        Me.lblAssetS1Name.AutoSize = True
        Me.lblAssetS1Name.Location = New System.Drawing.Point(8, 51)
        Me.lblAssetS1Name.Name = "lblAssetS1Name"
        Me.lblAssetS1Name.Size = New System.Drawing.Size(72, 13)
        Me.lblAssetS1Name.TabIndex = 3
        Me.lblAssetS1Name.Text = "Tên phân loại"
        Me.lblAssetS1Name.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(169, 97)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(76, 22)
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "&Lưu"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(252, 97)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(76, 22)
        Me.btnNext.TabIndex = 2
        Me.btnNext.Text = "Nhập &tiếp"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(335, 97)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(76, 22)
        Me.btnClose.TabIndex = 3
        Me.btnClose.Text = "Đó&ng"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'D02F3001
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(417, 130)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.grp1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D02F3001"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CËp nhËt ph¡n loÁi - D02F3001"
        Me.grp1.ResumeLayout(False)
        Me.grp1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grp1 As System.Windows.Forms.GroupBox
    Private WithEvents txtAssetS1ID As System.Windows.Forms.TextBox
    Private WithEvents lblAssetS1ID As System.Windows.Forms.Label
    Private WithEvents chkDisabled As System.Windows.Forms.CheckBox
    Private WithEvents txtAssetS1Name As System.Windows.Forms.TextBox
    Private WithEvents lblAssetS1Name As System.Windows.Forms.Label
    Private WithEvents btnSave As System.Windows.Forms.Button
    Private WithEvents btnNext As System.Windows.Forms.Button
    Private WithEvents btnClose As System.Windows.Forms.Button
End Class