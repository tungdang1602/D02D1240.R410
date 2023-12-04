<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D02F0201
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D02F0201))
        Me.txtSourceID = New System.Windows.Forms.TextBox()
        Me.lblSourceID = New System.Windows.Forms.Label()
        Me.grp1 = New System.Windows.Forms.GroupBox()
        Me.txtSourceName = New System.Windows.Forms.TextBox()
        Me.chkDisabled = New System.Windows.Forms.CheckBox()
        Me.lblSourceName = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.grp1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtSourceID
        '
        Me.txtSourceID.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSourceID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtSourceID.Location = New System.Drawing.Point(100, 15)
        Me.txtSourceID.MaxLength = 20
        Me.txtSourceID.Name = "txtSourceID"
        Me.txtSourceID.Size = New System.Drawing.Size(114, 22)
        Me.txtSourceID.TabIndex = 1
        '
        'lblSourceID
        '
        Me.lblSourceID.AutoSize = True
        Me.lblSourceID.Location = New System.Drawing.Point(4, 19)
        Me.lblSourceID.Name = "lblSourceID"
        Me.lblSourceID.Size = New System.Drawing.Size(55, 13)
        Me.lblSourceID.TabIndex = 1
        Me.lblSourceID.Text = "Mã nguồn"
        Me.lblSourceID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'grp1
        '
        Me.grp1.Controls.Add(Me.txtSourceName)
        Me.grp1.Controls.Add(Me.chkDisabled)
        Me.grp1.Controls.Add(Me.txtSourceID)
        Me.grp1.Controls.Add(Me.lblSourceID)
        Me.grp1.Controls.Add(Me.lblSourceName)
        Me.grp1.Location = New System.Drawing.Point(9, 5)
        Me.grp1.Name = "grp1"
        Me.grp1.Size = New System.Drawing.Size(390, 73)
        Me.grp1.TabIndex = 0
        Me.grp1.TabStop = False
        '
        'txtSourceName
        '
        Me.txtSourceName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtSourceName.Location = New System.Drawing.Point(100, 42)
        Me.txtSourceName.MaxLength = 50
        Me.txtSourceName.Name = "txtSourceName"
        Me.txtSourceName.Size = New System.Drawing.Size(280, 22)
        Me.txtSourceName.TabIndex = 3
        '
        'chkDisabled
        '
        Me.chkDisabled.AutoSize = True
        Me.chkDisabled.Location = New System.Drawing.Point(282, 18)
        Me.chkDisabled.Name = "chkDisabled"
        Me.chkDisabled.Size = New System.Drawing.Size(98, 17)
        Me.chkDisabled.TabIndex = 2
        Me.chkDisabled.Text = "Không sử dụng"
        Me.chkDisabled.UseVisualStyleBackColor = True
        '
        'lblSourceName
        '
        Me.lblSourceName.AutoSize = True
        Me.lblSourceName.Location = New System.Drawing.Point(3, 47)
        Me.lblSourceName.Name = "lblSourceName"
        Me.lblSourceName.Size = New System.Drawing.Size(59, 13)
        Me.lblSourceName.TabIndex = 4
        Me.lblSourceName.Text = "Tên nguồn"
        Me.lblSourceName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(171, 84)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(72, 22)
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "&Lưu"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(249, 84)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(72, 22)
        Me.btnNext.TabIndex = 2
        Me.btnNext.Text = "Nhập &tiếp"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(327, 84)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(72, 22)
        Me.btnClose.TabIndex = 3
        Me.btnClose.Text = "Đó&ng"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'D02F0201
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(409, 113)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.grp1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D02F0201"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CËp nhËt nguän hØnh thªnh - D02F0201"
        Me.grp1.ResumeLayout(False)
        Me.grp1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Private WithEvents txtSourceID As System.Windows.Forms.TextBox
    Private WithEvents lblSourceID As System.Windows.Forms.Label
    Private WithEvents grp1 As System.Windows.Forms.GroupBox
    Private WithEvents txtSourceName As System.Windows.Forms.TextBox
    Private WithEvents chkDisabled As System.Windows.Forms.CheckBox
    Private WithEvents lblSourceName As System.Windows.Forms.Label
    Private WithEvents btnSave As System.Windows.Forms.Button
    Private WithEvents btnNext As System.Windows.Forms.Button
    Private WithEvents btnClose As System.Windows.Forms.Button
End Class