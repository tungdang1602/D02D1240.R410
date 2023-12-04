'#---------------------------------------------------------------------------------------------------
'# Title: D02E0201
'# Created User: Lê Thị Thanh Hiền
'# Created Date: 31/07/2007 10:51:43
'# Modified User: 
'# Modified Date: 
'# Description: 
'#---------------------------------------------------------------------------------------------------
Imports System.Text
Imports System
Public Class D02F0201

    Private _savedOK As Boolean
    Public ReadOnly Property SavedOK() As Boolean
        Get
            Return _savedOK
        End Get
    End Property

    Private _sourceID As String
    Public Property SourceID() As String
        Get
            Return _sourceID
        End Get
        Set(ByVal Value As String)
            If _sourceID = Value Then
                _sourceID = ""
                Return
            End If
            _sourceID = Value
        End Set
    End Property
    Dim bLoadFormState As Boolean = False
	Private _FormState As EnumFormState
    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal value As EnumFormState)
	bLoadFormState = True
	LoadInfoGeneral()
            _FormState = value
            Select Case _FormState
                Case EnumFormState.FormAdd
                    btnNext.Enabled = False

                Case EnumFormState.FormEdit
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                    LoadEdit()
                Case EnumFormState.FormView
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                    btnSave.Enabled = False
                    LoadEdit()

            End Select
        End Set
    End Property
    Private Sub SetBackColorObligatory()
        txtSourceID.BackColor = COLOR_BACKCOLOROBLIGATORY
        txtSourceName.BackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub

    Private Sub LoadEdit()
        txtSourceID.Enabled = False
        Dim sSQL As String = ""
        sSQL = "select * from D02T0013 WITH(NOLOCK) where SourceID = " & SQLString(_sourceID)
        Dim dt As DataTable
        dt = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            txtSourceID.Text = dt.Rows(0).Item("SourceID").ToString
            txtSourceName.Text = dt.Rows(0).Item("SourceName" & UnicodeJoin(gbUnicode)).ToString
            chkDisabled.Checked = Convert.ToBoolean(dt.Rows(0).Item("Disabled"))
        End If
    End Sub

    Private Sub D02F0201_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '12/10/2020, id 144622-Tài sản cố định_Lỗi chưa cảnh báo khi lưu
        If _FormState = EnumFormState.FormEdit Then
            If Not _savedOK Then
                If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
            End If
        ElseIf _FormState = EnumFormState.FormAdd Then
            If (txtSourceID.Text <> "") Then
                If Not _savedOK Then
                    If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub D02F0201_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
    End Sub

    Private Sub D02F0201_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
	If bLoadFormState = False Then FormState = _formState
        Loadlanguage()
        SetBackColorObligatory()
        CheckIdTextBox(txtSourceID)
        InputbyUnicode(Me, gbUnicode)
        SetResolutionForm(Me)
    End Sub

    Private Function AllowSave() As Boolean
        If txtSourceID.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rl3("Ma_nguon"))
            txtSourceID.Focus()
            Return False
        End If
        If txtSourceName.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rl3("Ten_nguon"))
            txtSourceName.Focus()
            Return False
        End If
        If _FormState = EnumFormState.FormAdd Then
            Dim sSQL As String
            sSQL = "select top 1 1 from D02T0013 WITH(NOLOCK) where SourceID = " & SQLString(txtSourceID.Text)
            Dim dt As DataTable
            dt = ReturnDataTable(sSQL)
            If dt.Rows.Count > 0 Then
                D99C0008.MsgDuplicatePKey()
                txtSourceID.Focus()
                Return False
            End If
        End If

        Return True
    End Function

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        If Not AllowSave() Then Exit Sub

        'Kiểm tra Ngày phiếu có phù hợp với kỳ kế toán hiện tại không (gọi hàm CheckVoucherDateInPeriod)

        btnSave.Enabled = False
        btnClose.Enabled = False
        _savedOK = False
        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder

        Select Case _FormState
            Case EnumFormState.FormAdd
                sSQL = SQLInsertD02T0013()
            Case EnumFormState.FormEdit
                sSQL = SQLUpdateD02T0013()
        End Select

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            _savedOK = True
            btnClose.Enabled = True
            Select Case _FormState
                Case EnumFormState.FormAdd
                    _sourceID = txtSourceID.Text
                    btnNext.Enabled = True
                    btnNext.Focus()
                Case EnumFormState.FormEdit
                    btnSave.Enabled = True
                    btnClose.Focus()
            End Select
        Else
            SaveNotOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
        End If
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T0013
    '# Created User: Lê Thị Thanh Hiền
    '# Created Date: 30/07/2007 11:52:32
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T0013() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Insert Into D02T0013(")
        sSQL.Append("SourceID, SourceNameU, Disabled, CreateUserID, CreateDate, ")
        sSQL.Append("LastModifyUserID, LastModifyDate")
        sSQL.Append(") Values(")
        sSQL.Append(SQLString(txtSourceID.Text) & COMMA) 'SourceID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLStringUnicode(txtSourceName.Text, gbUnicode, True) & COMMA) 'SourceNameU, varchar[50], NULL
        sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, bit, NOT NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL
        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL
        sSQL.Append("GetDate()") 'LastModifyDate, datetime, NULL
        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0013
    '# Created User: Lê Thị Thanh Hiền
    '# Created Date: 30/07/2007 02:58:47
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0013() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0013 Set ")
        sSQL.Append("SourceNameU = " & SQLStringUnicode(txtSourceName.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
        sSQL.Append("Disabled = " & SQLNumber(chkDisabled.Checked) & COMMA) 'bit, NOT NULL
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
        sSQL.Append("LastModifyDate = GetDate()") 'datetime, NULL
        sSQL.Append(" Where ")
        sSQL.Append("SourceID = " & SQLString(txtSourceID.Text))

        Return sSQL
    End Function

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        btnSave.Enabled = True
        btnNext.Enabled = False
        txtSourceID.Text = ""
        txtSourceName.Text = ""
        chkDisabled.Checked = False
        txtSourceID.Focus()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Cap_nhat_nguon_hinh_thanh_-_D02F0201") & UnicodeCaption(gbUnicode) 'CËp nhËt nguän hØnh thªnh - D02F0201
        '================================================================ 
        lblSourceID.Text = rl3("Ma_nguon") 'Mã nguồn
        lblSourceName.Text = rl3("Ten_nguon") 'Tên nguồn
        '================================================================ 
        btnSave.Text = rl3("_Luu") '&Lưu
        btnNext.Text = rl3("Nhap__tiep") 'Nhập &tiếp
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        '================================================================ 
        chkDisabled.Text = rl3("Khong_su_dung") 'Không sử dụng
        '================================================================ 

    End Sub

End Class