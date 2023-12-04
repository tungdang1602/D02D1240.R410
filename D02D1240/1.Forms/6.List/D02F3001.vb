'#-------------------------------------------------------------------------------------
'# Created Date: 27/08/2007 3:46:17 PM
'# Created User: Trần Thị ÁiTrâm
'# Modify Date: 
'# Modify User: 
'#-------------------------------------------------------------------------------------
Imports System.Text
Imports System

Public Class D02F3001

    Private _savedOk As Boolean
    Public ReadOnly Property  SavedOk() As Boolean
        Get
            Return _savedOk
        End Get
    End Property

    Private _assetID As String = ""
  

    Public Property AssetID() As String
        Get
            Return _assetID
        End Get
        Set(ByVal Value As String)
            _assetID = Value
        End Set
    End Property

    Private _indexTab As Integer = 0

    Public Property IndexTab() As Integer
        Get
            Return _indexTab
        End Get
        Set(ByVal Value As Integer)
            _indexTab = Value
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
                    btnSave.Enabled = True
                    btnNext.Enabled = False
                    LoadAddNew()
                Case EnumFormState.FormEdit
                    btnSave.Enabled = True
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                    LoadEdit()
                Case EnumFormState.FormView
                    btnSave.Enabled = False
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                    LoadEdit()
            End Select
        End Set
    End Property

    Private Sub LoadMaster()
        Dim sSQL As String = ""
        Dim dt As DataTable

        Select Case _indexTab
            Case 0
                sSQL = "Select AssetS1ID, AssetS1Name, AssetS1NameU, Disabled From D02T1000 WITH(NOLOCK) Where AssetS1ID=" & SQLString(_assetID)
                dt = ReturnDataTable(sSQL)
                If dt.Rows.Count <= 0 Then Exit Sub
                With dt.Rows(0)
                    txtAssetS1ID.Text = .Item("AssetS1ID").ToString
                    chkDisabled.Checked = CBool(.Item("Disabled"))
                    txtAssetS1Name.Text = .Item("AssetS1Name" & UnicodeJoin(gbUnicode)).ToString
                End With
            Case 1
                sSQL = "Select AssetS2ID, AssetS2Name, AssetS2NameU, Disabled From D02T2000 WITH(NOLOCK) Where AssetS2ID=" & SQLString(_assetID)
                dt = ReturnDataTable(sSQL)
                If dt.Rows.Count <= 0 Then Exit Sub
                With dt.Rows(0)
                    txtAssetS1ID.Text = .Item("AssetS2ID").ToString
                    chkDisabled.Checked = CBool(.Item("Disabled"))
                    txtAssetS1Name.Text = .Item("AssetS2Name" & UnicodeJoin(gbUnicode)).ToString
                End With
            Case 2
                sSQL = "Select AssetS3ID, AssetS3Name, AssetS3NameU, Disabled From D02T3000 WITH(NOLOCK) Where AssetS3ID=" & SQLString(_assetID)
                dt = ReturnDataTable(sSQL)
                If dt.Rows.Count <= 0 Then Exit Sub
                With dt.Rows(0)
                    txtAssetS1ID.Text = .Item("AssetS3ID").ToString
                    chkDisabled.Checked = CBool(.Item("Disabled"))
                    txtAssetS1Name.Text = .Item("AssetS3Name" & UnicodeJoin(gbUnicode)).ToString
                End With
            Case 3
                sSQL = "Select AssetS4ID, AssetS4Name, AssetS4NameU, Disabled From D02T4000 WITH(NOLOCK) Where AssetS4ID=" & SQLString(_assetID)
                dt = ReturnDataTable(sSQL)
                If dt.Rows.Count <= 0 Then Exit Sub
                With dt.Rows(0)
                    txtAssetS1ID.Text = .Item("AssetS4ID").ToString
                    chkDisabled.Checked = CBool(.Item("Disabled"))
                    txtAssetS1Name.Text = .Item("AssetS4Name" & UnicodeJoin(gbUnicode)).ToString
                End With
            Case 4
                sSQL = "Select AssetS5ID, AssetS5Name, AssetS5NameU, Disabled From D02T5003 WITH(NOLOCK) Where AssetS5ID=" & SQLString(_assetID)
                dt = ReturnDataTable(sSQL)
                If dt.Rows.Count <= 0 Then Exit Sub
                With dt.Rows(0)
                    txtAssetS1ID.Text = .Item("AssetS5ID").ToString
                    chkDisabled.Checked = CBool(.Item("Disabled"))
                    txtAssetS1Name.Text = .Item("AssetS5Name" & UnicodeJoin(gbUnicode)).ToString
                End With
        End Select
    End Sub

    Private Sub LoadAddNew()
        txtAssetS1ID.Text = ""
        txtAssetS1Name.Text = ""
        chkDisabled.Visible = False
    End Sub

    Private Sub LoadEdit()
        txtAssetS1ID.Enabled = False
        chkDisabled.Visible = True
        LoadMaster()
        txtAssetS1Name.Focus()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub D02F3001_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
    End Sub

    Private Sub D02F3001_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
	If bLoadFormState = False Then FormState = _formState
        Loadlanguage()
        SetBackColorObligatory()
        CheckIdTextBox(txtAssetS1ID)
        InputbyUnicode(Me, gbUnicode)
    SetResolutionForm(Me)
Me.Cursor = Cursors.Default
End Sub

    Private Sub SetBackColorObligatory()
        txtAssetS1ID.BackColor = COLOR_BACKCOLOROBLIGATORY
        txtAssetS1Name.BackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub

    Private Sub btnNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNext.Click
        btnSave.Enabled = True
        btnNext.Enabled = False
        LoadAddNew()
        txtAssetS1ID.Focus()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        If Not AllowSave() Then Exit Sub

        'Kiểm tra Ngày phiếu có phù hợp với kỳ kế toán hiện tại không (gọi hàm CheckVoucherDateInPeriod)

        btnSave.Enabled = False
        btnClose.Enabled = False
        _savedOk = False
        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder
        Select Case _FormState
            Case EnumFormState.FormAdd
                Select Case _indexTab
                    Case 0
                        sSQL.Append(SQLInsertD02T1000)
                    Case 1
                        sSQL.Append(SQLInsertD02T2000)
                    Case 2
                        sSQL.Append(SQLInsertD02T3000)
                    Case 3
                        sSQL.Append(SQLInsertD02T4000)
                    Case 4
                        sSQL.Append(SQLInsertD02T5003)
                End Select
                'Lưu LastKey của Số phiếu xuống Database (gọi hàm CreateIGEVoucherNo bật cờ True)
                'Kiểm tra trùng Số phiếu (gọi hàm CheckDuplicateVoucherNo)

            Case EnumFormState.FormEdit
                Select Case _indexTab
                    Case 0
                        sSQL.Append(SQLUpdateD02T1000)
                    Case 1
                        sSQL.Append(SQLUpdateD02T2000)
                    Case 2
                        sSQL.Append(SQLUpdateD02T3000)
                    Case 3
                        sSQL.Append(SQLUpdateD02T4000)
                    Case 4
                        sSQL.Append(SQLUpdateD02T5003)
                End Select

        End Select

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            _savedOk = True
            btnClose.Enabled = True
            Select Case _FormState
                Case EnumFormState.FormAdd
                    _assetID = txtAssetS1ID.Text
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

    Private Function AllowSave() As Boolean
        Dim sSQL As String = ""
        If txtAssetS1ID.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rl3("Ma_phan_loai") & " " & (_indexTab + 1))
            txtAssetS1ID.Focus()
            Return False
        End If
        sSQL = "Select S1Length, AssetS1Enabled, S2Length, AssetS2Enabled, S3Length, AssetS3Enabled, S4Length, AssetS4Enabled, S5Length, AssetS5Enabled From D02T0000 WITH(NOLOCK)"
        Dim dt As DataTable = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            If txtAssetS1ID.Text.Trim <> "" Then
                Select Case _indexTab
                    Case 0
                        If CBool(dt.Rows(0).Item("AssetS1Enabled")) = True Then
                            If txtAssetS1ID.Text.Trim.Length > CInt(dt.Rows(0).Item("S1Length")) Then
                                D99C0008.MsgL3(rl3("Do_dai") & " " & rl3("Ma_phan_loai") & " " & (_indexTab + 1) & " " & rl3("khong_duoc_vuot_qua") & " " & CInt(dt.Rows(0).Item("S1Length")) & " " & rl3("_ky_tu"))
                                txtAssetS1ID.Focus()
                                Return False
                            End If
                        End If

                    Case 1
                        If CBool(dt.Rows(0).Item("AssetS2Enabled")) = True Then
                            If txtAssetS1ID.Text.Trim.Length > CInt(dt.Rows(0).Item("S2Length")) Then
                                D99C0008.MsgL3(rl3("Do_dai") & " " & rl3("Ma_phan_loai") & " " & (_indexTab + 1) & " " & rl3("khong_duoc_vuot_qua") & " " & CInt(dt.Rows(0).Item("S2Length")) & " " & rl3("_ky_tu"))
                                txtAssetS1ID.Focus()
                                Return False
                            End If
                        End If

                    Case 2
                        If CBool(dt.Rows(0).Item("AssetS3Enabled")) = True Then
                            If txtAssetS1ID.Text.Trim.Length > CInt(dt.Rows(0).Item("S3Length")) Then
                                D99C0008.MsgL3(rl3("Do_dai") & " " & rl3("Ma_phan_loai") & " " & (_indexTab + 1) & " " & rl3("khong_duoc_vuot_qua") & " " & CInt(dt.Rows(0).Item("S3Length")) & " " & rl3("_ky_tu"))
                                txtAssetS1ID.Focus()
                                Return False
                            End If
                        End If

                    Case 3
                        If CBool(dt.Rows(0).Item("AssetS4Enabled")) = True Then
                            If txtAssetS1ID.Text.Trim.Length > CInt(dt.Rows(0).Item("S4Length")) Then
                                D99C0008.MsgL3(rL3("Do_dai") & " " & rL3("Ma_phan_loai") & " " & (_indexTab + 1) & " " & rL3("khong_duoc_vuot_qua") & " " & CInt(dt.Rows(0).Item("S4Length")) & " " & rL3("_ky_tu"))
                                txtAssetS1ID.Focus()
                                Return False
                            End If
                        End If

                    Case 4
                        If CBool(dt.Rows(0).Item("AssetS5Enabled")) = True Then
                            If txtAssetS1ID.Text.Trim.Length > CInt(dt.Rows(0).Item("S5Length")) Then
                                D99C0008.MsgL3(rL3("Do_dai") & " " & rL3("Ma_phan_loai") & " " & (_indexTab + 1) & " " & rL3("khong_duoc_vuot_qua") & " " & CInt(dt.Rows(0).Item("S5Length")) & " " & rL3("_ky_tu"))
                                txtAssetS1ID.Focus()
                                Return False
                            End If
                        End If

                End Select

            End If
        End If

        If txtAssetS1Name.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rl3("Ten_phan_loai") & " " & (_indexTab + 1))
            txtAssetS1Name.Focus()
            Return False
        End If

        If _FormState = EnumFormState.FormAdd Then
            Select Case _indexTab
                Case 0
                    sSQL = "Select AssetS1ID From D02T1000 WITH(NOLOCK) Where AssetS1ID=" & SQLString(txtAssetS1ID.Text)
                Case 1
                    sSQL = "Select AssetS2ID From D02T2000 WITH(NOLOCK) Where AssetS2ID=" & SQLString(txtAssetS1ID.Text)
                Case 2
                    sSQL = "Select AssetS3ID From D02T3000 WITH(NOLOCK) Where AssetS3ID=" & SQLString(txtAssetS1ID.Text)
                Case 3
                    sSQL = "Select AssetS4ID From D02T4000 WITH(NOLOCK) Where AssetS4ID=" & SQLString(txtAssetS1ID.Text)
                Case 4
                    sSQL = "Select AssetS5ID From D02T5003 WITH(NOLOCK) Where AssetS5ID=" & SQLString(txtAssetS1ID.Text)
            End Select

            If ExistRecord(sSQL) Then
                D99C0008.MsgDuplicatePKey()
                txtAssetS1ID.Focus()
                Return False
            End If
        End If

        Return True
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T1000
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 28/08/2007 08:26:56
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T1000() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Insert Into D02T1000(")
        sSQL.Append("AssetS1ID, AssetS1NameU, Disabled, CreateUserID, CreateDate, ")
        sSQL.Append("LastModifyUserID, LastModifyDate")
        sSQL.Append(") Values(")
        sSQL.Append(SQLString(txtAssetS1ID.Text) & COMMA) 'AssetS1ID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLStringUnicode(txtAssetS1Name.Text, gbUnicode, True) & COMMA) 'AssetS1Name, varchar[50], NULL
        sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, bit, NOT NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL
        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL
        sSQL.Append("GetDate()") 'LastModifyDate, datetime, NULL
        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T2000
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 28/08/2007 08:28:49
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T2000() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Insert Into D02T2000(")
        sSQL.Append("AssetS2ID, AssetS2NameU, Disabled, CreateUserID, CreateDate, ")
        sSQL.Append("LastModifyUserID, LastModifyDate")
        sSQL.Append(") Values(")
        sSQL.Append(SQLString(txtAssetS1ID.Text) & COMMA) 'AssetS1ID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLStringUnicode(txtAssetS1Name.Text, gbUnicode, True) & COMMA) 'AssetS1Name, varchar[50], NULL
        sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, bit, NOT NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL
        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL
        sSQL.Append("GetDate()") 'LastModifyDate, datetime, NULL
        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T3000
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 28/08/2007 08:29:01
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T3000() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Insert Into D02T3000(")
        sSQL.Append("AssetS3ID, AssetS3NameU, Disabled, CreateUserID, CreateDate, ")
        sSQL.Append("LastModifyUserID, LastModifyDate")
        sSQL.Append(") Values(")
        sSQL.Append(SQLString(txtAssetS1ID.Text) & COMMA) 'AssetS1ID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLStringUnicode(txtAssetS1Name.Text, gbUnicode, True) & COMMA) 'AssetS1Name, varchar[50], NULL
        sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, bit, NOT NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL
        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL
        sSQL.Append("GetDate()") 'LastModifyDate, datetime, NULL
        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T3000
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 28/08/2007 08:29:01
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T4000() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Insert Into D02T4000(")
        sSQL.Append("AssetS4ID, AssetS4NameU, Disabled, CreateUserID, CreateDate, ")
        sSQL.Append("LastModifyUserID, LastModifyDate")
        sSQL.Append(") Values(")
        sSQL.Append(SQLString(txtAssetS1ID.Text) & COMMA) 'AssetS1ID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLStringUnicode(txtAssetS1Name.Text, gbUnicode, True) & COMMA) 'AssetS1Name, varchar[50], NULL
        sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, bit, NOT NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL
        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL
        sSQL.Append("GetDate()") 'LastModifyDate, datetime, NULL
        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T3000
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 28/08/2007 08:29:01
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T5003() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Insert Into D02T5003(")
        sSQL.Append("AssetS5ID, AssetS5NameU, Disabled, CreateUserID, CreateDate, ")
        sSQL.Append("LastModifyUserID, LastModifyDate")
        sSQL.Append(") Values(")
        sSQL.Append(SQLString(txtAssetS1ID.Text) & COMMA) 'AssetS1ID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLStringUnicode(txtAssetS1Name.Text, gbUnicode, True) & COMMA) 'AssetS1Name, varchar[50], NULL
        sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, bit, NOT NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL
        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL
        sSQL.Append("GetDate()") 'LastModifyDate, datetime, NULL
        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T1000
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 28/08/2007 08:29:17
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T1000() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T1000 Set ")
        sSQL.Append("AssetS1NameU = " & SQLStringUnicode(txtAssetS1Name.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
        sSQL.Append("Disabled = " & SQLNumber(chkDisabled.Checked) & COMMA) 'bit, NOT NULL
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
        sSQL.Append("LastModifyDate = GetDate()") 'datetime, NULL
        sSQL.Append(" Where ")
        sSQL.Append("AssetS1ID = " & SQLString(txtAssetS1ID.Text))

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T2000
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 28/08/2007 08:29:25
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T2000() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T2000 Set ")
        sSQL.Append("AssetS2NameU = " & SQLStringUnicode(txtAssetS1Name.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
        sSQL.Append("Disabled = " & SQLNumber(chkDisabled.Checked) & COMMA) 'bit, NOT NULL
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
        sSQL.Append("LastModifyDate = GetDate()") 'datetime, NULL
        sSQL.Append(" Where ")
        sSQL.Append("AssetS2ID = " & SQLString(txtAssetS1ID.Text))

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T3000
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 28/08/2007 08:29:34
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T3000() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T3000 Set ")
        sSQL.Append("AssetS3NameU = " & SQLStringUnicode(txtAssetS1Name.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
        sSQL.Append("Disabled = " & SQLNumber(chkDisabled.Checked) & COMMA) 'bit, NOT NULL
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
        sSQL.Append("LastModifyDate = GetDate()") 'datetime, NULL
        sSQL.Append(" Where ")
        sSQL.Append("AssetS3ID = " & SQLString(txtAssetS1ID.Text))

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T4000
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 28/08/2007 08:29:34
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T4000() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T4000 Set ")
        sSQL.Append("AssetS4NameU = " & SQLStringUnicode(txtAssetS1Name.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
        sSQL.Append("Disabled = " & SQLNumber(chkDisabled.Checked) & COMMA) 'bit, NOT NULL
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
        sSQL.Append("LastModifyDate = GetDate()") 'datetime, NULL
        sSQL.Append(" Where ")
        sSQL.Append("AssetS4ID = " & SQLString(txtAssetS1ID.Text))

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T5003
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 28/08/2007 08:29:34
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T5003() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T5003 Set ")
        sSQL.Append("AssetS5NameU = " & SQLStringUnicode(txtAssetS1Name.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
        sSQL.Append("Disabled = " & SQLNumber(chkDisabled.Checked) & COMMA) 'bit, NOT NULL
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
        sSQL.Append("LastModifyDate = GetDate()") 'datetime, NULL
        sSQL.Append(" Where ")
        sSQL.Append("AssetS5ID = " & SQLString(txtAssetS1ID.Text))

        Return sSQL
    End Function

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Cap_nhat_phan_loai") & " " & (_indexTab + 1) & " - " & Me.Name & UnicodeCaption(gbUnicode)
        '================================================================ 
        lblAssetS1ID.Text = rl3("Ma_phan_loai") & " " & (_indexTab + 1) 'Mã phân loại
        lblAssetS1Name.Text = rl3("Ten_phan_loai") & " " & (_indexTab + 1) 'Tên phân loại
        '================================================================ 
        btnSave.Text = rl3("_Luu") '&Lưu
        btnNext.Text = rl3("Nhap__tiep") 'Nhập &tiếp
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        '================================================================ 
        chkDisabled.Text = rl3("Khong_su_dung") 'Không sử dụng
        '================================================================ 
    End Sub
End Class