'#-------------------------------------------------------------------------------------
'# Created Date: 26/11/2007 10:36:29 AM
'# Created User: Trần Thị ÁiTrâm
'# Modify Date: 26/11/2007 10:36:29 AM
'# Modify User: Trần Thị ÁiTrâm
'#-------------------------------------------------------------------------------------
Imports System.Text
Imports System.Runtime.InteropServices
Imports System

Public Class D02F0101

    Private _savedOK As Boolean
    Public ReadOnly Property SavedOK() As Boolean
        Get
            Return _savedOK
        End Get
    End Property


#Region "Const of tdbg - Total of Columns: 13"
    Private Const COL_PeriodID As Integer = 0   ' Kỳ sản xuất
    Private Const COL_Ana01ID As Integer = 1    ' Khoản mục 01
    Private Const COL_Ana02ID As Integer = 2    ' Khoản mục 02
    Private Const COL_Ana03ID As Integer = 3    ' Khoản mục 03
    Private Const COL_Ana04ID As Integer = 4    ' Khoản mục 04
    Private Const COL_Ana05ID As Integer = 5    ' Khoản mục 05
    Private Const COL_Ana06ID As Integer = 6    ' Khoản mục 06
    Private Const COL_Ana07ID As Integer = 7    ' Khoản mục 07
    Private Const COL_Ana08ID As Integer = 8    ' Khoản mục 08
    Private Const COL_Ana09ID As Integer = 9    ' Khoản mục 09
    Private Const COL_Ana10ID As Integer = 10   ' Khoản mục 10
    Private Const COL_ProjectID As Integer = 11 ' Dự án
    Private Const COL_TaskID As Integer = 12    ' Hạng mục
#End Region


#Region "Const of tdbgGuide"
    Private Const COLG_VariableID As Integer = 0  ' VariableID
    Private Const COLG_Description As Integer = 1 ' Description
#End Region

    Private _assignmentID As String
    Private dtObject, dtProject, dtTaskID As DataTable
    Private iLastCol As Integer

    '---Kiểm tra khoản mục theo chuẩn gồm 6 bước
    '--- Chuẩn Khoản mục b1: Khai báo biến

#Region "Biến khai báo cho khoản mục"

    Private Const SplitAna As Int16 = 0 ' Ghi nhận Khoản mục chứa ở Split nào
    Dim bUseAna As Boolean 'Kiểm tra có sử dụng Khoản mục không, để set thuộc tính Enabled nút Khoản mục 
 
#End Region

    Public Property AssignmentID() As String
        Get
            Return _assignmentID
        End Get
        Set(ByVal value As String)
            _assignmentID = value
        End Set
    End Property

    Private _auditCode As String
    Public WriteOnly Property sAuditCode() As String
        Set(ByVal value As String)
            _auditCode = value
        End Set
    End Property

    Dim bLoadFormState As Boolean = False
    Dim clsFilterDropdown As Lemon3.Controls.FilterDropdown
    Dim oFilterCombo As Lemon3.Controls.FilterCombo

	Private _FormState As EnumFormState
    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal value As EnumFormState)
	bLoadFormState = True
	LoadInfoGeneral()
            _FormState = value
            '--- Chuẩn Khoản mục b2: Lấy caption cho 10 khoản mục
            bUseAna = LoadTDBGridAnalysisCaption(D02, tdbg, COL_Ana01ID, SplitAna, , gbUnicode)
            '------------------------------------
            oFilterCombo = New Lemon3.Controls.FilterCombo
            oFilterCombo.CheckD91 = True 'Giá trị mặc định True: kiểm tra theo DxxFormat.LoadFormNotINV. Ngược lại luôn luôn Filter dạng mới (dùng cho Novaland)
            oFilterCombo.UseFilterCombo(tdbcAccountID)
            oFilterCombo.AddPairObject(tdbcObjectTypeID, tdbcObjectID) 'Đã bổ sung cột Loại ĐT
            oFilterCombo.UseFilterComboObjectID(True)
            LoadTDBCombo()
            LoadTDBDropDown()
            Select Case _FormState
                Case EnumFormState.FormAdd
                    btnSave.Enabled = True
                    btnNext.Enabled = False
                    LoadAdd()
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

    Private Sub D02F0101_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '12/10/2020, id 144622-Tài sản cố định_Lỗi chưa cảnh báo khi lưu
        If _FormState = EnumFormState.FormEdit Then
            If Not _savedOK Then
                If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
            End If
        ElseIf _FormState = EnumFormState.FormAdd Then
            If (txtAssignmentID.Text <> "" Or txtAssignmentName.Text <> "" Or tdbcAccountID.Text <> "") Then
                If Not _savedOK Then
                    If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub D02F0101_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
        If e.Control Then
            If e.KeyCode = Keys.D1 Or e.KeyCode = Keys.NumPad1 Then
                txtAssignmentID.Focus()
            ElseIf e.KeyCode = Keys.D2 Or e.KeyCode = Keys.NumPad2 Then
                chkSource.Focus()
            ElseIf e.KeyCode = Keys.D3 Or e.KeyCode = Keys.NumPad3 Then
                chkAna.Focus()
            End If
        End If
        If e.KeyCode = Keys.F11 Then
            HotKeyF11(Me, tdbg)
        End If
    End Sub

    Private Sub D02F0101_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
	If bLoadFormState = False Then FormState = _formState
        Me.Cursor = Cursors.WaitCursor
        Loadlanguage()
        SetBackColorObligatory()
        grpGuide.BringToFront()
        grp3.BringToFront()
        LoadTDBGridAnalysisCaption(D02, tdbg, COL_Ana01ID, SPLIT0, True, gbUnicode)
        LoadTDBGridGuide()
        iLastCol = CountCol(tdbg, SPLIT0)
        CheckIdTextBox(txtAssignmentID, 50)
        CheckIdTextBox(txtFormular1, txtFormular1.MaxLength, True)
        CheckIdTextBox(txtFormular2, txtFormular2.MaxLength, True)
        CheckIdTextBox(txtFormular3, txtFormular3.MaxLength, True)
        CheckIdTextBox(txtCondition1, txtCondition1.MaxLength, True)
        CheckIdTextBox(txtCondition2, txtCondition2.MaxLength, True)
        CheckIdTextBox(txtCondition3, txtCondition3.MaxLength, True)
        InputbyUnicode(Me, gbUnicode)

        clsFilterDropdown = New Lemon3.Controls.FilterDropdown()
        clsFilterDropdown.CheckD91 = True 'Giá trị mặc định True: kiểm tra theo DxxFormat.LoadFormNotINV. Ngược lại luôn luôn Filter dạng mới (dùng cho Novaland)
        clsFilterDropdown.UseFilterDropdown(tdbg, COL_PeriodID, COL_Ana01ID, COL_Ana02ID, COL_Ana03ID, COL_Ana04ID, COL_Ana05ID, COL_Ana06ID, COL_Ana07ID, COL_Ana08ID, COL_Ana09ID, COL_Ana10ID, COL_ProjectID, COL_TaskID)

        SetResolutionForm(Me)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Cap_nhat_tieu_thuc_phan_bo_khau_hao_-_D02F0101") & UnicodeCaption(gbUnicode) 'CËp nhËt ti£u th÷c ph¡n bå khÊu hao - D02F0101
        '================================================================ 
        lblAssignmentID.Text = rl3("Ma_phan_bo") 'Mã phân bổ
        lblAssignmentName.Text = rl3("Ten_phan_bo") 'Tên phân bổ
        lblAccountID.Text = rl3("TK_no") 'TK nợ
        lblObjectTypeID.Text = rl3("DT_phan_bo") 'ĐT phân bổ & 
        lblObjDesc.Text = "(" & rL3("Phan_bo_theo") & " " & rL3("Bo_phanU") & " " & rL3("chiu_phi_neu_chon_") & ")" '(Phân bổ theo Bộ phận chịu phí nếu chọn %)
        lblCondition.Text = rl3("Dieu_kien") 'Điều kiện
        lblFormular.Text = rl3("Cong_thuc_tinh_Gia_tri_phan_bo") 'Công thức tính Giá trị phân bổ
        lblFormular1.Text = rl3("Truong_hop") & Space(1) & "1" 'Trường hợp 1
        lblFormular2.Text = rl3("Truong_hop") & Space(1) & "2" 'Trường hợp 2
        lblFormular3.Text = rl3("Truong_hop") & Space(1) & "3" 'Trường hợp 3
        '================================================================ 
        btnSave.Text = rl3("_Luu") '&Lưu
        btnNext.Text = rl3("Nhap__tiep") 'Nhập &tiếp
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        btnGuide.Text = rl3("_Huong_dan") '&Hướng dẫn
        '================================================================ 
        chkDisabled.Text = rl3("Khong_su_dung") 'Không sử dụng
       
        chkSource.Text = "2. " & rl3("Phan_bo_theo_nguon") & "   " 'Phân bổ theo nguồn
        chkAna.Text = "3. " & rl3("Thong_ke_theo_khoan_muc_du_an") & "   " 'Thống kê theo khoản mục, dự án

        chkExtend.Text = rl3("Mo_rong") 'Mở rộng
        '================================================================ 
        optDifference.Text = rl3("Chenh_lech") 'Chênh lệch
        optNorm.Text = rl3("Dinh_muc") 'Định mức
        grp1.Text = "1. " & rl3("Thong_tin_chung") '1. Thông tin chung

        '================================================================ 
        tdbcObjectID.Columns("ObjectID").Caption = rl3("Ma") 'Mã
        tdbcObjectID.Columns("ObjectName").Caption = rl3("Ten") 'Tên
        tdbcObjectTypeID.Columns("ObjectTypeID").Caption = rl3("Ma") 'Mã
        tdbcObjectTypeID.Columns("ObjectTypeName").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcAccountID.Columns("AccountID").Caption = rl3("Ma") 'Mã
        tdbcAccountID.Columns("AccountName").Caption = rl3("Ten") 'Tên
        tdbcSourceID.Columns("SourceID").Caption = rl3("Ma") 'Mã
        tdbcSourceID.Columns("SourceName").Caption = rl3("Ten") 'Tên
        '================================================================ 
        tdbdAna10ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna10ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdPeriodID.Columns("PeriodID").Caption = rl3("Ky_san_xuat") 'Kỳ sản xuất
        tdbdPeriodID.Columns("PeriodName").Caption = rl3("Ten") 'Tên
        tdbdPeriodID.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbdAna09ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna09ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna08ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna08ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna07ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna07ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna01ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna01ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna06ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna06ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna02ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna02ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna05ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna05ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna03ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna03ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna04ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna04ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdProjectID.Columns("ProjectID").Caption = rl3("Ma") 'Mã dự án
        tdbdProjectID.Columns("ProjectName").Caption = rL3("Ten") 'Tên dự án
        tdbdTaskID.Columns("TaskID").Caption = rL3("Ma") 'Mã hạng mục
        tdbdTaskID.Columns("TaskName").Caption = rL3("Ten") 'Tên hạng mục
        '================================================================ 
        tdbg.Columns("PeriodID").Caption = rL3("Ky_san_xuat") 'Kỳ sản xuất
        tdbg.Columns("ProjectID").Caption = rL3("Du_an") 'Dự án
        tdbg.Columns("TaskID").Caption = rL3("Hang_muc") 'Hạng mục

        'ID : 245825 : -	Bổ sung check Ưu tiên khoản mục của phiếu hình thành
        '================================================================ 
        chkIsKCodeByTrans.Text = rL3("Uu_tien_khoan_muc_cua_phieu_hinh_thanh") 'Ưu tiên khoản mục của phiếu hình thành
        '================================================================ 
        chkIsManagement.Text = rL3("Bo_phan_quan_ly") 'Bộ phận quản lý
        chkIsReceive.Text = rL3("Bo_phan_tiep_nhan") 'Bộ phận tiếp nhận
        '================================================================ 



    End Sub

    Private Sub SetBackColorObligatory()
        txtAssignmentID.BackColor = COLOR_BACKCOLOROBLIGATORY
        txtAssignmentName.BackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcAccountID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub

    Private Sub LoadTDBCombo()
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, gbUnicode)

        Dim sSQL As String = ""
        'Load tdbcAccountID
        sSQL = "Select AccountID," & IIf(geLanguage = EnumLanguage.Vietnamese, " AccountName", " AccountName01").ToString() & sUnicode & " as AccountName " & vbCrLf
        sSQL &= "From D90T0001 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where Disabled = 0 And AccountStatus = 0 And OffAccount = 0" & vbCrLf
        sSQL &= "Order By AccountID   "
        LoadDataSource(tdbcAccountID, sSQL, gbUnicode)

        'Load tdbcObjectTypeID
        LoadObjectTypeIDAll(tdbcObjectTypeID, gbUnicode)

        'Load tdbcObjectID
        'sSQL = "Select 0 as DisplayOrder,'%' as ObjectID, " & sLanguage & " as ObjectName, '' As  VATNo, '%' As ObjectTypeID " & vbCrLf
        'sSQL &= "Union All" & vbCrLf
        'sSQL &= "Select 1 as DisplayOrder,ObjectID, ObjectName" & sUnicode & " As ObjectName, VATNo, ObjectTypeID " & vbCrLf
        'sSQL &= "From Object WITH(NOLOCK)  " & vbCrLf
        'sSQL &= "Where Disabled = 0 " & vbCrLf
        'sSQL &= "Order By DisplayOrder,ObjectID "
        'dtObject = ReturnDataTable(sSQL)

        dtObject = oFilterCombo.LoadObjectID(Lemon3.Data.eUnionAll.All)
        oFilterCombo.LoadtdbcObjectID(tdbcObjectID, dtObject, "")

        'Load tdbcSourceID
        sSQL = "Select SourceID, SourceName" & sUnicode & " As SourceName From D02T0013 WITH(NOLOCK) Where Disabled = 0"
        LoadDataSource(tdbcSourceID, sSQL, gbUnicode)
    End Sub

    Private Sub LoadtdbcObjectID(ByVal ID As String)
        If tdbcObjectTypeID.Text = "%" Then
            'LoadDataSource(tdbcObjectID, dtObject, gbUnicode)
            oFilterCombo.LoadtdbcObjectID(tdbcObjectID, dtObject, "")
        Else
            'LoadDataSource(tdbcObjectID, ReturnTableFilter(dtObject, "ObjectTypeID=" & SQLString(ID), True), gbUnicode)
            oFilterCombo.LoadtdbcObjectID(tdbcObjectID, dtObject, ID)
        End If
    End Sub

    Private Sub LoadTDBDropDown()
        Dim sSQL As String = ""
        'Load tdbdPeriodID
        sSQL = "Select PeriodID, WorkOrderNo as PeriodName, Note" & UnicodeJoin(gbUnicode) & " As Description  " & vbCrLf
        sSQL &= "From D08N0100 (" & SQLString(gsDivisionID) & "," & SQLNumber(giTranMonth) & "," & SQLNumber(giTranYear) & ", 2) " & vbCrLf
        sSQL &= "Where  (  DAGroupID = ''  " & vbCrLf
        sSQL &= "Or    DAGroupID  In   ( Select  DAGroupID  " & vbCrLf
        sSQL &= "From   LEMONSYS.DBO.D00V0080" & vbCrLf
        sSQL &= "Where UserID = " & SQLString(gsUserID) & ")" & vbCrLf
        sSQL &= "Or " & SQLString(gsUserID) & " =  'LEMONADMIN' )  " & vbCrLf
        sSQL &= "Order By PeriodID"

        LoadDataSource(tdbdPeriodID, sSQL, gbUnicode)
        '--- Chuẩn Khoản mục b3: Load 10 khoản mục
        LoadTDBDropDownAna(tdbdAna01ID, tdbdAna02ID, tdbdAna03ID, tdbdAna04ID, tdbdAna05ID, tdbdAna06ID, tdbdAna07ID, tdbdAna08ID, tdbdAna09ID, tdbdAna10ID, tdbg, COL_Ana01ID, gbUnicode)
        '------------------------------------------
        'Thêm ngày 25/9/2012 theo incident 51216
        LoadProject(tdbdProjectID, dtProject, True) '14/1/2020, Phạm Thị Mỹ Tiên:id 131799-Sửa câu đổ nguồn dropdown dự án, hạng mục

        'ID 78571 13/08/2015
        LoadTask(tdbdTaskID, dtTaskID, , True) '14/1/2020, Phạm Thị Mỹ Tiên:id 131799-Sửa câu đổ nguồn dropdown dự án, hạng mục
    End Sub

    Private Sub LoadTaskID(ByVal sProjectID As String)
        LoadDataSource(tdbdTaskID, ReturnTableFilter(dtTaskID, "ProjectID ='%' Or ProjectID=" & SQLString(sProjectID), True), True) '14/1/2020, Phạm Thị Mỹ Tiên:id 131799-Sửa câu đổ nguồn dropdown dự án, hạng mục
    End Sub

    Private Sub LoadTDBGridGuide()
        Dim sSQL As String
        sSQL = "Select VariableID, Description" & UnicodeJoin(gbUnicode) & " As Description" & vbCrLf
        sSQL &= "From D02V0055" & vbCrLf
        sSQL &= "Where Language = " & SQLString(geLanguage) & " Order By OrderNum"
        LoadDataSource(tdbgGuide, sSQL, gbUnicode)
    End Sub

    Private Sub LoadForm()
        Dim sSQL As String
        sSQL = "Select  AssignmentID, AssignmentName, AssignmentNameU, Disabled, DebitAccountID as AccountID, DebitObjectTypeID as ObjectTypeID, DebitObjectID as ObjectID, " & vbCrLf
        sSQL &= "SourceID, PeriodID, Ana01ID,  Ana02ID, Ana03ID, Ana04ID, Ana05ID, Ana06ID, Ana07ID, Ana08ID, " & vbCrLf
        sSQL &= "Ana09ID, Ana10ID,ProjectID, Extend, Formular1, Formular2, Formular3, Condition1, Condition2, Condition3, TaskID, IsKCodeByTrans, IsReceive, IsManagement " & vbCrLf
        sSQL &= "From D02T0002 WITH(NOLOCK)" & vbCrLf
        sSQL &= "Where AssignmentID = " & SQLString(_assignmentID)
        Dim dt As DataTable = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            With dt.Rows(0)
                txtAssignmentID.Text = .Item("AssignmentID").ToString
                txtAssignmentName.Text = .Item("AssignmentName" & UnicodeJoin(gbUnicode)).ToString
                chkDisabled.Checked = CBool(.Item("Disabled"))
                tdbcAccountID.Text = .Item("AccountID").ToString
                txtAccountName.Text = tdbcAccountID.Columns("AccountName").Text
                tdbcObjectTypeID.Text = .Item("ObjectTypeID").ToString
                tdbcObjectID.SelectedValue = .Item("ObjectID").ToString
                chkIsKCodeByTrans.Checked = CBool(.Item("IsKCodeByTrans"))
                chkIsReceive.Checked = CBool(.Item("IsReceive"))
                chkIsManagement.Checked = CBool(.Item("IsManagement"))
                If IsDBNull(.Item("SourceID")) Or .Item("SourceID").ToString = "" Then
                    tdbcSourceID.Enabled = False
                    tdbcSourceID.SelectedValue = ""
                Else
                    chkSource.Checked = True
                    tdbcSourceID.Text = .Item("SourceID").ToString
                    txtSourceName.Text = tdbcSourceID.Columns("SourceName").Text
                End If
                txtFormular1.Text = .Item("Formular1").ToString
                txtFormular2.Text = .Item("Formular2").ToString
                txtFormular3.Text = .Item("Formular3").ToString
                txtCondition1.Text = .Item("Condition1").ToString
                txtCondition2.Text = .Item("Condition2").ToString
                txtCondition3.Text = .Item("Condition3").ToString
                chkExtend.Checked = CBool(.Item("Extend"))

                Select Case CInt(.Item("Extend"))
                    Case 0
                        grp3.Visible = False
                    Case 1
                        chkExtend.Enabled = False
                        grp3.Visible = True
                        optNorm.Checked = True
                        optNorm_Click(Nothing, Nothing)
                    Case 2
                        chkExtend.Enabled = False
                        grp3.Visible = True
                        optDifference.Checked = True
                        optDifference_Click(Nothing, Nothing)
                End Select

                If (Not IsDBNull(.Item("PeriodID").ToString) And .Item("PeriodID").ToString <> "") OrElse (Not IsDBNull(.Item("Ana01ID").ToString) And .Item("Ana01ID").ToString <> "") OrElse .Item("Ana02ID").ToString <> "" OrElse .Item("Ana03ID").ToString <> "" OrElse .Item("Ana04ID").ToString <> "" OrElse .Item("Ana05ID").ToString <> "" OrElse .Item("Ana06ID").ToString <> "" OrElse .Item("Ana07ID").ToString <> "" OrElse .Item("Ana08ID").ToString <> "" OrElse .Item("Ana09ID").ToString <> "" OrElse .Item("Ana10ID").ToString <> "" OrElse .Item("ProjectID").ToString <> "" Then
                    chkAna.Checked = True
                Else
                    chkAna.Checked = False
                End If
                chkAna_Click(Nothing, Nothing)
            End With
        End If
        LoadDataSource(tdbg, dt, gbUnicode)

    End Sub

    Private Sub LoadAdd()
        _assignmentID = ""
        chkIsKCodeByTrans.Checked = True
        chkDisabled.Visible = False
        chkIsReceive.Enabled = False
        chkIsManagement.Enabled = False
        chkSource_Click(Nothing, Nothing)
        chkAna_Click(Nothing, Nothing)
        LoadForm()
    End Sub

    Private Sub LoadEdit()
        txtAssignmentID.Enabled = False
        chkDisabled.Visible = True
        LoadForm()
        txtAssignmentName.Focus()
        EnabledControl()
    End Sub

#Region "Events tdbcAccountID with txtAccountName"

    'Private Sub tdbcAccountID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAccountID.Close
    '    If tdbcAccountID.FindStringExact(tdbcAccountID.Text) = -1 Then
    '        tdbcAccountID.Text = ""
    '        txtAccountName.Text = ""
    '    End If
    'End Sub

    Private Sub tdbcAccountID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAccountID.SelectedValueChanged
        txtAccountName.Text = tdbcAccountID.Columns(1).Value.ToString
    End Sub

    Private Sub tdbcAccountID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAccountID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcAccountID.Text = ""
            txtAccountName.Text = ""
        End If
    End Sub
    Private Sub tdbcAccountID_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAccountID.Validated
        If tdbcAccountID.FindStringExact(tdbcAccountID.Text) = -1 Then
            tdbcAccountID.Text = ""
            txtAccountName.Text = ""
        End If
    End Sub


#End Region

#Region "Events tdbcObjectTypeID load tdbcObjectID with txtObjectName"

    Private Sub tdbcObjectTypeID_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.Close
        If tdbcObjectTypeID.FindStringExact(tdbcObjectTypeID.Text) = -1 Then tdbcObjectTypeID.Text = ""
    End Sub

    Private Sub tdbcObjectTypeID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.SelectedValueChanged
        If Not (tdbcObjectTypeID.Tag Is Nothing OrElse tdbcObjectTypeID.Tag.ToString = "") Then
            tdbcObjectTypeID.Tag = ""
            Exit Sub
        End If
        'Bổ sung đoạn code sau:
        Dim sObjectID As String = ""
        If bLoadObectID = False Then sObjectID = ReturnValueC1Combo(tdbcObjectID)
        Dim sObjectTypeID As String = ReturnValueC1Combo(tdbcObjectTypeID)
        tdbcObjectID.Splits(0).DisplayColumns("ObjectTypeID").Visible = (sObjectTypeID = "" Or sObjectTypeID = "-1" Or sObjectTypeID = "%") 'Xử lý cho dạng cũ
        oFilterCombo.LoadtdbcObjectID(tdbcObjectID, dtObject, sObjectTypeID)
        If bLoadObectID = False Then tdbcObjectID.SelectedValue = sObjectID : Exit Sub
        tdbcObjectID.Text = ""
        txtObjectName.Text = ""

        'If tdbcObjectTypeID.SelectedValue Is Nothing Then
        '    LoadtdbcObjectID("-1")
        '    Exit Sub
        'End If
        'LoadtdbcObjectID(tdbcObjectTypeID.SelectedValue.ToString())
        'tdbcObjectID.Text = ""
        'txtObjectName.Text = ""
    End Sub

    Private Sub tdbcObjectTypeID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcObjectTypeID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcObjectTypeID.Text = ""
    End Sub

    'Private Sub tdbcObjectID_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectID.Close
    '    If tdbcObjectID.FindStringExact(tdbcObjectID.Text) = -1 Then
    '        tdbcObjectID.Text = ""
    '        txtObjectName.Text = ""
    '    End If
    'End Sub

    Private Sub tdbcObjectID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectID.SelectedValueChanged
        txtObjectName.Text = tdbcObjectID.Columns(1).Value.ToString()
        If ReturnValueC1Combo(tdbcObjectID) = "%" Then
            chkIsReceive.Enabled = True
            chkIsManagement.Enabled = True
            chkIsReceive.Checked = True
        Else
            chkIsReceive.Enabled = False
            chkIsManagement.Enabled = False
            chkIsReceive.Checked = False
            chkIsManagement.Checked = False
        End If
    End Sub

    Private Sub tdbcObjectID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcObjectID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcObjectID.Text = ""
            txtObjectName.Text = ""
        End If
    End Sub

#End Region

#Region "Events tdbcSourceID with txtSourceName"

    Private Sub tdbcSourceID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcSourceID.Close
        If tdbcSourceID.FindStringExact(tdbcSourceID.Text) = -1 Then
            tdbcSourceID.Text = ""
            txtSourceName.Text = ""
        End If
    End Sub

    Private Sub tdbcSourceID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcSourceID.SelectedValueChanged
        txtSourceName.Text = tdbcSourceID.Columns("SourceName").Value.ToString
    End Sub

    Private Sub tdbcSourceID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcSourceID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcSourceID.Text = ""
            txtSourceName.Text = ""
        End If
    End Sub

#End Region

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNext.Click
        btnNext.Enabled = False
        btnSave.Enabled = True
        txtAssignmentID.Text = ""
        txtAssignmentName.Text = ""
        tdbcAccountID.Text = ""
        tdbcObjectTypeID.SelectedValue = ""
        tdbcObjectID.SelectedValue = ""
        txtObjectName.Text = ""

        chkSource.Checked = False
        tdbcSourceID.Text = ""
        txtSourceName.Text = ""
        txtFormular1.Text = ""
        txtFormular2.Text = ""
        txtFormular3.Text = ""
        txtCondition1.Text = ""
        txtCondition2.Text = ""
        txtCondition3.Text = ""
        txtAccountName.Text = ""
        optNorm.Checked = True

        chkAna.Checked = False
        chkExtend.Checked = False
        grp3.Visible = False
        grpGuide.Visible = False

        LoadAdd()
        txtAssignmentID.Focus()
        dtCheckAccount = Nothing
    End Sub

    Private Sub btnGuide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGuide.Click
        If grpGuide.Visible = False Then
            grpGuide.Visible = True
        Else
            grpGuide.Visible = False
        End If
    End Sub

    Private Sub btnCloseFrameGuide_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCloseFrameGuide.Click
        grpGuide.Visible = False
    End Sub

    Private Sub btnCloseFrame_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCloseFrame.Click
        grp3.Visible = False
        chkExtend.Checked = False
    End Sub

    Private Sub chkExtend_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkExtend.Click
        If chkExtend.Checked Then
            grp3.Visible = True
        Else
            grp3.Visible = False
        End If
    End Sub

    Private Sub chkAna_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAna.Click
        tdbg.Enabled = chkAna.Checked
        If chkAna.Checked = False And tdbg.RowCount > 0 Then
            tdbg.Delete()
        End If
    End Sub

    Private Sub chkSource_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkSource.Click
        tdbcSourceID.Enabled = chkSource.Checked
        If chkSource.Checked = False Then
            tdbcSourceID.Text = ""
            txtSourceName.Text = ""
        End If
    End Sub

#Region "Grid"

    Private Sub tdbg_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg.BeforeColUpdate
        Select Case e.ColIndex
            Case COL_PeriodID
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg.Columns(COL_PeriodID).Text <> tdbdPeriodID.Columns("PeriodID").Text Then
                    tdbg.Columns(COL_PeriodID).Text = ""
                End If
                '--- Chuẩn Khoản mục b5: Kiểm tra Khoản mục lúc nhập liệu
                '---------------------------------------------
            Case COL_Ana01ID
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg.Columns(COL_Ana01ID).Text <> tdbdAna01ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(0) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana01ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana01ID).Text.Length > giArrAnaLength(0) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana01ID).Text = ""
                        End If
                    End If
                End If

            Case COL_Ana02ID
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg.Columns(COL_Ana02ID).Text <> tdbdAna02ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(1) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana02ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana02ID).Text.Length > giArrAnaLength(1) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana02ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana03ID
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg.Columns(COL_Ana03ID).Text <> tdbdAna03ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(2) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana03ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana03ID).Text.Length > giArrAnaLength(2) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana03ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana04ID
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg.Columns(COL_Ana04ID).Text <> tdbdAna04ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(3) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana04ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana04ID).Text.Length > giArrAnaLength(3) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana04ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana05ID
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg.Columns(COL_Ana05ID).Text <> tdbdAna05ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(4) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana05ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana05ID).Text.Length > giArrAnaLength(4) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana05ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana06ID
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg.Columns(COL_Ana06ID).Text <> tdbdAna06ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(5) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana06ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana06ID).Text.Length > giArrAnaLength(5) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana06ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana07ID
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg.Columns(COL_Ana07ID).Text <> tdbdAna07ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(6) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana07ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana07ID).Text.Length > giArrAnaLength(6) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana07ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana08ID
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg.Columns(COL_Ana08ID).Text <> tdbdAna08ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(7) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana08ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana08ID).Text.Length > giArrAnaLength(7) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana08ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana09ID
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg.Columns(COL_Ana09ID).Text <> tdbdAna09ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(8) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana09ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana09ID).Text.Length > giArrAnaLength(8) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana09ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana10ID
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg.Columns(COL_Ana10ID).Text <> tdbdAna10ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(9) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana10ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana10ID).Text.Length > giArrAnaLength(9) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana10ID).Text = ""
                        End If
                    End If
                End If
                '---------------------------------------------
            Case COL_ProjectID
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg.Columns(COL_ProjectID).Text <> tdbdProjectID.Columns("ProjectID").Text Then
                    tdbg.Columns(COL_ProjectID).Text = ""
                End If
            Case COL_TaskID
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg.Columns(COL_TaskID).Text <> tdbdTaskID.Columns("TaskID").Text Then
                    tdbg.Columns(COL_TaskID).Text = ""
                End If
        End Select
    End Sub

    Private Sub tdbg_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.AfterColUpdate
        '--- Gán giá trị cột sau khi tính toán và giá trị phụ thuộc từ Dropdown
        Select Case e.ColIndex
            'Case COL_PeriodID
            '    Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(e.Column.DataColumn.DataField)
            '    If tdbd Is Nothing Then Exit Select
            '    If clsFilterDropdown.IsNewFilter Then
            '        Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg, e, tdbd)
            '        AfterColUpdate(e.ColIndex, dr)
            '        Exit Sub
            '    Else ' Nhập liệu dạng cũ (xổ dropdown)
            '        Dim row As DataRow = ReturnDataRow(tdbd, tdbd.DisplayMember & "=" & SQLString(tdbg.Columns(e.ColIndex).Text))
            '        AfterColUpdate(e.ColIndex, row)
            '    End If
            Case COL_PeriodID, COL_ProjectID, COL_Ana01ID To COL_Ana10ID, COL_TaskID
                Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(e.Column.DataColumn.DataField)
                If tdbd Is Nothing Then Exit Select
                If DxxFormat.LoadFormNotINV = 1 Then
                    Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg, e, tdbd, False)
                    AfterColUpdate(e.ColIndex, dr)
                    Exit Sub
                Else ' Nhập liệu dạng cũ (xổ dropdown)
                    Dim row As DataRow = ReturnDataRow(tdbd, tdbd.DisplayMember & "=" & SQLString(tdbg.Columns(e.ColIndex).Text))
                    AfterColUpdate(e.ColIndex, row)
                End If
                'Case COL_ProjectID
        End Select
    End Sub

    Private Sub AfterColUpdate(ByVal iCol As Integer, ByVal dr() As DataRow)
        Dim iRow As Integer = tdbg.Row
        If dr Is Nothing OrElse dr.Length = 0 Then
            Dim row As DataRow = Nothing
            AfterColUpdate(iCol, row)
        ElseIf dr.Length = 1 Then
            If tdbg.Bookmark <> tdbg.Row AndAlso tdbg.RowCount = tdbg.Row Then 'Đang đứng dòng mới
                Dim dtGrid As DataTable = CType(tdbg.DataSource, DataTable)
                Dim dr1 As DataRow = dtGrid.NewRow
                dtGrid.Rows.InsertAt(dr1, tdbg.Row)
                tdbg.Bookmark = tdbg.Row
            End If
            AfterColUpdate(iCol, dr(0))
        Else
            For Each row As DataRow In dr
                tdbg.Bookmark = iRow
                tdbg.Row = iRow
                AfterColUpdate(iCol, row)
                tdbg.UpdateData()
                iRow += 1
            Next
            tdbg.Focus()
        End If
    End Sub

    Private Sub AfterColUpdate(ByVal iCol As Integer, ByVal dr As DataRow)
        'Gán lại các giá trị phụ thuộc vào Dropdown
        Select Case iCol
            Case COL_PeriodID
                If dr Is Nothing OrElse dr.Item("PeriodID").ToString = "" Then
                    'Gắn rỗng các cột liên quan
                    tdbg.Columns(COL_PeriodID).Text = ""
                    Exit Select
                End If
                tdbg.Columns(COL_PeriodID).Text = dr.Item("PeriodID").ToString
            Case COL_Ana01ID To COL_Ana10ID
                If dr Is Nothing OrElse dr.Item("AnaID").ToString = "" Then
                    'Gắn rỗng các cột liên quan
                    tdbg.Columns(iCol).Text = ""
                Else
                    tdbg.Columns(iCol).Text = dr.Item("AnaID").ToString
                End If
            Case COL_ProjectID
                If dr Is Nothing OrElse dr.Item("ProjectID").ToString = "" Then
                    'Gắn rỗng các cột liên quan
                    tdbg.Columns(iCol).Text = ""
                    Exit Select
                End If
                tdbg.Columns(iCol).Text = dr.Item("ProjectID").ToString

            Case COL_TaskID
                If dr Is Nothing OrElse dr.Item("TaskID").ToString = "" Then
                    'Gắn rỗng các cột liên quan
                    tdbg.Columns(iCol).Text = ""
                    Exit Select
                End If
                tdbg.Columns(iCol).Text = dr.Item("TaskID").ToString
        End Select

SumFooter:
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        If CheckKeydownFilterDropdown(tdbg, e, False) Then
            Select Case tdbg.Col
                Case COL_PeriodID, COL_ProjectID, COL_Ana01ID To COL_Ana10ID, COL_TaskID
                    Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg.Columns(tdbg.Col).DataField)
                    If tdbd Is Nothing Then Exit Select
                    Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg, e, tdbd, False)
                    AfterColUpdate(tdbg.Col, dr)
                    Exit Sub
            End Select
        End If

        If e.KeyCode = Keys.Enter Then
            If tdbg.Col = iLastCol Then
                HotKeyEnterGrid(tdbg, COL_PeriodID, e)
            End If
        End If
        'If tdbg.Enabled And e.KeyCode = Keys.F2 Then
        '    Select Case tdbg.Col
        '        Case COL_PeriodID
        '            HotKeyF2("85", tdbg.Col)
        '        Case COL_Ana01ID To COL_Ana10ID
        '            HotKeyF2("93", tdbg.Col)
        '        Case COL_ProjectID
        '            HotKeyF2("97", tdbg.Col)
        '    End Select
        'End If
    End Sub


    'Private Sub HotKeyF2(ByVal sInListID As String, ByVal iCol As Integer)
    '    Try
    '        Dim arrPro() As StructureProperties = Nothing
    '        SetProperties(arrPro, "InListID", sInListID)
    '        Dim frm As Form = CallFormShowDialog("D91D0240", "D91F6010", arrPro)
    '        Dim sKey As String = GetProperties(frm, "Output01").ToString
    '        If sKey <> "" Then
    '            'Load dữ liệu
    '            tdbg.Columns(iCol).Text = sKey
    '            tdbg.UpdateData()

    '        End If
    '    Catch ex As Exception
    '        D99C0008.MsgL3(ex.Message)
    '    End Try
    'End Sub

    Private Sub tdbg_RowColChange(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.RowColChangeEventArgs) Handles tdbg.RowColChange
  If e IsNot Nothing AndAlso e.LastRow = -1 Then Exit Sub
        If tdbg.Row > 0 Then
            tdbg.AllowAddNew = False
        Else
            tdbg.AllowAddNew = True
        End If

        Select Case tdbg.Col
            Case COL_TaskID
                LoadTaskID(tdbg(tdbg.Row, COL_ProjectID).ToString)
        End Select
    End Sub

#End Region

    Private Sub SetTagTDBG()
        For i As Integer = 0 To tdbg.Columns.Count - 1
            If tdbg.Splits(0).DisplayColumns(i).Visible Then
                tdbg.Columns(i).Tag = "1"
            End If
        Next
    End Sub

    Dim dtCheckAccount As DataTable 'table kiểm tra tài khoản
    Private Function AllowSave() As Boolean
        Dim sKCode As String = ""
        If txtAssignmentID.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rL3("Ma_phan_bo"))
            txtAssignmentID.Focus()
            Return False
        End If
        If txtAssignmentName.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rL3("Ten_phan_bo"))
            txtAssignmentName.Focus()
            Return False
        End If
        If tdbcAccountID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("TK_no"))
            tdbcAccountID.Focus()
            Return False
        End If

        If _FormState = EnumFormState.FormAdd Then
            If IsExistKey("D02T0002", "AssignmentID", txtAssignmentID.Text) Then
                D99C0008.MsgDuplicatePKey()
                txtAssignmentID.Focus()
                Return False
            End If
        End If
        Dim bEmpty As Boolean = False

        'ID 83828 10.05.2016 Bo sung bat buoc nhap Du an,hang muc theo tai khoan
        If chkAna.Checked Then
            SetTagTDBG() 'Gán Tag cho các cot tren luoi
            If tdbg.RowCount <= 0 Then
                D99C0008.MsgNoDataInGrid()
                tdbg.Focus()
                Return False
            End If
            For i As Integer = 0 To tdbg.RowCount - 1
                'If tdbg(i, COL_PeriodID).ToString <> "" OrElse tdbg(i, COL_Ana01ID).ToString <> "" OrElse tdbg(i, COL_Ana02ID).ToString <> "" OrElse tdbg(i, COL_Ana03ID).ToString <> "" OrElse tdbg(i, COL_Ana04ID).ToString <> "" OrElse tdbg(i, COL_Ana05ID).ToString <> "" OrElse tdbg(i, COL_Ana06ID).ToString <> "" OrElse tdbg(i, COL_Ana07ID).ToString <> "" OrElse tdbg(i, COL_Ana08ID).ToString <> "" OrElse tdbg(i, COL_Ana09ID).ToString <> "" OrElse tdbg(i, COL_Ana10ID).ToString <> "" OrElse tdbg(i, COL_ProjectID).ToString <> "" Then
                '    bEmpty = True
                '    Exit For
                'End If
                ' Chuẩn kiểm tra theo thiết lập Tài khoản B3: Kiểm tra trước khi lưu ' Duyệt theo bảng dữ liệu do store D91P9310 để kiểm tra cho từng cột Kcode
                If Not AllowKCode(dtCheckAccount, tdbg, i, Nothing, sKCode, ReturnValueC1Combo(tdbcAccountID)) Then
                    tdbg.Focus()
                    tdbg.SplitIndex = 0
                    tdbg.Col = IndexOfColumn(tdbg, sKCode)
                    tdbg.Bookmark = i
                    Return False
                End If
            Next
            'If bEmpty = False Then
            '    D99C0008.MsgNoDataInGrid()
            '    tdbg.Focus()
            '    Return False
            'End If
        End If

        If chkIsReceive.Enabled And chkIsManagement.Enabled Then
            If Not chkIsReceive.Checked And Not chkIsManagement.Checked Then
                D99C0008.Msg(rL3("Ban_phai_chon") & Space(1) & rL3("Bo_phanU") & Space(1) & rL3("chiu_phi"))
                Return False
            End If
        End If
        Return True
    End Function

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        tdbg.UpdateData()

        If Not AllowSave() Then Exit Sub
        btnSave.Enabled = False
        btnClose.Enabled = False
        _savedOK = False

        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder
        Select Case _FormState
            Case EnumFormState.FormAdd
                sSQL.Append(SQLInsertD02T0002)
            Case EnumFormState.FormEdit
                sSQL.Append(SQLUpdateD02T0002)
        End Select

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            _savedOK = True
            btnClose.Enabled = True
            Select Case _FormState
                Case EnumFormState.FormAdd
                    _assignmentID = txtAssignmentID.Text
                    btnNext.Enabled = True
                    btnNext.Focus()
                Case EnumFormState.FormEdit
                    'ExecuteAuditLog(_auditCode, "02", txtAssignmentID.Text, tdbcAccountID.Text, tdbcObjectTypeID.Text, tdbcSourceID.Text)
                    Lemon3.D91.RunAuditLog("02", _auditCode, "02", txtAssignmentID.Text, tdbcAccountID.Text, tdbcObjectTypeID.Text, tdbcSourceID.Text)
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
    '# Title: SQLStoreD02P1000
    '# Created User: 
    '# Created Date: 27/11/2007 10:47:31
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1000() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P1000 "
        sSQL &= SQLString(_assignmentID) & COMMA 'AssignmentID, varchar[20], NOT NULL
        sSQL &= SQLNumber(geLanguage) 'Language, tinyint, NOT NULL
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T0002
    '# Created User: 
    '# Created Date: 27/11/2007 02:05:23
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T0002() As StringBuilder
        Dim nExtend As Byte
        If Not chkExtend.Checked Then
            nExtend = 0
        ElseIf optNorm.Checked Then
            nExtend = 1
        Else
            nExtend = 2
        End If
        Dim sSQL As New StringBuilder
        If tdbg.RowCount > 0 Then
            sSQL.Append("Insert Into D02T0002(")
            sSQL.Append("AssignmentID, AssignmentNameU, DebitAccountID, DebitObjectTypeID, DebitObjectID, ")
            sSQL.Append("Disabled, SourceID, CreateUserID, CreateDate, LastModifyUserID, ")
            sSQL.Append("LastModifyDate, Ana01ID, Ana02ID, Ana03ID, Ana04ID, ")
            sSQL.Append("Ana05ID, Ana06ID, Ana07ID, Ana08ID, Ana09ID, ")
            sSQL.Append("Ana10ID,ProjectID, Extend, Formular1, Formular2, Formular3, ")
            sSQL.Append("Condition1, Condition2, Condition3, PeriodID, TaskID, IsKCodeByTrans, IsReceive, IsManagement")
            sSQL.Append(") Values(")
            sSQL.Append(SQLString(txtAssignmentID.Text) & COMMA) 'AssignmentID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLStringUnicode(txtAssignmentName.Text, gbUnicode, True) & COMMA) 'AssignmentNameU, varchar[50], NULL
            sSQL.Append(SQLString(tdbcAccountID.Text) & COMMA) 'DebitAccountID, varchar[20], NULL
            sSQL.Append(SQLString(tdbcObjectTypeID.Text) & COMMA) 'DebitObjectTypeID, varchar[20], NULL
            sSQL.Append(SQLString(tdbcObjectID.Text) & COMMA) 'DebitObjectID, varchar[20], NULL
            sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, tinyint, NOT NULL
            sSQL.Append(SQLString(tdbcSourceID.Text) & COMMA) 'SourceID, varchar[20], NULL
            sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL
            sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL
            sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NULL
            sSQL.Append(SQLString(tdbg(0, COL_Ana01ID)) & COMMA) 'Ana01ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(0, COL_Ana02ID)) & COMMA) 'Ana02ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(0, COL_Ana03ID)) & COMMA) 'Ana03ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(0, COL_Ana04ID)) & COMMA) 'Ana04ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(0, COL_Ana05ID)) & COMMA) 'Ana05ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(0, COL_Ana06ID)) & COMMA) 'Ana06ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(0, COL_Ana07ID)) & COMMA) 'Ana07ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(0, COL_Ana08ID)) & COMMA) 'Ana08ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(0, COL_Ana09ID)) & COMMA) 'Ana09ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(0, COL_Ana10ID)) & COMMA) 'Ana10ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(0, COL_ProjectID)) & COMMA) 'ProjectID, varchar[50], NULL
            sSQL.Append(SQLNumber(nExtend) & COMMA) 'Extend, tinyint, NOT NULL
            sSQL.Append(SQLString(txtFormular1.Text) & COMMA) 'Formular1, varchar[100], NOT NULL
            sSQL.Append(SQLString(txtFormular2.Text) & COMMA) 'Formular2, varchar[100], NOT NULL
            sSQL.Append(SQLString(txtFormular3.Text) & COMMA) 'Formular3, varchar[100], NOT NULL
            sSQL.Append(SQLString(txtCondition1.Text) & COMMA) 'Condition1, varchar[100], NOT NULL
            sSQL.Append(SQLString(txtCondition2.Text) & COMMA) 'Condition2, varchar[100], NULL
            sSQL.Append(SQLString(txtCondition3.Text) & COMMA) 'Condition3, varchar[100], NULL
            sSQL.Append(SQLString(tdbg(0, COL_PeriodID).ToString) & COMMA) 'PeriodID, varchar[20], NOT NULL
            sSQL.Append(SQLString(tdbg(0, COL_TaskID).ToString) & COMMA) 'TaskID, varchar[20], NOT NULL
            sSQL.Append(SQLNumber(chkIsKCodeByTrans.Checked) & COMMA) 'IsKCodeByTrans , tinyint, NOT NULL
            sSQL.Append(SQLNumber(chkIsReceive.Checked) & COMMA) 'IsReceive
            sSQL.Append(SQLNumber(chkIsManagement.Checked)) 'IsManagement

        Else
            sSQL.Append("Insert Into D02T0002(")
            sSQL.Append("AssignmentID, AssignmentNameU, DebitAccountID, DebitObjectTypeID, DebitObjectID, ")
            sSQL.Append("Disabled, SourceID, CreateUserID, CreateDate, LastModifyUserID, ")
            sSQL.Append("LastModifyDate, Extend, Formular1, Formular2, Formular3, ")
            sSQL.Append("Condition1, Condition2, Condition3, IsKCodeByTrans , IsReceive, IsManagement")
            sSQL.Append(") Values(")
            sSQL.Append(SQLString(txtAssignmentID.Text) & COMMA) 'AssignmentID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLStringUnicode(txtAssignmentName.Text, gbUnicode, True) & COMMA) 'AssignmentNameU, varchar[50], NULL
            sSQL.Append(SQLString(tdbcAccountID.Text) & COMMA) 'DebitAccountID, varchar[20], NULL
            sSQL.Append(SQLString(tdbcObjectTypeID.Text) & COMMA) 'DebitObjectTypeID, varchar[20], NULL
            sSQL.Append(SQLString(tdbcObjectID.Text) & COMMA) 'DebitObjectID, varchar[20], NULL
            sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, tinyint, NOT NULL
            sSQL.Append(SQLString(tdbcSourceID.Text) & COMMA) 'SourceID, varchar[20], NULL
            sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL
            sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL
            sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NULL
            sSQL.Append(SQLNumber(nExtend) & COMMA) 'Extend, tinyint, NOT NULL
            sSQL.Append(SQLString(txtFormular1.Text) & COMMA) 'Formular1, varchar[100], NOT NULL
            sSQL.Append(SQLString(txtFormular2.Text) & COMMA) 'Formular2, varchar[100], NOT NULL
            sSQL.Append(SQLString(txtFormular3.Text) & COMMA) 'Formular3, varchar[100], NOT NULL
            sSQL.Append(SQLString(txtCondition1.Text) & COMMA) 'Condition1, varchar[100], NOT NULL
            sSQL.Append(SQLString(txtCondition2.Text) & COMMA) 'Condition2, varchar[100], NULL
            sSQL.Append(SQLString(txtCondition3.Text) & COMMA) 'Condition3, varchar[100], NULL
            sSQL.Append(SQLNumber(chkIsKCodeByTrans.Checked) & COMMA) 'IsKCodeByTrans , tinyint, NOT NULL
            sSQL.Append(SQLNumber(chkIsReceive.Checked) & COMMA) 'IsReceive
            sSQL.Append(SQLNumber(chkIsManagement.Checked)) 'IsManagement

        End If
        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0002
    '# Created User: 
    '# Created Date: 27/11/2007 02:40:53
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0002() As StringBuilder
        Dim nExtend As Byte
        If Not chkExtend.Checked Then
            nExtend = 0
        ElseIf optNorm.Checked Then
            nExtend = 1
        Else
            nExtend = 2
        End If
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0002 Set ")

        sSQL.Append("AssignmentNameU = " & SQLStringUnicode(txtAssignmentName.Text, gbUnicode, True) & COMMA) 'varchar[50], NULL
		sSQL.Append("DebitAccountID = " & SQLString(tdbcAccountID.Text) & COMMA) 'varchar[20], NULL
		sSQL.Append("DebitObjectTypeID = " & SQLString(tdbcObjectTypeID.Text) & COMMA) 'varchar[20], NULL
		sSQL.Append("DebitObjectID = " & SQLString(tdbcObjectID.Text) & COMMA) 'varchar[20], NULL
		sSQL.Append("Disabled = " & SQLNumber(chkDisabled.Checked) & COMMA) 'tinyint, NOT NULL
		sSQL.Append("SourceID = " & SQLString(tdbcSourceID.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
        sSQL.Append("Extend = " & SQLNumber(nExtend) & COMMA) 'tinyint, NOT NULL
		sSQL.Append("Formular1 = " & SQLString(txtFormular1.Text) & COMMA) 'varchar[100], NOT NULL
		sSQL.Append("Formular2 = " & SQLString(txtFormular2.Text) & COMMA) 'varchar[100], NOT NULL
		sSQL.Append("Formular3 = " & SQLString(txtFormular3.Text) & COMMA) 'varchar[100], NOT NULL
		sSQL.Append("Condition1 = " & SQLString(txtCondition1.Text) & COMMA) 'varchar[100], NOT NULL
		sSQL.Append("Condition2 = " & SQLString(txtCondition2.Text) & COMMA) 'varchar[100], NULL
        sSQL.Append("Condition3 = " & SQLString(txtCondition3.Text) & COMMA) 'varchar[100], NULL
        sSQL.Append("IsKCodeByTrans = " & SQLNumber(chkIsKCodeByTrans.Checked) & COMMA) 'tinyint, NOT NULL
        If tdbg.RowCount > 0 Then
            sSQL.Append("Ana01ID = " & SQLString(tdbg(0, COL_Ana01ID)) & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana02ID = " & SQLString(tdbg(0, COL_Ana02ID)) & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana03ID = " & SQLString(tdbg(0, COL_Ana03ID)) & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana04ID = " & SQLString(tdbg(0, COL_Ana04ID)) & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana05ID = " & SQLString(tdbg(0, COL_Ana05ID)) & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana06ID = " & SQLString(tdbg(0, COL_Ana06ID)) & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana07ID = " & SQLString(tdbg(0, COL_Ana07ID)) & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana08ID = " & SQLString(tdbg(0, COL_Ana08ID)) & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana09ID = " & SQLString(tdbg(0, COL_Ana09ID)) & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana10ID = " & SQLString(tdbg(0, COL_Ana10ID)) & COMMA) 'varchar[20], NULL
            sSQL.Append("PeriodID = " & SQLString(tdbg(0, COL_PeriodID)) & COMMA) 'varchar[20], NOT NULL
            sSQL.Append("ProjectID = " & SQLString(tdbg(0, COL_ProjectID)) & COMMA) 'varchar[50], NOT NULL
            sSQL.Append("TaskID = " & SQLString(tdbg(0, COL_TaskID)) & COMMA) 'TaskID[50], NOT NULL
        Else
            sSQL.Append("Ana01ID = " & SQLString("") & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana02ID = " & SQLString("") & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana03ID = " & SQLString("") & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana04ID = " & SQLString("") & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana05ID = " & SQLString("") & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana06ID = " & SQLString("") & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana07ID = " & SQLString("") & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana08ID = " & SQLString("") & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana09ID = " & SQLString("") & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana10ID = " & SQLString("") & COMMA) 'varchar[20], NULL
            sSQL.Append("PeriodID = " & SQLString("") & COMMA) 'varchar[20], NOT NULL
            sSQL.Append("ProjectID = " & SQLString("") & COMMA) 'varchar[50], NOT NULL
            sSQL.Append("TaskID = " & SQLString("") & COMMA) 'varchar[50], NOT NULL
        End If
        sSQL.Append("IsReceive = " & SQLNumber(chkIsReceive.Checked) & COMMA)
        sSQL.Append("IsManagement =" & SQLNumber(chkIsManagement.Checked))
        sSQL.Append(" Where ")
        sSQL.Append("AssignmentID = " & SQLString(_assignmentID))

        Return sSQL
    End Function

    Private Sub optDifference_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles optDifference.Click
        If optNorm.Checked Then
            txtFormular1.Enabled = True
            txtFormular2.Enabled = True
            txtFormular3.Enabled = True
            txtCondition1.Enabled = True
            txtCondition2.Enabled = True
            txtCondition3.Enabled = True
            txtFormular1.Focus()
        ElseIf optDifference.Checked Then
            txtFormular1.Enabled = False
            txtFormular2.Enabled = False
            txtFormular3.Enabled = False
            txtCondition1.Enabled = False
            txtCondition2.Enabled = False
            txtCondition3.Enabled = False
        End If
    End Sub

    Private Sub optNorm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles optNorm.Click
        If optNorm.Checked Then
            txtFormular1.Enabled = True
            txtFormular2.Enabled = True
            txtFormular3.Enabled = True
            txtCondition1.Enabled = True
            txtCondition2.Enabled = True
            txtCondition3.Enabled = True
            txtFormular1.Focus()
        ElseIf optDifference.Checked Then
            txtFormular1.Enabled = False
            txtFormular2.Enabled = False
            txtFormular3.Enabled = False
            txtCondition1.Enabled = False
            txtCondition2.Enabled = False
            txtCondition3.Enabled = False
        End If
    End Sub

    Private Sub txtAssignmentName_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAssignmentName.Validated
        If Microsoft.VisualBasic.Left(txtAssignmentName.Text, 1) <> UCase(Microsoft.VisualBasic.Left(txtAssignmentName.Text, 1)) Then
            txtAssignmentName.Text = UCase(Microsoft.VisualBasic.Left(txtAssignmentName.Text, 1)) + Microsoft.VisualBasic.Right(txtAssignmentName.Text, Len(txtAssignmentName.Text) - 1)
        End If
    End Sub

    Private Sub EnabledControl()
        Dim bEnabled As Boolean

        Dim sSQL As String
        Dim dt As DataTable
        sSQL = SQLStoreD02P1000()
        dt = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            With dt.Rows(0)
                If .Item("Status").ToString = "1" Then
                    bEnabled = False
                    tdbcAccountID.Enabled = bEnabled
                    tdbcObjectTypeID.Enabled = bEnabled
                    tdbcObjectID.Enabled = bEnabled
                    chkSource.Enabled = bEnabled
                    tdbcSourceID.Enabled = bEnabled
                    chkIsReceive.Enabled = bEnabled
                    chkIsManagement.Enabled = bEnabled
                Else
                    bEnabled = True
                End If
            End With
        End If
    End Sub

    Private Sub tdbg_ButtonClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.ButtonClick
        If DxxFormat.LoadFormNotINV = 0 Then Exit Sub
        If tdbg.AllowUpdate = False Then Exit Sub
        If tdbg.Splits(tdbg.SplitIndex).DisplayColumns(tdbg.Col).Locked Then Exit Sub

        Select Case tdbg.Col
            Case COL_PeriodID, COL_ProjectID, COL_Ana01ID To COL_Ana10ID, COL_TaskID
                Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg.Columns(tdbg.Col).DataField)
                If tdbd Is Nothing Then Exit Select
                Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg, e, tdbd, False)
                AfterColUpdate(e.ColIndex, dr)
                Exit Sub
        End Select
    End Sub


    Private bLoadObectID As Boolean = True
    Private Sub tdbcObjectID_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcObjectID.Validated
        oFilterCombo.FilterCombo(tdbcObjectID, e)
        If tdbcObjectID.FindStringExact(tdbcObjectID.Text) = -1 Then
            tdbcObjectID.Text = ""
            txtObjectName.Text = ""
        End If
        'Xử lý dạng cũ khi Loại ĐT ='' thì Đối tượng gắn lại Loại ĐT => sinh sự kiện bị load lại combo ĐT => nên chăn lại
        If Not oFilterCombo.IsNewFilter AndAlso tdbcObjectID.Splits(0).DisplayColumns("ObjectTypeID").Visible Then
            bLoadObectID = False 'Chặn không cho Load lại Combo Loại ĐT
            tdbcObjectTypeID.SelectedValue = ReturnValueC1Combo(tdbcObjectID, "ObjectTypeID")
        End If
        bLoadObectID = True
    End Sub

    Private Sub chkIsReceive_CheckedChanged(sender As Object, e As EventArgs) Handles chkIsReceive.CheckedChanged
        If chkIsReceive.Checked Then
            chkIsManagement.Checked = False
        End If
    End Sub

    Private Sub chkIsManagement_CheckedChanged(sender As Object, e As EventArgs) Handles chkIsManagement.CheckedChanged
        If chkIsManagement.Checked Then
            chkIsReceive.Checked = False
        End If
    End Sub
End Class