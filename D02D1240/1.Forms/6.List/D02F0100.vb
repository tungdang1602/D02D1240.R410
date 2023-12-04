'#-------------------------------------------------------------------------------------
'# Created Date: 05/05/2011 8:32:09 AM
'# Created User: Nguyễn Đức Trọng
'# Modify Date: 05/05/2011 8:32:09 AM
'# Modify User: Nguyễn Đức Trọng
'#-------------------------------------------------------------------------------------
Public Class D02F0100
	Dim report As D99C2003
	Private _formIDPermission As String = "D02F0100"
	Public WriteOnly Property FormIDPermission() As String
		Set(ByVal Value As String)
			       _formIDPermission = Value
		   End Set
	End Property


#Region "Const of tdbg - Total of Columns: 21"
    Private Const COL_AssignmentID As Integer = 0      ' Mã phân bổ
    Private Const COL_AssignmentName As Integer = 1    ' Tên phân bổ
    Private Const COL_DebitAccountID As Integer = 2    ' Tài khoản phân bổ
    Private Const COL_DebitObjectTypeID As Integer = 3 ' Đối tượng phân bổ
    Private Const COL_SourceID As Integer = 4          ' Nguồn vốn
    Private Const COL_PeriodID As Integer = 5          ' Tập phí
    Private Const COL_Ana01ID As Integer = 6           ' Khoản mục 01
    Private Const COL_Ana02ID As Integer = 7           ' Khoản mục 02
    Private Const COL_Ana03ID As Integer = 8           ' Khoản mục 03
    Private Const COL_Ana04ID As Integer = 9           ' Khoản mục 04
    Private Const COL_Ana05ID As Integer = 10          ' Khoản mục 05
    Private Const COL_Ana06ID As Integer = 11          ' Khoản mục 06
    Private Const COL_Ana07ID As Integer = 12          ' Khoản mục 07
    Private Const COL_Ana08ID As Integer = 13          ' Khoản mục 08
    Private Const COL_Ana09ID As Integer = 14          ' Khoản mục 09
    Private Const COL_Ana10ID As Integer = 15          ' Khoản mục 10
    Private Const COL_Disabled As Integer = 16         ' KSD
    Private Const COL_CreateUserID As Integer = 17     ' CreateUserID
    Private Const COL_CreateDate As Integer = 18       ' CreateDate
    Private Const COL_LastModifyUserID As Integer = 19 ' LastModifyUserID
    Private Const COL_LastModifyDate As Integer = 20   ' LastModifyDate
#End Region


    Private Const sAuditCode As String = "AllocationCodes"
    Private dtGrid, dtCaptionCols As DataTable
    Dim bRefreshFilter As Boolean
    Dim sFilter As New System.Text.StringBuilder()

    Private Sub D02F0100_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
	LoadInfoGeneral()
        Me.Cursor = Cursors.WaitCursor
        gbEnabledUseFind = False
        ResetColorGrid(tdbg)
        Loadlanguage()
        LoadTDBGridAnalysisCaption(D02, tdbg, COL_Ana01ID, SPLIT0, True, gbUnicode) 'ID : 245776

        LoadTDBGrid()
        SetShortcutPopupMenu(Me, TableToolStrip, ContextMenuStrip1)
        SetResolutionForm(Me, ContextMenuStrip1)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Danh_muc_tieu_thuc_phan_bo_khau_hao_-_D02F0100") & UnicodeCaption(gbUnicode) 'Danh móc ti£u th÷c ph¡n bå khÊu hao - D02F0100
        '================================================================ 
        tdbg.Columns("AssignmentID").Caption = rl3("Ma_phan_bo") 'Mã phân bổ
        tdbg.Columns("AssignmentName").Caption = rl3("Ten_phan_bo") 'Tên phân bổ
        tdbg.Columns("DebitAccountID").Caption = rl3("Tai_khoan_phan_bo") 'Tài khoản phân bổ
        tdbg.Columns("DebitObjectTypeID").Caption = rl3("Doi_tuong_phan_bo") 'Đối tượng phân bổ
        tdbg.Columns("SourceID").Caption = rl3("Nguon_von") 'Nguồn vốn
        tdbg.Columns("Disabled").Caption = rl3("KSD") 'Không sử dụng 
        '================================================================ 
        chkShowDisabled.Text = rL3("Hien_thi_danh_muc_khong_su_dung") 'Hiển thị danh mục không sử dụng

        '================================================================ 
        tdbg.Columns(COL_PeriodID).Caption = rL3("Tap_phi") 'Tập phí
        tdbg.Columns(COL_Ana01ID).Caption = rL3("Khoan_muc") & " 01" 'Khoản mục 01
        tdbg.Columns(COL_Ana02ID).Caption = rL3("Khoan_muc") & " 02" 'Khoản mục 02
        tdbg.Columns(COL_Ana03ID).Caption = rL3("Khoan_muc") & " 03" 'Khoản mục 03
        tdbg.Columns(COL_Ana04ID).Caption = rL3("Khoan_muc") & " 04" 'Khoản mục 04
        tdbg.Columns(COL_Ana05ID).Caption = rL3("Khoan_muc") & " 05" 'Khoản mục 05
        tdbg.Columns(COL_Ana06ID).Caption = rL3("Khoan_muc") & " 06" 'Khoản mục 06
        tdbg.Columns(COL_Ana07ID).Caption = rL3("Khoan_muc") & " 07" 'Khoản mục 07
        tdbg.Columns(COL_Ana08ID).Caption = rL3("Khoan_muc") & " 08" 'Khoản mục 08
        tdbg.Columns(COL_Ana09ID).Caption = rL3("Khoan_muc") & " 09" 'Khoản mục 09
        tdbg.Columns(COL_Ana10ID).Caption = rL3("Khoan_muc") & " 10" 'Khoản mục 10
        tdbg.Columns(COL_Disabled).Caption = rL3("KSD") 'KSD

    End Sub

    Private Sub LoadTDBGrid(Optional ByVal FlagAdd As Boolean = False, Optional ByVal sKey As String = "")
        Dim sSQL As String
        sSQL = "Select AssignmentID, AssignmentName" & UnicodeJoin(gbUnicode) & " As AssignmentName, DebitAccountID, DebitObjectTypeID, SourceID, Disabled, " & vbCrLf
        sSQL &= "CreateUserID, CreateDate, LastModifyUserID, LastModifyDate, PeriodID, Ana01ID, Ana02ID, Ana03ID, Ana04ID" & vbCrLf
        sSQL &= ", Ana05ID, Ana06ID, Ana07ID, Ana08ID, Ana09ID, Ana10ID " & vbCrLf
        sSQL &= "From D02T0002 WITH(NOLOCK) Order By AssignmentID"
        dtGrid = ReturnDataTable(sSQL)

        gbEnabledUseFind = dtGrid.Rows.Count > 0

        If FlagAdd Then ' Thêm mới thì set Filter = "" và sFind =""
            ResetFilter(tdbg, sFilter, bRefreshFilter)
            sFind = ""
            sFilter = New System.Text.StringBuilder("")
        End If

        LoadDataSource(tdbg, dtGrid, gbUnicode)
        ReLoadTDBGrid()

        If sKey <> "" Then
            Dim dt1 As DataTable = dtGrid.DefaultView.ToTable
            Dim dr() As DataRow = dt1.Select("AssignmentID = " & SQLString(sKey), dt1.DefaultView.Sort)
            If dr.Length > 0 Then tdbg.Row = dt1.Rows.IndexOf(dr(0))
        End If

        If Not tdbg.Focused Then tdbg.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới
    End Sub

    Private Sub ReLoadTDBGrid()
        Dim strFind As String = sFind
        If sFilter.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
        strFind &= sFilter.ToString

        If Not chkShowDisabled.Checked Then
            If strFind <> "" Then strFind &= " And "
            strFind &= "Disabled =0"
        End If
        dtGrid.DefaultView.RowFilter = strFind
        ResetGrid()
    End Sub

    Private Sub ResetGrid()
        CheckMenu(_formIDPermission, TableToolStrip, tdbg.RowCount, gbEnabledUseFind, False, ContextMenuStrip1, , "D02F0100")
        FooterTotalGrid(tdbg, COL_AssignmentName)
    End Sub

#Region "Active Find Client - List All "
    Private WithEvents Finder As New D99C1001
	Dim gbEnabledUseFind As Boolean = False
    'Cần sửa Tìm kiếm như sau:
	'Bỏ sự kiện Finder_FindClick.
	'Sửa tham số Me.Name -> Me
	'Phải tạo biến properties có tên chính xác strNewFind và strNewServer
	'Sửa gdtCaptionExcel thành dtCaptionCols: biến toàn cục trong form
	'Nếu có F12 dùng D09U1111 thì Sửa dtCaptionCols thành ResetTableByGrid(usrOption, dtCaptionCols.DefaultView.ToTable)
    Private sFind As String = ""
	Public WriteOnly Property strNewFind() As String
		Set(ByVal Value As String)
			sFind = Value
			ReLoadTDBGrid()'Làm giống sự kiện Finder_FindClick. Ví dụ đối với form Báo cáo thường gọi btnPrint_Click(Nothing, Nothing): sFind = "
		End Set
	End Property


    Private Sub tsbFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbFind.Click, tsmFind.Click, mnsFind.Click
        ''If Not CallMenuFromGrid(sender, tdbg) Then Exit Sub

        gbEnabledUseFind = True
        '*****************************************
        'Chuẩn hóa D09U1111 : Tìm kiếm dùng table caption có sẵn
        tdbg.UpdateData()
        'If dtCaptionCols Is Nothing OrElse dtCaptionCols.Rows.Count < 1 Then 'Incident 72333
        Dim Arr As New ArrayList
        AddColVisible(tdbg, SPLIT0, Arr, , , , gbUnicode)
        dtCaptionCols = CreateTableForExcelOnly(tdbg, Arr)
        'End If
        ShowFindDialogClient(Finder, dtCaptionCols, Me, "0", gbUnicode)
    End Sub

    Private Sub tsbListAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbListAll.Click, tsmListAll.Click, mnsListAll.Click
        ''If Not CallMenuFromGrid(sender, tdbg) Then Exit Sub
        sFind = ""
        ResetFilter(tdbg, sFilter, bRefreshFilter)
        ReLoadTDBGrid()
    End Sub

#End Region

#Region "Menu bar"

    Private Sub tsbAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbAdd.Click, tsmAdd.Click, mnsAdd.Click
        Dim f As New D02F0101
        With f
            .AssignmentID = ""
            .sAuditCode = sAuditCode
            .FormState = EnumFormState.FormAdd
            .ShowDialog()
            If f.SavedOK Then LoadTDBGrid(True, .AssignmentID)
            .Dispose()
        End With
    End Sub

    Private Sub tsbView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbView.Click, tsmView.Click, mnsView.Click
        Dim f As New D02F0101
        With f
            .AssignmentID = tdbg.Columns(COL_AssignmentID).Text
            .FormState = EnumFormState.FormView
            .ShowDialog()
            .Dispose()
        End With
    End Sub

    Private Sub tsbEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbEdit.Click, tsmEdit.Click, mnsEdit.Click
        Dim f As New D02F0101
        With f
            .AssignmentID = tdbg.Columns(COL_AssignmentID).Text
            .sAuditCode = sAuditCode
            .FormState = EnumFormState.FormEdit
            .ShowDialog()
            .Dispose()
        End With
        If f.SavedOK Then LoadTDBGrid(False, tdbg.Columns(COL_AssignmentID).Text)
    End Sub

    Private Sub tsbDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbDelete.Click, tsmDelete.Click, mnsDelete.Click
        If AskDelete() = Windows.Forms.DialogResult.No Then Exit Sub
        If Not AllowDelete() Then Exit Sub 'Kiểm tra điều kiện xóa

        Dim sSQL As String = ""
        sSQL &= "Delete From D02T0002 Where AssignmentID = " & SQLString(tdbg.Columns(COL_AssignmentID).Text)
        Dim bResult As Boolean = ExecuteSQL(sSQL)
        If bResult = True Then
            'ExecuteAuditLog(sAuditCode, "03", tdbg.Columns(COL_AssignmentID).Text, tdbg.Columns(COL_DebitAccountID).Text, tdbg.Columns(COL_DebitObjectTypeID).Text, tdbg.Columns(COL_SourceID).Text)
            Lemon3.D91.RunAuditLog("02", sAuditCode, "03", tdbg.Columns(COL_AssignmentID).Text, tdbg.Columns(COL_DebitAccountID).Text, tdbg.Columns(COL_DebitObjectTypeID).Text, tdbg.Columns(COL_SourceID).Text)
            DeleteGridEvent(tdbg, dtGrid, gbEnabledUseFind)
            ResetGrid()
            DeleteOK()
        Else
            DeleteNotOK()
        End If
    End Sub

    Private Function AllowDelete() As Boolean
        Return CheckStore(SQLStoreD02P1000)
    End Function

    Private Sub tsbSysInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbSysInfo.Click, tsmSysInfo.Click, mnsSysInfo.Click
        ShowSysInfoDialog(Me,tdbg.Columns(COL_CreateUserID).Text, tdbg.Columns(COL_CreateDate).Text, tdbg.Columns(COL_LastModifyUserID).Text, tdbg.Columns(COL_LastModifyDate).Text)
    End Sub

    Private Sub tsbClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbClose.Click
        Me.Close()
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0100
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 22/03/2012 11:10:34
    '# Modified User: 
    '# Modified Date: 
    '# Description: In
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0100() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0100 "
        sSQL &= SQLString("%") & COMMA 'AssignmentID, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function


    Private Sub tsbPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbPrint.Click, tsmPrint.Click, mnsPrint.Click

        Me.Cursor = Cursors.WaitCursor
        'Dim report As New D99C1003
        'Đưa vể đầu tiên hàm In trước khi gọi AllowPrint()
		If Not AllowNewD99C2003(report, Me) Then Exit Sub
		'************************************
        Dim conn As New SqlConnection(gsConnectionString)
        Dim sReportName As String = "D02R0100"
        Dim sSubReportName As String = "D02R0000"
        Dim sReportCaption As String = ""
        Dim sPathReport As String = ""
        Dim sSQL As String = ""
        Dim sSQLSub As String = ""

        sReportCaption = rL3("Danh_muc_tieu_thuc_phan_bo_khau_hao") & " - " & sReportName
        sPathReport = UnicodeGetReportPath(gbUnicode, D02Options.ReportLanguage, "") & sReportName & ".rpt"
        ' sSQL = "Select * From D02V0100"
        sSQLSub = "Select Top 1 * From D91T0025 WITH(NOLOCK)"
        UnicodeSubReport(sSubReportName, sSQLSub, , gbUnicode)

        With report
            .OpenConnection(conn)
            .AddSub(sSQLSub, sSubReportName & ".rpt")
            .AddMain(SQLStoreD02P0100)
            .PrintReport(sPathReport, sReportCaption)
        End With
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub chkShowDisabled_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowDisabled.CheckedChanged
        If dtGrid Is Nothing Then Exit Sub
        ReLoadTDBGrid()
    End Sub

#End Region

#Region "Grid"

    Private Sub tdbg_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.DoubleClick
        If tdbg.FilterActive Then Exit Sub
        If tsbEdit.Enabled Then
            tsbEdit_Click(sender, Nothing)
        ElseIf tsbView.Enabled Then
            tsbView_Click(sender, Nothing)
        End If
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        If e.KeyCode = Keys.Enter Then tdbg_DoubleClick(Nothing, Nothing)
        HotKeyCtrlVOnGrid(tdbg, e)
    End Sub

    Private Sub tdbg_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.FilterChange
        Try
            If (dtGrid Is Nothing) Then Exit Sub
            If bRefreshFilter Then Exit Sub 'set FilterText ="" thì thoát
            FilterChangeGrid(tdbg, sFilter)
            ReLoadTDBGrid()
        Catch ex As Exception
            'MessageBox.Show(ex.Message & " - " & ex.Source)
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub

#End Region

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1000
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 26/11/2007 11:47:21
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1000() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P1000 "
        sSQL &= SQLString(tdbg.Columns(COL_AssignmentID).Text) & COMMA 'AssignmentID, varchar[20], NOT NULL
        sSQL &= SQLNumber(gsLanguage) 'Language, tinyint, NOT NULL
        Return sSQL
    End Function

    Private Sub mnsImportData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnsImportData.Click, tsbImportData.Click, tsmImportData.Click
        If CallShowDialogD80F2090(D02, "D02F0100", Me.Name) Then LoadTDBGrid(, tdbg.Columns(COL_AssignmentID).Text)
    End Sub
    'ID : 252769
    Private Sub mnsExportToExcel_Click(sender As Object, e As EventArgs) Handles mnsExportToExcel.Click, tsbExportToExcel.Click, tsmExportToExcel.Click
        '*****************************************
        '    If dtCaptionCols Is Nothing OrElse dtCaptionCols.Rows.Count < 1 Then
        'Những cột bắt buộc nhập
        Dim arrColObligatory() As Integer = {}
        Dim Arr As New ArrayList
        AddColVisible(tdbg, SPLIT0, Arr, arrColObligatory, False, , gbUnicode)
        AddColVisible(tdbg, SPLIT1, Arr, arrColObligatory, False, , gbUnicode)
        'Tạo tableCaption: đưa tất cả các cột trên lưới có Visible = True vào table 
        dtCaptionCols = CreateTableForExcelOnly(tdbg, Arr)
        '   End If

        Dim frm As New D99F2222
        With frm
            .FormID = Me.Name
            .UseUnicode = gbUnicode
            .dtLoadGrid = dtCaptionCols
            .GroupColumns = gsGroupColumns
            .dtExportTable = dtGrid
            .ShowDialog()
            .Dispose()
        End With

        '*****************************************
    End Sub
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        AnchorResizeColumnsGrid(EnumAnchorStyles.TopLeftRightBottom, tdbg)
        AnchorForControl(EnumAnchorStyles.BottomLeft, chkShowDisabled)
    End Sub
  
End Class