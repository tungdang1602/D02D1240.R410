'#---------------------------------------------------------------------------------------------------
'# Title: D02E0200
'# Created User: Lê Thị Thanh Hiền
'# Created Date: 31/07/2007 10:51:43
'# Modify Date: 05/05/2011 8:32:25 AM
'# Modify User: Nguyễn Đức Trọng
'# Description: 
'#---------------------------------------------------------------------------------------------------

Public Class D02F0200
	Dim report As D99C2003


	Private _formIDPermission As String = "D02F0200"
	Public WriteOnly Property FormIDPermission() As String
		Set(ByVal Value As String)
			       _formIDPermission = Value
		   End Set
	End Property


#Region "Const of tdbg"
    Private Const COL_SourceID As Integer = 0         ' Mã nguồn hình thành
    Private Const COL_SourceName As Integer = 1       ' Tên nguồn hình thành
    Private Const COL_Disabled As Integer = 2         ' Không sử dụng
    Private Const COL_CreateUserID As Integer = 3     ' Người tạo
    Private Const COL_LastModifyUserID As Integer = 4 ' LastModifyUserID
    Private Const COL_LastModifyDate As Integer = 5   ' LastModifyDate
    Private Const COL_CreateDate As Integer = 6       ' Ngày tạo
#End Region

    Private dtGrid As DataTable
    Dim bRefreshFilter As Boolean
    Dim sFilter As New System.Text.StringBuilder()

    Private Sub D02F0200_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
    End Sub

    Private Sub D02F0200_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
	LoadInfoGeneral()
        SetShortcutPopupMenu(Me, TableToolStrip, ContextMenuStrip1)
        Loadlanguage()
        ResetColorGrid(tdbg)
        LoadTDBGrid()
        InputbyUnicode(Me, gbUnicode)
        SetResolutionForm(Me, ContextMenuStrip1)
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Danh_muc_nguon_hinh_thanh_-_D02F0200") & UnicodeCaption(gbUnicode) 'Danh móc nguän hØnh thªnh - D02F0200
        '================================================================ 
        tdbg.Columns("SourceID").Caption = rl3("Ma_nguon_hinh_thanh") 'Mã nguồn hình thành
        tdbg.Columns("SourceName").Caption = rl3("Ten_nguon_hinh_thanh") 'Tên nguồn hình thành
        tdbg.Columns("Disabled").Caption = rl3("KSD") 'Không sử dụng
        '================================================================ 
        chkShowDisabled.Text = rl3("Hien_thi_danh_muc_khong_su_dung") 'Hiển thị danh mục không sử dụng
    End Sub

    Private Sub LoadTDBGrid(Optional ByVal FlagAdd As Boolean = False, Optional ByVal sKey As String = "")
        Dim sSQL As String
        sSQL = "Select      SourceID, SourceName" & UnicodeJoin(gbUnicode) & " As SourceName, "
        sSQL &= "Disabled, CreateUserID, CreateDate, LastModifyUserID, LastModifyDate" & vbCrLf
        sSQL &= "From       D02T0013 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Order by   SourceID" & vbCrLf
        dtGrid = ReturnDataTable(sSQL)

        gbEnabledUseFind = dtGrid.Rows.Count > 0
        If FlagAdd Then ' Thêm mới thì set Filter = "" và sFind =""
            ResetFilter(tdbg, sFilter, bRefreshFilter)
        End If

        LoadDataSource(tdbg, dtGrid, gbUnicode)
        ReLoadTDBGrid()

        If sKey <> "" Then
            Dim dt1 As DataTable = dtGrid.DefaultView.ToTable
            Dim dr() As DataRow = dt1.Select("SourceID = " & SQLString(sKey), dt1.DefaultView.Sort)
            If dr.Length > 0 Then tdbg.Row = dt1.Rows.IndexOf(dr(0))
        End If

        If Not tdbg.Focused Then tdbg.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới
    End Sub

    Private Sub ReLoadTDBGrid()
        Dim strFind As String = sFilter.ToString

        If Not chkShowDisabled.Checked Then
            If strFind <> "" Then strFind &= " And "
            strFind &= "Disabled = 0"
        End If
        dtGrid.DefaultView.RowFilter = strFind
        ResetGrid()
    End Sub

    Private Sub ResetGrid()
        CheckMenu(_formIDPermission, TableToolStrip, tdbg.RowCount, gbEnabledUseFind, False, ContextMenuStrip1)
        FooterTotalGrid(tdbg, COL_SourceName)
    End Sub

    Private Sub chkShowDisabled_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowDisabled.CheckedChanged
        If dtGrid Is Nothing Then Exit Sub
        ReLoadTDBGrid()
    End Sub

#Region "Menu Bar"

    Private Sub tsbAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbAdd.Click
        Dim f As New D02F0201()
        With f
            .SourceID = ""
            .FormState = EnumFormState.FormAdd
            .ShowDialog()
            If .SavedOK = True Then LoadTDBGrid(True, .SourceID)
            .Dispose()
        End With
    End Sub

    Private Sub tsbEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbEdit.Click
        Dim f As New D02F0201()
        With f
            .SourceID = tdbg.Columns(COL_SourceID).Text
            .FormState = EnumFormState.FormEdit
            .ShowDialog()
            .Dispose()
        End With
        If f.SavedOK = True Then LoadTDBGrid(False, tdbg.Columns(COL_SourceID).Text)
    End Sub

    Private Sub tsbView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbView.Click
        Dim f As New D02F0201()
        With f
            .SourceID = tdbg.Columns(COL_SourceID).Text
            .FormState = EnumFormState.FormView
            .ShowDialog()
            .Dispose()
        End With
    End Sub

    Private Sub tsbDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbDelete.Click
        If D99C0008.MsgAskDelete = Windows.Forms.DialogResult.No Then Exit Sub
        If NotAllowDelete() Then
            D99C0008.MsgCanNotDelete()
            Exit Sub
        End If

        Dim sSQL As String = SQLDeleteD02T0013()
        Dim bResult As Boolean = ExecuteSQL(sSQL)
        If bResult Then
            DeleteGridEvent(tdbg, dtGrid, gbEnabledUseFind)
            ResetGrid()
            DeleteOK()
        Else
            DeleteNotOK()
        End If
    End Sub

    Private Function NotAllowDelete() As Boolean
        Dim sSQL As String = ""
        sSQL = "Select Top 1 1 From D02T0012 WITH(NOLOCK) Where SourceID= " & SQLString(tdbg.Columns(COL_SourceID).Text)
        Return ExistRecord(sSQL)
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD02T0013
    '# Created User: Lê Thị Thanh Hiền
    '# Created Date: 30/07/2007 01:42:31
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD02T0013() As String
        Dim sSQL As String = ""
        sSQL &= "Delete From D02T0013"
        sSQL &= " Where "
        sSQL &= "SourceID = " & SQLString(tdbg.Columns(COL_SourceID).Text)
        Return sSQL
    End Function

    Private Sub tsbSysInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbSysInfo.Click
        ShowSysInfoDialog(Me,tdbg.Columns(COL_CreateUserID).Text, tdbg.Columns(COL_CreateDate).Text, tdbg.Columns(COL_LastModifyUserID).Text, tdbg.Columns(COL_LastModifyDate).Text)
    End Sub

    Private Sub tsbClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbClose.Click
        Me.Close()
    End Sub

    Private Sub tsbPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbPrint.Click
        Me.Cursor = Cursors.WaitCursor

        'Dim report As New D99C1003
        'Đưa vể đầu tiên hàm In trước khi gọi AllowPrint()
		If Not AllowNewD99C2003(report, Me) Then Exit Sub
		'************************************
        Dim conn As New SqlConnection(gsConnectionString)
        Dim sReportName As String = "D02R0200"
        Dim sSubReportName As String = "D02R0000"
        Dim sReportCaption As String = rl3("Bao_cao_danh_moc_nguon_hinh_thanh")
        Dim sPathReport As String = ""
        Dim sSQL As String = ""
        Dim sSQLSub As String = ""

        sReportCaption = sReportCaption & " - " & sReportName
        sPathReport = gsApplicationSetup & "\XReports\" & sReportName & ".rpt"
        sSQL = "Select * from D02T0013 WITH(NOLOCK) order by SourceID"
        sSQLSub = "Select TOP 1*  from D91T0025 WITH(NOLOCK)"
        With report
            .OpenConnection(conn)
            .AddSub(sSQLSub, sSubReportName & ".rpt")
            .AddMain(sSQL)
            .PrintReport(sPathReport, sReportCaption)
        End With
        Me.Cursor = Cursors.Default
    End Sub
#End Region

#Region "Menu Dropdown"

    Private Sub tsmAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsmAdd.Click, mnsAdd.Click
        tsbAdd_Click(Nothing, Nothing)
    End Sub

    Private Sub tsmEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsmEdit.Click, mnsEdit.Click
        tsbEdit_Click(Nothing, Nothing)
    End Sub

    Private Sub tsmView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmView.Click, mnsView.Click
        tsbView_Click(Nothing, Nothing)
    End Sub

    Private Sub tsmDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmDelete.Click, mnsDelete.Click
        tsbDelete_Click(Nothing, Nothing)
    End Sub

    Private Sub tsmSysInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmSysInfo.Click, mnsSysInfo.Click
        tsbSysInfo_Click(Nothing, Nothing)
    End Sub

    Private Sub tsmPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmPrint.Click, mnsPrint.Click
        tsbPrint_Click(Nothing, Nothing)
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

End Class