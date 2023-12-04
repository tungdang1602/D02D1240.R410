'#-------------------------------------------------------------------------------------
'# Created Date: 27/08/2007 3:31:15 PM
'# Created User: Trần Thị ÁiTrâm
'# Modify Date: 
'# Modify User: 
'#-------------------------------------------------------------------------------------
Imports System.Text
Imports System

Public Class D02F3000
	Dim report As D99C2003

    Private _savedOK As Boolean
    Public ReadOnly Property SavedOK() As Boolean
        Get
            Return _savedOK
        End Get
    End Property

	Private _formIDPermission As String = "D02F3000"
	Public WriteOnly Property FormIDPermission() As String
		Set(ByVal Value As String)
			       _formIDPermission = Value
		   End Set
	End Property


#Region "Const of tdbg1"
    Private Const COL1_AssetS1ID As Integer = 0        ' Mã phân loại 1
    Private Const COL1_AssetS1Name As Integer = 1      ' Tên phân loại 1
    Private Const COL1_CreateUserID As Integer = 2     ' CreateUserID
    Private Const COL1_CreateDate As Integer = 3       ' CreateDate
    Private Const COL1_LastModifyUserID As Integer = 4 ' LastModifyUserID
    Private Const COL1_LastModifyDate As Integer = 5   ' LastModifyDate
    Private Const COL1_Disabled As Integer = 6         ' Không sử dụng
#End Region

#Region "Const of tdbg2"
    Private Const COL2_AssetS2ID As Integer = 0        ' Mã phân loại 2
    Private Const COL2_AssetS2Name As Integer = 1      ' Tên phân loại 2
    Private Const COL2_Disabled As Integer = 2         ' Không sử dụng
    Private Const COL2_CreateUserID As Integer = 3     ' CreateUserID
    Private Const COL2_CreateDate As Integer = 4       ' CreateDate
    Private Const COL2_LastModifyUserID As Integer = 5 ' LastModifyUserID
    Private Const COL2_LastModifyDate As Integer = 6   ' LastModifyDate
#End Region

#Region "Const of tdbg3"
    Private Const COL3_AssetS3ID As Integer = 0        ' Mã phân loại 3
    Private Const COL3_AssetS3Name As Integer = 1      ' Tên phân loại 3
    Private Const COL3_Disabled As Integer = 2         ' Không sử dụng
    Private Const COL3_CreateUserID As Integer = 3     ' CreateUserID
    Private Const COL3_CreateDate As Integer = 4       ' CreateDate
    Private Const COL3_LastModifyUserID As Integer = 5 ' LastModifyUserID
    Private Const COL3_LastModifyDate As Integer = 6   ' LastModifyDate
#End Region

#Region "Const of tdbg4 - Total of Columns: 7"
    Private Const COL4_AssetS4ID As Integer = 0        ' Mã phân loại 4
    Private Const COL4_AssetS4Name As Integer = 1      ' Tên phân loại 4
    Private Const COL4_Disabled As Integer = 2         ' KSD
    Private Const COL4_CreateUserID As Integer = 3     ' CreateUserID
    Private Const COL4_CreateDate As Integer = 4       ' CreateDate
    Private Const COL4_LastModifyUserID As Integer = 5 ' LastModifyUserID
    Private Const COL4_LastModifyDate As Integer = 6   ' LastModifyDate
#End Region


#Region "Const of tdbg5 - Total of Columns: 7"
    Private Const COL5_AssetS5ID As Integer = 0        ' Mã phân loại 5
    Private Const COL5_AssetS5Name As Integer = 1      ' Tên phân loại 5
    Private Const COL5_Disabled As Integer = 2         ' KSD
    Private Const COL5_CreateUserID As Integer = 3     ' CreateUserID
    Private Const COL5_CreateDate As Integer = 4       ' CreateDate
    Private Const COL5_LastModifyUserID As Integer = 5 ' LastModifyUserID
    Private Const COL5_LastModifyDate As Integer = 6   ' LastModifyDate
#End Region


    Dim dtGrid1, dtGrid2, dtGrid3, dtGrid4, dtGrid5 As DataTable

    Dim bRefreshFilter1, bRefreshFilter2, bRefreshFilter3, bRefreshFilter4, bRefreshFilter5 As Boolean
    Dim sFilter1 As New System.Text.StringBuilder()
    Dim sFilter2 As New System.Text.StringBuilder()
    Dim sFilter3 As New System.Text.StringBuilder()
    Dim sFilter4 As New System.Text.StringBuilder()
    Dim sFilter5 As New System.Text.StringBuilder()

    Private Sub D02F3000_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
        If e.Alt = True And (e.KeyCode = Keys.D1 Or e.KeyCode = Keys.NumPad1) Then
            Application.DoEvents()
            tabMain.SelectedTab = TabPage1
            Application.DoEvents()
            tdbg1.Focus()

        End If
        If e.Alt = True And (e.KeyCode = Keys.D2 Or e.KeyCode = Keys.NumPad2) Then
            Application.DoEvents()
            tabMain.SelectedTab = TabPage2
            Application.DoEvents()
            tdbg2.Focus()

        End If
        If e.Alt = True And (e.KeyCode = Keys.D3 Or e.KeyCode = Keys.NumPad3) Then
            Application.DoEvents()
            tabMain.SelectedTab = TabPage3
            Application.DoEvents()
            tdbg3.Focus()
        End If
        If e.Alt = True And (e.KeyCode = Keys.D4 Or e.KeyCode = Keys.NumPad4) Then
            Application.DoEvents()
            tabMain.SelectedTab = TabPage4
            Application.DoEvents()
            tdbg4.Focus()
        End If
        If e.Alt = True And (e.KeyCode = Keys.D5 Or e.KeyCode = Keys.NumPad5) Then
            Application.DoEvents()
            tabMain.SelectedTab = TabPage5
            Application.DoEvents()
            tdbg5.Focus()
        End If
    End Sub

    Private Sub D02F3000_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
	LoadInfoGeneral()
        SetShortcutPopupMenu(Me, TableToolStrip, ContextMenuStrip1)
        Loadlanguage()
        ResetColorGrid(tdbg1)
        ResetColorGrid(tdbg2)
        ResetColorGrid(tdbg3)
        ResetColorGrid(tdbg4)
        ResetColorGrid(tdbg5)
        gbEnabledUseFind = False
        LoadTDBGrid()
        InputbyUnicode(Me, gbUnicode)
        VisibleMenuImport() '29/8/2019, Lê Thị Phú Hà:id 122897-Hỗ trợ import Excel cho mã phân loại 2 TSCĐ
        SetResolutionForm(Me, ContextMenuStrip1)
    End Sub

    'Private Sub Loadlanguage()
    '    '================================================================ 
    '    Me.Text = rl3("Danh_muc_phan_loai_-_D02F3000") & UnicodeCaption(gbUnicode) 'Danh móc ph¡n loÁi - D02F3000
    '    '================================================================ 
    '    TabPage1.Text = "1. " & rl3("Phan_loai") & " 1" '1. Phân loại 1
    '    TabPage2.Text = "2. " & rl3("Phan_loai") & " 2" '2. Phân loại 2
    '    TabPage3.Text = "3. " & rl3("Phan_loai") & " 3" '3. Phân loại 3
    '    '================================================================ 
    '    tdbg1.Columns("AssetS1ID").Caption = rl3("Ma_phan_loai") & " 1" 'Mã phân loại 1
    '    tdbg1.Columns("AssetS1Name").Caption = rl3("Ten_phan_loai") & " 1" 'Tên phân loại 1
    '    tdbg1.Columns("Disabled").Caption = rl3("KSD") 'KSD
    '    tdbg2.Columns("AssetS2ID").Caption = rl3("Ma_phan_loai") & " 2" 'Mã phân loại 2
    '    tdbg2.Columns("AssetS2Name").Caption = rl3("Ten_phan_loai") & " 2" 'Tên phân loại 2
    '    tdbg2.Columns("Disabled").Caption = rl3("KSD") 'KSD
    '    tdbg3.Columns("AssetS3ID").Caption = rl3("Ma_phan_loai") & " 3" 'Mã phân loại 3
    '    tdbg3.Columns("AssetS3Name").Caption = rl3("Ten_phan_loai") & " 3" 'Tên phân loại 3
    '    tdbg3.Columns("Disabled").Caption = rl3("KSD") 'KSD
    '    '================================================================ 
    '    chkShowDisabled1.Text = rl3("Hien_thi_danh_muc_khong_su_dung") 'Hiển thị danh mục không sử dụng
    '    chkShowDisabled2.Text = rl3("Hien_thi_danh_muc_khong_su_dung") 'Hiển thị danh mục không sử dụng
    '    chkShowDisabled3.Text = rl3("Hien_thi_danh_muc_khong_su_dung") 'Hiển thị danh mục không sử dụng
    'End Sub

    Private Sub LoadLanguage()
        '================================================================ 
        Me.Text = rl3("Danh_muc_phan_loai") & " - " & Me.Name & UnicodeCaption(gbUnicode) 'Danh móc ph¡n loÁi
        '================================================================ 
        chkShowDisabled1.Text = rl3("Hien_thi_danh_muc_khong_su_dung") 'Hiển thị danh mục không sử dụng
        chkShowDisabled2.Text = rl3("Hien_thi_danh_muc_khong_su_dung") 'Hiển thị danh mục không sử dụng
        chkShowDisabled3.Text = rl3("Hien_thi_danh_muc_khong_su_dung") 'Hiển thị danh mục không sử dụng
        chkShowDisabled4.Text = rl3("Hien_thi_danh_muc_khong_su_dung") 'Hiển thị danh mục không sử dụng
        chkShowDisabled5.Text = rl3("Hien_thi_danh_muc_khong_su_dung") 'Hiển thị danh mục không sử dụng
        '================================================================ 
        TabPage1.Text = "1. " & rl3("Phan_loai") & " 1" '1. Phân loại 1
        TabPage2.Text = "2. " & rl3("Phan_loai") & " 2" '2. Phân loại 2
        TabPage3.Text = "3. " & rl3("Phan_loai") & " 3" '3. Phân loại 3
        TabPage4.Text = "4. " & rl3("Phan_loai") & " 4" '4. Phân loại 4
        TabPage5.Text = "5. " & rl3("Phan_loai") & " 5" '5. Phân loại 5
        '================================================================ 
        tdbg1.Columns(COL1_AssetS1ID).Caption = rl3("Ma_phan_loai") & " 1" 'Mã phân loại 1
        tdbg1.Columns(COL1_AssetS1Name).Caption = rl3("Ten_phan_loai") & " 1" 'Tên phân loại 1
        tdbg1.Columns(COL1_Disabled).Caption = rl3("KSD") 'KSD
        tdbg2.Columns(COL2_AssetS2ID).Caption = rl3("Ma_phan_loai") & " 2" 'Mã phân loại 2
        tdbg2.Columns(COL2_AssetS2Name).Caption = rl3("Ten_phan_loai") & " 2" 'Tên phân loại 2
        tdbg2.Columns(COL2_Disabled).Caption = rl3("KSD") 'KSD
        tdbg3.Columns(COL3_AssetS3ID).Caption = rl3("Ma_phan_loai") & " 3" 'Mã phân loại 3
        tdbg3.Columns(COL3_AssetS3Name).Caption = rl3("Ten_phan_loai") & " 3" 'Tên phân loại 3
        tdbg3.Columns(COL3_Disabled).Caption = rl3("KSD") 'KSD
        tdbg4.Columns(COL4_AssetS4ID).Caption = rL3("Ma_phan_loai") & " 4" 'Mã phân loại 4
        tdbg4.Columns(COL4_AssetS4Name).Caption = rL3("Ten_phan_loai") & " 4" 'Tên phân loại 4
        tdbg4.Columns(COL4_Disabled).Caption = rL3("KSD") 'KSD
        tdbg5.Columns(COL5_AssetS5ID).Caption = rL3("Ma_phan_loai") & " 5" 'Mã phân loại 5
        tdbg5.Columns(COL5_AssetS5Name).Caption = rL3("Ten_phan_loai") & " 5" 'Tên phân loại 5
        tdbg5.Columns(COL5_Disabled).Caption = rL3("KSD") 'KSD
    End Sub



    Private Sub LoadTDBGrid(Optional ByVal FlagAdd As Boolean = False, Optional ByVal sKey As String = "")
        Dim sSQL As String = ""
        Select Case tabMain.SelectedIndex
            Case 0
                sSQL = "Select      AssetS1ID, AssetS1Name" & UnicodeJoin(gbUnicode) & " As AssetS1Name, Disabled, " & vbCrLf
                sSQL &= "           CreateUserID, CreateDate, LastModifyUserID, LastModifyDate" & vbCrLf
                sSQL &= "From       D02T1000 WITH(NOLOCK) " & vbCrLf
                sSQL &= "Order By   AssetS1ID"
                dtGrid1 = ReturnDataTable(sSQL)

                gbEnabledUseFind = dtGrid1.Rows.Count > 0

                If FlagAdd Then ' Thêm mới thì set Filter = "" và sFind =""
                    ResetFilter(tdbg1, sFilter1, bRefreshFilter1)
                    sFilter1 = New System.Text.StringBuilder("")
                End If

                LoadDataSource(tdbg1, dtGrid1, gbUnicode)
                ReLoadTDBGrid()

                If sKey <> "" Then
                    Dim dt1 As DataTable = dtGrid1.DefaultView.ToTable
                    Dim dr() As DataRow = dt1.Select("AssetS1ID = " & SQLString(sKey), dt1.DefaultView.Sort)
                    If dr.Length > 0 Then tdbg1.Row = dt1.Rows.IndexOf(dr(0))
                End If

                If Not tdbg1.Focused Then tdbg1.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới
            Case 1
                sSQL = "Select      AssetS2ID, AssetS2Name" & UnicodeJoin(gbUnicode) & " As AssetS2Name, Disabled, " & vbCrLf
                sSQL &= "           CreateUserID, CreateDate, LastModifyUserID, LastModifyDate" & vbCrLf
                sSQL &= "From       D02T2000 WITH(NOLOCK)" & vbCrLf
                sSQL &= "Order By   AssetS2ID"
                dtGrid2 = ReturnDataTable(sSQL)

                gbEnabledUseFind = dtGrid1.Rows.Count > 0

                If FlagAdd Then ' Thêm mới thì set Filter = "" và sFind =""
                    ResetFilter(tdbg2, sFilter2, bRefreshFilter2)
                    sFilter2 = New System.Text.StringBuilder("")
                End If

                LoadDataSource(tdbg2, dtGrid2, gbUnicode)
                ReLoadTDBGrid()

                If sKey <> "" Then
                    Dim dt1 As DataTable = dtGrid2.DefaultView.ToTable
                    Dim dr() As DataRow = dt1.Select("AssetS2ID = " & SQLString(sKey), dt1.DefaultView.Sort)
                    If dr.Length > 0 Then tdbg2.Row = dt1.Rows.IndexOf(dr(0))
                End If

                If Not tdbg2.Focused Then tdbg2.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới

            Case 2
                sSQL = "Select      AssetS3ID, AssetS3Name" & UnicodeJoin(gbUnicode) & " As AssetS3Name, Disabled, " & vbCrLf
                sSQL &= "           CreateUserID, CreateDate, LastModifyUserID, LastModifyDate" & vbCrLf
                sSQL &= "From       D02T3000 WITH(NOLOCK) " & vbCrLf
                sSQL &= "Order By   AssetS3ID"
                dtGrid3 = ReturnDataTable(sSQL)

                gbEnabledUseFind = dtGrid3.Rows.Count > 0

                If FlagAdd Then ' Thêm mới thì set Filter = "" và sFind =""
                    ResetFilter(tdbg3, sFilter3, bRefreshFilter3)
                    sFilter3 = New System.Text.StringBuilder("")
                End If

                LoadDataSource(tdbg3, dtGrid3, gbUnicode)
                ReLoadTDBGrid()

                If sKey <> "" Then
                    Dim dt1 As DataTable = dtGrid3.DefaultView.ToTable
                    Dim dr() As DataRow = dt1.Select("AssetS3ID = " & SQLString(sKey), dt1.DefaultView.Sort)
                    If dr.Length > 0 Then tdbg3.Row = dt1.Rows.IndexOf(dr(0))
                End If

                If Not tdbg3.Focused Then tdbg3.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới
            Case 3
                sSQL = "Select      AssetS4ID, AssetS4Name" & UnicodeJoin(gbUnicode) & " As AssetS4Name, Disabled, " & vbCrLf
                sSQL &= "           CreateUserID, CreateDate, LastModifyUserID, LastModifyDate" & vbCrLf
                sSQL &= "From       D02T4000 WITH(NOLOCK) " & vbCrLf
                sSQL &= "Order By   AssetS4ID"
                dtGrid4 = ReturnDataTable(sSQL)

                gbEnabledUseFind = dtGrid4.Rows.Count > 0

                If FlagAdd Then ' Thêm mới thì set Filter = "" và sFind =""
                    ResetFilter(tdbg4, sFilter4, bRefreshFilter4)
                    sFilter4 = New System.Text.StringBuilder("")
                End If

                LoadDataSource(tdbg4, dtGrid4, gbUnicode)
                ReLoadTDBGrid()

                If sKey <> "" Then
                    Dim dt1 As DataTable = dtGrid4.DefaultView.ToTable
                    Dim dr() As DataRow = dt1.Select("AssetS4ID = " & SQLString(sKey), dt1.DefaultView.Sort)
                    If dr.Length > 0 Then tdbg4.Row = dt1.Rows.IndexOf(dr(0))
                End If

                If Not tdbg4.Focused Then tdbg4.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới
            Case 4
                sSQL = "Select      AssetS5ID, AssetS5Name" & UnicodeJoin(gbUnicode) & " As AssetS5Name, Disabled, " & vbCrLf
                sSQL &= "           CreateUserID, CreateDate, LastModifyUserID, LastModifyDate" & vbCrLf
                sSQL &= "From       D02T5003 WITH(NOLOCK) " & vbCrLf
                sSQL &= "Order By   AssetS5ID"
                dtGrid5 = ReturnDataTable(sSQL)

                gbEnabledUseFind = dtGrid5.Rows.Count > 0

                If FlagAdd Then ' Thêm mới thì set Filter = "" và sFind =""
                    ResetFilter(tdbg5, sFilter5, bRefreshFilter5)
                    sFilter5 = New System.Text.StringBuilder("")
                End If

                LoadDataSource(tdbg5, dtGrid5, gbUnicode)
                ReLoadTDBGrid()

                If sKey <> "" Then
                    Dim dt1 As DataTable = dtGrid5.DefaultView.ToTable
                    Dim dr() As DataRow = dt1.Select("AssetS5ID = " & SQLString(sKey), dt1.DefaultView.Sort)
                    If dr.Length > 0 Then tdbg5.Row = dt1.Rows.IndexOf(dr(0))
                End If

                If Not tdbg5.Focused Then tdbg5.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới
        End Select
    End Sub

    Private Sub ReLoadTDBGrid(Optional ByVal bUseFilterBar As Boolean = False)
        Dim strFind As String = ""

        Select Case tabMain.SelectedIndex
            Case 0
                If sFilter1.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
                strFind &= sFilter1.ToString

                If Not chkShowDisabled1.Checked Then
                    If strFind <> "" Then strFind &= " And "
                    strFind &= "Disabled = 0"
                End If
                dtGrid1.DefaultView.RowFilter = strFind
            Case 1
                If sFilter2.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
                strFind &= sFilter2.ToString

                If Not chkShowDisabled2.Checked Then
                    If strFind <> "" Then strFind &= " And "
                    strFind &= "Disabled = 0"
                End If
                dtGrid2.DefaultView.RowFilter = strFind
            Case 2
                If sFilter3.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
                strFind &= sFilter3.ToString

                If Not chkShowDisabled3.Checked Then
                    If strFind <> "" Then strFind &= " And "
                    strFind &= "Disabled = 0"
                End If
                dtGrid3.DefaultView.RowFilter = strFind
            Case 3
                If sFilter4.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
                strFind &= sFilter4.ToString

                If Not chkShowDisabled4.Checked Then
                    If strFind <> "" Then strFind &= " And "
                    strFind &= "Disabled = 0"
                End If
                dtGrid4.DefaultView.RowFilter = strFind
            Case 4
                If sFilter5.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
                strFind &= sFilter5.ToString

                If Not chkShowDisabled5.Checked Then
                    If strFind <> "" Then strFind &= " And "
                    strFind &= "Disabled = 0"
                End If
                dtGrid5.DefaultView.RowFilter = strFind
        End Select

        ResetGrid()
    End Sub

    Private Sub ResetGrid()

        Select Case tabMain.SelectedIndex
            Case 0
                CheckMenu(_formIDPermission, TableToolStrip, tdbg1.RowCount, gbEnabledUseFind, False, ContextMenuStrip1)
                FooterTotalGrid(tdbg1, COL1_AssetS1ID)
            Case 1
                CheckMenu(_formIDPermission, TableToolStrip, tdbg2.RowCount, gbEnabledUseFind, False, ContextMenuStrip1, False, "D02F5607")
                FooterTotalGrid(tdbg2, COL2_AssetS2ID)
            Case 2
                CheckMenu(_formIDPermission, TableToolStrip, tdbg3.RowCount, gbEnabledUseFind, False, ContextMenuStrip1)
                FooterTotalGrid(tdbg3, COL3_AssetS3ID)
            Case 3
                CheckMenu(_formIDPermission, TableToolStrip, tdbg4.RowCount, gbEnabledUseFind, False, ContextMenuStrip1)
                FooterTotalGrid(tdbg4, COL4_AssetS4ID)
            Case 4
                CheckMenu(_formIDPermission, TableToolStrip, tdbg5.RowCount, gbEnabledUseFind, False, ContextMenuStrip1)
                FooterTotalGrid(tdbg5, COL5_AssetS5ID)
        End Select
    End Sub

    Private Sub tabMain_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabMain.SelectedIndexChanged
        LoadTDBGrid()
        VisibleMenuImport() '29/8/2019, Lê Thị Phú Hà:id 122897-Hỗ trợ import Excel cho mã phân loại 2 TSCĐ
    End Sub

#Region "Menu bar"

    Private Sub tsbAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbAdd.Click, tsmAdd.Click, mnsAdd.Click
        Dim f As New D02F3001
        With f
            .AssetID = ""
            .IndexTab = tabMain.SelectedIndex
            .FormState = EnumFormState.FormAdd
            .ShowDialog()
            If f.SavedOk Then LoadTDBGrid(True, .AssetID)
            .Dispose()
        End With
    End Sub

    Private Sub tsbView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbView.Click, tsmView.Click, mnsView.Click
        Dim f As New D02F3001
        With f
            Select Case tabMain.SelectedIndex
                Case 0
                    .AssetID = tdbg1.Columns(COL1_AssetS1ID).Text
                Case 1
                    .AssetID = tdbg2.Columns(COL2_AssetS2ID).Text
                Case 2
                    .AssetID = tdbg3.Columns(COL3_AssetS3ID).Text
                Case 3
                    .AssetID = tdbg4.Columns(COL4_AssetS4ID).Text
                Case 4
                    .AssetID = tdbg5.Columns(COL5_AssetS5ID).Text
            End Select
            .IndexTab = tabMain.SelectedIndex
            .FormState = EnumFormState.FormView
            .ShowDialog()
            .Dispose()
        End With
    End Sub

    Private Sub tsbEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbEdit.Click, tsmEdit.Click, mnsEdit.Click
        Dim f As New D02F3001
        With f
            Select Case tabMain.SelectedIndex
                Case 0
                    .AssetID = tdbg1.Columns(COL1_AssetS1ID).Text
                Case 1
                    .AssetID = tdbg2.Columns(COL2_AssetS2ID).Text
                Case 2
                    .AssetID = tdbg3.Columns(COL3_AssetS3ID).Text
                Case 3
                    .AssetID = tdbg4.Columns(COL4_AssetS4ID).Text
                Case 4
                    .AssetID = tdbg5.Columns(COL5_AssetS5ID).Text
            End Select
            .IndexTab = tabMain.SelectedIndex
            .FormState = EnumFormState.FormEdit
            .ShowDialog()
            If f.SavedOk = True Then LoadTDBGrid(False, .AssetID)
            .Dispose()
        End With
    End Sub

    Private Sub tsbDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbDelete.Click, tsmDelete.Click, mnsDelete.Click
        Dim sSQL As String = ""

        If AskDelete() = Windows.Forms.DialogResult.No Then Exit Sub
        Select Case tabMain.SelectedIndex
            Case 0
                If Not CheckBeforeDelete("AssetS1ID", tdbg1.Columns(COL1_AssetS1ID).Text) Then Exit Sub
                sSQL = "Delete D02T1000 Where AssetS1ID = " & SQLString(tdbg1.Columns(COL1_AssetS1ID).Text)
            Case 1
                If Not CheckBeforeDelete("AssetS2ID", tdbg2.Columns(COL2_AssetS2ID).Text) Then Exit Sub
                sSQL = "Delete D02T2000 Where AssetS2ID = " & SQLString(tdbg2.Columns(COL2_AssetS2ID).Text)
            Case 2
                If Not CheckBeforeDelete("AssetS3ID", tdbg3.Columns(COL3_AssetS3ID).Text) Then Exit Sub
                sSQL = "Delete D02T3000 Where AssetS3ID = " & SQLString(tdbg3.Columns(COL3_AssetS3ID).Text)
            Case 3
                If Not CheckBeforeDelete("AssetS4ID", tdbg4.Columns(COL4_AssetS4ID).Text) Then Exit Sub
                sSQL = "Delete D02T4000 Where AssetS4ID = " & SQLString(tdbg4.Columns(COL4_AssetS4ID).Text)
            Case 4
                If Not CheckBeforeDelete("AssetS5ID", tdbg5.Columns(COL5_AssetS5ID).Text) Then Exit Sub
                sSQL = "Delete D02T5003 Where AssetS5ID = " & SQLString(tdbg5.Columns(COL5_AssetS5ID).Text)
        End Select

        Dim bResult As Boolean = ExecuteSQL(sSQL)
        If bResult Then
            Select Case tabMain.SelectedIndex
                Case 0
                    DeleteGridEvent(tdbg1, dtGrid1, gbEnabledUseFind)
                Case 1
                    DeleteGridEvent(tdbg2, dtGrid2, gbEnabledUseFind)
                Case 2
                    DeleteGridEvent(tdbg3, dtGrid3, gbEnabledUseFind)
                Case 3
                    DeleteGridEvent(tdbg4, dtGrid4, gbEnabledUseFind)
                Case 4
                    DeleteGridEvent(tdbg5, dtGrid5, gbEnabledUseFind)
            End Select
            ResetGrid()
            DeleteOK()
        Else
            DeleteNotOK()
        End If
    End Sub

    Private Function CheckBeforeDelete(ByVal sField As String, ByVal AssetID As String) As Boolean
        Dim sSQL As String = ""
        sSQL = "Select Top 1 1 From D02T0001 WITH(NOLOCK) Where " & sField & "= " & SQLString(AssetID)
        If ExistRecord(sSQL) Then
            D99C0008.MsgCanNotDelete()
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub tsbSysInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbSysInfo.Click, tsmSysInfo.Click, mnsSysInfo.Click
        If tabMain.SelectedIndex = 0 Then
            ShowSysInfoDialog(Me,tdbg1.Columns(COL1_CreateUserID).Text, tdbg1.Columns(COL1_CreateDate).Text, tdbg1.Columns(COL1_LastModifyUserID).Text, tdbg1.Columns(COL1_LastModifyDate).Text)
        ElseIf tabMain.SelectedIndex = 1 Then
            ShowSysInfoDialog(Me,tdbg2.Columns(COL2_CreateUserID).Text, tdbg2.Columns(COL2_CreateDate).Text, tdbg2.Columns(COL2_LastModifyUserID).Text, tdbg2.Columns(COL2_LastModifyDate).Text)
        ElseIf tabMain.SelectedIndex = 2 Then
            ShowSysInfoDialog(Me, tdbg3.Columns(COL3_CreateUserID).Text, tdbg3.Columns(COL3_CreateDate).Text, tdbg3.Columns(COL3_LastModifyUserID).Text, tdbg3.Columns(COL3_LastModifyDate).Text)
        ElseIf tabMain.SelectedIndex = 3 Then
            ShowSysInfoDialog(Me, tdbg4.Columns(COL4_CreateUserID).Text, tdbg4.Columns(COL4_CreateDate).Text, tdbg4.Columns(COL4_LastModifyUserID).Text, tdbg4.Columns(COL4_LastModifyDate).Text)
        ElseIf tabMain.SelectedIndex = 4 Then
            ShowSysInfoDialog(Me, tdbg5.Columns(COL5_CreateUserID).Text, tdbg5.Columns(COL5_CreateDate).Text, tdbg5.Columns(COL5_LastModifyUserID).Text, tdbg5.Columns(COL5_LastModifyDate).Text)
        End If
    End Sub

    Private Sub tsbClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbClose.Click
        Me.Close()
    End Sub

    Private Sub tsbPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbPrint.Click, tsmPrint.Click, mnsPrint.Click

        'Dim report As New D99C1003
        'Đưa vể đầu tiên hàm In trước khi gọi AllowPrint()
		If Not AllowNewD99C2003(report, Me) Then Exit Sub
		'************************************
        Dim conn As New SqlConnection(gsConnectionString)
        Dim sReportName As String = ""
        Dim sSubReportName As String = "D02R0000"
        Dim sReportCaption As String = ""
        Dim sPathReport As String = ""
        Dim sSQL As String = ""
        Dim sSQLSub As String = ""

        Select Case tabMain.SelectedIndex
            Case 0
                sReportName = "D02R1000"
                sSQL = "Select * From D02T1000 WITH(NOLOCK) Order By AssetS1ID"
            Case 1
                sReportName = "D02R2000"
                sSQL = "Select * From D02T2000 WITH(NOLOCK) Order By AssetS2ID"
            Case 2
                sReportName = "D02R3000"
                sSQL = "Select * From D02T3000 WITH(NOLOCK) Order By AssetS3ID"
            Case 3
                sReportName = "D02R4000"
                sSQL = "Select * From D02T4000 WITH(NOLOCK) Order By AssetS4ID"
            Case 4
                sReportName = "D02R5000"
                sSQL = "Select * From D02T5003 WITH(NOLOCK) Order By AssetS5ID"
        End Select

        sReportCaption = rl3("Danh_muc_phan_loai") & " " & (tabMain.SelectedIndex + 1) & " - " & sReportName
        sPathReport = UnicodeGetReportPath(gbUnicode, D02Options.ReportLanguage, "") & sReportName & ".rpt"

        sSQLSub = "Select Top 1 * From D91T0025 WITH(NOLOCK)"
        UnicodeSubReport(sSubReportName, sSQLSub, , gbUnicode)

        With report
            .OpenConnection(conn)
            .AddSub(sSQLSub, sSubReportName & ".rpt")
            '.AddMain(sSQL)
            Select Case tabMain.SelectedIndex
                Case 0
                    .AddMain(dtGrid1.DefaultView.ToTable)
                Case 1
                    .AddMain(dtGrid2.DefaultView.ToTable)
                Case 2
                    .AddMain(dtGrid3.DefaultView.ToTable)
                Case 3
                    .AddMain(dtGrid4.DefaultView.ToTable)
                Case 4
                    .AddMain(dtGrid5.DefaultView.ToTable)
            End Select
            .PrintReport(sPathReport, sReportCaption)
        End With

    End Sub

    Private Sub chkShowDisabled1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowDisabled1.CheckedChanged
        If dtGrid1 Is Nothing Then Exit Sub
        ReLoadTDBGrid()
    End Sub

    Private Sub chkShowDisabled2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowDisabled2.CheckedChanged
        If dtGrid2 Is Nothing Then Exit Sub
        ReLoadTDBGrid()
    End Sub

    Private Sub chkShowDisabled3_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowDisabled3.CheckedChanged
        If dtGrid3 Is Nothing Then Exit Sub
        ReLoadTDBGrid()
    End Sub

    Private Sub chkShowDisabled4_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowDisabled4.CheckedChanged
        If dtGrid4 Is Nothing Then Exit Sub
        ReLoadTDBGrid()
    End Sub

    Private Sub chkShowDisabled5_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowDisabled5.CheckedChanged
        If dtGrid5 Is Nothing Then Exit Sub
        ReLoadTDBGrid()
    End Sub
#End Region

#Region "Grid"

    Private Sub tdbg1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg1.DoubleClick
        If tdbg1.FilterActive Then Exit Sub

        If tsbEdit.Enabled Then
            tsbEdit_Click(sender, Nothing)
        ElseIf tsbView.Enabled Then
            tsbView_Click(sender, Nothing)
        End If
    End Sub

    Private Sub tdbg1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg1.KeyDown
        If e.KeyCode = Keys.Enter Then tdbg1_DoubleClick(Nothing, Nothing)
        HotKeyCtrlVOnGrid(tdbg1, e)
    End Sub

    Private Sub tdbg1_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg1.FilterChange
        Try
            If (dtGrid1 Is Nothing) Then Exit Sub
            If bRefreshFilter1 Then Exit Sub 'set FilterText ="" thì thoát
            FilterChangeGrid(tdbg1, sFilter1)
            ReLoadTDBGrid()
        Catch ex As Exception
            'MessageBox.Show(ex.Message & " - " & ex.Source)
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub


    Private Sub tdbg2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg2.DoubleClick
        If tdbg2.FilterActive Then Exit Sub

        If tsbEdit.Enabled Then
            tsbEdit_Click(sender, Nothing)
        ElseIf tsbView.Enabled Then
            tsbView_Click(sender, Nothing)
        End If
    End Sub

    Private Sub tdbg2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg2.KeyDown
        If e.KeyCode = Keys.Enter Then tdbg2_DoubleClick(Nothing, Nothing)
        HotKeyCtrlVOnGrid(tdbg2, e)
    End Sub

    Private Sub tdbg2_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg2.FilterChange
        Try
            If (dtGrid2 Is Nothing) Then Exit Sub
            If bRefreshFilter2 Then Exit Sub 'set FilterText ="" thì thoát
            FilterChangeGrid(tdbg2, sFilter2)
            ReLoadTDBGrid()
        Catch ex As Exception
            'MessageBox.Show(ex.Message & " - " & ex.Source)
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub


    Private Sub tdbg3_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg3.DoubleClick
        If tdbg3.FilterActive Then Exit Sub

        If tsbEdit.Enabled Then
            tsbEdit_Click(sender, Nothing)
        ElseIf tsbView.Enabled Then
            tsbView_Click(sender, Nothing)
        End If
    End Sub

    Private Sub tdbg3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg3.KeyDown
        If e.KeyCode = Keys.Enter Then tdbg3_DoubleClick(Nothing, Nothing)
        HotKeyCtrlVOnGrid(tdbg3, e)
    End Sub

    Private Sub tdbg3_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg3.FilterChange
        Try
            If (dtGrid3 Is Nothing) Then Exit Sub
            If bRefreshFilter3 Then Exit Sub 'set FilterText ="" thì thoát
            FilterChangeGrid(tdbg3, sFilter3)
            ReLoadTDBGrid()
        Catch ex As Exception
            'MessageBox.Show(ex.Message & " - " & ex.Source)
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub
#End Region

    Private Sub tdbg4_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg4.FilterChange
        Try
            If (dtGrid4 Is Nothing) Then Exit Sub
            If bRefreshFilter4 Then Exit Sub
            FilterChangeGrid(tdbg4, sFilter4) 'Nếu có Lọc khi In
            ReLoadTDBGrid()
        Catch ex As Exception
            'Update 11/05/2011: Tạm thời có lỗi thì bỏ qua không hiện message
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub

    Private Sub tdbg4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg4.KeyDown
        If e.KeyCode = Keys.Enter Then tdbg4_DoubleClick(Nothing, Nothing)
        HotKeyCtrlVOnGrid(tdbg4, e)
    End Sub

    Private Sub tdbg4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg4.KeyPress
        If tdbg4.Columns(tdbg4.Col).ValueItems.Presentation = C1.Win.C1TrueDBGrid.PresentationEnum.CheckBox Then
            e.Handled = CheckKeyPress(e.KeyChar)
        ElseIf tdbg4.Splits(tdbg4.SplitIndex).DisplayColumns(tdbg4.Col).Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far Then
            e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End If
    End Sub

    Private Sub tdbg4_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg4.DoubleClick
        If tdbg4.FilterActive Then Exit Sub

        If tsbEdit.Enabled Then
            tsbEdit_Click(sender, Nothing)
        ElseIf tsbView.Enabled Then
            tsbView_Click(sender, Nothing)
        End If
    End Sub



    'Lưu ý: gọi hàm ResetFilter(tdbg4, sFilter, bRefreshFilter) tại btnFilter_Click và tsbListAll_Click
    'Bổ sung vào đầu sự kiện tdbg4_DoubleClick(nếu có) câu lệnh If tdbg4.RowCount <= 0 OrElse tdbg4.FilterActive Then Exit Sub


    Private Sub tdbg5_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg5.FilterChange
        Try
            If (dtGrid5 Is Nothing) Then Exit Sub
            If bRefreshFilter5 Then Exit Sub
            FilterChangeGrid(tdbg5, sFilter5) 'Nếu có Lọc khi In
            ReLoadTDBGrid()
        Catch ex As Exception
            'Update 11/05/2011: Tạm thời có lỗi thì bỏ qua không hiện message
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub

    Private Sub tdbg5_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg5.KeyDown
        If e.KeyCode = Keys.Enter Then tdbg5_DoubleClick(Nothing, Nothing)
        HotKeyCtrlVOnGrid(tdbg5, e)
    End Sub

    Private Sub tdbg5_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg5.DoubleClick
        If tdbg5.FilterActive Then Exit Sub

        If tsbEdit.Enabled Then
            tsbEdit_Click(sender, Nothing)
        ElseIf tsbView.Enabled Then
            tsbView_Click(sender, Nothing)
        End If
    End Sub

    '29/8/2019, Lê Thị Phú Hà:id 122897-Hỗ trợ import Excel cho mã phân loại 2 TSCĐ
    Private Sub VisibleMenuImport()
        Dim bVisible As Boolean = (tabMain.SelectedIndex = 1)
        ToolStripSeparator5.Visible = bVisible
        tsbImportData.Visible = bVisible
        ToolStripSeparator6.Visible = bVisible
        tsmImportData.Visible = bVisible
        ToolStripSeparator9.Visible = bVisible
        mnsImportData.Visible = bVisible
    End Sub
  
    'Import dữ liệu
    Private Sub tsbImportData_Click(sender As Object, e As EventArgs) Handles tsbImportData.Click, tsmImportData.Click, mnsImportData.Click
        '29/8/2019, Lê Thị Phú Hà:id 122897-Hỗ trợ import Excel cho mã phân loại 2 TSCĐ
        If CallShowDialogD80F2090(D02, "D02F5607", "D02F3000") Then LoadTDBGrid()
    End Sub
End Class