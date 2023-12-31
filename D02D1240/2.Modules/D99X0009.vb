'#######################################################################################
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Ngày cập nhật cuối cùng: 28/08/2012 
'# Người cập nhật cuối cùng: Nguyễn Thị Ánh
'# Diễn giải: Các hàm chung của nhóm G4
'Bổ sung Điều điện IsUseD15/IsUseD21=1 cho combo Nhân viên
'Bổ sung thêm "." vào hàm IndexIdCharactor và đổ nguồn Nhóm nhân viên - EmpGroupID
'Bổ sung thêm đổ nguồn Nhóm nhân viên cho dropdown - EmpGroupID
'Bổ sung thêm ký tự đặc biệt của hàm IndexIdCharactor
'Bổ sung đổ nguồn dropdown Khối, Phòng ban, Tổ nhóm, Nhân viên
'Bổ sung Nhóm nhân viên theo Đơn vị, Khối
'Sửa lại câu đổ nguồn theo Tiếng Anh: Khối, Phòng ban, Tổ nhóm, Nhóm nhân viên
'# Bổ sung WITH (NOLOCK) vào table, trong bảng D91T0000 23/9/2013
'#######################################################################################

Module D99X0009
#Region "General Function of G4"
    ''' <summary>
    ''' Câu SQL load SubReport cho G4 (có Combo Đơn vị hiện tại)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Function SQLSubReportG4(Optional ByVal sDivisionID As String = "") As String
        If sDivisionID = "" Then sDivisionID = gsDivisionID
        Dim sSQL As String
        sSQL = "Select CompanyName  as  Company, CompanyAddress as  Address, CompanyPhone  as  Telephone, CompanyFax  as  Fax, BankAccountName as BankAccountName, BankAccountNo,  VATCode" & vbCrLf
        sSQL &= " FROM D91V0016 " & vbCrLf
        sSQL &= " WHERE   DivisionID = " & SQLString(sDivisionID)

        Return sSQL
    End Function

#Region "Đơn vị"
    Public Function ReturnTableDivisionIDD09(ByVal sModuleID As String, Optional ByVal bHavePercent As Boolean = False, Optional ByVal bUseUnicode As Boolean = False) As DataTable
        'Bổ sung Field Unicode
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, bUseUnicode)
        '***************
        Dim sSQL As String = "--Do nguon Don vi" & vbCrLf
        sSQL &= "Select Distinct T99.DivisionID as DivisionID, "
        sSQL &= "T16.DivisionName" & sUnicode & " as DivisionName,"
        sSQL &= " 1 as DisplayOrder" & vbCrLf
        sSQL &= " From " & sModuleID & "T9999 T99 WITH(NOLOCK) Inner Join D91T0016 T16 WITH(NOLOCK) On T99.DivisionID = T16.DivisionID " & vbCrLf
        ' Nếu là nhóm G4 thì lấy table D91T0061
        sSQL &= " Inner Join D91T0061 T60 WITH(NOLOCK) On T99.DivisionID = T60.DivisionID " & vbCrLf
        sSQL &= " Where T16.Disabled = 0 And T60.UserID = '" & gsUserID & "'" & vbCrLf
        If bHavePercent Then
            sSQL &= " Union All " & vbCrLf
            sSQL &= " Select '%' as DivisionID," & sLanguage & " as DivisionName ,0 as DisplayOrder" & vbCrLf
        End If
        sSQL &= " Order By DisplayOrder, DivisionName"

        Dim dt As DataTable
        dt = ReturnDataTable(sSQL)
        Return dt
    End Function

    ''' <summary>
    ''' Load data of combo DivisionID in Inquiry/Transaction/List (no Report) 
    ''' </summary>
    ''' <param name="tdbcDivisionID">combo DivisionID</param>
    ''' <param name="sModuleID">Dxx</param>
    ''' <param name="bHavePercent">True: contain %; False(Default): no contain %</param>
    ''' <remarks></remarks>
    ''' 
    Public Sub LoadCboDivisionIDD09(ByVal tdbcDivisionID As C1.Win.C1List.C1Combo, ByVal sModuleID As String, Optional ByVal bHavePercent As Boolean = False, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbcDivisionID, ReturnTableDivisionIDD09(sModuleID, bHavePercent, bUseUnicode), bUseUnicode)
        tdbcDivisionID.DisplayMember = "DivisionName"
        tdbcDivisionID.ValueMember = "DivisionID"
    End Sub


    ''' <summary>
    ''' Load data of combo DivisionID in Report
    ''' </summary>
    ''' <param name="tdbc">Combo DivisionID</param>
    ''' <param name="sModuleID">ModuleID (Dxx)</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub LoadCboDivisionIDReportD09(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal sModuleID As String, Optional ByVal bUseUnicode As Boolean = False)

        'Bổ sung Field Unicode
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, bUseUnicode)
        '***************

        Dim sSQL As String = ""
        sSQL = "Select Distinct T99.DivisionID as DivisionID, T16.DivisionName" & sUnicode & " as DivisionName,1 as DisplayOrder "
        sSQL &= " From " & sModuleID & "T9999 T99 WITH(NOLOCK) Inner Join D91T0016 T16 WITH(NOLOCK) On T99.DivisionID = T16.DivisionID "
        ' Nếu là nhóm G4 thì lấy table D91T0061
        sSQL &= " Inner Join D91T0061 T60 WITH(NOLOCK) On T99.DivisionID = T60.DivisionID "
        sSQL &= " Where T16.Disabled = 0 And T60.UserID = '" & gsUserID & "' "
        If giModuleAdmin = 1 Then
            sSQL &= " Union All " & vbCrLf
            sSQL &= " Select '%' as DivisionID," & sLanguage & " as DivisionName,0 as DisplayOrder " & vbCrLf
        End If
        sSQL &= " Order By DisplayOrder, DivisionName"

        LoadDataSource(tdbc, sSQL, bUseUnicode)

        tdbc.DisplayMember = "DivisionName"
        tdbc.ValueMember = "DivisionID"
    End Sub
#End Region

    Private Function ReturnFilter(ByVal arrValue() As String, ByVal arrField() As String) As String
        Dim sFilter As String = ""
        If arrValue.Length = 0 OrElse arrValue.Length <> arrField.Length Then
            D99C0008.MsgL3("Điều kiện lọc không đúng.")
            Return sFilter
        End If

        For i As Integer = 0 To arrValue.Length - 1
            If arrValue(i) <> "%" Then
                If sFilter <> "" Then sFilter &= " And "
                sFilter &= arrField(i) & " =" & SQLString(arrValue(i))
            End If
        Next
        If sFilter <> "" Then sFilter = arrField(0) & " ='%' or (" & sFilter & ")"
        Return sFilter
    End Function

#Region "Khối"
    'Load tdbcBlockID: Khối
    Public Function ReturnTableBlockID(Optional ByVal bCurDivision As Boolean = False, Optional ByVal bHavePercent As Boolean = True, Optional ByVal bUseUnicode As Boolean = False) As DataTable

        'Bổ sung Field Unicode
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, bUseUnicode)
        '***************

        Dim sSQL As String = "--Do nguon Khoi" & vbCrLf
        If bHavePercent Then
            'sSQL = "SELECT 	'%' As BlockID, '<" & r("Tat_ca") & ">' As BlockName, '%' As DivisionID,0 as DisplayOrder" & vbCrLf
            sSQL &= "SELECT 	'%' As BlockID, " & sLanguage & " As BlockName, '%' As DivisionID,  0 as BlockDisplayOrder, 0 as DisplayOrder" & vbCrLf
            sSQL &= "UNION" & vbCrLf
        End If
        sSQL &= "SELECT BlockID, BlockName" & IIf(geLanguage = EnumLanguage.English, "01", "").ToString & sUnicode & " As BlockName, DivisionID, BlockDisplayOrder, 1 as DisplayOrder" & vbCrLf
        sSQL &= "FROM 	D09T1140 WITH(NOLOCK) WHERE	Disabled = 0 " & vbCrLf
        If bCurDivision Then sSQL &= " And DivisionID = " & SQLString(gsDivisionID) & vbCrLf
        sSQL &= "ORDER BY 	DisplayOrder, BlockDisplayOrder, BlockName"
        Return ReturnDataTable(sSQL)
    End Function


    'Load tdbcBlockID by Current DivisionID: no have tdbcDivisionID
    Public Sub LoadtdbcBlockID(ByVal tdbc As C1.Win.C1List.C1Combo, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbc, ReturnTableBlockID(True, , bUseUnicode), bUseUnicode) 'Load by Current Division
        ' tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
        tdbc.DisplayMember = "BlockName"
        tdbc.ValueMember = "BlockID"
    End Sub

    'Load tdbcBlockID by tdbcDivisionID
    Public Sub LoadtdbcBlockID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        If sDivisionID = "%" Then
            LoadDataSource(tdbc, dtOriginal.Copy, bUseUnicode)
        Else
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DivisionID='%' or DivisionID =" & SQLString(sDivisionID), True), bUseUnicode)
        End If
        Try
            tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False 'Không cần design
        Catch ex As Exception

        End Try

        tdbc.DisplayMember = "BlockName"
        tdbc.ValueMember = "BlockID"
    End Sub

    Public Sub LoadtdbdBlockID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbd, ReturnTableBlockID(True, False, bUseUnicode), bUseUnicode) 'Load by Current Division
    End Sub

    Public Sub LoadtdbdBlockID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sDivisionID}
        Dim sField() As String = {"DivisionID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField), True), bUseUnicode)
    End Sub
#End Region
   
#Region "Phòng ban"
    '************************

    'Load tdbcDepartmentID: Phòng ban
    'bCurDivision=True: by Current DivisionID
    'bCurDivision=False: by tdbcDivisionID
    Public Function ReturnTableDepartmentID(Optional ByVal bCurDivision As Boolean = False, Optional ByVal bHavePercent As Boolean = True, Optional ByVal bUseUnicode As Boolean = False) As DataTable

        'Bổ sung Field Unicode
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, bUseUnicode)
        '***************
        Dim sSQL As String = "--Do nguon Phong ban" & vbCrLf
        If bHavePercent Then
            sSQL &= "SELECT 	'%' As DepartmentID, " & sLanguage & " As DepartmentName, '%' As DivisionID, '%' As BlockID,0 as DepDisplayOrder, 0 as DisplayOrder " & vbCrLf
            sSQL &= "UNION" & vbCrLf
        End If
        sSQL &= "SELECT DepartmentID, DepartmentName" & IIf(geLanguage = EnumLanguage.English, "01", "").ToString & sUnicode & " As DepartmentName, DivisionID,BlockID, DepDisplayOrder, 1 as DisplayOrder" & vbCrLf
        sSQL &= "FROM 	D91T0012 WITH(NOLOCK) WHERE	Disabled = 0 " & vbCrLf
        If bCurDivision Then sSQL &= " And DivisionID = " & SQLString(gsDivisionID) & vbCrLf
        sSQL &= "ORDER BY DisplayOrder, DepDisplayOrder, DepartmentName"

        Return ReturnDataTable(sSQL)
    End Function

    ''' <summary>
    ''' Load tdbcDepartmentID by BlockID and gsDivisionID 
    ''' </summary>
    ''' <param name="tdbc">tdbcDepartmentID</param>
    ''' <param name="dtOriginal">get value from ReturnTableDepartmentID(True) function</param>
    ''' <param name="sBlockID">value of tdbcBlockID</param>
    ''' <remarks></remarks>
    Public Sub LoadtdbcDepartmentID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, Optional ByVal bUseUnicode As Boolean = False)
        If sBlockID = "%" Then
            LoadDataSource(tdbc, dtOriginal.Copy, bUseUnicode)
        Else
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or BlockID =" & SQLString(sBlockID)), bUseUnicode)
        End If
        'tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
        Try
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
        Catch ex As Exception

        End Try

        tdbc.DisplayMember = "DepartmentName"
        tdbc.ValueMember = "DepartmentID"
    End Sub

    ''' <summary>
    ''' Load tdbcDepartmentID by BlockID and tdbcDivisionID 
    ''' </summary>
    ''' <param name="tdbc">tdbcDepartmentID</param>
    ''' <param name="dtOriginal">get value from ReturnTableDepartmentID() function</param>
    ''' <param name="sBlockID">tdbcBlockID</param>
    ''' <param name="sDivisionID">tdbcDivisionID</param>
    ''' <remarks></remarks>
    Public Sub LoadtdbcDepartmentID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        If sDivisionID = "%" And sBlockID = "%" Then 'No Filter
            LoadDataSource(tdbc, dtOriginal.Copy, bUseUnicode)
        ElseIf sDivisionID = "%" And sBlockID <> "%" Then 'Filter by BlockID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or BlockID =" & SQLString(sBlockID), True), bUseUnicode)
        ElseIf sDivisionID <> "%" And sBlockID = "%" Then 'Filter by DivisionID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or DivisionID =" & SQLString(sDivisionID), True), bUseUnicode)
        Else
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "DepartmentID='%' or (DivisionID =" & SQLString(sDivisionID) & " And BlockID =" & SQLString(sBlockID) & ")", True), bUseUnicode)
        End If
        Try
            tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
        Catch ex As Exception
            
        End Try
        tdbc.DisplayMember = "DepartmentName"
        tdbc.ValueMember = "DepartmentID"
    End Sub

    Public Sub LoadtdbdDepartmentID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID}
        Dim sField() As String = {"BlockID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField), True), bUseUnicode)
    End Sub

    Public Sub LoadtdbdDepartmentID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDivisionID}
        Dim sField() As String = {"BlockID", "DivisionID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField), True), bUseUnicode)
    End Sub
#End Region

#Region "Tổ nhóm"
    'Load tdbcTeamID: Tổ nhóm
    'bCurDivision=True: by Current DivisionID
    'bCurDivision=False: by tdbcDivisionID
    Public Function ReturnTableTeamID(Optional ByVal bCurDivision As Boolean = False, Optional ByVal bHavePercent As Boolean = True, Optional ByVal bUseUnicode As Boolean = False) As DataTable
        'Bổ sung Field Unicode
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, bUseUnicode)
        '***************

        Dim sSQL As String = "--Do nguon To nhom" & vbCrLf

        If bHavePercent Then
            sSQL &= "SELECT 	'%' As TeamID, " & sLanguage & " As TeamName, '%' As DivisionID , '%' As BlockID, '%' As DepartmentID, 0 As TeamDisplayOrder, 0 As DisplayOrder " & vbCrLf
            sSQL &= "UNION" & vbCrLf
        End If
        sSQL &= "SELECT  D01.TeamID, D01.TeamName" & IIf(geLanguage = EnumLanguage.English, "01", "").ToString & sUnicode & " as TeamName, D02.DivisionID, D02.BlockID, D01.DepartmentID, TeamDisplayOrder, 1 As DisplayOrder" & vbCrLf
        sSQL &= "FROM 	D09T0227 D01 WITH(NOLOCK) " & vbCrLf
        sSQL &= "INNER JOIN D91T0012 D02 WITH(NOLOCK) On D02.DepartmentID = D01.DepartmentID" & vbCrLf
        sSQL &= "WHERE	D01.Disabled = 0 " & vbCrLf
        If bCurDivision Then sSQL &= " And D02.DivisionID = " & SQLString(gsDivisionID) & vbCrLf
        sSQL &= "ORDER BY  DisplayOrder,TeamDisplayOrder, TeamName"
        Return ReturnDataTable(sSQL)
    End Function

    ''' <summary>
    ''' Load tdbcTeamID by BlockID and gsDivisionID 
    ''' </summary>
    ''' <param name="tdbc">tdbcTeamID</param>
    ''' <param name="dtOriginal">get value from ReturnTableTeamID(True) function</param>
    ''' <param name="sBlockID">value of tdbcBlockID</param>
    ''' <remarks></remarks>
    Public Sub LoadtdbcTeamID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, Optional ByVal bUseUnicode As Boolean = False)
        If sDepartmentID = "%" And sBlockID = "%" Then 'No Filter
            LoadDataSource(tdbc, dtOriginal.Copy, bUseUnicode)
        ElseIf sDepartmentID = "%" And sBlockID <> "%" Then 'Filter by BlockID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or BlockID =" & SQLString(sBlockID), True), bUseUnicode)
        ElseIf sDepartmentID <> "%" And sBlockID = "%" Then 'Filter by Department
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or DepartmentID =" & SQLString(sDepartmentID), True), bUseUnicode)
        Else
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or (DepartmentID =" & SQLString(sDepartmentID) & " And BlockID =" & SQLString(sBlockID) & ")", True), bUseUnicode)
        End If
        'tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
        Try
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentID").Visible = False
        Catch ex As Exception
            
        End Try
        tdbc.DisplayMember = "TeamName"
        tdbc.ValueMember = "TeamID"
    End Sub

    Public Sub LoadtdbdTeamID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDepartmentID}
        Dim sField() As String = {"BlockID", "DepartmentID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField), True), bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load tdbcTeamID by BlockID and tdbcDivisionID 
    ''' </summary>
    ''' <param name="tdbc">tdbcTeamID</param>
    ''' <param name="dtOriginal">get value from ReturnTableTeamID() function</param>
    ''' <param name="sBlockID">tdbcBlockID</param>
    ''' <param name="sDivisionID">tdbcDivisionID</param>
    ''' <remarks></remarks>
    Public Sub LoadtdbcTeamID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)

        If sDivisionID = "%" And sBlockID = "%" And sDepartmentID = "%" Then 'No Filter
            LoadDataSource(tdbc, dtOriginal.Copy, bUseUnicode)

        ElseIf sDivisionID = "%" And sBlockID <> "%" And sDepartmentID = "%" Then 'Filter by BlockID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or BlockID =" & SQLString(sBlockID), True), bUseUnicode)

        ElseIf sDivisionID = "%" And sBlockID <> "%" And sDepartmentID <> "%" Then 'Filter by BlockID and DepartmentID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or (DepartmentID =" & SQLString(sDepartmentID) & " And BlockID =" & SQLString(sBlockID) & ")", True), bUseUnicode)

        ElseIf sDivisionID = "%" And sBlockID = "%" And sDepartmentID <> "%" Then 'Filter by DepartmentID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or DepartmentID =" & SQLString(sDepartmentID), True), bUseUnicode)

        ElseIf sDivisionID <> "%" And sBlockID = "%" And sDepartmentID = "%" Then 'Filter by DivisionID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or DivisionID =" & SQLString(sDivisionID), True), bUseUnicode)

        ElseIf sDivisionID <> "%" And sBlockID = "%" And sDepartmentID <> "%" Then 'Filter by DivisionID and DepartmentID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or (DepartmentID =" & SQLString(sDepartmentID) & " And DivisionID =" & SQLString(sDivisionID) & ")", True), bUseUnicode)

        ElseIf sDivisionID <> "%" And sBlockID <> "%" And sDepartmentID = "%" Then 'Filter by DivisionID and BlockID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or (DivisionID =" & SQLString(sDivisionID) & " And BlockID =" & SQLString(sBlockID) & ")", True), bUseUnicode)
        Else 'Filter by DivisionID and BlockID and DepartmentID
            LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, "TeamID='%' or (DivisionID =" & SQLString(sDivisionID) & " And BlockID =" & SQLString(sBlockID) & " And DepartmentID =" & SQLString(sDepartmentID) & ")", True), bUseUnicode)
        End If
        Try
            tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentID").Visible = False
        Catch ex As Exception
            
        End Try
        tdbc.DisplayMember = "TeamName"
        tdbc.ValueMember = "TeamID"
    End Sub


    Public Sub LoadtdbdTeamID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDepartmentID, sDivisionID}
        Dim sField() As String = {"BlockID", "DepartmentID", "DivisionID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField), True), bUseUnicode)
    End Sub
#End Region

#Region "Hình thức làm việc"
    'Load tdbcWorkingStatusID: Hình thức làm việc
    Public Function ReturnTableWorkingStatusID(Optional ByVal bHavePercent As Boolean = True, Optional ByVal bUseUnicode As Boolean = False) As DataTable
        'Bổ sung Field Unicode
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, bUseUnicode)
        '***************

        Dim sSQL As String = "--Do nguon Hinh thuc lam viec" & vbCrLf
        If bHavePercent Then
            sSQL &= "SELECT 	'%' As WorkingStatusID, " & sLanguage & "  As WorkingStatusName,0 as DisplayOrder" & vbCrLf
            sSQL &= "UNION" & vbCrLf
        End If
        sSQL &= "SELECT  WorkingStatusID, WorkingStatusName" & sUnicode & " as WorkingStatusName,1 as DisplayOrder" & vbCrLf
        sSQL &= "FROM 	D09T0070 WITH(NOLOCK) " & vbCrLf
        sSQL &= "WHERE	Disabled = 0 " & vbCrLf
        sSQL &= "ORDER BY 	DisplayOrder,WorkingStatusName"

        Return ReturnDataTable(sSQL)
    End Function

    'Load tdbcWorkingStatusID 
    Public Sub LoadtdbcWorkingStatusID(ByVal tdbc As C1.Win.C1List.C1Combo, Optional ByVal bHavePercent As Boolean = True, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbc, ReturnTableWorkingStatusID(bHavePercent, bUseUnicode), bUseUnicode) 'Load by Current Division
        tdbc.DisplayMember = "WorkingStatusName"
        tdbc.ValueMember = "WorkingStatusID"
    End Sub

    Public Sub LoadtdbdWorkingStatusID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, Optional ByVal bUseUnicode As Boolean = False)
        LoadDataSource(tdbd, ReturnTableWorkingStatusID(False, bUseUnicode), bUseUnicode)
    End Sub
#End Region

#Region "Nhân viên"
    'Load tdbcEmployeeID: Nhân viên
    'bCurDivision=True: by Current DivisionID
    'bCurDivision=False: by tdbcDivisionID
    'sModuleID: Dxx
    Public Function ReturnTableEmployeeID(Optional ByVal bCurDivision As Boolean = False, Optional ByVal bHavePercent As Boolean = True, Optional ByVal bUseUnicode As Boolean = False, Optional ByVal sModuleID As String = "") As DataTable
        'Bổ sung Field Unicode
        Dim sUnicode As String = ""
        Dim sLanguage As String = ""
        UnicodeAllString(sUnicode, sLanguage, bUseUnicode)
        '***************
        Dim sSQL As String = "--Do nguon Nhan vien" & vbCrLf
        If bHavePercent Then
            sSQL &= "SELECT 	'%' As EmployeeID, " & sLanguage & "  As EmployeeName, '%' As DivisionID , '%' As BlockID, '%' As DepartmentID, '%' as  TeamID, '%' as  WorkingStatusID, '%' as EmpGroupID, 0 as EmpDisplayOrder, 0 as DisplayOrder" & vbCrLf
            sSQL &= "UNION" & vbCrLf
        End If
        sSQL &= "SELECT 	D01.EmployeeID,Isnull(D01.LastName" & sUnicode & ",'') + CASE WHEN  D01.MiddleName" & sUnicode & " ='' THEN '' ELSE ' ' + D01.MiddleName" & sUnicode & " END  + ' '+ Isnull(D01.FirstName" & sUnicode & ",'') as EmployeeName, D01.DivisionID, D02.BlockID,  D01.DepartmentID, D01.TeamID, D01.WorkingTypeID AS WorkingStatusID, D01.EmpGroupID, EmpDisplayOrder, 1 as DisplayOrder  " & vbCrLf
        sSQL &= "FROM 	D09T0201 D01 WITH (NOLOCK) " & vbCrLf
        sSQL &= "INNER JOIN D91T0012 D02  WITH(NOLOCK) ON D01.DepartmentID = D02.DepartmentID" & vbCrLf
        sSQL &= "WHERE	D01.Disabled = 0 " & vbCrLf
        If bCurDivision Then sSQL &= " And D02.DivisionID = " & SQLString(gsDivisionID) & vbCrLf
        If sModuleID.Contains("21") OrElse sModuleID.Contains("15") Then sSQL &= " And D01.IsUseD" & Microsoft.VisualBasic.Right(sModuleID, 2) & " = 1" & vbCrLf
        sSQL &= "ORDER BY DisplayOrder, EmpDisplayOrder, EmployeeName"
        Return ReturnDataTable(sSQL)
    End Function


    Public Sub LoadtdbcEmployeeID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, ByVal sWorkingStatusID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sFilter As String = ""
        If sBlockID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "BlockID =" & SQLString(sBlockID)
        End If
        If sDepartmentID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "DepartmentID =" & SQLString(sDepartmentID)
        End If
        If sTeamID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "TeamID =" & SQLString(sTeamID)
        End If
        If sWorkingStatusID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "WorkingStatusID =" & SQLString(sWorkingStatusID)
        End If
        If sFilter <> "" Then sFilter = "EmployeeID='%' or (" & sFilter & ")"
        LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, sFilter, True), gbUnicode)

        'tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
        Try
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentID").Visible = False
            tdbc.Splits(0).DisplayColumns("TeamID").Visible = False
            tdbc.Splits(0).DisplayColumns("WorkingStatusID").Visible = False
        Catch ex As Exception
            
        End Try
        tdbc.DisplayMember = "EmployeeName"
        tdbc.ValueMember = "EmployeeID"
    End Sub

    Public Sub LoadtdbdEmployeeID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, ByVal sWorkingStatusID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDepartmentID, sTeamID, sWorkingStatusID}
        Dim sField() As String = {"BlockID", "DepartmentID", "TeamID", "WorkingStatusID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField), True), bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load tdbcEmployeeID by BlockID,DepartmentID,TeamID,WorkingStatusID and tdbcDivisionID 
    ''' </summary>
    ''' <param name="tdbc">tdbcEmployeeID</param>
    ''' <param name="dtOriginal">get value from ReturnTableEmployeeID() function</param>
    ''' <param name="sBlockID">tdbcBlockID</param>
    ''' <param name="sDivisionID">tdbcDivisionID</param>
    ''' <remarks></remarks>
    Public Sub LoadtdbcEmployeeID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, ByVal sWorkingStatusID As String, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sFilter As String = ""
        If sDivisionID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "DivisionID =" & SQLString(sDivisionID)
        End If
        If sBlockID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "BlockID =" & SQLString(sBlockID)
        End If
        If sDepartmentID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "DepartmentID =" & SQLString(sDepartmentID)
        End If
        If sTeamID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "TeamID =" & SQLString(sTeamID)
        End If
        If sWorkingStatusID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "WorkingStatusID =" & SQLString(sWorkingStatusID)
        End If
        If sFilter <> "" Then sFilter = "EmployeeID='%' or (" & sFilter & ")"
        LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, sFilter, True), gbUnicode)

        Try
            tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentID").Visible = False
            tdbc.Splits(0).DisplayColumns("TeamID").Visible = False
            tdbc.Splits(0).DisplayColumns("WorkingStatusID").Visible = False
        Catch ex As Exception

        End Try
        tdbc.DisplayMember = "EmployeeName"
        tdbc.ValueMember = "EmployeeID"
    End Sub

    Public Sub LoadtdbdEmployeeID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, ByVal sWorkingStatusID As String, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDepartmentID, sTeamID, sWorkingStatusID, sDivisionID}
        Dim sField() As String = {"BlockID", "DepartmentID", "TeamID", "WorkingStatusID", "DivisionID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField), True), bUseUnicode)
    End Sub

    ''' <summary>
    ''' Load tdbcEmployeeID by BlockID,DepartmentID,TeamID,WorkingStatusID, EmpGroupID and tdbcDivisionID 
    ''' </summary>
    ''' <param name="tdbc">tdbcEmployeeID</param>
    ''' <param name="dtOriginal">get value from ReturnTableEmployeeID() function</param>
    ''' <param name="sBlockID">tdbcBlockID</param>
    ''' <param name="sDivisionID">tdbcDivisionID</param>
    ''' <remarks></remarks>
    Public Sub LoadtdbcEmployeeID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, ByVal sWorkingStatusID As String, ByVal sDivisionID As String, ByVal sEmpGroupID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sFilter As String = ""
        If sDivisionID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "DivisionID =" & SQLString(sDivisionID)
        End If
        If sBlockID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "BlockID =" & SQLString(sBlockID)
        End If
        If sDepartmentID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "DepartmentID =" & SQLString(sDepartmentID)
        End If
        If sTeamID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "TeamID =" & SQLString(sTeamID)
        End If
        If sWorkingStatusID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "WorkingStatusID =" & SQLString(sWorkingStatusID)
        End If
        If sEmpGroupID <> "%" Then
            If sFilter <> "" Then sFilter &= " And "
            sFilter &= "EmpGroupID =" & SQLString(sEmpGroupID)
        End If
        If sFilter <> "" Then sFilter = "EmployeeID='%' or (" & sFilter & ")"
        LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, sFilter, True), gbUnicode)

        Try
            tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentID").Visible = False
            tdbc.Splits(0).DisplayColumns("TeamID").Visible = False
            tdbc.Splits(0).DisplayColumns("WorkingStatusID").Visible = False
        Catch ex As Exception
            
        End Try
        tdbc.DisplayMember = "EmployeeName"
        tdbc.ValueMember = "EmployeeID"
    End Sub

    Public Sub LoadtdbdEmployeeID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, ByVal sWorkingStatusID As String, ByVal sDivisionID As String, ByVal sEmpGroupID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDepartmentID, sTeamID, sWorkingStatusID, sDivisionID, sEmpGroupID}
        Dim sField() As String = {"BlockID", "DepartmentID", "TeamID", "WorkingStatusID", "DivisionID", "EmpGroupID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField), True), bUseUnicode)
    End Sub
#End Region

#Region "Nhóm nhân viên"
    ''' <summary>
    ''' Đổ nguồn Nhóm nhân viên
    ''' </summary>
    ''' <param name="bHavePercent"></param>
    ''' <param name="bUseUnicode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReturnTableEmpGroupID(Optional ByVal bHavePercent As Boolean = True, Optional ByVal bUseUnicode As Boolean = False, Optional ByVal bCurDivision As Boolean = False) As DataTable
        Dim sSQL As String = "--Do nguon Nhom nhan vien" & vbCrLf
        If bHavePercent Then
            sSQL &= "SELECT 	" & AllCode & " As EmpGroupID, " & AllName & "  As EmpGroupName, '%' As DepartmentID, '%' as  TeamID, '%' As DivisionID , '%' As BlockID, 0 as EGDisplayOrder, 0 As DisplayOrder" & vbCrLf
            sSQL &= "UNION" & vbCrLf
        End If
        sSQL &= "SELECT 	EmpGroupID, EmpGroupName" & gsLanguage & UnicodeJoin(gbUnicode) & " As EmpGroupName, T1.DepartmentID, T1. TeamID, T2.DivisionID, T2.BlockID, EGDisplayOrder, 1 As DisplayOrder  " & vbCrLf
        sSQL &= "FROM 	    D09T1210 T1 WITH(NOLOCK) " & vbCrLf
        sSQL &= " INNER JOIN	 D91T0012 T2 WITH(NOLOCK) ON	T1.DepartmentID = T2.DepartmentID" & vbCrLf
        sSQL &= "WHERE	    T1.Disabled = 0 " & vbCrLf
        If bCurDivision Then sSQL &= " And T2.DivisionID = " & SQLString(gsDivisionID) & vbCrLf
        sSQL &= "ORDER BY   DisplayOrder, EGDisplayOrder, EmpGroupName"
        Return ReturnDataTable(sSQL)
    End Function

    Public Sub LoadtdbcEmpGroupID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDepartmentID, sTeamID}
        Dim sField() As String = {"BlockID", "DepartmentID", "TeamID"}
        LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField), True), bUseUnicode)

        Try
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentID").Visible = False
            tdbc.Splits(0).DisplayColumns("TeamID").Visible = False
            tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
        Catch ex As Exception

        End Try
        tdbc.DisplayMember = "EmpGroupName"
        tdbc.ValueMember = "EmpGroupID"
    End Sub

    Public Sub LoadtdbcEmpGroupID(ByVal tdbc As C1.Win.C1List.C1Combo, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sDivisionID, sBlockID, sDepartmentID, sTeamID}
        Dim sField() As String = {"DivisionID", "BlockID", "DepartmentID", "TeamID"}
        LoadDataSource(tdbc, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField), True), bUseUnicode)

        Try
            tdbc.Splits(0).DisplayColumns("BlockID").Visible = False
            tdbc.Splits(0).DisplayColumns("DepartmentID").Visible = False
            tdbc.Splits(0).DisplayColumns("TeamID").Visible = False
            tdbc.Splits(0).DisplayColumns("DivisionID").Visible = False
        Catch ex As Exception

        End Try
        tdbc.DisplayMember = "EmpGroupName"
        tdbc.ValueMember = "EmpGroupID"
    End Sub

    Public Sub LoadtdbdEmpGroupID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDepartmentID, sTeamID}
        Dim sField() As String = {"BlockID", "DepartmentID", "TeamID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField), True), bUseUnicode)
        Try
            tdbd.DisplayColumns("DivisionID").Visible = False
            tdbd.DisplayColumns("BlockID").Visible = False
            tdbd.DisplayColumns("DepartmentID").Visible = False
            tdbd.DisplayColumns("TeamID").Visible = False
        Catch ex As Exception

        End Try
        tdbd.DisplayMember = "EmpGroupName"
        tdbd.ValueMember = "EmpGroupID"
    End Sub

    Public Sub LoadtdbdEmpGroupID(ByVal tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown, ByVal dtOriginal As DataTable, ByVal sBlockID As String, ByVal sDepartmentID As String, ByVal sTeamID As String, ByVal sDivisionID As String, Optional ByVal bUseUnicode As Boolean = False)
        Dim sValue() As String = {sBlockID, sDepartmentID, sTeamID, sDivisionID}
        Dim sField() As String = {"BlockID", "DepartmentID", "TeamID", "DivisionID"}
        LoadDataSource(tdbd, ReturnTableFilter(dtOriginal, ReturnFilter(sValue, sField), True), bUseUnicode)
        Try
            tdbd.DisplayColumns("DivisionID").Visible = False
            tdbd.DisplayColumns("BlockID").Visible = False
            tdbd.DisplayColumns("DepartmentID").Visible = False
            tdbd.DisplayColumns("TeamID").Visible = False
        Catch ex As Exception

        End Try
        tdbd.DisplayMember = "EmpGroupName"
        tdbd.ValueMember = "EmpGroupID"
    End Sub
#End Region
#End Region

#Region "Kiểm tra nhập Mã, Công thức nhóm G4"
    ''' <summary>
    ''' Thay đổi vị trí Select của chuỗi Vni
    ''' </summary>
    ''' <param name="str"></param>
    ''' <param name="posFrom">vị trí bắt đầu</param>
    ''' <param name="posTo">Số ký tự được Select</param>
    ''' <remarks>Không cần kiểm tra khi Unicode</remarks>
    Private Sub ChangePositionIndexVNI(ByVal str As String, ByRef posFrom As Integer, ByRef posTo As Integer)
        If str = "" OrElse posFrom < 0 OrElse posFrom >= str.Length - 1 Then Exit Sub

        Dim arrChar() As String = {"Â", "Á", "À", "Å", "Ä", "Ã", "Ù", "Ø", "Û", "Õ", "Ï", "É", "È", "Ú", "Ü", "Ë", "Ê"}
        Dim c As String = (str.Substring(posFrom, 1)).ToUpper
        '"Ö", "Ô"
        Select Case c
            Case "Ö", "Ô" 'Ö: Ư; Ô: Ơ - không tăng vị trí, ngược lại thì tăng thêm 1 vị trí
                If L3FindArrString(arrChar, (str.Substring(posFrom + 1, 1)).ToUpper) Then posTo = 2
            Case Else 'kiểm tra trong danh sách arrChar
                If L3FindArrString(arrChar, c) Then
                    If posFrom > 0 Then posFrom -= 1
                    posTo = 2
                End If
        End Select
    End Sub

    'Kiểm tra Button Đóng có đặt Tên "Close"
    Private Function CheckContinue(ByVal ctrl As Control) As Boolean
        Try
            Dim form As Form = CType(ctrl.TopLevelControl, Form)
            If form.Controls.ContainsKey("btnClose") Then
                Dim btnClose As Control = CType(form.Controls("btnClose"), System.Windows.Forms.Button)
                If btnClose Is Nothing Then Return True 'không có nút đóng
                If btnClose.Focused Then Return False 'Nhấn vào nút Đóng
                Dim arr() As String = ctrl.Tag.ToString.Split(" "c) 'Nhấn ALT + N
                If arr.Length > 2 Then Return False
                '************
            End If
        Catch ex As Exception

        End Try
        Return True
    End Function

    Private Sub txtID_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Dim txtID As TextBox = CType(sender, TextBox)
        If txtID.ReadOnly OrElse txtID.Enabled = False Then Exit Sub
        'Nếu nhấn đóng thì không cần hiện thông báo
        If CheckContinue(txtID) = False Then Exit Sub
        '************
        Dim posFrom As Integer
        'Bổ sung kiểm tra ký tự đặc biệt chuỗi truyền vào
        Dim arrTag() As String = Nothing
        If txtID.Tag IsNot Nothing Then arrTag = txtID.Tag.ToString.Split(" "c)
        Dim bFormula As Boolean = False
        Dim sCheckID As String = ""
        If arrTag IsNot Nothing OrElse arrTag.Length > 0 Then
            bFormula = L3Bool(arrTag(0))
            If arrTag.Length > 1 Then sCheckID = arrTag(1)
        End If
        If bFormula Then
            posFrom = IndexFormulaCharactor(txtID.Text, sCheckID)
        Else
            posFrom = IndexIdCharactor(txtID.Text, sCheckID)
        End If
        '***********************
        Dim posTo As Integer = 1
        If txtID.Font.Name.Contains("Lemon3") Then ChangePositionIndexVNI(txtID.Text, posFrom, posTo)
        Select Case posFrom
            Case -1 'thỏa điều kiện
            Case -2 'Vượt chiều dài
                'D99C0008.MsgL3("Chiều dài vượt quá quy định.")
                'txtID.SelectAll()
                'e.Cancel = True
            Case Else 'vi phạm
                D99C0008.MsgL3(r("Ma_co_ky_tu_khong_hop_le"))
                e.Cancel = True
                txtID.Select(posFrom, posTo)
        End Select
    End Sub

    Private Sub txtID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.Modifiers <> Keys.Alt Then Exit Sub
        If e.KeyCode = Keys.N Then
            Dim txtID As TextBox = CType(sender, TextBox)
            txtID.Tag = txtID.Tag.ToString & " True"
        End If
    End Sub

    ''' <summary>
    ''' Kiểm tra TextBox Nhập Mã/Công thức
    ''' </summary>
    ''' <param name="txtID">Control cần kiểm tra</param>
    ''' <param name="iLength">Chiều dài nhập liệu</param>
    ''' <param name="bFormula">Theo kiểu công thức (Default: False - cho Mã)</param>
    ''' <remarks>Cho một textbox, đối số bFormula : Kiểu kiểm tra là Mã hay Công thức</remarks>
    Public Sub CheckIdTextBoxG4(ByRef txtID As TextBox, Optional ByVal iLength As Integer = 20, Optional ByVal bFormula As Boolean = False, Optional ByVal sCheckID As String = "")
        txtID.CharacterCasing = CharacterCasing.Upper
        txtID.MaxLength = iLength

        If bFormula = False Then sCheckID &= ":;\{},()""<>=&~" 'TH nhập mã
        txtID.Tag = bFormula.ToString & IIf(sCheckID = "", "", " " & sCheckID).ToString
        AddHandler txtID.KeyDown, AddressOf txtID_KeyDown 'Khi nhấn ALT + N thì không kiểm tra
        AddHandler txtID.Validating, AddressOf txtID_Validating
    End Sub

    ''' <summary>
    ''' Kiểm tra nhiều TextBox Nhập Mã/Công thức
    ''' </summary>
    ''' <param name="txtID">Control cần kiểm tra</param>
    ''' <param name="iLength">Chiều dài nhập liệu</param>
    ''' <param name="bFormula">Theo kiểu công thức (Default: False - cho Mã)</param>
    ''' <remarks>Cho một textbox, đối số bFormula : Kiểu kiểm tra là Mã hay Công thức</remarks>
    Public Sub CheckIdTextBoxG4(ByRef txtID() As TextBox, Optional ByVal iLength As Integer = 20, Optional ByVal bFormula As Boolean = False, Optional ByVal sCheckID As String = "")
        For i As Integer = 0 To txtID.Length - 1
            CheckIdTextBoxG4(txtID(i), iLength, bFormula, sCheckID)
        Next
    End Sub

    ''' <summary>
    ''' Kiểm tra Mã hợp lệ 
    ''' </summary>
    ''' <param name="str">Chuỗi kiểm tra</param>
    ''' <returns>Vị trí ký tự vi phạm</returns>
    ''' <remarks></remarks>
    Private Function IndexIdCharactor(ByVal str As String, Optional ByVal sCheckID As String = ":;\{},()""<>=&~") As Integer
        'BackSpace: 8
        For Each c As Char In str
            Select Case AscW(c)
                Case 13, 10 'Mutiline của textbox và phím Enter
                    Continue For
                Case Is < 33, Is > 127, 37, 39, 42, 43, 45, 46, 47, 91, 93, 94 'Các ký tự đặc biệt: 37(%) 39(') 42 (*) 43 (+) 45 (-) 46 (.) 47 (/) 91([) 93(]) 94(^)
                    Return str.IndexOf(c)
            End Select
            If sCheckID <> "" Then
                Dim index As Integer = Strings.InStr(sCheckID, c, CompareMethod.Text)
                If index > 0 Then Return str.IndexOf(c)
            End If
        Next
        Return -1
    End Function


    '''' Kiểm tra công thức hợp lệ
    '''' </summary>
    '''' <param name="str">Chuỗi kiểm tra</param>
    '''' <returns>Vị trí ký tự vi phạm</returns>
    '''' <remarks></remarks>
    Private Function IndexFormulaCharactor(ByVal str As String, Optional ByVal sCheckID As String = "") As Integer
        '  If str.Length > iLength Then Return -2 'vượt chiều dài
        'BackSpace: 8
        For Each c As Char In str
            Select Case AscW(c)
                Case 13, 10 'Mutiline của textbox và phím Enter
                    Continue For
                Case Is < 33, Is > 127, 94 ''Các ký tự đặc biệt: 94(^)
                    Return str.IndexOf(c)
            End Select
            If sCheckID <> "" Then
                Dim index As Integer = Strings.InStr(sCheckID, c, CompareMethod.Text)
                If index > 0 Then Return str.IndexOf(c)
            End If
        Next
        Return -1
    End Function

    Private Function CheckFormulaCharactor(ByVal str As String) As Boolean
        Return IndexFormulaCharactor(str) >= 0
    End Function

#End Region

End Module
