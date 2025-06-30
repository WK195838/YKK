Imports System.Data
Imports System.Data.OleDb

Partial Class HR_VacationInfor
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wDepoID As String
    Dim wEMPID As String
    Dim xUserID As String                   ' 使用者ID

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()
        If Not Me.IsPostBack Then
            ShowPersonInfor()    '顯示個人基本資訊
            ShowVacationInfor()  '顯示休假資訊
            BBaseDate.Attributes("onclick") = "calendarPicker('form1.DBaseDate');"    '日期選擇
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        xUserID = UCase(Request.QueryString("pUserID"))
        '
        If UCase(xUserID) = "IT003" Or UCase(xUserID) = "GAS007" Then
        Else
            'BSimulate.Style("left") = -500 & "px"
        End If
        BBaseDate.Style("left") = -500 & "px"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     顯示個人基本資訊
    '**
    '*****************************************************************
    Sub ShowPersonInfor()
        Dim SQL As String
        Dim wLevel, wName, wDivision As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        '當日
        DDate.Text = CStr(DateTime.Now.Date)
        '基準日
        DBaseDate.Text = CStr(DateTime.Now.Date)
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++
        '++  設定部門/姓名權限
        '++  顯示個人基本資訊
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++
        '個人基本資訊
        DBDataSet1.Clear()
        SQL = "Select isnull(Name,'') as UserName, isnull(HRWDivName,'') as HRWDivName, "
        SQL = SQL + "EmpID, JobName, JobID "
        SQL = SQL + "From V_WorkTime_01 "
        SQL = SQL + "Where UserID='" & Request.QueryString("pUserID") & "' "
        SQL = SQL + "  And Not HRWDivName is Null "
        SQL = SQL + "  And Not Name is Null "
        SQL = SQL + "Order by Name "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        If DBTable1.Rows.Count > 0 Then
            wDivision = Trim(DBTable1.Rows(0).Item("HRWDivName"))
            wName = Trim(DBTable1.Rows(0).Item("UserName"))
            DEmpID.Text = Trim(DBDataSet1.Tables("M_Users").Rows(0).Item("EmpID"))
            DJobTitle.Text = Trim(DBDataSet1.Tables("M_Users").Rows(0).Item("JobName"))
            DJobCode.Text = Trim(DBDataSet1.Tables("M_Users").Rows(0).Item("JobID"))
        End If

        DBDataSet1.Clear()
        SQL = "Select * From M_Referp  "
        SQL = SQL + "Where Cat='1999'  "
        SQL = SQL + "  and DKey='" & "AUTHORITY-" & Request.QueryString("pUserID") & "' "
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "M_Referp")
        If DBDataSet1.Tables("M_Referp").Rows.Count > 0 Then
            If DBDataSet1.Tables("M_Referp").Rows(0).Item("Data") = "ALL" Then
                wLevel = "ALL"
            Else
                wLevel = "DIVISION"
            End If
        Else
            wLevel = "PERSON"
        End If

        DBDataSet1.Clear()
        DDivision.Items.Clear()
        DName.Items.Clear()
        If wLevel = "PERSON" Then
            '取得個人資訊
            SQL = "Select isnull(Name,'') as UserName, isnull(HRWDivName,'') as HRWDivName "
            SQL = SQL + "From V_WorkTime_01 "
            SQL = SQL + "Where UserID='" & Request.QueryString("pUserID") & "' "
            SQL = SQL + "  And Not HRWDivName is Null "
            SQL = SQL + "  And Not Name is Null "
            SQL = SQL + "Order by Name "
            Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter4.Fill(DBDataSet1, "M_Users")
            DBTable1 = DBDataSet1.Tables("M_Users")
            If DBTable1.Rows.Count > 0 Then
                Dim ListItem1 As New ListItem
                ListItem1.Text = Trim(DBTable1.Rows(0).Item("HRWDivName"))
                ListItem1.Value = Trim(DBTable1.Rows(0).Item("HRWDivName"))
                DDivision.Items.Add(ListItem1)
                '姓名
                Dim ListItem2 As New ListItem
                ListItem2.Text = Trim(DBTable1.Rows(0).Item("UserName"))
                ListItem2.Value = Trim(DBTable1.Rows(0).Item("UserName"))
                DName.Items.Add(ListItem2)
            End If
        Else
            If wLevel = "DIVISION" Then
                '取得所指定部門
                SQL = "Select * From M_Referp  "
                SQL = SQL + "Where Cat='1999'  "
                SQL = SQL + "  and DKey='" & "DIVISION-" & Request.QueryString("pUserID") & "' "
                Dim DBAdapter5 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter5.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = Trim(DBTable1.Rows(i).Item("Data"))
                    ListItem1.Value = Trim(DBTable1.Rows(i).Item("Data"))
                    If ListItem1.Value = wDivision Then ListItem1.Selected = True
                    DDivision.Items.Add(ListItem1)
                Next
                '取得個人資訊
                DBDataSet1.Clear()
                SQL = "Select isnull(Name,'') as UserName "
                SQL = SQL + "From V_WorkTime_01 "
                SQL = SQL + "Where HRWDivName = '" & DDivision.SelectedValue & "' "
                SQL = SQL + "  And Not HRWDivName is Null "
                SQL = SQL + "  And Not Name is Null "
                SQL = SQL + "Group by Name, EmpID, JobName, JobID "
                SQL = SQL + "Order by Name, EmpID, JobName, JobID "
                Dim DBAdapter6 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter6.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    '姓名
                    Dim ListItem2 As New ListItem
                    ListItem2.Text = Trim(DBTable1.Rows(i).Item("UserName"))
                    ListItem2.Value = Trim(DBTable1.Rows(i).Item("UserName"))
                    If ListItem2.Value = wName Then ListItem2.Selected = True
                    DName.Items.Add(ListItem2)
                Next
            Else
                '取得全部部門
                If wLevel = "ALL" Then
                    SQL = "Select isnull(HRWDivName,'') as HRWDivName From V_WorkTime_01 "
                    SQL = SQL + "Where Not HRWDivName is Null "
                    SQL = SQL + "  And Not Name is Null "
                    SQL = SQL + "Group by HRWDivName "
                    SQL = SQL + "Order by HRWDivName "
                    Dim DBAdapter7 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter7.Fill(DBDataSet1, "M_Users")
                    DBTable1 = DBDataSet1.Tables("M_Users")
                    For i = 0 To DBTable1.Rows.Count - 1
                        Dim ListItem1 As New ListItem
                        ListItem1.Text = Trim(DBTable1.Rows(i).Item("HRWDivName"))
                        ListItem1.Value = Trim(DBTable1.Rows(i).Item("HRWDivName"))
                        If ListItem1.Value = wDivision Then ListItem1.Selected = True
                        DDivision.Items.Add(ListItem1)
                    Next

                    '姓名
                    DBDataSet1.Clear()
                    SQL = "Select isnull(Name,'') as UserName "
                    SQL = SQL + "From V_WorkTime_01 "
                    SQL = SQL + "Where HRWDivName = '" & DDivision.SelectedValue & "' "
                    SQL = SQL + "  And Not HRWDivName is Null "
                    SQL = SQL + "  And Not Name is Null "
                    SQL = SQL + "Group by Name, EmpID, JobName, JobID "
                    SQL = SQL + "Order by Name, EmpID, JobName, JobID "
                    Dim DBAdapter8 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter8.Fill(DBDataSet1, "M_Users")
                    DBTable1 = DBDataSet1.Tables("M_Users")
                    For i = 0 To DBTable1.Rows.Count - 1
                        '姓名
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = Trim(DBTable1.Rows(i).Item("UserName"))
                        ListItem2.Value = Trim(DBTable1.Rows(i).Item("UserName"))
                        If ListItem2.Value = wName Then ListItem2.Selected = True
                        DName.Items.Add(ListItem2)
                    Next
                End If
            End If
        End If
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     顯示休假資訊
    '**
    '*****************************************************************
    Sub ShowVacationInfor()
        Dim InDate As String = ""
        Dim OldInDate As String = ""

        Dim NowYY As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        '入社日
        DBDataSet1.Clear()
        SQL = "Select DepoID, EmpID, StartDate As ArriveDate From V_WorkTime_01 "
        SQL = SQL & " Where HRWDivName = '" & DDivision.SelectedValue & "'"
        SQL = SQL & "   And Name       = '" & DName.SelectedValue & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "WorkTime")
        If DBDataSet1.Tables("WorkTime").Rows.Count > 0 Then
            DInDate.Text = DBDataSet1.Tables("WorkTime").Rows(0).Item("ArriveDate")
            wDepoID = DBDataSet1.Tables("WorkTime").Rows(0).Item("DepoID")
            wEMPID = DBDataSet1.Tables("WorkTime").Rows(0).Item("EmpID")
        End If
        '留職停薪
        DBDataSet1.Clear()
        SQL = "Select Days From HR_StopJobTime "
        SQL = SQL & " Where DepoCode = '" & wDepoID & "'"
        SQL = SQL & "   And EmpID    = '" & wEMPID & "'"
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "StopJob")
        If DBDataSet1.Tables("StopJob").Rows.Count > 0 Then
            OldInDate = DInDate.Text

            InDate = DateAdd(DateInterval.Day, DBDataSet1.Tables("StopJob").Rows(0).Item("Days"), CDate(DInDate.Text))
            DInDate.Text = CDate(InDate).Year.ToString + "/" + CDate(InDate).Month.ToString + "/" + CDate(InDate).Day.ToString
            '調閱請假
            LStopJobList.Style("left") = 275 & "px"
            LStopJobList.NavigateUrl = "HR_StopJobInforList_01.aspx?pDepoID=" + wDepoID + "&pEmpID=" + wEMPID
        Else
            '調閱請假
            LStopJobList.Style("left") = -500 & "px"
        End If
        '
        OleDbConnection1.Close()
        '基準年
        NowYY = CStr(CDate(DBaseDate.Text).Year)
        '年資計算
        DNensi.Text = CalNensi(DInDate.Text, DBaseDate.Text, wDepoID, wEMPID)
        '年休
        DA_Lastyear.Text = GetLastYear(wDepoID, wEMPID, "A")
        DA_Base.Text = GetBase(wDepoID, wEMPID, "A")
        'DA_Base.Text = GetBase(CDbl(DNensi.Text))
        DA_Sum.Text = CDbl(DA_Lastyear.Text) + CDbl(DA_Base.Text)
        DA_Finish.Text = GetFinish(NowYY, wDepoID, wEMPID, "A")
        DA_Vacation.Text = CDbl(DA_Sum.Text) - CDbl(DA_Finish.Text)
        '事假
        DI_Lastyear.Text = "0"
        DI_Base.Text = "14"
        DI_Sum.Text = CDbl(DI_Lastyear.Text) + CDbl(DI_Base.Text)
        DI_Finish.Text = GetFinish(NowYY, wDepoID, wEMPID, "I")
        DI_Vacation.Text = CDbl(DI_Sum.Text) - CDbl(DI_Finish.Text)
        '家顧
        DY_Lastyear.Text = "0"
        DY_Base.Text = "7"
        DY_Sum.Text = CDbl(DY_Lastyear.Text) + CDbl(DY_Base.Text)
        DY_Finish.Text = GetFinish(NowYY, wDepoID, wEMPID, "Y")
        DY_Vacation.Text = CDbl(DY_Sum.Text) - CDbl(DY_Finish.Text)
        '病假
        DS_Lastyear.Text = "0"
        DS_Base.Text = "30"
        DS_Sum.Text = CDbl(DS_Lastyear.Text) + CDbl(DS_Base.Text)
        DS_Finish.Text = GetFinish(NowYY, wDepoID, wEMPID, "S")
        DS_Vacation.Text = CDbl(DS_Sum.Text) - CDbl(DS_Finish.Text)
        '生理
        DZ_Lastyear.Text = "0"
        DZ_Base.Text = "12"
        DZ_Sum.Text = CDbl(DZ_Lastyear.Text) + CDbl(DZ_Base.Text)
        DZ_Finish.Text = GetFinish(NowYY, wDepoID, wEMPID, "Z")
        DZ_Vacation.Text = CDbl(DZ_Sum.Text) - CDbl(DZ_Finish.Text)
        '公假
        DB_Lastyear.Text = "0"
        DB_Base.Text = "0"
        DB_Sum.Text = CDbl(DB_Lastyear.Text) + CDbl(DB_Base.Text)
        DB_Finish.Text = GetFinish(NowYY, wDepoID, wEMPID, "B")
        DB_Vacation.Text = CDbl(DB_Sum.Text) - CDbl(DB_Finish.Text)
        '公傷
        DG_Lastyear.Text = "0"
        DG_Base.Text = "0"
        DG_Sum.Text = CDbl(DG_Lastyear.Text) + CDbl(DG_Base.Text)
        DG_Finish.Text = GetFinish(NowYY, wDepoID, wEMPID, "G")
        DG_Vacation.Text = CDbl(DG_Sum.Text) - CDbl(DG_Finish.Text)
        '婚假
        DM_Lastyear.Text = "0"
        DM_Base.Text = "8"
        DM_Sum.Text = CDbl(DM_Lastyear.Text) + CDbl(DM_Base.Text)
        DM_Finish.Text = GetFinish(NowYY, wDepoID, wEMPID, "M")
        DM_Vacation.Text = CDbl(DM_Sum.Text) - CDbl(DM_Finish.Text)
        '喪假
        DX_Lastyear.Text = "0"
        DX_Base.Text = "8"
        DX_Sum.Text = CDbl(DX_Lastyear.Text) + CDbl(DX_Base.Text)
        DX_Finish.Text = GetFinish(NowYY, wDepoID, wEMPID, "X")
        DX_Vacation.Text = CDbl(DX_Sum.Text) - CDbl(DX_Finish.Text)
        '產假
        DP_Lastyear.Text = "0"
        DP_Base.Text = "56"
        DP_Sum.Text = CDbl(DP_Lastyear.Text) + CDbl(DP_Base.Text)
        DP_Finish.Text = GetFinish(NowYY, wDepoID, wEMPID, "P")
        DP_Vacation.Text = CDbl(DP_Sum.Text) - CDbl(DP_Finish.Text)
        '陪產
        DQ_Lastyear.Text = "0"
        DQ_Base.Text = "3"
        DQ_Sum.Text = CDbl(DQ_Lastyear.Text) + CDbl(DQ_Base.Text)
        DQ_Finish.Text = GetFinish(NowYY, wDepoID, wEMPID, "Q")
        DQ_Vacation.Text = CDbl(DQ_Sum.Text) - CDbl(DQ_Finish.Text)
        '返台假
        DH_Lastyear.Text = "0"
        DH_Base.Text = "0"
        DH_Sum.Text = CDbl(DH_Lastyear.Text) + CDbl(DH_Base.Text)
        DH_Finish.Text = GetFinish(NowYY, wDepoID, wEMPID, "H")
        DH_Vacation.Text = CDbl(DH_Sum.Text) - CDbl(DH_Finish.Text)
        '其他
        DE_Lastyear.Text = "0"
        DE_Base.Text = "0"
        DE_Sum.Text = CDbl(DE_Lastyear.Text) + CDbl(DE_Base.Text)
        DE_Finish.Text = GetFinish(NowYY, wDepoID, wEMPID, "E")
        DE_Vacation.Text = CDbl(DE_Sum.Text) - CDbl(DE_Finish.Text)
        '合計
        DSum_Lastyear.Text = CDbl(DA_Lastyear.Text) + CDbl(DB_Lastyear.Text) + CDbl(DE_Lastyear.Text) + CDbl(DG_Lastyear.Text) + _
                             CDbl(DH_Lastyear.Text) + CDbl(DI_Lastyear.Text) + CDbl(DM_Lastyear.Text) + CDbl(DP_Lastyear.Text) + _
                             CDbl(DQ_Lastyear.Text) + CDbl(DS_Lastyear.Text) + CDbl(DX_Lastyear.Text) + CDbl(DY_Lastyear.Text) + _
                             CDbl(DZ_Lastyear.Text)
        DSum_Base.Text = CDbl(DA_Base.Text) + CDbl(DB_Base.Text) + CDbl(DE_Base.Text) + CDbl(DG_Base.Text) + _
                             CDbl(DH_Base.Text) + CDbl(DI_Base.Text) + CDbl(DM_Base.Text) + CDbl(DP_Base.Text) + _
                             CDbl(DQ_Base.Text) + CDbl(DS_Base.Text) + CDbl(DX_Base.Text) + CDbl(DY_Base.Text) + _
                             CDbl(DZ_Base.Text)
        DSum_Sum.Text = CDbl(DA_Sum.Text) + CDbl(DB_Sum.Text) + CDbl(DE_Sum.Text) + CDbl(DG_Sum.Text) + _
                             CDbl(DH_Sum.Text) + CDbl(DI_Sum.Text) + CDbl(DM_Sum.Text) + CDbl(DP_Sum.Text) + _
                             CDbl(DQ_Sum.Text) + CDbl(DS_Sum.Text) + CDbl(DX_Sum.Text) + CDbl(DY_Sum.Text) + _
                             CDbl(DZ_Sum.Text)
        DSum_Finish.Text = CDbl(DA_Finish.Text) + CDbl(DB_Finish.Text) + CDbl(DE_Finish.Text) + CDbl(DG_Finish.Text) + _
                             CDbl(DH_Finish.Text) + CDbl(DI_Finish.Text) + CDbl(DM_Finish.Text) + CDbl(DP_Finish.Text) + _
                             CDbl(DQ_Finish.Text) + CDbl(DS_Finish.Text) + CDbl(DX_Finish.Text) + CDbl(DY_Finish.Text) + _
                             CDbl(DZ_Finish.Text)
        DSum_Vacation.Text = CDbl(DA_Vacation.Text) + CDbl(DB_Vacation.Text) + CDbl(DE_Vacation.Text) + CDbl(DG_Vacation.Text) + _
                             CDbl(DH_Vacation.Text) + CDbl(DI_Vacation.Text) + CDbl(DM_Vacation.Text) + CDbl(DP_Vacation.Text) + _
                             CDbl(DQ_Vacation.Text) + CDbl(DS_Vacation.Text) + CDbl(DX_Vacation.Text) + CDbl(DY_Vacation.Text) + _
                             CDbl(DZ_Vacation.Text)
        '調閱請假
        LVacationList.NavigateUrl = "HR_VacationInforList_01.aspx?pBaseDate=" + DBaseDate.Text + "&pDepoID=" + wDepoID + "&pEmpID=" + wEMPID + "&pInDate=" + DInDate.Text + "&pOldInDate=" + OldInDate
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     基準部門點選
    '**
    '*****************************************************************
    Protected Sub DDivision_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDivision.SelectedIndexChanged
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()

        DName.Items.Clear()
        SQL = "Select isnull(name,'') as UserName, "
        SQL = SQL + "EmpID, JobName, JobID "
        SQL = SQL + "From V_WorkTime_01 "
        SQL = SQL + "Where HRWDivName = '" & DDivision.SelectedValue & "' "
        SQL = SQL + "  And Not HRWDivName is Null "
        SQL = SQL + "  And Not Name is Null "
        SQL = SQL + "Group by Name, EmpID, JobName, JobID "
        SQL = SQL + "Order by Name, EmpID, JobName, JobID "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            '姓名
            Dim ListItem1 As New ListItem
            ListItem1.Text = Trim(DBTable1.Rows(i).Item("UserName"))
            ListItem1.Value = Trim(DBTable1.Rows(i).Item("UserName"))
            If i = 0 Then
                ListItem1.Selected = True
                DEmpID.Text = Trim(DBDataSet1.Tables("M_Users").Rows(0).Item("EmpID"))
                DJobTitle.Text = Trim(DBDataSet1.Tables("M_Users").Rows(0).Item("JobName"))
                DJobCode.Text = Trim(DBDataSet1.Tables("M_Users").Rows(0).Item("JobID"))
            End If
            DName.Items.Add(ListItem1)
        Next

        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     基準姓名點選
    '**
    '*****************************************************************
    Protected Sub DName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DName.SelectedIndexChanged
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()

        SQL = "Select isnull(name,'') as UserName, "
        SQL = SQL + "EmpID, JobName, JobID "
        SQL = SQL + "From V_WorkTime_01 "
        SQL = SQL + "Where HRWDivName = '" & DDivision.SelectedValue & "' "
        SQL = SQL + "  And Name = '" & DName.SelectedValue & "' "
        SQL = SQL + "  And Not HRWDivName is Null "
        SQL = SQL + "  And Not Name is Null "
        SQL = SQL + "Group by Name, EmpID, JobName, JobID "
        SQL = SQL + "Order by Name, EmpID, JobName, JobID "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        If DBTable1.Rows.Count > 0 Then
            DEmpID.Text = Trim(DBTable1.Rows(0).Item("EmpID"))
            DJobTitle.Text = Trim(DBTable1.Rows(0).Item("JobName"))
            DJobCode.Text = Trim(DBTable1.Rows(0).Item("JobID"))
        End If

        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     計算年資 
    '**
    '*****************************************************************
    Function CalNensi(ByVal pInDate As String, ByVal pBaseDate As String, ByVal pDepo As String, ByVal pID As String) As String
        '計算年資
        Dim wBaseDD As Integer = CDate(pBaseDate).Day
        Dim wInDD As Integer = CDate(pInDate).Day
        Dim Months As Integer = DateDiff(DateInterval.Month, CDate(pInDate), CDate(pBaseDate))

        If wBaseDD < wInDD Then Months = Months - 1
        'MsgBox(CStr(DateDiff(DateInterval.Month, CDate(pInDate), CDate(pBaseDate))) + "-" + CStr(wMinusMM) + "=" + CStr(Months))
        CalNensi = CStr(Fix(Months / 12))
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     去年剩餘年休 
    '**
    '*****************************************************************
    Function GetLastYear(ByVal pDepo As String, ByVal pID As String, ByVal pCode As String) As String
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()

        SQL = "Select isnull(sum(codevalue), 0)  as days From HR_LastYearVacation "
        SQL = SQL & " Where DepoCode = '" & pDepo & "'"
        SQL = SQL & "   And EmpID    = '" & pID & "'"
        SQL = SQL & "   And Code     = '" & pCode & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "LastYear")
        GetLastYear = CStr(DBDataSet1.Tables("LastYear").Rows(0).Item("Days"))

        OleDbConnection1.Close()
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     基準日年休 
    '**
    '*****************************************************************
    'Function GetBase(ByVal pNensi As Integer) As String
    Function GetBase(ByVal pDepo As String, ByVal pID As String, ByVal pCode As String) As String
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()

        SQL = "Select isnull(sum(closeydays), 0)  as days From HR_LastYearVacation "
        SQL = SQL & " Where DepoCode = '" & pDepo & "'"
        SQL = SQL & "   And EmpID    = '" & pID & "'"
        SQL = SQL & "   And Code     = '" & pCode & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "LastYear")
        GetBase = CStr(DBDataSet1.Tables("LastYear").Rows(0).Item("Days"))

        OleDbConnection1.Close()

        'Dim wYears As Integer = 0
        'Select Case pNensi
        '    Case 0
        '        wYears = 0
        '    Case 0.5
        '        wYears = 3
        '    Case 1
        '        wYears = 7
        '    Case 2
        '        wYears = 10
        '    Case 3, 4
        '        wYears = 14
        '    Case 5 To 9
        '        wYears = 15
        '    Case Else
        '        wYears = pNensi + 6
        'End Select
        'If wYears > 30 Then wYears = 30
        'GetBase = CStr(wYears)
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     已休休假 
    '**
    '*****************************************************************
    Function GetFinish(ByVal pYY As String, ByVal pDepo As String, ByVal pID As String, ByVal pCode As String) As String
        Dim SQL As String
        Dim wDays As Double = 0
        Dim wSYYMM As String = ""
        Dim wEYYMM As String = ""
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()

        '年休假特別處理
        If pCode = "A" Then    '年休   
            '計算起迄日期
            If CDate(DBaseDate.Text) >= CDate(pYY + Mid(DInDate.Text, 5, 6)) Then
                wSYYMM = pYY + Mid(DInDate.Text, 5, 6)
                wEYYMM = CStr(CInt(pYY) + 1) + Mid(DInDate.Text, 5, 6)
            Else
                wSYYMM = CStr(CInt(pYY) - 1) + Mid(DInDate.Text, 5, 6)
                wEYYMM = pYY + Mid(DInDate.Text, 5, 6)
            End If
            '
            '上線前
            SQL = "Select isnull(sum(CodeValue), 0)  as days From HR_BeforeVacationList "
            SQL = SQL & " Where VacationDate >= '" & wSYYMM & "'"
            SQL = SQL & "   And VacationDate <  '" & wEYYMM & "'"
            SQL = SQL & "   And DepoCode = '" & pDepo & "'"
            SQL = SQL & "   And EmpID    = '" & pID & "'"
            SQL = SQL & "   And Code = '" & pCode & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "FinishBefore")
            wDays = DBDataSet1.Tables("FinishBefore").Rows(0).Item("Days")
            '上線後
            DBDataSet1.Clear()
            SQL = "Select isnull(sum(ADays), 0)  as days From F_TimeOffSheet "
            SQL = SQL & " Where AStartDate >= '" & wSYYMM & "'"
            SQL = SQL & "   And AEndDate   <  '" & wEYYMM & "'"
            SQL = SQL & "   And DepoCode = '" & pDepo & "'"
            SQL = SQL & "   And EmpID    = '" & pID & "'"
            SQL = SQL & "   And VacationCode = '" & pCode & "'"
            SQL = SQL & "   And Sts      = '1' "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "Finish")
            GetFinish = CStr(wDays + DBDataSet1.Tables("Finish").Rows(0).Item("Days"))
        Else
            '計算起迄日期
            '非年休
            If CDate(DBaseDate.Text) >= CDate(pYY + "/4/1") Then
                wSYYMM = pYY + "/04"
                wEYYMM = CStr(CInt(pYY) + 1) + "/03"
            Else
                wSYYMM = CStr(CInt(pYY) - 1) + "/04"
                wEYYMM = pYY + "/03"
            End If
            '
            SQL = "Select isnull(sum(ADays), 0)  as days From F_TimeOffSheet "
            SQL = SQL & " Where SalaryYM >= '" & wSYYMM & "'"
            SQL = SQL & "   And SalaryYM <= '" & wEYYMM & "'"
            SQL = SQL & "   And DepoCode = '" & pDepo & "'"
            SQL = SQL & "   And EmpID    = '" & pID & "'"
            SQL = SQL & "   And VacationCode = '" & pCode & "'"
            SQL = SQL & "   And Sts      = '1' "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "Finish")
            GetFinish = CStr(DBDataSet1.Tables("Finish").Rows(0).Item("Days"))
        End If

        OleDbConnection1.Close()
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     計算按鈕 
    '**
    '*****************************************************************
    Protected Sub BSimulate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSimulate.Click
        If IsDate(DBaseDate.Text) Then
            ShowVacationInfor()
        Else
            Response.Write(YKK.ShowMessage("基準日輸入錯誤, 請確認!"))
        End If
    End Sub
End Class
