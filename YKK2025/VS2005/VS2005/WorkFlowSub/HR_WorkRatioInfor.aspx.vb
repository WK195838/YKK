Imports System.Data
Imports System.Data.OleDb

Partial Class HR_WorkRatioInfor
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
    Dim startdate As String
    Dim enddate As String

 


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        SetParameter()

        If Not Me.IsPostBack Then
            ShowPersonInfor()    '顯示個人基本資訊
            'ShowVacationInfor()  '顯示休假資訊

        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     顯示個人基本資訊
    '**
    '*****************************************************************
    Sub ShowPersonInfor()
        ClearData()
        Dim SQL As String
        Dim wLevel, wName, wDivision As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()

        '--------------------------------------------------------------------------------------------------------
        ' 年度
        '--------------------------------------------------------------------------------------------------------

        SQL = " SELECT * FROM M_REFERP "
        SQL &= " WHERE CAT = '016' "
        SQL &= " AND DKEY = 'WORKRATIO-YY'"

        Dim DBAdapter10 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter10.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")

        If DBTable1.Rows.Count > 0 Then
            DYear.Text = DBTable1.Rows(0).Item("Data")
        End If
        DBDataSet1.Clear()


        '當日
        DDate.Text = CStr(DateTime.Now.Date)

        '開始日期
        startdate = DYear.Text + "/01/01"
        enddate = DYear.Text + "/12/31"

        '基準日
        DBaseDate.Text = enddate


        '+++++++++++++++++++++++++++++++++++++++++++++++++++++
        '++  設定部門/姓名權限
        '++  顯示個人基本資訊
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++
        '個人基本資訊
        DBDataSet1.Clear()
        SQL = "Select distinct isnull(Name,'') as UserName, isnull(HRWDivName,'') as HRWDivName, "
        SQL = SQL + "EmpID, JobName, JobID, DepoID "
        SQL = SQL + "From V_WorkTime_01 "
        SQL = SQL + "Where UserID='" & Request.QueryString("pUserID") & "' "
        SQL = SQL + "  And HRWDivName <>'' "
        SQL = SQL + "  And Name <>'' "
        SQL = SQL + " and convert(char(10),cdate,111) between '" + startdate + "' and '" + enddate + "' "
        'SQL = SQL + "Order by Name "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        If DBTable1.Rows.Count > 0 Then
            wDepoID = Trim(DBTable1.Rows(0).Item("DepoID"))
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
            SQL = "Select  distinct isnull(Name,'') as UserName, isnull(HRWDivName,'') as HRWDivName "
            SQL = SQL + "From V_WorkTime_01 "
            SQL = SQL + "Where UserID='" & Request.QueryString("pUserID") & "' "
            SQL = SQL + "  And HRWDivName <> '' "
            SQL = SQL + "  And Name <>'' "
            SQL = SQL + " and convert(char(10),cdate,111) between '" + startdate + "' and '" + enddate + "' "
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
                SQL = SQL + "  And DepoID in ('10', '11') "
                SQL = SQL + "  And HRWDivName <>'' "
                SQL = SQL + "  And Name <>'' "
                SQL = SQL + "  And LEAV_CD ='' "
                SQL = SQL + "Group by Name, EmpID, JobName, JobID "
                'SQL = SQL + "Order by Name, EmpID, JobName, JobID "
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
                    SQL = SQL + "Where DepoID in ('10', '11') "
                    SQL = SQL + " And HRWDivName <> '' "
                    SQL = SQL + " And Name <>'' "
                    SQL = SQL + " And convert(char(10),cdate,111) between '" + startdate + "' and '" + enddate + "' "
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
                    SQL = SQL + " And DepoID in ('10', '11') "
                    SQL = SQL + " And HRWDivName  <>'' "
                    SQL = SQL + " And Name  <>'' "
                    SQL = SQL + " And LEAV_CD ='' "
                    SQL = SQL + " and convert(char(10),cdate,111) between '" + startdate + "' and '" + enddate + "' "
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
        Count()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     基準部門點選
    '**
    '*****************************************************************
    Protected Sub DDivision_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDivision.SelectedIndexChanged
        ClearData()
        '開始日期
        startdate = DYear.Text + "/01/01"
        enddate = DBaseDate.Text

        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()

        DName.Items.Clear()
        SQL = "Select isnull(name,'') as UserName, "
        SQL = SQL + "DepoID, EmpID, JobName, JobID "
        SQL = SQL + "From V_WorkTime_01 "
        SQL = SQL + "Where HRWDivName = '" & DDivision.SelectedValue & "' "
        SQL = SQL + "  And DepoID in ('10', '11') "
        SQL = SQL + "  And HRWDivName <>'' "
        SQL = SQL + "  And Name <>'' "
        SQL = SQL + "  And LEAV_CD ='' "
        SQL = SQL + " and convert(char(10),cdate,111) between '" + startdate + "' and '" + enddate + "' "
        SQL = SQL + "Group by Name, DepoID, EmpID, JobName, JobID "
        SQL = SQL + "Order by Name, DepoID, EmpID, JobName, JobID "
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
                wDepoID = Trim(DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID"))
                DEmpID.Text = Trim(DBDataSet1.Tables("M_Users").Rows(0).Item("EmpID"))
                DJobTitle.Text = Trim(DBDataSet1.Tables("M_Users").Rows(0).Item("JobName"))
                DJobCode.Text = Trim(DBDataSet1.Tables("M_Users").Rows(0).Item("JobID"))
            End If
            DName.Items.Add(ListItem1)
        Next

        OleDbConnection1.Close()
        Count()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     基準姓名點選
    '**
    '*****************************************************************
    Protected Sub DName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DName.SelectedIndexChanged
        ClearData()
        '開始日期
        startdate = DYear.Text + "/01/01"
        enddate = DBaseDate.Text

        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()

        SQL = "Select isnull(name,'') as UserName, "
        SQL = SQL + "DepoID, EmpID, JobName, JobID "
        SQL = SQL + "From V_WorkTime_01 "
        SQL = SQL + "Where HRWDivName = '" & DDivision.SelectedValue & "' "
        SQL = SQL + "  And Name = '" & DName.SelectedValue & "' "
        SQL = SQL + "  And HRWDivName <> '' "
        SQL = SQL + "  And Name <> '' "
        SQL = SQL + " and convert(char(10),cdate,111) between '" + startdate + "' and '" + enddate + "' "
        SQL = SQL + "Group by Name, DepoID, EmpID, JobName, JobID "
        ' SQL = SQL + "Order by Name, EmpID, JobName, JobID "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        If DBTable1.Rows.Count > 0 Then
            wDepoID = Trim(DBTable1.Rows(0).Item("DepoID"))
            DEmpID.Text = Trim(DBTable1.Rows(0).Item("EmpID"))
            DJobTitle.Text = Trim(DBTable1.Rows(0).Item("JobName"))
            DJobCode.Text = Trim(DBTable1.Rows(0).Item("JobID"))
        End If

        OleDbConnection1.Close()
        Count()
    End Sub
    '-------------------------------------------------------------------------------------------------------
    '  模擬計算
    '-------------------------------------------------------------------------------------------------------
    Sub Count()

        ClearData() '清空欄位資料  
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim depo As String = ""
        Dim i As Integer
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()
        '開始日期
        startdate = DYear.Text + "/01/01"
        enddate = DBaseDate.Text

        '--------------------------------------------------------------------------------------------------------
        ' 應出勤
        '--------------------------------------------------------------------------------------------------------

        '應際出勤日
        SQL = " select depo,yearday,count(*)cun  from ( "
        SQL &= " Select  depo,datediff(day,'" + startdate + "','" + enddate + "')+1 as YearDay "
        SQL &= " from m_emp a , M_Vacation b"
        SQL &= " where  NAME = '" + DName.SelectedValue + "' and id = '" + DEmpID.Text + "'"
        SQL &= " and  yy ='" + DYear.Text + "' and  convert(char(10),ymd,111) between '" + startdate + "' and '" + enddate + "' "
        SQL &= " and a.calendargroup =b.depo"
        SQL &= " )a group by depo,yearday"

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")
        DVacationday.Text = "0"
        If DBTable1.Rows.Count > 0 Then
            DYearDay.Text = Trim(DBTable1.Rows(0).Item("yearday"))
            DVacationday.Text = Trim(DBTable1.Rows(0).Item("cun"))
            depo = Trim(DBTable1.Rows(0).Item("depo"))
        End If
        DBDataSet1.Clear()
        '應出勤1日分數
        SQL = "select * from M_referp where  cat ='016' and dkey = '" + depo + "'+'-Y1'"
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")
        Dim Y1 As Integer
        If DBTable1.Rows.Count > 0 Then
            Y1 = Trim(DBTable1.Rows(0).Item("Data"))
        Else
            Y1 = 500
        End If
        DBworkday.Text = CStr(Y1)
        DBDataSet1.Clear()

        '應出勤調整分數
        SQL = "select * from M_referp where  cat ='016' and dkey ='" + depo + "'+'-Y2'"
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")

        If DBTable1.Rows.Count > 0 Then
            DY2.Text = Trim(DBTable1.Rows(0).Item("Data"))
        Else
            DY2.Text = 0
        End If
        Dim Y2 As Integer
        Y2 = Int(DY2.Text)
        DBDataSet1.Clear()

        '應出勤個人調整分數
        SQL = " select isnull(sum(AdjustTime),0) as AdjustTime from M_PersonAdjustTime where active ='1'"
        SQL &= " and depo = '" + depo + "' and empid ='" + DEmpID.Text + "'"
        SQL &= " and yy ='" + DYear.Text + "' and AdjustType='Y3'"

        Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter4.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")

        If DBTable1.Rows.Count > 0 Then
            DY3.Text = Trim(DBTable1.Rows(0).Item("AdjustTime"))
        Else
            DY3.Text = "0"
        End If
        Dim Y3 As Integer
        Y3 = Int(DY3.Text)
        DBDataSet1.Clear()
        'DBworkday.Text = Str((CDbl(DYearDay.Text) - CDbl(DVacationday.Text)) * Y1 + Y2 + Y3)
        Dim Y As Integer
        Y = (CDbl(DYearDay.Text) - CDbl(DVacationday.Text)) * Y1 + Y2 + Y3
        '--------------------------------------------------------------------------------------------------------
        ' 實際出勤
        '--------------------------------------------------------------------------------------------------------

        '實際出勤1日分數
        SQL = "select * from M_referp where  cat ='016' and dkey = '" + depo + "'+'-X1'"
        Dim DBAdapter8 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter8.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")
        Dim X1 As Integer
        If DBTable1.Rows.Count > 0 Then
            X1 = Trim(DBTable1.Rows(0).Item("Data"))
        Else
            X1 = 480
        End If
        DAWorkday.Text = CStr(X1)
        DBDataSet1.Clear()


        DBDataSet1.Clear()
        '一般請假(B1-1)
        SQL = " select  empid,name, vacationcode,sum(adays)adays  from  F_TimeoffSheet "
        SQL &= " where  empid = '" + DEmpID.Text + "' and name ='" + DName.SelectedValue + "' and convert(char(10),AEndDate,111) between '" + startdate + "' and '" + enddate + "' "
        SQL &= " and sts =1  and vacationcode in ('I', 'Y', 'S', 'Z', 'M', 'X')"
        SQL &= " group by  empid,name, vacationcode"


        Dim DBAdapter5 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter5.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")

        For i = 0 To DBTable1.Rows.Count - 1
            '假別          
            If Trim(DBTable1.Rows(i).Item("vacationcode")) = "I" Then '事假 

                DIVcation.Text = Trim(DBTable1.Rows(i).Item("adays"))
                DIVcationM.Text = X1 * -CDbl(DIVcation.Text)

            ElseIf Trim(DBTable1.Rows(i).Item("vacationcode")) = "Y" Then '家庭照顧假

                DYVcation.Text = Trim(DBTable1.Rows(i).Item("adays"))
                DYVcationM.Text = X1 * -CDbl(DYVcation.Text)


            ElseIf Trim(DBTable1.Rows(i).Item("vacationcode")) = "S" Then '病假

                DSVcation.Text = Trim(DBTable1.Rows(i).Item("adays"))
                DSVcationM.Text = X1 * -CDbl(DSVcation.Text)

            ElseIf Trim(DBTable1.Rows(i).Item("vacationcode")) = "Z" Then '生理假

                DZVcation.Text = Trim(DBTable1.Rows(i).Item("adays"))
                DZVcationM.Text = X1 * -CDbl(DZVcation.Text)

            ElseIf Trim(DBTable1.Rows(i).Item("vacationcode")) = "M" Then '婚假

                DMVcation.Text = Trim(DBTable1.Rows(i).Item("adays"))
                DMVcationM.Text = X1 * -CDbl(DMVcation.Text)

            ElseIf Trim(DBTable1.Rows(i).Item("vacationcode")) = "X" Then '喪假

                DXVcation.Text = Trim(DBTable1.Rows(i).Item("adays"))
                DXVcationM.Text = X1 * -CDbl(DXVcation.Text)
            End If
        Next
        DBDataSet1.Clear()
        '產假陪產假(B1-2) 
        SQL = " select vacationcode,adays-cun as adays  from ("
        SQL &= " select  empid,name,vacationcode,adays,count(*)cun   from ("
        SQL &= "select  empid,name ,vacationcode,sum(adays)adays,min(astartdate)astartdate,max(aenddate)aenddate  from  F_TimeoffSheet  "
        SQL &= " where  empid = '" + DEmpID.Text + "' and name ='" + DName.SelectedValue + "'"
        SQL &= " and convert(char(10),AEndDate,111) between '" + startdate + "' and '" + enddate + "' "
        SQL &= " and vacationcode in ('p','q') "
        SQL &= " and sts  ='1'"
        SQL &= " group by  empid,name ,vacationcode"
        SQL &= " )a ,  M_Vacation b"
        SQL &= " where  depo ='" + depo + "'"
        SQL &= " and yy ='" + DYear.Text + "'"
        SQL &= " and active = '1' "
        SQL &= " and ymd between astartdate and aenddate "
        SQL &= " group by  empid,name,vacationcode,adays"
        SQL &= " )a "




        Dim DBAdapter6 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter6.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")
        DPVcation.Text = "0"
        DQVcation.Text = "0"

        For i = 0 To DBTable1.Rows.Count - 1
            '假別          
            If Trim(DBTable1.Rows(i).Item("vacationcode")) = "P" Then '產假 
                DPVcation.Text = Trim(DBTable1.Rows(i).Item("adays"))
                DPVcationM.Text = X1 * -CDbl(DPVcation.Text)
            ElseIf Trim(DBTable1.Rows(i).Item("vacationcode")) = "Q" Then '陪產假 
                DQVcation.Text = Trim(DBTable1.Rows(i).Item("adays"))
                DQVcationM.Text = X1 * -CDbl(DQVcation.Text)
            End If

        Next
        DBDataSet1.Clear()

        '曠職
        SQL = " select isnull(sum(AdjustTime),0) as AdjustTime from M_PersonAdjustTime where active ='1'"
        SQL &= " and depo = '" + depo + "' and empid ='" + DEmpID.Text + "'"
        SQL &= " and yy ='" + DYear.Text + "' and AdjustType='Z1'"

        Dim DBAdapter7 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter7.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")

        If DBTable1.Rows.Count > 0 Then
            DZ1M.Text = Trim(DBTable1.Rows(0).Item("AdjustTime"))

        Else
            DZ1M.Text = "0"

        End If

        DBDataSet1.Clear()



        '實際出勤個人調整分數
        SQL = " select isnull(sum(AdjustTime),0) as AdjustTime from M_PersonAdjustTime where active ='1'"
        SQL &= " and depo = '" + depo + "' and empid ='" + DEmpID.Text + "'"
        SQL &= " and yy ='" + DYear.Text + "' and AdjustType='X2'"

        Dim DBAdapter9 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter9.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")

        If DBTable1.Rows.Count > 0 Then
            DX2.Text = Trim(DBTable1.Rows(0).Item("AdjustTime"))

        Else
            DX2.Text = "0"
        End If

        DBDataSet1.Clear()
        Dim sumVacation As Double
        sumVacation = CInt(DIVcationM.Text) + CInt(DYVcationM.Text) + CInt(DSVcationM.Text) + _
        CInt(DZVcationM.Text) + CInt(DMVcationM.Text) + _
        CInt(DXVcationM.Text) + CInt(DPVcationM.Text) + CInt(DQVcationM.Text) + CInt(DZ1M.Text) + CInt(DX2.Text)

        '實際出勤日
        'DAWorkday.Text = CDbl(DBworkday.Text) - sumVacation
        Dim X As Integer
        X = Y + sumVacation
        DSum.Text = "          (" + CStr(Y) + "+ (" + CStr(sumVacation) + ")    )            /" + CStr(Y)
        '--------------------------------------------------------------------------------------------------------
        '出勤率
        '--------------------------------------------------------------------------------------------------------
        If Y = 0 Then
            DWorkRate.Text = 0
        Else
            DWorkRate.Text = FormatNumber((Y + (sumVacation)) / Y, 4)

        End If


        DYearDay.Style.Add("TEXT-ALIGN", "right")
        DVacationday.Style.Add("TEXT-ALIGN", "right")
        DY2.Style.Add("TEXT-ALIGN", "right")
        DY3.Style.Add("TEXT-ALIGN", "right")
        DIVcation.Style.Add("TEXT-ALIGN", "right")
        DYVcation.Style.Add("TEXT-ALIGN", "right")
        DSVcation.Style.Add("TEXT-ALIGN", "right")
        DZVcation.Style.Add("TEXT-ALIGN", "right")
        DMVcation.Style.Add("TEXT-ALIGN", "right")
        DXVcation.Style.Add("TEXT-ALIGN", "right")
        DPVcation.Style.Add("TEXT-ALIGN", "right")
        DQVcation.Style.Add("TEXT-ALIGN", "right")
        DIVcationM.Style.Add("TEXT-ALIGN", "right")
        DYVcationM.Style.Add("TEXT-ALIGN", "right")
        DSVcationM.Style.Add("TEXT-ALIGN", "right")
        DZVcationM.Style.Add("TEXT-ALIGN", "right")
        DMVcationM.Style.Add("TEXT-ALIGN", "right")
        DXVcationM.Style.Add("TEXT-ALIGN", "right")
        DPVcationM.Style.Add("TEXT-ALIGN", "right")
        DQVcationM.Style.Add("TEXT-ALIGN", "right")

        DZ1M.Style.Add("TEXT-ALIGN", "right")
        DX2.Style.Add("TEXT-ALIGN", "right")
        DAWorkday.Style.Add("TEXT-ALIGN", "right")
        DBworkday.Style.Add("TEXT-ALIGN", "right")

        '其它說明
        LOtherDescription.NavigateUrl = "HR_WorkRatioOtherInfor.aspx?pYear=" + DYear.Text + "&pDepo=" + depo + "&pEmpid=" + DEmpID.Text
        LOtherDescription.Visible = True
        '休假調閱
        LVacationList.NavigateUrl = "HR_VacationInforList_02.aspx?pYear=" + DYear.Text + "&pDepo=" + wDepoID + "&pEmpid=" + DEmpID.Text
        LVacationList.Visible = True
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     資料清空
    '**
    '*****************************************************************
    Sub ClearData()
        LVacationList.Visible = False
        DAWorkday.Text = "0"
        DBworkday.Text = "0"
        DYearDay.Text = "0"
        DVacationday.Text = "0"
        DY2.Text = "0"
        DY3.Text = "0"
        DIVcation.Text = "0"
        DYVcation.Text = "0"
        DSVcation.Text = "0"
        DZVcation.Text = "0"
        DMVcation.Text = "0"
        DXVcation.Text = "0"
        DPVcation.Text = "0"
        DQVcation.Text = "0"
        DIVcationM.Text = "0"
        DYVcationM.Text = "0"
        DSVcationM.Text = "0"
        DZVcationM.Text = "0"
        DMVcationM.Text = "0"
        DXVcationM.Text = "0"
        DPVcationM.Text = "0"
        DQVcationM.Text = "0"
        DWorkRate.Text = "0"

        DZ1M.Text = "0"
        DX2.Text = "0"
    End Sub

  
End Class
