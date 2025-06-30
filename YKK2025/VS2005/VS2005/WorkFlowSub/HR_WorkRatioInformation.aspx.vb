Imports System.Data
Imports System.Data.OleDb

Partial Class HR_WorkRatioInformation
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
    Dim wDepoID As String           'Depo
    Dim wCalendarGroup As String     '行事曆群組
    Dim wEMPID As String            'EMP-ID
    Dim Startdate As String         '計算起始日
    Dim Enddate As String           '計算截止日
    Dim Firstdate As String         '元旦日
    Dim SumDays As Double = 0       '日數合計 
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()
        If Not IsPostBack Then
            SetImageButtonImageFile(0)
            SetImageButtonEvent()
            '
            ShowPersonInfor()       '個人資訊 & 權限設定
            '
            WorkRatio()             '出勤率
            '
            ShowVacationInfor()     '請假履歷
            '
            ShowOtherInfor()        '其他說明
        Else
        End If
    End Sub
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
    End Sub
    '*****************************************************************
    '**
    '**     個人資訊 & 權限設定
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
        '--------------------------------------------------------------------------------------------------------
        ' 當日
        '--------------------------------------------------------------------------------------------------------
        DDate.Text = CStr(DateTime.Now.Date)
        '--------------------------------------------------------------------------------------------------------
        ' 開始日期
        '--------------------------------------------------------------------------------------------------------
        Startdate = DYear.Text + "/04/01"
        Enddate = CStr(CInt(DYear.Text) + 1) + "/03/31"
        Firstdate = CStr(CInt(DYear.Text) + 1) + "/01/01"
        '--------------------------------------------------------------------------------------------------------
        ' 基準日
        '--------------------------------------------------------------------------------------------------------
        DBaseDate.Text = enddate
        '--------------------------------------------------------------------------------------------------------
        ' 個人基本資訊
        '--------------------------------------------------------------------------------------------------------
        DBDataSet1.Clear()
        SQL = "Select distinct isnull(Name,'') as UserName, isnull(HRWDivName,'') as HRWDivName, "
        SQL = SQL + "EmpID, JobName, JobID, DepoID "
        SQL = SQL + "From V_WorkTime_01 "
        SQL = SQL + "Where UserID='" & Request.QueryString("pUserID") & "' "
        SQL = SQL + "  And HRWDivName <>'' "
        SQL = SQL + "  And Name <>'' "
        SQL = SQL + " and convert(char(10),cdate,111) between '" + Startdate + "' and '" + Enddate + "' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        If DBTable1.Rows.Count > 0 Then
            wDepoID = Trim(DBTable1.Rows(0).Item("DepoID"))
            wDivision = Trim(DBTable1.Rows(0).Item("HRWDivName"))
            wName = Trim(DBTable1.Rows(0).Item("UserName"))
            DEmpID.Text = Trim(DBTable1.Rows(0).Item("EmpID"))
            DJobTitle.Text = Trim(DBTable1.Rows(0).Item("JobName"))
            DJobCode.Text = Trim(DBTable1.Rows(0).Item("JobID"))
        End If
        '--------------------------------------------------------------------------------------------------------
        ' 權限
        '--------------------------------------------------------------------------------------------------------
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

        OleDbConnection1.Close()
    End Sub
    '-------------------------------------------------------------------------------------------------------
    '  出勤率
    '-------------------------------------------------------------------------------------------------------
    Sub WorkRatio()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim i As Integer
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()
        '
        ClearWorkRatioData() '清空欄位資料  
        '
        '--------------------------------------------------------------------------------------------------------
        ' 應出勤
        '--------------------------------------------------------------------------------------------------------
        '應際出勤日
        SQL = " select depo,yearday,count(*)cun  from ( "
        SQL &= " Select  depo,datediff(day,'" + startdate + "','" + enddate + "')+1 as YearDay "
        SQL &= " from m_emp a , M_Vacation b"
        SQL &= " where  NAME = '" + DName.SelectedValue + "' and id = '" + DEmpID.Text + "'"
        SQL &= " and  a.Com_Code in ('10', '11') "
        SQL &= " and  convert(char(10),ymd,111) between '" + Startdate + "' and '" + Enddate + "' "
        SQL &= " and a.calendargroup =b.depo"
        SQL &= " )a group by depo,yearday"

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")
        DVacationday.Text = "0"
        If DBTable1.Rows.Count > 0 Then
            DYearDay.Text = Trim(DBTable1.Rows(0).Item("yearday"))
            DVacationday.Text = Trim(DBTable1.Rows(0).Item("cun"))
            wCalendarGroup = Trim(DBTable1.Rows(0).Item("depo"))
        End If
        DBDataSet1.Clear()
        '應出勤1日分數
        SQL = "select * from M_referp where  cat ='016' and dkey = '" + wCalendarGroup + "'+'-Y1'"
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")
        Dim Y1 As Integer
        If DBTable1.Rows.Count > 0 Then
            Y1 = Trim(DBTable1.Rows(0).Item("Data"))
        Else
            Y1 = 0
        End If
        DBworkday.Text = CStr(Y1)
        DBDataSet1.Clear()

        '應出勤調整分數
        SQL = "select * from M_referp where  cat ='016' and dkey ='" + wCalendarGroup + "'+'-Y2'"
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
        SQL &= " and depo = '" + wCalendarGroup + "' and empid ='" + DEmpID.Text + "'"
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
        SQL = "select * from M_referp where  cat ='016' and dkey = '" + wCalendarGroup + "'+'-X1'"
        Dim DBAdapter8 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter8.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")
        Dim X1 As Integer
        If DBTable1.Rows.Count > 0 Then
            X1 = Trim(DBTable1.Rows(0).Item("Data"))
        Else
            X1 = 0
        End If
        DAWorkday.Text = CStr(X1)
        DBDataSet1.Clear()


        DBDataSet1.Clear()
        '一般請假(B1-1)
        SQL = " select  empid,name, vacationcode,sum(adays)adays  from  F_TimeoffSheet "
        SQL &= " where  empid = '" + DEmpID.Text + "' and name ='" + DName.SelectedValue + "' and convert(char(10),AEndDate,111) between '" + Startdate + "' and '" + Enddate + "' "
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

        'SQL = " select vacationcode,adays-cun as adays  from ("
        SQL = " select vacationcode,adays as adays  from ("
        SQL &= " select  empid,name,vacationcode,adays,count(*)-1 As cun from ("
        SQL &= "select  empid,name ,vacationcode,sum(adays)adays,min(astartdate)astartdate,max(aenddate)aenddate  from  F_TimeoffSheet  "
        SQL &= " where  empid = '" + DEmpID.Text + "' and name ='" + DName.SelectedValue + "'"
        SQL &= " and convert(char(10),AEndDate,111) between '" + Startdate + "' and '" + Enddate + "' "
        SQL &= " and vacationcode in ('p','q','w') "
        SQL &= " and sts  ='1'"
        SQL &= " group by  empid,name ,vacationcode"
        SQL &= " )a ,  M_Vacation b"
        SQL &= " where  depo ='" + wCalendarGroup + "'"
        SQL &= " and ymd between '" + Startdate + "' and '" + Enddate + "' "
        SQL &= " and active = '1' "

        SQL &= " and (ymd between astartdate and aenddate or ymd = '" + Firstdate + "') "

        SQL &= " group by  empid,name,vacationcode,adays"
        SQL &= " )a "

        Dim DBAdapter6 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter6.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")
        DPVcation.Text = "0"
        DQVcation.Text = "0"

        For i = 0 To DBTable1.Rows.Count - 1
            '假別          
            If Trim(DBTable1.Rows(i).Item("vacationcode")) = "P" Or Trim(DBTable1.Rows(i).Item("vacationcode")) = "W" Then '產假 
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
        SQL &= " and depo = '" + wCalendarGroup + "' and empid ='" + DEmpID.Text + "'"
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
        SQL &= " and depo = '" + wCalendarGroup + "' and empid ='" + DEmpID.Text + "'"
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
    End Sub
    '*****************************************************************
    '**
    '**     請假履歷
    '**
    '*****************************************************************
    Sub ShowVacationInfor()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        Dim wHttp As String = System.Configuration.ConfigurationManager.AppSettings("Http")

        OleDbConnection1.Open()
        DBDataSet1.Clear()
        SQL = "Select No, VacationCode+':'+Vacation As Vacation, ADays, "
        SQL = SQL & "Convert(VARCHAR(10), AStartDate, 111) + ' ' + str(AStartH) + ':00' + '~' + "
        SQL = SQL & "Convert(VARCHAR(10), AEndDate, 111)   + ' ' + str(AEndH) + ':00' as VacationTime, "
        SQL = SQL + "'" + wHttp + "' + '/WorkFlow/HR_TimeOffSheet_02.aspx?' + "
        SQL = SQL + "'&pFormNo=' + FormNo + "
        SQL = SQL + "'&pFormSno=' + str(FormSno) "
        SQL = SQL + " As URL "
        SQL = SQL & "From F_TimeOffSheet "
        SQL = SQL & " Where DepoCode = '" & wDepoID & "'"
        SQL = SQL & "   And EmpID    = '" & DEmpID.Text & "'"
        SQL = SQL & "   And Sts      = '1' "
        SQL = SQL & "   And ( "
        SQL = SQL & "           VacationCode >= 'I' "
        SQL = SQL & "        OR VacationCode >= 'Y' "
        SQL = SQL & "        OR VacationCode >= 'S' "
        SQL = SQL & "        OR VacationCode >= 'Z' "
        SQL = SQL & "        OR VacationCode >= 'M' "
        SQL = SQL & "        OR VacationCode >= 'X' "
        SQL = SQL & "        OR VacationCode >= 'P' "
        SQL = SQL & "        OR VacationCode >= 'Q' "
        SQL = SQL & "       ) "
        SQL = SQL & "   And AStartDate >= '" & Startdate & "'"
        SQL = SQL & "   And AEndDate   <= '" & Enddate & "'"
        SQL = SQL & " Order by VacationCode, AStartDate "
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "Data")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()

        OleDbConnection1.Close()
    End Sub
    '*****************************************************************
    '**
    '**     請假日數合計
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.Header Then
            SumDays = 0      '日數合計 
        End If
        If e.Row.RowType = DataControlRowType.DataRow Then
            SumDays = SumDays + CDbl(e.Row.Cells(3).Text.ToString)
        End If
        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(2).Text = "合計"
            e.Row.Cells(3).Text = FormatNumber(SumDays, 1)
        End If
    End Sub
    '*****************************************************************
    '**
    '**     其他說明
    '**
    '*****************************************************************
    Sub ShowOtherInfor()
        Dim SQL As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '
        ClearOtherInforData() '清空欄位資料  
        '---------------------------------------------------------------------------------------------------------
        'Y2  應出勤時間
        '---------------------------------------------------------------------------------------------------------
        OleDbConnection1.Open()

        SQL = " select  *  from  m_referp "
        SQL &= " where cat ='016' "
        SQL &= " and dkey like '" + wCalendarGroup + "-Y2%" + "'"

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")
        DY2Sum.Text = "0"
        If DBTable1.Rows.Count > 0 Then
            For i = 0 To DBTable1.Rows.Count - 1
                If DBTable1.Rows(i).Item("DKEY") = wCalendarGroup + "-Y2" Then
                    DY2Min.Text = Trim(DBTable1.Rows(i).Item("Data"))
                    DY2Sum.Text = Trim(DBTable1.Rows(i).Item("Data"))
                Else
                    DY2Desc.Text = Trim(DBTable1.Rows(i).Item("Data"))
                End If
            Next
        End If
        DBDataSet1.Clear()

        '---------------------------------------------------------------------------------------------------------
        'Y3  應出勤時間
        '---------------------------------------------------------------------------------------------------------
        SQL = " select top 5 * from  M_PersonAdjustTime "
        SQL &= " where active = 1 and yy = '" + DYear.Text + "'"
        SQL &= " and adjusttype ='Y3' "
        SQL &= " and empid = '" + DEmpID.Text + "'"
        SQL &= " order by adjusttype,seqno "

        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")
        Dim wDY3SUM As Integer = 0


        If DBTable1.Rows.Count > 0 Then
            For i = 0 To DBTable1.Rows.Count - 1
                If i = 0 Then
                    DY3Min1.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DY3Desc1.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 1 Then
                    DY3Min2.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DY3Desc2.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 2 Then
                    DY3Min3.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DY3Desc3.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 3 Then
                    DY3Min4.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DY3Desc4.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 4 Then
                    DY3Min5.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DY3Desc5.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                End If
                wDY3SUM = wDY3SUM + Int(Trim(DBTable1.Rows(i).Item("AdjustTime")))
            Next
        End If
        DY3Sum.Text = wDY3SUM

        DBDataSet1.Clear()

        '---------------------------------------------------------------------------------------------------------
        'X2  實際出勤時間
        '---------------------------------------------------------------------------------------------------------

        SQL = " select top 5 * from  M_PersonAdjustTime "
        SQL &= " where active = 1 and yy = '" + DYear.Text + "'"
        SQL &= " and adjusttype ='X2'"
        SQL &= " and empid = '" + DEmpID.Text + "'"
        SQL &= " order by adjusttype,seqno "

        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")

        Dim wDX2SUM As Integer = 0

        If DBTable1.Rows.Count > 0 Then

            For i = 0 To DBTable1.Rows.Count - 1
                If i = 0 Then
                    DX2Min1.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DX2Desc1.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 1 Then
                    DX2Min2.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DX2Desc2.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 2 Then
                    DX2Min3.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DX2Desc3.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 3 Then
                    DX2Min4.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DX2Desc4.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 4 Then
                    DX2Min5.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DX2Desc5.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                End If
                wDX2SUM = wDX2SUM + Int(Trim(DBTable1.Rows(i).Item("AdjustTime")))
            Next
        End If
        DX2Sum.Text = wDX2SUM
        DBDataSet1.Clear()

        '---------------------------------------------------------------------------------------------------------
        'ZZ  曠職
        '---------------------------------------------------------------------------------------------------------

        SQL = " select top 5 * from  M_PersonAdjustTime "
        SQL &= " where active = 1 and yy = '" + DYear.Text + "'"
        SQL &= " and adjusttype ='Z1'"
        SQL &= " and empid = '" + DEmpID.Text + "'"
        SQL &= " order by adjusttype,seqno "

        Dim DBAdapter5 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter5.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")

        Dim wDZ1SUM As Integer = 0

        If DBTable1.Rows.Count > 0 Then

            For i = 0 To DBTable1.Rows.Count - 1
                If i = 0 Then
                    DZZMin1.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DZZDesc1.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 1 Then
                    DZZMin2.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DZZDesc2.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 2 Then
                    DZZMin3.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DZZDesc3.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 3 Then
                    DZZMin4.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DZZDesc4.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 4 Then
                    DZZMin5.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DZZDesc5.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                End If
                wDZ1SUM = wDZ1SUM + Int(Trim(DBTable1.Rows(i).Item("AdjustTime")))
            Next
        End If
        DZZSum.Text = wDZ1SUM
        DBDataSet1.Clear()

        DY2Min.Style.Add("TEXT-ALIGN", "right")
        DY2Sum.Style.Add("TEXT-ALIGN", "right")

        DY3Min1.Style.Add("TEXT-ALIGN", "right")
        DY3Min2.Style.Add("TEXT-ALIGN", "right")
        DY3Min3.Style.Add("TEXT-ALIGN", "right")
        DY3Min4.Style.Add("TEXT-ALIGN", "right")
        DY3Min5.Style.Add("TEXT-ALIGN", "right")
        DY3Sum.Style.Add("TEXT-ALIGN", "right")

        DX2Min1.Style.Add("TEXT-ALIGN", "right")
        DX2Min2.Style.Add("TEXT-ALIGN", "right")
        DX2Min3.Style.Add("TEXT-ALIGN", "right")
        DX2Min4.Style.Add("TEXT-ALIGN", "right")
        DX2Min5.Style.Add("TEXT-ALIGN", "right")
        DX2Sum.Style.Add("TEXT-ALIGN", "right")

        DZZMin1.Style.Add("TEXT-ALIGN", "right")
        DZZMin2.Style.Add("TEXT-ALIGN", "right")
        DZZMin3.Style.Add("TEXT-ALIGN", "right")
        DZZMin4.Style.Add("TEXT-ALIGN", "right")
        DZZMin5.Style.Add("TEXT-ALIGN", "right")
        DZZSum.Style.Add("TEXT-ALIGN", "right")
    End Sub
    '*****************************************************************
    '**
    '**     基準部門點選
    '**
    '*****************************************************************
    Protected Sub DDivision_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDivision.SelectedIndexChanged
        '開始日期
        Startdate = DYear.Text + "/04/01"
        Enddate = DBaseDate.Text
        Firstdate = CStr(CInt(DYear.Text) + 1) + "/01/01"

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
        SQL = SQL + " and convert(char(10),cdate,111) between '" + Startdate + "' and '" + Enddate + "' "
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
        '
        OleDbConnection1.Close()
        '
        WorkRatio()             '出勤率
        '
        ShowVacationInfor()     '請假履歷
        '
        ShowOtherInfor()        '其他說明
    End Sub
    '*****************************************************************
    '**
    '**     基準姓名點選
    '**
    '*****************************************************************
    Protected Sub DName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DName.SelectedIndexChanged
        '開始日期
        Startdate = DYear.Text + "/04/01"
        Enddate = DBaseDate.Text
        Firstdate = CStr(CInt(DYear.Text) + 1) + "/01/01"

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
        SQL = SQL + " and convert(char(10),cdate,111) between '" + Startdate + "' and '" + Enddate + "' "
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
        '
        OleDbConnection1.Close()
        '
        WorkRatio()             '出勤率
        '
        ShowVacationInfor()     '請假履歷
        '
        ShowOtherInfor()        '其他說明
    End Sub
    '*****************************************************************
    '**
    '**     出勤率資料清空
    '**
    '*****************************************************************
    Sub ClearWorkRatioData()
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
    '*****************************************************************
    '**
    '**     其他說明資料清空
    '**
    '*****************************************************************
    Sub ClearOtherInforData()
        DY2Min.Text = ""
        DY2Sum.Text = ""

        DY3Min1.Text = ""
        DY3Min2.Text = ""
        DY3Min3.Text = ""
        DY3Min4.Text = ""
        DY3Min5.Text = ""
        DY3Sum.Text = ""

        DX2Min1.Text = ""
        DX2Min2.Text = ""
        DX2Min3.Text = ""
        DX2Min4.Text = ""
        DX2Min5.Text = ""
        DX2Sum.Text = ""

        DZZMin1.Text = ""
        DZZMin2.Text = ""
        DZZMin3.Text = ""
        DZZMin4.Text = ""
        DZZMin5.Text = ""
        DZZSum.Text = ""
    End Sub
    '*****************************************************************
    '**
    '**     設定ImageButton的圖檔
    '**
    '*****************************************************************
    Protected Sub SetImageButtonImageFile(ByVal pType As Integer)
        If pType = 0 Then
            MultiView1.ActiveViewIndex = pType
            ImageButton1.ImageUrl = "Images/WorkRatioInfor_B1_Blank.jpg"
            ImageButton2.ImageUrl = "Images/WorkRatioInfor_B2.jpg"
            ImageButton3.ImageUrl = "Images/WorkRatioInfor_B3.jpg"
        End If
        If pType = 1 Then
            MultiView1.ActiveViewIndex = pType
            ImageButton1.ImageUrl = "Images/WorkRatioInfor_B1.jpg"
            ImageButton2.ImageUrl = "Images/WorkRatioInfor_B2_Blank.jpg"
            ImageButton3.ImageUrl = "Images/WorkRatioInfor_B3.jpg"
        End If
        If pType = 2 Then
            MultiView1.ActiveViewIndex = pType
            ImageButton1.ImageUrl = "Images/WorkRatioInfor_B1.jpg"
            ImageButton2.ImageUrl = "Images/WorkRatioInfor_B2.jpg"
            ImageButton3.ImageUrl = "Images/WorkRatioInfor_B3_Blank.jpg"
        End If
    End Sub
    '*****************************************************************
    '**
    '**     設定ImageButton事件處理程序
    '**
    '*****************************************************************
    Protected Sub SetImageButtonEvent()
        'ImageButton1.Attributes.Add("OnMouseOver", "ChangeImage(this)")
        'ImageButton2.Attributes.Add("OnMouseOver", "ChangeImage(this)")
        'ImageButton3.Attributes.Add("OnMouseOver", "ChangeImage(this)")
    End Sub
    '*****************************************************************
    '**
    '**     ImageButton3 事件處理程序
    '**
    '*****************************************************************
    Protected Sub ImageButton3_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton3.Click
        SetImageButtonImageFile(2)
    End Sub
    '*****************************************************************
    '**
    '**     ImageButton2 事件處理程序
    '**
    '*****************************************************************
    Protected Sub ImageButton2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton2.Click
        SetImageButtonImageFile(1)
    End Sub
    '*****************************************************************
    '**
    '**     ImageButton1 事件處理程序
    '**
    '*****************************************************************
    Protected Sub ImageButton1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton1.Click
        SetImageButtonImageFile(0)
    End Sub

End Class
