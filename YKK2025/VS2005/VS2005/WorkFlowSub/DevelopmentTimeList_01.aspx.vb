Imports System.Data
Imports System.Data.OleDb

Partial Class DevelopmentTimeList_01
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

    Dim oSchedule As New Schedule.ScheduleService

    Dim NowDateTime As String       '現在日期
    Dim NowYear As String           '現在年
    Dim NowMonth As String          '現在月
    Dim wUserID As String           'UserID
    Dim StartDate As Date           '起始日
    Dim EndDate As Date             '迄
    Dim WorkTable As String = ""    'Temp Table
    'Table Field Define
    Dim fNo As String = ""
    Dim fBuyer As String = ""
    Dim fMapNo As String = ""
    Dim fLevel As String = ""
    '
    Dim fMapCreateTime As String = ""
    Dim fMapFinishTime As String = ""
    Dim fMapSts As String = ""
    Dim fMapModCount As Integer = 0
    '
    Dim fManufInURL As String = ""
    Dim fManufInCreateTime As String = ""
    Dim fManufInFinishTime As String = ""
    Dim fManufInSts As String = ""
    Dim fManufInModCount As Integer = 0
    '
    Dim fManufOutURL As String = ""
    Dim fManufOutCreateTime As String = ""
    Dim fManufOutFinishTime As String = ""
    Dim fManufOutSts As String = ""
    Dim fManufOutModCount As Integer = 0
    Dim pDevelopTime1 As Integer
    Dim pDevelopTime As String
    '
    'Add-Start by Joy 2009/11/20(2010行事曆對應)
    '群組行事曆
    Dim wDepo As String = "CL1"      '中壢行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
    'Add-End

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        Response.Cookies("PGM").Value = "DevelopmentTimeList_01.aspx"
        Server.ScriptTimeout = 900   '900秒

        SetParameter()
        If Not Me.IsPostBack Then
            SetSearchItem()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(DateTime.Now.Date) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)
        NowYear = CStr(DateTime.Now.Year)
        NowMonth = CStr(DateTime.Now.Month)
        wUserID = Request.QueryString("pUserID")
        WorkTable = "Temp_DevelopmentTime01_" & wUserID   'Temp Table
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選條件
    '**
    '*****************************************************************
    Sub SetSearchItem()
        Dim i As Integer
        Dim ListItem1 As New ListItem
        '年 
        For i = 2007 To 2023
            DSYY.Items.Add(CStr(i))
            DEYY.Items.Add(CStr(i))
        Next
        ListItem1.Text = NowYear
        ListItem1.Value = NowYear
        DSYY.SelectedIndex = DSYY.Items.IndexOf(ListItem1)
        DEYY.SelectedIndex = DEYY.Items.IndexOf(ListItem1)
        '月 
        For i = 1 To 12
            If i < 10 Then
                DSMM.Items.Add("0" + CStr(i))
                DEMM.Items.Add("0" + CStr(i))
            Else
                DSMM.Items.Add(CStr(i))
                DEMM.Items.Add(CStr(i))
            End If
        Next
        ListItem1.Text = NowMonth
        ListItem1.Value = NowMonth
        DSMM.SelectedIndex = DSMM.Items.IndexOf(ListItem1)
        DEMM.SelectedIndex = DEMM.Items.IndexOf(ListItem1)
        'Buyer 
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        DBuyer.Items.Add("ALL")

        OleDbConnection1.Open()
        SQL = "Select Data From M_Referp "
        SQL = SQL + "Where Cat='700' "
        SQL = SQL + "  And DKey='BUYER' "
        SQL = SQL + "Order by Data "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Buyer")
        DBTable1 = DBDataSet1.Tables("Buyer")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem2 As New ListItem
            ListItem2.Text = DBTable1.Rows(i).Item("Data")
            ListItem2.Value = DBTable1.Rows(i).Item("Data")
            DBuyer.Items.Add(ListItem2)
        Next
        '資料更新時點
        DBDataSet1.Clear()
        DLastUpdateTime.Text = ""
        SQL = "SELECT Convert(VARCHAR(20), CreateTime, 120) as CreateTime From W_CommissionStructure "
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "Sheet")
        If DBDataSet1.Tables("Sheet").Rows.Count > 0 Then
            DLastUpdateTime.Text = "資料時點：" + DBDataSet1.Tables("Sheet").Rows(0).Item("CreateTime")
        End If

        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Go-Button
    '**
    '*****************************************************************
    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        StartDate = CDate(DSYY.SelectedValue + "/" + DSMM.SelectedValue + "/" + "1")
        EndDate = CDate(DEYY.SelectedValue + "/" + DEMM.SelectedValue + "/" + "1").AddMonths(1)
        'If DateDiff(DateInterval.Day, StartDate, EndDate) > 62 Then
        '    Response.Write(YKK.ShowMessage("期間指定不可超過２個月！"))
        'Else
        SelectData()
        'End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選資料
    '**
    '*****************************************************************
    Sub SelectData()
        Dim i, j As Integer
        Dim SQL As String
        Dim DBDataSet1, DBDataSet2, DBDataSet3, DBDataSet4 As New DataSet
        Dim DBTable1, DBTable2, DBTable3, DBTable4 As DataTable
        Dim OleDBCommand1 As New OleDbCommand
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        'Break Key
        Dim BKey As String = ""
        Dim FirstRecord As Boolean = False
        '
        OleDbConnection1.Open()
        '符合篩選條件資料
        DBDataSet1.Clear()
        SQL = "Select No, Buyer, MapNo, Level From F_MapSheet "
        SQL = SQL + "Where Date  >= '" + CStr(StartDate) + "' "
        SQL = SQL + "  And Date  <  '" + CStr(EndDate) + "' "
        'Buyer
        If DBuyer.SelectedValue <> "ALL" Then
            SQL = SQL + "  And Buyer =  '" + DBuyer.SelectedValue + "' "
        End If
        SQL = SQL + "  And MapNo <> '" + "" + "' "
        SQL = SQL + "Order by Date "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Map")
        DBTable1 = DBDataSet1.Tables("Map")
        If DBTable1.Rows.Count > 0 Then
            'Call Stored Procedure Create WorkTable 
            SQL = "Exec sp_Temp_DevelopmentTime01 " & WorkTable & " "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDBCommand1.ExecuteNonQuery()
        Else
            SQL = "Delete From " & WorkTable & " "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDBCommand1.ExecuteNonQuery()
        End If
        For i = 0 To DBTable1.Rows.Count - 1
            'Break Key                                                                     
            BKey = ""
            FirstRecord = False
            'Table Field-初始值
            fNo = DBTable1.Rows(i).Item("No")
            fBuyer = DBTable1.Rows(i).Item("Buyer")
            fMapNo = DBTable1.Rows(i).Item("MapNo")
            fLevel = DBTable1.Rows(i).Item("Level")
            '
            fMapCreateTime = ""
            fMapFinishTime = ""
            fMapSts = ""
            fMapModCount = 0
            '
            fManufInURL = ""
            fManufInCreateTime = ""
            fManufInFinishTime = ""
            fManufInSts = ""
            fManufInModCount = 0
            '
            fManufOutURL = ""
            fManufOutCreateTime = ""
            fManufOutFinishTime = ""
            fManufOutSts = ""
            fManufOutModCount = 0
            '搜尋結構資料
            DBDataSet2.Clear()
            SQL = "Select FormNo, FormSno From V_CommissionStructure_01 "
            SQL = SQL + "Where MapNo  = '" + DBTable1.Rows(i).Item("MapNo") + "' "
            SQL = SQL + "  And ( (FormNo between '000001' and '000003')  or FormNo='000007') "
            SQL = SQL + "Order by FormNo, BaseTime "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "Structure")
            DBTable2 = DBDataSet2.Tables("Structure")
            For j = 0 To DBTable2.Rows.Count - 1
                DBDataSet3.Clear()
                If DBTable2.Rows(j).Item("FormNo") = "000001" Or DBTable2.Rows(j).Item("FormNo") = "000002" Then
                    '----圖面or圖面修改處理-----------------------------------------------------------------
                    SQL = "Select FormNo, FormSno, FSts, "
                    SQL = SQL + "Convert(VARCHAR(20), ApplyTime, 120) as ApplyTime, "
                    SQL = SQL + "Convert(VARCHAR(20), CompletedTime, 120) as CompletedTime, "
                    SQL = SQL + "Case FSts When 0 Then '開發中' When 1 Then '完成' When 2 Then 'ＮＧ' Else '取消' End As FStsDesc "
                    SQL = SQL + "From V_Waithandle_01 "
                    SQL = SQL + "Where FormNo   = '" + DBTable2.Rows(j).Item("FormNo") + "' "
                    SQL = SQL + "  And FormSno  = '" + DBTable2.Rows(j).Item("FormSno").ToString + "' "
                    SQL = SQL + "  And Step     = '" + "130" + "' "
                    SQL = SQL + "Order by CompletedTime Desc "
                    Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter3.Fill(DBDataSet3, "WaitHandle")
                    DBTable3 = DBDataSet3.Tables("WaitHandle")
                    If DBTable3.Rows.Count > 0 Then
                        '圖面or圖面修改--130工程完成
                        If DBTable3.Rows(0).Item("FormNo") = "000001" Then
                            fMapCreateTime = DBTable3.Rows(0).Item("ApplyTime").ToString
                        End If
                        '    fMapFinishTime = DBTable3.Rows(0).Item("CompletedTime").ToString
                        fMapSts = DBTable3.Rows(0).Item("FStsDesc")
                    Else
                        ''圖面or圖面修改--130工程未完成
                        'DBDataSet4.Clear()
                        'SQL = "Select FormNo, FormSno, FSts, "
                        'SQL = SQL + "Convert(VARCHAR(20), ApplyTime, 120) as ApplyTime, "
                        'SQL = SQL + "Convert(VARCHAR(20), CompletedTime, 120) as CompletedTime, "
                        'SQL = SQL + "Case FSts When 0 Then '開發中' When 1 Then '完成' When 2 Then 'ＮＧ' Else '取消' End As FStsDesc "
                        'SQL = SQL + "From V_Waithandle_01 "
                        'SQL = SQL + "Where FormNo   = '" + DBTable2.Rows(j).Item("FormNo") + "' "
                        'SQL = SQL + "  And FormSno  = '" + DBTable2.Rows(j).Item("FormSno").ToString + "' "
                        'SQL = SQL + "Order by CompletedTime Desc "
                        'Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
                        'DBAdapter4.Fill(DBDataSet4, "WaitHandle")
                        'DBTable4 = DBDataSet4.Tables("WaitHandle")
                        'If DBTable4.Rows.Count > 0 Then
                        '    If DBTable4.Rows(0).Item("FSts") <> 1 Then  '上線前資料判斷(狀態=完成 & 沒有130工程時不列入)
                        '        If DBTable4.Rows(0).Item("FormNo") = "000001" Then
                        '            fMapCreateTime = DBTable4.Rows(0).Item("ApplyTime").ToString
                        '        End If
                        '        fMapSts = DBTable4.Rows(0).Item("FStsDesc")
                        '    End If
                        'End If
                    End If
                    fMapModCount = fMapModCount + 1
                    '--------------------------------------------------------------------
                Else
                    '----內製or外注處理-----------------------------------------------------------------
                    'Set Break Key
                    If BKey <> DBTable2.Rows(j).Item("FormNo") Then
                        FirstRecord = True
                        BKey = DBTable2.Rows(j).Item("FormNo")
                    Else
                        FirstRecord = False
                    End If
                    SQL = "Select FormNo, FormSno, FSts, ViewURL, "
                    SQL = SQL + "Convert(VARCHAR(20), ApplyTime, 120) as ApplyTime, "
                    SQL = SQL + "Convert(VARCHAR(20), CompletedTime, 120) as CompletedTime, "
                    SQL = SQL + "Case FSts When 0 Then '開發中' When 1 Then '完成' When 2 Then 'ＮＧ' Else '取消' End As FStsDesc "
                    SQL = SQL + "From V_Waithandle_01 "
                    SQL = SQL + "Where FormNo   = '" + DBTable2.Rows(j).Item("FormNo") + "' "
                    SQL = SQL + "  And FormSno  = '" + DBTable2.Rows(j).Item("FormSno").ToString + "' "
                    SQL = SQL + "  And Step     = '" + "170" + "' "
                    SQL = SQL + "Order by CompletedTime Desc "
                    Dim DBAdapter13 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter13.Fill(DBDataSet3, "WaitHandle")
                    DBTable3 = DBDataSet3.Tables("WaitHandle")
                    If DBTable3.Rows.Count > 0 Then
                        '內製or外注--170工程完成
                        If DBTable3.Rows(0).Item("FormNo") = "000003" Then
                            If FirstRecord Then
                                '      fManufInCreateTime = DBTable3.Rows(0).Item("ApplyTime").ToString
                            End If
                            fManufInURL = DBTable3.Rows(0).Item("ViewURL")
                            fManufInFinishTime = DBTable3.Rows(0).Item("CompletedTime").ToString
                            ' fManufInSts = DBTable3.Rows(0).Item("FStsDesc")
                            ' fManufInModCount = fManufInModCount + 1
                        Else
                            If FirstRecord Then
                                '    fManufOutCreateTime = DBTable3.Rows(0).Item("ApplyTime").ToString
                            End If
                            fManufOutURL = DBTable3.Rows(0).Item("ViewURL")
                            fManufOutFinishTime = DBTable3.Rows(0).Item("CompletedTime").ToString
                            ' fManufOutSts = DBTable3.Rows(0).Item("FStsDesc")
                            ' fManufOutModCount = fManufOutModCount + 1
                        End If
                    Else
                        ''內製or外注--160工程未完成
                        'DBDataSet4.Clear()
                        'SQL = "Select FormNo, FormSno, FSts, ViewURL, "
                        'SQL = SQL + "Convert(VARCHAR(20), ApplyTime, 120) as ApplyTime, "
                        'SQL = SQL + "Convert(VARCHAR(20), CompletedTime, 120) as CompletedTime, "
                        'SQL = SQL + "Case FSts When 0 Then '開發中' When 1 Then '完成' When 2 Then 'ＮＧ' Else '取消' End As FStsDesc "
                        'SQL = SQL + "From V_Waithandle_01 "
                        'SQL = SQL + "Where FormNo   = '" + DBTable2.Rows(j).Item("FormNo") + "' "
                        'SQL = SQL + "  And FormSno  = '" + DBTable2.Rows(j).Item("FormSno").ToString + "' "
                        'SQL = SQL + "Order by CompletedTime Desc "
                        'Dim DBAdapter14 As New OleDbDataAdapter(SQL, OleDbConnection1)
                        'DBAdapter14.Fill(DBDataSet4, "WaitHandle")
                        'DBTable4 = DBDataSet4.Tables("WaitHandle")
                        'If DBTable4.Rows.Count > 0 Then
                        '    If DBTable4.Rows(0).Item("FSts") <> 1 Then  '上線前資料判斷(狀態=完成 & 沒有160工程時不列入)
                        '        If DBTable4.Rows(0).Item("FormNo") = "000003" Then
                        '            If FirstRecord Then
                        '                fManufInCreateTime = DBTable4.Rows(0).Item("ApplyTime").ToString
                        '            End If
                        '            fManufInURL = DBTable4.Rows(0).Item("ViewURL")
                        '            fManufInSts = DBTable4.Rows(0).Item("FStsDesc")
                        '            fManufInModCount = fManufInModCount + 1
                        '        Else
                        '            If FirstRecord Then
                        '                fManufOutCreateTime = DBTable4.Rows(0).Item("ApplyTime").ToString
                        '            End If
                        '            fManufOutURL = DBTable4.Rows(0).Item("ViewURL")
                        '            fManufOutSts = DBTable4.Rows(0).Item("FStsDesc")
                        '            fManufOutModCount = fManufOutModCount + 1
                        '        End If
                        '    End If
                        'End If
                    End If
                    '--------------------------------------------------------------------
                End If
            Next    '搜尋結構資料
            OleDbConnection1.Close()

            '寫至Work File
            If fMapCreateTime = "" And fMapFinishTime = "" And _
               fManufInCreateTime = "" And fManufInFinishTime = "" And _
               fManufOutCreateTime = "" And fManufOutFinishTime = "" Then
                '無效資料
            Else
                '有效資料
               
                pDevelopTime1 = 0
                pDevelopTime = 0
                Dim pStartTime = ""
                Dim pEndTime = ""
                pStartTime = fMapCreateTime

                If fManufInFinishTime > fManufOutFinishTime Then
                    pEndTime = fManufInFinishTime
                Else
                    pEndTime = fManufOutFinishTime
                End If
                If oSchedule.GetDevelopTime(pStartTime, pEndTime, pDevelopTime1, "CL1") = 0 Then  '計算開發時間
                    pDevelopTime = FormatNumber(pDevelopTime1 / 60, 0)
                Else
                    pDevelopTime = ""
                End If
                CreateWorkFile()
            End If
        Next        '符合篩選條件資料
        '顯示至畫面
        ShowData()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     產生WorkFile
    '**
    '*****************************************************************
    Sub CreateWorkFile()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDBCommand1 As New OleDbCommand
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        '建立Table
        SQL = "Insert into " & WorkTable & " "
        SQL = SQL + "( "
        SQL = SQL + "MNo, Buyer, MapNo, MLevel, MapCreateTime,MapSts,MapET, "
        SQL = SQL + "MISUrl, ManufInFinishTime,"
        SQL = SQL + "MOSUrl, ManufOutFinishTime,"
        SQL = SQL + "CreateUser, CreateTime "
        SQL = SQL + ") "
        SQL = SQL + "VALUES ("
        SQL = SQL + " '" + fNo + "', "
        SQL = SQL + " '" + fBuyer + "', "
        SQL = SQL + " '" + fMapNo + "', "
        SQL = SQL + " '" + fLevel + "', "
        SQL = SQL + " '" + fMapCreateTime + "', "
        SQL = SQL + " '" + fMapSts + "', " '狀態
        SQL = SQL + " '" + pDevelopTime + "', " '開發時間   
       
        SQL = SQL + " '" + fManufInURL + "', "  '內製url
        SQL = SQL + " '" + fManufInFinishTime + "', " '內製完成時間

        SQL = SQL + " '" + fManufOutURL + "', " '外製url
        SQL = SQL + " '" + fManufOutFinishTime + "', " '外製完成時間
        
        SQL = SQL + " '" + wUserID + "', "     '作成者
        SQL = SQL + " '" + NowDateTime + "' "       '作成時間
        SQL = SQL + ") "
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQL
        OleDBCommand1.ExecuteNonQuery()
        '
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     顯示畫面
    '**
    '*****************************************************************
    Sub ShowData()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '資料顯示
        OleDbConnection1.Open()
        DBDataSet1.Clear()
        SQL = "SELECT * From " + WorkTable
        SQL = SQL + " Order by Buyer, MapCreateTime "
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "Data")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     計算時間
    '**
    '*****************************************************************
    Function GetTime(ByVal StartDateTime As Date, ByVal EndDateTime As Date) As Integer
        Dim SQL As String
        Dim DBDataSet1, DBDataSet2 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '
        Dim StartDate As String = CStr(StartDateTime.Date)
        Dim StartMinute As Integer = CInt(StartDateTime.Hour) * 60 + CInt(StartDateTime.Minute)
        Dim EndDate As String = CStr(EndDateTime.Date)
        Dim EndMinute As Integer = CInt(EndDateTime.Hour) * 60 + CInt(EndDateTime.Minute)
        Dim StartNo As Integer = 0
        Dim EndNo As Integer = 0
        '
        GetTime = 0
        OleDbConnection1.Open()
        '取得開始時間
        DBDataSet1.Clear()
        SQL = "Select * From M_Calendar "
        SQL = SQL + "Where YMD    =  '" + StartDate + "' "
        SQL = SQL + "  And Hour  >= '" + CStr(StartMinute) + "' "
        SQL = SQL + "  And Active = '1' "

        'Modify-Start by Joy 2009/11/20(2010行事曆對應)
        'SQL = SQL + "  and Depo = 'CL' "
        '
        SQL = SQL + "  And Depo   = '" + wDepo + "' "
        'Modify-End

        SQL = SQL + "Order by YMD, Hour "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Calendar")
        DBTable1 = DBDataSet1.Tables("Calendar")
        If DBTable1.Rows.Count > 0 Then
            StartNo = DBTable1.Rows(0).Item("SeqNo")
        Else
            DBDataSet2.Clear()
            SQL = "Select * From M_Calendar "
            SQL = SQL + "Where YMD   >  '" + StartDate + "' "
            SQL = SQL + "  And Active = '1' "

            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
            'SQL = SQL + "  and Depo = 'CL' "
            '
            SQL = SQL + "  And Depo   = '" + wDepo + "' "
            'Modify-End

            SQL = SQL + "Order by YMD, Hour "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "Calendar")
            DBTable1 = DBDataSet2.Tables("Calendar")
            If DBTable1.Rows.Count > 0 Then
                StartNo = DBTable1.Rows(0).Item("SeqNo")
            End If
        End If
        '取得完成時間
        DBDataSet1.Clear()
        SQL = "Select * From M_Calendar "
        SQL = SQL + "Where YMD    =  '" + EndDate + "' "
        SQL = SQL + "  And Hour  >= '" + CStr(EndMinute) + "' "
        SQL = SQL + "  And Active = '1' "

        'Modify-Start by Joy 2009/11/20(2010行事曆對應)
        'SQL = SQL + "  and Depo = 'CL' "
        '
        SQL = SQL + "  And Depo   = '" + wDepo + "' "
        'Modify-End

        SQL = SQL + "Order by YMD, Hour "
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "Calendar")
        DBTable1 = DBDataSet1.Tables("Calendar")
        If DBTable1.Rows.Count > 0 Then
            EndNo = DBTable1.Rows(0).Item("SeqNo")
        Else
            DBDataSet2.Clear()
            SQL = "Select * From M_Calendar "
            SQL = SQL + "Where YMD   >  '" + EndDate + "' "
            SQL = SQL + "  And Active = '1' "

            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
            'SQL = SQL + "  and Depo = 'CL' "
            '
            SQL = SQL + "  And Depo   = '" + wDepo + "' "
            'Modify-End

            SQL = SQL + "Order by YMD, Hour "
            Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter4.Fill(DBDataSet2, "Calendar")
            DBTable1 = DBDataSet2.Tables("Calendar")
            If DBTable1.Rows.Count > 0 Then
                EndNo = DBTable1.Rows(0).Item("SeqNo")
            End If
        End If
        OleDbConnection1.Close()

        GetTime = (EndNo - StartNo) * 10
    End Function

    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click

        Response.AppendHeader("Content-Disposition", "attachment;filename=DevelopmentTimeList_01.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

    End Sub
End Class
