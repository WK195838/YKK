Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class DevelopmentTimeList
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DBuyer As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DSYear As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSMonth As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DEYear As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DEMonth As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label
    Protected WithEvents Label4 As System.Web.UI.WebControls.Label

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

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
    Dim pFormNo As String
    Dim NowYear As String
    Dim NowDateTime As String       '現在日期時間
    Dim SCreateDate, ECreateDate As String       ''製圖委託書委託日期(起,迄)


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "DevelopmentTimeList.aspx"

        SetParameter()
        If Not Me.IsPostBack Then
            SetYear()
            SetBuyer()
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '現在日時
        NowYear = CStr(DateTime.Now.Year)

        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選製圖委託書委託年份
    '**
    '*****************************************************************
    Sub SetYear()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection

        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "SELECT CONVERT(varchar(4), CreateTime, 120) AS SetYear FROM F_MapSheet "
        SQL = SQL + "GROUP BY CONVERT(varchar(4), CreateTime, 120)"

        DSYear.Items.Clear()
        DEYear.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Year")
        DBTable1 = DBDataSet1.Tables("Year")
        '將篩選出的年份加入DropDownList
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("SetYear")
            ListItem1.Value = DBTable1.Rows(i).Item("SetYear")
            If ListItem1.Value = NowYear Then ListItem1.Selected = True
            DSYear.Items.Add(ListItem1)
            DEYear.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選月份BUYER
    '**
    '*****************************************************************
    Sub SetBuyer()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection

        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()


        If CDate(DSYear.SelectedValue + "-" + DSMonth.SelectedValue + "-1") > CDate(DEYear.SelectedValue + "-" + DEMonth.SelectedValue + "-1") Then
            SCreateDate = DEYear.SelectedValue + "-" + DEMonth.SelectedValue + "-1"
            ECreateDate = DSYear.SelectedValue + "-" + DSMonth.SelectedValue + "-1"
        Else
            SCreateDate = DSYear.SelectedValue + "-" + DSMonth.SelectedValue + "-1"
            ECreateDate = DEYear.SelectedValue + "-" + DEMonth.SelectedValue + "-1"
        End If

        SQL = "SELECT a.Buyer FROM F_MapSheet a, T_WaitHandle b "       '在委託日期為該月份的製圖委託書中,有工程130的篩選出來
        SQL = SQL + "WHERE (a.CreateTime > DATEADD(d, 0, '" + SCreateDate + "')) AND (a.CreateTime < DATEADD(m, 1, '" + ECreateDate + "')) "
        SQL = SQL + "AND NOT(a.Sts = '3') AND a.Formno = b.Formno AND a.Formsno = b.Formsno and b.Step = '130' "
        SQL = SQL + "AND NOT(b.CompletedTime IS NULL) GROUP BY a.Buyer"

        DBuyer.Items.Clear()
        DBuyer.Items.Add("ALL")
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "BuyerName")
        DBTable1 = DBDataSet1.Tables("BuyerName")
        '將篩選出的BUYER加入DropDownListr 
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("Buyer")
            ListItem1.Value = DBTable1.Rows(i).Item("Buyer")
            DBuyer.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        If DBuyer.SelectedValue <> "" Then
            CreateWorkTable()
        Else
            DataGrid1.Visible = False
            Label1.Visible = False
            Label2.Visible = False
            Label3.Visible = False
            Label4.Visible = False
        End If
    End Sub

    Sub CreateWorkTable()
        Dim SQL, SQL1, SQL2, wStep As String
        Dim i, j, Count As Integer
        Dim MapModStsCount As Integer   '圖面修改狀態為NG及取消的數量
        Dim MapNo, MapCreateTime, MapFinishTime, MapModTime, MapSts As String               '圖號,繪圖起單時間,繪圖完成時間,繪圖經過時間,繪圖開發狀態
        Dim MapModCount As Integer      '繪圖修改次數
        Dim ManufInCreateTime, ManufInFinishTime, ManufInModTime, ManufInSts, MapToInET As String     '內製起單時間,內製完成時間,內製經過時間,內製開發狀態,繪圖~內製經過時間
        Dim ManufInModCount As Integer  '內製修改次數
        Dim ManufOutCreateTime, ManufOutFinishTime, ManufOutModTime, ManufOutSts, MapToOutET As String '外注起單時間,外注完成時間,外注經過時間,外注開發狀態,繪圖~外注經過時間
        Dim ManufOutModCount As Integer '外注修改次數
        Dim Buyer As String
        Dim MOSUrl, MISUrl As String    '外注,內製連結暫存

        Dim WorkTable As String = "Temp_DevelopmentTimeList_" & Request.Cookies("UserID").Value

        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBDataSet3 As New DataSet
        Dim DBDataSet4 As New DataSet
        Dim DBDataSet5 As New DataSet
        Dim DBDataSet6 As New DataSet
        Dim DBDataSet7 As New DataSet
        Dim DBDataSet8 As New DataSet
        Dim DBTable1 As DataTable
        Dim DBDataRow As DataRow

        Dim OleDBCommand1 As New OleDbCommand
        Dim OleDbConnection1 As New OleDbConnection

        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        If CDate(DSYear.SelectedValue + "-" + DSMonth.SelectedValue + "-1") > CDate(DEYear.SelectedValue + "-" + DEMonth.SelectedValue + "-1") Then
            SCreateDate = DEYear.SelectedValue + "-" + DEMonth.SelectedValue + "-1"
            ECreateDate = DSYear.SelectedValue + "-" + DSMonth.SelectedValue + "-1"
        Else
            SCreateDate = DSYear.SelectedValue + "-" + DSMonth.SelectedValue + "-1"
            ECreateDate = DEYear.SelectedValue + "-" + DEMonth.SelectedValue + "-1"
        End If

        DataGrid1.Visible = True
        Label1.Visible = True
        Label2.Visible = True
        Label3.Visible = True
        Label4.Visible = True

        If DBuyer.SelectedValue <> "ALL" Then
            SQL1 = "AND a.Buyer = '" + DBuyer.SelectedValue + "' "
        End If

        '篩選條件1.年月,2.BUYER,3.狀態非取消,4.Formno & Formsno相同,5.製圖委託書有130工程
        SQL = "SELECT a.No,a.Buyer,a.Level,a.Sts,a.CreateTime,b.CompletedTime,a.ModMap,a.MapNo FROM F_MapSheet a,T_WaitHandle b "
        SQL = SQL + "WHERE  (a.CreateTime > DATEADD(d, 0, '" + SCreateDate + "')) AND (a.CreateTime < DATEADD(m, 1, '" + ECreateDate + "')) "
        SQL = SQL + SQL1    'Buyer
        SQL = SQL + "AND NOT (a.Sts = '3') AND a.FormNo = b.FormNo AND a.Formsno = b.Formsno "
        SQL = SQL + "AND b.Step = '130' "
        SQL = SQL + "AND NOT(b.CompletedTime IS NULL) "
        SQL = SQL + "ORDER BY a.Buyer,a.CreateTime"

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "MapSheetInfo")

        'Call Stored Procedure Create WorkTable 
        SQL = "Exec sp_Temp_DevelopmentTimeList '" & WorkTable & "'"
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQL
        OleDBCommand1.ExecuteNonQuery()


        For i = 0 To DBDataSet1.Tables("MapSheetInfo").Rows.Count - 1
            SQL1 = ""
            SQL2 = ""
            MapCreateTime = CStr(DBDataSet1.Tables("MapSheetInfo").Rows(i).Item("createtime"))
            MapModCount = 0
            MapFinishTime = ""
            MapModTime = ""
            MapNo = DBDataSet1.Tables("MapSheetInfo").Rows(i).Item("MapNo")
            MapFinishTime = CStr(DBDataSet1.Tables("MapSheetInfo").Rows(i).Item("CompletedTime"))
            MapSts = DBDataSet1.Tables("MapSheetInfo").Rows(i).Item("Sts")

            If DBDataSet1.Tables("MapSheetInfo").Rows(i).Item("ModMap") = 1 Then    '判斷是否有圖面修改
                MapModStsCount = 0
                '篩選條件圖號相同
                SQL = "SELECT FormNo,Formsno,Sts,MapNo FROM F_MapModSheet WHERE OriMapNo = '" + MapNo + " ' ORDER BY Unique_ID DESC"

                DBDataSet2.Clear()
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet2, "MapModSheetInfo")

                MapModCount = DBDataSet2.Tables("MapModSheetInfo").Rows.Count       '記錄繪圖修改次數

                For j = 0 To MapModCount - 1
                    '篩選條件1.Formno & Formsno相同,2.製圖委託書有130工程
                    SQL = "SELECT CompletedTime FROM T_WaitHandle WHERE FormNo = '"
                    SQL = SQL + DBDataSet2.Tables("MapModSheetInfo").Rows(j).Item("Formno") + "' AND Formsno = '"
                    SQL = SQL + CStr(DBDataSet2.Tables("MapModSheetInfo").Rows(j).Item("Formsno")) + "' AND Step = '130' "
                    SQL = SQL + "AND NOT(CompletedTime IS NULL) "

                    DBDataSet3.Clear()
                    Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter3.Fill(DBDataSet3, "WaitHandleInfo")

                    '判斷是否有130工程且狀態為完成OK或開發中
                    If DBDataSet3.Tables("WaitHandleInfo").Rows.Count >= 1 Then 'And DBDataSet2.Tables("MapModSheetInfo").Rows(j).Item("Sts") < 2 Then
                        MapSts = DBDataSet2.Tables("MapModSheetInfo").Rows(j).Item("Sts")
                        MapFinishTime = CStr(DBDataSet3.Tables("WaitHandleInfo").Rows(0).Item("CompletedTime"))
                        Exit For
                    End If
                Next j

                Count = 0

                '將所有圖面修改的MapNo抓出來
                For j = 0 To MapModCount - 1
                    If DBDataSet2.Tables("MapModSheetInfo").Rows(j).Item("mapno") <> "" Then    '判斷圖號是否為空白
                        If Count = 0 Then   '判斷是否為第一筆
                            SQL1 = SQL1 + " a.MapNo = '" + DBDataSet2.Tables("MapModSheetInfo").Rows(j).Item("mapno") + "' "
                            Count = 1
                        Else
                            SQL1 = SQL1 + "OR a.MapNo = '" + DBDataSet2.Tables("MapModSheetInfo").Rows(j).Item("mapno") + "' "
                        End If
                    End If
                Next j

                If SQL1 <> "" Then      '判斷SQL1是否為空白
                    SQL2 = "SELECT a.Unique_ID,a.No,a.Formno,a.Formsno,a.Sts,a.CreateTime,b.ViewUrl FROM F_ManufInSheet a "
                    SQL2 = SQL2 + "LEFT OUTER JOIN V_WaitHandle_01 b ON a.FormNo=b.FormNo And a.FormSno=b.FormSno WHERE "
                    SQL2 = SQL2 + SQL1
                    SQL2 = SQL2 + "GROUP BY a.Unique_ID,a.No,a.Formno,a.Formsno,a.Sts,a.CreateTime,b.ViewUrl "
                    SQL2 = SQL2 + "UNION "
                    SQL1 = "SELECT a.Unique_ID,LEFT(a.No,1) AS No,a.Formno,a.Formsno,a.Sts,a.CreateTime,b.ViewUrl FROM F_ManufOutSheet a LEFT OUTER JOIN V_WaitHandle_01 b ON a.FormNo=b.FormNo And a.FormSno=b.FormSno WHERE " + SQL1
                    SQL1 = SQL1 + "UNION "
                End If
            End If

            If MapFinishTime <> "" Then '判斷繪圖完成時間是否為空白,若不是的話,則到GetTime計算經過時間
                MapModTime = GetTime(CDate(MapCreateTime), CDate(MapFinishTime))
            End If

            '以製圖或圖面修改的圖號找出外注圖號是相同的RECORD
            SQL = SQL1 + "SELECT a.Unique_ID,LEFT(a.No,1) AS No,a.Formno,a.Formsno,a.Sts,a.CreateTime,b.ViewUrl FROM F_ManufOutSheet a "
            SQL = SQL + "LEFT OUTER JOIN V_WaitHandle_01 b ON a.FormNo=b.FormNo And a.FormSno=b.FormSno WHERE "
            SQL = SQL + "a.MapNo = '" + MapNo + "' "
            SQL = SQL + "GROUP BY a.Unique_ID,a.No,a.Formno,a.Formsno,a.Sts,a.CreateTime,b.ViewUrl "
            SQL = SQL + "ORDER BY a.Unique_ID DESC"

            DBDataSet4.Clear()
            Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter4.Fill(DBDataSet4, "ManufOutSheetInfo")

            ManufOutCreateTime = ""
            ManufOutFinishTime = ""
            ManufOutModTime = ""
            ManufOutSts = ""
            MapToOutET = ""
            ManufOutModCount = 0
            MOSUrl = ""

            '判斷是否有找到
            If DBDataSet4.Tables("ManufOutSheetInfo").Rows.Count > 0 Then
                wStep = "160"   '非自主開發STEP160
                ManufOutModCount = DBDataSet4.Tables("ManufOutSheetInfo").Rows.Count    '外注的修改次數
                ManufOutCreateTime = DBDataSet4.Tables("ManufOutSheetInfo").Rows(ManufOutModCount - 1).Item("CreateTime")   '建立日期為第一筆的完成日期
                MOSUrl = DBDataSet4.Tables("ManufOutSheetInfo").Rows(0).Item("ViewUrl")
                ManufOutSts = DBDataSet4.Tables("ManufOutSheetInfo").Rows(0).Item("Sts")

                '判斷是否為E開頭的委託書(自主開發)
                For j = 0 To ManufOutModCount - 1
                    If DBDataSet4.Tables("ManufOutSheetInfo").Rows(j).Item("No") <> "" And DBDataSet4.Tables("ManufOutSheetInfo").Rows(j).Item("No") = "E" Then
                        wStep = "165"   '若為自主開發STEP為165
                        Exit For
                    End If
                Next j

                For j = 0 To ManufOutModCount - 1
                    '篩選條件1.Formno & Formsno相同,2.外注委託書有160或165工程
                    SQL = "SELECT CompletedTime FROM T_WaitHandle WHERE FormNo = '"
                    SQL = SQL + DBDataSet4.Tables("ManufOutSheetInfo").Rows(j).Item("Formno") + "' AND Formsno = '"
                    SQL = SQL + CStr(DBDataSet4.Tables("ManufOutSheetInfo").Rows(j).Item("Formsno")) + "' AND STEP = '"
                    SQL = SQL + wStep + "' "
                    SQL = SQL + "AND NOT(CompletedTime IS NULL) "

                    DBDataSet5.Clear()
                    Dim DBAdapter5 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter5.Fill(DBDataSet5, "WaitHandleInfo")

                    '判斷是否有160或165工程
                    If DBDataSet5.Tables("WaitHandleInfo").Rows.Count >= 1 Then
                        ManufOutSts = DBDataSet4.Tables("ManufOutSheetInfo").Rows(j).Item("Sts")
                        ManufOutFinishTime = CStr(DBDataSet5.Tables("WaitHandleInfo").Rows(0).Item("CompletedTime"))
                        MOSUrl = DBDataSet4.Tables("ManufOutSheetInfo").Rows(j).Item("ViewUrl")
                        Exit For
                    End If
                Next j
            End If

            '判斷外注建立及完成時間是否為空白,若不是的話,則到GetTime計算經過時間
            If ManufOutCreateTime <> "" And ManufOutFinishTime <> "" Then
                ManufOutModTime = GetTime(CDate(ManufOutCreateTime), CDate(ManufOutFinishTime))
                MapToOutET = GetTime(CDate(MapCreateTime), CDate(ManufOutFinishTime))
            End If

            '以製圖或圖面修改的圖號找出內製圖號是相同的RECORD
            SQL = SQL2 + "SELECT a.Unique_ID,a.No,a.Formno,a.Formsno,a.Sts,a.CreateTime,b.ViewUrl FROM F_ManufInSheet a "
            SQL = SQL + "LEFT OUTER JOIN V_WaitHandle_01 b ON a.FormNo=b.FormNo And a.FormSno=b.FormSno WHERE "
            SQL = SQL + "a.MapNo = '" + MapNo + "' "
            SQL = SQL + "GROUP BY a.Unique_ID,a.No,a.Formno,a.Formsno,a.Sts,a.CreateTime,b.ViewUrl "
            SQL = SQL + "ORDER BY a.Unique_ID DESC"

            DBDataSet6.Clear()
            Dim DBAdapter6 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter6.Fill(DBDataSet6, "ManufInSheetInfo")

            ManufInCreateTime = ""
            ManufInFinishTime = ""
            ManufInModTime = ""
            ManufInSts = ""
            MapToInET = ""
            ManufInModCount = 0
            MISUrl = ""

            '判斷是否有找到
            If DBDataSet6.Tables("ManufInSheetInfo").Rows.Count > 0 Then
                wStep = "160"
                ManufInModCount = DBDataSet6.Tables("ManufInSheetInfo").Rows.Count  '內製的修改次數
                ManufInCreateTime = DBDataSet6.Tables("ManufInSheetInfo").Rows(ManufInModCount - 1).Item("CreateTime")  '建立日期為第一筆的完成日期
                MISUrl = DBDataSet6.Tables("ManufInSheetInfo").Rows(0).Item("ViewUrl")
                ManufInSts = DBDataSet6.Tables("ManufInSheetInfo").Rows(0).Item("Sts")

                For j = 0 To ManufInModCount - 1
                    '篩選條件1.Formno & Formsno相同,2.內製委託書有160工程
                    SQL = "SELECT CompletedTime FROM T_WaitHandle WHERE FormNo = '"
                    SQL = SQL + DBDataSet6.Tables("ManufInSheetInfo").Rows(j).Item("Formno") + "' AND Formsno = '"
                    SQL = SQL + CStr(DBDataSet6.Tables("ManufInSheetInfo").Rows(j).Item("Formsno")) + "' AND STEP = '"
                    SQL = SQL + wStep + "' "
                    SQL = SQL + "AND NOT(CompletedTime IS NULL) "

                    DBDataSet7.Clear()
                    Dim DBAdapter7 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter7.Fill(DBDataSet7, "WaitHandleInfo")

                    '判斷是否有160工程
                    If DBDataSet7.Tables("WaitHandleInfo").Rows.Count >= 1 Then
                        ManufInSts = DBDataSet6.Tables("ManufInSheetInfo").Rows(j).Item("Sts")
                        ManufInFinishTime = CStr(DBDataSet7.Tables("WaitHandleInfo").Rows(0).Item("CompletedTime"))
                        MISUrl = DBDataSet6.Tables("ManufInSheetInfo").Rows(j).Item("ViewUrl")
                        Exit For
                    End If
                Next j
            End If

            '判斷內製建立及完成時間是否為空白,若不是的話,則到GetTime計算經過時間
            If ManufInCreateTime <> "" And ManufInFinishTime <> "" Then
                ManufInModTime = GetTime(CDate(ManufInCreateTime), CDate(ManufInFinishTime))
                MapToInET = GetTime(CDate(MapCreateTime), CDate(ManufInFinishTime))
            End If

            '建立Table
            SQL = "Insert into " & WorkTable & " "
            SQL = SQL + "( "
            SQL = SQL + "MNo, Buyer, MLevel, MapCreateTime, MapFinishTime, MapModTime, MapModCount, MapSts, "
            SQL = SQL + "ManufInCreateTime, MISUrl, ManufInFinishTime, ManufInModTime, ManufInModCount, ManufInSts, MapToInET, "
            SQL = SQL + "ManufOutCreateTime, MOSUrl, ManufOutFinishTime, ManufOutModTime, ManufOutModCount, ManufOutSts, MapToOutET, "
            SQL = SQL + "CreateUser, CreateTime "
            SQL = SQL + ") "
            SQL = SQL + "VALUES ("
            SQL = SQL + " '" + DBDataSet1.Tables("MapSheetInfo").Rows(i).Item("No") + "', "
            SQL = SQL + " '" + DBDataSet1.Tables("MapSheetInfo").Rows(i).Item("Buyer") + "', "
            SQL = SQL + " '" + DBDataSet1.Tables("MapSheetInfo").Rows(i).Item("Level") + "', "
            SQL = SQL + " '" + MapCreateTime + "', "
            SQL = SQL + " '" + MapFinishTime + "', "
            SQL = SQL + " '" + MapModTime + "', "
            SQL = SQL + " '" + CStr(MapModCount) + "', "
            SQL = SQL + " '" + MapSts + "', "
            SQL = SQL + " '" + ManufInCreateTime + "', "
            SQL = SQL + " '" + MISUrl + "', "
            SQL = SQL + " '" + ManufInFinishTime + "', "
            SQL = SQL + " '" + ManufInModTime + "', "
            SQL = SQL + " '" + CStr(ManufInModCount) + "', "
            SQL = SQL + " '" + ManufInSts + "', "
            SQL = SQL + " '" + MapToInET + "', "
            SQL = SQL + " '" + ManufOutCreateTime + "', "
            SQL = SQL + " '" + MOSUrl + "', "
            SQL = SQL + " '" + ManufOutFinishTime + "', "
            SQL = SQL + " '" + ManufOutModTime + "', "
            SQL = SQL + " '" + CStr(ManufOutModCount) + "', "
            SQL = SQL + " '" + ManufOutSts + "', "
            SQL = SQL + " '" + MapToOutET + "', "
            SQL = SQL + " '" + Request.Cookies("UserID").Value + "', "     '作成者
            SQL = SQL + " '" + NowDateTime + "' "       '作成時間
            SQL = SQL + ") "

            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDBCommand1.ExecuteNonQuery()

        Next i

        SQL = "SELECT MNo, Buyer, MLevel, MapCreateTime, MapFinishTime, MapModTime, MapModCount , CASE MapSts When '0' Then '開發中' When '1' Then '開發完成(OK)' When '2' Then '開發完成(NG)' When '3' Then '開發取消' Else '' End As MapSts, "
        SQL = SQL + "ManufInCreateTime, MISUrl, ManufInFinishTime, ManufInModTime, CASE ManufInModCount WHEN '0' THEN '' ELSE ManufInModCount END AS ManufInModCount, CASE ManufInSts When '0' Then '開發中' When '1' Then '開發完成(OK)' When '2' Then '開發完成(NG)' When '3' Then '開發取消' Else '' End As ManufInSts, MapToInET, "
        SQL = SQL + "ManufOutCreateTime, MOSUrl, ManufOutFinishTime, ManufOutModTime, CASE ManufOutModCount WHEN '0' THEN '' ELSE ManufOutModCount END AS ManufOutModCount, CASE ManufOutSts When '0' Then '開發中' When '1' Then '開發完成(OK)' When '2' Then '開發完成(NG)' When '3' Then '開發取消' Else '' End As ManufOutSts,  MapToOutET "
        SQL = SQL + "FROM "
        SQL = SQL + WorkTable
        SQL = SQL + " ORDER BY Buyer,MapCreateTime"

        Dim DBAdapter8 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter8.Fill(DBDataSet8, "DevelopmentTimeList")
        DBTable1 = DBDataSet8.Tables("DevelopmentTimeList")

        ''設定第一列及第二列表頭內容
        'DBDataRow = DBTable1.NewRow()
        'DBDataRow.Item(0) = "委託No"
        'DBDataRow.Item(1) = "Buyer"
        'DBDataRow.Item(2) = "繪圖階段"
        'DBDataRow.Item(3) = "內製委託書"
        'DBDataRow.Item(4) = "外注委託書"

        'DBTable1.Rows.InsertAt(DBDataRow, 0)    '插入在第一列

        'DBDataRow = DBTable1.NewRow()
        'DBDataRow.Item(0) = "委託No"
        'DBDataRow.Item(1) = "Buyer"
        'DBDataRow.Item(2) = "委託起始日"
        'DBDataRow.Item(3) = "最終完成日"
        'DBDataRow.Item(4) = "經過時間"
        'DBDataRow.Item(5) = "修改次數"
        'DBDataRow.Item(6) = "狀態"
        'DBDataRow.Item(7) = "委託起始日"
        'DBDataRow.Item(8) = "最終完成日"
        'DBDataRow.Item(9) = "經過時間"
        'DBDataRow.Item(10) = "開發次數"
        'DBDataRow.Item(11) = "狀態"
        'DBDataRow.Item(12) = "委託起始日"
        'DBDataRow.Item(13) = "最終完成日"
        'DBDataRow.Item(14) = "經過時間"
        'DBDataRow.Item(15) = "開發次數"
        'DBDataRow.Item(16) = "狀態"
        'DBTable1.Rows.InsertAt(DBDataRow, 0)    '插入在第二列

        DataGrid1.DataSource = DBTable1
        DataGrid1.DataBind()

        SetFieldWidth() '設定跨欄及欄寬

        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     經過時間計算
    '**
    '*****************************************************************
    Function GetTime(ByVal SDate As Date, ByVal EDate As Date) As Integer
        Dim i, TmpHour As Integer
        Dim TmpDate, TmpYMD As Date
        Dim SQL As String
        Dim SSeqNo, ESeqNo As Integer

        Dim DBDataSetA As New DataSet
        Dim DBDataSetB As New DataSet

        Dim OleDBCommand2 As New OleDbCommand
        Dim OleDbConnection2 As New OleDbConnection

        OleDbConnection2.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection2.Open()

        '共兩次,一為開始,一為結束
        For i = 1 To 2
            If i = 1 Then TmpDate = SDate Else TmpDate = EDate
            TmpYMD = TmpDate.Date
            TmpHour = CInt(TmpDate.Hour) * 60 + CInt(TmpDate.Minute)    '將時間轉換成分

            SQL = "SELECT MIN(SeqNo) AS SeqNo FROM M_Calendar WHERE YMD = '"
            SQL = SQL + TmpYMD + "' AND Hour >= '"

            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
            'SQL = SQL + CStr(TmpHour) + "' AND Active = '1' AND Depo = 'cl'"
            '
            '群組行事曆
            SQL = SQL + CStr(TmpHour) + "' AND Active = '1' AND Depo = 'CL1'"     '中壢行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
            'Modify-End

            DBDataSetA.Clear()
            Dim DBAdapterA As New OleDbDataAdapter(SQL, OleDbConnection2)
            DBAdapterA.Fill(DBDataSetA, "MapModTime")

            '判斷是否有找到資料
            If IsDBNull(DBDataSetA.Tables("MapModTime").Rows(0).Item("SeqNo")) <> True Then
                If i = 1 Then
                    SSeqNo = DBDataSetA.Tables("MapModTime").Rows(0).Item("SeqNo")
                Else
                    ESeqNo = DBDataSetA.Tables("MapModTime").Rows(0).Item("SeqNo")
                End If
            Else
                SQL = "SELECT MIN(SeqNo) AS SeqNo FROM M_Calendar WHERE YMD > '"

                'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                SQL = SQL + TmpYMD + "' AND Active = '1' AND Depo = 'cl'"
                '
                '群組行事曆
                SQL = SQL + TmpYMD + "' AND Active = '1' AND Depo = 'CL1'"     '中壢行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
                'Modify-End

                DBDataSetB.Clear()
                Dim DBAdapterB As New OleDbDataAdapter(SQL, OleDbConnection2)
                DBAdapterB.Fill(DBDataSetB, "MapModTime")
                If i = 1 Then
                    SSeqNo = DBDataSetB.Tables("MapModTime").Rows(0).Item("SeqNo")
                Else
                    ESeqNo = DBDataSetB.Tables("MapModTime").Rows(0).Item("SeqNo")
                End If
            End If
        Next i
        OleDbConnection2.Close()
        GetTime = (ESeqNo - SSeqNo) * 10 / 60    '計算經過時間(小時)
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定跨欄及欄寬
    '**
    '*****************************************************************
    Sub SetFieldWidth()
        Dim i As Integer
        Dim TempWidth As Integer

        'DataGrid1.Items(0).Cells(0).RowSpan = 2     '將第一列第一欄顯示為直向跨欄
        'DataGrid1.Items(0).Cells(1).RowSpan = 2     '將第一列第一欄顯示為直向跨欄
        'DataGrid1.Items(0).Cells(2).ColumnSpan = 5  '將第一列第二欄顯示為橫向跨欄
        'DataGrid1.Items(0).Cells(3).ColumnSpan = 5  '將第一列第三欄顯示為橫向跨欄
        'DataGrid1.Items(0).Cells(4).ColumnSpan = 5  '將第一列第四欄顯示為橫向跨欄

        'For i = 5 To DataGrid1.Columns.Count - 1
        '    DataGrid1.Items(0).Cells(i).Visible = False '隱藏橫向跨欄後面12欄
        'Next i

        'DataGrid1.Items(1).Cells(0).Visible = False '隱藏直向跨欄後面1欄
        'DataGrid1.Items(1).Cells(1).Visible = False '隱藏直向跨欄後面1欄

        ''設定第一列第二、三、四欄寬度為
        'DataGrid1.Items(0).Cells(2).Width = New Unit(240)
        'DataGrid1.Items(0).Cells(3).Width = New Unit(240)
        'DataGrid1.Items(0).Cells(4).Width = New Unit(240)

        '將第一、二列所有欄位字設定為置中
        'For i = 0 To DataGrid1.Columns.Count - 1
        '    DataGrid1.Items(0).Cells(i).HorizontalAlign = HorizontalAlign.Center
        'DataGrid1.Items(1).Cells(i).HorizontalAlign = HorizontalAlign.Center
        'Next i

        '設定第一、二列所有欄位背景及字體顏色
        'DataGrid1.Items(0).BackColor = Color.FromArgb(13395456)
        'DataGrid1.Items(0).ForeColor = System.Drawing.Color.White
        'DataGrid1.Items(1).BackColor = Color.FromArgb(13395456)
        'DataGrid1.Items(1).ForeColor = System.Drawing.Color.White



        For i = 0 To 19
            Select Case i
                Case 0
                    TempWidth = 80
                Case 1
                    TempWidth = 80
                Case 2
                    TempWidth = 40
                Case 3
                    TempWidth = 80
                Case 4
                    TempWidth = 80
                Case 5
                    TempWidth = 30
                Case 6
                    TempWidth = 20
                Case 7
                    TempWidth = 30
                Case 8
                    TempWidth = 80
                Case 9
                    TempWidth = 80
                Case 10
                    TempWidth = 30
                Case 11
                    TempWidth = 20
                Case 12
                    TempWidth = 30
                Case 13
                    TempWidth = 30
                Case 14
                    TempWidth = 80
                Case 15
                    TempWidth = 80
                Case 16
                    TempWidth = 30
                Case 17
                    TempWidth = 20
                Case 18
                    TempWidth = 30
                Case 19
                    TempWidth = 30
            End Select
            DataGrid1.Columns.Item(i).ItemStyle.Width = New Unit(TempWidth)
            DataGrid1.Columns.Item(i).HeaderStyle.Width = New Unit(TempWidth)
            DataGrid1.Width = New Unit(1161)
            'width = 80+80+40+80+80+30+20+40+80+80+30+20+40+30+80+80+30+20+40+30=950,格線=21(20+1),預留空間4*2(兩邊)*20=160,total=1161
        Next
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '程式別不同

        ''pFormNo = DFormName.SelectedValue
        DataGrid1.AllowPaging = False   '程式別不同
        CreateWorkTable()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=Commission_List.xls")     '程式別不同
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        DataGrid1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        DataGrid1.AllowPaging = wAllowPaging        '程式別不同
    End Sub

    Private Sub DSYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DSYear.SelectedIndexChanged
        SetBuyer()
    End Sub

    Private Sub DSMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DSMonth.SelectedIndexChanged
        SetBuyer()
    End Sub

    Private Sub DEYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DEYear.SelectedIndexChanged
        SetBuyer()
    End Sub

    Private Sub DEMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DEMonth.SelectedIndexChanged
        SetBuyer()
    End Sub
End Class
