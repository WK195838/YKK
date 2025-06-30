Imports System.Data
Imports System.Data.OleDb

Partial Class NewReport
    Inherits System.Web.UI.Page


    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_Class   'YKK SPD系共通涵式

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime, xNowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期
    Dim wID As String               'Unique_ID
    Dim wUserID As String           'User ID
    Dim wUserName As String         'User Name
    Dim wDivision As String         'Division
    Dim wFun As String              '新增(ADD)/新增(ADD1)-整體行程/修改(MOD)
    Dim wInqKey As String           '搜尋Key值
    Dim wPageIdx As String          'GridView Page Index
    Dim wNowYear As String          '全體行程之閱讀年
    Dim wWeeks As Integer           '全體行程之閱讀週

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        SetCommonParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            DisplayFieldFormat()
            If wFun = "MOD" Or wFun = "VIEW" Then
                ShowData()          '顯示資料(修改行程)
            Else
                SetFormatData()     '設定初值(新行程)
            End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetCommonParameter()
        '基本資料設定
        Response.Cookies("PGM").Value = "NewReport.aspx"
        xNowDateTime = Now.ToString("yyyyMMdd HHmmss")   '現在日期時間
        NowDateTime = CStr(DateTime.Now.Date) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)       '現在日時
        NowDate = Now.ToString("yyyy/MM/dd")          '現在日
        '附帶參數
        wUserID = Request.QueryString("pUserID")      'User ID
        wFun = Request.QueryString("pFun")            '新增/變更
        wID = Request.QueryString("pID")              'Unique_ID
        wInqKey = Request.QueryString("pInqKey")      '搜尋Key
        wPageIdx = Request.QueryString("pPageIdx")    'GridView Page Index
        wNowYear = Request.QueryString("pReportYear")  '全體行程之閱讀年
        wWeeks = CInt(Request.QueryString("pWeeks"))  '全體行程之閱讀週
        '取得業務員資訊
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As New DataTable
        Dim SQL As String

        Dim OleDbConnection1 As New OleDbConnection
        Dim OleDBCommand1 As New OleDbCommand
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn1")  'SQL連結設定

        SQL = "Select Name, Division From V_Member_01 Where UserID = '" + wUserID + "' "
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Member")
        DBTable1 = DBDataSet1.Tables("Member")
        If DBTable1.Rows.Count > 0 Then
            wUserName = DBTable1.Rows(0).Item("Name")
            wDivision = DBTable1.Rows(0).Item("Division")
        End If

        If wFun = "MOD" Then
            DBDataSet1.Clear()
            SQL = "Select * From WeekReport "
            SQL = SQL & "Where CreateUser = '" & wUserID & "' "
            SQL = SQL & "And Unique_ID = " & wID & " "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "WeekReport")
            DBTable1 = DBDataSet1.Tables("WeekReport")
            If DBTable1.Rows.Count <= 0 Then
                wFun = "VIEW"
            End If
        End If
        OleDbConnection1.Close()
        '設定表單欄位屬性
        SetFormFieldAttribute()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定表單欄位屬性
    '**
    '*****************************************************************
    Sub SetFormFieldAttribute()
        '按鈕設定
        BSave.Attributes("onclick") = "var ok=window.confirm('" + "是否儲存？" + "');if(!ok){return false;} else {}"
        BDel.Attributes("onclick") = "var ok=window.confirm('" + "是否刪除？" + "');if(!ok){return false;} else {}"

        If wFun = "VIEW" Then
            BSave.Visible = False
            BDel.Visible = False
            BReportDate.Enabled = False       '日期選擇
        Else
            BSave.Visible = True
            BDel.Visible = True
            BReportDate.Enabled = True        '日期選擇
        End If
        BReportDate.Attributes("onclick") = "CalendarPicker('Form1.DReportDate');"      '日期選擇
        'ReadOnly欄位屬性設定
        DReportDate.Attributes("readonly") = "readonly"
        DName.Attributes("readonly") = "readonly"
        DDivision.Attributes("readonly") = "readonly"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     顯示欄位初值處理
    '**
    '*****************************************************************
    Sub DisplayFieldFormat()
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn1")  'SQL連結設定
        '
        DReportDate.Text = ""
        DID.Text = ""
        DDivision.Text = ""
        DName.Text = ""
        DContent.Text = ""
        DRemark.Text = ""
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定初值(新行程)
    '**
    '*****************************************************************
    Sub SetFormatData()
        DReportDate.Text = NowDate
        DName.Text = wUserName
        DDivision.Text = wDivision
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     顯示資料(修改行程)
    '**
    '*****************************************************************
    Sub ShowData()
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As New DataTable
        Dim SQL As String

        Dim OleDbConnection1 As New OleDbConnection
        Dim OleDBCommand1 As New OleDbCommand
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn1")  'SQL連結設定

        SQL = "Select * From WeekReport Where Unique_ID = '" + wID + "' "
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "WeekReport")
        DBTable1 = DBDataSet1.Tables("WeekReport")
        If DBTable1.Rows.Count > 0 Then
            '行程日期
            DReportDate.Text = Format(DBTable1.Rows(0).Item("ReportDate"), "Short Date")
            'ID
            DID.Text = DBTable1.Rows(0).Item("Unique_ID").ToString
            'Name
            DName.Text = DBTable1.Rows(0).Item("Name").ToString
            'Division
            DDivision.Text = DBTable1.Rows(0).Item("Division").ToString
            '
            DSection.Text = DBTable1.Rows(0).Item("Section").ToString
            '
            DType.Text = DBTable1.Rows(0).Item("WorkType").ToString
            '
            DContent.Text = DBTable1.Rows(0).Item("WorkContent").ToString
            '
            DRemark.Text = DBTable1.Rows(0).Item("Remark").ToString
        End If
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     儲存
    '**
    '*****************************************************************
    Protected Sub BSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSave.Click
        Dim SQL As String
        Dim ErrCode As Integer
        Dim wHeadDivision As String = ""
        Dim wSection As String = ""
        Dim wSalesCode As String = ""
        Dim wEmpID As String = ""
        Dim wVisitDate As String = ""   '前拜訪日期

        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As New DataTable
        Dim OleDbConnection1 As New OleDbConnection
        Dim OleDBCommand1 As New OleDbCommand
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn1")  'SQL連結設定

        '帳號遺失是否
        If ErrCode = 0 Then
            If wUserID = "" Then ErrCode = 9000
        End If
        '
        '異常是否
        If ErrCode = 0 Then
            BSave.Enabled = False
            BDel.Visible = False

            '取得業務員相關資料
            OleDbConnection1.Open()
            SQL = "Select HeadDivision, Section, SalesCode, EmpID From V_Member_01 Where UserID = '" + wUserID + "' "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "Member")
            DBTable1 = DBDataSet1.Tables("Member")
            If DBTable1.Rows.Count > 0 Then
                wHeadDivision = DBTable1.Rows(0).Item("HeadDivision")
                wSection = DBTable1.Rows(0).Item("Section")
                wSalesCode = DBTable1.Rows(0).Item("SalesCode")
                wEmpID = DBTable1.Rows(0).Item("EmpID")
            End If
            '新增 or 修改 WeekReport
            If wFun = "ADD" Then
                '***新增作業
                SQL = "Insert into WeekReport "
                SQL = SQL + "( "
                SQL = SQL + "Sts, ReportDate, "
                SQL = SQL + "HeadDivision, Division, Section, Name, "
                SQL = SQL + "WorkType, WorkContent, Remark, "
                SQL = SQL + "CreateUser, CreateTime, ModifyUser, ModifyTime "              '21~24
                SQL = SQL + ")  "
                SQL = SQL + "VALUES( "
                SQL = SQL + " '" + "1" + "', "                                         '啟用
                SQL = SQL + " '" + CDate(DReportDate.Text) + "', "                     '行程日
                SQL = SQL + " N'" + YKK.ReplaceString(wHeadDivision) + "', "           '事業部
                SQL = SQL + " N'" + YKK.ReplaceString(DDivision.Text) + "', "          '部門
                SQL = SQL + " N'" + YKK.ReplaceString(DSection.Text) + "', "           '課
                SQL = SQL + " N'" + YKK.ReplaceString(DName.Text) + "', "              'member
                SQL = SQL + " N'" + YKK.ReplaceString(DType.Text) + "', "
                SQL = SQL + " N'" + YKK.ReplaceString(DContent.Text) + "', "
                SQL = SQL + " N'" + YKK.ReplaceString(DRemark.Text) + "', "
                SQL = SQL + " '" + wUserID + "', "                  '作成者ID
                SQL = SQL + " '" + NowDateTime + "', "              '作成時間
                SQL = SQL + " '" + "" + "', "                       '修改者
                SQL = SQL + " '" + NowDateTime + "' "               '修改時間
                SQL = SQL + " ) "
                OleDBCommand1.Connection = OleDbConnection1
                OleDBCommand1.CommandText = SQL
                OleDBCommand1.ExecuteNonQuery()
            Else
                '更新變更後訪問新行程
                SQL = "Update WeekReport Set "
                SQL = SQL + " ReportDate = '" + CDate(DReportDate.Text) + "',"
                SQL = SQL + " Division = N'" + YKK.ReplaceString(DDivision.Text) + "',"
                SQL = SQL + " Section = N'" + YKK.ReplaceString(DSection.Text) + "',"
                SQL = SQL + " WorkType = N'" + YKK.ReplaceString(DType.Text) + "',"
                SQL = SQL + " WorkContent = N'" + YKK.ReplaceString(DContent.Text) + "',"
                SQL = SQL + " Remark = N'" + YKK.ReplaceString(DRemark.Text) + "',"
                SQL = SQL + " ModifyUser = '" + wUserID + "',"
                SQL = SQL + " ModifyTime = '" + NowDateTime + "' "
                SQL = SQL + " Where Unique_ID =  '" + DID.Text + "'"
                OleDBCommand1.Connection = OleDbConnection1
                OleDBCommand1.CommandText = SQL
                OleDBCommand1.ExecuteNonQuery()
            End If
            BSave.Enabled = True   'Enable 儲存Button
            BDel.Visible = True
            OleDbConnection1.Close()
        Else   '異常
            If ErrCode = 9000 Then ShowMessage("登錄者資料遺失")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     刪除
    '**
    '*****************************************************************
    Protected Sub BDel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BDel.Click
        Dim SQL As String
        Dim OleDbConnection1 As New OleDbConnection
        Dim OleDBCommand1 As New OleDbCommand
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn1")  'SQL連結設定
        OleDbConnection1.Open()
        '
        SQL = "Delete  From WeekReport "
        SQL = SQL + " Where Unique_ID =  '" + DID.Text + "'"
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = Sql
        OleDBCommand1.ExecuteNonQuery()
        '
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     彈出警告訊息
    '**
    '*****************************************************************
    Sub ShowMessage(ByVal Msg As String)
        Dim myMsg As New Literal()
        myMsg.Text = "<script>alert('" & Msg & "')</script><br>"
        Me.Page.Controls.Add(myMsg)
    End Sub
End Class
