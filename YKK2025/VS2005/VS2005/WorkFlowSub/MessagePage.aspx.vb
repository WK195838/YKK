Imports System.Data
Imports System.Data.OleDb

Partial Class MessagePage
    Inherits System.Web.UI.Page


    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        FocusBox.Focus()
        Message.Text = ""

        If Request.QueryString("pMSGID") = "C" Then CompletedMessage()
        If Request.QueryString("pMSGID") = "S" Then SavedMessage()
        If Request.QueryString("pMSGID") = "N" Then NGMessage()

        If Not Me.IsPostBack Then
            WaitHandle_DataList()
        End If

    End Sub

    Sub WaitHandle_DataList()
        'Dim i As Integer = 0
        'Dim SQL As String
        'Dim DBDataSet1 As New DataSet
        'Dim OleDbConnection1 As New OleDbConnection
        Dim Http As String = System.Configuration.ConfigurationManager.AppSettings("Http")  'Http Address
        'OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定


        'OleDbConnection1.Open()
        ''一般件
        'SQL = "SELECT "
        'SQL = SQL + "Count(*) As Low "
        'SQL = SQL + "FROM V_WaitHandle_01 "
        'SQL = SQL + "Where Active = '1' "
        'SQL = SQL + "  And (Sts = '0'  Or  Sts = '4') "
        'SQL = SQL + "  And (Delay = '0' or Delay='1') "
        'SQL = SQL + "  And DecideID = '" + Request.QueryString("pUserID") + "'"
        'Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter2.Fill(DBDataSet1, "WaitHandle")
        Dim Meta As New HtmlMeta
        Meta.Attributes.Add("http-equiv", "refresh")
        'If DBDataSet1.Tables("WaitHandle").Rows(0).Item("Low") > 0 Then

        Meta.Content = "5; url=" + Http + "/WorkFlow/WaitHandle.aspx?pUserID=" + Request.QueryString("pUserID")



        'Else
        'Meta.Content = "5; url=" + Http + "/Portal/"

        'End If
        Header.Controls.Add(Meta)
        ' OleDbConnection1.Close()

    End Sub

    Sub CompletedMessage()
        Dim NowDateTime = CStr(DateTime.Now.Date) + " " + _
                          CStr(DateTime.Now.Hour) + ":" + _
                          CStr(DateTime.Now.Minute) + ":" + _
                          CStr(DateTime.Now.Second)     '現在日時

        Dim i As Integer = 0
        Dim SQL As String
        Dim wFormName, wStepName, wUserName, wApplyName As String

        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        '表單名
        SQL = "SELECT FormName FROM M_FORM "
        SQL = SQL + " Where FormNo = '" + Request.QueryString("pFormNo") + "'"
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_FORM")
        DBTable1 = DBDataSet1.Tables("M_FORM")
        For i = 0 To DBTable1.Rows.Count - 1
            wFormName = DBTable1.Rows(i).Item("FormName")
        Next
        OleDbConnection1.Close()

        '工程名
        SQL = "SELECT StepName FROM M_FLOW "
        SQL = SQL + " Where FormNo = '" + Request.QueryString("pFormNo") + "'"
        SQL = SQL + "   And Step   = '" + Request.QueryString("pStep") + "'"
        SQL = SQL + "   And Active = '1' "
        SQL = SQL + "   And Action = '0' "
        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_FLOW")
        DBTable1 = DBDataSet1.Tables("M_FLOW")
        For i = 0 To DBTable1.Rows.Count - 1
            wStepName = DBTable1.Rows(i).Item("StepName")
        Next
        OleDbConnection1.Close()

        '簽核者
        SQL = "SELECT UserName FROM M_Users "
        SQL = SQL + " Where UserID = '" + Request.QueryString("pUserID") + "'"
        OleDbConnection1.Open()
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            wUserName = DBTable1.Rows(i).Item("UserName")
        Next
        OleDbConnection1.Close()

        '委託者
        SQL = "SELECT UserName FROM M_Users "
        SQL = SQL + " Where UserID = '" + Request.QueryString("pApplyID") + "'"
        OleDbConnection1.Open()
        Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter4.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            wApplyName = DBTable1.Rows(i).Item("UserName")
        Next
        OleDbConnection1.Close()

        '下一關核定人
        Dim pName As String = oCommon.GetUserName(Request.QueryString("pNextGate"))

        Message.Text = "[ " & wUserName & " ]" & "於" & NowDateTime & _
                       "將" & "[ " & wApplyName & " ]" & "之" & "[ " & wFormName & " ]" & _
                       "成功送至" & "[ " & wStepName & "(" & pName & ") ]．"
    End Sub

    Sub SavedMessage()
        Dim NowDateTime = CStr(DateTime.Now.Date) + " " + _
                          CStr(DateTime.Now.Hour) + ":" + _
                          CStr(DateTime.Now.Minute) + ":" + _
                          CStr(DateTime.Now.Second)     '現在日時
        Dim i As Integer = 0
        Dim SQL As String
        Dim wFormName, wStepName, wUserName, wApplyName As String

        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        '表單名
        SQL = "SELECT FormName FROM M_FORM "
        SQL = SQL + " Where FormNo = '" + Request.QueryString("pFormNo") + "'"
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_FORM")
        DBTable1 = DBDataSet1.Tables("M_FORM")
        For i = 0 To DBTable1.Rows.Count - 1
            wFormName = DBTable1.Rows(i).Item("FormName")
        Next
        OleDbConnection1.Close()

        '工程名
        SQL = "SELECT StepName FROM M_FLOW "
        SQL = SQL + " Where FormNo = '" + Request.QueryString("pFormNo") + "'"
        SQL = SQL + "   And Step   = '" + Request.QueryString("pStep") + "'"
        SQL = SQL + "   And Active = '1' "
        SQL = SQL + "   And Action = '0' "
        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_FLOW")
        DBTable1 = DBDataSet1.Tables("M_FLOW")
        For i = 0 To DBTable1.Rows.Count - 1
            wStepName = DBTable1.Rows(i).Item("StepName")
        Next
        OleDbConnection1.Close()

        '簽核者
        SQL = "SELECT UserName FROM M_Users "
        SQL = SQL + " Where UserID = '" + Request.QueryString("pUserID") + "'"
        OleDbConnection1.Open()
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            wUserName = DBTable1.Rows(i).Item("UserName")
        Next
        OleDbConnection1.Close()

        '委託者
        SQL = "SELECT UserName FROM M_Users "
        SQL = SQL + " Where UserID = '" + Request.QueryString("pApplyID") + "'"
        OleDbConnection1.Open()
        Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter4.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            wApplyName = DBTable1.Rows(i).Item("UserName")
        Next
        OleDbConnection1.Close()

        Message.Text = "[ " & wStepName & ":" & wUserName & " ]" & "於" & NowDateTime & _
                       "將" & "[ " & wApplyName & " ]" & "之" & "[ " & wFormName & " ]" & "資料儲存成功．"
    End Sub

    Sub NGMessage()
        Dim NowDateTime = CStr(DateTime.Now.Date) + " " + _
                          CStr(DateTime.Now.Hour) + ":" + _
                          CStr(DateTime.Now.Minute) + ":" + _
                          CStr(DateTime.Now.Second)     '現在日時
        Dim i As Integer = 0
        Dim SQL As String
        Dim wFormName, wStepName, wUserName, wApplyName As String

        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        '表單名
        SQL = "SELECT FormName FROM M_FORM "
        SQL = SQL + " Where FormNo = '" + Request.QueryString("pFormNo") + "'"
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_FORM")
        DBTable1 = DBDataSet1.Tables("M_FORM")
        For i = 0 To DBTable1.Rows.Count - 1
            wFormName = DBTable1.Rows(i).Item("FormName")
        Next
        OleDbConnection1.Close()

        '工程名
        SQL = "SELECT StepName FROM M_FLOW "
        SQL = SQL + " Where FormNo = '" + Request.QueryString("pFormNo") + "'"
        SQL = SQL + "   And Step   = '" + Request.QueryString("pStep") + "'"
        SQL = SQL + "   And Active = '1' "
        SQL = SQL + "   And Action = '1' "
        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_FLOW")
        DBTable1 = DBDataSet1.Tables("M_FLOW")
        For i = 0 To DBTable1.Rows.Count - 1
            wStepName = DBTable1.Rows(i).Item("StepName")
        Next
        OleDbConnection1.Close()

        '簽核者
        SQL = "SELECT UserName FROM M_Users "
        SQL = SQL + " Where UserID = '" + Request.QueryString("pUserID") + "'"
        OleDbConnection1.Open()
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            wUserName = DBTable1.Rows(i).Item("UserName")
        Next
        OleDbConnection1.Close()

        '委託者
        SQL = "SELECT UserName FROM M_Users "
        SQL = SQL + " Where UserID = '" + Request.QueryString("pApplyID") + "'"
        OleDbConnection1.Open()
        Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter4.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            wApplyName = DBTable1.Rows(i).Item("UserName")
        Next
        OleDbConnection1.Close()

        Message.Text = "[ " & wStepName & ":" & wUserName & " ]" & "於" & NowDateTime & _
                       "將" & "[ " & wApplyName & " ]" & "之" & "[ " & wFormName & " ]" & "資料NG成功．"
    End Sub


End Class
