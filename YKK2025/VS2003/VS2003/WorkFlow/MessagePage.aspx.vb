Imports System.Data
Imports System.Data.OleDb

Public Class CompletedPage
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Message As System.Web.UI.WebControls.TextBox

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim URL As String = "http://10.245.1.6/WorkFlowSub/MessagePage.aspx?" + _
                            "pMSGID=" + Request.QueryString("pMSGID") + "&" + _
                            "pFormNo=" + Request.QueryString("pFormNo") + "&" + _
                            "pStep=" + Request.QueryString("pStep") + "&" + _
                            "pUserID=" + Request.QueryString("pUserID") + "&" + _
                            "pApplyID=" + Request.QueryString("pApplyID")

        Response.Redirect(URL)
        '------------------------------------------------------------------
        Message.Text = ""

        If Request.QueryString("pMSGID") = "C" Then CompletedMessage()
        If Request.QueryString("pMSGID") = "S" Then SavedMessage()
        If Request.QueryString("pMSGID") = "N" Then NGMessage()

    End Sub

    Sub CompletedMessage()
        Dim NowDateTime As String = CStr(DateTime.Now.Today) + " " + _
                                    CStr(DateTime.Now.Hour) + ":" + _
                                    CStr(DateTime.Now.Minute) + ":" + _
                                    CStr(DateTime.Now.Second)           '現在日時
        Dim i As Integer = 0
        Dim SQL As String
        Dim wFormName, wStepName, wUserName, wApplyName As String

        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

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
        Dim oGetUserName As Object
        Dim pUser As String = Request.QueryString("pNextGate")
        Dim pName As String = ""
        Dim RtnCode As Integer = 0

        oGetUserName = Server.CreateObject("GetUserName.WFGetUserName")
        RtnCode = oGetUserName.UserName(pUser, pName)

        Message.Text = "[ " & wUserName & " ]" & "於" & NowDateTime & _
                       "將" & "[ " & wApplyName & " ]" & "之" & "[ " & wFormName & " ]" & _
                       "成功送至" & "[ " & wStepName & "(" & pName & ") ]．"
    End Sub

    Sub SavedMessage()
        Dim NowDateTime As String = CStr(DateTime.Now.Today) + " " + _
                                    CStr(DateTime.Now.Hour) + ":" + _
                                    CStr(DateTime.Now.Minute) + ":" + _
                                    CStr(DateTime.Now.Second)           '現在日時
        Dim i As Integer = 0
        Dim SQL As String
        Dim wFormName, wStepName, wUserName, wApplyName As String

        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

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
        Dim NowDateTime As String = CStr(DateTime.Now.Today) + " " + _
                                    CStr(DateTime.Now.Hour) + ":" + _
                                    CStr(DateTime.Now.Minute) + ":" + _
                                    CStr(DateTime.Now.Second)           '現在日時
        Dim i As Integer = 0
        Dim SQL As String
        Dim wFormName, wStepName, wUserName, wApplyName As String

        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

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
