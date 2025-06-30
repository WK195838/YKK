Imports System.Data
Imports System.Data.OleDb

Public Class WinPop
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents TextBox1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents TextBox2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Button2 As System.Web.UI.WebControls.Button
    Protected WithEvents TextBox3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Button3 As System.Web.UI.WebControls.Button
    Protected WithEvents Button4 As System.Web.UI.WebControls.Button
    Protected WithEvents Button5 As System.Web.UI.WebControls.Button
    Protected WithEvents Button6 As System.Web.UI.WebControls.Button
    Protected WithEvents Button7 As System.Web.UI.WebControls.Button
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DStatus As System.Web.UI.WebControls.TextBox
    Protected WithEvents LOPContact As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LSliderDetail As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LContact As System.Web.UI.WebControls.HyperLink
    Protected WithEvents Hyperlink1 As System.Web.UI.WebControls.HyperLink

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
        If Not Me.IsPostBack Then   '不是PostBack
            Button1.Attributes("onclick") = "aaa();"
            SetOPDataList()
        End If
    End Sub

    Sub SetOPDataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        SQL = "SELECT "
        SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
        SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, "
        SQL = SQL + "'預定開始：[' + Convert(VarChar, BStartTime, 20) + '], ' + "
        SQL = SQL + "'預定完成：[' + Convert(VarChar, BEndTime, 20) + '], ' + "
        SQL = SQL + "'實際開始：[' + Convert(VarChar, AStartTime, 20) + '], ' + "
        SQL = SQL + "'實際完成：[' + AEndTimeDesc + '], ' + "
        SQL = SQL + "'工程說明：[' + DecideDescA +  '], ' + "
        SQL = SQL + "'延遲原因：[' + ReasonA + '], ' + "
        SQL = SQL + "'延遲說明：[' + ReasonDescA + ']' "
        SQL = SQL + " As Description, URL "
        SQL = SQL + "FROM V_WaitHandle_01 "
        SQL = SQL + "Where FormNo  = '000002' "
        SQL = SQL + "  And FormSno = '6' "
        SQL = SQL + "Order by FormNo, FormSno, Step, SeqNo "

        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "WaitHandle")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()
        OleDbConnection1.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim oPutSubFile As Object
        Dim RtnCode

        oPutSubFile = Server.CreateObject("MSubFile.SubFile")
        RtnCode = oPutSubFile.PutSlider(TextBox2.Text, "aaa", "000003", 1, "Joy")

        If RtnCode = 1 Then
            TextBox3.Text = "Error"
        Else
            TextBox3.Text = "OK"
        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim oPutSubFile As Object
        Dim RtnCode

        oPutSubFile = Server.CreateObject("MSubFile.SubFile")
        RtnCode = oPutSubFile.PutSample(TextBox2.Text, "000003", 1, "Joy")

        If RtnCode = 1 Then
            TextBox3.Text = "Error"
        Else
            TextBox3.Text = "OK"
        End If

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim oPutSubFile As Object
        Dim RtnCode

        oPutSubFile = Server.CreateObject("MSubFile.SubFile")
        RtnCode = oPutSubFile.PutPrice(TextBox2.Text, "000003", 1, "Joy")

        If RtnCode = 1 Then
            TextBox3.Text = "Error"
        Else
            TextBox3.Text = "OK"
        End If


    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim oPutSubFile As Object
        Dim RtnCode

        oPutSubFile = Server.CreateObject("MSubFile.SubFile")
        RtnCode = oPutSubFile.PutQA1(TextBox2.Text, "000003", 1, "Joy")

        If RtnCode = 1 Then
            TextBox3.Text = "Error"
        Else
            TextBox3.Text = "OK"
        End If

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim oPutSubFile As Object
        Dim RtnCode

        oPutSubFile = Server.CreateObject("MSubFile.SubFile")
        RtnCode = oPutSubFile.PutQA2(TextBox2.Text, "000003", 1, "Joy")

        If RtnCode = 1 Then
            TextBox3.Text = "Error"
        Else
            TextBox3.Text = "OK"
        End If

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        Dim scriptString As String = "<script language=JavaScript> "
        scriptString = scriptString & "aaa(); "
        scriptString = scriptString & "</script>"

        If (Not Me.IsStartupScriptRegistered("Startup")) Then
            Me.RegisterStartupScript("Startup", scriptString)
        End If

    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex   'DataGrid跳上下頁
        'DataGrid1.DataBind()
        SetOPDataList()

    End Sub
End Class
