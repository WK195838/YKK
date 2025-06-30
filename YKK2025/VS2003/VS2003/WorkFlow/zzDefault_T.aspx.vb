Imports System.Random
Imports System.Data
Imports System.Data.OleDb

Public Class Default_T
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label
    Protected WithEvents DFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DStep As System.Web.UI.WebControls.TextBox
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents RequiredFieldValidator1 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents DSeqNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label4 As System.Web.UI.WebControls.Label
    Protected WithEvents Button2 As System.Web.UI.WebControls.Button
    Protected WithEvents Label5 As System.Web.UI.WebControls.Label
    Protected WithEvents Label6 As System.Web.UI.WebControls.Label
    Protected WithEvents DUserID As System.Web.UI.WebControls.TextBox
    Protected WithEvents DApplyID As System.Web.UI.WebControls.TextBox

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    Public YKK As New YKK_SPDClass   'YKK共通

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not Me.IsPostBack Then   '不是PostBack
        End If

        '在這裡放置使用者程式碼以初始化網頁
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Response.Cookies("Cookies_UserID").Value = DUserID.Text
        'Response.Cookies("Cookies_ApplyID").Value = DApplyID.Text

        Dim LinkAdr As String
        LinkAdr = "MapSheet_01.aspx?pFormNo=" & DFormNo.Text & "&pFormSno=" & DFormSno.Text & "&pStep=" & DStep.Text & "&pSeqNo=" & DSeqNo.Text & "&pApplyID=" & DApplyID.Text
        Response.Redirect(LinkAdr)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = "Provider=SQLOLEDB.1;" & _
                                            "User ID=" & "sa" & ";" & _
                                            "Data Source='" & "10.245.0.112" & "';" & _
                                            "Initial Catalog=" & "SPD"
        Dim OleDBCommand1 As New OleDbCommand

        Dim NowDateTime As String = CStr(DateTime.Now.Today) + " " + CStr(DateTime.Now.Hour) + ":" + CStr(DateTime.Now.Minute) + ":" + CStr(DateTime.Now.Second)


        Dim SQL As String
        SQL = "Insert into aaa (CompletedTime) Values( '" & NowDateTime & "' ) "
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQL
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()
    End Sub
End Class
