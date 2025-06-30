Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class CalendarList_TP2
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents cal As System.Web.UI.WebControls.Calendar

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
        Response.Cookies("PGM").Value = "CalendarList_TP2.aspx"
        If Not Me.IsPostBack Then
        End If
    End Sub

    Sub Cal_DayRender(ByVal sender As Object, ByVal e As DayRenderEventArgs)
        Dim SQL As String
        Dim objDb As New SqlConnection(ConfigurationSettings.AppSettings.Get("SqlConn1"))    'SQL連結設定

        Dim objCom As New SqlCommand("SELECT sid,title,sdate,stime,etime FROM V_Vacation_01 WHERE sdate=@sdate and Depo=@Depo ORDER BY stime ASC", objDb)
        objCom.Parameters.Add("@sdate", e.Day.Date)
        objCom.Parameters.Add("@Depo", "TP2")
        objDb.Open()
        Dim objDr As SqlDataReader = objCom.ExecuteReader()

        Do While objDr.Read()
            e.Cell.Controls.Add(New LiteralControl("<br/><font color='#FF0000'>" & String.Format("{2}", objDr.GetString(3), objDr.GetString(4), objDr.GetString(1)) & "</font>"))
        Loop
        objDb.Close()
    End Sub


End Class
