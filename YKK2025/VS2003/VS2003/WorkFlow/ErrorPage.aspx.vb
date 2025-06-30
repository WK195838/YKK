Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ErrorPage
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Message As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label

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
        Write_WaitSend()
    End Sub

    Sub Write_WaitSend()
        Dim NowDateTime As String       '現在日期時間
        Dim SQl As String
        Dim Str, Str1, Str2, Str3, Str4 As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '現在日時

        Str = "網頁發生系統異常"
        Str1 = "ID= (" + Request.Cookies("UserID").Value + ")"
        Str2 = "PGM= (" + Request.Cookies("PGM").Value + ")"
        Str3 = ""
        Str4 = ""
        If Not Request.Cookies("PGMFORMSNO") Is Nothing Then
            Str3 = "FORMSNO= (" + Request.Cookies("PGMFORMSNO").Value + ")"
        End If
        If Not Request.Cookies("PGMSTEP") Is Nothing Then
            Str4 = "STEP= (" + Request.Cookies("PGMSTEP").Value + ")"
        End If

        SQl = "Insert into Q_WaitSend "
        SQl = SQl + "( "
        SQl = SQl + "Sts, FromID, FromMail, FromName, ToID, "      '1~5
        SQl = SQl + "ToMail, ToName, CCMail, "                     '6~8
        SQl = SQl + "FormNo, FormSno, FormName, Step, StepName, "  '9~13
        SQl = SQl + "ApplyID, ApplyName, MSG, MSGName, "           '14~17
        SQl = SQl + "CreateTime "                                  '18
        SQl = SQl + ")  "
        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                'Sts
        SQl = SQl + " 'Admin', "            'FromID
        SQl = SQl + " 'spd@ykk.com.tw', "   'FromMail
        SQl = SQl + " '系統管理', "         'FromName
        SQl = SQl + " 'Joy', "              'ToID
        SQl = SQl + " 'joo@ykk.com.tw;jimmy-sung@ykk.com.tw', "   'ToMail
        SQl = SQl + " '系統開發', "         'ToName
        SQl = SQl + " 'hunter.joo@msa.hinet.net;spd@ykk.com.tw', "     'CCMail
        SQl = SQl + " '', "                 'FormNo

        SQl = SQl + " '" + Str3 + "', "     'FormSno
        SQl = SQl + " '" + Str1 + "', "     'FormName
        SQl = SQl + " '" + Str4 + "', "     'Step
        SQl = SQl + " '" + Str2 + "', "     'StepName

        SQl = SQl + " '', "                 'ApplyID
        SQl = SQl + " '', "                 'ApplyName
        SQl = SQl + " 'SYSERROR', "         'MSG Type
        SQl = SQl + " '" + Str + "', "      'MSG Name
        SQl = SQl + " '" + NowDateTime + "' "       '作成時間
        SQl = SQl + " ) "
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()

        '--郵件傳送---------
        Dim oMail As Object
        oMail = Server.CreateObject("SendMail.WFMail")
        oMail.SendMail()

    End Sub
End Class
