Imports System.Random
Imports System.Data
Imports System.Data.OleDb

Public Class Default_T
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
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

    '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
    '�ФŧR���β��ʥ��C
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
        '�ФŨϥε{���X�s�边�i��ק�C
        InitializeComponent()
    End Sub

#End Region

    Public YKK As New YKK_SPDClass   'YKK�@�q

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not Me.IsPostBack Then   '���OPostBack
        End If

        '�b�o�̩�m�ϥΪ̵{���X�H��l�ƺ���
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
