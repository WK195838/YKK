Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ErrorPage
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Message As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label

    '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
    '�ФŧR���β��ʥ��C
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
        '�ФŨϥε{���X�s�边�i��ק�C
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Write_WaitSend()
    End Sub

    Sub Write_WaitSend()
        Dim NowDateTime As String       '�{�b����ɶ�
        Dim SQl As String
        Dim Str, Str1, Str2, Str3, Str4 As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '�{�b���

        Str = "�����o�ͨt�β��`"
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
        SQl = SQl + " '�t�κ޲z', "         'FromName
        SQl = SQl + " 'Joy', "              'ToID
        SQl = SQl + " 'joo@ykk.com.tw;jimmy-sung@ykk.com.tw', "   'ToMail
        SQl = SQl + " '�t�ζ}�o', "         'ToName
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
        SQl = SQl + " '" + NowDateTime + "' "       '�@���ɶ�
        SQl = SQl + " ) "
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()

        '--�l��ǰe---------
        Dim oMail As Object
        oMail = Server.CreateObject("SendMail.WFMail")
        oMail.SendMail()

    End Sub
End Class
