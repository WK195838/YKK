Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class CalendarList_TP2
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents cal As System.Web.UI.WebControls.Calendar

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
        Response.Cookies("PGM").Value = "CalendarList_TP2.aspx"
        If Not Me.IsPostBack Then
        End If
    End Sub

    Sub Cal_DayRender(ByVal sender As Object, ByVal e As DayRenderEventArgs)
        Dim SQL As String
        Dim objDb As New SqlConnection(ConfigurationSettings.AppSettings.Get("SqlConn1"))    'SQL�s���]�w

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
