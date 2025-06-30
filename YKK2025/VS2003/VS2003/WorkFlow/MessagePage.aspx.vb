Imports System.Data
Imports System.Data.OleDb

Public Class CompletedPage
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Message As System.Web.UI.WebControls.TextBox

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
                                    CStr(DateTime.Now.Second)           '�{�b���
        Dim i As Integer = 0
        Dim SQL As String
        Dim wFormName, wStepName, wUserName, wApplyName As String

        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        '���W
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

        '�u�{�W
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

        'ñ�֪�
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

        '�e�U��
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

        '�U�@���֩w�H
        Dim oGetUserName As Object
        Dim pUser As String = Request.QueryString("pNextGate")
        Dim pName As String = ""
        Dim RtnCode As Integer = 0

        oGetUserName = Server.CreateObject("GetUserName.WFGetUserName")
        RtnCode = oGetUserName.UserName(pUser, pName)

        Message.Text = "[ " & wUserName & " ]" & "��" & NowDateTime & _
                       "�N" & "[ " & wApplyName & " ]" & "��" & "[ " & wFormName & " ]" & _
                       "���\�e��" & "[ " & wStepName & "(" & pName & ") ]�D"
    End Sub

    Sub SavedMessage()
        Dim NowDateTime As String = CStr(DateTime.Now.Today) + " " + _
                                    CStr(DateTime.Now.Hour) + ":" + _
                                    CStr(DateTime.Now.Minute) + ":" + _
                                    CStr(DateTime.Now.Second)           '�{�b���
        Dim i As Integer = 0
        Dim SQL As String
        Dim wFormName, wStepName, wUserName, wApplyName As String

        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        '���W
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

        '�u�{�W
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

        'ñ�֪�
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

        '�e�U��
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

        Message.Text = "[ " & wStepName & ":" & wUserName & " ]" & "��" & NowDateTime & _
                       "�N" & "[ " & wApplyName & " ]" & "��" & "[ " & wFormName & " ]" & "����x�s���\�D"
    End Sub

    Sub NGMessage()
        Dim NowDateTime As String = CStr(DateTime.Now.Today) + " " + _
                                    CStr(DateTime.Now.Hour) + ":" + _
                                    CStr(DateTime.Now.Minute) + ":" + _
                                    CStr(DateTime.Now.Second)           '�{�b���
        Dim i As Integer = 0
        Dim SQL As String
        Dim wFormName, wStepName, wUserName, wApplyName As String

        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        '���W
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

        '�u�{�W
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

        'ñ�֪�
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

        '�e�U��
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

        Message.Text = "[ " & wStepName & ":" & wUserName & " ]" & "��" & NowDateTime & _
                       "�N" & "[ " & wApplyName & " ]" & "��" & "[ " & wFormName & " ]" & "���NG���\�D"
    End Sub

End Class
