Imports System.Data
Imports System.Data.OleDb

Public Class CodeNoPicker
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DKey As System.Web.UI.WebControls.TextBox
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents BKey As System.Web.UI.WebControls.Button

    '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
    '�ФŧR���β��ʥ��C
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
        '�ФŨϥε{���X�s�边�i��ק�C
        InitializeComponent()
    End Sub

#End Region

    Public YKK As New YKK_SPDClass   'YKK�@�q�[��

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܼ�
    '**
    '*****************************************************************
    Dim pKey As String     'Search Key

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not Me.IsPostBack Then   '���OPostBack
            If Request.Cookies("CodeNo").Value <> "" Then
                DKey.Text = Request.Cookies("CodeNo").Value
                pKey = Request.Cookies("CodeNo").Value
            Else
                DKey.Text = ""
                pKey = "ALL"
            End If
            CodeData()  '���o���
        End If
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandSource.CommandName = "Select" Then  '�I�����ˬd
            Dim Key As String = DataGrid1.DataKeys(e.Item.ItemIndex)  '�ҿ����Code

            Dim wNo As String = ""
            Dim wFormNo As String = ""
            Dim wFormSno As String = ""
            Dim wBuyer As String = ""
            Dim wSellVendor As String = ""

            Dim SQL As String
            Dim DBDataSet1 As New DataSet

            'DB�s���]�w
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
            OleDbConnection1.Open()

            If Request.QueryString("field") = "Suface" Then
                SQL = "Select Code, No, FormNo, FormSno, Buyer, SellVendor From F_SufaceSheet "
            End If

            SQL = SQL & " Where Sts =  '1' "
            If pKey <> "ALL" Then
                SQL = SQL & "   And  Code Like '%" & Key & "%'"
            End If

            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "F_SufaceSheet")
            If DBDataSet1.Tables("F_SufaceSheet").Rows.Count > 0 Then
                wNo = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("No")
                wFormNo = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FormNo")
                wFormSno = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FormSno")
                wBuyer = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Buyer")
                wSellVendor = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("SellVendor")
            End If
            'DB�s������
            OleDbConnection1.Close()

            '�ҿ����Map No�^������������
            Dim Cmd As String
            Response.Cookies("CodeNo").Value = Key

            If Request.QueryString("pFormNo") = "000015" Then
                Cmd = String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.opener.document.{4}.value = '{5:d}'; window.opener.document.{6}.value = '{7:d}'; window.opener.document.{8}.value = '{9:d}'; window.opener.document.{10}.value = '{11:d}'; window.close();</script>", "Form1.DCode", Key, "Form1.DONo", wNo, "Form1.DOFormNo", wFormNo, "Form1.DOFormSno", wFormSno, "Form1.DBuyer", wBuyer, "Form1.DSellVendor", wSellVendor)
            End If

            Response.Write(Cmd)
        End If
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex   'DataGrid���W�U��
        'DataGrid1.DataBind()
        CodeData()
    End Sub

    Sub CodeData()
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim SQL As String

        SQL = "SELECT "   '���oDB���
        SQL = SQL & "Code, No, "

        If Request.QueryString("field") = "Suface" Then
            SQL = SQL & "'SufaceSheet_02.aspx?pFormNo=' + FormNo + '&pFormSno=' + str(FormSno, Len(FormSno)) As URL "
            SQL = SQL & "FROM F_SufaceSheet "
        End If
        SQL = SQL & " Where Sts =  '1' "

        If Request.QueryString("pFormNo") = "000015" Then
            SQL = SQL + "And   UPDSts = '0' "
        End If

        If (pKey <> "ALL") And (pKey <> "") Then
            SQL = SQL + "And Code Like '%" & pKey & "%' "
        End If
        SQL = SQL + "Order by CompletedTime Desc "

        Dim DBDataSet1 As New DataSet
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_SufaceSheet")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()

        OleDbConnection1.Close()
    End Sub

    Private Sub BKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BKey.Click
        DataGrid1.CurrentPageIndex = 0  'DataGrid���W�U��
        pKey = DKey.Text
        CodeData()
    End Sub

End Class
