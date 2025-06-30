Imports System.Data
Imports System.Data.OleDb

Public Class MapPicker_01
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents BKey As System.Web.UI.WebControls.Button
    Protected WithEvents DKey As System.Web.UI.WebControls.TextBox

    '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
    '�ФŧR���β��ʥ��C
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
        '�ФŨϥε{���X�s�边�i��ק�C
        InitializeComponent()
    End Sub

#End Region

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܼ�
    '**
    '*****************************************************************
    Dim pKey As String     'Search Key

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not Me.IsPostBack Then   '���OPostBack
            If Request.Cookies("MapNo").Value <> "" Then
                DKey.Text = Request.Cookies("MapNo").Value
                pKey = Request.Cookies("MapNo").Value
            Else
                DKey.Text = ""
                pKey = "ALL"
            End If
            MapData()  '���o���
        End If
    End Sub


    Sub MapData()
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim SQL As String

        If Request.QueryString("field") = "Ori" Then
            SQL = "SELECT "   '���oDB���
            SQL = SQL & "MapNo, 'MapSheet_02.aspx?pFormNo=' + FormNo + '&pFormSno=' + str(FormSno, Len(FormSno)) As URL "
            SQL = SQL & "FROM F_MapSheet "
        Else
            SQL = "SELECT "   '���oDB���
            SQL = SQL & "MapNo, 'MapModSheet_02.aspx?pFormNo=' + FormNo + '&pFormSno=' + str(FormSno, Len(FormSno)) As URL "
            SQL = SQL & "FROM F_MapModSheet "
        End If
        SQL = SQL + "Where Sts = '1' "
        If (pKey <> "ALL") And (pKey <> "") Then
            SQL = SQL + "And MapNo Like '%" & pKey & "%' "
        End If
        SQL = SQL + "Order by CompletedTime Desc "

        Dim DBDataSet1 As New DataSet
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Map")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()

        OleDbConnection1.Close()
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex   'DataGrid���W�U��
        MapData()
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandSource.CommandName = "Select" Then  '�I�����ˬd
            Dim Key As String = DataGrid1.DataKeys(e.Item.ItemIndex)  '�ҿ����Map No
            Response.Cookies("MapNo").Value = Key

            Dim wLevel As String = ""
            Dim wFormNo As String = ""
            Dim wFormSno As Integer = 0

            Dim SQL As String
            Dim DBDataSet1 As New DataSet

            'DB�s���]�w
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
            OleDbConnection1.Open()

            If Request.QueryString("field") = "Ori" Then
                SQL = "SELECT "   '���oDB���
                SQL = SQL & "FormNo, FormSno, Level "
                SQL = SQL & "FROM F_MapSheet "
            Else
                SQL = "SELECT "   '���oDB���
                SQL = SQL & "FormNo, FormSno, Level "
                SQL = SQL & "FROM F_MapModSheet "
            End If
            SQL = SQL & " Where MapNo =  '" & Key & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "F_MapSheet")
            If DBDataSet1.Tables("F_MapSheet").Rows.Count > 0 Then
                wLevel = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Level")
                wFormNo = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("FormNo")
                wFormSno = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("FormSno")
            End If
            'DB�s������
            OleDbConnection1.Close()

            If Request.QueryString("pFormNo") = "000004" Or Request.QueryString("pFormNo") = "000008" Then
                Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.close();</script>", "Form1.DMapNo", Key, "Form1.DLevel", wLevel))
            End If

            If Request.QueryString("pFormNo") = "000005" Or Request.QueryString("pFormNo") = "000009" Then
                Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.close();</script>", "Form1.DMapNo", Key))
            End If

        End If

    End Sub

    Private Sub BKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BKey.Click
        DataGrid1.CurrentPageIndex = 0  'DataGrid���W�U��
        pKey = DKey.Text
        MapData()
    End Sub

End Class
