Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class SpecPicker
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DChainType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBody As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BDown As System.Web.UI.WebControls.ImageButton
    Protected WithEvents BClose As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents Dsize As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQASheet As System.Web.UI.WebControls.Image
    Protected WithEvents LChainType As System.Web.UI.WebControls.HyperLink

    '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
    '�ФŧR���β��ʥ��C
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
        '�ФŨϥε{���X�s�边�i��ק�C
        InitializeComponent()
    End Sub

#End Region

    '*****************************************************************
    '**
    '**     �ۭq�[���w
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD�t�@�q�[��

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not Me.IsPostBack Then   '���OPostBack
        
            Datalist()
        End If
    End Sub

    Private Sub BDown_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BDown.Click
        Dim str As String
        If Dsize.SelectedValue <> "�s��" Then
            str = "[" & Dsize.SelectedValue & "-" & DChainType.SelectedValue & "-" & DBody.SelectedValue & "]"
            If DSpec.Text = "" Then
                DSpec.Text = str
            Else
                DSpec.Text = DSpec.Text & ";" & str
            End If
        Else
            str = DBody.SelectedValue
            If DSpec.Text = "" Then
                DSpec.Text = DBody.SelectedValue
            Else
                DSpec.Text = DSpec.Text & ";" & str
            End If


        End If
      
    End Sub

    Private Sub BClose_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BClose.Click
        If Request.QueryString("fun") = "MANUF" Then

            '--��LSubFile---------
            Dim pTarget(10) As String
            Dim RtnCode, i As Integer
            '
            For i = 1 To 10
                pTarget(i) = ""
            Next
            '
            i = 1
            Dim Str As String = DSpec.Text
            Do Until InStr(Str, ";") = 0
                pTarget(i) = Mid(Str, 1, InStr(Str, ";") - 1)
                Str = Mid(Str, InStr(Str, ";") + 1, Len(Str))
                i = i + 1
            Loop
            pTarget(i) = Str

            Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}';</script>", "Form1.DSampleA1", pTarget(1)))
            Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}';</script>", "Form1.DPriceA1", pTarget(1)))
            'Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}';</script>", "Form1.DQAA2", pTarget(1)))

            Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}';</script>", "Form1.DSampleB1", pTarget(2)))
            Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}';</script>", "Form1.DPriceB1", pTarget(2)))
            'Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}';</script>", "Form1.DQAB2", pTarget(2)))

            Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}';</script>", "Form1.DSampleC1", pTarget(3)))
            Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}';</script>", "Form1.DPriceC1", pTarget(3)))
            'Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}';</script>", "Form1.DQAC2", pTarget(3)))

            Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}';</script>", "Form1.DPriceD1", pTarget(4)))
            'Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}';</script>", "Form1.DQAD2", pTarget(4)))

            Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}';</script>", "Form1.DPriceE1", pTarget(5)))
            Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}';</script>", "Form1.DPriceF1", pTarget(6)))
            Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}';</script>", "Form1.DPriceG1", pTarget(7)))
            Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}';</script>", "Form1.DPriceH1", pTarget(8)))
            Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}';</script>", "Form1.DPriceI1", pTarget(9)))
            Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}';</script>", "Form1.DPriceJ1", pTarget(10)))
        End If

                Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.close();</script>", "Form1.DSpec", DSpec.Text))
    End Sub

    Private Sub Dsize_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Dsize.SelectedIndexChanged
        If Dsize.SelectedValue = "�s��" Then
            'DB�s���]�w
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
            Dim DBDataSet1 As New DataSet
            Dim DBTable1 As DataTable
            Dim SQL As String
            Dim i As Integer

            OleDbConnection1.Open()

            'Size
            SQL = "Select  substring(dkey,12,len(dkey)-11) Data From M_Referp Where Cat='100' and substring(dkey,1,11) =  'SIZECHAINGR' Order by unique_id "
            DChainType.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            DChainType.Items.Add("")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                DChainType.Items.Add(ListItem1)
            Next

            DBDataSet1.Clear()
            DBody.Items.Clear()
        Else
            'DB�s���]�w
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
            Dim DBDataSet1 As New DataSet
            Dim DBTable1 As DataTable
            Dim SQL As String
            Dim i As Integer

            OleDbConnection1.Open()



            'ChainType
            SQL = "Select * From M_Referp Where Cat='100' and DKey='CHAINTYPE' Order by Data "
            DChainType.Items.Clear()
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "M_Referp1")
            DBTable1 = DBDataSet1.Tables("M_Referp1")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                DChainType.Items.Add(ListItem1)
            Next

            DBDataSet1.Clear()

            '����
            SQL = "Select * From M_Referp Where Cat='100' and DKey='BODY' Order by Data "
            DBody.Items.Clear()
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet1, "M_Referp2")
            DBTable1 = DBDataSet1.Tables("M_Referp2")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                DBody.Items.Add(ListItem1)
            Next
            'DB�s������
            OleDbConnection1.Close()
        End If
    End Sub

    Sub Datalist()
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim SQL As String
        Dim i As Integer

        OleDbConnection1.Open()

        'Size
        SQL = "Select * From M_Referp Where Cat='100' and DKey='SIZE' Order by Data "
        Dsize.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Referp")
        DBTable1 = DBDataSet1.Tables("M_Referp")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("Data")
            ListItem1.Value = DBTable1.Rows(i).Item("Data")
            Dsize.Items.Add(ListItem1)
        Next

        DBDataSet1.Clear()

        'ChainType
        SQL = "Select * From M_Referp Where Cat='100' and DKey='CHAINTYPE' Order by Data "
        DChainType.Items.Clear()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Referp1")
        DBTable1 = DBDataSet1.Tables("M_Referp1")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("Data")
            ListItem1.Value = DBTable1.Rows(i).Item("Data")
            DChainType.Items.Add(ListItem1)
        Next

        DBDataSet1.Clear()

        '����
        SQL = "Select * From M_Referp Where Cat='100' and DKey='BODY' Order by Data "
        DBody.Items.Clear()
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "M_Referp2")
        DBTable1 = DBDataSet1.Tables("M_Referp2")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("Data")
            ListItem1.Value = DBTable1.Rows(i).Item("Data")
            DBody.Items.Add(ListItem1)
        Next
        'DB�s������
        OleDbConnection1.Close()
    End Sub


    Private Sub DropDownList1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub DChainType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DChainType.SelectedIndexChanged
        If Mid(DChainType.SelectedValue, 1, 1) = "#" Then
            'DB�s���]�w
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
            Dim DBDataSet1 As New DataSet
            Dim DBTable1 As DataTable
            Dim SQL As String
            Dim i As Integer

            OleDbConnection1.Open()

            'Size
            SQL = "select  Data  From M_Referp Where Cat='100' and dkey =  'SIZECHAINGR'+'" + DChainType.SelectedValue + "'"
            DBody.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                DBody.Items.Add(ListItem1)
            Next

            DBDataSet1.Clear()
        Else
            'DB�s���]�w
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
            Dim DBDataSet1 As New DataSet
            Dim DBTable1 As DataTable
            Dim SQL As String
            Dim i As Integer

            OleDbConnection1.Open()



            '����
            SQL = "Select * From M_Referp Where Cat='100' and DKey='BODY' Order by Data "
            DBody.Items.Clear()
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet1, "M_Referp2")
            DBTable1 = DBDataSet1.Tables("M_Referp2")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                DBody.Items.Add(ListItem1)
            Next
            'DB�s������
            OleDbConnection1.Close()
        End If
    End Sub
End Class
