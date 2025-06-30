Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class QA1Picker
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DQASheet As System.Web.UI.WebControls.Image
    Protected WithEvents BDown As System.Web.UI.WebControls.ImageButton
    Protected WithEvents BDate As System.Web.UI.WebControls.Button
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAssembly As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BClose As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DSurface2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSurface1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DGentani As System.Web.UI.WebControls.TextBox
    Protected WithEvents DKyoudo As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DYellow As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMityaku As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNyuryoku As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DKensin As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWater As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDry As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DCPSC As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DContent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBody As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DChainType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Dsize As System.Web.UI.WebControls.DropDownList

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

            BDate.Attributes("onclick") = "CalendarPicker('QAForm1.DDate');"  '������

        End If
    End Sub

    Private Sub BDown_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BDown.Click
        Dim str As String = "[" & _
                               DDate.Text & "," & _
                               Dsize.SelectedValue & "," & DChainType.SelectedValue & "," & DBody.SelectedValue & "," & _
                               DAssembly.SelectedValue & "," & DSurface1.Text & "," & DSurface2.Text & "," & DGentani.Text & "," & _
                               DKyoudo.SelectedValue & "," & DNyuryoku.SelectedValue & "," & DKensin.SelectedValue & "," & _
                               DWater.SelectedValue & "," & DDry.SelectedValue & "," & _
                               DYellow.Text & "," & DMityaku.Text & "," & DCPSC.SelectedValue & _
                            "]"
        If Len(DContent.Text) + Len(str) < 255 Then
            If DContent.Text = "" Then
                DContent.Text = str
            Else
                DContent.Text = DContent.Text + ";" + str
            End If
        Else
            Response.Write(YKK.ShowMessage("�ҿ�ܤ��e�w�W�L�W������"))
        End If
    End Sub

    Private Sub BClose_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BClose.Click
        Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.close();</script>", "Form1.DQuality1", DContent.Text))
    End Sub
End Class
