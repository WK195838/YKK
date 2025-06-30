Imports System.Data
Imports System.Data.OleDb

Public Class SPDNewCommission
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DNewSheet As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DLevel As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSimulation As System.Web.UI.WebControls.Button
    Protected WithEvents DNew As System.Web.UI.WebControls.Button
    Protected WithEvents TextBox1 As System.Web.UI.WebControls.TextBox
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

    '*****************************************************************
    '**
    '**     �ۭq�[���w
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD�t�@�q�[��

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "SPDNewCommission.aspx"
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
        If Not Me.IsPostBack Then
            SetNewSheet()
            SetLevel()
        End If
    End Sub

    Sub SetLevel()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        Dim wFormNo As String = DNewSheet.SelectedValue          '��渹�X
        SQL = "SELECT * FROM M_Referp  "
        SQL = SQL + "Where Cat = '007' "
        Select Case wFormNo
            Case "000001"       '�ϭ��e�U
                SQL = SQL + "and (DKey like '0%' or DKey like 'Z%') "
            Case "000002"       '�ϭ��ק�e�U
                SQL = SQL + "and (DKey like '0%' or DKey like 'Z%') "
            Case "000003"       '���s�e�U
                SQL = SQL + "and DKey like 'In%' "
            Case "000004"       '�u�{�s����
                SQL = SQL + "and DKey like 'In%' "
            Case "000005"       '�~�ȳs����
                SQL = SQL + "and DKey like '0%' "
            Case "000007"       '�~�`�e�U
                SQL = SQL + "and DKey like 'Out%' "
            Case "000008"       '�~�`�u�{�s����
                SQL = SQL + "and DKey like 'Out%' "
            Case "000009"       '�~�`�~�ȳs����
                SQL = SQL + "and DKey like '0%' "
            Case "000010"       '���Y�ӥس�
                SQL = SQL + "and DKey like '0%' "
            Case "000011"       '�i�f/�Ȩѩ��Y
                SQL = SQL + "and DKey like '0%' "
            Case "000012"       '���O����l�[�e�U��
                SQL = SQL + "and DKey like '0%' "
            Case "000013"       '�i�f/�Ȩѩ��Y�~�ȳs����
                SQL = SQL + "and DKey like '0%' "
            Case "000014"       '���B�z�e�U��
                SQL = SQL + "and DKey like '0%' "
            Case "000015"       '���B�z�l�[�e�U��
                SQL = SQL + "and DKey like '0%' "
            Case "800001"
                SQL = SQL + "and DKey like '0%' "
            Case "900050"       '�ª��ϭ��e�U
                SQL = SQL + "and (DKey like '0%' or DKey like 'Z%') "
            Case Else
                SQL = SQL + "and DKey like 'Other%' "
        End Select
        SQL = SQL + "Order by Cat, DKey, Data "

        OleDbConnection1.Open()
        DLevel.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Referp")
        DBTable1 = DBDataSet1.Tables("M_Referp")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("Data")
            ListItem1.Value = DBTable1.Rows(i).Item("Data")
            DLevel.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����I���ƥ�
    '**
    '*****************************************************************
    Private Sub DNewSheet_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DNewSheet.SelectedIndexChanged
        SetLevel()
    End Sub

    Sub SetNewSheet()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        SQL = "SELECT FormName, FormNo FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And ( (FormNo >= '000001' And FormNo <= '001000' "
        SQL = SQL + "         And (IniAuthority = '0'  "
        SQL = SQL + "             Or (IniAuthority = '1' "
        SQL = SQL + "                And (   IniUser     like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                    Or  IniUserName like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                    ) "
        SQL = SQL + "                ) "
        SQL = SQL + "             ) "
        SQL = SQL + "        ) "
        SQL = SQL + "        or "
        SQL = SQL + "        (FormNo >= '900050' And FormNo <= '900999' "
        SQL = SQL + "         And (IniAuthority = '0'  "
        SQL = SQL + "             Or (IniAuthority = '1' "
        SQL = SQL + "                And (   IniUser     like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                    Or  IniUserName like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                    ) "
        SQL = SQL + "                ) "
        SQL = SQL + "             ) "
        SQL = SQL + "        ) "
        SQL = SQL + "      ) "
        SQL = SQL + "Order by FormName "

        OleDbConnection1.Open()
        DNewSheet.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_FORM")
        DBTable1 = DBDataSet1.Tables("M_FORM")

        DNew.Visible = False
        DSimulation.Visible = False
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("FormName")
            ListItem1.Value = DBTable1.Rows(i).Item("FormNo")
            DNewSheet.Items.Add(ListItem1)
            DNew.Visible = True
            DSimulation.Visible = True
        Next
        OleDbConnection1.Close()
    End Sub

    Private Sub DNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DNew.Click
        Dim wFormNo As String = DNewSheet.SelectedValue          '��渹�X
        Dim wFormSno As Integer = 0       '���y����
        Dim wStep As Integer = 1          '�u�{�N�X
        Dim SPDStep As Integer = 1        'SPD�u�{�N�X
        Dim wSeqNo As Integer = 0         '�Ǹ�
        Dim URL As String = ""            'URL

        If Request.QueryString("pUserID") <> "" Then
            Dim SQL As String
            Dim DBDataSet1 As New DataSet
            'DB�s���]�w
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
            OleDbConnection1.Open()

            SQL = "Select * From M_Users "
            SQL = SQL & " Where Active =  '1' "
            SQL = SQL & "   And UserID =  '" & Request.QueryString("pUserID") & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Users")
            If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then
                If DBDataSet1.Tables("M_Users").Rows(0).Item("DivID") = "1118211" Then SPDStep = 2 '�x��
                If DBDataSet1.Tables("M_Users").Rows(0).Item("DivID") = "1018950" Then SPDStep = 2 'EA
                'DB�s������
                OleDbConnection1.Close()

                Select Case wFormNo
                    Case "000001"       '�ϭ��e�U
                        URL = "MapSheet_01.aspx?pFormNo=" & wFormNo & _
                                              "&pFormSno=" & wFormSno & _
                                              "&pStep=" & SPDStep & _
                                              "&pSeqNo=" & wSeqNo & _
                                              "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000002"       '�ϭ��ק�e�U
                        URL = "MapModSheet_01.aspx?pFormNo=" & wFormNo & _
                                                 "&pFormSno=" & wFormSno & _
                                                 "&pStep=" & SPDStep & _
                                                 "&pSeqNo=" & wSeqNo & _
                                                 "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000003"       '���s�e�U
                        URL = "ManufInSheet_01.aspx?pFormNo=" & wFormNo & _
                                                 "&pFormSno=" & wFormSno & _
                                                 "&pStep=" & SPDStep & _
                                                 "&pSeqNo=" & wSeqNo & _
                                                 "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000004"       '�u�{�s����
                        URL = "ManufInOPSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000005"       '�~�ȳs����
                        URL = "ManufInCTSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000007"       '�~�`�e�U
                        URL = "ManufOutSheet_01.aspx?pFormNo=" & wFormNo & _
                                                 "&pFormSno=" & wFormSno & _
                                                 "&pStep=" & SPDStep & _
                                                 "&pSeqNo=" & wSeqNo & _
                                                 "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000008"       '�~�`�u�{�s����
                        URL = "ManufOutOPSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000009"       '�~�`�~�ȳs����
                        URL = "ManufOutCTSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000010"       '���Y�ӥس�
                        URL = "ManufOutSDSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000011"       '�i�f/�Ȩѩ��Y
                        URL = "ImportSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000012"       '���O����l�[�e�U��
                        URL = "AppendSpecSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000013"       '�i�f/�Ȩѩ��Y�s����
                        URL = "ImportCTSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000014"       '���B�z�e�U��
                        URL = "SufaceSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000015"       '���B�z�l�[�e�U��
                        URL = "SufaceAppendSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "800001"
                        URL = "ModifyDataSheet_01.aspx?pFormNo=" & wFormNo & _
                                                 "&pFormSno=" & wFormSno & _
                                                 "&pStep=" & wStep & _
                                                 "&pSeqNo=" & wSeqNo & _
                                                 "&pApplyID=" & Request.QueryString("pUserID")
                    Case "900050"        '�¹ϭ��e�U��
                        URL = "MapBefSheet_01.aspx?pFormNo=" & wFormNo & _
                                                 "&pFormSno=" & wFormSno & _
                                                 "&pStep=" & wStep & _
                                                 "&pSeqNo=" & wSeqNo & _
                                                 "&pApplyID=" & Request.QueryString("pUserID")
                    Case "900051"        '�¤��s�e�U��
                        URL = "ManufInBefSheet_01.aspx?pFormNo=" & wFormNo & _
                                                 "&pFormSno=" & wFormSno & _
                                                 "&pStep=" & wStep & _
                                                 "&pSeqNo=" & wSeqNo & _
                                                 "&pApplyID=" & Request.QueryString("pUserID")
                    Case "900052"        '�¥~�`�e�U��
                        URL = "ManufOutBefSheet_01.aspx?pFormNo=" & wFormNo & _
                                                 "&pFormSno=" & wFormSno & _
                                                 "&pStep=" & wStep & _
                                                 "&pSeqNo=" & wSeqNo & _
                                                 "&pApplyID=" & Request.QueryString("pUserID")
                    Case Else
                End Select

                Response.Redirect(URL)
            End If
        End If
        Response.Write(YKK.ShowMessage(Request.QueryString("pUserID")))
    End Sub

    Private Sub DSimulation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DSimulation.Click
        Dim wFormNo As String = DNewSheet.SelectedValue    '��渹�X
        Dim wFormSno As String = 0       '�渹
        Dim wStep As Integer = 1         '�u�{�N�X
        Dim SPDStep As Integer = 1        'SPD�u�{�N�X
        Dim wLevel As String = DLevel.SelectedValue    '������
        Dim URL As String = ""            'URL

        If Request.QueryString("pUserID") <> "" Then
            Dim SQL As String
            Dim DBDataSet1 As New DataSet
            'DB�s���]�w
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
            OleDbConnection1.Open()
            SQL = "Select * From M_Users "
            SQL = SQL & " Where Active =  '1' "
            SQL = SQL & "   And UserID =  '" & Request.QueryString("pUserID") & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Users")
            If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then
                If DBDataSet1.Tables("M_Users").Rows(0).Item("DIVID") = "1118211" Or _
                   DBDataSet1.Tables("M_Users").Rows(0).Item("DIVID") = "1018950" Then
                    SPDStep = 2
                End If
            End If
            'DB�s������
            OleDbConnection1.Close()

            Select Case wFormNo
                Case "000001"       '�ϭ��e�U
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000002"       '�ϭ��ק�e�U
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000003"       '���s�e�U
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000004"       '�u�{�s����
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000005"       '�~�ȳs����
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000007"       '�~�`�e�U
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000008"       '�~�`�u�{�s����
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000009"       '�~�`�~�ȳs����
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000010"       '�~�ȳs����
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000011"       '�i�f/�Ȩ�
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000012"       '���O/����
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000013"       '�i�f/�Ȩѷ~�ȳs����
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000014"       '���B�z�e�U��
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000015"       '���B�z�l�[�e�U��
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "800001"
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & wStep & _
                                          "&pLevel=" & wLevel
                Case Else
            End Select

            Response.Redirect(URL)
        End If
    End Sub

End Class