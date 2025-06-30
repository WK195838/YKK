Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class SingleOPPerformance
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DStep As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSDate As System.Web.UI.WebControls.Button
    Protected WithEvents DEDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BEDate As System.Web.UI.WebControls.Button
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents ChartUnit4 As System.Web.UI.WebControls.Label
    Protected WithEvents ChartUnit3 As System.Web.UI.WebControls.Label
    Protected WithEvents ChartUnit2 As System.Web.UI.WebControls.Label
    Protected WithEvents ChartTitle4 As System.Web.UI.WebControls.Label
    Protected WithEvents ChartTitle3 As System.Web.UI.WebControls.Label
    Protected WithEvents ChartTitle2 As System.Web.UI.WebControls.Label
    Protected WithEvents ChartUnit1 As System.Web.UI.WebControls.Label
    Protected WithEvents ChartTitle1 As System.Web.UI.WebControls.Label
    Protected WithEvents BarImage4 As System.Web.UI.WebControls.Image
    Protected WithEvents BarImage3 As System.Web.UI.WebControls.Image
    Protected WithEvents BarImage2 As System.Web.UI.WebControls.Image
    Protected WithEvents BarImage1 As System.Web.UI.WebControls.Image
    Protected WithEvents DMsgBox As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSts As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSeqNo As System.Web.UI.WebControls.DropDownList

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
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܼ�
    '**
    '*****************************************************************
    Dim pFormNo As String
    Dim NowDateTime As String       '�{�b����ɶ�

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "SingleOPPerformance.aspx"

        SetParameter()          '�]�w�@�ΰѼ�
        If Not Me.IsPostBack Then
            pFormNo = "000001"
            SetSearchItem()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�@�ΰѼ�
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(DateTime.Now.Year) + _
                      CStr(DateTime.Now.Month) + _
                      CStr(DateTime.Now.Day) + _
                      CStr(DateTime.Now.Hour) + _
                      CStr(DateTime.Now.Minute) + _
                      CStr(DateTime.Now.Second)     '�{�b���

        pFormNo = Request.QueryString("pFormNo")    '��渹�X
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")

        DMsgBox.Visible = True
        DMsgBox.Text = "���\��t�s�@�έp�ϩҥH�t�׷|����C�A�Э@�ߵ���.."

        BarImage1.Visible = False   '��ƦX�p
        BarImage2.Visible = False   '��Ƥ�v
        BarImage3.Visible = False   '�w�w���Z�X�p
        BarImage4.Visible = False   '�w�w���Z��v

        ChartTitle1.Visible = False
        ChartTitle2.Visible = False
        ChartTitle3.Visible = False
        ChartTitle4.Visible = False
        ChartUnit1.Visible = False
        ChartUnit2.Visible = False
        ChartUnit3.Visible = False
        ChartUnit4.Visible = False

        DataGrid1.Visible = False
    End Sub

    Sub SetSearchItem()
        Dim i As Integer = 0
        Dim wSts As String
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        '���
        SQL = "SELECT FormName, FormNo FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And ( (FormNo >= '000001' And FormNo <= '001000') or FormNo='800001' ) "
        SQL = SQL + "  And (InqAuthority = '0' "
        SQL = SQL + "       Or (InqAuthority = '1' "
        SQL = SQL + "           And (InqUser like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                Or InqUserName like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "               ) "
        SQL = SQL + "          ) "
        SQL = SQL + "      ) "
        SQL = SQL + "Order by FormNo "

        OleDbConnection1.Open()
        DFormName.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Form")
        DBTable1 = DBDataSet1.Tables("M_Form")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("FormName")
            ListItem1.Value = DBTable1.Rows(i).Item("FormNo")
            If ListItem1.Value = pFormNo Then ListItem1.Selected = True
            DFormName.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()

        'Step
        DBDataSet1.Clear()
        SQL = "Select Step, str(Step) + ':' + StepName As StepName From M_Flow "
        SQL = SQL + "Where Step     <  '900' "
        SQL = SQL + "  and FormNo   =  '" + pFormNo + "' "
        SQL = SQL + "  and FlowType <> '0'   "
        SQL = SQL + "Group by Step, StepName "
        SQL = SQL + "Order by Step, StepName "
        DStep.Items.Clear()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Flow")
        DBTable1 = DBDataSet1.Tables("M_Flow")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("StepName")
            ListItem1.Value = DBTable1.Rows(i).Item("Step")
            DStep.Items.Add(ListItem1)
        Next

        '�̿ೡ��
        DBDataSet1.Clear()
        SQL = "Select DivName From M_Users Group by DivName Order by DivName "
        DDivision.Items.Clear()
        DDivision.Items.Add("ALL")
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("DivName")
            ListItem1.Value = DBTable1.Rows(i).Item("DivName")
            DDivision.Items.Add(ListItem1)
        Next

        '���
        DSDate.Text = CStr(DateTime.Now.Year) & "/" & CStr(DateTime.Now.Month) & "/1"
        DEDate.Text = CStr(DateTime.Now.Today)

        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DEDate');"

        pFormNo = DFormName.SelectedValue
        Search_Item_Attribute()

    End Sub

    Sub Search_Item_Attribute()
        DPerson.ReadOnly = True
        DBuyer.ReadOnly = True

        '�̿��
        DPerson.Text = "�̿��"
        'Buyer
        DBuyer.Text = "Buyer"

        DPerson.ReadOnly = False
        If pFormNo = "000001" Then
            DBuyer.ReadOnly = False
        End If
        If pFormNo = "000002" Then
            DBuyer.ReadOnly = False
        End If
        If pFormNo = "000003" Or pFormNo = "000004" Or pFormNo = "000005" Or + _
           pFormNo = "000007" Or pFormNo = "000008" Or pFormNo = "000009" Then
            DBuyer.ReadOnly = False
        End If
        If pFormNo = "000010" Then
        End If
        If pFormNo = "000011" Then
            DBuyer.ReadOnly = False
        End If
        If pFormNo = "000012" Then
            DBuyer.ReadOnly = False
        End If
        If pFormNo = "000013" Then
        End If
    End Sub

    Sub DataList()
        Dim Item(10) As String
        Dim Data(10) As Double
        Dim Color(10) As String    'RED , BLUE, YELLOW, GREEN, WHITE, GRAY, ORANGE, PURPLE
        Dim Buffer_Data(10) As Integer
        Dim Real_Data As Double = 0

        Dim wTableName As String = ""
        Dim wSDate As String = DSDate.Text & " 00:00:00"
        Dim wEDate As String = DEDate.Text & " 23:59:59"

        Dim i As Integer = 0
        Dim Records As Integer = 0
        Dim RtnCode As Integer = 0

        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection

        For i = 1 To 10
            Item(i) = ""
            Data(i) = 0
            Buffer_Data(i) = 0
            Color(i) = ""
        Next

        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()
        SQL = "SELECT FormNo, FormName, TableName1 FROM M_Form "
        SQL = SQL + " Where Active  =  '1' "
        SQL = SQL + "   And FormNo = '" + DFormName.SelectedValue + "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Form")
        If DBDataSet1.Tables("Form").Rows.Count > 0 Then
            wTableName = "V_" + DBDataSet1.Tables("Form").Rows(0).Item("TableName1") + "_01"
        End If
        OleDbConnection1.Close()

        If wTableName <> "" Then
            '�`��� 
            DBDataSet2.Clear()
            OleDbConnection1.Open()
            SQL = "SELECT isnull(Count(*),0) As RCount "
            SQL = SQL + "FROM " + wTableName + " "
            SQL = SQL + "Where T_Active = '0' "
            SQL = SQL + "  And T_FormNo = '" + DFormName.SelectedValue + "' "
            SQL = SQL + "  And T_Step   = '" + DStep.SelectedValue + "' "
            SQL = SQL + "  And T_CompletedTime >= '" + wSDate + "' "
            SQL = SQL + "  And T_CompletedTime <= '" + wEDate + "' "
            SQL = SQL + "  And FormSno  > '2000' "
            SQL = SQL + "  And Action = '0' "

            If DSts.SelectedValue <> "ALL" Then
                SQL = SQL + " And Sts = '" + DSts.SelectedValue + "'"
            End If
            If DDivision.SelectedValue <> "ALL" Then
                SQL = SQL + " And Division = '" + DDivision.SelectedValue + "'"
            End If
            If DPerson.Text <> "�̿��" And DPerson.Text <> "" Then
                SQL = SQL + " And Person Like '%" + DPerson.Text + "%'"
            End If
            If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
                SQL = SQL + " And Buyer Like '%" + DBuyer.Text + "%'"
            End If
            SQL = SQL + " And T_SeqNo = '" + DSeqNo.SelectedValue + "'"

            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "WaitHandle")
            If DBDataSet2.Tables("WaitHandle").Rows.Count > 0 Then
                Item(1) = "Total"
                Data(1) = DBDataSet2.Tables("WaitHandle").Rows(0).Item("RCount")
                Buffer_Data(1) = Data(1)
                Records = Data(1)
                Color(1) = "RED"
            End If
            OleDbConnection1.Close()

            '�}�o����� 
            DBDataSet2.Clear()
            OleDbConnection1.Open()
            SQL = "SELECT isnull(Count(*),0) As RCount "
            SQL = SQL + "FROM " + wTableName + " "
            SQL = SQL + "Where T_Active = '0' "
            SQL = SQL + "  And T_FormNo = '" + DFormName.SelectedValue + "' "
            SQL = SQL + "  And T_Step   = '" + DStep.SelectedValue + "' "
            SQL = SQL + "  And T_CompletedTime >= '" + wSDate + "' "
            SQL = SQL + "  And T_CompletedTime <= '" + wEDate + "' "
            SQL = SQL + "  And FormSno  > '2000' "
            SQL = SQL + "  And Sts      = '0' "
            SQL = SQL + "  And Action = '0' "

            If DSts.SelectedValue <> "ALL" Then
                SQL = SQL + " And Sts = '" + DSts.SelectedValue + "'"
            End If
            If DDivision.SelectedValue <> "ALL" Then
                SQL = SQL + " And Division = '" + DDivision.SelectedValue + "'"
            End If
            If DPerson.Text <> "�̿��" And DPerson.Text <> "" Then
                SQL = SQL + " And Person Like '%" + DPerson.Text + "%'"
            End If
            If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
                SQL = SQL + " And Buyer Like '%" + DBuyer.Text + "%'"
            End If
            SQL = SQL + " And T_SeqNo = '" + DSeqNo.SelectedValue + "'"

            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet2, "WaitHandle")
            If DBDataSet2.Tables("WaitHandle").Rows.Count > 0 Then
                Item(2) = "�}�o��"
                Data(2) = DBDataSet2.Tables("WaitHandle").Rows(0).Item("RCount")
                Buffer_Data(2) = Data(2)
                Color(2) = "BLUE"
            End If
            OleDbConnection1.Close()
            '����-OK��� 
            DBDataSet2.Clear()
            OleDbConnection1.Open()
            SQL = "SELECT isnull(Count(*),0) As RCount "
            SQL = SQL + "FROM " + wTableName + " "
            SQL = SQL + "Where T_Active = '0' "
            SQL = SQL + "  And T_FormNo = '" + DFormName.SelectedValue + "' "
            SQL = SQL + "  And T_Step   = '" + DStep.SelectedValue + "' "
            SQL = SQL + "  And T_CompletedTime >= '" + wSDate + "' "
            SQL = SQL + "  And T_CompletedTime <= '" + wEDate + "' "
            SQL = SQL + "  And FormSno  > '2000' "
            SQL = SQL + "  And Sts      = '1' "
            SQL = SQL + "  And Action = '0' "

            If DSts.SelectedValue <> "ALL" Then
                SQL = SQL + " And Sts = '" + DSts.SelectedValue + "'"
            End If
            If DDivision.SelectedValue <> "ALL" Then
                SQL = SQL + " And Division = '" + DDivision.SelectedValue + "'"
            End If
            If DPerson.Text <> "�̿��" And DPerson.Text <> "" Then
                SQL = SQL + " And Person Like '%" + DPerson.Text + "%'"
            End If
            If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
                SQL = SQL + " And Buyer Like '%" + DBuyer.Text + "%'"
            End If
            SQL = SQL + " And T_SeqNo = '" + DSeqNo.SelectedValue + "'"

            Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter4.Fill(DBDataSet2, "WaitHandle")
            If DBDataSet2.Tables("WaitHandle").Rows.Count > 0 Then
                Item(3) = "OK"
                Data(3) = DBDataSet2.Tables("WaitHandle").Rows(0).Item("RCount")
                Buffer_Data(3) = Data(3)
                Color(3) = "YELLOW"
            End If
            OleDbConnection1.Close()
            '����-NG��� 
            DBDataSet2.Clear()
            OleDbConnection1.Open()
            SQL = "SELECT isnull(Count(*),0) As RCount "
            SQL = SQL + "FROM " + wTableName + " "
            SQL = SQL + "Where T_Active = '0' "
            SQL = SQL + "  And T_FormNo = '" + DFormName.SelectedValue + "' "
            SQL = SQL + "  And T_Step   = '" + DStep.SelectedValue + "' "
            SQL = SQL + "  And T_CompletedTime >= '" + wSDate + "' "
            SQL = SQL + "  And T_CompletedTime <= '" + wEDate + "' "
            SQL = SQL + "  And FormSno  > '2000' "
            SQL = SQL + "  And Sts      = '2' "
            SQL = SQL + "  And Action = '0' "

            If DSts.SelectedValue <> "ALL" Then
                SQL = SQL + " And Sts = '" + DSts.SelectedValue + "'"
            End If
            If DDivision.SelectedValue <> "ALL" Then
                SQL = SQL + " And Division = '" + DDivision.SelectedValue + "'"
            End If
            If DPerson.Text <> "�̿��" And DPerson.Text <> "" Then
                SQL = SQL + " And Person Like '%" + DPerson.Text + "%'"
            End If
            If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
                SQL = SQL + " And Buyer Like '%" + DBuyer.Text + "%'"
            End If
            SQL = SQL + " And T_SeqNo = '" + DSeqNo.SelectedValue + "'"

            Dim DBAdapter5 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter5.Fill(DBDataSet2, "WaitHandle")
            If DBDataSet2.Tables("WaitHandle").Rows.Count > 0 Then
                Item(4) = "NG"
                Data(4) = DBDataSet2.Tables("WaitHandle").Rows(0).Item("RCount")
                Buffer_Data(4) = Data(4)
                Color(4) = "GREEN"
            End If
            OleDbConnection1.Close()
            '������� 
            DBDataSet2.Clear()
            OleDbConnection1.Open()
            SQL = "SELECT isnull(Count(*),0) As RCount "
            SQL = SQL + "FROM " + wTableName + " "
            SQL = SQL + "Where T_Active = '0' "
            SQL = SQL + "  And T_FormNo = '" + DFormName.SelectedValue + "' "
            SQL = SQL + "  And T_Step   = '" + DStep.SelectedValue + "' "
            SQL = SQL + "  And T_CompletedTime >= '" + wSDate + "' "
            SQL = SQL + "  And T_CompletedTime <= '" + wEDate + "' "
            SQL = SQL + "  And FormSno  > '2000' "
            SQL = SQL + "  And Sts      = '3' "
            SQL = SQL + "  And Action = '0' "

            If DSts.SelectedValue <> "ALL" Then
                SQL = SQL + " And Sts = '" + DSts.SelectedValue + "'"
            End If
            If DDivision.SelectedValue <> "ALL" Then
                SQL = SQL + " And Division = '" + DDivision.SelectedValue + "'"
            End If
            If DPerson.Text <> "�̿��" And DPerson.Text <> "" Then
                SQL = SQL + " And Person Like '%" + DPerson.Text + "%'"
            End If
            If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
                SQL = SQL + " And Buyer Like '%" + DBuyer.Text + "%'"
            End If
            SQL = SQL + " And T_SeqNo = '" + DSeqNo.SelectedValue + "'"

            Dim DBAdapter6 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter6.Fill(DBDataSet2, "WaitHandle")
            If DBDataSet2.Tables("WaitHandle").Rows.Count > 0 Then
                Item(5) = "����"
                Data(5) = DBDataSet2.Tables("WaitHandle").Rows(0).Item("RCount")
                Buffer_Data(5) = Data(5)
                Color(5) = "ORANGE"
            End If
            OleDbConnection1.Close()

            'BarChart
            Dim HttpPath As String = ConfigurationSettings.AppSettings.Get("Http") & _
                                     ConfigurationSettings.AppSettings.Get("ChartPath") & _
                                     "Operation-Record-" & NowDateTime & ".bmp"

            Dim FileName As String = ConfigurationSettings.AppSettings.Get("BMPPath") & "Operation-Record-" & NowDateTime & ".bmp"

            Dim oChart As Object
            oChart = Server.CreateObject("Chart1.Chart")
            RtnCode = oChart.BarChart(FileName, Item, Data, Color, 5)  'BMP�ɦW, �R���ɦW, ����, �ƭ�, �C��, ���ؼ� 
            If RtnCode = 0 Then
                BarImage1.ImageUrl = HttpPath
                BarImage1.Visible = True
                ChartTitle1.Visible = True
                ChartUnit1.Visible = True
            End If
            '**************************************************************************************************
            For i = 1 To 10
                Item(i) = ""
                Data(i) = 0
                Color(i) = ""
            Next
            '�}�o��Ƥ�v
            If Buffer_Data(1) <> 0 Then
                Item(1) = "�}�o��"
                Real_Data = Buffer_Data(2) * 100 / Buffer_Data(1)
                Data(1) = CInt(Real_Data.ToString("0"))
                Color(1) = "BLUE"
                'OK��Ƥ�v
                Item(2) = "OK"
                Real_Data = Buffer_Data(3) * 100 / Buffer_Data(1)
                Data(2) = CInt(Real_Data.ToString("0"))
                Color(2) = "YELLOW"
                'NG��Ƥ�v
                Item(3) = "NG"
                Real_Data = Buffer_Data(4) * 100 / Buffer_Data(1)
                Data(3) = CInt(Real_Data.ToString("0"))
                Color(3) = "GREEN"
                '������Ƥ�v
                Item(4) = "����"
                Real_Data = Buffer_Data(5) * 100 / Buffer_Data(1)
                Data(4) = CInt(Real_Data.ToString("0"))
                Color(4) = "ORANGE"
            Else
                Item(1) = "�}�o��"
                Color(1) = "BLUE"
                'OK��Ƥ�v
                Item(2) = "OK"
                Color(2) = "YELLOW"
                'NG��Ƥ�v
                Item(3) = "NG"
                Color(3) = "GREEN"
                '������Ƥ�v
                Item(4) = "����"
                Color(4) = "ORANGE"
            End If
            'BarChart
            HttpPath = ConfigurationSettings.AppSettings.Get("Http") & _
                       ConfigurationSettings.AppSettings.Get("ChartPath") & _
                       "Operation-Record-Ratio-" & NowDateTime & ".bmp"

            FileName = ConfigurationSettings.AppSettings.Get("BMPPath") & "Operation-Record-Ratio-" & NowDateTime & ".bmp"

            RtnCode = oChart.BarChart(FileName, Item, Data, Color, 4)   'BMP�ɦW, �R���ɦW, ����, �ƭ�, �C��, ���ؼ� 
            If RtnCode = 0 Then
                BarImage2.ImageUrl = HttpPath
                BarImage2.Visible = True
                ChartUnit2.Visible = True
                ChartTitle2.Visible = True
            End If
            '**************************************************************************************************
            For i = 1 To 10
                Item(i) = ""
                Data(i) = 0
                Buffer_Data(i) = 0
                Color(i) = ""
            Next

            '�w�w/���Z�`�ɼ� 
            Dim BSum As Integer = 0
            Dim ASum As Integer = 0

            DBDataSet2.Clear()
            OleDbConnection1.Open()
            SQL = "SELECT isnull(Sum(BWorkTime),0) as BWorkTime, isnull(Sum(AWorkTime),0) as AWorkTime "
            SQL = SQL + "FROM " + wTableName + " "
            SQL = SQL + "Where T_Active = '0' "
            SQL = SQL + "  And T_FormNo = '" + DFormName.SelectedValue + "' "
            SQL = SQL + "  And T_Step   = '" + DStep.SelectedValue + "' "
            SQL = SQL + "  And T_CompletedTime >= '" + wSDate + "' "
            SQL = SQL + "  And T_CompletedTime <= '" + wEDate + "' "
            SQL = SQL + "  And FormSno  > '2000' "
            SQL = SQL + "  And Action = '0' "

            If DSts.SelectedValue <> "ALL" Then
                SQL = SQL + " And Sts = '" + DSts.SelectedValue + "'"
            End If
            If DDivision.SelectedValue <> "ALL" Then
                SQL = SQL + " And Division = '" + DDivision.SelectedValue + "'"
            End If
            If DPerson.Text <> "�̿��" And DPerson.Text <> "" Then
                SQL = SQL + " And Person Like '%" + DPerson.Text + "%'"
            End If
            If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
                SQL = SQL + " And Buyer Like '%" + DBuyer.Text + "%'"
            End If
            SQL = SQL + " And T_SeqNo = '" + DSeqNo.SelectedValue + "'"

            Dim DBAdapter10 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter10.Fill(DBDataSet2, "WaitHandle")
            If DBDataSet2.Tables("WaitHandle").Rows.Count > 0 Then
                '�w�w�`�ɼ�
                If DBDataSet2.Tables("WaitHandle").Rows(0).Item("BWorkTime") > 0 Then
                    BSum = BSum + DBDataSet2.Tables("WaitHandle").Rows(0).Item("BWorkTime")
                End If

                '���Z�`�ɼ�
                If DBDataSet2.Tables("WaitHandle").Rows(0).Item("AWorkTime") > 0 Then
                    ASum = ASum + DBDataSet2.Tables("WaitHandle").Rows(0).Item("AWorkTime")
                End If
            End If
            OleDbConnection1.Close()
            '�w�w�`�ɼ�
            If BSum > 0 Then
                Item(1) = "�w�w"
                Real_Data = BSum / 60
                Data(1) = CInt(Real_Data.ToString("0"))
                Buffer_Data(1) = Data(1)
                Color(1) = "RED"
            End If

            '���Z�`�ɼ� 
            If ASum > 0 Then
                Item(2) = "���Z"
                Real_Data = ASum / 60
                Data(2) = CInt(Real_Data.ToString("0"))
                Buffer_Data(2) = Data(2)
                Color(2) = "BLUE"
            End If

            'BarChart
            HttpPath = ConfigurationSettings.AppSettings.Get("Http") & _
                       ConfigurationSettings.AppSettings.Get("ChartPath") & _
                       "Operation-Time-" & NowDateTime & ".bmp"

            FileName = ConfigurationSettings.AppSettings.Get("BMPPath") & "Operation-Time-" & NowDateTime & ".bmp"

            RtnCode = oChart.BarChart(FileName, Item, Data, Color, 2)   'BMP�ɦW, �R���ɦW, ����, �ƭ�, �C��, ���ؼ� 
            If RtnCode = 0 Then
                BarImage3.ImageUrl = HttpPath
                BarImage3.Visible = True
                ChartTitle3.Visible = True
                ChartUnit3.Visible = True
            End If
            '**************************************************************************************************
            For i = 1 To 10
                Item(i) = ""
                Data(i) = 0
                Color(i) = ""
            Next

            If Records <> 0 Then
                '�w�w
                Item(1) = "�w�w"
                Real_Data = Buffer_Data(1) / Records
                Data(1) = Real_Data.ToString("0.00")
                Color(1) = "RED"
                '���Z
                Item(2) = "���Z"
                Real_Data = Buffer_Data(2) / Records
                Data(2) = Real_Data.ToString("0.00")
                Color(2) = "BLUE"
            Else
                '�w�w
                Item(1) = "�w�w"
                Color(1) = "RED"
                '�w�w
                Item(2) = "���Z"
                Color(2) = "BLUE"
            End If
            'BarChart
            HttpPath = ConfigurationSettings.AppSettings.Get("Http") & _
                       ConfigurationSettings.AppSettings.Get("ChartPath") & _
                       "Operation-Time-Ratio-" & NowDateTime & ".bmp"

            FileName = ConfigurationSettings.AppSettings.Get("BMPPath") & "Operation-Time-Ratio-" & NowDateTime & ".bmp"

            RtnCode = oChart.BarChart(FileName, Item, Data, Color, 2)   'BMP�ɦW, �R���ɦW, ����, �ƭ�, �C��, ���ؼ� 
            If RtnCode = 0 Then
                BarImage4.ImageUrl = HttpPath
                BarImage4.Visible = True
                ChartUnit4.Visible = True
                ChartTitle4.Visible = True
            End If

            '**************************************************************************************************
            DBDataSet2.Clear()
            SQL = "SELECT No, T_FormName as FormName, "
            SQL = SQL + "Case Sts When 0 Then '�}�o��' When 1 Then 'OK' When 2 Then 'NG' Else '����' End As Sts, "
            SQL = SQL + "Convert(VarChar, T_BStartTime, 20) + '~' + Convert(VarChar, T_BEndTime, 20) As BTime, "
            SQL = SQL + "BWorkTime, "
            SQL = SQL + "Convert(VarChar, T_AStartTime, 20) + '~' + Convert(VarChar, T_AEndTime, 20) As ATime, "
            SQL = SQL + "AWorkTime, "
            SQL = SQL + "BWorkTime-AWorkTime As Balance, "
            SQL = SQL + "Division, "
            If DFormName.SelectedValue = "000001" Then
                SQL = SQL + "Buyer, "
                SQL = SQL + "Case HalfFinish When 'Yes' Then '�b���~' Else '���~' End As Field1, "
                SQL = SQL + "Suppiler As Field2, "
                SQL = SQL + "Level    As Field3 "
            End If
            If DFormName.SelectedValue = "000002" Then
                SQL = SQL + "Buyer, "
                SQL = SQL + "'' As Field1, "
                SQL = SQL + "Suppiler As Field2, "
                SQL = SQL + "Level    As Field3 "
            End If
            If DFormName.SelectedValue = "000003" Or DFormName.SelectedValue = "000007" Then
                SQL = SQL + "Buyer, "
                SQL = SQL + "SliderType2 As Field1, "
                SQL = SQL + "Suppiler As Field2, "
                SQL = SQL + "Level    As Field3 "
            End If
            If DFormName.SelectedValue = "000005" Or DFormName.SelectedValue = "000009" Or DFormName.SelectedValue = "000010" Or DFormName.SelectedValue = "000013" Then
                SQL = SQL + "'' As Buyer, "
                SQL = SQL + "'' As Field1, "
                SQL = SQL + "'' As Field2, "
                SQL = SQL + "'' As Field3 "
            End If
            If DFormName.SelectedValue = "000011" Or DFormName.SelectedValue = "000012" Then
                SQL = SQL + "Buyer, "
                SQL = SQL + "'' As Field1, "
                SQL = SQL + "'' As Field2, "
                SQL = SQL + "'' As Field3 "
            End If

            SQL = SQL + "FROM " + wTableName + " "
            SQL = SQL + "Where T_Active = '0' "
            SQL = SQL + "  And T_FormNo = '" + DFormName.SelectedValue + "' "
            SQL = SQL + "  And T_Step   = '" + DStep.SelectedValue + "' "
            SQL = SQL + "  And T_CompletedTime >= '" + wSDate + "' "
            SQL = SQL + "  And T_CompletedTime <= '" + wEDate + "' "
            SQL = SQL + "  And FormSno  > '2000' "
            SQL = SQL + "  And Action = '0' "

            If DSts.SelectedValue <> "ALL" Then
                SQL = SQL + " And Sts = '" + DSts.SelectedValue + "'"
            End If
            If DDivision.SelectedValue <> "ALL" Then
                SQL = SQL + " And Division = '" + DDivision.SelectedValue + "'"
            End If
            If DPerson.Text <> "�̿��" And DPerson.Text <> "" Then
                SQL = SQL + " And Person Like '%" + DPerson.Text + "%'"
            End If
            If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
                SQL = SQL + " And Buyer Like '%" + DBuyer.Text + "%'"
            End If
            SQL = SQL + " And T_SeqNo = '" + DSeqNo.SelectedValue + "'"

            SQL = SQL + " Order by " + wTableName + ".T_CompletedTime "

            OleDbConnection1.Open()
            Dim DBAdapter11 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter11.Fill(DBDataSet2, "WaitHandle")
            DataGrid1.DataSource = DBDataSet2
            DataGrid1.DataBind()
            OleDbConnection1.Close()

            DataGrid1.Visible = True

        End If  'TableName

    End Sub

    Private Sub DFormName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged
        pFormNo = DFormName.SelectedValue
        BarImage1.Visible = False   '��ƦX�p
        BarImage2.Visible = False   '��Ƥ�v
        BarImage3.Visible = False   '�w�w���Z�X�p
        BarImage4.Visible = False   '�w�w���Z��v

        ChartTitle1.Visible = False
        ChartTitle2.Visible = False
        ChartTitle3.Visible = False
        ChartTitle4.Visible = False
        ChartUnit1.Visible = False
        ChartUnit2.Visible = False
        ChartUnit3.Visible = False
        ChartUnit4.Visible = False

        DataGrid1.Visible = False
        SetSearchItem()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        pFormNo = DFormName.SelectedValue
        DMsgBox.Visible = True
        DMsgBox.Text = "�s�@�έp�Ϥ�...�Э@�ߵ���"

        DataList()
    End Sub

End Class
