Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class SCD_SampleSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DOP6 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCDITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCSITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTDRITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTSRITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTNRITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCNITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTDLITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTSLITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTNLITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHER As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTHCOL As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCCOL As System.Web.UI.WebControls.TextBox
    Protected WithEvents DECOL As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTACOL As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDEVPRD As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDEVNO As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTAWIDTH As System.Web.UI.WebControls.TextBox
    Protected WithEvents LSAMPLEFILE As System.Web.UI.WebControls.Image
    Protected WithEvents DCODENO As System.Web.UI.WebControls.TextBox
    Protected WithEvents DITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSIZENO As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTALINE As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DSampleSheet2 As System.Web.UI.WebControls.Image
    Protected WithEvents DDATE As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAPPBUYER As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents WF5Name As System.Web.UI.WebControls.TextBox
    Protected WithEvents WF6Name As System.Web.UI.WebControls.TextBox
    Protected WithEvents WF7Name As System.Web.UI.WebControls.TextBox
    Protected WithEvents WF4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents WF3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents WF2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF4Name As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF3Name As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF7 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF6 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF7Name As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF6Name As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF5Name As System.Web.UI.WebControls.TextBox
    Protected WithEvents DO1Item As System.Web.UI.WebControls.TextBox
    Protected WithEvents DO2Item As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOther1 As System.Web.UI.WebControls.Label
    Protected WithEvents DOther2 As System.Web.UI.WebControls.Label
    Protected WithEvents LQCFILE1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFILE4 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFILE3 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFILE2 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFILE5 As System.Web.UI.WebControls.HyperLink

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
    Dim wFormNo As String           '��渹�X
    Dim wFormSno As Integer         '���y����
    Dim NowDateTime As String       '�{�b����ɶ�

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "SCD_SampleSheet_02.aspx"

        SetParameter()          '�]�w�@�ΰѼ�

        If Not Me.IsPostBack Then   '���OPostBack
            ShowFormData()      '��ܪ����
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�@�ΰѼ�
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '�{�b���
        wFormNo = Request.QueryString("pFormNo")    '��渹�X
        wFormSno = Request.QueryString("pFormSno")  '���y����
    End Sub


    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("SampleFilePath")
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_SampleSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_SampleSheet")
        If DBDataSet1.Tables("F_SampleSheet").Rows.Count > 0 Then
            '�����
            DNo.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("No")                     'No
            DAPPBUYER.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("AppBuyer")         'Customer
            DDATE.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Date")                 '�o���
            DSIZENO.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("SizeNo")             'Size
            DITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Item")                 'Item
            DCODENO.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CodeNo")             'Code No
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("SampleFile") <> "" Then          '��ڼ˫~
                LSAMPLEFILE.ImageUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("SampleFile")
            Else
                LSAMPLEFILE.Visible = False
            End If
            DTAWIDTH.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TAWidth")           '���a�e��
            DDEVNO.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("DevNo")               '�}�oNo
            DDEVPRD.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("DevPrd")             '�}�o����
            DTACOL.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TACol")               '���aColor
            DTALINE.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TALine")             '�����uColor
            DECOL.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("ECol")                 '�Ⱦ�
            DCCOL.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CCol")                 '�Y��
            DTHCOL.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("THCol")               '�_�u�u
            DOTHER.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Other")               '��L

            '980410 update by alin
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile1") <> "" Then              '�~�����i1
                LQCFILE1.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile1")
            Else
                LQCFILE1.Visible = False
            End If
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile2") <> "" Then              '�~�����i2
                LQCFILE2.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile2")
            Else
                LQCFILE2.Visible = False
            End If
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile3") <> "" Then              '�~�����i3
                LQCFILE3.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile3")
            Else
                LQCFILE3.Visible = False
            End If
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile4") <> "" Then              '�~�����i4
                LQCFILE4.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile4")
            Else
                LQCFILE4.Visible = False
            End If
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile5") <> "" Then              '�~�����i5
                LQCFILE5.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile5")
            Else
                LQCFILE5.Visible = False
            End If

            DOther1.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Other1")           'Other1
            DOther2.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("other2")           'Other2
            DO1Item.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("O1Item")           'O1Item
            DO2Item.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("O2item")           'O2Item
            '=====================


            DTNLITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TNLItem")           'TAPE NAT-��
            DTNRITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TNRItem")           'TAPE NAT-�k
            DTSLITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TSLItem")           'TAPE SET-��
            DTSRITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TSRItem")           'TAPE SET-�k
            DTDLITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TDLItem")           'TAPE DYED-��
            DTDRITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TDRItem")           'TAPE DYED-�k
            DCNITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CNItem")             'CHAIN NAT
            DCSITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CSItem")             'CHAIN SET
            DCDITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CDItem")             'CHAIN DYED
            DCITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CItem")               'CORD
            DOP1.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP1")                   '�u�{1
            DOP2.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP2")                   '�u�{2
            DOP3.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP3")                   '�u�{3
            DOP4.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP4")                   '�u�{4
            DOP5.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP5")                   '�u�{5
            DOP6.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP6")                   '�u�{6
            DWF1.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF1")                   '�ӻ{WF1
            DWF2.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF2")                   '�ӻ{WF2
            DWF3.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF3")                   '�ӻ{WF3
            DWF4.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF4")                   '�ӻ{WF4
            DWF5.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF5")                   '�ӻ{WF5
            DWF6.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF6")                   '�ӻ{WF6
            DWF7.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF7")                   '�ӻ{WF7
            DWF3Name.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF3Name")           '�ӻ{�̳���WF3
            DWF4Name.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF4Name")           '�ӻ{�̳���WF4
            DWF5Name.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF5Name")           '�ӻ{�̳���WF5
            DWF6Name.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF6Name")           '�ӻ{�̳���WF6
            DWF7Name.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF7Name")           '�ӻ{�̳���WF7
        End If
        'DB�s������
        OleDbConnection1.Close()
    End Sub

End Class
