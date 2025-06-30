Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class MapSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DLevel As System.Web.UI.WebControls.TextBox
    Protected WithEvents LMapFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LRefMapFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSurface As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBackground As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCramper As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDescription As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMapReqDelDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMapSheet3 As System.Web.UI.WebControls.Image
    Protected WithEvents DMaterial As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMaterialDetail As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFrontBack As System.Web.UI.WebControls.TextBox
    Protected WithEvents DLight As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMapSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DCPSC As System.Web.UI.WebControls.TextBox
    Protected WithEvents DHalfFinish As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMakeMap As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSample As System.Web.UI.WebControls.TextBox
    Protected WithEvents Image1 As System.Web.UI.WebControls.Image
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DManufType As System.Web.UI.WebControls.TextBox
    Protected WithEvents Panel1 As System.Web.UI.WebControls.Panel
    Protected WithEvents DStatus As System.Web.UI.WebControls.TextBox
    Protected WithEvents LMMap As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DSuppiler As System.Web.UI.WebControls.TextBox
    Protected WithEvents LPdfFile As System.Web.UI.WebControls.HyperLink

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

    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "MapSheet_02.aspx"

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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("MapFilePath")
        Dim PathOld1 As String = ConfigurationSettings.AppSettings.Get("HttpOld")
        Dim PathOld2 As String = ConfigurationSettings.AppSettings.Get("MapFilePath")
        Dim RtnCode As Integer = 0

        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_MapSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_MapSheet")
        If DBDataSet1.Tables("F_MapSheet").Rows.Count > 0 Then
            If DBDataSet1.Tables("F_MapSheet").Rows(0).Item("UPDSts") = 1 Then          '�ק�u�{���A
                DStatus.Text = "�ק�u�{�i�椤"
                DStatus.Visible = True
                Panel1.Visible = True
            Else
                DStatus.Visible = False
                Panel1.Visible = False
            End If

            If DBDataSet1.Tables("F_MapSheet").Rows(0).Item("ModMap") = 1 Then          '�ק�u�{���A
                LMMap.NavigateUrl = "MapList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno & "&pFun=MAP"
                LMMap.Visible = True
            Else
                LMMap.Visible = False
            End If

            DNo.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("No")         'No
            DDate.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Date")                   '���
            DBuyer.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Buyer")                 'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("SellVendor")       '�e�U�t��
            DDivision.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Division")           '����
            DPerson.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Person")               '���
            DBackground.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Background")       '�}�o�I��
            DSpec.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Spec")                   'Spec
            DCramper.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Cramper")             'Cramper
            DSurface.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Surface")             '���B�z
            DFrontBack.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("FrontBack")         'Puller--���ϭ�
            DHalfFinish.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("HalfFinish")       '�b���~
            DMaterial.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Material")           '����
            DMaterialDetail.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("MaterialDetail") '����Ӷ�
            If DBDataSet1.Tables("F_MapSheet").Rows(0).Item("RefMapFile") <> "" Then


                '�ʦs���٭�  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_MapSheet").Rows(0).Item("RefMapFile"))

                LRefMapFile.NavigateUrl = Path & DBDataSet1.Tables("F_MapSheet").Rows(0).Item("RefMapFile")   '��ϸ�
            Else
                LRefMapFile.Visible = False
            End If
            DDescription.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Description")     '�Ƶ�
            DManufType.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("ManufType")     '���~�s
            DSuppiler.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Suppiler")     '�~�`��

            DCPSC.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("CPSC")                   'CPSC
            DMapReqDelDate.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("MapReqDelDate") '�ϭ��Ʊ���
            DLight.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Light")                 '���y��
            DSample.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Sample")               '�˫~
            DMapNo.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("MapNo")                 '�ϸ�
            If DBDataSet1.Tables("F_MapSheet").Rows(0).Item("MapFile") <> "" Then

                '�ʦs���٭�  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_MapSheet").Rows(0).Item("MapFile"))
                LMapFile.NavigateUrl = Path & DBDataSet1.Tables("F_MapSheet").Rows(0).Item("MapFile")  '����
            Else
                LMapFile.Visible = False
            End If

            If DBDataSet1.Tables("F_MapSheet").Rows(0).Item("PdfFile") <> "" Then

                '�ʦs���٭�  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_MapSheet").Rows(0).Item("pdfFile"))
                LPdfFile.NavigateUrl = Path & DBDataSet1.Tables("F_MapSheet").Rows(0).Item("PdfFile")  '����
            Else
                LPdfFile.Visible = False
            End If


            DLevel.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Level")     '������
            DMakeMap.Text = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("MakeMap") '�s�Ϫ�
        End If

        DFormSno.Text = "�渹�G" & CStr(wFormSno)
        'DB�s������
        OleDbConnection1.Close()
    End Sub

End Class
