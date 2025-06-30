Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class MapSheetMod_02
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DMapSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DModContent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DModReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBefFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBefFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOriFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBefMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOriMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents LRefAttach As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents LMapFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMapSheet3 As System.Web.UI.WebControls.Image
    Protected WithEvents DOriFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents DModReasonCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMakeMap As System.Web.UI.WebControls.TextBox
    Protected WithEvents DLevel As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSample As System.Web.UI.WebControls.TextBox
    Protected WithEvents LOriMap As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LBefMap As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DManufType As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCpsc As System.Web.UI.WebControls.TextBox
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
        Response.Cookies("PGM").Value = "MapModSheet_02.aspx"

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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("MapModFilePath")
        Dim PathOld1 As String = ConfigurationSettings.AppSettings.Get("HttpOld")
        Dim PathOld2 As String = ConfigurationSettings.AppSettings.Get("MapModFilePath")
        Dim RtnCode As Integer = 0


        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_MapModSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_MapModSheet")
        If DBDataSet1.Tables("F_MapModSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("No")         'No
            DDate.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Date")                   '���
            DBuyer.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Buyer")                 'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("SellVendor")       '�e�U�t��
            DDivision.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Division")           '����
            DSample.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Sample")           'sample
            DPerson.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Person")               '���
            DCpsc.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Cpsc")               'Cpsc
            DOriMapNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriMapNo")        '��l�ϸ�
            DOriFormNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormNo")      '��l��渹�X
            DOriFormSno.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormSno")    '��l�渹
            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriMapNo") <> "" And _
               DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormNo") <> "" And _
               DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormSno") > 0 Then          '��ϸ�
                LOriMap.NavigateUrl = "MapSheet_02.aspx?pFormNo=" & DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormNo") & _
                                                      "&pFormSno=" & CStr(DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormSno"))
                LOriMap.Visible = True
            Else
                LOriMap.Visible = False
            End If

            DLevel.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Level")              '������
            DBefMapNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefMapNo")        '�e�ϸ�
            DBefFormNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormNo")      '�e��渹�X
            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormSno") > 0 Then
                DBefFormSno.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormSno")    '�e�渹
            Else
                DBefFormSno.Text = ""
            End If
            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefMapNo") <> "" And _
               DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormNo") <> "" And _
               DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormSno") > 0 Then            '��ϸ�
                LBefMap.NavigateUrl = "MapModSheet_02.aspx?pFormNo=" & DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormNo") & _
                                                         "&pFormSno=" & CStr(DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormSno"))
            Else
                LBefMap.Visible = False
            End If

            DModReasonCode.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("ModReasonCode")  '��]���O�N�X
            DModReasonDesc.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("ModReasonDesc")  '��]����
            DModContent.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("ModContent")        '�ק�Ӹ`����

            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("RefAttach") <> "" Then              '�ѦҪ���
                LRefAttach.NavigateUrl = Path & DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("RefAttach")
            Else
                LRefAttach.Visible = False
            End If

            DMapNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("MapNo")                     '�ϸ�
            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("MapFile") <> "" Then                   '����
                '�ʦs���٭�  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("MapFile"))

                LMapFile.NavigateUrl = Path & DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("MapFile")
            Else
                LMapFile.Visible = False
            End If

            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("PdfFile") <> "" Then                   '����
                '�ʦs���٭�  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("PdfFile"))

                LPdfFile.NavigateUrl = Path & DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("PdfFile")
            Else
                LPdfFile.Visible = False
            End If

            DMakeMap.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("MakeMap")                '�s�Ϫ�
            DManufType.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("ManufType")            '���~�`
            DSuppiler.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Suppiler")     '�~�`��
        End If

        DFormSno.Text = "�渹�G" & CStr(wFormSno)
        'DB�s������
        OleDbConnection1.Close()
    End Sub

End Class
