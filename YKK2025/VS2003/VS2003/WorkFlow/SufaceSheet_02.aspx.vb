Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class SufaceSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DOrderTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCResult3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCResult2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCRemark3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCRemark2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCDate3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCDate2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCRemark1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCResult1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCDate1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEADesc1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEACheck1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck12 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck11 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck10 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck9 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck8 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck5 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck4 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LFinalSampleFile As System.Web.UI.WebControls.Image
    Protected WithEvents DEnglishName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBFinalDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAllowSample As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQty As System.Web.UI.WebControls.TextBox
    Protected WithEvents DColor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DManufType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSliderSample As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReqQty As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReqDelDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDevReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSufaceSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DSuppiler As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBuyer As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LQCReqFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOPManualFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LEACheckFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFinalFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents LContactFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LForcastFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LManufFlowFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAttachSample As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DORNO As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSufaceSheet2 As System.Web.UI.WebControls.Image
    Protected WithEvents LCustSampleFile As System.Web.UI.WebControls.Image
    Protected WithEvents Panel1 As System.Web.UI.WebControls.Panel
    Protected WithEvents DStatus As System.Web.UI.WebControls.TextBox
    Protected WithEvents LSuface As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSchedule As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCap As System.Web.UI.WebControls.TextBox
    Protected WithEvents LEACheckFile1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DQCCheck14 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck13 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DLOSS As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCCheck15 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DYearSeason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCCheck16 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LPFASFile As System.Web.UI.WebControls.HyperLink

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
        Response.Cookies("PGM").Value = "SufaceSheet_02.aspx"

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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("SufaceFilePath")
        Dim PathOld1 As String = ConfigurationSettings.AppSettings.Get("HttpOld")
        Dim PathOld2 As String = ConfigurationSettings.AppSettings.Get("SufaceFilePath")
        Dim RtnCode As Integer = 0
        Dim i As Integer
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_SufaceSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_SufaceSheet")
        If DBDataSet1.Tables("F_SufaceSheet").Rows.Count > 0 Then
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("UPDSts") = 1 Then          '�l�[�u�{���A
                DStatus.Text = "�l�[�u�{�i�椤"
                DStatus.Visible = True
                Panel1.Visible = True
            Else
                DStatus.Visible = False
                Panel1.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Suface") = 1 Then          '�l�[���B�z
                LSuface.NavigateUrl = "SufaceList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno & "&pFun=Suface"
                LSuface.Visible = True
            Else
                LSuface.Visible = False
            End If
            '------------------------------------------------------------------------------------------------
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("CustSampleFile") <> "" Then        '�Ȥ�˫~��
                '�ʦs���٭�  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("CustSampleFile"))

                LCustSampleFile.ImageUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("CustSampleFile")
            Else
                LCustSampleFile.Visible = False
            End If
            DNo.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Date")                   '���
            SetFieldData("Division", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Division"))  '����
            SetFieldData("Person", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Person"))      '���
            DSpec.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Spec")                   '�W��
            SetFieldData("Buyer", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Buyer"))        'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("SellVendor")       '�e�U�t��
            DReqDelDate.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ReqDelDate")       '�Ʊ���
            DReqQty.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ReqQty")               '�w���q
            DSliderSample.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("SliderSample")   '�˫~���Y
            SetFieldData("AttachSample", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("AttachSample"))    '����
            DORNO.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ORNO")                   'OR-NO
            DOrderTime.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("OrderTime")         '�U��ɶ�
            DPrice.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Price")                 '���
            DDevReason.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("DevReason")         '�}�o�z��
            DYearSeason.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("YearSeason")         '�~�u

            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FinalSampleFile") <> "" Then       '�̲׼˫~��
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FinalSampleFile"))
                LFinalSampleFile.ImageUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FinalSampleFile")
            Else
                LFinalSampleFile.Visible = False
            End If
            SetFieldData("ManufType", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ManufType"))  '���s/�~�`
            SetFieldData("Suppiler", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Suppiler"))    '�~�`��
            DColor.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Color")                   'Color
            DQty.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Qty")                       '�ƶq
            DCap.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Cap")                   '�鲣��
            DSchedule.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Schedule")                       '��Ǥ�{
            DFReason.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FReason")                       '�z��
            SetFieldData("AllowSample", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("AllowSample"))    '���׼˫~
            DBFinalDate.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("BFinalDate")        '�w�w������
            DCode.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Code")                    'Code
            DEnglishName.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EnglishName")      '�^��W��
            DLOSS.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("LOSS")                    'LOSS
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCReqFile") <> "" Then              '�~��̿��
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCReqFile"))
                LQCReqFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCReqFile")
            Else
                LQCReqFile.Visible = False
            End If

            SetFieldData("QCCheck1", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck1"))   '�f�|�o�k
            SetFieldData("QCCheck2", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck2"))   '�P�ʩ��
            SetFieldData("QCCheck3", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck3"))   'LOCK�j��
            SetFieldData("QCCheck4", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck4"))   '90�ױj��
            SetFieldData("QCCheck5", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck5"))   '��O
            SetFieldData("QCCheck15", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck15"))   'N-ANTI
            SetFieldData("QCCheck16", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck16"))   'PFAS


            '  If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck6File") <> "" Then           '�q�ὤ�p
            '  LQCCheck6File.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck6File")
            'Else
            'LQCCheck6File.Visible = False
            'End If
            SetFieldData("QCCheck7", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck7"))   '�˰w
            SetFieldData("QCCheck8", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck8"))   'AATCC
            SetFieldData("QCCheck9", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck9"))   '���~
            SetFieldData("QCCheck10", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck10")) '�Q���Q��
            SetFieldData("QCCheck11", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck11")) '�@���K��
            SetFieldData("QCCheck12", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck12")) '�G���K��
            SetFieldData("QCCheck13", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck13")) 'Oeko-tex
            SetFieldData("QCCheck14", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck14")) 'A01
            SetFieldData("EACheck1", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheck1"))   '���`����
            DEADesc1.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EADesc1")              '���`����Ƶ�

            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile") <> "" Then            'Oeko-tex���`������i
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile"))
                LEACheckFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile")
            Else
                LEACheckFile.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile1") <> "" Then           'A01���`������i
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile1"))
                LEACheckFile1.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile1")
            Else
                LEACheckFile1.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCFinalFile") <> "" Then            '���ճ��i��
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCFinalFile"))
                LQCFinalFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCFinalFile")
            Else
                LQCFinalFile.Visible = False
            End If

            DQCDate1.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCDate1")               '���-1
            SetFieldData("QCResult1", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCResult1"))  '�˴����G-1
            DQCRemark1.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCRemark1")           '�Ƶ�-1
            DQCDate2.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCDate2")               '���-2
            SetFieldData("QCResult2", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCResult2"))  '�˴����G-2
            DQCRemark2.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCRemark2")           '�Ƶ�-2
            DQCDate3.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCDate3")               '���-3
            SetFieldData("QCResult3", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCResult3"))  '�˴����G-3
            DQCRemark3.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCRemark3")           '�Ƶ�-3

            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ManufFlowFile") <> "" Then           '�s�y�y�{��
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ManufFlowFile"))
                LManufFlowFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ManufFlowFile")
            Else
                LManufFlowFile.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ForcastFile") <> "" Then             '������
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ForcastFile"))
                LForcastFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ForcastFile")
            Else
                LForcastFile.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("OPManualFile") <> "" Then            '�@�~�зǮ�
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("OPManualFile"))
                LOPManualFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("OPManualFile")
            Else
                LOPManualFile.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ContactFile") <> "" Then             '������
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ContactFile"))
                LContactFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ContactFile")
            Else
                LContactFile.Visible = False
            End If

            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("PFASFile") <> "" Then             '???
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("PFASFile"))
                LPFASFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("PFASFile")
            Else
                LPFASFile.Visible = False
            End If
            '------------------------------------------------------------------------------------------------
        End If
        DFormSno.Text = "�渹�G" & CStr(wFormSno)
        'DB�s������
        OleDbConnection1.Close()
    End Sub

    '*****************************************************************
    '**(ShowSheetField)
    '**     �ظm�U�Ԧ�����l��
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        '����
        If pFieldName = "Division" Then
            DDivision.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DDivision.Items.Add(ListItem1)
        End If
        '���
        If pFieldName = "Person" Then
            DPerson.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DPerson.Items.Add(ListItem1)
        End If

        'Buyer
        If pFieldName = "Buyer" Then
            DBuyer.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DBuyer.Items.Add(ListItem1)
        End If

        '����
        If pFieldName = "AttachSample" Then
            DAttachSample.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAttachSample.Items.Add(ListItem1)
        End If

        '���s/�~�`
        If pFieldName = "ManufType" Then
            DManufType.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DManufType.Items.Add(ListItem1)
        End If

        'Suppiler
        If pFieldName = "Suppiler" Then
            DSuppiler.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSuppiler.Items.Add(ListItem1)
        End If

        '���׼˫~
        If pFieldName = "AllowSample" Then
            DAllowSample.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAllowSample.Items.Add(ListItem1)
        End If

        '�f�|�o�k
        If pFieldName = "QCCheck1" Then
            DQCCheck1.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck1.Items.Add(ListItem1)
        End If

        '�P�ʩ��
        If pFieldName = "QCCheck2" Then
            DQCCheck2.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck2.Items.Add(ListItem1)
        End If

        'LOCK�j��
        If pFieldName = "QCCheck3" Then
            DQCCheck3.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck3.Items.Add(ListItem1)
        End If

        '90�ױj��
        If pFieldName = "QCCheck4" Then
            DQCCheck4.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck4.Items.Add(ListItem1)
        End If

        '��O
        If pFieldName = "QCCheck5" Then
            DQCCheck5.Items.Clear()
            Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck5.Items.Add(ListItem1)
        End If

        'N-ANTI
        If pFieldName = "QCCheck15" Then
            DQCCheck15.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck15.Items.Add(ListItem1)
        End If



        'PFAS
        If pFieldName = "QCCheck16" Then
            DQCCheck16.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck16.Items.Add(ListItem1)
        End If




        '�˰w
        If pFieldName = "QCCheck7" Then
            DQCCheck7.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck7.Items.Add(ListItem1)
        End If

        'AATCC
        If pFieldName = "QCCheck8" Then
            DQCCheck8.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck8.Items.Add(ListItem1)
        End If

        '���~
        If pFieldName = "QCCheck9" Then
            DQCCheck9.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck9.Items.Add(ListItem1)
        End If

        '�Q���Q��
        If pFieldName = "QCCheck10" Then
            DQCCheck10.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck10.Items.Add(ListItem1)
        End If

        '�@���K��
        If pFieldName = "QCCheck11" Then
            DQCCheck11.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck11.Items.Add(ListItem1)
        End If

        '�G���K��
        If pFieldName = "QCCheck12" Then
            '�G���K��
            DQCCheck12.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck12.Items.Add(ListItem1)
        End If

        'Oeko-tex
        If pFieldName = "QCCheck13" Then
            DQCCheck13.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck13.Items.Add(ListItem1)
        End If

        'A01
        If pFieldName = "QCCheck14" Then
            DQCCheck14.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCCheck14.Items.Add(ListItem1)
        End If

        '���`����
        If pFieldName = "EACheck1" Then
            DEACheck1.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DEACheck1.Items.Add(ListItem1)
        End If

        '�˴����G-1
        If pFieldName = "QCResult1" Then
            DQCResult1.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCResult1.Items.Add(ListItem1)
        End If

        '�˴����G-2
        If pFieldName = "QCResult2" Then
            DQCResult2.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCResult2.Items.Add(ListItem1)
        End If

        '�˴����G-3
        If pFieldName = "QCResult3" Then
            DQCResult3.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCResult3.Items.Add(ListItem1)
        End If

        '������
        'If pFieldName = "Level" Then
        'DLevel.Items.Clear()
        'If idx = 0 Then
        'If pName <> "ZZZZZZ" Then
        'Dim ListItem1 As New ListItem
        'ListItem1.Text = pName
        'ListItem1.Value = pName
        'DLevel.Items.Add(ListItem1)
        'End If
        'Else
        'SQL = "Select * From M_Referp Where Cat='007' and (DKey like 'In%' or DKey = '') Order by DKey, Data "
        'Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter1.Fill(DBDataSet1, "M_Referp")
        'DBTable1 = DBDataSet1.Tables("M_Referp")
        'For i = 0 To DBTable1.Rows.Count - 1
        'Dim ListItem1 As New ListItem
        'ListItem1.Text = DBTable1.Rows(i).Item("Data")
        'ListItem1.Value = DBTable1.Rows(i).Item("Data")
        'If ListItem1.Value = pName Then ListItem1.Selected = True
        'DLevel.Items.Add(ListItem1)
        'Next
        'End If
        'End If
    End Sub
End Class
