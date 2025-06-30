Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ModManufOutSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DManuaInSheet2 As System.Web.UI.WebControls.Image
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSave As System.Web.UI.WebControls.Button
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderGRCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSpec As System.Web.UI.WebControls.Button
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents BOMapNo As System.Web.UI.WebControls.Button
    Protected WithEvents LMapFile1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LMapFile2 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DLevel As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAssembler As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderType1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSliderType2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DManufPlace As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMaterialDetail As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMaterial As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMaterialOther As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents LConfirmFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LAuthorizeFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DDevReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSliderCode As System.Web.UI.WebControls.Button
    Protected WithEvents BSample As System.Web.UI.WebControls.Button
    Protected WithEvents BPrice As System.Web.UI.WebControls.Button
    Protected WithEvents DArMoldFee As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPurMold As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPullerPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMoldName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMoldQty As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMoldPoint As System.Web.UI.WebControls.TextBox
    Protected WithEvents BQuality1 As System.Web.UI.WebControls.Button
    Protected WithEvents BQuality2 As System.Web.UI.WebControls.Button
    Protected WithEvents LQAAttachFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LSampleFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents BMMapNo As System.Web.UI.WebControls.Button
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSample As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQuality1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQuality2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents LSliderCode As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LSample As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LPrice As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQuality1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQuality2 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BDate As System.Web.UI.WebControls.Button
    Protected WithEvents wMapFile1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents wMapFile2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents wConfirmFile As System.Web.UI.WebControls.TextBox
    Protected WithEvents wAuthorizeFile As System.Web.UI.WebControls.TextBox
    Protected WithEvents wSampleFile As System.Web.UI.WebControls.TextBox
    Protected WithEvents wQAAttachFile As System.Web.UI.WebControls.TextBox
    Protected WithEvents DManuaInSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DSampleFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DQAAttachFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DAuthorizeFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DConfirmFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DMapFile2 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DMapFile1 As System.Web.UI.HtmlControls.HtmlInputFile

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
    Dim FieldName(44) As String     '�U���
    Dim Attribute(44) As Integer    '�U����ݩ�    
    Dim Top As Integer              '�ʺA����Top��m
    Dim wFormNo As String           '��渹�X
    Dim wFormSno As Integer         '���y����
    Dim wOFormNo As String          '���渹�X
    Dim wOFormSno As Integer        '����y����
    Dim rFormNo As String           '�Ǧ^��渹�X
    Dim rFormSno As Integer         '�Ǧ^���y����
    Dim wStep As Integer            '�u�{�N�X
    Dim NowDateTime As String       '�{�b����ɶ�

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()          '�]�w�@�ΰѼ�
        TopPosition()           '���s��RequestedField��Top��m

        If Not Me.IsPostBack Then   '���OPostBack
            ShowSheetField("New")   '��������ܤ�����J�ˬd
            ShowSheetFunction()     '���\����s���
            If wFormSno > 0 Then    '�P�_�O�_�����
                ShowFormData()      '��ܪ����
            Else
                ShowOriFormData()   '��ܭ�����
            End If
            SetPopupFunction()      '�]�w�u�X�����ƥ�
        Else
            ShowSheetField("Post")  '��������ܤ�����J�ˬd
            'ShowMessage()           '�W�Ǹ���ˬd����ܰT��
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
                      CStr(DateTime.Now.Second)       '�{�b���
        wFormNo = Request.QueryString("pFormNo")
        wFormSno = Request.QueryString("pFormSno")
        wOFormNo = Request.QueryString("pOFormNo")    '���渹�X
        wOFormSno = Request.QueryString("pOFormSno")  '����y����
        wStep = Request.QueryString("pStep")          '�u�{�N�X
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     '�W�Ǹ���ˬd����ܰT��
    '**
    '*****************************************************************
    Sub ShowMessage()
        Dim Message As String = ""
        'Check����-1
        If DMapFile1.Visible Then
            If DMapFile1.PostedFile.FileName <> "" Then
                Message = "����-1"
            End If
        End If
        'Check����-2
        If DMapFile2.Visible Then
            If DMapFile2.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "����-2"
                Else
                    Message = Message & ", " & "����-2"
                End If
            End If
        End If
        'Check�T�{��
        If DConfirmFile.Visible Then
            If DConfirmFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "�T�{��"
                Else
                    Message = Message & ", " & "�T�{��"
                End If
            End If
        End If
        'Check���v��
        If DAuthorizeFile.Visible Then
            If DAuthorizeFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "���v��"
                Else
                    Message = Message & ", " & "���v��"
                End If
            End If
        End If
        'Check�˫~��
        If DSampleFile.Visible Then
            If DSampleFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "�˫~��"
                Else
                    Message = Message & ", " & "�˫~��"
                End If
            End If
        End If
        'CheckQA����
        If DQAAttachFile.Visible Then
            If DQAAttachFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "QA����"
                Else
                    Message = Message & ", " & "QA����"
                End If
            End If
        End If

        If Message <> "" Then
            Message = "�U�C�ҳ]�w�����[�ɮױN�|�� (" & Message & ") " & ",�Э��s�]�w!"
            Response.Write(YKK.ShowMessage(Message))
        End If
    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     ���s��RequestedField��Top��m
    '**
    '*****************************************************************
    Sub TopPosition()
        Top = 1120
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetFunction)
    '**     ���\����s���
    '**
    '*****************************************************************
    Sub ShowSheetFunction()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        If wFormSno > 0 And wStep > 1 Then    '�P�_�O�_[ñ��]
            'Sheet���
            DManuaInSheet1.Visible = True   '���Sheet-1
            DManuaInSheet2.Visible = True   '���Sheet-2
            '���s��m
            BSave.Style.Add("Top", Top)    '�x�s���s
            DFormSno.Style.Add("Top", Top) '�渹
        Else
            '���s��m
            BSave.Style.Add("Top", Top)    '�x�s���s
            DFormSno.Style.Add("Top", Top) '�渹
        End If

        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '900003' "
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Flow")
        If DBDataSet1.Tables("M_Flow").Rows.Count > 0 Then
            '�q�lñ�����ϥ�
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("SignImage") = 1 Then
            Else
            End If
            '���[�ɮץ��ϥ�(��FormField���]�w)
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("Attach") = 1 Then
                'DRefMapFile.Visible = True
                'DMapFile.Visible = True
            Else
                'DRefMapFile.Visible = False
                'DMapFile.Visible = False
            End If
            '�x�s���s
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("SaveFun") = 1 Then
                BSave.Visible = True
            Else
                BSave.Visible = False
            End If
        End If
        'DB�s������
        OleDbConnection1.Close()
    End Sub

    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String
        If wOFormNo = "000003" Then
            Path = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ManufInFilePath")
        Else
            Path = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ManufOutFilePath")
        End If

        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_ModManufSheet "
        SQL = SQL & " Where Sts = 0 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ModManufSheet")
        If DBDataSet1.Tables("F_ModManufSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Date")                   '���
            SetFieldData("Division", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Division"))  '����
            SetFieldData("Person", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Person"))      '���
            DSliderGRCode.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderGRCode")   'Slider Group Code
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderCode") = 1 Then              'Slider Code
                Dim DBTable1 As DataTable
                SQL = "Select * From F_SliderList "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "SliderList")
                DBTable1 = DBDataSet1.Tables("SliderList")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & DBTable1.Rows(i).Item("SliderCode") & "]"
                    If DSliderCode.Text = "" Then
                        DSliderCode.Text = Str
                    Else
                        DSliderCode.Text = DSliderCode.Text + ";" + Str
                    End If
                Next
                LSliderCode.NavigateUrl = "SliderCodeList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno
            Else
                LSliderCode.Visible = False
            End If

            DSpec.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Spec")               '�W��(Size,ChainType,���骺���X)
            DMapNo.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapNo")             '�ϸ�
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile1") <> "" Then          '����1
                LMapFile1.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile1")
            Else
                LMapFile1.Visible = False
            End If
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile2") <> "" Then          '����2 
                LMapFile2.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile2")
            Else
                LMapFile2.Visible = False
            End If
            SetFieldData("Level", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Level"))                '������
            DAssembler.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Assembler")                 '�ե�
            SetFieldData("SliderType1", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderType1"))    '���Y����1
            SetFieldData("SliderType2", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderType2"))    '���Y����2
            SetFieldData("ManufPlace", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ManufPlace"))      '�Ͳ��a
            SetFieldData("Material", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Material"))              '����
            SetFieldData("MaterialDetail", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MaterialDetail"))  '����Ӷ�
            DMaterialOther.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MaterialOther")             '�����L
            DBuyer.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Buyer")             'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SellVendor")   '�e�U�t��
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ConfirmFile") <> "" Then       '�T�{��
                LConfirmFile.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ConfirmFile")
            Else
                LConfirmFile.Visible = False
            End If
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("AuthorizeFile") <> "" Then     '���v��
                LAuthorizeFile.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("AuthorizeFile")
            Else
                LAuthorizeFile.Visible = False
            End If
            DDevReason.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("DevReason")     '�}�o�z��

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Sample") = 1 Then              '�˫~
                Dim DBTable1 As DataTable
                SQL = "Select * From F_SampleList "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "SampleList")
                DBTable1 = DBDataSet1.Tables("SampleList")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & DBTable1.Rows(i).Item("Size") & "," & DBTable1.Rows(i).Item("ChainType") & "," & DBTable1.Rows(i).Item("Class") & "," & _
                                DBTable1.Rows(i).Item("Class") & "," & DBTable1.Rows(i).Item("Qty") & "]"
                    If DSample.Text = "" Then
                        DSample.Text = Str
                    Else
                        DSample.Text = DSample.Text + ";" + Str
                    End If
                Next
                LSample.NavigateUrl = "SampleList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno
            Else
                LSample.Visible = False
            End If

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SampleFile") <> "" Then        '�˫~��
                LSampleFile.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SampleFile")
            Else
                LSampleFile.Visible = False
            End If

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Price") = 1 Then               '���
                Dim DBTable1 As DataTable
                SQL = "Select * From F_PriceList "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "PriceList")
                DBTable1 = DBDataSet1.Tables("PriceList")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & DBTable1.Rows(i).Item("Size") & "," & DBTable1.Rows(i).Item("ChainType") & "," & DBTable1.Rows(i).Item("Class") & "," & _
                                DBTable1.Rows(i).Item("Price") & "]"
                    If DPrice.Text = "" Then
                        DPrice.Text = Str
                    Else
                        DPrice.Text = DPrice.Text + ";" + Str
                    End If
                Next
                LPrice.NavigateUrl = "PriceList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno
            Else
                LPrice.Visible = False
            End If

            DArMoldFee.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ArMoldFee")     '�����Ҩ�O
            DPurMold.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("PurMold")         '�Ҩ��ʤJ�O
            DPullerPrice.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("PullerPrice") '�ޤ��ʤJ��
            DMoldName.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldName")       '�Ҩ�W��
            DMoldQty.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldQty")         '�ҫ�
            DMoldPoint.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldPoint")     '�ި�

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Quality1") = 1 Then            '�~��1
                Dim DBTable1 As DataTable
                SQL = "Select * From F_QAList1 "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "QAList1")
                DBTable1 = DBDataSet1.Tables("QAList1")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & _
                            DBTable1.Rows(i).Item("Date") & "," & _
                            DBTable1.Rows(i).Item("Size") & "," & DBTable1.Rows(i).Item("ChainType") & "," & DBTable1.Rows(i).Item("Class") & "," & _
                            DBTable1.Rows(i).Item("Assembler") & "," & DBTable1.Rows(i).Item("Surface1") & "," & DBTable1.Rows(i).Item("Surface2") & "," & DBTable1.Rows(i).Item("Gentani") & "," & _
                            DBTable1.Rows(i).Item("Kyoudo") & "," & DBTable1.Rows(i).Item("Nyuryoku") & "," & DBTable1.Rows(i).Item("Kensin") & "," & _
                            DBTable1.Rows(i).Item("Water") & "," & DBTable1.Rows(i).Item("Dry") & "," & _
                            DBTable1.Rows(i).Item("Yellow") & "," & DBTable1.Rows(i).Item("Mityaku") & "," & DBTable1.Rows(i).Item("CPSC") & _
                          "]"
                    If DQuality1.Text = "" Then
                        DQuality1.Text = Str
                    Else
                        DQuality1.Text = DQuality1.Text + ";" + Str
                    End If
                Next
                LQuality1.NavigateUrl = "Quality1List.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno
            Else
                LQuality1.Visible = False
            End If

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Quality2") = 1 Then            '�~��2
                Dim DBTable1 As DataTable
                SQL = "Select * From F_QAList2 "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "QAList2")
                DBTable1 = DBDataSet1.Tables("QAList2")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & DBTable1.Rows(i).Item("Date") & "," & DBTable1.Rows(i).Item("QACheck") & "," & DBTable1.Rows(i).Item("Remark") & "]"
                    If DQuality2.Text = "" Then
                        DQuality2.Text = Str
                    Else
                        DQuality2.Text = DQuality2.Text + ";" + Str
                    End If
                Next
                LQuality2.NavigateUrl = "Quality2List.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno
            Else
                LQuality2.Visible = False
            End If

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("QAAttachFile") <> "" Then      '�~�����
                LQAAttachFile.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("QAAttachFile")
            Else
                LQAAttachFile.Visible = False
            End If

            DFormSno.Text = "�渹�G" & CStr(wFormSno) '�渹
        End If
        'DB�s������
        OleDbConnection1.Close()
    End Sub

    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowOriFormData()
        Dim Path As String
        If wOFormNo = "000003" Then
            Path = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ManufInFilePath")
        Else
            Path = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ManufOutFilePath")
        End If
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_ManufOutSheet "
        SQL = SQL & " Where Sts = 2 "
        SQL = SQL & "   And FormNo =  '" & wOFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ModManufSheet")
        If DBDataSet1.Tables("F_ModManufSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Date")                   '���
            SetFieldData("Division", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Division"))  '����
            SetFieldData("Person", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Person"))      '���
            DSliderGRCode.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderGRCode")   'Slider Group Code

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderCode") = 1 Then              'Slider Code
                Dim DBTable1 As DataTable
                SQL = "Select * From F_SliderList "
                SQL = SQL & " Where FormNo =  '" & wOFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "SliderList")
                DBTable1 = DBDataSet1.Tables("SliderList")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & DBTable1.Rows(i).Item("SliderCode") & "]"
                    If DSliderCode.Text = "" Then
                        DSliderCode.Text = Str
                    Else
                        DSliderCode.Text = DSliderCode.Text + ";" + Str
                    End If
                Next
                LSliderCode.NavigateUrl = "SliderCodeList.aspx?pFormNo=" & wOFormNo & "&pFormSno=" & wOFormSno
            Else
                LSliderCode.Visible = False
            End If

            DSpec.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Spec")               '�W��(Size,ChainType,���骺���X)
            DMapNo.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapNo")             '�ϸ�
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile1") <> "" Then          '����1
                LMapFile1.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile1")
                wMapFile1.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile1")
            Else
                LMapFile1.Visible = False
            End If
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile2") <> "" Then          '����2 
                LMapFile2.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile2")
                wMapFile2.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile2")
            Else
                LMapFile2.Visible = False
            End If
            SetFieldData("Level", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Level"))                '������
            DAssembler.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Assembler")                 '�ե�
            SetFieldData("SliderType1", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderType1"))    '���Y����1
            SetFieldData("SliderType2", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderType2"))    '���Y����2
            SetFieldData("ManufPlace", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ManufPlace"))      '�Ͳ��a
            SetFieldData("Material", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Material"))              '����
            SetFieldData("MaterialDetail", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MaterialDetail"))  '����Ӷ�
            DMaterialOther.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MaterialOther")             '�����L
            DBuyer.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Buyer")             'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SellVendor")   '�e�U�t��
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ConfirmFile") <> "" Then       '�T�{��
                LConfirmFile.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ConfirmFile")
                wConfirmFile.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ConfirmFile")
            Else
                LConfirmFile.Visible = False
            End If
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("AuthorizeFile") <> "" Then     '���v��
                LAuthorizeFile.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("AuthorizeFile")
                wAuthorizeFile.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("AuthorizeFile")
            Else
                LAuthorizeFile.Visible = False
            End If
            DDevReason.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("DevReason")     '�}�o�z��

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Sample") = 1 Then              '�˫~
                Dim DBTable1 As DataTable
                SQL = "Select * From F_SampleList "
                SQL = SQL & " Where FormNo =  '" & wOFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "SampleList")
                DBTable1 = DBDataSet1.Tables("SampleList")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & DBTable1.Rows(i).Item("Size") & "," & DBTable1.Rows(i).Item("ChainType") & "," & DBTable1.Rows(i).Item("Class") & "," & _
                                DBTable1.Rows(i).Item("Class") & "," & DBTable1.Rows(i).Item("Qty") & "]"
                    If DSample.Text = "" Then
                        DSample.Text = Str
                    Else
                        DSample.Text = DSample.Text + ";" + Str
                    End If
                Next
                LSample.NavigateUrl = "SampleList.aspx?pFormNo=" & wOFormNo & "&pFormSno=" & wOFormSno
            Else
                LSample.Visible = False
            End If

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SampleFile") <> "" Then        '�˫~��
                LSampleFile.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SampleFile")
                wSampleFile.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SampleFile")
            Else
                LSampleFile.Visible = False
            End If

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Price") = 1 Then               '���
                Dim DBTable1 As DataTable
                SQL = "Select * From F_PriceList "
                SQL = SQL & " Where FormNo =  '" & wOFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "PriceList")
                DBTable1 = DBDataSet1.Tables("PriceList")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & DBTable1.Rows(i).Item("Size") & "," & DBTable1.Rows(i).Item("ChainType") & "," & DBTable1.Rows(i).Item("Class") & "," & _
                                DBTable1.Rows(i).Item("Price") & "]"
                    If DPrice.Text = "" Then
                        DPrice.Text = Str
                    Else
                        DPrice.Text = DPrice.Text + ";" + Str
                    End If
                Next
                LPrice.NavigateUrl = "PriceList.aspx?pFormNo=" & wOFormNo & "&pFormSno=" & wOFormSno
            Else
                LPrice.Visible = False
            End If

            DArMoldFee.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ArMoldFee")     '�����Ҩ�O
            DPurMold.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("PurMold")         '�Ҩ��ʤJ�O
            DPullerPrice.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("PullerPrice") '�ޤ��ʤJ��
            DMoldName.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldName")       '�Ҩ�W��
            DMoldQty.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldQty")         '�ҫ�
            DMoldPoint.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldPoint")     '�ި�

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Quality1") = 1 Then            '�~��1
                Dim DBTable1 As DataTable
                SQL = "Select * From F_QAList1 "
                SQL = SQL & " Where FormNo =  '" & wOFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "QAList1")
                DBTable1 = DBDataSet1.Tables("QAList1")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & _
                            DBTable1.Rows(i).Item("Date") & "," & _
                            DBTable1.Rows(i).Item("Size") & "," & DBTable1.Rows(i).Item("ChainType") & "," & DBTable1.Rows(i).Item("Class") & "," & _
                            DBTable1.Rows(i).Item("Assembler") & "," & DBTable1.Rows(i).Item("Surface1") & "," & DBTable1.Rows(i).Item("Surface2") & "," & DBTable1.Rows(i).Item("Gentani") & "," & _
                            DBTable1.Rows(i).Item("Kyoudo") & "," & DBTable1.Rows(i).Item("Nyuryoku") & "," & DBTable1.Rows(i).Item("Kensin") & "," & _
                            DBTable1.Rows(i).Item("Water") & "," & DBTable1.Rows(i).Item("Dry") & "," & _
                            DBTable1.Rows(i).Item("Yellow") & "," & DBTable1.Rows(i).Item("Mityaku") & "," & DBTable1.Rows(i).Item("CPSC") & _
                          "]"
                    If DQuality1.Text = "" Then
                        DQuality1.Text = Str
                    Else
                        DQuality1.Text = DQuality1.Text + ";" + Str
                    End If
                Next
                LQuality1.NavigateUrl = "Quality1List.aspx?pFormNo=" & wOFormNo & "&pFormSno=" & wOFormSno
            Else
                LQuality1.Visible = False
            End If

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Quality2") = 1 Then            '�~��2
                Dim DBTable1 As DataTable
                SQL = "Select * From F_QAList2 "
                SQL = SQL & " Where FormNo =  '" & wOFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "QAList2")
                DBTable1 = DBDataSet1.Tables("QAList2")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & DBTable1.Rows(i).Item("Date") & "," & DBTable1.Rows(i).Item("QACheck") & "," & DBTable1.Rows(i).Item("Remark") & "]"
                    If DQuality2.Text = "" Then
                        DQuality2.Text = Str
                    Else
                        DQuality2.Text = DQuality2.Text + ";" + Str
                    End If
                Next
                LQuality2.NavigateUrl = "Quality2List.aspx?pFormNo=" & wOFormNo & "&pFormSno=" & wOFormSno
            Else
                LQuality2.Visible = False
            End If

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("QAAttachFile") <> "" Then      '�~�����
                LQAAttachFile.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("QAAttachFile")
                wQAAttachFile.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("QAAttachFile")
            Else
                LQAAttachFile.Visible = False
            End If

            DFormSno.Text = "�渹�G" & CStr(wOFormSno)       '�渹
        End If
        'DB�s������
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     �]�w�u�X�����ƥ�
    '**
    '*****************************************************************
    Sub SetPopupFunction()
        BDate.Attributes("onclick") = "CalendarPicker('Form1.DDate');"  '������
        BSliderCode.Attributes("onclick") = "SliderPicker('Form1.DSliderCode');"  'Slider Code
        BSpec.Attributes("onclick") = "SpecPicker('Form1.DSpec');"      '�W��
        BOMapNo.Attributes("onclick") = "MapPicker('Ori');"    '��l�ϸ�
        BMMapNo.Attributes("onclick") = "MapPicker('Mod');"    '�ק�ϸ�
        BSample.Attributes("onclick") = "SamplePicker('Form1.DSample');"    '�˫~
        BPrice.Attributes("onclick") = "PricePicker('Form1.DPrice');"    '���
        BQuality1.Attributes("onclick") = "QA1Picker('Form1.DQuality1');"    '�~����R1
        BQuality2.Attributes("onclick") = "QA2Picker('Form1.DQuality2');"    '�~����R1
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
    '**     ��������ܤ�����J�ˬd
    '**
    '*****************************************************************
    Sub ShowSheetField(ByVal pPost As String)
        '�ظm�����ݩʰ}�C
        Dim oFieldAttribute As Object
        oFieldAttribute = Server.CreateObject("GetFieldAttb.WFField")
        oFieldAttribute.pFormNo = "900003"     '��渹�X
        oFieldAttribute.pStep = wStep          '�u�{���d���X
        oFieldAttribute.GetFieldAttribute(FieldName, Attribute)       '���W,����ݩ�

        SetFieldAttribute(pPost)     '���U����ݩʤ�����J�ˬd���]�w
    End Sub
    '*****************************************************************
    '**(ShowSheetField)
    '**     ���U����ݩʤ�����J�ˬd���]�w
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        '���U����ݩʤ�����J�ˬd���]�w
        'No
        Select Case FindFieldInf("No")
            Case 0  '���
                DNo.BackColor = Color.PeachPuff
                DNo.ReadOnly = True
                DNo.Visible = True
            Case 1  '�ק�+�ˬd
                DNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DnoRqd", "Dno", "���`�G�ݿ�J�ܢ�")
                DNo.Visible = True
            Case 2  '�ק�
                DNo.BackColor = Color.Yellow
                DNo.Visible = True
            Case Else   '����
                DNo.Visible = False
        End Select
        If pPost = "New" Then DNo.Text = ""
        '���
        Select Case FindFieldInf("Date")
            Case 0  '���
                DDate.BackColor = Color.PeachPuff
                DDate.ReadOnly = True
                DDate.Visible = True
                BDate.Visible = False
            Case 1  '�ק�+�ˬd
                DDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDateRqd", "DDate", "���`�G�ݿ�J���")
                DDate.Visible = True
                BDate.Visible = True
            Case 2  '�ק�
                DDate.BackColor = Color.Yellow
                DDate.Visible = True
                BDate.Visible = True
            Case Else   '����
                DDate.Visible = False
                BDate.Visible = False
        End Select
        If pPost = "New" Then DDate.Text = CStr(DateTime.Now.Today)
        '����
        Select Case FindFieldInf("Division")
            Case 0  '���
                DDivision.BackColor = Color.PeachPuff
                DDivision.Enabled = False
                DDivision.Visible = True
            Case 1  '�ק�+�ˬd
                DDivision.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDivisionRqd", "DDivision", "���`�G�ݿ�J����")
                DDivision.Visible = True
            Case 2  '�ק�
                DDivision.BackColor = Color.Yellow
                DDivision.Visible = True
            Case Else   '����
                DDivision.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Division", "ZZZZZZ")
        '���
        Select Case FindFieldInf("Person")
            Case 0  '���
                DPerson.BackColor = Color.PeachPuff
                DPerson.Enabled = False
                DPerson.Visible = True
            Case 1  '�ק�+�ˬd
                DPerson.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPersonRqd", "DPerson", "���`�G�ݿ�J���")
                DPerson.Visible = True
            Case 2  '�ק�
                DPerson.BackColor = Color.Yellow
                DPerson.Visible = True
            Case Else   '����
                DPerson.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Person", "ZZZZZZ")
        'Slider Group Code
        Select Case FindFieldInf("SliderGRCode")
            Case 0  '���
                DSliderGRCode.BackColor = Color.PeachPuff
                DSliderGRCode.ReadOnly = True
                DSliderGRCode.Visible = True
            Case 1  '�ק�+�ˬd
                DSliderGRCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderGRCodeRqd", "DSliderGRCode", "���`�G�ݿ�J Slider Code")
                DSliderGRCode.Visible = True
            Case 2  '�ק�
                DSliderGRCode.BackColor = Color.Yellow
                DSliderGRCode.Visible = True
            Case Else   '����
                DSliderGRCode.Visible = False
        End Select
        If pPost = "New" Then DSliderGRCode.Text = ""
        'Slider Code
        Select Case FindFieldInf("SliderCode")
            Case 0  '���
                DSliderCode.Visible = False
                BSliderCode.Visible = False
                LSliderCode.Visible = True
            Case 1  '�ק�+�ˬd
                DSliderCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderCodeRqd", "DSliderCode", "���`�G�ݿ�J Slider Code")
                DSliderCode.Visible = True
                BSliderCode.Visible = True
                LSliderCode.Visible = True
            Case 2  '�ק�
                DSliderCode.BackColor = Color.Yellow
                DSliderCode.Visible = True
                BSliderCode.Visible = True
                LSliderCode.Visible = True
            Case Else   '����
                DSliderCode.Visible = False
                BSliderCode.Visible = False
                LSliderCode.Visible = False
        End Select
        If pPost = "New" Then DSliderCode.Text = ""
        'Spec
        Select Case FindFieldInf("Spec")
            Case 0  '���
                DSpec.BackColor = Color.PeachPuff
                DSpec.ReadOnly = True
                DSpec.Visible = True
                BSpec.Visible = False
            Case 1  '�ק�+�ˬd
                DSpec.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSpecRqd", "DSpec", "���`�G�ݿ�J�W��(Size, Chain Type, ����)")
                DSpec.Visible = True
                BSpec.Visible = True
            Case 2  '�ק�
                DSpec.BackColor = Color.Yellow
                DSpec.Visible = True
                BSpec.Visible = True
            Case Else   '����
                DSpec.Visible = False
                BSpec.Visible = False
        End Select
        If pPost = "New" Then DSpec.Text = ""
        '�ϸ�
        Select Case FindFieldInf("MapNo")
            Case 0  '���
                DMapNo.BackColor = Color.PeachPuff
                DMapNo.ReadOnly = True
                DMapNo.Visible = True
                BOMapNo.Visible = False
                BMMapNo.Visible = False
            Case 1  '�ק�+�ˬd
                DMapNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMapNoRqd", "DMapNo", "���`�G�ݿ�J�ϸ�")
                DMapNo.Visible = True
                BOMapNo.Visible = True
                BMMapNo.Visible = True
            Case 2  '�ק�
                DMapNo.BackColor = Color.Yellow
                DMapNo.Visible = True
                BOMapNo.Visible = True
                BMMapNo.Visible = True
            Case Else   '����
                DMapNo.Visible = False
                BOMapNo.Visible = False
                BMMapNo.Visible = False
        End Select
        If pPost = "New" Then DMapNo.Text = ""

        '����1
        Select Case FindFieldInf("MapFile1")
            Case 0  '���
                DMapFile1.Visible = False
                LMapFile1.Visible = True
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DMapFile1Rqd", "DMapFile1", "���`�G�ݿ�J����")
                DMapFile1.Visible = True
                LMapFile1.Visible = True
            Case 2  '�ק�
                DMapFile1.Visible = True
                LMapFile1.Visible = True
            Case Else   '����
                DMapFile1.Visible = False
                LMapFile1.Visible = False
        End Select
        '����2
        Select Case FindFieldInf("MapFile2")
            Case 0  '���
                DMapFile2.Visible = False
                LMapFile2.Visible = True
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DMapFile2Rqd", "DMapFile2", "���`�G�ݿ�J����")
                DMapFile2.Visible = True
                LMapFile2.Visible = True
            Case 2  '�ק�
                DMapFile2.Visible = True
                LMapFile2.Visible = True
            Case Else   '����
                DMapFile2.Visible = False
                LMapFile2.Visible = False
        End Select
        '������
        Select Case FindFieldInf("Level")
            Case 0  '���
                DLevel.BackColor = Color.PeachPuff
                DLevel.Enabled = False
                DLevel.Visible = True
            Case 1  '�ק�+�ˬd
                DLevel.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DLevelRqd", "DLevel", "���`�G�ݿ�J������")
                DLevel.Visible = True
            Case 2  '�ק�
                DLevel.BackColor = Color.Yellow
                DLevel.Visible = True
            Case Else   '����
                DLevel.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Level", "ZZZZZZ")
        '�եߧP�O
        Select Case FindFieldInf("Assembler")
            Case 0  '���
                DAssembler.BackColor = Color.PeachPuff
                DAssembler.ReadOnly = True
                DAssembler.Visible = True
            Case 1  '�ק�+�ˬd
                DAssembler.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAssemblerRqd", "DAssembler", "���`�G�ݿ�J�եߧP�w")
                DAssembler.Visible = True
            Case 2  '�ק�
                DAssembler.BackColor = Color.Yellow
                DAssembler.Visible = True
            Case Else   '����
                DAssembler.Visible = False
        End Select
        If pPost = "New" Then DAssembler.Text = ""
        '���Y����(���s.�~�`...)
        Select Case FindFieldInf("SliderType1")
            Case 0  '���
                DSliderType1.BackColor = Color.PeachPuff
                DSliderType1.Enabled = False
                DSliderType1.Visible = True
            Case 1  '�ק�+�ˬd
                DSliderType1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderType1Rqd", "DSliderType1", "���`�G�ݿ�J���Y����")
                DSliderType1.Visible = True
            Case 2  '�ק�
                DSliderType1.BackColor = Color.Yellow
                DSliderType1.Visible = True
            Case Else   '����
                DSliderType1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SliderType1", "ZZZZZZ")
        '���Y����(�b���~.���~...)
        Select Case FindFieldInf("SliderType2")
            Case 0  '���
                DSliderType2.BackColor = Color.PeachPuff
                DSliderType2.Enabled = False
                DSliderType2.Visible = True
            Case 1  '�ק�+�ˬd
                DSliderType2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderType2Rqd", "DSliderType2", "���`�G�ݿ�J���Y����")
                DSliderType2.Visible = True
            Case 2  '�ק�
                DSliderType2.BackColor = Color.Yellow
                DSliderType2.Visible = True
            Case Else   '����
                DSliderType2.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SliderType2", "ZZZZZZ")
        '�Ͳ��a
        Select Case FindFieldInf("ManufPlace")
            Case 0  '���
                DManufPlace.BackColor = Color.PeachPuff
                DManufPlace.Enabled = False
                DManufPlace.Visible = True
            Case 1  '�ק�+�ˬd
                DManufPlace.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DManufPlaceRqd", "DManufPlace", "���`�G�ݿ�J�Ͳ��a")
                DManufPlace.Visible = True
            Case 2  '�ק�
                DManufPlace.BackColor = Color.Yellow
                DManufPlace.Visible = True
            Case Else   '����
                DManufPlace.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ManufPlace", "ZZZZZZ")
        '����
        Select Case FindFieldInf("Material")
            Case 0  '���
                DMaterial.BackColor = Color.PeachPuff
                DMaterial.Enabled = False
                DMaterial.Visible = True
            Case 1  '�ק�+�ˬd
                DMaterial.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMaterialRqd", "DMaterial", "���`�G�ݿ�J����")
                DMaterial.Visible = True
            Case 2  '�ק�
                DMaterial.BackColor = Color.Yellow
                DMaterial.Visible = True
            Case Else   '����
                DMaterial.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Material", "ZZZZZZ")
        '����Ӷ�
        Select Case FindFieldInf("MaterialDetail")
            Case 0  '���
                DMaterialDetail.BackColor = Color.PeachPuff
                DMaterialDetail.Enabled = False
                DMaterialDetail.Visible = True
            Case 1  '�ק�+�ˬd
                DMaterialDetail.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMaterialDetailRqd", "DMaterialDetail", "���`�G�ݿ�J����Ӷ�")
                DMaterialDetail.Visible = True
            Case 2  '�ק�
                DMaterialDetail.BackColor = Color.Yellow
                DMaterialDetail.Visible = True
            Case Else   '����
                DMaterialDetail.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("MaterialDetail", "ZZZZZZ")
        '����Ƶ�
        Select Case FindFieldInf("MaterialOther")
            Case 0  '���
                DMaterialOther.BackColor = Color.PeachPuff
                DMaterialOther.ReadOnly = True
                DMaterialOther.Visible = True
            Case 1  '�ק�+�ˬd
                DMaterialOther.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMaterialOtherRqd", "DMaterialOther", "���`�G�ݿ�J����Ƶ�")
                DMaterialOther.Visible = True
            Case 2  '�ק�
                DMaterialOther.BackColor = Color.Yellow
                DMaterialOther.Visible = True
            Case Else   '����
                DMaterialOther.Visible = False
        End Select
        If pPost = "New" Then DMaterialOther.Text = ""
        '�e�U�t��
        Select Case FindFieldInf("SellVendor")
            Case 0  '���
                DSellVendor.BackColor = Color.PeachPuff
                DSellVendor.ReadOnly = True
                DSellVendor.Visible = True
            Case 1  '�ק�+�ˬd
                DSellVendor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSellVendorRqd", "DSellVendor", "���`�G�ݿ�J�e�U�t��")
                DSellVendor.Visible = True
            Case 2  '�ק�
                DSellVendor.BackColor = Color.Yellow
                DSellVendor.Visible = True
            Case Else   '����
                DSellVendor.Visible = False
        End Select
        If pPost = "New" Then DSellVendor.Text = ""
        'Buyer
        Select Case FindFieldInf("Buyer")
            Case 0  '���
                DBuyer.BackColor = Color.PeachPuff
                DBuyer.ReadOnly = True
                DBuyer.Visible = True
            Case 1  '�ק�+�ˬd
                DBuyer.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBuyerRqd", "DBuyer", "���`�G�ݿ�JBuyer")
                DBuyer.Visible = True
            Case 2  '�ק�
                DBuyer.BackColor = Color.Yellow
                DBuyer.Visible = True
            Case Else   '����
                DBuyer.Visible = False
        End Select
        If pPost = "New" Then DBuyer.Text = ""
        '�T�{��
        Select Case FindFieldInf("ConfirmFile")
            Case 0  '���
                DConfirmFile.Visible = False
                LConfirmFile.Visible = True
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DConfirmFileRqd", "DConfirmFile", "���`�G�ݿ�J�T�{��")
                DConfirmFile.Visible = True
                LConfirmFile.Visible = True
            Case 2  '�ק�
                DConfirmFile.Visible = True
                LConfirmFile.Visible = True
            Case Else   '����
                DConfirmFile.Visible = False
                LConfirmFile.Visible = False
        End Select
        '���v��
        Select Case FindFieldInf("AuthorizeFile")
            Case 0  '���
                DAuthorizeFile.Visible = False
                LAuthorizeFile.Visible = True
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DAuthorizeFileRqd", "DAuthorizeFile", "���`�G�ݿ�J���v��")
                DAuthorizeFile.Visible = True
                LAuthorizeFile.Visible = True
            Case 2  '�ק�
                DAuthorizeFile.Visible = True
                LAuthorizeFile.Visible = True
            Case Else   '����
                DAuthorizeFile.Visible = False
                LAuthorizeFile.Visible = False
        End Select
        '�}�o�z��
        Select Case FindFieldInf("DevReason")
            Case 0  '���
                DDevReason.BackColor = Color.PeachPuff
                DDevReason.ReadOnly = True
                DDevReason.Visible = True
            Case 1  '�ק�+�ˬd
                DDevReason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDevReasonRqd", "DDevReason", "���`�G�ݿ�J�}�o�z��")
                DDevReason.Visible = True
            Case 2  '�ק�
                DDevReason.BackColor = Color.Yellow
                DDevReason.Visible = True
            Case Else   '����
                DDevReason.Visible = False
        End Select
        If pPost = "New" Then DDevReason.Text = ""
        'Sample
        Select Case FindFieldInf("Sample")
            Case 0  '���
                DSample.Visible = False
                BSample.Visible = False
                LSample.Visible = True
            Case 1  '�ק�+�ˬd
                DSample.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSampleRqd", "DSample", "���`�G�ݿ�J�˫~��T")
                DSample.Visible = True
                BSample.Visible = True
                LSample.Visible = True
            Case 2  '�ק�
                DSample.BackColor = Color.Yellow
                DSample.Visible = True
                BSample.Visible = True
                LSample.Visible = True
            Case Else   '����
                DSample.Visible = False
                BSample.Visible = False
                LSample.Visible = False
        End Select
        If pPost = "New" Then DSample.Text = ""
        '�˫~��
        Select Case FindFieldInf("SampleFile")
            Case 0  '���
                DSampleFile.Visible = False
                LSampleFile.Visible = True
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DSampleFileRqd", "DSampleFile", "���`�G�ݿ�J�˫~��")
                DSampleFile.Visible = True
                LSampleFile.Visible = True
            Case 2  '�ק�
                DSampleFile.Visible = True
                LSampleFile.Visible = True
            Case Else   '����
                DSampleFile.Visible = False
                LSampleFile.Visible = False
        End Select
        'Price
        Select Case FindFieldInf("Price")
            Case 0  '���
                DPrice.Visible = False
                BPrice.Visible = False
                LPrice.Visible = True
            Case 1  '�ק�+�ˬd
                DPrice.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPriceRqd", "DPrice", "���`�G�ݿ�J�����T")
                DPrice.Visible = True
                BPrice.Visible = True
                LPrice.Visible = True
            Case 2  '�ק�
                DPrice.BackColor = Color.Yellow
                DPrice.Visible = True
                BPrice.Visible = True
                LPrice.Visible = True
            Case Else   '����
                DPrice.Visible = False
                BPrice.Visible = False
                LPrice.Visible = False
        End Select
        If pPost = "New" Then DPrice.Text = ""
        '�����Ҩ�O
        Select Case FindFieldInf("ArMoldFee")
            Case 0  '���
                DArMoldFee.BackColor = Color.PeachPuff
                DArMoldFee.ReadOnly = True
                DArMoldFee.Visible = True
            Case 1  '�ק�+�ˬd
                DArMoldFee.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DArMoldFeeRqd", "DArMoldFee", "���`�G�ݿ�J�����Ҩ�O")
                DArMoldFee.Visible = True
            Case 2  '�ק�
                DArMoldFee.BackColor = Color.Yellow
                DArMoldFee.Visible = True
            Case Else   '����
                DArMoldFee.Visible = False
        End Select
        If pPost = "New" Then DArMoldFee.Text = ""
        '�Ҩ��ʤJ��
        Select Case FindFieldInf("PurMold")
            Case 0  '���
                DPurMold.BackColor = Color.PeachPuff
                DPurMold.ReadOnly = True
                DPurMold.Visible = True
            Case 1  '�ק�+�ˬd
                DPurMold.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPurMoldRqd", "DPurMold", "���`�G�ݿ�J�Ҩ��ʤJ��")
                DPurMold.Visible = True
            Case 2  '�ק�
                DPurMold.BackColor = Color.Yellow
                DPurMold.Visible = True
            Case Else   '����
                DPurMold.Visible = False
        End Select
        If pPost = "New" Then DPurMold.Text = ""
        '�ޤ��ʤJ���
        Select Case FindFieldInf("PullerPrice")
            Case 0  '���
                DPullerPrice.BackColor = Color.PeachPuff
                DPullerPrice.ReadOnly = True
                DPullerPrice.Visible = True
            Case 1  '�ק�+�ˬd
                DPullerPrice.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPullerPriceRqd", "DPullerPrice", "���`�G�ݿ�J�ޤ��ʤJ���")
                DPullerPrice.Visible = True
            Case 2  '�ק�
                DPullerPrice.BackColor = Color.Yellow
                DPullerPrice.Visible = True
            Case Else   '����
                DPullerPrice.Visible = False
        End Select
        If pPost = "New" Then DPullerPrice.Text = ""
        '�Ҩ�W��
        Select Case FindFieldInf("MoldName")
            Case 0  '���
                DMoldName.BackColor = Color.PeachPuff
                DMoldName.ReadOnly = True
                DMoldName.Visible = True
            Case 1  '�ק�+�ˬd
                DMoldName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMoldNameRqd", "DMoldName", "���`�G�ݿ�J�Ҩ�W��")
                DMoldName.Visible = True
            Case 2  '�ק�
                DMoldName.BackColor = Color.Yellow
                DMoldName.Visible = True
            Case Else   '����
                DMoldName.Visible = False
        End Select
        If pPost = "New" Then DMoldName.Text = ""
        '�ҫ��Ө�-�ҫ�
        Select Case FindFieldInf("MoldQty")
            Case 0  '���
                DMoldQty.BackColor = Color.PeachPuff
                DMoldQty.ReadOnly = True
                DMoldQty.Visible = True
            Case 1  '�ק�+�ˬd
                DMoldQty.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMoldQtyRqd", "DMoldQty", "���`�G�ݿ�J�ҫ��Ө�-�ҫ�")
                DMoldQty.Visible = True
            Case 2  '�ק�
                DMoldQty.BackColor = Color.Yellow
                DMoldQty.Visible = True
            Case Else   '����
                DMoldQty.Visible = False
        End Select
        If pPost = "New" Then DMoldQty.Text = ""
        '�ҫ��Ө�-�ި�
        Select Case FindFieldInf("MoldPoint")
            Case 0  '���
                DMoldPoint.BackColor = Color.PeachPuff
                DMoldPoint.ReadOnly = True
                DMoldPoint.Visible = True
            Case 1  '�ק�+�ˬd
                DMoldPoint.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMoldPointRqd", "DMoldPoint", "���`�G�ݿ�J�ҫ��Ө�-�ި�")
                DMoldPoint.Visible = True
            Case 2  '�ק�
                DMoldPoint.BackColor = Color.Yellow
                DMoldPoint.Visible = True
            Case Else   '����
                DMoldPoint.Visible = False
        End Select
        If pPost = "New" Then DMoldPoint.Text = ""
        'Quality1
        Select Case FindFieldInf("Quality1")
            Case 0  '���
                DQuality1.Visible = False
                BQuality1.Visible = False
                LQuality1.Visible = True
            Case 1  '�ק�+�ˬd
                DQuality1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQuality1Rqd", "DQuality1", "���`�G�ݿ�J�~����R��T")
                DQuality1.Visible = True
                BQuality1.Visible = True
                LQuality1.Visible = True
            Case 2  '�ק�
                DQuality1.BackColor = Color.Yellow
                DQuality1.Visible = True
                BQuality1.Visible = True
                LQuality1.Visible = True
            Case Else   '����
                DQuality1.Visible = False
                BQuality1.Visible = False
                LQuality1.Visible = False
        End Select
        If pPost = "New" Then DQuality1.Text = ""
        'Quality2
        Select Case FindFieldInf("Quality2")
            Case 0  '���
                DQuality2.Visible = False
                BQuality2.Visible = False
                LQuality2.Visible = True
            Case 1  '�ק�+�ˬd
                DQuality2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQuality2Rqd", "DQuality2", "���`�G�ݿ�J�~����R��T")
                DQuality2.Visible = True
                BQuality2.Visible = True
                LQuality2.Visible = True
            Case 2  '�ק�
                DQuality2.BackColor = Color.Yellow
                DQuality2.Visible = True
                BQuality2.Visible = True
                LQuality2.Visible = True
            Case Else   '����
                DQuality2.Visible = False
                BQuality2.Visible = False
                LQuality2.Visible = False
        End Select
        If pPost = "New" Then DQuality2.Text = ""
        '�~����R��
        Select Case FindFieldInf("QAAttachFile")
            Case 0  '���
                DQAAttachFile.Visible = False
                LQAAttachFile.Visible = True
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DQAAttachFileRqd", "DQAAttachFile", "���`�G�ݿ�J�~����R��")
                DQAAttachFile.Visible = True
                LQAAttachFile.Visible = True
            Case 2  '�ק�
                DQAAttachFile.Visible = True
                LQAAttachFile.Visible = True
            Case Else   '����
                DQAAttachFile.Visible = False
                LQAAttachFile.Visible = False
        End Select
        '�渹
        If pPost = "New" Then DFormSno.Text = ""
    End Sub
    '*****************************************************************
    '**(ShowSheetField)
    '**     �إߪ��ݿ�J���
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator
        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.ErrorMessage = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", CStr(Top + 25))
        rqdVal.Style.Add("Left", "8")
        rqdVal.Style.Add("Height", "20")
        rqdVal.Style.Add("Width", "250")
        rqdVal.Style.Add("POSITION", "absolute")
        Page.Controls(1).Controls.Add(rqdVal)
    End Sub

    '*****************************************************************
    '**(ShowSheetField)
    '**     �ظm�U�Ԧ�����l��
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim SQL As String
        Dim i As Integer
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()
        '����
        If pFieldName = "Division" Then
            SQL = "Select * From M_Referp Where Cat='100' and DKey='DIVISION' Order by Data "
            DDivision.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DDivision.Items.Add(ListItem1)
            Next
        End If
        '���
        If pFieldName = "Person" Then
            SQL = "Select * From M_Referp Where Cat='100' and DKey='PERSON' Order by Data "
            DPerson.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DPerson.Items.Add(ListItem1)
            Next
        End If
        '������
        If pFieldName = "Level" Then
            SQL = "Select * From M_Referp Where Cat='007' and DKey<>'Z' Order by DKey, Data "
            DLevel.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DLevel.Items.Add(ListItem1)
            Next
        End If
        '���Y����(���s.�~�`...)
        If pFieldName = "SliderType1" Then
            SQL = "Select * From M_Referp Where Cat='200' and DKey='SLIDERTYPE1' Order by Data "
            DSliderType1.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DSliderType1.Items.Add(ListItem1)
            Next
        End If
        '���Y����(�b���~.���~...)
        If pFieldName = "SliderType2" Then
            SQL = "Select * From M_Referp Where Cat='200' and DKey='SLIDERTYPE2' Order by Data "
            DSliderType2.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DSliderType2.Items.Add(ListItem1)
            Next
        End If
        '�Ͳ��a
        If pFieldName = "ManufPlace" Then
            SQL = "Select * From M_Referp Where Cat='200' and DKey='MANUFPLACE' Order by Data "
            DManufPlace.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DManufPlace.Items.Add(ListItem1)
            Next
        End If
        '����
        If pFieldName = "Material" Then
            SQL = "Select * From M_Referp Where Cat='100' and DKey='MATERIAL' Order by Data "
            DMaterial.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DMaterial.Items.Add(ListItem1)
            Next
        End If
        '����Ӷ�
        If pFieldName = "MaterialDetail" Then
            SQL = "Select * From M_Referp Where Cat='101' and DKey= '" & DMaterial.SelectedValue & "' Order by Data "
            DMaterialDetail.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DMaterialDetail.Items.Add(ListItem1)
            Next
        End If
        'DB�s������
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
    '**     �M��������ݩ�
    '**
    '*****************************************************************
    Function FindFieldInf(ByVal pFieldName As String) As Integer
        Dim Run As Boolean
        Dim i As Integer
        Run = True
        FindFieldInf = 9
        While i < 44 And Run
            If FieldName(i) = pFieldName Then
                FindFieldInf = Attribute(i)
                Run = False
            End If
            i = i + 1
        End While
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����I���ƥ�
    '**
    '*****************************************************************
    Private Sub DMaterial_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DMaterial.SelectedIndexChanged
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim SQL As String
        Dim i As Integer
        SQL = "Select * From M_Referp Where Cat='101' and DKey= '" & DMaterial.SelectedValue & "' Order by Data "
        DMaterialDetail.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Referp")
        DBTable1 = DBDataSet1.Tables("M_Referp")
        For i = 0 To DBTable1.Rows.Count - 1        '��ܲӶ��������e
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("Data")
            ListItem1.Value = DBTable1.Rows(i).Item("Data")
            DMaterialDetail.Items.Add(ListItem1)
        Next
        'DB�s������
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �x�s���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSave.Click
        Dim ErrCode As Integer = 0   '�ŧi�@��Error�ΨӧP�O�O�_��ƥ��`�A�w�]��False
        Dim Message As String = ""
        '�W�Ǥ��
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)

        Dim NewFormSno As Integer = wFormSno    '���y����
        Dim RtnCode As Integer = 0
        '--��LSubFile---------
        Dim oPutSubFile As Object
        '--���o���y�����]�w---------
        Dim oGetSeqNo As Object

        'Check����-1Size�ή榡
        If ErrCode = 0 Then
            If DMapFile1.Visible Then
                If DMapFile1.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DMapFile1)
                End If
            End If
        End If
        'Check����-2Size�ή榡
        If ErrCode = 0 Then
            If DMapFile2.Visible Then
                If DMapFile2.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DMapFile2)
                End If
            End If
        End If
        'Check�T�{��Size�ή榡
        If ErrCode = 0 Then
            If DConfirmFile.Visible Then
                If DConfirmFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DConfirmFile)
                End If
            End If
        End If
        'Check���v��Size�ή榡
        If ErrCode = 0 Then
            If DAuthorizeFile.Visible Then
                If DAuthorizeFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DAuthorizeFile)
                End If
            End If
        End If
        'Check�˫~��Size�ή榡
        If ErrCode = 0 Then
            If DSampleFile.Visible Then
                If DSampleFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DSampleFile)
                End If
            End If
        End If
        'CheckQA����Size�ή榡
        If ErrCode = 0 Then
            If DQAAttachFile.Visible Then
                If DQAAttachFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DQAAttachFile)
                End If
            End If
        End If

        '�x�s���
        If ErrCode = 0 Then
            Dim Path As String
            If wOFormNo = "000003" Then
                Path = Server.MapPath(ConfigurationSettings.AppSettings.Get("ManufInFilePath"))
            Else
                Path = Server.MapPath(ConfigurationSettings.AppSettings.Get("ManufOutFilePath"))
            End If
            Dim FileName As String
            Dim SQL As String

            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
            Dim OleDBCommand1 As New OleDbCommand

            '�x�s�����
            If wFormSno > 0 Then    '�P�_�O�_�����
                SQL = "Update F_ModManufSheet Set "
                SQL = SQL + " No = '" & DNo.Text & "',"
                SQL = SQL + " Date = '" & DDate.Text & "',"
                SQL = SQL + " Division = '" & DDivision.SelectedValue & "',"
                SQL = SQL + " Person = '" & DPerson.SelectedValue & "',"
                If DSliderCode.Text <> "" Then
                    SQL = SQL + " SliderCode = '" & "1" & "',"
                Else
                    SQL = SQL + " SliderCode = '" & "0" & "',"
                End If
                SQL = SQL + " SliderGRCode = '" & DSliderGRCode.Text & "',"
                SQL = SQL + " Spec = '" & DSpec.Text & "',"
                SQL = SQL + " MapNo = '" & DMapNo.Text & "',"

                If DMapFile1.Visible Then
                    If DMapFile1.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                        FileName = CStr(wFormSno) & "-" & "Map1" & "-" & UploadDateTime & "-" & Right(DMapFile1.PostedFile.FileName, InStr(StrReverse(DMapFile1.PostedFile.FileName), "\") - 1)
                        Try    '�W�ǹ���
                            DMapFile1.PostedFile.SaveAs(Path + FileName)
                        Catch ex As Exception
                        End Try
                        SQL = SQL + " MapFile1 = '" & FileName & "',"
                    End If
                End If

                If DMapFile2.Visible Then
                    If DMapFile2.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                        FileName = CStr(wFormSno) & "-" & "Map2" & "-" & UploadDateTime & "-" & Right(DMapFile2.PostedFile.FileName, InStr(StrReverse(DMapFile2.PostedFile.FileName), "\") - 1)
                        Try    '�W�ǹ���
                            DMapFile2.PostedFile.SaveAs(Path + FileName)
                        Catch ex As Exception
                        End Try
                        SQL = SQL + " MapFile2 = '" & FileName & "',"
                    End If
                End If

                SQL = SQL + " Level = '" & DLevel.SelectedValue & "',"
                SQL = SQL + " Assembler = '" & DAssembler.Text & "',"
                SQL = SQL + " SliderType1 = '" & DSliderType1.SelectedValue & "',"
                SQL = SQL + " SliderType2 = '" & DSliderType2.SelectedValue & "',"
                SQL = SQL + " ManufPlace = '" & DManufPlace.SelectedValue & "',"
                SQL = SQL + " Material = '" & DMaterial.SelectedValue & "',"
                SQL = SQL + " MaterialDetail = '" & DMaterialDetail.SelectedValue & "',"
                SQL = SQL + " MaterialOther = '" & DMaterialOther.Text & "',"
                SQL = SQL + " SellVendor = '" & DSellVendor.Text & "',"
                SQL = SQL + " Buyer = '" & DBuyer.Text & "',"

                If DConfirmFile.Visible Then
                    If DConfirmFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                        FileName = CStr(wFormSno) & "-" & "Confirm" & "-" & UploadDateTime & "-" & Right(DConfirmFile.PostedFile.FileName, InStr(StrReverse(DConfirmFile.PostedFile.FileName), "\") - 1)
                        Try    '�W�ǹ���
                            DConfirmFile.PostedFile.SaveAs(Path + FileName)
                        Catch ex As Exception
                        End Try
                        SQL = SQL + " ConfirmFile = '" & FileName & "',"
                    End If
                End If

                If DAuthorizeFile.Visible Then
                    If DAuthorizeFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                        FileName = CStr(wFormSno) & "-" & "Authorize" & "-" & UploadDateTime & "-" & Right(DAuthorizeFile.PostedFile.FileName, InStr(StrReverse(DAuthorizeFile.PostedFile.FileName), "\") - 1)
                        Try    '�W�ǹ���
                            DAuthorizeFile.PostedFile.SaveAs(Path + FileName)
                        Catch ex As Exception
                        End Try
                        SQL = SQL + " AuthorizeFile = '" & FileName & "',"
                    End If
                End If

                SQL = SQL + " DevReason = '" & DDevReason.Text & "',"
                If DSample.Text <> "" Then
                    SQL = SQL + " Sample = '" & "1" & "',"
                Else
                    SQL = SQL + " Sample = '" & "0" & "',"
                End If
                If DPrice.Text <> "" Then
                    SQL = SQL + " Price = '" & "1" & "',"
                Else
                    SQL = SQL + " Price = '" & "0" & "',"
                End If
                SQL = SQL + " ArMoldFee = '" & DArMoldFee.Text & "',"
                SQL = SQL + " PurMold = '" & DPurMold.Text & "',"
                SQL = SQL + " PullerPrice = '" & DPullerPrice.Text & "',"
                SQL = SQL + " MoldName = '" & DMoldName.Text & "',"
                SQL = SQL + " MoldQty = '" & DMoldQty.Text & "',"
                SQL = SQL + " MoldPoint = '" & DMoldPoint.Text & "',"
                If DQuality1.Text <> "" Then
                    SQL = SQL + " Quality1 = '" & "1" & "',"
                Else
                    SQL = SQL + " Quality1 = '" & "0" & "',"
                End If
                If DQuality2.Text <> "" Then
                    SQL = SQL + " Quality2 = '" & "1" & "',"
                Else
                    SQL = SQL + " Quality2 = '" & "0" & "',"
                End If

                If DQAAttachFile.Visible Then
                    If DQAAttachFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                        FileName = CStr(wFormSno) & "-" & "QA" & "-" & UploadDateTime & "-" & Right(DQAAttachFile.PostedFile.FileName, InStr(StrReverse(DQAAttachFile.PostedFile.FileName), "\") - 1)
                        Try    '�W�ǹ���
                            DQAAttachFile.PostedFile.SaveAs(Path + FileName)
                        Catch ex As Exception
                        End Try
                        SQL = SQL + " QAAttachFile = '" & FileName & "',"
                    End If
                End If

                If DSampleFile.Visible Then
                    If DSampleFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                        FileName = CStr(wFormSno) & "-" & "Sample" & "-" & UploadDateTime & "-" & Right(DSampleFile.PostedFile.FileName, InStr(StrReverse(DSampleFile.PostedFile.FileName), "\") - 1)
                        Try    '�W�ǹ���
                            DSampleFile.PostedFile.SaveAs(Path + FileName)
                        Catch ex As Exception
                        End Try
                        SQL = SQL + " SampleFile = '" & FileName & "',"
                    End If
                End If
                SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
                SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
                SQL = SQL + " Where FormNo  =  '" & wFormNo & "'"
                SQL = SQL + "   And FormSno =  '" & CStr(wFormSno) & "'"
                OleDBCommand1.Connection = OleDbConnection1
                OleDBCommand1.CommandText = SQL
                OleDbConnection1.Open()
                OleDBCommand1.ExecuteNonQuery()
                OleDbConnection1.Close()

                'Slider Table�B�z
                oPutSubFile = Server.CreateObject("MSubFile.SubFile")
                RtnCode = oPutSubFile.PutSlider(DSliderCode.Text, DSliderGRCode.Text, "900003", wFormSno, Request.Cookies("UserID").Value)
                'Sample Table�B�z
                RtnCode = oPutSubFile.PutSample(DSample.Text, "900003", wFormSno, Request.Cookies("UserID").Value)
                'Price Table�B�z
                RtnCode = oPutSubFile.PutPrice(DPrice.Text, "900003", wFormSno, Request.Cookies("UserID").Value)
                'QA1 Table�B�z
                RtnCode = oPutSubFile.PutQA1(DQuality1.Text, "900003", wFormSno, Request.Cookies("UserID").Value)
                'QA2 Table�B�z
                RtnCode = oPutSubFile.PutQA2(DQuality2.Text, "900003", wFormSno, Request.Cookies("UserID").Value)
            Else
                '���o���y����
                oGetSeqNo = Server.CreateObject("GetSeqno.WFFormInf")
                RtnCode = oGetSeqNo.Seqno(wFormNo, NewFormSno)   '��渹�X, ���y����
                If RtnCode <> 0 Then
                    ErrCode = 9110
                Else
                    SQL = "Insert into F_ModManufSheet "
                    SQL = SQL + "( "
                    SQL = SQL + "Sts, CompletedTime, FormNo, FormSNo, "                         '1~4
                    SQL = SQL + "No, Date, Division, Person, SliderCode, "                      '5~9
                    SQL = SQL + "SliderGRCode, Spec, MapNo, MapFile1, MapFile2, "               '10~14
                    SQL = SQL + "Level, Assembler, SliderType1, SliderType2, ManufPlace, "      '15~19
                    SQL = SQL + "Material, MaterialDetail, MaterialOther, SellVendor, Buyer, "  '20~24
                    SQL = SQL + "ConfirmFile, AuthorizeFile, DevReason, Sample, Price, "        '25~29
                    SQL = SQL + "ArMoldFee, PurMold, PullerPrice, MoldName, MoldQty, "          '29~33
                    SQL = SQL + "MoldPoint, Quality1, Quality2, QAAttachFile, SampleFile,  "    '34~38
                    SQL = SQL + "CreateUser, CreateTime, ModifyUser, ModifyTime "               '39~42
                    SQL = SQL + ")  "
                    SQL = SQL + "VALUES( "
                    SQL = SQL + " '0', "                                '���A(0:����,1:�w��NG,2:�w��OK)
                    SQL = SQL + " '" + NowDateTime + "', "              '���פ�
                    SQL = SQL + " '900003', "                           '���N��
                    SQL = SQL + " '" + Str(NewFormSno) + "', "            '���y����

                    SQL = SQL + " '" + DNo.Text + "', "                 'NO
                    SQL = SQL + " '" + DDate.Text + "', "               '���
                    SQL = SQL + " '" + DDivision.SelectedValue + "', "  '����
                    SQL = SQL + " '" + DPerson.SelectedValue + "', "    '���
                    If DSliderCode.Text <> "" Then                      'Slider Code(��ܥ�)
                        SQL = SQL + " '1', "
                    Else
                        SQL = SQL + " '0', "
                    End If
                    SQL = SQL + " '" + DSliderGRCode.Text + "', "       'Slider Group Code
                    SQL = SQL + " '" + DSpec.Text + "', "               '�W��
                    SQL = SQL + " '" + DMapNo.Text + "', "              '�ϸ�

                    If DMapFile1.Visible Then                           '����1
                        If DMapFile1.PostedFile.FileName <> "" Then     '�P�_���ɮפW��
                            FileName = CStr(wFormSno) & "-" & "Map1" & "-" & UploadDateTime & "-" & Right(DMapFile1.PostedFile.FileName, InStr(StrReverse(DMapFile1.PostedFile.FileName), "\") - 1)
                            Try    '�W�ǹ���
                                DMapFile1.PostedFile.SaveAs(Path + FileName)
                            Catch ex As Exception
                            End Try
                        Else
                            FileName = wMapFile1.Text
                        End If
                    Else
                        FileName = wMapFile1.Text
                    End If
                    SQL = SQL + " '" + FileName + "', "

                    If DMapFile2.Visible Then                           '����2
                        If DMapFile2.PostedFile.FileName <> "" Then     '�P�_���ɮפW��
                            FileName = CStr(wFormSno) & "-" & "Map2" & "-" & UploadDateTime & "-" & Right(DMapFile2.PostedFile.FileName, InStr(StrReverse(DMapFile2.PostedFile.FileName), "\") - 1)
                            Try    '�W�ǹ���
                                DMapFile2.PostedFile.SaveAs(Path + FileName)
                            Catch ex As Exception
                            End Try
                        Else
                            FileName = wMapFile2.Text
                        End If
                    Else
                        FileName = wMapFile2.Text
                    End If
                    SQL = SQL + " '" + FileName + "', "

                    SQL = SQL + " '" + DLevel.SelectedValue + "', "           '������
                    SQL = SQL + " '" + DAssembler.Text + "', "                '�եߧP�w
                    SQL = SQL + " '" + DSliderType1.SelectedValue + "', "     '���Y����-1
                    SQL = SQL + " '" + DSliderType2.SelectedValue + "', "     '���Y����-2
                    SQL = SQL + " '" + DManufPlace.SelectedValue + "', "      '�Ͳ��a
                    SQL = SQL + " '" + DMaterial.SelectedValue + "', "        '����
                    SQL = SQL + " '" + DMaterialDetail.SelectedValue + "', "  '����Ӷ�
                    SQL = SQL + " '" + DMaterialOther.Text + "', "            '�����L����
                    SQL = SQL + " '" + DSellVendor.Text + "', "               '�e�U�t��
                    SQL = SQL + " '" + DBuyer.Text + "', "                    'Buyer
                    If DConfirmFile.Visible Then                              '�T�{��
                        If DConfirmFile.PostedFile.FileName <> "" Then     '�P�_���ɮפW��
                            FileName = CStr(wFormSno) & "-" & "Confirm" & "-" & UploadDateTime & "-" & Right(DConfirmFile.PostedFile.FileName, InStr(StrReverse(DConfirmFile.PostedFile.FileName), "\") - 1)
                            Try    '�W�ǹ���
                                DConfirmFile.PostedFile.SaveAs(Path + FileName)
                            Catch ex As Exception
                            End Try
                        Else
                            FileName = wConfirmFile.Text
                        End If
                    Else
                        FileName = wConfirmFile.Text
                    End If
                    SQL = SQL + " '" + FileName + "', "

                    If DAuthorizeFile.Visible Then                            '���v��
                        If DAuthorizeFile.PostedFile.FileName <> "" Then     '�P�_���ɮפW��
                            FileName = CStr(wFormSno) & "-" & "Authorize" & "-" & UploadDateTime & "-" & Right(DAuthorizeFile.PostedFile.FileName, InStr(StrReverse(DAuthorizeFile.PostedFile.FileName), "\") - 1)
                            Try    '�W�ǹ���
                                DAuthorizeFile.PostedFile.SaveAs(Path + FileName)
                            Catch ex As Exception
                            End Try
                        Else
                            FileName = wAuthorizeFile.Text
                        End If
                    Else
                        FileName = wAuthorizeFile.Text
                    End If
                    SQL = SQL + " '" + FileName + "', "

                    SQL = SQL + " '" + DDevReason.Text + "', "          '�}�o�z��
                    If DSample.Text <> "" Then                          'Sample(��ܥ�)
                        SQL = SQL + " '1', "
                    Else
                        SQL = SQL + " '0', "
                    End If
                    If DPrice.Text <> "" Then                           'Price(��ܥ�)
                        SQL = SQL + " '1', "
                    Else
                        SQL = SQL + " '0', "
                    End If
                    SQL = SQL + " '" + DArMoldFee.Text + "', "          '�����Ҩ�O
                    SQL = SQL + " '" + DPurMold.Text + "', "            '�Ҩ��ʤJ�O
                    SQL = SQL + " '" + DPullerPrice.Text + "', "        '�ޤ��ʤJ���
                    SQL = SQL + " '" + DMoldName.Text + "', "           '�Ҩ�W��
                    SQL = SQL + " '" + DMoldQty.Text + "', "            '�ҫ���
                    SQL = SQL + " '" + DMoldPoint.Text + "', "          '�ި���
                    If DQuality1.Text <> "" Then                        '�~����R-1(��ܥ�)
                        SQL = SQL + " '1', "
                    Else
                        SQL = SQL + " '0', "
                    End If
                    If DQuality2.Text <> "" Then                        '�~����R-2(��ܥ�)
                        SQL = SQL + " '1', "
                    Else
                        SQL = SQL + " '0', "
                    End If

                    If DQAAttachFile.Visible Then                      '�~����R����
                        If DQAAttachFile.PostedFile.FileName <> "" Then     '�P�_���ɮפW��
                            FileName = CStr(wFormSno) & "-" & "QA" & "-" & UploadDateTime & "-" & Right(DQAAttachFile.PostedFile.FileName, InStr(StrReverse(DQAAttachFile.PostedFile.FileName), "\") - 1)
                            Try    '�W�ǹ���
                                DQAAttachFile.PostedFile.SaveAs(Path + FileName)
                            Catch ex As Exception
                            End Try
                        Else
                            FileName = wQAAttachFile.Text
                        End If
                    Else
                        FileName = wQAAttachFile.Text
                    End If
                    SQL = SQL + " '" + FileName + "', "

                    If DSampleFile.Visible Then                      '�˫~����
                        If DSampleFile.PostedFile.FileName <> "" Then     '�P�_���ɮפW��
                            FileName = CStr(wFormSno) & "-" & "Sample" & "-" & UploadDateTime & "-" & Right(DSampleFile.PostedFile.FileName, InStr(StrReverse(DSampleFile.PostedFile.FileName), "\") - 1)
                            Try    '�W�ǹ���
                                DSampleFile.PostedFile.SaveAs(Path + FileName)
                            Catch ex As Exception
                            End Try
                        Else
                            FileName = wSampleFile.Text
                        End If
                    Else
                        FileName = wSampleFile.Text
                    End If
                    SQL = SQL + " '" + FileName + "', "

                    SQL = SQL + " '" + Request.Cookies("UserID").Value + "', "     '�@����
                    SQL = SQL + " '" + NowDateTime + "', "       '�@���ɶ�
                    SQL = SQL + " '" + "" + "', "                       '�ק��
                    SQL = SQL + " '" + NowDateTime + "' "       '�ק�ɶ�
                    SQL = SQL + " ) "
                    OleDBCommand1.Connection = OleDbConnection1
                    OleDBCommand1.CommandText = SQL
                    OleDbConnection1.Open()
                    OleDBCommand1.ExecuteNonQuery()
                    OleDbConnection1.Close()

                    'Slider Table�B�z
                    oPutSubFile = Server.CreateObject("MSubFile.SubFile")
                    RtnCode = oPutSubFile.PutSlider(DSliderCode.Text, DSliderGRCode.Text, "900003", NewFormSno, Request.Cookies("UserID").Value)
                    'Sample Table�B�z
                    RtnCode = oPutSubFile.PutSample(DSample.Text, "900003", NewFormSno, Request.Cookies("UserID").Value)
                    'Price Table�B�z
                    RtnCode = oPutSubFile.PutPrice(DPrice.Text, "900003", NewFormSno, Request.Cookies("UserID").Value)
                    'QA1 Table�B�z
                    RtnCode = oPutSubFile.PutQA1(DQuality1.Text, "900003", NewFormSno, Request.Cookies("UserID").Value)
                    'QA2 Table�B�z
                    RtnCode = oPutSubFile.PutQA2(DQuality2.Text, "900003", NewFormSno, Request.Cookies("UserID").Value)
                End If
            End If
        Else
            If ErrCode = 9110 Then Message = "���o���y�����p�ⲧ�`,�нT�{�γs���t�ΤH��!"
            If ErrCode = 9010 Then Message = "�W���ɮ׮榡���~,�нT�{!"
            If ErrCode = 9020 Then Message = "�W���ɮ�Size�W�L1024KB,�нT�{!"
            If ErrCode = 9030 Then Message = "�W���ɮרS���w,�нT�{!"
            If ErrCode = 9210 Then Message = "����z�Ѭ���L�ɻݶ�g����,�нT�{!"
            Response.Write(YKK.ShowMessage(Message))
        End If      '�W���ɮ�ErrCode=0

        If ErrCode = 0 Then
            If wFormSno > 0 And wStep > 1 Then    '�P�_�O�_[ñ��]
                Response.Write(String.Format("<script>window.close();</script>"))
            Else
                Dim SQL As String
                Dim DBDataSet1 As New DataSet
                'DB�s���]�w
                Dim OleDbConnection1 As New OleDbConnection
                OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
                OleDbConnection1.Open()

                SQL = "Select FormNo, FormSno From F_ModManufSheet "
                SQL = SQL & " Where Sts = 0 "
                SQL = SQL & "   And FormNo =  '900003' "
                SQL = SQL & "   And FormSno =  '" & CStr(NewFormSno) & "'"
                SQL = SQL & " Order by Unique_ID Desc "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "F_ModManufSheet")
                If DBDataSet1.Tables("F_ModManufSheet").Rows.Count > 0 Then
                    rFormNo = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("FormNo")
                    rFormSno = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("FormSno")
                End If
                'DB�s������
                OleDbConnection1.Close()
                Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.close();</script>", "Form1.DNFormNo", rFormNo, "Form1.DNFormSno", rFormSno))
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �ˬd�W���ɮ�
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As HtmlInputFile) As Integer
        Dim fileExtension As String     '�ŧi�@���ܼƦs���ɮ׮榡(���ɦW)
        Dim allowedExtensions As String() = {".jpg", ".jpeg", ".gif", ".xls", ".doc"}   '�w�q���\���ɮ׮榡
        Dim i As Integer

        UPFileIsNormal = 0

        fileExtension = IO.Path.GetExtension(UPFile.PostedFile.FileName).ToLower   '���o�ɮ׮榡
        For i = 0 To allowedExtensions.Length - 1           '�v�@�ˬd���\���榡���O�_���W�Ǫ��榡
            If fileExtension = allowedExtensions(i) Then
                UPFileIsNormal = 0
                Exit For
            Else
                UPFileIsNormal = 9010
            End If
        Next
        'Check�W���ɮ�Size
        If UPFileIsNormal = 0 Then
            If UPFile.PostedFile.ContentLength <= 1000 * 1024 Then
                UPFileIsNormal = 0
            Else
                UPFileIsNormal = 9020
            End If
        End If

        'If UPFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
        'Check�W���ɮ׮榡
        'Else
        'UPFileIsNormal = 9030
        'End If
    End Function

End Class
