Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ImportModSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DImportSheet As System.Web.UI.WebControls.Image
    Protected WithEvents Label22 As System.Web.UI.WebControls.Label
    Protected WithEvents Label21 As System.Web.UI.WebControls.Label
    Protected WithEvents Label20 As System.Web.UI.WebControls.Label
    Protected WithEvents Label19 As System.Web.UI.WebControls.Label
    Protected WithEvents Label18 As System.Web.UI.WebControls.Label
    Protected WithEvents Label17 As System.Web.UI.WebControls.Label
    Protected WithEvents DBuyer As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BFlow As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DPriceJ1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceI3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceI2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceI1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceH3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceH2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceJ3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceJ2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceA3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceA2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceH1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceG3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceG2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceG1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceF3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceF2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceF1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceE3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceE2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceE1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceD3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceD2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceD1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceC3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceC2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceC1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceB3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceB2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceB1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceA1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSliderCode As System.Web.UI.WebControls.Button
    Protected WithEvents LRefFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents BDate As System.Web.UI.WebControls.Button
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSliderPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDevReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DManufPlace As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSliderType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BSpec As System.Web.UI.WebControls.Button
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderGRCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents LMapFile As System.Web.UI.WebControls.Image
    Protected WithEvents BPrint As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DMapFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DRefFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton

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
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܼ�
    '**
    '*****************************************************************
    Dim FieldName(60) As String     '�U���
    Dim Attribute(60) As Integer    '�U����ݩ�    
    Dim Top As Integer              '�ʺA����Top��m
    Dim wFormNo As String           '��渹�X
    Dim wFormSno As Integer         '���y����
    Dim wOFormNo As String          '���渹�X
    Dim wOFormSno As Integer        '����y����
    Dim rFormNo As String           '�Ǧ^��渹�X
    Dim rFormSno As Integer         '�Ǧ^���y����
    Dim wStep As Integer            '�u�{�N�X
    Dim wSeqNo As Integer           '�Ǹ�
    Dim wApplyID As String          '�ӽЪ�ID
    Dim wAgentID As String          '�Q�N�z�HID
    Dim NowDateTime As String       '�{�b����ɶ�
    Dim wNextGate As String         '�U�@���H
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    'Modify-Start by Joy 2009/11/20(2010��ƾ����)
    'Dim wDepo As String = "CL"      '���c��ƾ�(CL->���c, TP->�x�_)
    '
    '�s�զ�ƾ�
    Dim wDepo As String = "CL1"      '���c��ƾ�(TP1->�x�_-����, TP2->�x�_-�ا�, CL1->���c-����, CL2->���c-�s�y)
    'Modify-End

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
            'If wFormSno > 0 And wStep > 2 Then    '�P�_�O�_�����
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
                      CStr(DateTime.Now.Second)     '�{�b���
        wFormNo = Request.QueryString("pFormNo")    '��渹�X
        wFormSno = Request.QueryString("pFormSno")  '���y����
        wOFormNo = Request.QueryString("pOFormNo")    '���渹�X
        wOFormSno = Request.QueryString("pOFormSno")  '����y����
        wStep = Request.QueryString("pStep")        '�u�{�N�X
        wSeqNo = Request.QueryString("pSeqNo")      '�Ǹ�
        wApplyID = Request.QueryString("pApplyID")  '�ӽЪ�ID
        wAgentID = Request.QueryString("pAgentID")  '�Q�N�z�HID
        Response.Cookies("MapNo").Value = ""        '�ϸ�, MapPicker�ϥ�

        Response.Cookies("PGM").Value = "ImportModSheet_01.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '���y����
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '�u�{�N�X
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     '�W�Ǹ���ˬd����ܰT��
    '**
    '*****************************************************************
    Sub ShowMessage()
        Dim Message As String = ""
        'Check�ϼ�
        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then
                Message = "�ϼ�"
            End If
        End If
        'Check��L����
        If DRefFile.Visible Then
            If DRefFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "��L����"
                Else
                    Message = Message & ", " & "��L����"
                End If
            End If
        End If

        If Message <> "" Then
            Message = "�U�C�ҳ]�w�����[�ɮױN�|�� (" & Message & ") " & ",�Э��s�]�w!"
            Response.Write(YKK.ShowMessage(Message))
        End If
    End Sub

    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ImportFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_ImportModSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ImportModSheet")
        If DBDataSet1.Tables("F_ImportModSheet").Rows.Count > 0 Then

            If DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("MapFile") <> "" Then          '�ϼ�
                LMapFile.ImageUrl = Path & DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("MapFile")
            Else
                LMapFile.Visible = False
            End If

            DNo.Text = DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("Date")                   '���
            SetFieldData("Division", DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("Division"))  '����
            SetFieldData("Person", DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("Person"))      '���
            DSliderCode.Text = DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("SliderCode")       'Slider Code
            DSliderGRCode.Text = DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("SliderGRCode")   'Slider Group Code
            DSpec.Text = DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("Spec")                   '�W��
            SetFieldData("SliderType", DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("SliderType"))    '���Y����
            SetFieldData("ManufPlace", DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("ManufPlace"))      '�Ͳ��a
            SetFieldData("Buyer", DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("Buyer"))    'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("SellVendor")   '�e�U�t��
            DDevReason.Text = DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("Reason")     '�}�o�z��

            DSliderPrice.Text = DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("SliderPrice") '���Y�ʤJ��
            If DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("Price") = 1 Then               '���
                Dim DBTable1 As DataTable
                SQL = "Select * From F_ManufCOPrice "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet1, "ManufCOPrice")
                DBTable1 = DBDataSet1.Tables("ManufCOPrice")
                For i = 0 To DBTable1.Rows.Count - 1
                    Select Case DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Seqno")
                        Case 1
                            DPriceA1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceA2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceA3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 2
                            DPriceB1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceB2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceB3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 3
                            DPriceC1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceC2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceC3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 4
                            DPriceD1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceD2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceD3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 5
                            DPriceE1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceE2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceE3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 6
                            DPriceF1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceF2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceF3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 7
                            DPriceG1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceG2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceG3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 8
                            DPriceH1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceH2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceH3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 9
                            DPriceI1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceI2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceI3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case Else
                            DPriceJ1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceJ2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceJ3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                    End Select
                Next
            End If

            If DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("RefFile") <> "" Then          '��L���� 
                LRefFile.NavigateUrl = Path & DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("RefFile")
            Else
                LRefFile.Visible = False
            End If

            DFormSno.Text = "�渹�G" & DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("FormSno")       '�渹
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
        Path = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ImportFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_ImportSheet "
        SQL = SQL & " Where Sts = 1 "
        SQL = SQL & "   And FormNo =  '" & wOFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ImportSheet")
        If DBDataSet1.Tables("F_ImportSheet").Rows.Count > 0 Then

            If DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("MapFile") <> "" Then          '�ϼ�
                LMapFile.ImageUrl = Path & DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("MapFile")
            Else
                LMapFile.Visible = False
            End If

            DNo.Text = DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("No")                       'No
            SetFieldData("Date", DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Date"))                   '���
            SetFieldData("Division", DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Division"))  '����
            SetFieldData("Person", DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Person"))      '���
            DSliderCode.Text = DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SliderCode")       'Slider Code
            DSliderGRCode.Text = DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SliderGRCode")   'Slider Group Code
            DSpec.Text = DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Spec")                   '�W��
            SetFieldData("SliderType", DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SliderType"))    '���Y����
            SetFieldData("ManufPlace", DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("ManufPlace"))      '�Ͳ��a
            SetFieldData("Buyer", DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Buyer"))    'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SellVendor")   '�e�U�t��
            DDevReason.Text = DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Reason")     '�}�o�z��

            DSliderPrice.Text = DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SliderPrice") '���Y�ʤJ��
            If DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Price") = 1 Then               '���
                Dim DBTable1 As DataTable
                SQL = "Select * From F_ManufCOPrice "
                SQL = SQL & " Where FormNo =  '" & wOFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet1, "ManufCOPrice")
                DBTable1 = DBDataSet1.Tables("ManufCOPrice")
                For i = 0 To DBTable1.Rows.Count - 1
                    Select Case DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Seqno")
                        Case 1
                            DPriceA1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceA2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceA3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 2
                            DPriceB1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceB2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceB3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 3
                            DPriceC1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceC2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceC3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 4
                            DPriceD1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceD2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceD3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 5
                            DPriceE1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceE2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceE3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 6
                            DPriceF1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceF2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceF3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 7
                            DPriceG1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceG2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceG3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 8
                            DPriceH1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceH2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceH3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 9
                            DPriceI1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceI2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceI3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case Else
                            DPriceJ1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceJ2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceJ3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                    End Select
                Next
            End If

            If DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("RefFile") <> "" Then          '��L���� 
                LRefFile.NavigateUrl = Path & DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("RefFile")
            Else
                LRefFile.Visible = False
            End If

            DFormSno.Text = "�渹�G" & CStr(wOFormSno)       '�渹
        End If
        'DB�s������
        OleDbConnection1.Close()

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

        'If wFormSno > 0 And wStep > 2 Then      '�P�_�O�_[ñ��]
        If wFormSno > 0 Then    '�P�_�O�_�����
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                'Sheet���
                DImportSheet.Visible = True     '���Sheet-1
                '�s�����---�ݦA�ק�
                LMapFile.Visible = True         '�ϼ�
                LRefFile.Visible = True         '��L����
                '���s��m
                BSAVE.Style.Add("Top", Top)    '�x�s���s
                BOK.Style.Add("Top", Top)      '�\Ū����
                DFormSno.Style.Add("Top", Top) '�渹
            End If
        Else
            'Sheet���
            DImportSheet.Visible = True   '���Sheet-1
            '�s�����---�ݦA�ק�
            LMapFile.Visible = True         '�ϼ�
            LRefFile.Visible = True         '��L����
            '���s��m
            BSAVE.Style.Add("Top", Top)    '�x�s���s
            BOK.Style.Add("Top", Top)      '�\Ū����
            DFormSno.Style.Add("Top", Top) '�渹
        End If

        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
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
                BOK.Visible = False
                BSAVE.Visible = True
                BSAVE.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("SaveDesc")
                '-- ������s���G�� JOY 16/08/17
                '��
                'BSAVE.Attributes("onclick") = "Button('SAVE', '" + BSAVE.Value + "');"
                '�s
                BSAVE.Attributes("onclick") = "this.disabled = true;" & "Button('SAVE', '" + BSAVE.Value + "');"
            Else
                BOK.Visible = True
                BSAVE.Visible = False
                '-- ������s���G�� JOY 16/08/17
                '��
                'BOK.Attributes("onclick") = "Button('OK', '�\Ū����');"
                '�s
                BOK.Attributes("onclick") = "this.disabled = true;" & "Button('OK', '�\Ū����');"
            End If
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
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     ���s��RequestedField��Top��m
    '**
    '*****************************************************************
    Sub TopPosition()
        Top = 656
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
    '**     ��������ܤ�����J�ˬd
    '**
    '*****************************************************************
    Sub ShowSheetField(ByVal pPost As String)
        '�ظm�����ݩʰ}�C
        oCommon.GetFieldAttribute(wFormNo, wStep, FieldName, Attribute)
        '��渹�X,�u�{���d���X,���W,����ݩ�

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
                DNo.BackColor = Color.LightGray
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
                DDate.BackColor = Color.LightGray
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
                DDivision.BackColor = Color.LightGray
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
                DPerson.BackColor = Color.LightGray
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
                DSliderGRCode.BackColor = Color.LightGray
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
                DSliderCode.BackColor = Color.LightGray
                DSliderCode.ReadOnly = True
                DSliderCode.Visible = True
                BSliderCode.Visible = False
            Case 1  '�ק�+�ˬd
                DSliderCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderCodeRqd", "DSliderCode", "���`�G�ݿ�J Slider Code")
                DSliderCode.Visible = True
                BSliderCode.Visible = True
            Case 2  '�ק�
                DSliderCode.BackColor = Color.Yellow
                DSliderCode.Visible = True
                BSliderCode.Visible = True
            Case Else   '����
                DSliderCode.Visible = False
                BSliderCode.Visible = False
        End Select
        If pPost = "New" Then DSliderCode.Text = ""

        'Spec
        Select Case FindFieldInf("Spec")
            Case 0  '���
                DSpec.BackColor = Color.LightGray
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

        '���Y����(���s.�~�`...)
        Select Case FindFieldInf("SliderType")
            Case 0  '���
                DSliderType.BackColor = Color.LightGray
                DSliderType.Visible = True
            Case 1  '�ק�+�ˬd
                DSliderType.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderTypeRqd", "DSliderType", "���`�G�ݿ�J���Y����")
                DSliderType.Visible = True
            Case 2  '�ק�
                DSliderType.BackColor = Color.Yellow
                DSliderType.Visible = True
            Case Else   '����
                DSliderType.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SliderType", "ZZZZZZ")

        '�Ͳ��a
        Select Case FindFieldInf("ManufPlace")
            Case 0  '���
                DManufPlace.BackColor = Color.LightGray
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

        '�e�U�t��
        Select Case FindFieldInf("SellVendor")
            Case 0  '���
                DSellVendor.BackColor = Color.LightGray
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
                DBuyer.BackColor = Color.LightGray
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
        If pPost = "New" Then SetFieldData("Buyer", "ZZZZZZ")

        '�}�o�z��
        Select Case FindFieldInf("Reason")
            Case 0  '���
                DDevReason.BackColor = Color.LightGray
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

        '���Y�ʤJ���
        Select Case FindFieldInf("SliderPrice")
            Case 0  '���
                DSliderPrice.BackColor = Color.LightGray
                DSliderPrice.ReadOnly = True
                DSliderPrice.Visible = True
            Case 1  '�ק�+�ˬd
                DSliderPrice.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderPriceRqd", "DSliderPrice", "���`�G�ݿ�J���Y�ʤJ���")
                DSliderPrice.Visible = True
            Case 2  '�ק�
                DSliderPrice.BackColor = Color.Yellow
                DSliderPrice.Visible = True
            Case Else   '����
                DSliderPrice.Visible = False
        End Select
        If pPost = "New" Then DSliderPrice.Text = ""

        '�ϼ�
        Select Case FindFieldInf("MapFile")
            Case 0  '���
                DMapFile.Visible = False
                DMapFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DMapFileRqd", "DMapFile", "���`�G�ݿ�J����")
                DMapFile.Visible = True
                DMapFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DMapFile.Visible = True
                DMapFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DMapFile.Visible = False
        End Select

        '��L����
        Select Case FindFieldInf("RefFile")
            Case 0  '���
                DRefFile.Visible = False
                DRefFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DRefFileRqd", "DRefFile", "���`�G�ݿ�J��L����")
                DRefFile.Visible = True
                DRefFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DRefFile.Visible = True
                DRefFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DRefFile.Visible = False
        End Select

        'Price
        Select Case FindFieldInf("Price")
            Case 0  '���
                DPriceA1.BackColor = Color.LightGray
                DPriceA2.BackColor = Color.LightGray
                DPriceA3.BackColor = Color.LightGray
                DPriceB1.BackColor = Color.LightGray
                DPriceB2.BackColor = Color.LightGray
                DPriceB3.BackColor = Color.LightGray
                DPriceC1.BackColor = Color.LightGray
                DPriceC2.BackColor = Color.LightGray
                DPriceC3.BackColor = Color.LightGray
                DPriceD1.BackColor = Color.LightGray
                DPriceD2.BackColor = Color.LightGray
                DPriceD3.BackColor = Color.LightGray
                DPriceE1.BackColor = Color.LightGray
                DPriceE2.BackColor = Color.LightGray
                DPriceE3.BackColor = Color.LightGray
                DPriceF1.BackColor = Color.LightGray
                DPriceF2.BackColor = Color.LightGray
                DPriceF3.BackColor = Color.LightGray
                DPriceG1.BackColor = Color.LightGray
                DPriceG2.BackColor = Color.LightGray
                DPriceG3.BackColor = Color.LightGray
                DPriceH1.BackColor = Color.LightGray
                DPriceH2.BackColor = Color.LightGray
                DPriceH3.BackColor = Color.LightGray
                DPriceI1.BackColor = Color.LightGray
                DPriceI2.BackColor = Color.LightGray
                DPriceI3.BackColor = Color.LightGray
                DPriceJ1.BackColor = Color.LightGray
                DPriceJ2.BackColor = Color.LightGray
                DPriceJ3.BackColor = Color.LightGray

                DPriceA1.ReadOnly = True
                DPriceA3.ReadOnly = True
                DPriceB1.ReadOnly = True
                DPriceB3.ReadOnly = True
                DPriceC1.ReadOnly = True
                DPriceC3.ReadOnly = True
                DPriceD1.ReadOnly = True
                DPriceD3.ReadOnly = True
                DPriceE1.ReadOnly = True
                DPriceE3.ReadOnly = True
                DPriceF1.ReadOnly = True
                DPriceF3.ReadOnly = True
                DPriceG1.ReadOnly = True
                DPriceG3.ReadOnly = True
                DPriceH1.ReadOnly = True
                DPriceH3.ReadOnly = True
                DPriceI1.ReadOnly = True
                DPriceI3.ReadOnly = True
                DPriceJ1.ReadOnly = True
                DPriceJ3.ReadOnly = True

                DPriceA1.Visible = True
                DPriceA2.Visible = True
                DPriceA3.Visible = True
                DPriceB1.Visible = True
                DPriceB2.Visible = True
                DPriceB3.Visible = True
                DPriceC1.Visible = True
                DPriceC2.Visible = True
                DPriceC3.Visible = True
                DPriceD1.Visible = True
                DPriceD2.Visible = True
                DPriceD3.Visible = True
                DPriceE1.Visible = True
                DPriceE2.Visible = True
                DPriceE3.Visible = True
                DPriceF1.Visible = True
                DPriceF2.Visible = True
                DPriceF3.Visible = True
                DPriceG1.Visible = True
                DPriceG2.Visible = True
                DPriceG3.Visible = True
                DPriceH1.Visible = True
                DPriceH2.Visible = True
                DPriceH3.Visible = True
                DPriceI1.Visible = True
                DPriceI2.Visible = True
                DPriceI3.Visible = True
                DPriceJ1.Visible = True
                DPriceJ2.Visible = True
                DPriceJ3.Visible = True
            Case 1  '�ק�+�ˬd
                DPriceA1.BackColor = Color.GreenYellow
                DPriceA2.BackColor = Color.GreenYellow
                DPriceA3.BackColor = Color.GreenYellow
                DPriceB1.BackColor = Color.GreenYellow
                DPriceB2.BackColor = Color.GreenYellow
                DPriceB3.BackColor = Color.GreenYellow
                DPriceC1.BackColor = Color.GreenYellow
                DPriceC2.BackColor = Color.GreenYellow
                DPriceC3.BackColor = Color.GreenYellow
                DPriceD1.BackColor = Color.GreenYellow
                DPriceD2.BackColor = Color.GreenYellow
                DPriceD3.BackColor = Color.GreenYellow
                DPriceE1.BackColor = Color.GreenYellow
                DPriceE2.BackColor = Color.GreenYellow
                DPriceE3.BackColor = Color.GreenYellow
                DPriceF1.BackColor = Color.GreenYellow
                DPriceF2.BackColor = Color.GreenYellow
                DPriceF3.BackColor = Color.GreenYellow
                DPriceG1.BackColor = Color.GreenYellow
                DPriceG2.BackColor = Color.GreenYellow
                DPriceG3.BackColor = Color.GreenYellow
                DPriceH1.BackColor = Color.GreenYellow
                DPriceH2.BackColor = Color.GreenYellow
                DPriceH3.BackColor = Color.GreenYellow
                DPriceI1.BackColor = Color.GreenYellow
                DPriceI2.BackColor = Color.GreenYellow
                DPriceI3.BackColor = Color.GreenYellow
                DPriceJ1.BackColor = Color.GreenYellow
                DPriceJ2.BackColor = Color.GreenYellow
                DPriceJ3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPriceA1Rqd", "DPriceA1", "���`�G�ݿ�J�����T(A1)")
                ShowRequiredFieldValidator("DPriceA2Rqd", "DPriceA2", "���`�G�ݿ�J�����T(A2)")
                ShowRequiredFieldValidator("DPriceA3Rqd", "DPriceA3", "���`�G�ݿ�J�����T(A3)")
                ShowRequiredFieldValidator("DPriceB1Rqd", "DPriceB1", "���`�G�ݿ�J�����T(B1)")
                ShowRequiredFieldValidator("DPriceB2Rqd", "DPriceB2", "���`�G�ݿ�J�����T(B2)")
                ShowRequiredFieldValidator("DPriceB3Rqd", "DPriceB3", "���`�G�ݿ�J�����T(B3)")
                ShowRequiredFieldValidator("DPriceC1Rqd", "DPriceC1", "���`�G�ݿ�J�����T(C1)")
                ShowRequiredFieldValidator("DPriceC2Rqd", "DPriceC2", "���`�G�ݿ�J�����T(C2)")
                ShowRequiredFieldValidator("DPriceC3Rqd", "DPriceC3", "���`�G�ݿ�J�����T(C3)")
                ShowRequiredFieldValidator("DPriceD1Rqd", "DPriceD1", "���`�G�ݿ�J�����T(D1)")
                ShowRequiredFieldValidator("DPriceD2Rqd", "DPriceD2", "���`�G�ݿ�J�����T(D2)")
                ShowRequiredFieldValidator("DPriceD3Rqd", "DPriceD3", "���`�G�ݿ�J�����T(D3)")
                ShowRequiredFieldValidator("DPriceE1Rqd", "DPriceE1", "���`�G�ݿ�J�����T(E1)")
                ShowRequiredFieldValidator("DPriceE2Rqd", "DPriceE2", "���`�G�ݿ�J�����T(E2)")
                ShowRequiredFieldValidator("DPriceE3Rqd", "DPriceE3", "���`�G�ݿ�J�����T(E3)")
                ShowRequiredFieldValidator("DPriceF1Rqd", "DPriceF1", "���`�G�ݿ�J�����T(F1)")
                ShowRequiredFieldValidator("DPriceF2Rqd", "DPriceF2", "���`�G�ݿ�J�����T(F2)")
                ShowRequiredFieldValidator("DPriceF3Rqd", "DPriceF3", "���`�G�ݿ�J�����T(F3)")
                ShowRequiredFieldValidator("DPriceG1Rqd", "DPriceG1", "���`�G�ݿ�J�����T(G1)")
                ShowRequiredFieldValidator("DPriceG2Rqd", "DPriceG2", "���`�G�ݿ�J�����T(G2)")
                ShowRequiredFieldValidator("DPriceG3Rqd", "DPriceG3", "���`�G�ݿ�J�����T(G3)")
                ShowRequiredFieldValidator("DPriceH1Rqd", "DPriceH1", "���`�G�ݿ�J�����T(H1)")
                ShowRequiredFieldValidator("DPriceH2Rqd", "DPriceH2", "���`�G�ݿ�J�����T(H2)")
                ShowRequiredFieldValidator("DPriceH3Rqd", "DPriceH3", "���`�G�ݿ�J�����T(H3)")
                ShowRequiredFieldValidator("DPriceI1Rqd", "DPriceI1", "���`�G�ݿ�J�����T(I1)")
                ShowRequiredFieldValidator("DPriceI2Rqd", "DPriceI2", "���`�G�ݿ�J�����T(I2)")
                ShowRequiredFieldValidator("DPriceI3Rqd", "DPriceI3", "���`�G�ݿ�J�����T(I3)")
                ShowRequiredFieldValidator("DPriceJ1Rqd", "DPriceJ1", "���`�G�ݿ�J�����T(J1)")
                ShowRequiredFieldValidator("DPriceJ2Rqd", "DPriceJ2", "���`�G�ݿ�J�����T(J2)")
                ShowRequiredFieldValidator("DPriceJ3Rqd", "DPriceJ3", "���`�G�ݿ�J�����T(J3)")
                DPriceA1.Visible = True
                DPriceA2.Visible = True
                DPriceA3.Visible = True
                DPriceB1.Visible = True
                DPriceB2.Visible = True
                DPriceB3.Visible = True
                DPriceC1.Visible = True
                DPriceC2.Visible = True
                DPriceC3.Visible = True
                DPriceD1.Visible = True
                DPriceD2.Visible = True
                DPriceD3.Visible = True
                DPriceE1.Visible = True
                DPriceE2.Visible = True
                DPriceE3.Visible = True
                DPriceF1.Visible = True
                DPriceF2.Visible = True
                DPriceF3.Visible = True
                DPriceG1.Visible = True
                DPriceG2.Visible = True
                DPriceG3.Visible = True
                DPriceH1.Visible = True
                DPriceH2.Visible = True
                DPriceH3.Visible = True
                DPriceI1.Visible = True
                DPriceI2.Visible = True
                DPriceI3.Visible = True
                DPriceJ1.Visible = True
                DPriceJ2.Visible = True
                DPriceJ3.Visible = True
            Case 2  '�ק�
                DPriceA1.BackColor = Color.Yellow
                DPriceA2.BackColor = Color.Yellow
                DPriceA3.BackColor = Color.Yellow
                DPriceB1.BackColor = Color.Yellow
                DPriceB2.BackColor = Color.Yellow
                DPriceB3.BackColor = Color.Yellow
                DPriceC1.BackColor = Color.Yellow
                DPriceC2.BackColor = Color.Yellow
                DPriceC3.BackColor = Color.Yellow
                DPriceD1.BackColor = Color.Yellow
                DPriceD2.BackColor = Color.Yellow
                DPriceD3.BackColor = Color.Yellow
                DPriceE1.BackColor = Color.Yellow
                DPriceE2.BackColor = Color.Yellow
                DPriceE3.BackColor = Color.Yellow
                DPriceF1.BackColor = Color.Yellow
                DPriceF2.BackColor = Color.Yellow
                DPriceF3.BackColor = Color.Yellow
                DPriceG1.BackColor = Color.Yellow
                DPriceG2.BackColor = Color.Yellow
                DPriceG3.BackColor = Color.Yellow
                DPriceH1.BackColor = Color.Yellow
                DPriceH2.BackColor = Color.Yellow
                DPriceH3.BackColor = Color.Yellow
                DPriceI1.BackColor = Color.Yellow
                DPriceI2.BackColor = Color.Yellow
                DPriceI3.BackColor = Color.Yellow
                DPriceJ1.BackColor = Color.Yellow
                DPriceJ2.BackColor = Color.Yellow
                DPriceJ3.BackColor = Color.Yellow
                DPriceA1.Visible = True
                DPriceA2.Visible = True
                DPriceA3.Visible = True
                DPriceB1.Visible = True
                DPriceB2.Visible = True
                DPriceB3.Visible = True
                DPriceC1.Visible = True
                DPriceC2.Visible = True
                DPriceC3.Visible = True
                DPriceD1.Visible = True
                DPriceD2.Visible = True
                DPriceD3.Visible = True
                DPriceE1.Visible = True
                DPriceE2.Visible = True
                DPriceE3.Visible = True
                DPriceF1.Visible = True
                DPriceF2.Visible = True
                DPriceF3.Visible = True
                DPriceG1.Visible = True
                DPriceG2.Visible = True
                DPriceG3.Visible = True
                DPriceH1.Visible = True
                DPriceH2.Visible = True
                DPriceH3.Visible = True
                DPriceI1.Visible = True
                DPriceI2.Visible = True
                DPriceI3.Visible = True
                DPriceJ1.Visible = True
                DPriceJ2.Visible = True
                DPriceJ3.Visible = True
            Case Else   '����
                DPriceA1.Visible = False
                DPriceA2.Visible = False
                DPriceA3.Visible = False
                DPriceB1.Visible = False
                DPriceB2.Visible = False
                DPriceB3.Visible = False
                DPriceC1.Visible = False
                DPriceC2.Visible = False
                DPriceC3.Visible = False
                DPriceD1.Visible = False
                DPriceD2.Visible = False
                DPriceD3.Visible = False
                DPriceE1.Visible = False
                DPriceE2.Visible = False
                DPriceE3.Visible = False
                DPriceF1.Visible = False
                DPriceF2.Visible = False
                DPriceF3.Visible = False
                DPriceG1.Visible = False
                DPriceG2.Visible = False
                DPriceG3.Visible = False
                DPriceH1.Visible = False
                DPriceH2.Visible = False
                DPriceH3.Visible = False
                DPriceI1.Visible = False
                DPriceI2.Visible = False
                DPriceI3.Visible = False
                DPriceJ1.Visible = False
                DPriceJ2.Visible = False
                DPriceJ3.Visible = False
        End Select
        If pPost = "New" Then
            DPriceA1.Text = ""
            DPriceA3.Text = ""
            DPriceB1.Text = ""
            DPriceB3.Text = ""
            DPriceC1.Text = ""
            DPriceC3.Text = ""
            DPriceD1.Text = ""
            DPriceD3.Text = ""
            DPriceE1.Text = ""
            DPriceE3.Text = ""
            DPriceF1.Text = ""
            DPriceF3.Text = ""
            DPriceG1.Text = ""
            DPriceG3.Text = ""
            DPriceH1.Text = ""
            DPriceH3.Text = ""
            DPriceI1.Text = ""
            DPriceI3.Text = ""
            DPriceJ1.Text = ""
            DPriceJ3.Text = ""
        End If

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
        Dim idx As Integer
        Dim i As Integer
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        If Left(pFieldName, 3) = "DQA" Then
            idx = FindFieldInf("Quality1")
        Else
            idx = FindFieldInf(pFieldName)
        End If

        '����
        If pFieldName = "Division" Then
            DDivision.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDivision.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select DivName From M_Users "
                SQL = SQL & " Where UserID = '" & Request.QueryString("pUserID") & "'"
                SQL = SQL & "   And Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("DivName")
                    ListItem1.Value = DBTable1.Rows(i).Item("DivName")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDivision.Items.Add(ListItem1)
                Next
            End If
        End If

        '���
        If pFieldName = "Person" Then
            DPerson.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPerson.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select UserName From M_Users "
                SQL = SQL & " Where UserID = '" & Request.QueryString("pUserID") & "'"
                SQL = SQL & "   And Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("UserName")
                    ListItem1.Value = DBTable1.Rows(i).Item("UserName")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPerson.Items.Add(ListItem1)
                Next
            End If
        End If

        'Buyer
        If pFieldName = "Buyer" Then
            DBuyer.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBuyer.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='700' and DKey='BUYER' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBuyer.Items.Add(ListItem1)
                Next
            End If
        End If

        '���Y����
        If pFieldName = "SliderType" Then
            DSliderType.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSliderType.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='225' and DKey='SLIDERTYPE' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSliderType.Items.Add(ListItem1)
                Next
            End If
        End If

        '�Ͳ��a
        If pFieldName = "ManufPlace" Then
            DManufPlace.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DManufPlace.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='200' and DKey='MANUFPLACE' Order by Data "
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
        While i <= 60 And Run
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
    '**     �x�s���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BSAVE_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSAVE.ServerClick
        If Request.Cookies("RunBSAVE").Value = True Then
            Dim ErrCode As Integer = 0   '�ŧi�@��Error�ΨӧP�O�O�_��ƥ��`�A�w�]��False
            Dim Message As String = ""
            Dim wQCLT As Integer = 0 'QC-L/T

            Dim NewFormSno As Integer = wFormSno    '���y����
            Dim RtnCode As Integer = 0

            'Check����-1Size�ή榡
            If ErrCode = 0 Then
                If DMapFile.Visible Then
                    If DMapFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                        ErrCode = UPFileIsNormal(DMapFile)
                    End If
                End If
            End If
            'Check�ѦҪ���-Size�ή榡
            If ErrCode = 0 Then
                If DRefFile.Visible Then
                    If DRefFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                        ErrCode = UPFileIsNormal(DRefFile)
                    End If
                End If
            End If

            '�x�s���
            If ErrCode = 0 Then
                'If wFormSno > 0 And wStep > 2 Then    '�P�_�O�_�����
                If wFormSno > 0 Then    '�P�_�O�_�����
                    ModifyData("MOD", wFormSno)
                Else
                    '���o���y����
                    RtnCode = oCommon.GetFormSeqNo(wFormNo, NewFormSno) '��渹�X, ���y����
                    If RtnCode <> 0 Then
                        ErrCode = 9110
                    Else
                        AppendData("ADD", NewFormSno)
                    End If
                End If
            Else
                If ErrCode = 9010 Then Message = "�W���ɮ׮榡���~,�нT�{!"
                If ErrCode = 9020 Then Message = "�W���ɮ�Size�W�L1024KB,�нT�{!"
                If ErrCode = 9030 Then Message = "�W���ɮרS���w,�нT�{!"
                If ErrCode = 9210 Then Message = "����z�Ѭ���L�ɻݶ�g����,�нT�{!"
                If ErrCode = 9050 Then Message = "QC L/T�ݬ����ļƦr,�нT�{!"
                Response.Write(YKK.ShowMessage(Message))
            End If      '�W���ɮ�ErrCode=0

            If ErrCode = 0 Then
                'If wFormSno > 0 And wStep > 2 Then      '�P�_�O�_[ñ��]
                If wFormSno > 0 Then    '�P�_�O�_�����
                    Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}';window.close();</script>", "Form1.DReady", "�w�\Ū", "Form1.DSliderCode", DSliderCode.Text))

                Else
                    Dim SQL As String
                    Dim DBDataSet1 As New DataSet
                    'DB�s���]�w
                    Dim OleDbConnection1 As New OleDbConnection
                    OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
                    OleDbConnection1.Open()

                    SQL = "Select FormNo, FormSno From F_ImportModSheet "
                    SQL = SQL & " Where Sts = 0 "
                    SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
                    SQL = SQL & "   And FormSno =  '" & CStr(NewFormSno) & "'"
                    SQL = SQL & " Order by Unique_ID Desc "
                    Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter1.Fill(DBDataSet1, "F_ImportModSheet")
                    If DBDataSet1.Tables("F_ImportModSheet").Rows.Count > 0 Then
                        rFormNo = DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("FormNo")
                        rFormSno = DBDataSet1.Tables("F_ImportModSheet").Rows(0).Item("FormSno")
                    End If
                    'DB�s������
                    OleDbConnection1.Close()

                    Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.opener.document.{4}.value = '{5:d}'; window.opener.document.{6}.value = '{7:d}'; window.close();</script>", "Form1.DNFormNo", rFormNo, "Form1.DNFormSno", rFormSno, "Form1.DReady", "�w�\Ū", "Form1.DSliderCode", DSliderCode.Text))
                End If
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �\Ū����
    '**
    '*****************************************************************
    Private Sub BOK_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.ServerClick
        Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.close();</script>", "Form1.DReady", "�w�\Ū"))

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     �s�W�����
    '**
    '*****************************************************************
    Sub AppendData(ByVal pFun As String, ByVal NewFormSno As Integer)
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("ImportFilePath"))
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim SQL As String
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_ImportSheet "
        SQL = SQL & " Where Sts = 1 "
        SQL = SQL & "   And FormNo =  '" & wOFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ImportSheet")
        If DBDataSet1.Tables("F_ImportSheet").Rows.Count > 0 Then
            Dim OleDbConnection2 As New OleDbConnection
            OleDbConnection2.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
            Dim OleDBCommand2 As New OleDbCommand

            SQL = "Insert into F_ImportModSheet "
            SQL = SQL + "( "
            SQL = SQL + "Sts, CompletedTime, FormNo, FormSNo, "         '1~4
            SQL = SQL + "No, Date, Division, Person, SliderCode, "      '5~9
            SQL = SQL + "SliderGRCode, Spec, MapFile, RefFile, "        '10~13
            SQL = SQL + "SliderType, ManufPlace, SellVendor, Buyer, "   '14~17
            SQL = SQL + "Reason, Price, SliderPrice, "                  '18~20
            SQL = SQL + "CreateUser, CreateTime, ModifyUser, ModifyTime "  '21~24
            SQL = SQL + ")  "
            SQL = SQL + "VALUES( "
            SQL = SQL + " '0', "                                '���A(0:����,1:�w��NG,2:�w��OK)
            SQL = SQL + " '" + NowDateTime + "', "              '���פ�
            SQL = SQL + " '" + wFormNo + "', "                  '���N��
            SQL = SQL + " '" + CStr(NewFormSno) + "', "           '���y����

            SQL = SQL + " '" + DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("No") + "', "        'NO
            SQL = SQL + " '" + DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Date") + "', "      '���
            SQL = SQL + " '" + DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Division") + "', "  '����
            SQL = SQL + " '" + DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Person") + "', "    '���
            SQL = SQL + " '" + DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SliderCode") + "', "     'Slider Code(��ܥ�)
            SQL = SQL + " '" + DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SliderGRCode") + "', "   'Slider Group Code
            SQL = SQL + " '" + DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Spec") + "', "      '�W��
            SQL = SQL + " '" + DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("MapFile") + "', "   '����1
            SQL = SQL + " '" + DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("RefFile") + "', "   '����2
            SQL = SQL + " '" + DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SliderType") + "', "  '���Y����
            SQL = SQL + " '" + DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("ManufPlace") + "', "   '�Ͳ��a
            SQL = SQL + " '" + DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SellVendor") + "', "       '�e�U�t��
            SQL = SQL + " '" + DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Buyer") + "', "            'Buyer
            SQL = SQL + " '" + DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Reason") + "', "    '�}�o�z��
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Price")) + "', "   'Price(��ܥ�)
            SQL = SQL + " '" + DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SliderPrice") + "', "    '���Y�ʤJ���

            SQL = SQL + " '" + Request.QueryString("pUserID") + "', "     '�@����
            SQL = SQL + " '" + NowDateTime + "', "       '�@���ɶ�
            SQL = SQL + " '" + "" + "', "                       '�ק��
            SQL = SQL + " '" + NowDateTime + "' "       '�ק�ɶ�
            SQL = SQL + " ) "
            OleDBCommand2.Connection = OleDbConnection2
            OleDBCommand2.CommandText = SQL
            OleDbConnection2.Open()
            OleDBCommand2.ExecuteNonQuery()
            OleDbConnection2.Close()
        End If
        'DB�s������
        OleDbConnection1.Close()

        ModifyData(pFun, NewFormSno)
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     ��s�����
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pFormSno As Integer)
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("ImportFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQL As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand
        Dim RtnCode As Integer = 0

        OleDbConnection1.Open()
        SQL = "Update F_ImportModSheet Set "
        SQL = SQL + " No = '" & YKK.ReplaceString(DNo.Text) & "',"
        SQL = SQL + " Date = '" & DDate.Text & "',"
        SQL = SQL + " Division = '" & DDivision.SelectedValue & "',"
        SQL = SQL + " Person = '" & DPerson.SelectedValue & "',"
        SQL = SQL + " SliderCode = '" & YKK.ReplaceString(DSliderCode.Text) & "',"
        SQL = SQL + " SliderGRCode = '" & YKK.ReplaceString(DSliderGRCode.Text) & "',"
        SQL = SQL + " Spec = '" & YKK.ReplaceString(DSpec.Text) & "',"

        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Map" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "Map" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DMapFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQL = SQL + " MapFile = N'" & FileName & "',"
            End If
        End If

        If DRefFile.Visible Then
            If DRefFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Ref" & "-" & UploadDateTime & "-" & Right(DRefFile.PostedFile.FileName, InStr(StrReverse(DRefFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "Ref" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DRefFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DRefFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQL = SQL + " RefFile = N'" & FileName & "',"
            End If
        End If

        SQL = SQL + " SliderType = '" & DSliderType.SelectedValue & "',"
        SQL = SQL + " ManufPlace = '" & DManufPlace.SelectedValue & "',"
        SQL = SQL + " SellVendor = N'" & YKK.ReplaceString(DSellVendor.Text) & "',"
        SQL = SQL + " Buyer = N'" & DBuyer.SelectedValue & "',"
        SQL = SQL + " Reason = N'" & YKK.ReplaceString(DDevReason.Text) & "',"

        If DPriceA1.Text <> "" Then
            SQL = SQL + " Price = '" & "1" & "',"
        Else
            SQL = SQL + " Price = '" & "0" & "',"
        End If

        SQL = SQL + " SliderPrice = '" & YKK.ReplaceString(DSliderPrice.Text) & "',"

        SQL = SQL + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
        SQL = SQL + " Where FormNo  =  '" & wFormNo & "'"
        SQL = SQL + "   And FormSno =  '" & CStr(pFormSno) & "'"
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQL
        OleDBCommand1.ExecuteNonQuery()
        '
        'Modify-Start by Joy 09-10-03
        'Price Table�B�z
        Dim i As Integer
        For i = 1 To 10
            SQL = "Insert into F_ManufCOPrice "
            SQL = SQL + "( "
            SQL = SQL + "Sts, CompletedTime, FormNo, FormSNo, SeqNo, "
            SQL = SQL + "Spec, Currency, Price, "
            SQL = SQL + "CreateUser, CreateTime, ModifyUser, ModifyTime "
            SQL = SQL + ")  "
            SQL = SQL + "VALUES( "
            SQL = SQL + " '0', "
            SQL = SQL + " '" + NowDateTime + "', "
            SQL = SQL + " '" + wFormNo + "', "
            SQL = SQL + " '" + CStr(pFormSno) + "', "
            SQL = SQL + " '" + CStr(i) + "', "
            Select Case i
                Case Is = 1
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceA1.Text) + "', "
                    SQL = SQL + " '" + DPriceA2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceA3.Text) + "', "
                Case Is = 2
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceB1.Text) + "', "
                    SQL = SQL + " '" + DPriceB2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceB3.Text) + "', "
                Case Is = 3
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceC1.Text) + "', "
                    SQL = SQL + " '" + DPriceC2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceC3.Text) + "', "
                Case Is = 4
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceD1.Text) + "', "
                    SQL = SQL + " '" + DPriceD2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceD3.Text) + "', "
                Case Is = 5
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceE1.Text) + "', "
                    SQL = SQL + " '" + DPriceE2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceE3.Text) + "', "
                Case Is = 6
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceF1.Text) + "', "
                    SQL = SQL + " '" + DPriceF2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceF3.Text) + "', "
                Case Is = 7
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceG1.Text) + "', "
                    SQL = SQL + " '" + DPriceG2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceG3.Text) + "', "
                Case Is = 8
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceH1.Text) + "', "
                    SQL = SQL + " '" + DPriceH2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceH3.Text) + "', "
                Case Is = 9
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceI1.Text) + "', "
                    SQL = SQL + " '" + DPriceI2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceI3.Text) + "', "
                Case Is = 10
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceJ1.Text) + "', "
                    SQL = SQL + " '" + DPriceJ2.SelectedValue + "', "
                    SQL = SQL + " '" + YKK.ReplaceString(DPriceJ3.Text) + "', "
            End Select
            SQL = SQL + " '" + Request.QueryString("pUserID") + "', "     '�@����
            SQL = SQL + " '" + NowDateTime + "', "       '�@���ɶ�
            SQL = SQL + " '" + "" + "', "                       '�ק��
            SQL = SQL + " '" + NowDateTime + "' "       '�ק�ɶ�
            SQL = SQL + " ) "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDBCommand1.ExecuteNonQuery()
        Next
        OleDbConnection1.Close()
        'Modify-end by Joy
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �ˬd�W���ɮ�
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As HtmlInputFile) As Integer
        Dim fileExtension As String     '�ŧi�@���ܼƦs���ɮ׮榡(���ɦW)
        Dim allowedExtensions As String() = {".bmp", ".jpg", ".jpeg", ".gif", ".pdf", ".ai", ".xls", ".doc", ".ppt"}   '�w�q���\���ɮ׮榡
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
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �C�L���
    '**
    '*****************************************************************
    Private Sub BPrint_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim URL As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()
        SQL = "Select ViewURL From V_WaitHandle_01 "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "URL")
        If DBDataSet1.Tables("URL").Rows.Count > 0 Then
            URL = DBDataSet1.Tables("URL").Rows(0).Item("ViewURL")
        End If
        'DB�s������
        OleDbConnection1.Close()

        'Call JavaScript
        Dim scriptString As String = "<script language=JavaScript> "
        scriptString = scriptString & "OpenPrintSheet('" & URL & "'); "
        scriptString = scriptString & "</script>"
        If (Not Me.IsStartupScriptRegistered("Startup")) Then
            Me.RegisterStartupScript("Startup", scriptString)
        End If
    End Sub

End Class
