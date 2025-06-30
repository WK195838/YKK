Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ImportSheet_02
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
    Protected WithEvents LRefFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSliderPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDevReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DManufPlace As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSliderType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderGRCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents LMapFile As System.Web.UI.WebControls.Image
    Protected WithEvents LContact As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOPContact As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LColorAppend As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LSurface As System.Web.UI.WebControls.HyperLink
    Protected WithEvents Panel1 As System.Web.UI.WebControls.Panel
    Protected WithEvents LSliderDetail As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DStatus As System.Web.UI.WebControls.TextBox

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
        Response.Cookies("PGM").Value = "ImportSheet_02.aspx"

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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ImportFilePath")
        Dim PathOld1 As String = ConfigurationSettings.AppSettings.Get("HttpOld")
        Dim PathOld2 As String = ConfigurationSettings.AppSettings.Get("ImportFilePath")
        Dim RtnCode As Integer = 0
        Dim i As Integer
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select *,replace(slidergrcode ,'+','%')NewSliderGRCode From F_ImportSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ImportSheet")
        If DBDataSet1.Tables("F_ImportSheet").Rows.Count > 0 Then

            If DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("OPSts") = 1 Or _
               DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("CTSts") = 1 Or _
               DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SDSts") = 1 Or _
               DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SFSts") = 1 Or _
               DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("CASts") = 1 Then          '�ק�u�{���A

                DStatus.Text = "�ק�u�{�i�椤"
                DStatus.Visible = True
                Panel1.Visible = True
            Else
                DStatus.Visible = False
                Panel1.Visible = False
            End If
            If DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("OPContact") = 1 Then       '�u�{�s����
                LOPContact.NavigateUrl = "ManufInList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno & "&pFun=OP"
                LOPContact.Visible = True
            Else
                LOPContact.Visible = False
            End If
            If DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Contact") = 1 Then         '�~�ȳs����
                LContact.NavigateUrl = "ContactList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno & "&pFun=Import"
                LContact.Visible = True
            Else
                LContact.Visible = False

            End If
            If DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Surface") = 1 Then         '���B�z
                LSurface.NavigateUrl = "ManufInList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno & "&pFun=SF"
                LSurface.Visible = True
            Else
                LSurface.Visible = False
            End If
            If DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("ColorAppend") = 1 Then     '�~�`��f
                LColorAppend.NavigateUrl = "ManufInList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno & "&pFun=CA"
                LColorAppend.Visible = True
            Else
                LColorAppend.Visible = False
            End If


            If DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SliderDetail") = 1 Then       '���Y�ӥ�
                LSliderDetail.NavigateUrl = "ManufOutList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno & "&pCode=" & DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("NewSliderGRCode") & "&pFun=SD"
                LSliderDetail.Visible = True
            Else
                LSliderDetail.Visible = False
            End If
            '------------------------------------------------------------------------------------------------

            If DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("MapFile") <> "" Then          '�ϼ�
                '�ʦs���٭�  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("MapFile"))

                LMapFile.ImageUrl = Path & DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("MapFile")
            Else
                LMapFile.Visible = False
            End If

            DNo.Text = DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Date")                   '���
            SetFieldData("Division", DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Division"))  '���
            SetFieldData("Person", DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Person"))      '���
            DSliderCode.Text = DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SliderCode")       'Slider Code
            DSliderGRCode.Text = DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SliderGRCode")   'Slider Group Code
            DSpec.Text = DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Spec")                   '�W��

            If DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("RefFile") <> "" Then          '��L���� 
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("RefFile"))
                LRefFile.NavigateUrl = Path & DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("RefFile")
            Else
                LRefFile.Visible = False
            End If

            SetFieldData("SliderType", DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SliderType"))    '���Y����
            SetFieldData("ManufPlace", DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("ManufPlace"))      '�Ͳ��a
            SetFieldData("Buyer", DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Buyer"))      '�Ͳ��a
            DSellVendor.Text = DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SellVendor")   '�e�U�t��
            DDevReason.Text = DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Reason")     '�}�o�z��

            If DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("Price") = 1 Then               '���
                Dim DBDataSet2 As New DataSet
                Dim DBTable1 As DataTable
                SQL = "Select * From F_ManufCOPrice "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet2, "ManufCOPrice")
                DBTable1 = DBDataSet2.Tables("ManufCOPrice")
                For i = 0 To DBTable1.Rows.Count - 1
                    Select Case DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Seqno")
                        Case 1
                            DPriceA1.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceA2.SelectedValue = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceA3.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 2
                            DPriceB1.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceB2.SelectedValue = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceB3.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 3
                            DPriceC1.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceC2.SelectedValue = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceC3.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 4
                            DPriceD1.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceD2.SelectedValue = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceD3.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 5
                            DPriceE1.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceE2.SelectedValue = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceE3.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 6
                            DPriceF1.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceF2.SelectedValue = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceF3.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 7
                            DPriceG1.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceG2.SelectedValue = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceG3.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 8
                            DPriceH1.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceH2.SelectedValue = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceH3.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 9
                            DPriceI1.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceI2.SelectedValue = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceI3.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case Else
                            DPriceJ1.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceJ2.SelectedValue = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceJ3.Text = DBDataSet2.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                    End Select
                Next
            End If

            DSliderPrice.Text = DBDataSet1.Tables("F_ImportSheet").Rows(0).Item("SliderPrice") '���Y�ʤJ��
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
        '���Y����(���s.�~�`...)
        If pFieldName = "SliderType" Then
            DSliderType.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DSliderType.Items.Add(ListItem1)
        End If
        '�Ͳ��a
        If pFieldName = "ManufPlace" Then
            DManufPlace.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DManufPlace.Items.Add(ListItem1)
        End If
        'Buyer
        If pFieldName = "Buyer" Then
            DBuyer.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DBuyer.Items.Add(ListItem1)
        End If
    End Sub

End Class
