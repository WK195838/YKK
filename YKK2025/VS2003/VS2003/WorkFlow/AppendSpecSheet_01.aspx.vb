Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class AppendSpecSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DOPReadyDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOPReady As System.Web.UI.WebControls.TextBox
    Protected WithEvents BFlow As System.Web.UI.WebControls.ImageButton
    Protected WithEvents BPrint As System.Web.UI.WebControls.ImageButton
    Protected WithEvents LOFormNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DOFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents BIn As System.Web.UI.WebControls.Button
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LBefOP As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DAEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents DDelivery As System.Web.UI.WebControls.Image
    Protected WithEvents DDescSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DContent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTarget As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BDate As System.Web.UI.WebControls.Button
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents BOut As System.Web.UI.WebControls.Button
    Protected WithEvents DPriceJ3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceJ2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceJ1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceI3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceI2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceI1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceH3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceH2 As System.Web.UI.WebControls.DropDownList
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
    Protected WithEvents DPriceA3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceA2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceA1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDescription As System.Web.UI.WebControls.TextBox
    Protected WithEvents LQAFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DQAFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DContactFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents LContactFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DRefFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents LRefFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DAppendSpecSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents BImport As System.Web.UI.WebControls.Button
    Protected WithEvents DAssembler As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LAssemblerFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DAssemblerFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DCPSC As System.Web.UI.WebControls.DropDownList

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
    Dim wApplyCalendar As String = ""       '�ӽЪ�
    Dim wDecideCalendar As String = ""      '�֩w��
    Dim wNextGateCalendar As String = ""    '�U�@�֩w��
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
            If wFormSno > 0 And wStep > 2 Then    '�P�_�O�_[ñ��]
                ShowFormData()      '��ܪ����
                UpdateTranFile()    '��s������
            End If
            SetPopupFunction()      '�]�w�u�X�����ƥ�
        Else
            ShowSheetField("Post")  '��������ܤ�����J�ˬd
            ShowMessage()           '�W�Ǹ���ˬd����ܰT��
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
        wStep = Request.QueryString("pStep")        '�u�{�N�X
        wSeqNo = Request.QueryString("pSeqNo")      '�Ǹ�
        wApplyID = Request.QueryString("pApplyID")  '�ӽЪ�ID
        wAgentID = Request.QueryString("pAgentID")  '�Q�N�z�HID
        'Add-Start by Joy 2009/11/20(2010��ƾ����)
        '�ӽЪ�-�s�զ�ƾ�(TP1->�x�_-����, TP2->�x�_-�ا�, CL1->���c-����, CL2->���c-�s�y)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        wDecideCalendar = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))
        'Add-End

        Response.Cookies("DevNo").Value = ""        '�}�oNo, DevNoPicker�ϥ�

        Response.Cookies("PGM").Value = "AppendSpecSheet_01.aspx"
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
        'Check���ճ��i
        If DQAFile.Visible Then
            If DQAFile.PostedFile.FileName <> "" Then
                Message = "���ճ��i"
            End If
        End If
        'Check������
        If DContactFile.Visible Then
            If DContactFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "������"
                Else
                    Message = Message & ", " & "������"
                End If
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

        'Check�եߧP�w��
        If DAssemblerFile.Visible Then
            If DAssemblerFile.PostedFile.FileName <> "" Then
                Message = "�եߧP�w��"
            End If
        End If


        If Message <> "" Then
            Message = "�U�C�ҳ]�w�����[�ɮױN�|�� (" & Message & ") " & ",�Э��s�]�w!"
            Response.Write(YKK.ShowMessage(Message))
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ��s������
    '**
    '*****************************************************************
    Sub UpdateTranFile()
        'Modify-Start by Joy 2009/11/20(2010��ƾ����)
        'oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDepo, Request.QueryString("pUserID"))
        '
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, Request.QueryString("pUserID"))
        '��渹�X,���y����,�u�{���d���X,�Ǹ�,��ƾ�,ñ�֪�
        'Modify-End
    End Sub

    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("AppendSpecFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_AppendSpecSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_AppendSpecSheet")
        If DBDataSet1.Tables("F_AppendSpecSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Date")                   '���
            SetFieldData("Division", DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Division"))  '����
            SetFieldData("Person", DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Person"))      '���
            SetFieldData("Assembler", DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Assembler"))      '�եߧP�w
            SetFieldData("CPSC", DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("CPSC"))          'CPSC
            DBuyer.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Buyer")                 'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("SellVendor")       '�e�U�t��
            DSliderCode.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("SliderCode")       'Slider Code


            DMapNo.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("MapNo")                 '�ϸ�

            DOFormNo.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("OFormNo")             '����
            DOFormSno.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("OFormSno")           '��渹
            If DOFormNo.Text <> "" And DOFormSno.Text <> "" Then
                If DOFormNo.Text = "000011" Then
                    LOFormNo.NavigateUrl = "ImportSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                Else
                    If DOFormNo.Text = "000003" Then
                        LOFormNo.NavigateUrl = "ManufInSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                    Else
                        LOFormNo.NavigateUrl = "ManufOutSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                    End If
                End If
                DOFormNo.Visible = False
                DOFormSno.Visible = False
            Else
                LOFormNo.Visible = False
            End If

            DTarget.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Target")           '�ت�
            DContent.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Content")         '���e
            DDescription.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Description") 'EA����

            If DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Price") = 1 Then               '���
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

            If DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("QAFile") <> "" Then          '���ճ��i
                LQAFile.NavigateUrl = Path & DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("QAFile")
            Else
                LQAFile.Visible = False
            End If

            If DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("AssemblerFile") <> "" Then          '�եߧP�w
                LAssemblerFile.NavigateUrl = Path & DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("AssemblerFile")
            Else
                LAssemblerFile.Visible = False
            End If

            If DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("ContactFile") <> "" Then     '������ 
                LContactFile.NavigateUrl = Path & DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("ContactFile")
            Else
                LContactFile.Visible = False
            End If
            If DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("RefFile") <> "" Then          '��L���� 
                LRefFile.NavigateUrl = Path & DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("RefFile")
            Else
                LRefFile.Visible = False
            End If
            '*******************************************************************
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
            SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                DDecideDesc.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("DecideDesc")       '����
                DBStartTime.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BStartTime")       '�w�w�}�l
                DBEndTime.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime")           '�w�w����
                DAStartTime.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("AStartTime")       '��ڶ}�l
                If Not IsDBNull(DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("AEndTime")) Then
                    DAEndTime.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("AEndTime")           '��ڧ���
                Else
                    DAEndTime.Text = ""
                End If
                LBefOP.NavigateUrl = "BefOPList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno & "&pStep=" & wStep

                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") < NowDateTime Then
                    If DDelay.Visible = True Then
                        SetFieldData("ReasonCode", DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonCode"))    '����z�ѥN�X
                        If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonCode") = "" Then
                            SetFieldData("Reason", DReasonCode.SelectedValue)    '����z��
                            DReasonDesc.Text = ""   '�����L����
                        Else
                            DReason.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("Reason")  '����z��
                            DReasonDesc.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonDesc")     '�����L����
                        End If

                    End If
                End If
                DFormSno.Text = "�渹�G" & DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("FormSno")       '�渹
            End If
        End If
        'DB�s������
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ����z�ѥN�X�I���ƥ�
    '**
    '*****************************************************************
    Private Sub DReasonCode_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DReasonCode.SelectedIndexChanged
        SetFieldData("Reason", DReasonCode.SelectedValue)    '����z��
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetFunction)
    '**     ���\����s���
    '**
    '*****************************************************************
    Sub ShowSheetFunction()
        Dim wDelay As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

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
                BSAVE.Visible = True
                BSAVE.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("SaveDesc")
                '-- ������s���G�� JOY 16/08/17
                '��
                'BSAVE.Attributes("onclick") = "Button('SAVE', '" + BSAVE.Value + "');"
                '�s
                BSAVE.Attributes("onclick") = "this.disabled = true;" & "Button('SAVE', '" + BSAVE.Value + "');"
                '--
            Else
                BSAVE.Visible = False
            End If
            'NG-1���s
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGDesc1")
                '-- ������s���G�� JOY 16/08/17
                '��
                'BNG1.Attributes("onclick") = "Button('NG1', '" + BNG1.Value + "');"
                '�s
                BNG1.Attributes("onclick") = "this.disabled = true;" & "Button('NG1', '" + BNG1.Value + "');"
                '--
                wNGSts1 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts1") + 1
            Else
                BNG1.Visible = False
            End If
            'NG-2���s
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGDesc2")
                '-- ������s���G�� JOY 16/08/17
                '��
                'BNG2.Attributes("onclick") = "Button('NG2', '" + BNG2.Value + "');"
                '�s
                BNG2.Attributes("onclick") = "this.disabled = true;" & "Button('NG2', '" + BNG2.Value + "');"
                '--
                wNGSts2 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts2") + 1
            Else
                BNG2.Visible = False
            End If
            'OK���s
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("OKDesc")
                '-- ������s���G�� JOY 16/08/17
                '��
                'BOK.Attributes("onclick") = "Button('OK', '" + BOK.Value + "');"
                '�s
                BOK.Attributes("onclick") = "this.disabled = true;" & "Button('OK', '" + BOK.Value + "');"
                '--
                wOKSts = DBDataSet1.Tables("M_Flow").Rows(0).Item("OKSts") + 1
            Else
                BOK.Visible = False
            End If
            '��Ǻ޲z
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("Delay") = 1 Then
                wDelay = 1
            End If
        End If

        If wFormSno > 0 And wStep > 2 Then    '�P�_�O�_[ñ��]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                'Sheet���
                DAppendSpecSheet.Visible = True '���Sheet-1
                DDescSheet.Visible = True       '����Sheet
                DDelivery.Visible = True        '���Sheet

                '��Ǻ޲z
                If wDelay = 1 Then
                    If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") < NowDateTime Then
                        DDelay.Visible = True   '����Sheet
                    Else
                        DDelay.Visible = False  '����Sheet
                    End If
                End If
                If DDelay.Visible = True Then
                    DReasonCode.Visible = True     '����z�ѥN�X
                    DReasonCode.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonCodeRqd", "DReasonCode", "���`�G�ݿ�J����z��")
                    DReason.Visible = True         '����z��
                    DReason.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonRqd", "DReason", "���`�G�ݿ�J����z��")
                    DReasonDesc.Visible = True     '�����L����
                    Top = 1000
                Else
                    DReasonCode.Visible = False     '����z�ѥN�X
                    DReason.Visible = False         '����z��
                    DReasonDesc.Visible = False     '�����L����
                    Top = 888
                End If

                '���
                DBStartTime.Visible = True      '�w�w�}�l
                DBEndTime.Visible = True        '�w�w����
                DAStartTime.Visible = True      '��ڶ}�l
                DAEndTime.Visible = True        '��ڧ���
                DOPReady.Visible = True     '�u�{�i���\Ū
                DOPReadyDesc.Visible = True
                '�ݾ\Ū
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("FlowType") = 0 Then
                    DOPReady.BackColor = Color.LightGray
                Else
                    '20190430 JESSICA ����
                    DOPReady.BackColor = Color.LightGray

                    'DOPReady.BackColor = Color.GreenYellow
                    'ShowRequiredFieldValidator("DOPReadyRqd", "DOPReady", "���`�G�ݾ\Ū�u�{�i��")
                End If

                '������
                DDecideDesc.Visible = True      '����
                '�����ݿ�J
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("FlowType") = 0 Then
                    DDecideDesc.BackColor = Color.LightGray
                Else
                    DDecideDesc.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDecideDescRqd", "DDecideDesc", "���`�G�ݿ�J����")
                End If

                '�s�����---�ݦA�ק�
                LOFormNo.Visible = True        '��e�U
                LQAFile.Visible = True          '���ճ��i
                LAssemblerFile.Visible = True          '�եߧP�w
                LContactFile.Visible = True     '������
                LRefFile.Visible = True         '��L����
                LBefOP.Visible = True           '�u�{�i��
                '���s��m
                BNG1.Style.Add("Top", Top)     'NG1���s
                BNG2.Style.Add("Top", Top)     'NG2���s
                BSAVE.Style.Add("Top", Top)    '�x�s���s
                BOK.Style.Add("Top", Top)      'OK���s
                DFormSno.Style.Add("Top", Top) '�渹
            End If
        Else
            Top = 704
            'Sheet����
            DDescSheet.Visible = False  '����Sheet
            DDelivery.Visible = False   '���Sheet
            DDelay.Visible = False      '����Sheet
            '�������
            DDecideDesc.Visible = False     '����
            DBStartTime.Visible = False     '�w�w�}�l
            DBEndTime.Visible = False       '�w�w����
            DAStartTime.Visible = False     '��ڶ}�l
            DAEndTime.Visible = False       '��ڧ���
            DReasonCode.Visible = False     '����z�ѥN�X
            DReason.Visible = False         '����z��
            DReasonDesc.Visible = False     '�����L����

            DOPReady.Visible = False    '�u�{�i���\Ū
            DOPReadyDesc.Visible = False

            '�s������
            LOFormNo.Visible = False         '��e�U
            LQAFile.Visible = False          '���ճ��i
            LAssemblerFile.Visible = False          '�եߧP�w
            LContactFile.Visible = False     '������
            LRefFile.Visible = False         '��L����
            LBefOP.Visible = False           '�u�{�i��
            '���s��m
            BNG1.Style.Add("Top", Top)     'NG1���s
            BNG2.Style.Add("Top", Top)     'NG2���s
            BSAVE.Style.Add("Top", Top)    '�x�s���s
            BOK.Style.Add("Top", Top)      'OK���s
            DFormSno.Style.Add("Top", Top) '�渹
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
        BIn.Attributes("onclick") = "DevNoPicker('In','000012');"       '���s�e�U��-No
        BOut.Attributes("onclick") = "DevNoPicker('Out','000012');"     '�~�`�e�U��-No
        BImport.Attributes("onclick") = "DevNoPicker('Import','000012');"     '�~�`�e�U��-No
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     ���s��RequestedField��Top��m
    '**
    '*****************************************************************
    Sub TopPosition()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()
        '���s��RequestedField��Top��m
        If wFormSno > 0 And wStep > 2 Then    '�P�_�O�_[ñ��]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") >= NowDateTime Then
                    Top = 888
                Else
                    If DDelay.Visible = True Then
                        Top = 1000
                    Else
                        Top = 888
                    End If
                End If
            End If
        Else
            Top = 704
        End If
        '----

        'DB�s������
        OleDbConnection1.Close()
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
                BIn.Visible = False
                BOut.Visible = False
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

        '���
        Select Case FindFieldInf("Assembler")
            Case 0  '���
                DAssembler.BackColor = Color.LightGray
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
        If pPost = "New" Then SetFieldData("Assembler", "ZZZZZZ")

        'CPSC
        Select Case FindFieldInf("CPSC")
            Case 0  '���
                DCPSC.BackColor = Color.LightGray
                DCPSC.Visible = True
            Case 1  '�ק�+�ˬd
                DCPSC.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCPSCRqd", "DCPSC", "���`�G�ݿ�JCPSC")
                DCPSC.Visible = True
            Case 2  '�ק�
                DCPSC.BackColor = Color.Yellow
                DCPSC.Visible = True
            Case Else   '����
                DCPSC.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("CPSC", "ZZZZZZ")



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
        If pPost = "New" Then DBuyer.Text = ""

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

        'Slider Code
        Select Case FindFieldInf("SliderCode")
            Case 0  '���
                DSliderCode.BackColor = Color.LightGray
                DSliderCode.ReadOnly = True
                DSliderCode.Visible = True
            Case 1  '�ק�+�ˬd
                DSliderCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderCodeRqd", "DSliderCode", "���`�G�ݿ�J Slider Code")
                DSliderCode.Visible = True
            Case 2  '�ק�
                DSliderCode.BackColor = Color.Yellow
                DSliderCode.Visible = True
            Case Else   '����
                DSliderCode.Visible = False
        End Select
        If pPost = "New" Then DSliderCode.Text = ""

        '�ϸ�
        Select Case FindFieldInf("MapNo")
            Case 0  '���
                DMapNo.BackColor = Color.LightGray
                DMapNo.ReadOnly = True
                DMapNo.Visible = True
            Case 1  '�ק�+�ˬd
                DMapNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMapNoRqd", "DMapNo", "���`�G�ݿ�J�ϸ�")
                DMapNo.Visible = True
            Case 2  '�ק�
                DMapNo.BackColor = Color.Yellow
                DMapNo.Visible = True
            Case Else   '����
                DMapNo.Visible = False
        End Select
        If pPost = "New" Then DMapNo.Text = ""

        'OFormNo
        Select Case FindFieldInf("OFormNo")
            Case 0  '���
                DOFormNo.BackColor = Color.LightGray
                DOFormNo.ReadOnly = True
                DOFormNo.Visible = True
            Case 1  '�ק�+�ˬd
                DOFormNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOFormNoRqd", "DOFormNo", "���`�G�ݿ�J��渹�X")
                DOFormNo.Visible = True
            Case 2  '�ק�
                DOFormNo.BackColor = Color.Yellow
                DOFormNo.Visible = True
            Case Else   '����
                DOFormNo.Visible = False
        End Select
        If pPost = "New" Then DOFormNo.Text = ""

        'OFormSno
        Select Case FindFieldInf("OFormSno")
            Case 0  '���
                DOFormSno.BackColor = Color.LightGray
                DOFormSno.ReadOnly = True
                DOFormSno.Visible = True
            Case 1  '�ק�+�ˬd
                DOFormSno.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOFormSnoRqd", "DOFormSno", "���`�G�ݿ�J�渹")
                DOFormSno.Visible = True
            Case 2  '�ק�
                DOFormSno.BackColor = Color.Yellow
                DOFormSno.Visible = True
            Case Else   '����
                DOFormSno.Visible = False
        End Select
        If pPost = "New" Then DOFormSno.Text = ""

        '�ت�
        Select Case FindFieldInf("Target")
            Case 0  '���
                DTarget.BackColor = Color.LightGray
                DTarget.ReadOnly = True
                DTarget.Visible = True
            Case 1  '�ק�+�ˬd
                DTarget.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTargetRqd", "DTarget", "���`�G�ݿ�J�ت�")
                DTarget.Visible = True
            Case 2  '�ק�
                DTarget.BackColor = Color.Yellow
                DTarget.Visible = True
            Case Else   '����
                DTarget.Visible = False
        End Select
        If pPost = "New" Then DTarget.Text = ""

        '���e
        Select Case FindFieldInf("Content")
            Case 0  '���
                DContent.BackColor = Color.LightGray
                DContent.ReadOnly = True
                DContent.Visible = True
            Case 1  '�ק�+�ˬd
                DContent.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DContentRqd", "DContent", "���`�G�ݿ�J���e")
                DContent.Visible = True
            Case 2  '�ק�
                DContent.BackColor = Color.Yellow
                DContent.Visible = True
            Case Else   '����
                DContent.Visible = False
        End Select
        If pPost = "New" Then DContent.Text = ""

        'EA����
        Select Case FindFieldInf("Description")
            Case 0  '���
                DDescription.BackColor = Color.LightGray
                DDescription.ReadOnly = True
                DDescription.Visible = True
            Case 1  '�ק�+�ˬd
                DDescription.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDescriptionRqd", "DDescription", "���`�G�ݿ�J����")
                DDescription.Visible = True
            Case 2  '�ק�
                DDescription.BackColor = Color.Yellow
                DDescription.Visible = True
            Case Else   '����
                DDescription.Visible = False
        End Select
        If pPost = "New" Then DDescription.Text = ""

        '���ճ��i
        Select Case FindFieldInf("QAFile")
            Case 0  '���
                DQAFile.Visible = False
                DQAFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DQAFileRqd", "DQAFile", "���`�G�ݿ�J���ճ��i")
                DQAFile.Visible = True
                DQAFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DQAFile.Visible = True
                DQAFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DQAFile.Visible = False
        End Select

        '������
        Select Case FindFieldInf("ContactFile")
            Case 0  '���
                DContactFile.Visible = False
                DContactFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DContactFileRqd", "DContactFile", "���`�G�ݿ�J������")
                DContactFile.Visible = True
                DContactFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DContactFile.Visible = True
                DContactFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DContactFile.Visible = False
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

        '�եߧP�w
        Select Case FindFieldInf("AssemblerFile")
            Case 0  '���
                DAssemblerFile.Visible = False
                DAssemblerFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DAssemblerFileRqd", "DAssemblerFile", "���`�G�ݿ�J�եߧP�w")
                DAssemblerFile.Visible = True
                DAssemblerFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DAssemblerFile.Visible = True
                DAssemblerFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DAssemblerFile.Visible = False
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

        '�եߧP�w
        If pFieldName = "Assembler" Then
            DAssembler.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAssembler.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='200' and Dkey = 'Assembler'  Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAssembler.Items.Add(ListItem1)
                Next
            End If
        End If


        'CPSC
        If pFieldName = "CPSC" Then
            DCPSC.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCPSC.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='100' and Dkey = 'CPSC'  Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCPSC.Items.Add(ListItem1)
                Next
            End If
        End If


        '����z�ѥN�X
        If pFieldName = "ReasonCode" Then
            DReasonCode.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DReasonCode.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='014' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("DKey")
                    ListItem1.Value = DBTable1.Rows(i).Item("DKey")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DReasonCode.Items.Add(ListItem1)
                Next
            End If
        End If

        '����z��
        If pFieldName = "Reason" Then
            SQL = "Select * From M_Referp Where Cat='014' and DKey= '" & DReasonCode.SelectedValue & "' Order by Data "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                DReason.Text = DBTable1.Rows(i).Item("Data")
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
    Private Sub BSAVE_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSAVE.ServerClick
        If Request.Cookies("RunBSAVE").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '����Button�B�@
                Dim ErrCode As Integer = 0   '�ŧi�@��Error�ΨӧP�O�O�_��ƥ��`�A�w�]��False
                Dim Message As String = ""

                'Check���ճ��iSize�ή榡
                If ErrCode = 0 Then
                    If DQAFile.Visible Then
                        If DQAFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                            ErrCode = UPFileIsNormal(DQAFile)
                        End If
                    End If
                End If

                'Check�եߧP�wSize�ή榡
                If ErrCode = 0 Then
                    If DAssemblerFile.Visible Then
                        If DAssemblerFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                            ErrCode = UPFileIsNormal(DAssemblerFile)
                        End If
                    End If
                End If

                'Check������Size�ή榡
                If ErrCode = 0 Then
                    If DContactFile.Visible Then
                        If DContactFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                            ErrCode = UPFileIsNormal(DContactFile)
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

                'Check�z��
                If ErrCode = 0 Then
                    If DReasonCode.Visible = True Then
                        If DReasonCode.SelectedValue = "99" Then
                            If DReasonDesc.Text = "" Then ErrCode = 9210
                        End If
                    End If
                End If
                '�x�s���
                If ErrCode = 0 Then
                    ModifyData("SAVE", "0")           '��s����� Sts=0(����)
                    ModifyTranData("SAVE", "0")       '��s������
                Else
                    If ErrCode = 9010 Then Message = "�W���ɮ׮榡���~,�нT�{!"
                    If ErrCode = 9020 Then Message = "�W���ɮ�Size�W�L1024KB,�нT�{!"
                    If ErrCode = 9030 Then Message = "�W���ɮרS���w,�нT�{!"
                    If ErrCode = 9210 Then Message = "����z�Ѭ���L�ɻݶ�g����,�нT�{!"
                    Response.Write(YKK.ShowMessage(Message))
                End If      '�W���ɮ�ErrCode=0

                If ErrCode = 0 Then
                    Dim URL As String = "http://10.245.1.6/WorkFlowSub/MessagePage.aspx?pMSGID=S&pFormNo=" & wFormNo & "&pStep=" & wStep & _
                                                        "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                    Response.Redirect(URL)
                End If
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BOK_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BOK.ServerClick
        If Request.Cookies("RunBOK").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '����Button�B�@
                FlowControl("OK", 0, "1")
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-1���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BNG1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG1.ServerClick
        If Request.Cookies("RunBNG1").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '����Button�B�@
                FlowControl("NG1", 1, "2")
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-2���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BNG2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG2.ServerClick
        If Request.Cookies("RunBNG2").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '����Button�B�@
                FlowControl("NG2", 2, "3")
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(FlowControl)
    '**     �y�{����
    '**        pFun=OK, NG1, NG2, SAVE  
    '**        pAction=0:OK, 1:NG1, 2:NG2, 3:Save   �U�@���d�ɨϥ� 
    '**        pSts=0:���B�z, 1:OK, 2:NG1, 3:NG2, 4:�w�\Ū, 5:�Q���  ��sT_Waithandle���A
    '**     
    '*****************************************************************
    Sub FlowControl(ByVal pFun As String, ByVal pAction As Integer, ByVal pSts As String)
        Dim ErrCode As Integer = 0   '�ŧi�@��Error�ΨӧP�O�O�_��ƥ��`�A�w�]��False
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim Message As String = ""
        Dim wQCLT As Integer = 0 'QC-L/T

        GetDataStatus()  '���o�����d�U���s��Data Status

        'Check���ճ��iSize�ή榡
        If ErrCode = 0 Then
            If DQAFile.Visible Then
                If DQAFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DQAFile)
                End If
            End If
        End If

        'Check�եߧP�wSize�ή榡
        If ErrCode = 0 Then
            If DAssemblerFile.Visible Then
                If DAssemblerFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DAssemblerFile)
                End If
            End If

        End If
        'Check������Size�ή榡
        If ErrCode = 0 Then
            If DContactFile.Visible Then
                If DContactFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DContactFile)
                End If
            End If
        End If
        'Check��L����-2Size�ή榡
        If ErrCode = 0 Then
            If DRefFile.Visible Then
                If DRefFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DRefFile)
                End If
            End If
        End If

        'Check����z��
        If ErrCode = 0 Then
            If DReasonCode.Visible = True Then
                If DReasonCode.SelectedValue = "99" Then
                    If DReasonDesc.Text = "" Then ErrCode = 9040
                End If
            End If
        End If

        '--�ˬd�e�U��No---------
        If ErrCode = 0 Then
            If DNo.Text <> "" Then
                ErrCode = oCommon.CommissionNo("000012", wFormSno, wStep, DNo.Text) '��渹�X, ���y����, �u�{, �e�U��No
                If ErrCode <> 0 Then
                    ErrCode = 9060
                End If
            End If
        End If

        If ErrCode = 0 Then
            Dim Run As Boolean = True           '�O�_����
            Dim RepeatRun As Boolean = False    '�O�_���а���
            Dim MultiJob As Integer = 0         '�h�H�֩w
            Dim wLevel As String = ""           '������

            While Run = True
                Run = False     '����Flag=������
                '--���o�U�@���ѼƳ]�w---------
                Dim pNextGate(10) As String
                Dim pAgentGate(10) As String
                Dim pNextStep As Integer = 0
                Dim pFlowType As Integer = 0    '0=�q��
                Dim pCount As Integer
                '--���oLeadTime�ѼƳ]�w---------
                Dim pCTime, pStartTime, pEndTime As DateTime
                '--���o�u�{�t���ѼƳ]�w---------
                Dim pLastTime As DateTime
                Dim pCount1 As Integer
                '--��L�ѼƳ]�w---------
                Dim RtnCode, i As Integer
                Dim NewFormSno As Integer = wFormSno    '���y����
                Dim pRunNextStep As Integer = 0         '�O�_����p��U�@��(�|ñ)
                Dim SQL As String

                '���o���y�����Χ�s������
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then    '�P�_�O�_�_��
                        '���o���y����
                        RtnCode = oCommon.GetFormSeqNo(wFormNo, NewFormSno) '��渹�X, ���y����
                        If RtnCode <> 0 Then
                            ErrCode = 9110
                        Else
                            '�ӽЬy�{��ƫظm
                            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
                            'oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDepo, Request.QueryString("pUserID"), wApplyID)
                            '
                            oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                            '��渹�X,���y����,�u�{���d���X,�Ǹ�,��ƾ�,ñ�֪�, �ӽЪ�
                            'Modify-End
                        End If
                        pRunNextStep = 1
                    Else
                        If RepeatRun = False Then   '���O�q�������а���
                            '��s������
                            ModifyTranData(pFun, pSts)

                            '�y�{��Ƶ���
                            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
                            'RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDepo, Request.QueryString("pUserID"), pRunNextStep)
                            '
                            RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDecideCalendar, Request.QueryString("pUserID"), pRunNextStep)
                            '��渹�X,���y����,�u�{���d���X,��ƾ�,ñ�֪�, �y�{�����_(�|ñ)
                            'Modify-End

                            If RtnCode <> 0 Then ErrCode = 9120
                        Else
                            pRunNextStep = 1    '�O�q�������а���
                        End If
                    End If
                End If

                '���o�U�@��
                If ErrCode = 0 And pRunNextStep = 1 Then
                    Dim wAllocateID As String = ""

                    RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, wAgentID, wAllocateID, MultiJob, _
                                                  pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)
                    '��渹�X,�u�{���d���X,ñ�֪�,�ӽЪ�,�Q�N�z��,�Q���w��,�h�H�֩w�u�{No,
                    '�U�@�u�{��, ���X, ����, �Q�N�z��, �H��, �B�z��k, �ʧ@(0:OK,1:NG,2:Save) 

                    If RtnCode <> 0 Then ErrCode = 9130
                    If pCount = 0 And pNextStep <> 999 Then ErrCode = 9131
                    If ErrCode = 0 Then pAction = 0
                End If

                '�ظm�y�{���
                If ErrCode = 0 And pRunNextStep = 1 Then
                    If pNextStep <> 999 Then
                        wNextGate = ""
                        For i = 1 To pCount
                            '���o�U�@���H��(�T���ɨϥ�)
                            If wNextGate = "" Then
                                wNextGate = pNextGate(i)
                            Else
                                wNextGate = "," & pNextGate(i)
                            End If

                            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
                            'oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)
                            'oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wDepo)
                            'RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wDepo, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '
                            '���o�֩w��-�s�զ�O��
                            wNextGateCalendar = oCommon.GetCalendarGroup(pNextGate(i))

                            '���o�u�{�t���̫���
                            oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)
                            '�֩w��, ��渹�X, �u�{���X, ���O(0:�q��,1:�֩w), �}�l���, �̫���, ���

                            '���o�w�w�}�l,������{�p��
                            oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wNextGateCalendar)
                            '��渹�X,�u�{���X,������,QC-L/T,�{�b�ɶ�, �w�w�}�l���, �w�w�������, ��ƾ�

                            '�ظm�y�{���
                            RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '��渹�X,���y����,�u�{���d���X,�Ǹ�,�ӽЪ�ID,��ƾ�,�ظm��, ñ�֪�, �Q�N�z��, �ӽЪ�, �w�w�}�l��, �w�w������, ���n��
                            'Modify-End

                            If RtnCode <> 0 Then
                                ErrCode = 9150
                                Exit For
                            End If
                        Next i
                    Else
                        'Modify-Start by 2009/11/20(2010��ƾ����)
                        'RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDepo, Request.QueryString("pUserID"), wApplyID)
                        '
                        RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                        '��渹�X,���y����,�u�{���d���X,�Ǹ�,��ƾ�,ñ�֪�, �ӽЪ�
                        'Modify-End

                        If RtnCode <> 0 Then ErrCode = 9160
                    End If
                End If
                '��u�{��{�վ�
                If ErrCode = 0 Then
                    If RepeatRun = False Then
                        'Modify-Start by 2009/11/20(2010��ƾ����)
                        'RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDepo)
                        '
                        RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDecideCalendar)
                        'ñ�֪�,��渹�X,���y����,�u�{���d���X,�Ǹ�,�{�b���,������,��ƾ�
                        'Modify-End
                    End If
                End If
                '�x�s�����
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then      '�P�_�O�_�_��
                        If NewFormSno <> 0 Then
                            AppendData(pFun, NewFormSno)    '�s�W�����
                            AddCommissionNo(wFormNo, NewFormSno)
                            ModifyManufData(pFun, 1, 1)   '��s���檬�A
                        End If  'pSeqno <> 0
                    Else    '�P�_�O�_�_��
                        If pNextStep = 999 Then     '�u�{������?
                            If pFun = "OK" Then ModifyData(pFun, CStr(wOKSts)) '��s�����
                            If pFun = "NG1" Then ModifyData(pFun, CStr(wNGSts1))
                            If pFun = "NG2" Then ModifyData(pFun, CStr(wNGSts2))
                            ModifyManufData(pFun, 0, 1)   '��s���檬�A
                        Else
                            ModifyData(pFun, "0")         '��s����� Sts=0(����)
                        End If
                        AddCommissionNo(wFormNo, wFormSno)
                    End If
                    '�ǰe�l��
                    If pNextStep <> 999 Then
                        For i = 1 To pCount
                            oCommon.Send(Request.QueryString("pUserID"), pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "FLOW")
                            '�e���, �����, �ӽЪ�, ��渹�X, ���y����, �u�{���d���X,�T�����O
                        Next i
                    Else
                        oCommon.Send(Request.QueryString("pUserID"), wApplyID, wApplyID, wFormNo, NewFormSno, pNextStep, "END")
                        '�e���, �����, �ӽЪ�, ��渹�X, ���y����, �u�{���d���X,�T�����O
                    End If

                    If pFlowType <> 3 Then
                        MultiJob = 0
                    Else
                        MultiJob = 1
                    End If

                    If (pRunNextStep = 1) And (pFlowType = 0 Or pFlowType = 3) Then
                        Run = True
                        RepeatRun = True
                        pAction = 0
                    End If

                    wStep = pNextStep     '�U�@�u�{���d���X
                    wFormSno = NewFormSno '�U�@�u�{���y����
                Else
                    EnabledButton()   '�_��Button�B�@
                    If ErrCode = 9110 Then Message = "���o���y�����p�ⲧ�`,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9120 Then Message = "�y�{��Ƨ�s���`,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9130 Then Message = "�U�u�{�p�ⲧ�`,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9131 Then Message = "�L�U�u�{�޲z�H,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9140 Then Message = "�u�{�w�w�}�l�Χ�����p�ⲧ�`,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9150 Then Message = "�U�@�u�{��ƫظm���`,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9160 Then Message = "�u�{������ƫظm���`(999),�нT�{�γs���t�ΤH��!"
                    Response.Write(YKK.ShowMessage(Message))
                End If      '�x�s���ErrCode=0
            End While       '���а���

            If ErrCode = 0 Then
                '--�l��ǰe---------
                oCommon.SendMail()

                Dim URL As String = "http://10.245.1.6/WorkFlowSub/MessagePage.aspx?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                    "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
            End If
        Else
            EnabledButton()   '�_��Button�B�@
            If ErrCode = 9010 Then Message = "�W���ɮ׮榡���~,�нT�{!"
            If ErrCode = 9020 Then Message = "�W���ɮ�Size�W�L1024KB,�нT�{!"
            If ErrCode = 9030 Then Message = "�W���ɮרS���w,�нT�{!"
            If ErrCode = 9040 Then Message = "����z�Ѭ���L�ɻݶ�g����,�нT�{!"
            If ErrCode = 9050 Then Message = "QC L/T�ݬ����ļƦr,�нT�{!"
            If ErrCode = 9060 Then Message = "�e�U��No.����,�нT�{�e�U��No.!"
            Response.Write(YKK.ShowMessage(Message))
        End If      '�W���ɮ�ErrCode=0
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetDataStatus)
    '**     ���o���i�ת��A
    '**
    '*****************************************************************
    Sub GetDataStatus()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Flow")
        If DBDataSet1.Tables("M_Flow").Rows.Count > 0 Then
            'NG-1���s
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun1") = 1 Then
                wNGSts1 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts1") + 1
            End If
            'NG-2���s
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun2") = 1 Then
                wNGSts2 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts2") + 1
            End If
            'OK���s
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("OKFun") = 1 Then
                wOKSts = DBDataSet1.Tables("M_Flow").Rows(0).Item("OKSts") + 1
            End If
        End If
        'DB�s������
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     �s�W�����
    '**
    '*****************************************************************
    Sub AppendData(ByVal pFun As String, ByVal NewFormSno As Integer)
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("AppendSpecFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String
        Dim RtnCode As Integer

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        OleDbConnection1.Open()
        SQl = "Insert into F_AppendSpecSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSNo, "         '1~4
        SQl = SQl + "No, Date, Division, Person,Assembler,CPSC, Buyer, "           '5~9
        SQl = SQl + "SellVendor, SliderCode, MapNo, OFormNo, OFormSno, "   '10~14
        SQl = SQl + "Target, Content, Description, Price, QAFile,AssemblerFile, "        '15~19
        SQl = SQl + "ContactFile, RefFile, "                               '20~21
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "      '22~25
        SQl = SQl + ")  "
        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                                '���A(0:����,1:�w��NG,2:�w��OK)
        SQl = SQl + " '" + NowDateTime + "', "              '���פ�
        SQl = SQl + " '000012', "                           '���N��
        SQl = SQl + " '" + CStr(NewFormSno) + "', "         '���y����

        SQl = SQl + " '" + YKK.ReplaceString(DNo.Text) + "', " 'NO
        SQl = SQl + " '" + DDate.Text + "', "                  '���
        SQl = SQl + " '" + DDivision.SelectedValue + "', "     '����
        SQl = SQl + " '" + DPerson.SelectedValue + "', "       '���
        SQl = SQl + " '" + DAssembler.SelectedValue + "', "       '�եߧP�w
        SQl = SQl + " '" + DCPSC.SelectedValue + "', "       'CPSC
        SQl = SQl + " '" + YKK.ReplaceString(DBuyer.Text) + "', "       'Buyer
        SQl = SQl + " N'" + YKK.ReplaceString(DSellVendor.Text) + "', "  '�e�U�t��
        SQl = SQl + " '" + YKK.ReplaceString(DSliderCode.Text) + "', "  'Slider Code
        SQl = SQl + " '" + DMapNo.Text + "', "              '�ϸ�
        SQl = SQl + " '" + DOFormNo.Text + "', "            '���No
        If DOFormSno.Text = "" Then                         '�渹
            SQl = SQl + " '0', "
        Else
            SQl = SQl + " '" + DOFormSno.Text + "', "
        End If
        SQl = SQl + " N'" + YKK.ReplaceString(DTarget.Text) + "', "   '�ت�
        SQl = SQl + " N'" + YKK.ReplaceString(DContent.Text) + "', "  '���e
        SQl = SQl + " N'" + YKK.ReplaceString(DDescription.Text) + "', "  'EA����

        If DPriceA1.Text <> "" Then                           'Price(��ܥ�)
            SQl = SQl + " '1', "
        Else
            SQl = SQl + " '0', "
        End If

        FileName = ""
        If DQAFile.Visible Then                           '���ճ��i
            If DQAFile.PostedFile.FileName <> "" Then     '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QA" & "-" & UploadDateTime & "-" & Right(DQAFile.PostedFile.FileName, InStr(StrReverse(DQAFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QA" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQAFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DQAFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        FileName = ""
        If DAssemblerFile.Visible Then                           '�եߧP�w
            If DAssemblerFile.PostedFile.FileName <> "" Then     '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QA" & "-" & UploadDateTime & "-" & Right(DAssemblerFile.PostedFile.FileName, InStr(StrReverse(DAssemblerFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "AssemblerFile" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DAssemblerFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DAssemblerFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "


        FileName = ""
        If DContactFile.Visible Then                           '������
            If DContactFile.PostedFile.FileName <> "" Then     '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "Contact" & "-" & UploadDateTime & "-" & Right(DContactFile.PostedFile.FileName, InStr(StrReverse(DContactFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "Contact" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DContactFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DContactFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        FileName = ""
        If DRefFile.Visible Then                           '����2
            If DRefFile.PostedFile.FileName <> "" Then     '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "Ref" & "-" & UploadDateTime & "-" & Right(DRefFile.PostedFile.FileName, InStr(StrReverse(DRefFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "Ref" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DRefFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DRefFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '�@����
        SQl = SQl + " '" + NowDateTime + "', "       '�@���ɶ�
        SQl = SQl + " '" + "" + "', "                       '�ק��
        SQl = SQl + " '" + NowDateTime + "' "       '�ק�ɶ�
        SQl = SQl + " ) "
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDBCommand1.ExecuteNonQuery()
        '
        'Modify-Start by Joy 09-10-03
        'Price Table�B�z
        Dim i As Integer
        For i = 1 To 10
            SQl = "Insert into F_ManufCOPrice "
            SQl = SQl + "( "
            SQl = SQl + "Sts, CompletedTime, FormNo, FormSNo, SeqNo, "
            SQl = SQl + "Spec, Currency, Price, "
            SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "
            SQl = SQl + ")  "
            SQl = SQl + "VALUES( "
            SQl = SQl + " '0', "
            SQl = SQl + " '" + NowDateTime + "', "
            SQl = SQl + " '000012', "
            SQl = SQl + " '" + CStr(NewFormSno) + "', "
            SQl = SQl + " '" + CStr(i) + "', "
            Select Case i
                Case Is = 1
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceA1.Text) + "', "
                    SQl = SQl + " '" + DPriceA2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceA3.Text) + "', "
                Case Is = 2
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceB1.Text) + "', "
                    SQl = SQl + " '" + DPriceB2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceB3.Text) + "', "
                Case Is = 3
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceC1.Text) + "', "
                    SQl = SQl + " '" + DPriceC2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceC3.Text) + "', "
                Case Is = 4
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceD1.Text) + "', "
                    SQl = SQl + " '" + DPriceD2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceD3.Text) + "', "
                Case Is = 5
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceE1.Text) + "', "
                    SQl = SQl + " '" + DPriceE2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceE3.Text) + "', "
                Case Is = 6
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceF1.Text) + "', "
                    SQl = SQl + " '" + DPriceF2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceF3.Text) + "', "
                Case Is = 7
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceG1.Text) + "', "
                    SQl = SQl + " '" + DPriceG2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceG3.Text) + "', "
                Case Is = 8
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceH1.Text) + "', "
                    SQl = SQl + " '" + DPriceH2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceH3.Text) + "', "
                Case Is = 9
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceI1.Text) + "', "
                    SQl = SQl + " '" + DPriceI2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceI3.Text) + "', "
                Case Is = 10
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceJ1.Text) + "', "
                    SQl = SQl + " '" + DPriceJ2.SelectedValue + "', "
                    SQl = SQl + " '" + YKK.ReplaceString(DPriceJ3.Text) + "', "
            End Select
            SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '�@����
            SQl = SQl + " '" + NowDateTime + "', "       '�@���ɶ�
            SQl = SQl + " '" + "" + "', "                       '�ק��
            SQl = SQl + " '" + NowDateTime + "' "       '�ק�ɶ�
            SQl = SQl + " ) "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQl
            OleDBCommand1.ExecuteNonQuery()
        Next
        OleDbConnection1.Close()
        'Modify-end by Joy
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     ��s�����
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("AppendSpecFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String
        Dim DBDataSet1 As New DataSet
        Dim RtnCode As Integer

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        OleDbConnection1.Open()
        SQl = "Update F_AppendSpecSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = '" & YKK.ReplaceString(DNo.Text) & "',"
        SQl = SQl + " Date = '" & DDate.Text & "',"
        SQl = SQl + " Division = '" & DDivision.SelectedValue & "',"
        SQl = SQl + " Person = '" & DPerson.SelectedValue & "',"
        SQl = SQl + " Assembler = '" & DAssembler.SelectedValue & "',"
        SQl = SQl + " CPSC = '" & DCPSC.SelectedValue & "',"
        SQl = SQl + " Buyer = '" & DBuyer.Text & "',"
        SQl = SQl + " SellVendor = N'" & YKK.ReplaceString(DSellVendor.Text) & "',"
        SQl = SQl + " SliderCode = '" & YKK.ReplaceString(DSliderCode.Text) & "',"
        SQl = SQl + " MapNo = '" & DMapNo.Text & "',"
        SQl = SQl + " OFormNo = '" & DOFormNo.Text & "',"
        SQl = SQl + " OFormSno = '" & DOFormSno.Text & "',"
        SQl = SQl + " Target = N'" & YKK.ReplaceString(DTarget.Text) & "',"
        SQl = SQl + " Content = N'" & YKK.ReplaceString(DContent.Text) & "',"
        SQl = SQl + " Description = N'" & YKK.ReplaceString(DDescription.Text) & "',"

        If DPriceA1.Text <> "" Then
            SQl = SQl + " Price = '" & "1" & "',"
        Else
            SQl = SQl + " Price = '" & "0" & "',"
        End If

        If DQAFile.Visible Then
            If DQAFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QA" & "-" & UploadDateTime & "-" & Right(DQAFile.PostedFile.FileName, InStr(StrReverse(DQAFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QA" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQAFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DQAFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QAFile = N'" & FileName & "',"
            End If
        End If

        If DAssemblerFile.Visible Then
            If DAssemblerFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QA" & "-" & UploadDateTime & "-" & Right(DAssemblerFile.PostedFile.FileName, InStr(StrReverse(DAssemblerFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "AssemblerFile" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DAssemblerFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DAssemblerFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " AssemblerFile = N'" & FileName & "',"
            End If
        End If


        If DContactFile.Visible Then
            If DContactFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Contact" & "-" & UploadDateTime & "-" & Right(DContactFile.PostedFile.FileName, InStr(StrReverse(DContactFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "Contact" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DContactFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DContactFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " ContactFile = N'" & FileName & "',"
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
                SQl = SQl + " RefFile = N'" & FileName & "',"
            End If
        End If

        SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
        SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
        SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDBCommand1.ExecuteNonQuery()
        '
        'Modify-Start by Joy 09-10-03
        'Price Table�B�z
        Dim i As Integer
        For i = 1 To 10
            DBDataSet1.Clear()
            SQl = "Select * From F_ManufCOPrice "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And SeqNo   =  '" & CStr(i) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQl, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "F_ManufCOPrice")
            If DBDataSet1.Tables("F_ManufCOPrice").Rows.Count > 0 Then
                SQl = "Update F_ManufCOPrice Set "
                Select Case i
                    Case Is = 1
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceA1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceA2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceA3.Text) + "', "
                    Case Is = 2
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceB1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceB2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceB3.Text) + "', "
                    Case Is = 3
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceC1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceC2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceC3.Text) + "', "
                    Case Is = 4
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceD1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceD2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceD3.Text) + "', "
                    Case Is = 5
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceE1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceE2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceE3.Text) + "', "
                    Case Is = 6
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceF1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceF2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceF3.Text) + "', "
                    Case Is = 7
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceG1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceG2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceG3.Text) + "', "
                    Case Is = 8
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceH1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceH2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceH3.Text) + "', "
                    Case Is = 9
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceI1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceI2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceI3.Text) + "', "
                    Case Is = 10
                        SQl = SQl + " Spec = '" + YKK.ReplaceString(DPriceJ1.Text) + "', "
                        SQl = SQl + " Currency = '" + DPriceJ2.SelectedValue + "', "
                        SQl = SQl + " Price = '" + YKK.ReplaceString(DPriceJ3.Text) + "', "
                End Select
                SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
                SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
                SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
                SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
                SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQl = SQl + "   And SeqNo   =  '" & CStr(i) & "'"
            Else
                SQl = "Insert into F_ManufCOPrice "
                SQl = SQl + "( "
                SQl = SQl + "Sts, CompletedTime, FormNo, FormSNo, SeqNo, "
                SQl = SQl + "Spec, Currency, Price, "
                SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '0', "
                SQl = SQl + " '" + NowDateTime + "', "
                SQl = SQl + " '000012', "
                SQl = SQl + " '" + CStr(wFormSno) + "', "
                SQl = SQl + " '" + CStr(i) + "', "
                Select Case i
                    Case Is = 1
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceA1.Text) + "', "
                        SQl = SQl + " '" + DPriceA2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceA3.Text) + "', "
                    Case Is = 2
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceB1.Text) + "', "
                        SQl = SQl + " '" + DPriceB2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceB3.Text) + "', "
                    Case Is = 3
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceC1.Text) + "', "
                        SQl = SQl + " '" + DPriceC2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceC3.Text) + "', "
                    Case Is = 4
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceD1.Text) + "', "
                        SQl = SQl + " '" + DPriceD2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceD3.Text) + "', "
                    Case Is = 5
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceE1.Text) + "', "
                        SQl = SQl + " '" + DPriceE2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceE3.Text) + "', "
                    Case Is = 6
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceF1.Text) + "', "
                        SQl = SQl + " '" + DPriceF2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceF3.Text) + "', "
                    Case Is = 7
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceG1.Text) + "', "
                        SQl = SQl + " '" + DPriceG2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceG3.Text) + "', "
                    Case Is = 8
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceH1.Text) + "', "
                        SQl = SQl + " '" + DPriceH2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceH3.Text) + "', "
                    Case Is = 9
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceI1.Text) + "', "
                        SQl = SQl + " '" + DPriceI2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceI3.Text) + "', "
                    Case Is = 10
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceJ1.Text) + "', "
                        SQl = SQl + " '" + DPriceJ2.SelectedValue + "', "
                        SQl = SQl + " '" + YKK.ReplaceString(DPriceJ3.Text) + "', "
                End Select
                SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '�@����
                SQl = SQl + " '" + NowDateTime + "', "       '�@���ɶ�
                SQl = SQl + " '" + "" + "', "                       '�ק��
                SQl = SQl + " '" + NowDateTime + "' "       '�ק�ɶ�
                SQl = SQl + " ) "
            End If
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQl
            OleDBCommand1.ExecuteNonQuery()
        Next
        OleDbConnection1.Close()
        'Modify-end by Joy
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyTranData)
    '**     ��s������
    '**
    '*****************************************************************
    Sub ModifyTranData(ByVal pFun As String, ByVal pSts As String)
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        If pFun <> "SAVE" Then      '<> Save
            SQl = "Update T_WaitHandle Set "
            SQl = SQl + " Active = '" & "0" & "',"
            SQl = SQl + " Sts = '" & pSts & "',"
            If pSts = "1" Then SQl = SQl + " StsDes = '" & BOK.Value & "',"
            If pSts = "2" Then SQl = SQl + " StsDes = '" & BNG1.Value & "',"
            If pSts = "3" Then SQl = SQl + " StsDes = '" & BNG2.Value & "',"

            SQl = SQl + " AEndTime = '" & NowDateTime & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
            If DDelay.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & YKK.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And Step    =  '" & CStr(wStep) & "'"
            SQl = SQl + "   And SeqNo   =  '" & CStr(wSeqNo) & "'"
            SQl = SQl + "   And Active =  '1' "
        Else
            SQl = "Update T_WaitHandle Set "
            If DReasonCode.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & YKK.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And Step =  '" & CStr(wStep) & "'"
            SQl = SQl + "   And SeqNo =  '" & CStr(wSeqNo) & "'"
        End If
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Add T_CommissionNo)
    '**     �l�[�����ƩM�e�U���Ӫ�
    '**
    '*****************************************************************
    Sub AddCommissionNo(ByVal pFormNo As String, ByVal pFormSno As Integer)
        Dim SQl As String
        Dim DBDataSet1 As New DataSet

        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand
        OleDbConnection1.Open()

        SQl = "Select * From T_CommissionNo "
        SQl = SQl & " Where FormNo =  '" & pFormNo & "'"
        SQl = SQl & "   And FormSno =  '" & CStr(pFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQl, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "T_CommissionNo")

        If DBDataSet1.Tables("T_CommissionNo").Rows.Count <= 0 Then
            If DNo.Text <> "" Then
                SQl = "Insert into T_CommissionNo "
                SQl = SQl + "( "
                SQl = SQl + "FormNo, FormSno, No, MapNo, CreateUser, CreateTime "    '1~5
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '" + pFormNo + "', "
                SQl = SQl + " '" + CStr(pFormSno) + "', "
                SQl = SQl + " '" + DNo.Text + "', "
                SQl = SQl + " '" + DMapNo.Text + "', "
                SQl = SQl + " '" + Request.QueryString("pUserID") + "', "
                SQl = SQl + " '" + NowDateTime + "' "
                SQl = SQl + " ) "
                OleDBCommand1.Connection = OleDbConnection1
                OleDBCommand1.CommandText = SQl
                OleDBCommand1.ExecuteNonQuery()
            End If
        Else
            If DNo.Text <> "" Then
                If DNo.Text <> DBDataSet1.Tables("T_CommissionNo").Rows(0).Item("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DNo.Text & "',"
                    SQl = SQl + " MapNo = '" & DMapNo.Text & "',"
                    SQl = SQl + " CreateUser = '" & Request.QueryString("pUserID") & "',"
                    SQl = SQl + " CreateTime = '" & NowDateTime & "' "
                    SQl = SQl + " Where FormNo  =  '" & pFormNo & "'"
                    SQl = SQl + "   And FormSno =  '" & CStr(pFormSno) & "'"
                    OleDBCommand1.Connection = OleDbConnection1
                    OleDBCommand1.CommandText = SQl
                    OleDBCommand1.ExecuteNonQuery()
                End If
            End If
        End If

        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyManufData)
    '**     ��s������
    '**
    '*****************************************************************
    Sub ModifyManufData(ByVal pFun As String, ByVal pSts As Integer, ByVal pContact As Integer)
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        If DOFormNo.Text = "000003" Then
            SQl = "Update F_ManufInSheet Set "
        End If
        If DOFormNo.Text = "000007" Then
            SQl = "Update F_ManufOutSheet Set "
        End If
        If DOFormNo.Text = "000011" Then
            SQl = "Update F_ImportSheet Set "
        End If

        SQl = SQl + " OPSts = '" & CStr(pSts) & "',"
        SQl = SQl + " OPContact = '" & CStr(pContact) & "',"
        SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
        SQl = SQl + " Where FormNo  =  '" & DOFormNo.Text & "'"
        SQl = SQl + "   And FormSno =  '" & DOFormSno.Text & "'"
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �ˬd�W���ɮ�
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As HtmlInputFile) As Integer
        Dim fileExtension As String     '�ŧi�@���ܼƦs���ɮ׮榡(���ɦW)
        Dim allowedExtensions As String() = {".bmp", ".jpg", ".jpeg", ".gif", ".pdf", ".ai", ".xls", ".doc", ".ppt", ".xlsx", ".docx", ".pptx"}   '�w�q���\���ɮ׮榡
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
            If UPFile.PostedFile.ContentLength <= 1500 * 1024 Then
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
    Private Sub BPrint_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BPrint.Click
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

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �y�{�������
    '**
    '*****************************************************************
    Private Sub BFlow_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BFlow.Click
        Dim URL As String = ""
        Dim wLevel As String = "" '������

        URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                              "&pFormSno=" & wFormSno & _
                              "&pStep=" & wStep & _
                              "&pLevel=" & wLevel

        'Call JavaScript
        Dim scriptString As String = "<script language=JavaScript> "
        scriptString = scriptString & "OpenSimulationSheet('" & URL & "'); "
        scriptString = scriptString & "</script>"
        If (Not Me.IsStartupScriptRegistered("Startup")) Then
            Me.RegisterStartupScript("Startup", scriptString)
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ����Button(OK, Ng1, NG2, Save)�B�@ 
    '**
    '*****************************************************************
    Private Sub DisabledButton()
        BOK.Disabled = True
        BNG1.Disabled = True
        BNG2.Disabled = True
        BSAVE.Disabled = True
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �_��Button(OK, Ng1, NG2, Save)�B�@ 
    '**
    '*****************************************************************
    Private Sub EnabledButton()
        BOK.Disabled = False
        BNG1.Disabled = False
        BNG2.Disabled = False
        BSAVE.Disabled = False
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ��J�ˬd
    '**
    '*****************************************************************
    Function InputCheck() As Integer
        InputCheck = 0
        'No
        If InputCheck = 0 Then
            If FindFieldInf("No") = 1 Then
                If DNo.Text = "" Then InputCheck = 1
            End If
        End If
        '���
        If InputCheck = 0 Then
            If FindFieldInf("Date") = 1 Then
                If DDate.Text = "" Then InputCheck = 1
            End If
        End If
        '����
        If InputCheck = 0 Then
            If FindFieldInf("Division") = 1 Then
                If DDivision.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '���
        If InputCheck = 0 Then
            If FindFieldInf("Person") = 1 Then
                If DPerson.SelectedValue = "" Then InputCheck = 1
            End If
        End If

        '�եߧP�w
        If InputCheck = 0 Then
            If FindFieldInf("Assembler") = 1 Then
                If DAssembler.SelectedValue = "" Then InputCheck = 1
            End If
        End If

        'Buyer
        If InputCheck = 0 Then
            If FindFieldInf("Buyer") = 1 Then
                If DBuyer.Text = "" Then InputCheck = 1
            End If
        End If
        '�e�U�t��
        If InputCheck = 0 Then
            If FindFieldInf("SellVendor") = 1 Then
                If DSellVendor.Text = "" Then InputCheck = 1
            End If
        End If
        'Slider Code
        If InputCheck = 0 Then
            If FindFieldInf("SliderCode") = 1 Then
                If DSliderCode.Text = "" Then InputCheck = 1
            End If
        End If
        '�ϸ�
        If InputCheck = 0 Then
            If FindFieldInf("MapNo") = 1 Then
                If DMapNo.Text = "" Then InputCheck = 1
            End If
        End If
        '��l��渹�X
        If InputCheck = 0 Then
            If FindFieldInf("OFormNo") = 1 Then
                If DOFormNo.Text = "" Then InputCheck = 1
            End If
        End If
        '��l�渹
        If InputCheck = 0 Then
            If FindFieldInf("OFormSno") = 1 Then
                If DOFormSno.Text = "" Then InputCheck = 1
            End If
        End If
        '�ت�
        If InputCheck = 0 Then
            If FindFieldInf("Target") = 1 Then
                If DTarget.Text = "" Then InputCheck = 1
            End If
        End If
        '���e
        If InputCheck = 0 Then
            If FindFieldInf("Content") = 1 Then
                If DContent.Text = "" Then InputCheck = 1
            End If
        End If
        'EA����
        If InputCheck = 0 Then
            If FindFieldInf("Description") = 1 Then
                If DDescription.Text = "" Then InputCheck = 1
            End If
        End If
        '���ճ��i
        If InputCheck = 0 Then
            If FindFieldInf("QAFile") = 1 Then
                If DQAFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '������
        If InputCheck = 0 Then
            If FindFieldInf("ContactFile") = 1 Then
                If DContactFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '��L����
        If InputCheck = 0 Then
            If FindFieldInf("RefFile") = 1 Then
                If DRefFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If

        '�եߧP�w��
        If InputCheck = 0 Then
            If FindFieldInf("AssemblerFile") = 1 Then
                If DAssemblerFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If


        'Price
        If InputCheck = 0 Then
            If FindFieldInf("Price") = 1 Then
                If DPriceA1.Text = "" Then InputCheck = 1
                If DPriceA2.SelectedValue = "" Then InputCheck = 1
                If DPriceA3.Text = "" Then InputCheck = 1
                If DPriceB1.Text = "" Then InputCheck = 1
                If DPriceB2.SelectedValue = "" Then InputCheck = 1
                If DPriceB3.Text = "" Then InputCheck = 1
                If DPriceC1.Text = "" Then InputCheck = 1
                If DPriceC2.SelectedValue = "" Then InputCheck = 1
                If DPriceC3.Text = "" Then InputCheck = 1
                If DPriceD1.Text = "" Then InputCheck = 1
                If DPriceD2.SelectedValue = "" Then InputCheck = 1
                If DPriceD3.Text = "" Then InputCheck = 1
                If DPriceE1.Text = "" Then InputCheck = 1
                If DPriceE2.SelectedValue = "" Then InputCheck = 1
                If DPriceE3.Text = "" Then InputCheck = 1
                If DPriceF1.Text = "" Then InputCheck = 1
                If DPriceF2.SelectedValue = "" Then InputCheck = 1
                If DPriceF3.Text = "" Then InputCheck = 1
                If DPriceG1.Text = "" Then InputCheck = 1
                If DPriceG2.SelectedValue = "" Then InputCheck = 1
                If DPriceG3.Text = "" Then InputCheck = 1
                If DPriceH1.Text = "" Then InputCheck = 1
                If DPriceH2.SelectedValue = "" Then InputCheck = 1
                If DPriceH3.Text = "" Then InputCheck = 1
                If DPriceI1.Text = "" Then InputCheck = 1
                If DPriceI2.SelectedValue = "" Then InputCheck = 1
                If DPriceI3.Text = "" Then InputCheck = 1
                If DPriceJ1.Text = "" Then InputCheck = 1
                If DPriceJ2.SelectedValue = "" Then InputCheck = 1
                If DPriceJ3.Text = "" Then InputCheck = 1
            End If
        End If
    End Function
End Class
