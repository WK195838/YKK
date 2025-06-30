Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DTMW_NewColorSheet02_05CNLSBS16
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܼ�
    '**
    '*****************************************************************
    Dim FieldName(60) As String     '�U���
    Dim Attribute(60) As Integer    '�U����ݩ�    
    Dim Top As String               '�w�]�������m
    Dim wFormNo As String           '��渹�X
    Dim wFormSno As Integer         '���y����
    Dim wStep As Integer            '�u�{�N�X
    Dim wSeqNo As Integer           '�Ǹ�
    Dim wApplyID As String          '�ӽЪ�ID
    Dim wUserID As String          '�ϥΪ�ID
    Dim wAgentID As String          '�Q�N�z�HID
    Dim NowDateTime As String       '�{�b����ɶ�
    Dim wNextGate As String         '�U�@���H
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    'Modify-Start by Joy 2009/11/20(2010��ƾ����)
    'Dim wDepo As String = "CL"      '�x�_��ƾ�(CL->���c, TP->�x�_)
    '
    '�s�զ�ƾ�
    Dim wApplyCalendar As String = ""       '�ӽЪ�
    Dim wDecideCalendar As String = ""      '�֩w��
    Dim wNextGateCalendar As String = ""    '�U�@�֩w��
    'Modify-End

    Dim wUserName As String = ""    '�m�W�N�z��
    Dim HolidayList As New List(Of Integer) '�ΥH�O�����骺�����ޭ�
    Dim indexList As New List(Of Integer)   '�ΥH�O�����ݩ�������������ޭ�
    Dim DateList As New List(Of String)     '�ΥH�O���ҿ�����@�P���

    ''' <summary>
    ''' �H�U���@�Ψ禡���ŧi
    ''' </summary>
    ''' <remarks></remarks>
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim YKK As New YKK_SPDClass   'YKK SPD�t�@�q�[��

    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    Dim aAgentID As String
    Dim Makemap1, Makemap2, Makemap3, Makemap4, Makemap5, Makemap6 As Integer
    Dim AQty, DevNo, Manufout As String
    Dim isOK As Boolean = True
    Dim Message As String = ""




    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '�{�b���
        wFormNo = Request.QueryString("pFormNo")    '��渹�X
        wFormSno = Request.QueryString("pFormSno")  '���y����
        LFormSno.Text = "�渹:" + CStr(wFormSno)



        If Now.ToString("yyyy/MM/dd") < "2025/04/01" Then
            LSheetName.Text = "����ZIPPER CHAIN (05 CNL SBS16)" '���W��"
        Else
            LSheetName.Text = "���O�ഫ-05CNPBS16" '���W��"
        End If

    End Sub



    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim SQL As String

        SQL = "Select * From F_NewColor05CNLSBS16 "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then

            If DBAdapter1.Rows(0).Item("CustomerCheck") = 1 Then
                DCustomerCheck.Checked = True
            End If
            If DBAdapter1.Rows(0).Item("FactoryCheck") = 1 Then
                DFactoryCheck.Checked = True
            End If
            If DBAdapter1.Rows(0).Item("VCACheck") = 1 Then
                DVCACheck.Checked = True
            End If

            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            '  DBarCode.ImageUrl = "imga.aspx?mycode=" & DNo.Text  BARCODE
            DDate.Text = DBAdapter1.Rows(0).Item("Date")                         'No 
            DDepoName.Text = DBAdapter1.Rows(0).Item("DepoName")                         'No
            DName.Text = DBAdapter1.Rows(0).Item("Name")                         'No
            DCustomer.Text = DBAdapter1.Rows(0).Item("Customer")                         'No
            DCustomerCode.Text = DBAdapter1.Rows(0).Item("CustomerCode")                         'No
            DBuyer.Text = DBAdapter1.Rows(0).Item("Buyer")                         'No
            DBuyerCode.Text = DBAdapter1.Rows(0).Item("BuyerCode")
            DCustomerColor.Text = DBAdapter1.Rows(0).Item("CustomerColor")                         'No3
            DCustomerColorCode.Text = DBAdapter1.Rows(0).Item("CustomerColorCode")
            DOverseaYKKCode.Text = DBAdapter1.Rows(0).Item("OverseaYKKCode")
            DReColorCode.Text = DBAdapter1.Rows(0).Item("RecolorCode")

            DPANTONECode.Text = DBAdapter1.Rows(0).Item("PANTONECode")
            SetFieldData("DevYear", DBAdapter1.Rows(0).Item("DevYear"))    '�~
            SetFieldData("DevSeason", DBAdapter1.Rows(0).Item("DevSeason"))    '�u

            If Mid(DBAdapter1.Rows(0).Item("ReceiveDate").ToString, 1, 4) = "1900" Then
                DReceiveDate.Text = ""
            Else
                DReceiveDate.Text = DBAdapter1.Rows(0).Item("ReceiveDate")               '�U��ɶ�
            End If


            SetFieldData("ColorLight1", DBAdapter1.Rows(0).Item("ColorLight1"))    '���O1
            SetFieldData("ColorLight2", DBAdapter1.Rows(0).Item("ColorLight2"))    '���O2
            SetFieldData("NewOldColor", DBAdapter1.Rows(0).Item("NewOldColor"))    '���O2

            DDeliveryDate.Text = DBAdapter1.Rows(0).Item("DeliveryDate")
            DDuplicateNo.Text = DBAdapter1.Rows(0).Item("DuplicateNo")
            If DBAdapter1.Rows(0).Item("NOCCS") = 1 Then
                DNOCCS.Checked = True
            End If
            DNOCCSReason.Text = DBAdapter1.Rows(0).Item("NOCCSReason")
            DCustomerNGColor.Text = DBAdapter1.Rows(0).Item("CustomerNGColor")
            DRemark.Text = DBAdapter1.Rows(0).Item("Remark")
            SetFieldData("CustomerSample", DBAdapter1.Rows(0).Item("CustomerSample"))    '���O2

            DColorSystem.Text = DBAdapter1.Rows(0).Item("ColorSystem")
            DPFBWire.Text = DBAdapter1.Rows(0).Item("PFBWire")
            DPFOpenParts.Text = DBAdapter1.Rows(0).Item("PFOpenParts")
            SetFieldData("YKKColorType", DBAdapter1.Rows(0).Item("YKKColorType"))    '���O2
            DYKKColorCode.Text = DBAdapter1.Rows(0).Item("YKKColorCode")
            DVersion.Text = DBAdapter1.Rows(0).Item("Version")
            DSLDColor.Text = DBAdapter1.Rows(0).Item("SLDColor")
            SetFieldData("ColorType", DBAdapter1.Rows(0).Item("Again"))    '���O2
            If DBAdapter1.Rows(0).Item("again") = 1 Then
                DAgain.SelectedValue = "�H��"
            ElseIf DBAdapter1.Rows(0).Item("again") = 2 Then
                DAgain.SelectedValue = "�@��"
            End If

            DVFColor.Text = DBAdapter1.Rows(0).Item("VFColor")
            DCheckNo.Text = DBAdapter1.Rows(0).Item("CheckNo")
            If DBAdapter1.Rows(0).Item("Complete") = 1 Then       '���s��̿৹����
                LComplete.NavigateUrl = "NewColoCompleteList.aspx?pNo=" & DNo.Text
                LComplete.Visible = True
            Else
                LComplete.Visible = False
            End If



            'If DBAdapter1.Rows(0).Item("Again") = 1 Then      '���s��β{������
            '    LAgain.NavigateUrl = "NewColorAgainList.aspx?pNo=" & DNo.Text
            '    LAgain.Visible = True
            'Else
            '    LAgain.Visible = False
            'End If

        End If







    End Sub


    '*****************************************************************
    '**(ShowSheetField)
    '**     �ظm�U�Ԧ�����l��
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim DBDataSet1 As New DataSet

        Dim SQL As String
        Dim idx As Integer
        Dim i As Integer



        idx = FindFieldInf(pFieldName)


        '����
        If pFieldName = "ColorLight1" Then
            DColorLight1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DColorLight1.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'ColorLight'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DColorLight1.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DColorLight1.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        '����
        If pFieldName = "ColorLight2" Then
            DColorLight2.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DColorLight2.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  data from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'ColorLight2' "
                SQL = SQL & " order by createtime desc "


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DColorLight2.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DColorLight2.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        'YKK��O
        If pFieldName = "YKKColorType" Then
            DYKKColorType.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DYKKColorType.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'YKKColorType'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DYKKColorType.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DYKKColorType.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'YKK��O
        If pFieldName = "ColorType" Then
            DAgain.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAgain.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'ColorType'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DAgain.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAgain.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If




        'YKK��O
        If pFieldName = "CustomerSample" Then

            DCustomerSample.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCustomerSample.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'CustomerSample'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DCustomerSample.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCustomerSample.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        '�}�o�~
        If pFieldName = "DevYear" Then

            DDevYear.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDevYear.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'DevYear'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DDevYear.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDevYear.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        '�}�o�u
        If pFieldName = "DevSeason" Then

            DDevSeason.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDevSeason.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'DevSeason'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DDevSeason.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDevSeason.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        '�s���¦�
        If pFieldName = "NewOldColor" Then
            DNewOldColor.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DNewOldColor.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'NewOldColor'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DNewOldColor.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DNewOldColor.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

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


End Class

