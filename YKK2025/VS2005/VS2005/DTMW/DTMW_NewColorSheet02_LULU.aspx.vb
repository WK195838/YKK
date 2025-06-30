Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DTMW_NewColorSheet02_LULU
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
    Dim wbFormSno As Integer        '�s��_��-���y����
    Dim wbStep As Integer           '�s��_��-�u�{�N�X
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
    Dim LastStep As String






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


        LSheetName.Text = "����ZIPPER CHAIN (LULU)" '���W��

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ��s������
    '**
    '*****************************************************************
    Sub UpdateTranFile()
        'Modify-Start by Joy 2009/11/20(2010��ƾ����)
        'oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDepo,wUserID)
        '
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, wUserID)
        '��渹�X,���y����,�u�{���d���X,�Ǹ�,��ƾ�,ñ�֪�
        'Modify-End
    End Sub

    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowFormData()


        Dim SQL As String

        SQL = "Select * From F_NewColorLULU "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then


            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            '   DBarCode.ImageUrl = "imga.aspx?mycode=" & DNo.Text  BARCODE
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
            SetFieldData("ColorLight3", DBAdapter1.Rows(0).Item("ColorLight3"))    '���O2
            SetFieldData("NewOldColor", DBAdapter1.Rows(0).Item("NewOldColor"))    '���O2

            If Mid(DBAdapter1.Rows(0).Item("DeliveryDate").ToString, 1, 4) = "1900" Then
                DDeliveryDate.Value = ""
            Else
                DDeliveryDate.Value = DBAdapter1.Rows(0).Item("DeliveryDate")               '�U��ɶ�
            End If

            If wStep = 500 Then  '(NG�A�e�X�Ʊ����ݭ��s���)
                DDeliveryDate.Value = ""
            End If


            DDuplicateNo.Text = DBAdapter1.Rows(0).Item("DuplicateNo")
            If DBAdapter1.Rows(0).Item("NOCCS") = 1 Then
                DNOCCS.Checked = True
            End If
            DNOCCSReason.Text = DBAdapter1.Rows(0).Item("NOCCSReason")
            DCustomerNGColor.Text = DBAdapter1.Rows(0).Item("CustomerNGColor")
            DRemark.Text = DBAdapter1.Rows(0).Item("Remark")
            SetFieldData("CustomerSample", DBAdapter1.Rows(0).Item("CustomerSample"))    '���O2

            DColorSystem.Text = DBAdapter1.Rows(0).Item("ColorSystem")

            DCheckNo.Text = DBAdapter1.Rows(0).Item("CheckNo")
            DPFBWire.Value = DBAdapter1.Rows(0).Item("PFBWire")
            DPFOpenParts.Value = DBAdapter1.Rows(0).Item("PFOpenParts")

            SetFieldData("YKKColorType", DBAdapter1.Rows(0).Item("YKKColorType"))    '���O2

            DYKKColorCode.Value = DBAdapter1.Rows(0).Item("YKKColorCode")
            DYKKColorCodeVF.Value = DBAdapter1.Rows(0).Item("YKKColorCodeVF")
            DYKKColorCodeSLD.Value = DBAdapter1.Rows(0).Item("YKKColorCodeSLD")


            Dim Version As Integer

            '�p��O�ĴX��
            SQL = " select  count(*)cun from  t_waithandle"
            SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & " and  ( "
            SQL = SQL & "  ( step in (525,25)  and sts =3)"
            SQL = SQL & " or "
            SQL = SQL & "  ( step in (20,520)  and sts =1)"
            SQL = SQL & " )"

            Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
            If DBAdapter3.Rows.Count > 0 Then
                If (wStep = 520 Or wStep = 525) Then
                    Version = DBAdapter3.Rows(0).Item("cun") + 1
                Else
                    Version = DBAdapter3.Rows(0).Item("cun")
                End If

            Else
                Version = 1
            End If

            DVersion.Text = Version
            DSLDColor.Text = DBAdapter1.Rows(0).Item("SLDColor")

            SetFieldData("ColorType", DBAdapter1.Rows(0).Item("Again"))    '���O2
            If DBAdapter1.Rows(0).Item("again") = 1 Then
                DAgain.SelectedValue = "�H��"
            ElseIf DBAdapter1.Rows(0).Item("again") = 2 Then
                DAgain.SelectedValue = "�@��"
            End If

            DVFColor.Text = DBAdapter1.Rows(0).Item("VFColor")
            DVFColor1.Text = DBAdapter1.Rows(0).Item("VFColor1")

            If DBAdapter1.Rows(0).Item("Complete") = 1 Then       '���s��̿৹����
                LComplete.NavigateUrl = "NewColoCompleteList.aspx?pNo=" & DNo.Text
                LComplete.Visible = True
            Else
                LComplete.Visible = False
            End If

            'If DBAdapter1.Rows(0).Item("DTNO") = 1 Then       '���l�[�֥i��
            ' LDTSheet.NavigateUrl = "NewColorDTSheetList.aspx?pNo=" & DNo.Text
            ' LDTSheet.Visible = True
            'Else
            '   LDTSheet.Visible = False
            'End If

        End If


        '�֩w�i�����
        SQL = "SELECT "
        SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
        SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
        SQL = SQL + "AEndTimeDesc As Description, "
        SQL = SQL + "URL "
        SQL = SQL + "FROM V_WaitHandle_01 "
        SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
        SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
        SQL = SQL + "Order by Unique_ID Desc "
        Dim dtWaitHandle1 As DataTable = uDataBase.GetDataTable(SQL)

        'Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter4.Fill(DBDataSet9, "DecideHistory")
        GridView2.DataSource = dtWaitHandle1
        GridView2.DataBind()

        'DB�s������
        'OleDbConnection1.Close()

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
                SQL = SQL & " and dkey = 'DevYear' order by data"


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



    Protected Sub DYKKColorCode_ServerChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles DYKKColorCode.ServerChange

    End Sub
End Class

