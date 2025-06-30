Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ContactOutSheet
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BDate As System.Web.UI.WebControls.Button
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DLevel As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BMMapNo As System.Web.UI.WebControls.Button
    Protected WithEvents BOMapNo As System.Web.UI.WebControls.Button
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTarget As System.Web.UI.WebControls.TextBox
    Protected WithEvents DContent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents BOK As System.Web.UI.WebControls.Button
    Protected WithEvents BNG As System.Web.UI.WebControls.Button
    Protected WithEvents BSave As System.Web.UI.WebControls.Button
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDescSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DDelivery As System.Web.UI.WebControls.Image
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents LBefOP As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents BIn As System.Web.UI.WebControls.Button
    Protected WithEvents DOFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents LAttachFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DNFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents LMapNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOFormNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents BModify As System.Web.UI.WebControls.Button
    Protected WithEvents LNFormNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DContactSheet As System.Web.UI.WebControls.Image
    Protected WithEvents Form1 As System.Web.UI.HtmlControls.HtmlForm
    Protected WithEvents DAttachFile As System.Web.UI.HtmlControls.HtmlInputFile

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
    Dim wStep As Integer            '�u�{�N�X
    Dim wSeqNo As Integer           '�Ǹ�
    Dim wApplyID As String          '�ӽЪ�ID
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
            If wFormSno > 0 And wStep > 1 Then    '�P�_�O�_[ñ��]
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
        Response.Cookies("DevNo").Value = ""        '�}�oNo, DevNoPicker�ϥ�
        Response.Cookies("MapNo").Value = ""        '�ϸ�, MapPicker�ϥ�
        Response.Cookies("Step").Value = Request.QueryString("pStep")  '���åΤu�{�N�X
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     '�W�Ǹ���ˬd����ܰT��
    '**
    '*****************************************************************
    Sub ShowMessage()
        Dim Message As String = ""
        'Check����
        If DAttachFile.Visible Then
            If DAttachFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "����"
                Else
                    Message = Message & ", " & "����"
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
    '**
    '**     ��s������
    '**
    '*****************************************************************
    Sub UpdateTranFile()
        Dim oUpdateFlow As Object
        oUpdateFlow = Server.CreateObject("FlowEngine.WFFlowEngine")
        oUpdateFlow.pFormNo = wFormNo      '��渹�X
        oUpdateFlow.pFormSno = wFormSno    '���y����
        oUpdateFlow.pStep = wStep          '�u�{���d���X
        oUpdateFlow.pSeqNo = wSeqNo        '�Ǹ�
        oUpdateFlow.UpdateFlow(Request.Cookies("UserID").Value)
    End Sub

    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ContactOutFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_ContactOutSheet "
        SQL = SQL & " Where Sts = 0 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ContactOutSheet")
        If DBDataSet1.Tables("F_ContactOutSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("Date")                   '���
            SetFieldData("Division", DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("Division"))  '����
            SetFieldData("Person", DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("Person"))      '���
            DSliderCode.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("SliderCode")       'Slider Code
            SetFieldData("Level", DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("Level"))                '������
            DMapNo.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("MapNo")             '�ϸ�
            If DMapNo.Text <> "" Then
                SQL = "Select FormNo, FormSno From F_MapSheet "
                SQL = SQL & " Where Sts = 2 "
                SQL = SQL & "   And MapNo =  '" & DMapNo.Text & "'"
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet1, "MapSheet")
                If DBDataSet1.Tables("MapSheet").Rows.Count > 0 Then
                    LMapNo.NavigateUrl = "MapSheet_02.aspx?pFormNo=" & DBDataSet1.Tables("MapSheet").Rows(0).Item("FormNo") & _
                                                         "&pFormSno=" & DBDataSet1.Tables("MapSheet").Rows(0).Item("FormSno")
                Else
                    SQL = "Select FormNo, FormSno From F_ModMapSheet "
                    SQL = SQL & " Where Sts = 2 "
                    SQL = SQL & "   And MapNo =  '" & DMapNo.Text & "'"
                    Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter4.Fill(DBDataSet1, "ModMapSheet")
                    If DBDataSet1.Tables("ModMapSheet").Rows.Count > 0 Then
                        LMapNo.NavigateUrl = "MapSheetMod_02.aspx?pFormNo=" & DBDataSet1.Tables("ModMapSheet").Rows(0).Item("FormNo") & _
                                                                "&pFormSno=" & DBDataSet1.Tables("ModMapSheet").Rows(0).Item("FormSno")
                    End If
                End If
            Else
                LMapNo.Visible = False
            End If

            DOFormNo.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("OFormNo")             '�ϸ�
            DOFormSno.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("OFormSno")             '�ϸ�
            If DOFormNo.Text <> "" And DOFormSno.Text <> "" Then
                If DOFormNo.Text = "000003" Then
                    LOFormNo.NavigateUrl = "ManufInSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                Else
                    LOFormNo.NavigateUrl = "ManufOutSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                End If
            Else
                LOFormNo.Visible = False
            End If

            DNFormNo.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("NFormNo")             '�ϸ�
            DNFormSno.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("NFormSno")             '�ϸ�
            If DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("NFormSno") = 0 Then
                DNFormSno.Text = ""
            Else
                DNFormSno.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("NFormSno")             '�ϸ�
            End If
            If DNFormNo.Text <> "" And DNFormSno.Text <> "" Then
                LNFormNo.NavigateUrl = "ModManufOutSheet_01.aspx?pFormNo=" & DNFormNo.Text & "&pFormSno=" & CInt(DNFormSno.Text) & "&pOFormNo=" & DOFormNo.Text & "&pOFormSno=" & CInt(DOFormSno.Text) & "&pStep=" & wStep
                If BModify.Visible = True Then
                    LNFormNo.Visible = False
                End If
            Else
                LNFormNo.Visible = False
            End If

            DTarget.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("Target")             '�ϸ�
            DContent.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("Content")             '�ϸ�
            DDReason.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("Reason")             '�ϸ�
            If DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("AttachFile") <> "" Then          '����1
                LAttachFile.NavigateUrl = Path & DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("AttachFile")
            Else
                LAttachFile.Visible = False
            End If

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
                    SetFieldData("ReasonCode", DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonCode"))    '����z�ѥN�X
                    If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonCode") = "" Then
                        SetFieldData("Reason", DReasonCode.SelectedValue)    '����z��
                        DReasonDesc.Text = ""   '�����L����
                    Else
                        DReason.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("Reason")  '����z��
                        DReasonDesc.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonDesc")     '�����L����
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
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        If wFormSno > 0 And wStep > 1 Then    '�P�_�O�_[ñ��]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                'Sheet���
                DContactSheet.Visible = True   '���Sheet-1
                DDescSheet.Visible = True       '����Sheet
                DDelivery.Visible = True        '���Sheet
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") < NowDateTime Then
                    DDelay.Visible = True   '����Sheet
                Else
                    DDelay.Visible = False  '����Sheet
                End If
                '������
                DDecideDesc.Visible = True      '����
                DBStartTime.Visible = True      '�w�w�}�l
                DBEndTime.Visible = True        '�w�w����
                DAStartTime.Visible = True      '��ڶ}�l
                DAEndTime.Visible = True        '��ڧ���
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") < NowDateTime Then
                    DReasonCode.Visible = True     '����z�ѥN�X
                    DReasonCode.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonCodeRqd", "DReasonCode", "���`�G�ݿ�J����z��")
                    DReason.Visible = True         '����z��
                    DReason.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonRqd", "DReason", "���`�G�ݿ�J����z��")
                    DReasonDesc.Visible = True     '�����L����
                Else
                    DReasonCode.Visible = False     '����z�ѥN�X
                    DReason.Visible = False         '����z��
                    DReasonDesc.Visible = False     '�����L����
                End If
                '�s�����---�ݦA�ק�
                LMapNo.Visible = True          '����
                LOFormNo.Visible = True        '��e�U
                LNFormNo.Visible = True        '�s�e�U
                LAttachFile.Visible = True     '����
                LBefOP.Visible = True          '�u�{�i��
                '���s��m
                BNG.Style.Add("Top", Top)      'NG���s
                BSave.Style.Add("Top", Top)    '�x�s���s
                BOK.Style.Add("Top", Top)      'OK���s
                DFormSno.Style.Add("Top", Top) '�渹
            End If
        Else
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
            '�s������
            LMapNo.Visible = False          '����
            LOFormNo.Visible = False        '��e�U
            LNFormNo.Visible = False        '�s�e�U
            LAttachFile.Visible = False     '����
            LBefOP.Visible = False          '�u�{�i��
            '���s��m
            BNG.Style.Add("Top", Top)       'NG���s
            BSave.Style.Add("Top", Top)     '�x�s���s
            BOK.Style.Add("Top", Top)       'OK���s
            DFormSno.Style.Add("Top", Top)  '�渹
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
                BSave.Visible = True
            Else
                BSave.Visible = False
            End If
            'NG���s
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun") = 1 Then
                BNG.Visible = True
            Else
                BNG.Visible = False
            End If
            'OK���s
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("OKFun") = 1 Then
                BOK.Visible = True
            Else
                BOK.Visible = False
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
        BIn.Attributes("onclick") = "DevNoPicker('Out','000009');"       '���s
        BOMapNo.Attributes("onclick") = "MapPicker('Ori');"             '��l�ϸ�
        BMMapNo.Attributes("onclick") = "MapPicker('Mod');"             '�ק�ϸ�
        BModify.Attributes("onclick") = "ModifySheet();"                '��e�U
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
        If wFormSno > 0 And wStep > 1 Then    '�P�_�O�_[ñ��]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") >= NowDateTime Then
                    Top = 648
                Else
                    Top = 752
                End If
            End If
        Else
            Top = 464
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
        Dim oFieldAttribute As Object
        oFieldAttribute = Server.CreateObject("GetFieldAttb.WFField")
        oFieldAttribute.pFormNo = wFormNo      '��渹�X
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
                BIn.Visible = False
            Case 1  '�ק�+�ˬd
                DNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DnoRqd", "Dno", "���`�G�ݿ�J�ܢ�")
                DNo.Visible = True
                BIn.Visible = True
            Case 2  '�ק�
                DNo.BackColor = Color.Yellow
                DNo.Visible = True
                BIn.Visible = True
            Case Else   '����
                DNo.Visible = False
                BIn.Visible = False
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
        'Slider Code
        Select Case FindFieldInf("SliderCode")
            Case 0  '���
                DSliderCode.BackColor = Color.PeachPuff
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
        'OFormNo
        Select Case FindFieldInf("OFormNo")
            Case 0  '���
                DOFormNo.BackColor = Color.PeachPuff
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
                DOFormSno.BackColor = Color.PeachPuff
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
        'NFormNo
        Select Case FindFieldInf("NFormNo")
            Case 0  '���
                DNFormNo.BackColor = Color.PeachPuff
                DNFormNo.ReadOnly = True
                DNFormNo.Visible = True
                BModify.Visible = False
            Case 1  '�ק�+�ˬd
                DNFormNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNFormNoRqd", "DNFormNo", "���`�G�ݿ�J��渹�X")
                DNFormNo.Visible = True
                BModify.Visible = True
            Case 2  '�ק�
                DNFormNo.BackColor = Color.Yellow
                DNFormNo.Visible = True
                BModify.Visible = True
            Case Else   '����
                DNFormNo.Visible = False
                BModify.Visible = False
        End Select
        If pPost = "New" Then DNFormNo.Text = ""
        'NFormSno
        Select Case FindFieldInf("NFormSno")
            Case 0  '���
                DNFormSno.BackColor = Color.PeachPuff
                DNFormSno.ReadOnly = True
                DNFormSno.Visible = True
                BModify.Visible = False
            Case 1  '�ק�+�ˬd
                DNFormSno.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNFormSnoRqd", "DNFormSno", "���`�G�ݿ�J�渹")
                DNFormSno.Visible = True
                BModify.Visible = True
            Case 2  '�ק�
                DNFormSno.BackColor = Color.Yellow
                DNFormSno.Visible = True
                BModify.Visible = True
            Case Else   '����
                DNFormSno.Visible = False
                BModify.Visible = False
        End Select
        If pPost = "New" Then DNFormSno.Text = ""
        'Target
        Select Case FindFieldInf("Target")
            Case 0  '���
                DTarget.BackColor = Color.PeachPuff
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
        'Content
        Select Case FindFieldInf("Content")
            Case 0  '���
                DContent.BackColor = Color.PeachPuff
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
        'Reason
        Select Case FindFieldInf("Reason")
            Case 0  '���
                DDReason.BackColor = Color.PeachPuff
                DDReason.ReadOnly = True
                DDReason.Visible = True
            Case 1  '�ק�+�ˬd
                DDReason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDReasonRqd", "DDReason", "���`�G�ݿ�J��]")
                DDReason.Visible = True
            Case 2  '�ק�
                DDReason.BackColor = Color.Yellow
                DDReason.Visible = True
            Case Else   '����
                DDReason.Visible = False
        End Select
        If pPost = "New" Then DDReason.Text = ""
        '����
        Select Case FindFieldInf("AttachFile")
            Case 0  '���
                DAttachFile.Visible = False
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DAttachFileRqd", "DAttachFile", "���`�G�ݿ�J����")
                DAttachFile.Visible = True
            Case 2  '�ק�
                DAttachFile.Visible = True
            Case Else   '����
                DAttachFile.Visible = False
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
        '����z�ѥN�X
        If pFieldName = "ReasonCode" Then
            SQL = "Select * From M_Referp Where Cat='014' Order by DKey, Data "
            DReasonCode.Items.Clear()
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
    '**     �x�s���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSave.Click
        Dim ErrCode As Integer = 0   '�ŧi�@��Error�ΨӧP�O�O�_��ƥ��`�A�w�]��False
        Dim Message As String = ""
        '�W�Ǥ��
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)

        'Check����Size�ή榡
        If ErrCode = 0 Then
            If DAttachFile.Visible Then
                If DAttachFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DAttachFile)
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
            Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("ContactOutFilePath"))
            Dim FileName As String
            Dim SQL As String

            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
            Dim OleDBCommand1 As New OleDbCommand

            '�x�s�����
            SQL = "Update F_ContactOutSheet Set "
            SQL = SQL + " No = '" & DNo.Text & "',"
            SQL = SQL + " Date = '" & DDate.Text & "',"
            SQL = SQL + " Division = '" & DDivision.SelectedValue & "',"
            SQL = SQL + " Person = '" & DPerson.SelectedValue & "',"
            SQL = SQL + " SliderCode = '" & DSliderCode.Text & "',"
            SQL = SQL + " Level = '" & DLevel.SelectedValue & "',"
            SQL = SQL + " MapNo = '" & DMapNo.Text & "',"
            SQL = SQL + " OFormNo = '" & DOFormNo.Text & "',"
            SQL = SQL + " OFormSno = '" & DOFormSno.Text & "',"
            SQL = SQL + " NFormNo = '" & DNFormNo.Text & "',"
            SQL = SQL + " NFormSno = '" & DNFormSno.Text & "',"
            SQL = SQL + " Target = '" & DTarget.Text & "',"
            SQL = SQL + " Content = '" & DContent.Text & "',"
            SQL = SQL + " Reason = '" & DDReason.Text & "',"

            If DAttachFile.Visible Then
                If DAttachFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    FileName = CStr(wFormSno) & "-" & "Attach" & "-" & UploadDateTime & "-" & Right(DAttachFile.PostedFile.FileName, InStr(StrReverse(DAttachFile.PostedFile.FileName), "\") - 1)
                    Try    '�W�ǹ���
                        DAttachFile.PostedFile.SaveAs(Path + FileName)
                    Catch ex As Exception
                    End Try
                    SQL = SQL + " AttachFile = '" & FileName & "',"
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

            '�x�s������
            SQL = "Update T_WaitHandle Set "
            If DReasonCode.Visible = True Then
                SQL = SQL + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQL = SQL + " Reason = '" & DReason.Text & "',"
                SQL = SQL + " ReasonDesc = '" & DReasonDesc.Text & "',"
            End If
            SQL = SQL + " DecideDesc = '" & DDecideDesc.Text & "',"
            SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
            SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
            SQL = SQL + " Where FormNo  =  '" & wFormNo & "'"
            SQL = SQL + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL + "   And Step =  '" & CStr(wStep) & "'"
            SQL = SQL + "   And SeqNo =  '" & CStr(wSeqNo) & "'"
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDbConnection1.Open()
            OleDBCommand1.ExecuteNonQuery()
            OleDbConnection1.Close()
        Else
            If ErrCode = 9010 Then Message = "�W���ɮ׮榡���~,�нT�{!"
            If ErrCode = 9020 Then Message = "�W���ɮ�Size�W�L1024KB,�нT�{!"
            If ErrCode = 9030 Then Message = "�W���ɮרS���w,�нT�{!"
            If ErrCode = 9210 Then Message = "����z�Ѭ���L�ɻݶ�g����,�нT�{!"
            Response.Write(YKK.ShowMessage(Message))
        End If      '�W���ɮ�ErrCode=0

        If ErrCode = 0 Then
            Dim URL As String = "MessagePage.aspx?pMSGID=S&pFormNo=" & wFormNo & "&pStep=" & wStep & _
                                                "&pUserID=" & Request.Cookies("UserID").Value & "&pApplyID=" & wApplyID
            Response.Redirect(URL)
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BOK.Click
        Dim ErrCode As Integer = 0   '�ŧi�@��Error�ΨӧP�O�O�_��ƥ��`�A�w�]��False
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim Message As String = ""

        'Check����Size�ή榡
        If ErrCode = 0 Then
            If DAttachFile.Visible Then
                If DAttachFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DAttachFile)
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

        If ErrCode = 0 Then
            Dim Run As Boolean = True           '�O�_����
            Dim RepeatRun As Boolean = False    '�O�_���а���
            Dim wLevel As String = DLevel.SelectedValue     '������

            While Run = True
                Run = False     '����Flag=������
                '--���o�U�@���ѼƳ]�w---------
                Dim oNextGate As Object
                Dim pNextGate(10) As String
                Dim pNextStep As Integer = 0
                Dim pFlowType As Integer = 0    '0=�q��
                Dim pCount As Integer
                '--���oLeadTime�ѼƳ]�w---------
                Dim oGetLeadTime As Object
                Dim pCTime, pStartTime, pEndTime As DateTime
                '--���o�u�{�t���ѼƳ]�w---------
                Dim oGetLoading As Object
                Dim pLastTime As DateTime
                Dim pCount1 As Integer
                '--���o���y�����]�w---------
                Dim oGetSeqNo As Object
                '--�y�{��Ƴ]�w---------
                Dim oNextFlow As Object
                '--�l��ǰe---------
                Dim oMail As Object
                '--��{�վ�---------
                Dim oSchedule As Object
                '--��LSubFile---------
                Dim oPutSubFile As Object

                Dim RtnCode, i As Integer
                Dim NewFormSno As Integer = wFormSno    '���y����
                Dim pRunNextStep As Integer = 0         '�O�_����p��U�@��(�|ñ)
                Dim SQL As String

                '���o���y�����Χ�s������
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep = 1 Then    '�P�_�O�_�_��
                        '���o���y����
                        oGetSeqNo = Server.CreateObject("GetSeqno.WFFormInf")
                        RtnCode = oGetSeqNo.Seqno(wFormNo, NewFormSno)   '��渹�X, ���y����
                        If RtnCode <> 0 Then
                            ErrCode = 9110
                        Else
                            '�ӽЬy�{��ƫظm
                            oNextFlow = Server.CreateObject("FlowEngine.WFFlowEngine")
                            oNextFlow.pFormNo = wFormNo       '��渹�X
                            oNextFlow.pFormSno = NewFormSno   '���y����
                            oNextFlow.pStep = 1               '�u�{���d���X
                            oNextFlow.pSeqNo = 1              '�Ǹ�
                            RtnCode = oNextFlow.NewFlow(Request.Cookies("UserID").Value, wApplyID)   '�ظm��, �ӽЪ�
                        End If
                        pRunNextStep = 1
                    Else
                        If RepeatRun = False Then   '���O�q�������а���
                            '��s������
                            Dim OleDbConnection1 As New OleDbConnection
                            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
                            Dim OleDBCommand1 As New OleDbCommand

                            SQL = "Update T_WaitHandle Set "
                            SQL = SQL + " Active = '" & "0" & "',"
                            SQL = SQL + " Sts = '" & "1" & "',"
                            SQL = SQL + " AEndTime = '" & NowDateTime & "',"
                            SQL = SQL + " CompletedTime = '" & NowDateTime & "',"
                            If DDelay.Visible = True Then
                                SQL = SQL + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                                SQL = SQL + " Reason = '" & DReason.Text & "',"
                                SQL = SQL + " ReasonDesc = '" & DReasonDesc.Text & "',"
                            End If
                            SQL = SQL + " DecideDesc = '" & DDecideDesc.Text & "',"
                            SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
                            SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
                            SQL = SQL + " Where FormNo  =  '" & wFormNo & "'"
                            SQL = SQL + "   And FormSno =  '" & CStr(wFormSno) & "'"
                            SQL = SQL + "   And Step    =  '" & CStr(wStep) & "'"
                            SQL = SQL + "   And EveryOne =  '1' "
                            'SQL = SQL + "   And SeqNo   =  '" & CStr(wSeqNo) & "'"
                            OleDBCommand1.Connection = OleDbConnection1
                            OleDBCommand1.CommandText = SQL
                            OleDbConnection1.Open()
                            OleDBCommand1.ExecuteNonQuery()
                            OleDbConnection1.Close()

                            '�y�{��Ƶ���
                            oNextFlow = Server.CreateObject("FlowEngine.WFFlowEngine")
                            oNextFlow.pFormNo = wFormNo   '��渹�X
                            oNextFlow.pFormSno = wFormSno                        '���y����
                            oNextFlow.pStep = wStep       '�u�{���d���X
                            RtnCode = oNextFlow.CheckFlow(Request.Cookies("UserID").Value, pRunNextStep)   '�ظm��, �y�{�����_(�|ñ)
                            If RtnCode <> 0 Then ErrCode = 9120
                        Else
                            pRunNextStep = 1    '�O�q�������а���
                        End If
                    End If
                End If

                '���o�U�@��
                If ErrCode = 0 And pRunNextStep = 1 Then
                    oNextGate = Server.CreateObject("NextGate.WFNextGate")
                    oNextGate.pFormNo = wFormNo     '��渹�X
                    oNextGate.pStep = wStep         '�u�{���d���X
                    oNextGate.pUserID = Request.Cookies("UserID").Value       'ñ�֪�ID
                    oNextGate.pApplyID = wApplyID                                     '�ӽЪ�ID
                    RtnCode = oNextGate.NextGate(pNextStep, pNextGate, pCount, pFlowType)  '�U�@�u�{��, ���X, ����, �H��, �B�z��k 
                    If RtnCode <> 0 Then ErrCode = 9130
                    If pCount = 0 And pNextStep <> 999 Then ErrCode = 9131
                End If

                '�ظm�y�{���
                If ErrCode = 0 And pRunNextStep = 1 Then
                    If pNextStep <> 999 Then
                        For i = 1 To pCount
                            '���o�u�{�t���̫���
                            oGetLoading = Server.CreateObject("GetLoading.WFOPINF")
                            RtnCode = oGetLoading.GetLastTime(pNextGate(i), wFormNo, pNextStep, NowDateTime, pLastTime, pCount1)  '��渹�X, �u�{���X, �}�l���, �̫���, ���
                            '���o�w�w�}�l,������{�p��
                            oGetLeadTime = Server.CreateObject("GetLeadTime.WFOPInf")
                            oGetLeadTime.pFormNo = wFormNo      '��渹�X
                            oGetLeadTime.pStep = pNextStep      '�u�{���X
                            oGetLeadTime.pLevel = wLevel        '������
                            RtnCode = oGetLeadTime.LeadTime(pLastTime, pStartTime, pEndTime)  '�{�b�ɶ�, �w�w�}�l���, �w�w�������
                            '�ظm�y�{���
                            oNextFlow = Server.CreateObject("FlowEngine.WFFlowEngine")
                            oNextFlow.pFormNo = wFormNo       '��渹�X
                            oNextFlow.pFormSno = NewFormSno   '���y����
                            oNextFlow.pStep = pNextStep       '�u�{���d���X
                            oNextFlow.pSeqNo = i              '�Ǹ�
                            oNextGate.pApplyID = wApplyID     '�ӽЪ�ID
                            RtnCode = oNextFlow.NextFlow(Request.Cookies("UserID").Value, pNextGate(i), wApplyID, pStartTime, pEndTime, 0)   '�ظm��, ñ�֪�, �ӽЪ�, �w�w�}�l��, �w�w������, ���n��
                            If RtnCode <> 0 Then
                                ErrCode = 9150
                                Exit For
                            End If
                        Next i
                    Else
                        oNextFlow = Server.CreateObject("FlowEngine.WFFlowEngine")
                        oNextFlow.pFormNo = wFormNo       '��渹�X
                        oNextFlow.pFormSno = wFormSno     '���y����
                        oNextFlow.pStep = pNextStep       '�u�{���d���X
                        oNextFlow.pSeqNo = 1              '�Ǹ�
                        RtnCode = oNextFlow.EndFlow(Request.Cookies("UserID").Value, wApplyID)   '�ظm��, �ӽЪ�
                        If RtnCode <> 0 Then ErrCode = 9160
                    End If
                End If
                '��u�{��{�վ�
                If ErrCode = 0 Then
                    If RepeatRun = False Then
                        oSchedule = Server.CreateObject("Schedule.WFSchedule")
                        RtnCode = oSchedule.AdjustSchedule(Request.Cookies("UserID").Value, wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel)
                    End If
                End If
                '�x�s�����
                If ErrCode = 0 Then
                    'If RepeatRun = False Then
                    Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("ContactOutFilePath"))
                    Dim FileName As String

                    Dim OleDbConnection1 As New OleDbConnection
                    OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
                    Dim OleDBCommand1 As New OleDbCommand

                    If wFormSno = 0 And wStep = 1 Then    '�P�_�O�_�_��
                        If NewFormSno <> 0 Then
                            '���Table�B�z
                            SQL = "Insert into F_ContactOutSheet "
                            SQL = SQL + "( "
                            SQL = SQL + "Sts, CompletedTime, FormNo, FormSNo, "          '1~4
                            SQL = SQL + "No, Date, Division, Person, SliderCode, "       '5~9
                            SQL = SQL + "Level, MapNo, OFormNo, OFormSno, NFormNo, NFormSno, "  '10~14
                            SQL = SQL + "Target, Content, Reason, AttachFile, "          '15~18
                            SQL = SQL + "CreateUser, CreateTime, ModifyUser, ModifyTime "  '19~22
                            SQL = SQL + ")  "
                            SQL = SQL + "VALUES( "
                            SQL = SQL + " '0', "                                '���A(0:����,1:�w��NG,2:�w��OK)
                            SQL = SQL + " '" + NowDateTime + "', "              '���פ�
                            SQL = SQL + " '000009', "                           '���N��
                            SQL = SQL + " '" + CStr(NewFormSno) + "', "         '���y����
                            SQL = SQL + " '" + DNo.Text + "', "                 'NO
                            SQL = SQL + " '" + DDate.Text + "', "               '���
                            SQL = SQL + " '" + DDivision.SelectedValue + "', "  '����
                            SQL = SQL + " '" + DPerson.SelectedValue + "', "    '���
                            SQL = SQL + " '" + DSliderCode.Text + "', "         'Slider Code
                            SQL = SQL + " '" + DLevel.SelectedValue + "', "     '������
                            SQL = SQL + " '" + DMapNo.Text + "', "              '�ϸ�
                            SQL = SQL + " '" + DOFormNo.Text + "', "            '���No
                            If DOFormSno.Text = "" Then                         '�渹
                                SQL = SQL + " '0', "
                            Else
                                SQL = SQL + " '" + DOFormSno.Text + "', "
                            End If
                            SQL = SQL + " '" + DNFormNo.Text + "', "            '���No
                            If DNFormSno.Text = "" Then                         '�渹
                                SQL = SQL + " '0', "
                            Else
                                SQL = SQL + " '" + DNFormSno.Text + "', "
                            End If
                            SQL = SQL + " '" + DTarget.Text + "', "             '�ت�
                            SQL = SQL + " '" + DContent.Text + "', "            '���e
                            SQL = SQL + " '" + DDReason.Text + "', "             '��]

                            FileName = ""
                            If DAttachFile.Visible Then                         '����
                                If DAttachFile.PostedFile.FileName <> "" Then     '�P�_���ɮפW��
                                    FileName = CStr(NewFormSno) & "-" & "Attach" & "-" & UploadDateTime & "-" & Right(DAttachFile.PostedFile.FileName, InStr(StrReverse(DAttachFile.PostedFile.FileName), "\") - 1)
                                    Try    '�W�ǹ���
                                        DAttachFile.PostedFile.SaveAs(Path + FileName)
                                    Catch ex As Exception
                                    End Try
                                End If
                            Else
                                FileName = ""
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

                            'Update ��e�U�檬�A
                            If DOFormNo.Text = "000003" Then
                                SQL = "Update F_ManufInSheet Set "
                            Else
                                SQL = "Update F_ManufOutSheet Set "
                            End If
                            SQL = SQL + " Status = '" & "1" & "',"
                            SQL = SQL + " StatusDesc = '" & "" & "',"
                            SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
                            SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
                            SQL = SQL + " Where FormNo  =  '" & DOFormNo.Text & "'"
                            SQL = SQL + "   And FormSno =  '" & DOFormSno.Text & "'"
                            OleDBCommand1.Connection = OleDbConnection1
                            OleDBCommand1.CommandText = SQL
                            OleDbConnection1.Open()
                            OleDBCommand1.ExecuteNonQuery()
                            OleDbConnection1.Close()

                        End If  'pSeqno <> 0
                    Else    '�P�_�O�_�_��
                        SQL = "Update F_ContactOutSheet Set "
                        If pNextStep = 999 Then     '�u�{������?
                            SQL = SQL + " Sts = '" & "2" & "',"
                        Else
                            SQL = SQL + " Sts = '" & "0" & "',"
                        End If
                        SQL = SQL + " CompletedTime = '" & NowDateTime & "',"
                        SQL = SQL + " No = '" & DNo.Text & "',"
                        SQL = SQL + " Date = '" & DDate.Text & "',"
                        SQL = SQL + " Division = '" & DDivision.SelectedValue & "',"
                        SQL = SQL + " Person = '" & DPerson.SelectedValue & "',"
                        SQL = SQL + " SliderCode = '" & DSliderCode.Text & "',"
                        SQL = SQL + " Level = '" & DLevel.SelectedValue & "',"
                        SQL = SQL + " MapNo = '" & DMapNo.Text & "',"
                        SQL = SQL + " OFormNo = '" & DOFormNo.Text & "',"
                        SQL = SQL + " OFormSno = '" & DOFormSno.Text & "',"
                        SQL = SQL + " NFormNo = '" & DNFormNo.Text & "',"
                        SQL = SQL + " NFormSno = '" & DNFormSno.Text & "',"
                        SQL = SQL + " Target = '" & DTarget.Text & "',"
                        SQL = SQL + " Content = '" & DContent.Text & "',"
                        SQL = SQL + " Reason = '" & DDReason.Text & "',"

                        If DAttachFile.Visible Then
                            If DAttachFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                                FileName = CStr(wFormSno) & "-" & "Attach" & "-" & UploadDateTime & "-" & Right(DAttachFile.PostedFile.FileName, InStr(StrReverse(DAttachFile.PostedFile.FileName), "\") - 1)
                                Try    '�W�ǹ���
                                    DAttachFile.PostedFile.SaveAs(Path + FileName)
                                Catch ex As Exception
                                End Try
                                SQL = SQL + " AttachFile = '" & FileName & "',"
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

                        If pNextStep = 999 Then     '�u�{������? �����ɱN�s�����ª�
                            If DNFormNo.Text <> "" And DNFormSno.Text <> "" Then
                                Dim DBDataSet1 As New DataSet
                                SQL = "Select * From F_ModManufSheet "
                                SQL = SQL & " Where FormNo =  '" & DNFormNo.Text & "'"
                                SQL = SQL & "   And FormSno =  '" & DNFormSno.Text & "'"
                                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                                DBAdapter1.Fill(DBDataSet1, "F_ModManufSheet")
                                If DBDataSet1.Tables("F_ModManufSheet").Rows.Count > 0 Then
                                    If DOFormNo.Text = "000003" Then
                                        SQL = "Update F_ManufInSheet Set "
                                    Else
                                        SQL = "Update F_ManufOutSheet Set "
                                    End If
                                    SQL = SQL + " No = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("No") & "',"
                                    SQL = SQL + " Date = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Date") & "',"
                                    SQL = SQL + " Division = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Division") & "',"
                                    SQL = SQL + " Person = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Person") & "',"
                                    SQL = SQL + " SliderCode = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderCode") & "',"
                                    SQL = SQL + " SliderGRCode = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderGRCode") & "',"
                                    SQL = SQL + " Spec = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Spec") & "',"
                                    SQL = SQL + " MapNo = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapNo") & "',"
                                    SQL = SQL + " MapFile1 = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile1") & "',"
                                    SQL = SQL + " MapFile2 = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile2") & "',"
                                    SQL = SQL + " Level = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Level") & "',"
                                    SQL = SQL + " Assembler = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Assembler") & "',"
                                    SQL = SQL + " SliderType1 = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderType1") & "',"
                                    SQL = SQL + " SliderType2 = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderType2") & "',"
                                    SQL = SQL + " ManufPlace = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ManufPlace") & "',"
                                    SQL = SQL + " Material = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Material") & "',"
                                    SQL = SQL + " MaterialDetail = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MaterialDetail") & "',"
                                    SQL = SQL + " MaterialOther = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MaterialOther") & "',"
                                    SQL = SQL + " SellVendor = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SellVendor") & "',"
                                    SQL = SQL + " Buyer = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Buyer") & "',"
                                    SQL = SQL + " ConfirmFile = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ConfirmFile") & "',"
                                    SQL = SQL + " AuthorizeFile = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("AuthorizeFile") & "',"
                                    SQL = SQL + " DevReason = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("DevReason") & "',"
                                    SQL = SQL + " Sample = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Sample") & "',"
                                    SQL = SQL + " Price = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Price") & "',"
                                    SQL = SQL + " ArMoldFee = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ArMoldFee") & "',"
                                    SQL = SQL + " PurMold = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("PurMold") & "',"
                                    SQL = SQL + " PullerPrice = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("PullerPrice") & "',"
                                    SQL = SQL + " MoldName = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldName") & "',"
                                    SQL = SQL + " MoldQty = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldQty") & "',"
                                    SQL = SQL + " MoldPoint = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldQty") & "',"
                                    SQL = SQL + " Quality1 = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Quality1") & "',"
                                    SQL = SQL + " Quality2 = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Quality2") & "',"
                                    SQL = SQL + " QAAttachFile = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("QAAttachFile") & "',"
                                    SQL = SQL + " SampleFile = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SampleFile") & "',"
                                    SQL = SQL + " Status = '" & "0" & "',"
                                    SQL = SQL + " StatusDesc = '" & "" & "',"
                                    SQL = SQL + " Contact = '" & "1" & "',"
                                    SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
                                    SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
                                    SQL = SQL + " Where FormNo  =  '" & DOFormNo.Text & "'"
                                    SQL = SQL + "   And FormSno =  '" & DOFormSno.Text & "'"
                                    OleDBCommand1.Connection = OleDbConnection1
                                    OleDBCommand1.CommandText = SQL
                                    OleDbConnection1.Open()
                                    OleDBCommand1.ExecuteNonQuery()
                                    OleDbConnection1.Close()
                                    'Work Sub File Transfer
                                    oPutSubFile = Server.CreateObject("MSubFile.SubFile")
                                    RtnCode = oPutSubFile.Transfer(DNFormNo.Text, DNFormSno.Text, DOFormNo.Text, DOFormSno.Text, Request.Cookies("UserID").Value)
                                End If
                            Else
                                If DOFormNo.Text = "000003" Then
                                    SQL = "Update F_ManufInSheet Set "
                                Else
                                    SQL = "Update F_ManufOutSheet Set "
                                End If
                                SQL = SQL + " Status = '" & "0" & "',"
                                SQL = SQL + " StatusDesc = '" & "" & "',"
                                SQL = SQL + " Contact = '" & "1" & "',"
                                SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
                                SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
                                SQL = SQL + " Where FormNo  =  '" & DOFormNo.Text & "'"
                                SQL = SQL + "   And FormSno =  '" & DOFormSno.Text & "'"
                                OleDBCommand1.Connection = OleDbConnection1
                                OleDBCommand1.CommandText = SQL
                                OleDbConnection1.Open()
                                OleDBCommand1.ExecuteNonQuery()
                                OleDbConnection1.Close()
                            End If  'NFormNo and NFormSno = ""
                        End If  '=999
                    End If  'Insert / Update
                    'End If
                    '�ǰe�l��
                    If pNextStep <> 999 Then
                        For i = 1 To pCount
                            oMail = Server.CreateObject("SendMail.WFMail")
                            oMail.Send(Request.Cookies("UserID").Value, pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "OK")
                            '�e���, �����, �ӽЪ�, ��渹�X, ���y����, �u�{���d���X
                        Next i
                    Else
                        oMail = Server.CreateObject("SendMail.WFMail")
                        oMail.Send(Request.Cookies("UserID").Value, wApplyID, wApplyID, wFormNo, NewFormSno, pNextStep, "OK")
                        '�e���, �����, �ӽЪ�, ��渹�X, ���y����, �u�{���d���X
                    End If

                    If pFlowType = 0 Then
                        Run = True
                        RepeatRun = True
                    End If
                    wStep = pNextStep     '�U�@�u�{���d���X
                Else
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
                Dim URL As String = "MessagePage.aspx?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & _
                                                    "&pUserID=" & Request.Cookies("UserID").Value & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
            End If
        Else
            If ErrCode = 9010 Then Message = "�W���ɮ׮榡���~,�нT�{!"
            If ErrCode = 9020 Then Message = "�W���ɮ�Size�W�L1024KB,�нT�{!"
            If ErrCode = 9030 Then Message = "�W���ɮרS���w,�нT�{!"
            If ErrCode = 9040 Then Message = "����z�Ѭ���L�ɻݶ�g����,�нT�{!"
            Response.Write(YKK.ShowMessage(Message))
        End If      '�W���ɮ�ErrCode=0
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BNG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG.Click
        Dim SQL As String
        Dim ErrCode As Integer = 0   '�ŧi�@��Error�ΨӧP�O�O�_��ƥ��`�A�w�]��False
        Dim RtnCode As Integer = 0
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("ContactOutFilePath"))
        Dim FileName As String
        Dim Message As String = ""
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        '--�y�{��Ƴ]�w---------
        Dim oNextFlow As Object
        Dim pRunNextStep As Integer = 0         '�O�_����p��U�@��(�|ñ)
        '--�l��ǰe---------
        Dim oMail As Object
        '--��{�վ�---------
        Dim oSchedule As Object

        'Check�z��
        If ErrCode = 0 Then
            If DReasonCode.Visible = True Then
                If DReasonCode.SelectedValue = "99" Then
                    If DReasonDesc.Text = "" Then ErrCode = 9310
                End If
            End If
        End If

        '��s������
        If ErrCode = 0 Then
            SQL = "Update T_WaitHandle Set "
            SQL = SQL + " Active = '" & "0" & "',"
            SQL = SQL + " Sts = '" & "2" & "',"
            SQL = SQL + " AEndTime = '" & NowDateTime & "',"
            SQL = SQL + " CompletedTime = '" & NowDateTime & "',"
            If DDelay.Visible = True Then
                SQL = SQL + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQL = SQL + " Reason = '" & DReason.Text & "',"
                SQL = SQL + " ReasonDesc = '" & DReasonDesc.Text & "',"
            End If
            SQL = SQL + " DecideDesc = '" & DDecideDesc.Text & "',"
            SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
            SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
            SQL = SQL + " Where FormNo  =  '" & wFormNo & "'"
            SQL = SQL + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL + "   And Step    =  '" & CStr(wStep) & "'"
            SQL = SQL + "   And EveryOne =  '1' "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDbConnection1.Open()
            OleDBCommand1.ExecuteNonQuery()
            OleDbConnection1.Close()

            '�y�{��Ƶ���
            oNextFlow = Server.CreateObject("FlowEngine.WFFlowEngine")
            oNextFlow.pFormNo = wFormNo   '��渹�X
            oNextFlow.pFormSno = wFormSno                        '���y����
            oNextFlow.pStep = wStep       '�u�{���d���X
            RtnCode = oNextFlow.CheckFlow(Request.Cookies("UserID").Value, pRunNextStep)  '�ظm��, �y�{�����_(�|ñ)
            If RtnCode <> 0 Then ErrCode = 9320
        End If

        '��u�{��{�վ�
        If ErrCode = 0 Then
            Dim wLevel As String = DLevel.SelectedValue     '������
            oSchedule = Server.CreateObject("Schedule.WFSchedule")
            RtnCode = oSchedule.AdjustSchedule(Request.Cookies("UserID").Value, wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel)
        End If

        '��s�����
        If ErrCode = 0 Then
            SQL = "Update F_ContactOutSheet Set "
            SQL = SQL + " Sts = '" & "1" & "',"
            SQL = SQL + " CompletedTime = '" & NowDateTime & "',"
            SQL = SQL + " No = '" & DNo.Text & "',"
            SQL = SQL + " Date = '" & DDate.Text & "',"
            SQL = SQL + " Division = '" & DDivision.SelectedValue & "',"
            SQL = SQL + " Person = '" & DPerson.SelectedValue & "',"
            SQL = SQL + " SliderCode = '" & DSliderCode.Text & "',"
            SQL = SQL + " Level = '" & DLevel.SelectedValue & "',"
            SQL = SQL + " MapNo = '" & DMapNo.Text & "',"
            SQL = SQL + " OFormNo = '" & DOFormNo.Text & "',"
            SQL = SQL + " OFormSno = '" & DOFormSno.Text & "',"
            SQL = SQL + " NFormNo = '" & DNFormNo.Text & "',"
            SQL = SQL + " NFormSno = '" & DNFormSno.Text & "',"
            SQL = SQL + " Target = '" & DTarget.Text & "',"
            SQL = SQL + " Content = '" & DContent.Text & "',"
            SQL = SQL + " Reason = '" & DDReason.Text & "',"

            If DAttachFile.Visible Then
                If DAttachFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    FileName = CStr(wFormSno) & "-" & "Attach" & "-" & UploadDateTime & "-" & Right(DAttachFile.PostedFile.FileName, InStr(StrReverse(DAttachFile.PostedFile.FileName), "\") - 1)
                    Try    '�W�ǹ���
                        DAttachFile.PostedFile.SaveAs(Path + FileName)
                    Catch ex As Exception
                    End Try
                    SQL = SQL + " AttachFile = '" & FileName & "',"
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

            '����B�z
            If DOFormNo.Text = "000003" Then
                SQL = "Update F_ManufInSheet Set "
            Else
                SQL = "Update F_ManufOutSheet Set "
            End If
            SQL = SQL + " Status = '" & "0" & "',"
            SQL = SQL + " StatusDesc = '" & "" & "',"
            SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
            SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
            SQL = SQL + " Where FormNo  =  '" & DOFormNo.Text & "'"
            SQL = SQL + "   And FormSno =  '" & DOFormSno.Text & "'"
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDbConnection1.Open()
            OleDBCommand1.ExecuteNonQuery()
            OleDbConnection1.Close()

            '�ǰe�l��
            oMail = Server.CreateObject("SendMail.WFMail")
            oMail.Send(Request.Cookies("UserID").Value, wApplyID, wApplyID, wFormNo, wFormSno, wStep, "NG")
            '�e���, �����, �ӽЪ�, ��渹�X, ���y����, �u�{���d���X
        End If

        If ErrCode = 0 Then
            Dim URL As String = "MessagePage.aspx?pMSGID=N&pFormNo=" & wFormNo & "&pStep=" & wStep & _
                                                "&pUserID=" & Request.Cookies("UserID").Value & "&pApplyID=" & wApplyID
            Response.Redirect(URL)
        Else
            If ErrCode = 9310 Then Message = "����z�Ѭ���L�ɻݶ�g����,�нT�{!"
            If ErrCode = 9320 Then Message = "�y�{��Ƨ�s���`,�нT�{�γs���t�ΤH��!"
            Response.Write(YKK.ShowMessage(Message))
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
