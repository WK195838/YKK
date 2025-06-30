Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class OPInqCommission_List
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DCpsc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMakeMap As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BEDate As System.Web.UI.WebControls.Button
    Protected WithEvents DEDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSDate As System.Web.UI.WebControls.Button
    Protected WithEvents DSDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DSts As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DProgress As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DMPSelect As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DKeepData As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSUPPILER As System.Web.UI.WebControls.TextBox
    Protected WithEvents LISIP As System.Web.UI.WebControls.HyperLink

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
    Dim pFormNo As String
    Dim SaveTime As String
    Dim NowDateTime As String       '�{�b����ɶ�

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "OPInqCommission_List.aspx"

        SetParameter()          '�]�w�@�ΰѼ�
        If Not Me.IsPostBack Then
            SetSearchItem()
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
        pFormNo = Request.QueryString("pFormNo")    '��渹�X
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
        '
        'DB�s���]�w
        Dim DBDataSet1 As New DataSet
        Dim SQL As String
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        ' SQL = "Select * From M_Referp Where Cat='018' and Dkey = 'ALL'  Order by DKey, Data "
        SQL = " select substring(convert(char(10),dateadd(year,-3,getdate()),111),1,8)+'01' as Data "


        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Referp")
        If DBDataSet1.Tables("M_Referp").Rows.Count > 0 Then
            SaveTime = DBDataSet1.Tables("M_Referp").Rows(0).Item("Data")
        Else
            SaveTime = ""
        End If

        'DB�s������
        OleDbConnection1.Close()


        LISIP.NavigateUrl = "http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & Request.QueryString("pUserID")

    End Sub

    Sub SetSearchItem()
        Dim i As Integer = 0
        Dim wSts As String
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        '���
        SQL = "SELECT FormName, FormNo FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And ( (FormNo >= '000001' And FormNo <= '001000') or FormNo='800001' or FormNo= '007001' ) "
        SQL = SQL + "  And (InqAuthority = '0' "
        SQL = SQL + "       Or (InqAuthority = '1' "
        SQL = SQL + "           And (InqUser like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                Or InqUserName like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "               ) "
        SQL = SQL + "          ) "
        SQL = SQL + "      ) "
        SQL = SQL + "Order by FormNo "

        OleDbConnection1.Open()
        DFormName.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Form")
        DBTable1 = DBDataSet1.Tables("M_Form")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("FormName")
            ListItem1.Value = DBTable1.Rows(i).Item("FormNo")
            If ListItem1.Value = pFormNo Then ListItem1.Selected = True
            DFormName.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()

        '�̿ೡ��
        DBDataSet1.Clear()
        SQL = "Select DivName From M_Users Group by DivName Order by DivName "
        DDivision.Items.Clear()
        DDivision.Items.Add("ALL")
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("DivName")
            ListItem1.Value = DBTable1.Rows(i).Item("DivName")
            DDivision.Items.Add(ListItem1)
        Next

        '���
        DSDate.Text = CStr(DateAdd("d", -180, DateTime.Now.Today))
        DEDate.Text = CStr(DateTime.Now.Today)
        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DEDate');"

        pFormNo = DFormName.SelectedValue

        Search_Item_Attribute()

    End Sub
    '****************************************************
    '  ���ʦs���
    '****************************************************
    Sub DataList()
        Dim i As Integer = 0
        Dim wTableName As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        OleDbConnection1.Open()
        SQL = "SELECT FormNo, FormName, TableName1 FROM M_Form "
        SQL = SQL + " Where Active  =  '1' "
        SQL = SQL + "   And FormNo = '" + DFormName.SelectedValue + "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Flow")
        If DBDataSet1.Tables("Flow").Rows.Count > 0 Then
            wTableName = "F_" + DBDataSet1.Tables("Flow").Rows(0).Item("TableName1")
        End If
        OleDbConnection1.Close()

        If wTableName <> "" Then

            If DMPSelect.SelectedValue = "1" Then   ''�P�_�O�_����Ҩ�O��
                DataGrid1.Columns.Item(0).HeaderText = "�e�UNo"
                DataGrid1.Columns.Item(1).HeaderText = "�}�o���A"
                DataGrid1.Columns.Item(2).HeaderText = "�[�u����"
                DataGrid1.Columns.Item(3).HeaderText = "�~�`�t��"

                SQL = "SELECT "

                SQL = SQL + "Case " + wTableName + ".No When '' Then '���s��' Else " + wTableName + ".No End As Field1, "
                SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '�}�o��' When '1' Then '�}�o����(OK)' When '2' Then '�}�o����(NG)' Else '�}�o����' End As Field2, "

                SQL = SQL + wTableName + ".Material as Field3, "
                SQL = SQL + wTableName + ".Suppiler as Field4, "
            Else
                DataGrid1.Columns.Item(0).HeaderText = "�e�UNo"
                DataGrid1.Columns.Item(1).HeaderText = "�}�o���A"
                DataGrid1.Columns.Item(2).HeaderText = "�̿��"
                DataGrid1.Columns.Item(3).HeaderText = "�̿��"
                SQL = "SELECT "
                'SQL = SQL + wTableName + ".No as Field1, "

                SQL = SQL + "Case " + wTableName + ".No When '' Then '���s��' Else " + wTableName + ".No End As Field1, "
                SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '�}�o��' When '1' Then '�}�o����(OK)' When '2' Then '�}�o����(NG)' Else '�}�o����' End As Field2, "

                SQL = SQL + wTableName + ".Division + '--' + "
                SQL = SQL + wTableName + ".Person as Field3, "


                SQL = SQL + "V_WaitHandle_01.ApplyTime as Field4, "
                'SQL = SQL + wTableName + ".CompletedTime as Field4, "
            End If

            If pFormNo = "000001" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "�]�p��"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".MakeMap as Field8, "
            End If

            If pFormNo = "000002" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "��ϸ�"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "�]�p��"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".OriMapNo as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".MakeMap as Field8, "
            End If

            If pFormNo = "000003" Or pFormNo = "000004" Or pFormNo = "000005" Or _
               pFormNo = "000007" Or pFormNo = "000008" Or pFormNo = "000009" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                If pFormNo = "000004" Or pFormNo = "000005" Or _
                   pFormNo = "000008" Or pFormNo = "000009" Then
                    If pFormNo = "000004" Or pFormNo = "000005" Then
                        SQL = SQL + "F_ManufInModSheet.Buyer   as Field5, "
                        SQL = SQL + "F_ManufInModSheet.Spec    as Field6, "
                        SQL = SQL + "F_ManufInModSheet.MapNo   as Field7, "
                        SQL = SQL + "F_ManufInModSheet.SliderCode as Field8, "
                    Else
                        SQL = SQL + "F_ManufOutModSheet.Buyer   as Field5, "
                        SQL = SQL + "F_ManufOutModSheet.Spec    as Field6, "
                        SQL = SQL + "F_ManufOutModSheet.MapNo   as Field7, "
                        SQL = SQL + "F_ManufOutModSheet.SliderCode as Field8, "
                    End If
                Else
                    SQL = SQL + wTableName + ".Buyer   as Field5, "
                    SQL = SQL + wTableName + ".Spec    as Field6, "
                    SQL = SQL + wTableName + ".MapNo   as Field7, "
                    SQL = SQL + wTableName + ".SliderCode as Field8, "
                End If
            End If

            If pFormNo = "000010" Then
                DataGrid1.Columns.Item(4).HeaderText = "G-Slider Code"
                DataGrid1.Columns.Item(5).HeaderText = "�s�e�U"
                DataGrid1.Columns.Item(6).HeaderText = ""
                DataGrid1.Columns.Item(7).HeaderText = ""

                SQL = SQL + wTableName + ".SliderGRCode as Field5, "
                SQL = SQL + wTableName + ".OFormSno as Field6, "
                SQL = SQL + "'  ' as Field7, "
                SQL = SQL + "'  ' as Field8, "
            End If
            If pFormNo = "000011" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + "'  ' as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If
            If pFormNo = "000012" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + "'  ' as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If
            If pFormNo = "000013" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + "'  ' as Field5, "
                SQL = SQL + "'  ' as Field6, "
                SQL = SQL + "'  ' as Field7, "
                SQL = SQL + "'  ' as Field8, "
            End If
            If pFormNo = "000014" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "Code"
                DataGrid1.Columns.Item(7).HeaderText = "OR-No"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".Code    as Field7, "
                SQL = SQL + wTableName + ".ORNO    as Field8, "
            End If
            If pFormNo = "000015" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "Code"
                DataGrid1.Columns.Item(7).HeaderText = "�l�[�z��"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".Code    as Field7, "
                SQL = SQL + wTableName + ".AppendReason    as Field8, "
            End If
            If pFormNo = "000016" Then
                DataGrid1.Columns.Item(4).HeaderText = "��f����"
                DataGrid1.Columns.Item(5).HeaderText = "�ʤJ��/�[��"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".ColorItem  as Field5, "
                SQL = SQL + wTableName + ".Pullerprice" + " +'/'+ " + wTableName + ".price as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "

            End If
            If pFormNo = "800001" Then
                DataGrid1.Columns.Item(4).HeaderText = "�e�U��"
                DataGrid1.Columns.Item(5).HeaderText = "���A"
                DataGrid1.Columns.Item(6).HeaderText = "�ק�z�����O"
                DataGrid1.Columns.Item(7).HeaderText = "�ק�z��"

                SQL = SQL + wTableName + ".Sheet   as Field5, "
                SQL = SQL + wTableName + ".Status  as Field6, "
                SQL = SQL + wTableName + ".MReasonType as Field7, "
                SQL = SQL + wTableName + ".MReason as Field8, "
            End If

            If pFormNo = "007001" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(6).HeaderText = "���Y�~�W"
                DataGrid1.Columns.Item(7).HeaderText = "�j�f��"


                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".MapNo  as Field6, "
                SQL = SQL + wTableName + ".SliderCode as Field7, "
                SQL = SQL + wTableName + ".YKKGroup as Field8, "
            End If

            If DMPSelect.SelectedValue = "1" Then   ''�P�_�O�_����Ҩ�O��
                Dim bc1 As New BoundColumn
                Dim bc2 As New BoundColumn

                bc1.DataField = "Field9"
                bc1.HeaderText = "�����Ҩ�O"
                bc2.DataField = "Field10"
                bc2.HeaderText = "�Ҩ��ʤJ��"

                DataGrid1.Columns.AddAt(8, bc1) ''�s�WDataGrid���
                DataGrid1.Columns.AddAt(9, bc2)
                SetFieldWidth()

                SQL = SQL + wTableName + ".ArMoldFee as Field9, "
                SQL = SQL + wTableName + ".PurMold as Field10, "
                Dim bc3 As New BoundColumn
                bc3.DataField = "Field11"
                bc3.HeaderText = "�e�U�t��"
                DataGrid1.Columns.AddAt(10, bc3) ''�s�WDataGrid���
                SQL = SQL + wTableName + ".SellVendor as Field11, "
                If pFormNo <> "000011" And pFormNo <> "000012" And pFormNo <> "000015" Then
                    Dim bc4 As New BoundColumn
                    bc4.DataField = "Field12"
                    bc4.HeaderText = "�~�`�t��"
                    DataGrid1.Columns.AddAt(11, bc4) ''�s�WDataGrid���
                    SQL = SQL + wTableName + ".SUPPILER as Field12, "
                End If


            End If
            '2012/5/16 Add-Start by Joy
            If pFormNo = "000003" Or pFormNo = "000007" Then
                SQL = SQL + wTableName + ".MMSSts as MSts, "
            End If

            '20180821 jessica 
            If pFormNo = "000001" Or pFormNo = "000002" Or pFormNo = "000003" Or pFormNo = "000007" Or pFormNo = "000011" Or pFormNo = "000012" Or pFormNo = "000014" Or pFormNo = "000015" Then
                If DMPSelect.SelectedValue <> "1" Then
                    Dim bc3 As New BoundColumn
                    bc3.DataField = "Field9"
                    bc3.HeaderText = "�e�U�t��"
                    DataGrid1.Columns.AddAt(8, bc3) ''�s�WDataGrid���
                    SQL = SQL + wTableName + ".SellVendor as Field9, "
                    If pFormNo <> "000011" And pFormNo <> "000012" And pFormNo <> "000015" Then
                        Dim bc4 As New BoundColumn
                        bc4.DataField = "Field10"
                        bc4.HeaderText = "�~�`�t��"
                        DataGrid1.Columns.AddAt(9, bc4) ''�s�WDataGrid���
                        SQL = SQL + wTableName + ".SUPPILER as Field10, "
                    End If

                End If
            End If
            'Add-End






            'Add-End
            SQL = SQL + " '....' as WorkFlow, ViewURL, "
            SQL = SQL + "'BefOPList.aspx?' + "
            SQL = SQL + "'pFormNo='   + V_WaitHandle_01.FormNo + "
            SQL = SQL + "'&pFormSno=' + str(V_WaitHandle_01.FormSno,Len(V_WaitHandle_01.FormSno)) + "
            SQL = SQL + "'&pStep='    + str(V_WaitHandle_01.Step,Len(V_WaitHandle_01.Step)) + "
            SQL = SQL + "'&pSeqNo='   + str(V_WaitHandle_01.SeqNo,Len(V_WaitHandle_01.SeqNo)) + "
            SQL = SQL + "'&pApplyID=' + V_WaitHandle_01.ApplyID "
            SQL = SQL + " As OPURL "

            SQL = SQL + "FROM " + wTableName + " "
            SQL = SQL + "Left Outer Join V_WaitHandle_01 ON " + wTableName + ".FormNo=V_WaitHandle_01.FormNo "
            SQL = SQL + "                               And " + wTableName + ".FormSno=V_WaitHandle_01.FormSno "
            '���s�u�{�~�ȳs��
            If pFormNo = "000004" Or pFormNo = "000005" Then
                SQL = SQL + "Left Outer Join F_ManufInModSheet ON " + wTableName + ".NFormNo=F_ManufInModSheet.FormNo "
                SQL = SQL + "                                 And " + wTableName + ".NFormSno=F_ManufInModSheet.FormSno "
            End If
            '�~�s, �u�{�~�ȳs��
            If pFormNo = "000008" Or pFormNo = "000009" Then
                SQL = SQL + "Left Outer Join F_ManufOutModSheet ON " + wTableName + ".NFormNo=F_ManufOutModSheet.FormNo "
                SQL = SQL + "                                  And " + wTableName + ".NFormSno=F_ManufOutModSheet.FormSno "
            End If
            '------------------------------------
            SQL = SQL + "Where V_WaitHandle_01.Step  < '10' "
            '���
            If DFormName.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".FormNo = '" + DFormName.SelectedValue + "'"
            End If

            '�}�o���A
            If DProgress.SelectedValue <> "ALL" Then
                If DProgress.SelectedValue = "1" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '0'  "
                End If
                If DProgress.SelectedValue = "2" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '1'  "
                End If
            End If

            '�}�o�������A
            If DSts.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".Sts = '" + DSts.SelectedValue + "'"
            End If

            '����
            If DDivision.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".Division = '" + DDivision.SelectedValue + "'"
            End If
            '�̿�H
            If DPerson.Text <> "�̿��" And DPerson.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".Person Like '%" + DPerson.Text + "%'"
            End If
            'Buyer
            If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.Buyer Like '%" + DBuyer.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.Buyer Like '%" + DBuyer.Text + "%'"
                    Else
                        SQL = SQL + " And " + wTableName + ".Buyer Like '%" + DBuyer.Text + "%'"
                    End If
                End If
            End If

            '20180821 jessica �W�[�e�U�t��
            'sellvendor
            If DSellVendor.Text <> "�e�U�t��" And DSellVendor.Text <> "" Then

                SQL = SQL + " And " + wTableName + ".SellVendor Like '%" + DSellVendor.Text + "%'"

            End If

            '202401191 jessica �W�[�~�` 
            'sellvendor
            If DSUPPILER.Text <> "�~�`�t��" And DSUPPILER.Text <> "" Then

                SQL = SQL + " And " + wTableName + ".SUPPILER Like '%" + DSUPPILER.Text + "%'"

            End If


            '�]�p��
            If DMakeMap.Text <> "�]�p��" And DMakeMap.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".MakeMap Like '%" + DMakeMap.Text + "%'"
            End If
            'Slider Code
            If DSliderCode.Text <> "Slider Code" And DSliderCode.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.SliderCode Like '%" + DSliderCode.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.SliderCode Like '%" + DSliderCode.Text + "%'"
                    Else
                        If pFormNo = "000010" Then
                            SQL = SQL + " And " + wTableName + ".SliderGRCode Like '%" + DSliderCode.Text + "%'"
                        Else
                            SQL = SQL + " And " + wTableName + ".SliderCode Like '%" + DSliderCode.Text + "%'"
                        End If
                    End If
                End If
            End If
            'Spec
            If DSpec.Text <> "Size-Chain-����" And DSpec.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.Spec Like '%" + DSpec.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.Spec Like '%" + DSpec.Text + "%'"
                    Else
                        SQL = SQL + " And " + wTableName + ".Spec Like '%" + DSpec.Text + "%'"
                    End If
                End If
            End If
            'MapNo
            If DMapNo.Text <> "�ϸ�" And DMapNo.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.MapNo Like '%" + DMapNo.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.MapNo Like '%" + DMapNo.Text + "%'"
                    Else
                        SQL = SQL + " And " + wTableName + ".MapNo Like '%" + DMapNo.Text + "%'"
                    End If
                End If
            End If
            'Cpsc
            If DCpsc.Text <> "CPSC" And DCpsc.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.Cpsc Like '%" + DCpsc.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.Cpsc Like '%" + DCpsc.Text + "%'"
                    Else
                        SQL = SQL + " And " + wTableName + ".Cpsc Like '%" + DCpsc.Text + "%'"
                    End If
                End If
            End If
            'No
            If DNo.Text <> "�e�U��No." And DNo.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".No Like '%" + DNo.Text + "%'"
            End If
            'Code
            If DCode.Text <> "Code" And DCode.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".Code Like '%" + DCode.Text + "%'"
            End If
            ''CompletedTime
            If DMPSelect.SelectedValue = "1" Then   ''�P�_�O�_����Ҩ�O��
                SQL = SQL + " And " + wTableName + ".CompletedTime between '" + DSDate.Text + " 00:00:00' and '" + DEDate.Text + " 23:59:59'"
            Else
                SQL = SQL + " And " + wTableName + ".CompletedTime between '" + DSDate.Text + " 00:00:00' and '" + DEDate.Text + " 23:59:59'"
            End If
            'Sort
            SQL = SQL + " Order by " + wTableName + ".No desc "

            OleDbConnection1.Open()
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "WaitHandle")
            '
            '2012/5/16 Add-Start by Joy
            If pFormNo = "000003" Or pFormNo = "000007" Then
                DBTable1 = DBDataSet2.Tables("WaitHandle")
                If DBTable1.Rows.Count > 0 Then
                    For i = 0 To DBTable1.Rows.Count - 1
                        If DBTable1.Rows(i).Item("MSts") = 1 Then
                            DBTable1.Rows(i).Item("Field1") = DBTable1.Rows(i).Item("Field1") + "-�Ҽo"
                        End If
                    Next
                End If
            End If
            'Add-End
            '
            DataGrid1.DataSource = DBDataSet2
            DataGrid1.DataBind()
            OleDbConnection1.Close()
        End If

    End Sub
    '****************************************************
    '  �w�ʦs���
    '****************************************************
    Sub KeepDataList()
        Dim i As Integer = 0
        Dim wTableName As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        OleDbConnection1.Open()
        SQL = "SELECT FormNo, FormName, TableName1 FROM M_Form "
        SQL = SQL + " Where Active  =  '1' "
        SQL = SQL + "   And FormNo = '" + DFormName.SelectedValue + "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Flow")
        If DBDataSet1.Tables("Flow").Rows.Count > 0 Then
            wTableName = "F_" + DBDataSet1.Tables("Flow").Rows(0).Item("TableName1")
        End If
        OleDbConnection1.Close()

        If wTableName <> "" Then

            If DMPSelect.SelectedValue = "1" Then   ''�P�_�O�_����Ҩ�O��
                DataGrid1.Columns.Item(0).HeaderText = "�e�UNo"
                DataGrid1.Columns.Item(1).HeaderText = "�}�o���A"
                DataGrid1.Columns.Item(2).HeaderText = "�[�u����"
                DataGrid1.Columns.Item(3).HeaderText = "�~�`�t��"

                SQL = "SELECT "

                SQL = SQL + "Case " + wTableName + ".No When '' Then '���s��' Else " + wTableName + ".No End As Field1, "
                SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '�}�o��' When '1' Then '�}�o����(OK)' When '2' Then '�}�o����(NG)' Else '�}�o����' End As Field2, "

                SQL = SQL + wTableName + ".Material as Field3, "
                SQL = SQL + wTableName + ".Suppiler as Field4, "
            Else
                DataGrid1.Columns.Item(0).HeaderText = "�e�UNo"
                DataGrid1.Columns.Item(1).HeaderText = "�}�o���A"
                DataGrid1.Columns.Item(2).HeaderText = "�̿��"
                DataGrid1.Columns.Item(3).HeaderText = "�̿��"
                SQL = "SELECT "
                'SQL = SQL + wTableName + ".No as Field1, "

                SQL = SQL + "Case " + wTableName + ".No When '' Then '���s��' Else " + wTableName + ".No End As Field1, "
                SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '�}�o��' When '1' Then '�}�o����(OK)' When '2' Then '�}�o����(NG)' Else '�}�o����' End As Field2, "

                SQL = SQL + wTableName + ".Division + '--' + "
                SQL = SQL + wTableName + ".Person as Field3, "
                SQL = SQL + "V_WaitHandle_OLD_01.ApplyTime as Field4, "
                'SQL = SQL + wTableName + ".CompletedTime as Field4, "
            End If

            If pFormNo = "000001" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "�]�p��"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".MakeMap as Field8, "
            End If

            If pFormNo = "000002" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "��ϸ�"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "�]�p��"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".OriMapNo as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".MakeMap as Field8, "
            End If

            If pFormNo = "000003" Or pFormNo = "000004" Or pFormNo = "000005" Or _
               pFormNo = "000007" Or pFormNo = "000008" Or pFormNo = "000009" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                If pFormNo = "000004" Or pFormNo = "000005" Or _
                   pFormNo = "000008" Or pFormNo = "000009" Then
                    If pFormNo = "000004" Or pFormNo = "000005" Then
                        SQL = SQL + "F_ManufInModSheet.Buyer   as Field5, "
                        SQL = SQL + "F_ManufInModSheet.Spec    as Field6, "
                        SQL = SQL + "F_ManufInModSheet.MapNo   as Field7, "
                        SQL = SQL + "F_ManufInModSheet.SliderCode as Field8, "
                    Else
                        SQL = SQL + "F_ManufOutModSheet.Buyer   as Field5, "
                        SQL = SQL + "F_ManufOutModSheet.Spec    as Field6, "
                        SQL = SQL + "F_ManufOutModSheet.MapNo   as Field7, "
                        SQL = SQL + "F_ManufOutModSheet.SliderCode as Field8, "
                    End If
                Else
                    SQL = SQL + wTableName + ".Buyer   as Field5, "
                    SQL = SQL + wTableName + ".Spec    as Field6, "
                    SQL = SQL + wTableName + ".MapNo   as Field7, "
                    SQL = SQL + wTableName + ".SliderCode as Field8, "
                End If
            End If


            If pFormNo = "000010" Then
                DataGrid1.Columns.Item(4).HeaderText = "G-Slider Code"
                DataGrid1.Columns.Item(5).HeaderText = "�s�e�U"
                DataGrid1.Columns.Item(6).HeaderText = ""
                DataGrid1.Columns.Item(7).HeaderText = ""

                SQL = SQL + wTableName + ".SliderGRCode as Field5, "
                SQL = SQL + wTableName + ".OFormSno as Field6, "
                SQL = SQL + "'  ' as Field7, "
                SQL = SQL + "'  ' as Field8, "
            End If
            If pFormNo = "000011" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + "'  ' as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If
            If pFormNo = "000012" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + "'  ' as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If
            If pFormNo = "000013" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + "'  ' as Field5, "
                SQL = SQL + "'  ' as Field6, "
                SQL = SQL + "'  ' as Field7, "
                SQL = SQL + "'  ' as Field8, "
            End If
            If pFormNo = "000014" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "Code"
                DataGrid1.Columns.Item(7).HeaderText = "OR-No"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".Code    as Field7, "
                SQL = SQL + wTableName + ".ORNO    as Field8, "
            End If
            If pFormNo = "000015" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "Code"
                DataGrid1.Columns.Item(7).HeaderText = "�l�[�z��"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".Code    as Field7, "
                SQL = SQL + wTableName + ".AppendReason    as Field8, "
            End If
            If pFormNo = "000016" Then
                DataGrid1.Columns.Item(4).HeaderText = "��f����"
                DataGrid1.Columns.Item(5).HeaderText = "�ʤJ��/�[��"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".ColorItem  as Field5, "
                SQL = SQL + wTableName + ".Pullerprice" + " +'/'+ " + wTableName + ".price as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If
            If pFormNo = "800001" Then
                DataGrid1.Columns.Item(4).HeaderText = "�e�U��"
                DataGrid1.Columns.Item(5).HeaderText = "���A"
                DataGrid1.Columns.Item(6).HeaderText = "�ק�z�����O"
                DataGrid1.Columns.Item(7).HeaderText = "�ק�z��"

                SQL = SQL + wTableName + ".Sheet   as Field5, "
                SQL = SQL + wTableName + ".Status  as Field6, "
                SQL = SQL + wTableName + ".MReasonType as Field7, "
                SQL = SQL + wTableName + ".MReason as Field8, "
            End If

            If pFormNo = "007001" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(6).HeaderText = "���Y�~�W"
                DataGrid1.Columns.Item(7).HeaderText = "�j�f��"


                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".MapNo  as Field6, "
                SQL = SQL + wTableName + ".SliderCode as Field7, "
                SQL = SQL + wTableName + ".YKKGroup as Field8, "
            End If


            If DMPSelect.SelectedValue = "1" Then   ''�P�_�O�_����Ҩ�O��
                Dim bc1 As New BoundColumn
                Dim bc2 As New BoundColumn

                bc1.DataField = "Field9"
                bc1.HeaderText = "�����Ҩ�O"
                bc2.DataField = "Field10"
                bc2.HeaderText = "�Ҩ��ʤJ��"

                DataGrid1.Columns.AddAt(8, bc1) ''�s�WDataGrid���
                DataGrid1.Columns.AddAt(9, bc2)
                SetFieldWidth()

                SQL = SQL + wTableName + ".ArMoldFee as Field9, "
                SQL = SQL + wTableName + ".PurMold as Field10, "
                Dim bc3 As New BoundColumn
                bc3.DataField = "Field11"
                bc3.HeaderText = "�e�U�t��"
                DataGrid1.Columns.AddAt(10, bc3) ''�s�WDataGrid���
                SQL = SQL + wTableName + ".SellVendor as Field11, "

                If pFormNo <> "000011" And pFormNo <> "000012" And pFormNo <> "000015" Then
                    Dim bc4 As New BoundColumn
                    bc4.DataField = "Field12"
                    bc4.HeaderText = "�~�`�t��"
                    DataGrid1.Columns.AddAt(11, bc4) ''�s�WDataGrid���
                    SQL = SQL + wTableName + ".Suppiler as Field12, "
                End If

            End If
            '2012/5/16 Add-Start by Joy
            If pFormNo = "000003" Or pFormNo = "000007" Then
                SQL = SQL + wTableName + ".MMSSts as MSts, "
            End If

            '20180821 jessica 
            If pFormNo = "000001" Or pFormNo = "000002" Or pFormNo = "000003" Or pFormNo = "000007" Or pFormNo = "000011" Or pFormNo = "000012" Or pFormNo = "000014" Or pFormNo = "000015" Then
                If DMPSelect.SelectedValue <> "1" Then
                    Dim bc3 As New BoundColumn
                    bc3.DataField = "Field9"
                    bc3.HeaderText = "�e�U�t��"
                    DataGrid1.Columns.AddAt(8, bc3) ''�s�WDataGrid���
                    SQL = SQL + wTableName + ".SellVendor as Field9, "

                    If pFormNo <> "000011" And pFormNo <> "000012" And pFormNo <> "000015" Then
                        Dim bc4 As New BoundColumn
                        bc4.DataField = "Field10"
                        bc4.HeaderText = "�~�`�t��"
                        DataGrid1.Columns.AddAt(9, bc4) ''�s�WDataGrid���
                        SQL = SQL + wTableName + ".SUPPILER as Field10, "

                    End If


                End If



            End If

            'Add-End

            'Add-End
            SQL = SQL + " '....' as WorkFlow, ViewURL, "
            SQL = SQL + "'BefOPList.aspx?' + "
            SQL = SQL + "'pFormNo='   + V_WaitHandle_OLD_01.FormNo + "
            SQL = SQL + "'&pFormSno=' + str(V_WaitHandle_OLD_01.FormSno,Len(V_WaitHandle_OLD_01.FormSno)) + "
            SQL = SQL + "'&pStep='    + str(V_WaitHandle_OLD_01.Step,Len(V_WaitHandle_OLD_01.Step)) + "
            SQL = SQL + "'&pSeqNo='   + str(V_WaitHandle_OLD_01.SeqNo,Len(V_WaitHandle_OLD_01.SeqNo)) + "
            SQL = SQL + "'&pKeepdata='   + rtrim(ltrim(str(1))) + "
            SQL = SQL + "'&pApplyID=' + V_WaitHandle_OLD_01.ApplyID "
            SQL = SQL + " As OPURL "

            SQL = SQL + "FROM " + wTableName + " "
            SQL = SQL + "Left Outer Join V_WaitHandle_OLD_01 ON " + wTableName + ".FormNo=V_WaitHandle_OLD_01.FormNo "
            SQL = SQL + "                               And " + wTableName + ".FormSno=V_WaitHandle_OLD_01.FormSno "
            '���s�u�{�~�ȳs��
            If pFormNo = "000004" Or pFormNo = "000005" Then
                SQL = SQL + "Left Outer Join F_ManufInModSheet ON " + wTableName + ".NFormNo=F_ManufInModSheet.FormNo "
                SQL = SQL + "                                 And " + wTableName + ".NFormSno=F_ManufInModSheet.FormSno "
            End If
            '�~�s, �u�{�~�ȳs��
            If pFormNo = "000008" Or pFormNo = "000009" Then
                SQL = SQL + "Left Outer Join F_ManufOutModSheet ON " + wTableName + ".NFormNo=F_ManufOutModSheet.FormNo "
                SQL = SQL + "                                  And " + wTableName + ".NFormSno=F_ManufOutModSheet.FormSno "
            End If
            '------------------------------------
            SQL = SQL + "Where V_WaitHandle_OLD_01.Step  < '10' "
            '���
            If DFormName.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".FormNo = '" + DFormName.SelectedValue + "'"
            End If

            '�}�o���A
            If DProgress.SelectedValue <> "ALL" Then
                If DProgress.SelectedValue = "1" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '0'  "
                End If
                If DProgress.SelectedValue = "2" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '1'  "
                End If
            End If

            '�}�o�������A
            If DSts.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".Sts = '" + DSts.SelectedValue + "'"
            End If

            '����
            If DDivision.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".Division = '" + DDivision.SelectedValue + "'"
            End If
            '�̿�H
            If DPerson.Text <> "�̿��" And DPerson.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".Person Like '%" + DPerson.Text + "%'"
            End If
            'Buyer
            If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.Buyer Like '%" + DBuyer.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.Buyer Like '%" + DBuyer.Text + "%'"
                    Else
                        SQL = SQL + " And " + wTableName + ".Buyer Like '%" + DBuyer.Text + "%'"
                    End If
                End If
            End If


            '20180821 jessica �W�[�e�U�t��
            'sellvendor
            If DSellVendor.Text <> "�e�U�t��" And DSellVendor.Text <> "" Then

                SQL = SQL + " And " + wTableName + ".SellVendor Like '%" + DSellVendor.Text + "%'"

            End If

            '202401191 jessica �W�[�~�` 
            'sellvendor
            If DSUPPILER.Text <> "�~�`�t��" And DSUPPILER.Text <> "" Then

                SQL = SQL + " And " + wTableName + ".SUPPILER Like '%" + DSUPPILER.Text + "%'"

            End If



            '�]�p��
            If DMakeMap.Text <> "�]�p��" And DMakeMap.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".MakeMap Like '%" + DMakeMap.Text + "%'"
            End If
            'Slider Code
            If DSliderCode.Text <> "Slider Code" And DSliderCode.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.SliderCode Like '%" + DSliderCode.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.SliderCode Like '%" + DSliderCode.Text + "%'"
                    Else
                        If pFormNo = "000010" Then
                            SQL = SQL + " And " + wTableName + ".SliderGRCode Like '%" + DSliderCode.Text + "%'"
                        Else
                            SQL = SQL + " And " + wTableName + ".SliderCode Like '%" + DSliderCode.Text + "%'"
                        End If
                    End If
                End If
            End If
            'Spec
            If DSpec.Text <> "Size-Chain-����" And DSpec.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.Spec Like '%" + DSpec.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.Spec Like '%" + DSpec.Text + "%'"
                    Else
                        SQL = SQL + " And " + wTableName + ".Spec Like '%" + DSpec.Text + "%'"
                    End If
                End If
            End If
            'MapNo
            If DMapNo.Text <> "�ϸ�" And DMapNo.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.MapNo Like '%" + DMapNo.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.MapNo Like '%" + DMapNo.Text + "%'"
                    Else
                        SQL = SQL + " And " + wTableName + ".MapNo Like '%" + DMapNo.Text + "%'"
                    End If
                End If
            End If
            'Cpsc
            If DCpsc.Text <> "CPSC" And DCpsc.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.Cpsc Like '%" + DCpsc.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.Cpsc Like '%" + DCpsc.Text + "%'"
                    Else
                        SQL = SQL + " And " + wTableName + ".Cpsc Like '%" + DCpsc.Text + "%'"
                    End If
                End If
            End If
            'No
            If DNo.Text <> "�e�U��No." And DNo.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".No Like '%" + DNo.Text + "%'"
            End If
            'Code
            If DCode.Text <> "Code" And DCode.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".Code Like '%" + DCode.Text + "%'"
            End If
            ''CompletedTime
            If DMPSelect.SelectedValue = "1" Then   ''�P�_�O�_����Ҩ�O��
                SQL = SQL + " And " + wTableName + ".CompletedTime between '" + DSDate.Text + " 00:00:00' and '" + DEDate.Text + " 23:59:59'"
            Else
                SQL = SQL + " And " + wTableName + ".CompletedTime between '" + DSDate.Text + " 00:00:00' and '" + DEDate.Text + " 23:59:59'"
            End If
            'Sort
            SQL = SQL + " Order by " + wTableName + ".No desc "

            OleDbConnection1.Open()
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "WaitHandle")
            '
            '2012/5/16 Add-Start by Joy
            If pFormNo = "000003" Or pFormNo = "000007" Then
                DBTable1 = DBDataSet2.Tables("WaitHandle")
                If DBTable1.Rows.Count > 0 Then
                    For i = 0 To DBTable1.Rows.Count - 1
                        If DBTable1.Rows(i).Item("MSts") = 1 Then
                            DBTable1.Rows(i).Item("Field1") = DBTable1.Rows(i).Item("Field1") + "-�Ҽo"
                        End If
                    Next
                End If
            End If
            'Add-End
            '
            DataGrid1.DataSource = DBDataSet2
            DataGrid1.DataBind()
            OleDbConnection1.Close()
        End If

    End Sub
    '****************************************************
    '  �����
    '****************************************************
    Sub FormDataList()
        Dim i As Integer = 0
        Dim wTableName As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        OleDbConnection1.Open()
        SQL = "SELECT FormNo, FormName, TableName1 FROM M_Form "
        SQL = SQL + " Where Active  =  '1' "
        SQL = SQL + "   And FormNo = '" + DFormName.SelectedValue + "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Flow")
        If DBDataSet1.Tables("Flow").Rows.Count > 0 Then
            wTableName = "F_" + DBDataSet1.Tables("Flow").Rows(0).Item("TableName1")
        End If
        OleDbConnection1.Close()

        If wTableName <> "" Then

            If DMPSelect.SelectedValue = "1" Then   ''�P�_�O�_����Ҩ�O��
                DataGrid1.Columns.Item(0).HeaderText = "�e�UNo"
                DataGrid1.Columns.Item(1).HeaderText = "�}�o���A"
                DataGrid1.Columns.Item(2).HeaderText = "�[�u����"
                DataGrid1.Columns.Item(3).HeaderText = "�~�`�t��"

                SQL = "SELECT "

                SQL = SQL + "Case " + wTableName + ".No When '' Then '���s��' Else " + wTableName + ".No End As Field1, "
                SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '�}�o��' When '1' Then '�}�o����(OK)' When '2' Then '�}�o����(NG)' Else '�}�o����' End As Field2, "

                SQL = SQL + wTableName + ".Material as Field3, "
                SQL = SQL + wTableName + ".Suppiler as Field4, "
            Else
                DataGrid1.Columns.Item(0).HeaderText = "�e�UNo"
                DataGrid1.Columns.Item(1).HeaderText = "�}�o���A"
                DataGrid1.Columns.Item(2).HeaderText = "�̿��"
                DataGrid1.Columns.Item(3).HeaderText = "�̿��"
                SQL = "SELECT "
                'SQL = SQL + wTableName + ".No as Field1, "

                SQL = SQL + "Case " + wTableName + ".No When '' Then '���s��' Else " + wTableName + ".No End As Field1, "
                SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '�}�o��' When '1' Then '�}�o����(OK)' When '2' Then '�}�o����(NG)' Else '�}�o����' End As Field2, "

                SQL = SQL + wTableName + ".Division + '--' + "
                SQL = SQL + wTableName + ".Person as Field3, "
                SQL = SQL + wTableName + ".CreateTime as Field4, "
            End If

            If pFormNo = "000001" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "�]�p��"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".MakeMap as Field8, "
            End If



            If pFormNo = "000002" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "��ϸ�"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "�]�p��"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".OriMapNo as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".MakeMap as Field8, "
            End If

            If pFormNo = "000003" Or pFormNo = "000004" Or pFormNo = "000005" Or _
               pFormNo = "000007" Or pFormNo = "000008" Or pFormNo = "000009" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                If pFormNo = "000004" Or pFormNo = "000005" Or _
                   pFormNo = "000008" Or pFormNo = "000009" Then
                    If pFormNo = "000004" Or pFormNo = "000005" Then
                        SQL = SQL + "F_ManufInModSheet.Buyer   as Field5, "
                        SQL = SQL + "F_ManufInModSheet.Spec    as Field6, "
                        SQL = SQL + "F_ManufInModSheet.MapNo   as Field7, "
                        SQL = SQL + "F_ManufInModSheet.SliderCode as Field8, "
                    Else
                        SQL = SQL + "F_ManufOutModSheet.Buyer   as Field5, "
                        SQL = SQL + "F_ManufOutModSheet.Spec    as Field6, "
                        SQL = SQL + "F_ManufOutModSheet.MapNo   as Field7, "
                        SQL = SQL + "F_ManufOutModSheet.SliderCode as Field8, "
                    End If
                Else
                    SQL = SQL + wTableName + ".Buyer   as Field5, "
                    SQL = SQL + wTableName + ".Spec    as Field6, "
                    SQL = SQL + wTableName + ".MapNo   as Field7, "
                    SQL = SQL + wTableName + ".SliderCode as Field8, "
                End If
            End If

            If pFormNo = "000010" Then
                DataGrid1.Columns.Item(4).HeaderText = "G-Slider Code"
                DataGrid1.Columns.Item(5).HeaderText = "�s�e�U"
                DataGrid1.Columns.Item(6).HeaderText = ""
                DataGrid1.Columns.Item(7).HeaderText = ""

                SQL = SQL + wTableName + ".SliderGRCode as Field5, "
                SQL = SQL + wTableName + ".OFormSno as Field6, "
                SQL = SQL + "'  ' as Field7, "
                SQL = SQL + "'  ' as Field8, "
            End If
            If pFormNo = "000011" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + "'  ' as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If
            If pFormNo = "000012" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "CPSC"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".CPSC   as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If
            If pFormNo = "000013" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + "'  ' as Field5, "
                SQL = SQL + "'  ' as Field6, "
                SQL = SQL + "'  ' as Field7, "
                SQL = SQL + "'  ' as Field8, "
            End If
            If pFormNo = "000014" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "Code"
                DataGrid1.Columns.Item(7).HeaderText = "OR-No"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".Code    as Field7, "
                SQL = SQL + wTableName + ".ORNO    as Field8, "
            End If
            If pFormNo = "000015" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "Code"
                DataGrid1.Columns.Item(7).HeaderText = "�l�[�z��"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".Code    as Field7, "
                SQL = SQL + wTableName + ".AppendReason    as Field8, "
            End If
            If pFormNo = "000016" Then
                DataGrid1.Columns.Item(4).HeaderText = "��f����"
                DataGrid1.Columns.Item(5).HeaderText = "�ʤJ��/�[��"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".ColorItem  as Field5, "
                SQL = SQL + wTableName + ".Pullerprice" + " +'/'+ " + wTableName + ".price as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If
            If pFormNo = "800001" Then
                DataGrid1.Columns.Item(4).HeaderText = "�e�U��"
                DataGrid1.Columns.Item(5).HeaderText = "���A"
                DataGrid1.Columns.Item(6).HeaderText = "�ק�z�����O"
                DataGrid1.Columns.Item(7).HeaderText = "�ק�z��"

                SQL = SQL + wTableName + ".Sheet   as Field5, "
                SQL = SQL + wTableName + ".Status  as Field6, "
                SQL = SQL + wTableName + ".MReasonType as Field7, "
                SQL = SQL + wTableName + ".MReason as Field8, "
            End If

            If pFormNo = "007001" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(6).HeaderText = "���Y�~�W"
                DataGrid1.Columns.Item(7).HeaderText = "�j�f��"


                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".MapNo  as Field6, "
                SQL = SQL + wTableName + ".SliderCode as Field7, "
                SQL = SQL + wTableName + ".YKKGroup as Field8, "
            End If


            If DMPSelect.SelectedValue = "1" Then   ''�P�_�O�_����Ҩ�O��
                Dim bc1 As New BoundColumn
                Dim bc2 As New BoundColumn

                bc1.DataField = "Field9"
                bc1.HeaderText = "�����Ҩ�O"
                bc2.DataField = "Field10"
                bc2.HeaderText = "�Ҩ��ʤJ��"

                DataGrid1.Columns.AddAt(8, bc1) ''�s�WDataGrid���
                DataGrid1.Columns.AddAt(9, bc2)
                SetFieldWidth()

                SQL = SQL + wTableName + ".ArMoldFee as Field9, "
                SQL = SQL + wTableName + ".PurMold as Field10, "

                Dim bc3 As New BoundColumn
                bc3.DataField = "Field11"
                bc3.HeaderText = "�e�U�t��"
                DataGrid1.Columns.AddAt(10, bc3) ''�s�WDataGrid���
                SQL = SQL + wTableName + ".SellVendor as Field11, "

                If pFormNo <> "000011" And pFormNo <> "000012" And pFormNo <> "000015" Then
                    Dim bc4 As New BoundColumn
                    bc4.DataField = "Field12"
                    bc4.HeaderText = "�~�`�t��"
                    DataGrid1.Columns.AddAt(11, bc4) ''�s�WDataGrid���
                    SQL = SQL + wTableName + ".SUPPILER as Field12, "

                End If



            End If
            '2012/5/16 Add-Start by Joy
            If pFormNo = "000003" Or pFormNo = "000007" Then
                SQL = SQL + wTableName + ".MMSSts as MSts, "
            End If

            '20180821 jessica 
            If pFormNo = "000001" Or pFormNo = "000002" Or pFormNo = "000003" Or pFormNo = "000007" Or pFormNo = "000011" Or pFormNo = "000012" Or pFormNo = "000014" Or pFormNo = "000015" Then
                If DMPSelect.SelectedValue <> "1" Then
                    Dim bc3 As New BoundColumn
                    bc3.DataField = "Field9"
                    bc3.HeaderText = "�e�U�t��"
                    DataGrid1.Columns.AddAt(8, bc3) ''�s�WDataGrid���
                    SQL = SQL + wTableName + ".SellVendor as Field9, "
                    If pFormNo <> "000011" And pFormNo <> "000012" And pFormNo <> "000015" Then
                        Dim bc4 As New BoundColumn
                        bc4.DataField = "Field10"
                        bc4.HeaderText = "�~�`�t��"
                        DataGrid1.Columns.AddAt(9, bc4) ''�s�WDataGrid���
                        SQL = SQL + wTableName + ".Suppiler as Field10, "

                    End If


                End If



            End If
            'jessica 20190805 
            'Add-End
            SQL = SQL + " (Select Top 1 Substring(ViewURL,1,CharIndex('?', ViewURL)) From V_WaitHandle_01  Where FormNo = " + wTableName + ".FormNo" + " and FormNo='" + DFormName.SelectedValue + "' ) "
            SQL = SQL + " + 'pFormNo=' + " + wTableName + ".FormNo + '&' + 'pFormSno=' + str(" + wTableName + ".FormSno, Len(" + wTableName + ".FormSno))"
            SQL = SQL + " As ViewURL, "
          
            SQL = SQL + "Case When " + wTableName + ".CreateTime <= '" + SaveTime + "' Then '�w�ʦs' Else '���ʦs' End As WorkFlow, "
            SQL = SQL + " '' As OPURL "

            SQL = SQL + "FROM " + wTableName + " "
            '���s�u�{�~�ȳs��
            If pFormNo = "000004" Or pFormNo = "000005" Then
                SQL = SQL + "Left Outer Join F_ManufInModSheet ON " + wTableName + ".NFormNo=F_ManufInModSheet.FormNo "
                SQL = SQL + "                                 And " + wTableName + ".NFormSno=F_ManufInModSheet.FormSno "
            End If
            '�~�s, �u�{�~�ȳs��
            If pFormNo = "000008" Or pFormNo = "000009" Then
                SQL = SQL + "Left Outer Join F_ManufOutModSheet ON " + wTableName + ".NFormNo=F_ManufOutModSheet.FormNo "
                SQL = SQL + "                                  And " + wTableName + ".NFormSno=F_ManufOutModSheet.FormSno "
            End If
            '------------------------------------
            SQL = SQL + "Where " + wTableName + ".FormNo <> '" + "@@@@@@@@@@@@@" + "'"
            '���
            If DFormName.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".FormNo = '" + DFormName.SelectedValue + "'"
            End If

            '�}�o���A
            If DProgress.SelectedValue <> "ALL" Then
                If DProgress.SelectedValue = "1" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '0'  "
                End If
                If DProgress.SelectedValue = "2" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '1'  "
                End If
            End If

            '�}�o�������A
            If DSts.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".Sts = '" + DSts.SelectedValue + "'"
            End If

            '����
            If DDivision.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".Division = '" + DDivision.SelectedValue + "'"
            End If
            '�̿�H
            If DPerson.Text <> "�̿��" And DPerson.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".Person Like '%" + DPerson.Text + "%'"
            End If
            'Buyer
            If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.Buyer Like '%" + DBuyer.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.Buyer Like '%" + DBuyer.Text + "%'"
                    Else
                        SQL = SQL + " And " + wTableName + ".Buyer Like '%" + DBuyer.Text + "%'"
                    End If
                End If
            End If


            '20180821 jessica �W�[�e�U�t��
            'sellvendor
            If DSellVendor.Text <> "�e�U�t��" And DSellVendor.Text <> "" Then

                SQL = SQL + " And " + wTableName + ".SellVendor Like '%" + DSellVendor.Text + "%'"

            End If

            '202401191 jessica �W�[�~�` 
            'sellvendor
            If DSUPPILER.Text <> "�~�`�t��" And DSUPPILER.Text <> "" Then

                SQL = SQL + " And " + wTableName + ".SUPPILER Like '%" + DSUPPILER.Text + "%'"

            End If




            '�]�p��
            If DMakeMap.Text <> "�]�p��" And DMakeMap.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".MakeMap Like '%" + DMakeMap.Text + "%'"
            End If
            'Slider Code
            If DSliderCode.Text <> "Slider Code" And DSliderCode.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.SliderCode Like '%" + DSliderCode.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.SliderCode Like '%" + DSliderCode.Text + "%'"
                    Else
                        If pFormNo = "000010" Then
                            SQL = SQL + " And " + wTableName + ".SliderGRCode Like '%" + DSliderCode.Text + "%'"
                        Else
                            SQL = SQL + " And " + wTableName + ".SliderCode Like '%" + DSliderCode.Text + "%'"
                        End If
                    End If
                End If
            End If
            'Spec
            If DSpec.Text <> "Size-Chain-����" And DSpec.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.Spec Like '%" + DSpec.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.Spec Like '%" + DSpec.Text + "%'"
                    Else
                        SQL = SQL + " And " + wTableName + ".Spec Like '%" + DSpec.Text + "%'"
                    End If
                End If
            End If
            'MapNo
            If DMapNo.Text <> "�ϸ�" And DMapNo.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.MapNo Like '%" + DMapNo.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.MapNo Like '%" + DMapNo.Text + "%'"
                    Else
                        SQL = SQL + " And " + wTableName + ".MapNo Like '%" + DMapNo.Text + "%'"
                    End If
                End If
            End If
            'Cpsc
            If DCpsc.Text <> "CPSC" And DCpsc.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.Cpsc Like '%" + DCpsc.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.Cpsc Like '%" + DCpsc.Text + "%'"
                    Else
                        SQL = SQL + " And " + wTableName + ".Cpsc Like '%" + DCpsc.Text + "%'"
                    End If
                End If
            End If
            'No
            If DNo.Text <> "�e�U��No." And DNo.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".No Like '%" + DNo.Text + "%'"
            End If
            'Code
            If DCode.Text <> "Code" And DCode.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".Code Like '%" + DCode.Text + "%'"
            End If
            ''CompletedTime
            If DMPSelect.SelectedValue = "1" Then   ''�P�_�O�_����Ҩ�O��
                SQL = SQL + " And " + wTableName + ".CompletedTime between '" + DSDate.Text + " 00:00:00' and '" + DEDate.Text + " 23:59:59'"
            Else
                SQL = SQL + " And " + wTableName + ".CompletedTime between '" + DSDate.Text + " 00:00:00' and '" + DEDate.Text + " 23:59:59'"
            End If
            'Sort
            SQL = SQL + " Order by " + wTableName + ".No desc "

            OleDbConnection1.Open()
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "WaitHandle")
            '
            '2012/5/16 Add-Start by Joy
            If pFormNo = "000003" Or pFormNo = "000007" Then
                DBTable1 = DBDataSet2.Tables("WaitHandle")
                If DBTable1.Rows.Count > 0 Then
                    For i = 0 To DBTable1.Rows.Count - 1
                        If DBTable1.Rows(i).Item("MSts") = 1 Then
                            DBTable1.Rows(i).Item("Field1") = DBTable1.Rows(i).Item("Field1") + "-�Ҽo"
                        End If
                    Next
                End If
            End If
            'Add-End
            '
            DataGrid1.DataSource = DBDataSet2
            DataGrid1.DataBind()
            OleDbConnection1.Close()
        End If

    End Sub

    Sub Search_Item_Attribute()
        DNo.ReadOnly = True
        DPerson.ReadOnly = True
        DBuyer.ReadOnly = True
        DMakeMap.ReadOnly = True
        DSpec.ReadOnly = True
        DMapNo.ReadOnly = True
        DSliderCode.ReadOnly = True
        DCpsc.ReadOnly = True
        DCode.ReadOnly = True

        '�̿��
        DPerson.Text = "�̿��"
        'Buyer
        DBuyer.Text = "Buyer"

        'SellVendor 
        DSellVendor.Text = "�e�U�t��"

        'SellVendor 
        DSUPPILER.Text = "�~�`�t��"

        '�]�p��
        DMakeMap.Text = "�]�p��"
        'SliderCode
        DSliderCode.Text = "Slider Code"
        'Spec
        DSpec.Text = "Size-Chain-����"
        'No
        DNo.Text = "�e�U��No."
        '�ϸ�
        DMapNo.Text = "�ϸ�"
        'Cpsc
        DCpsc.Text = "CPSC"
        'Code
        DCode.Text = "Code"

        DNo.ReadOnly = False
        DPerson.ReadOnly = False
        If pFormNo = "000001" Then
            DBuyer.ReadOnly = False
            DMakeMap.ReadOnly = False
            DSpec.ReadOnly = False
            DMapNo.ReadOnly = False
            DCpsc.ReadOnly = False
        End If
        If pFormNo = "000002" Then
            DBuyer.ReadOnly = False
            DMakeMap.ReadOnly = False
            DMapNo.ReadOnly = False
            DCpsc.ReadOnly = False
        End If
        If pFormNo = "000003" Or pFormNo = "000004" Or pFormNo = "000005" Or + _
           pFormNo = "000007" Or pFormNo = "000008" Or pFormNo = "000009" Then
            If pFormNo = "000003" Or pFormNo = "000007" Then    ''�P�_�Ҩ�O�άO�_�}��
                DMPSelect.Enabled = True
            End If
            DBuyer.ReadOnly = False
            DSliderCode.ReadOnly = False
            DSpec.ReadOnly = False
            DMapNo.ReadOnly = False
            DCpsc.ReadOnly = False
        End If
        If pFormNo = "000010" Then
            DSliderCode.ReadOnly = False
        End If
        If pFormNo = "000011" Then
            DBuyer.ReadOnly = False
            DSpec.ReadOnly = False
            DSliderCode.ReadOnly = False
        End If
        If pFormNo = "000012" Then
            DBuyer.ReadOnly = False
            DMapNo.ReadOnly = False
            DSliderCode.ReadOnly = False
        End If
        If pFormNo = "000013" Then
        End If
        If pFormNo = "000014" Then
            DBuyer.ReadOnly = False
            DSpec.ReadOnly = False
            DCode.ReadOnly = False
        End If
        If pFormNo = "000015" Then
            DBuyer.ReadOnly = False
            DSpec.ReadOnly = False
        End If
        If pFormNo = "000016" Then
            DMapNo.ReadOnly = False
            DSliderCode.ReadOnly = False
        End If
        If pFormNo = "800001" Then
        End If
        If pFormNo = "007001" Then
            DBuyer.ReadOnly = False
            DMapNo.ReadOnly = False
            DSliderCode.ReadOnly = False
        End If

        If pFormNo = "000001" Or pFormNo = "000002" Or pFormNo = "000003" Or pFormNo = "000007" Or pFormNo = "000011" Or pFormNo = "000012" Or pFormNo = "000014" Or pFormNo = "000015" Then
            DSellVendor.ReadOnly = False
            DSUPPILER.ReadOnly = False
        Else
            DSellVendor.ReadOnly = True
            DSUPPILER.ReadOnly = True
        End If


    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        pFormNo = DFormName.SelectedValue
        '
        If DKeepData.SelectedValue = "0" Then
            DataList()      '--���ʦs���
        Else
            If DKeepData.SelectedValue = "1" Then
                KeepDataList()  '--�w�ʦs���
            Else
                FormDataList()  '--�����
            End If
        End If
    End Sub


    Private Sub DFormName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged
        pFormNo = DFormName.SelectedValue

        DNo.ReadOnly = True
        DPerson.ReadOnly = True
        DBuyer.ReadOnly = True
        DSellVendor.ReadOnly = True
        DSUPPILER.ReadOnly = True
        DMakeMap.ReadOnly = True
        DSpec.ReadOnly = True
        DMapNo.ReadOnly = True
        DSliderCode.ReadOnly = True
        DCpsc.ReadOnly = True
        DCode.ReadOnly = True

        DProgress.Enabled = True    ''�w�]���Ұ�
        DProgress.SelectedIndex = 0 ''�w�]��ALL
        DDivision.Enabled = True    ''�w�]���Ұ�
        DSts.SelectedIndex = 0      ''�w�]��ALL
        DMPSelect.Enabled = False   ''�w�]�Ҩ�O�οﶵ�����}��
        DMPSelect.SelectedIndex = 0 ''�w�]���@��d��

        DataGrid1.DataBind()    ''�M��DataGrid


        '�̿��
        DPerson.Text = "�̿��"
        'Buyer
        DBuyer.Text = "Buyer"

        'sellvendor
        DSellVendor.Text = "�e�U�t��"

        'sUPPILER
        DSUPPILER.Text = "�~�`�t��"


        '�]�p��
        DMakeMap.Text = "�]�p��"
        'SliderCode
        DSliderCode.Text = "Slider Code"
        'Spec
        DSpec.Text = "Size-Chain-����"
        'No
        DNo.Text = "�e�U��No."
        '�ϸ�
        DMapNo.Text = "�ϸ�"
        'Cpsc
        DCpsc.Text = "CPSC"
        'Code
        DCode.Text = "Code"

        DNo.ReadOnly = False
        DPerson.ReadOnly = False
        If pFormNo = "000001" Then
            DBuyer.ReadOnly = False
            DMakeMap.ReadOnly = False
            DSpec.ReadOnly = False
            DMapNo.ReadOnly = False
            DCpsc.ReadOnly = False
        End If
        If pFormNo = "000002" Then
            DBuyer.ReadOnly = False
            DMakeMap.ReadOnly = False
            DMapNo.ReadOnly = False
            DCpsc.ReadOnly = False
        End If
        If pFormNo = "000003" Or pFormNo = "000004" Or pFormNo = "000005" Or + _
           pFormNo = "000007" Or pFormNo = "000008" Or pFormNo = "000009" Then
            If pFormNo = "000003" Or pFormNo = "000007" Then    ''�P�_�Ҩ�O�οﶵ�O�_�n�}��
                DMPSelect.Enabled = True
            End If
            DBuyer.ReadOnly = False
            DSliderCode.ReadOnly = False
            DSpec.ReadOnly = False
            DMapNo.ReadOnly = False
            DCpsc.ReadOnly = False
        End If
        If pFormNo = "000010" Then
            DSliderCode.ReadOnly = False
        End If
        If pFormNo = "000011" Then
            DBuyer.ReadOnly = False
            DSpec.ReadOnly = False
            DSliderCode.ReadOnly = False
        End If
        If pFormNo = "000012" Then
            DBuyer.ReadOnly = False
            DMapNo.ReadOnly = False
            DSliderCode.ReadOnly = False
            DCpsc.ReadOnly = False
        End If
        If pFormNo = "000013" Then
        End If
        If pFormNo = "000014" Then
            DBuyer.ReadOnly = False
            DSpec.ReadOnly = False
            DCode.ReadOnly = False
        End If
        If pFormNo = "000015" Then
            DBuyer.ReadOnly = False
            DSpec.ReadOnly = False
        End If
        If pFormNo = "000016" Then
            DMapNo.ReadOnly = False
            DSliderCode.ReadOnly = False
        End If
        If pFormNo = "007001" Then
            DBuyer.ReadOnly = False
            DMapNo.ReadOnly = False
            DSliderCode.ReadOnly = False
        End If


        If pFormNo = "000001" Or pFormNo = "000002" Or pFormNo = "000003" Or pFormNo = "000007" Or pFormNo = "000011" Or pFormNo = "000012" Or pFormNo = "000014" Or pFormNo = "000015" Then
            DSellVendor.ReadOnly = False
            DSUPPILER.ReadOnly = False
        Else
            DSellVendor.ReadOnly = True
            DSUPPILER.ReadOnly = True
        End If

        ''SetFiledDdefault()
    End Sub

    ''*****************************************************************
    ''**
    ''**     �ˬd�U���z����쪬�A
    ''**
    ''*****************************************************************

    ''Sub SetFiledDdefault()
    ''    Dim wTempTextBox As New TextBox
    ''    Dim wDefault As String
    ''    Dim i As Integer

    ''    For i = 1 To 6
    ''        Select Case i
    ''            Case 1
    ''                wTempTextBox = DBuyer
    ''                wDefault = "Buyer"
    ''            Case 2
    ''                wTempTextBox = DMapNo
    ''                wDefault = "�ϸ�"
    ''            Case 3
    ''                wTempTextBox = DMakeMap
    ''                wDefault = "�]�p��"
    ''            Case 4
    ''                wTempTextBox = DSliderCode
    ''                wDefault = "Slider Code"
    ''            Case 5
    ''                wTempTextBox = DSpec
    ''                wDefault = "Size-Chain-����"
    ''            Case 6
    ''                wTempTextBox = DCpsc
    ''                wDefault = "CPSC"
    ''        End Select

    ''        If wTempTextBox.ReadOnly Then       ''�Y��쬰��Ū,�h��줺�e��_���w�]
    ''            wTempTextBox.Text = wDefault
    ''        End If
    ''    Next
    ''End Sub
    ''*****************************************************************
    ''**
    ''**     �]�w�d�߼Ҩ�O����e
    ''**
    ''*****************************************************************

    Sub SetFieldWidth()     ''�]�w��e
        Dim i As Integer
        Dim TempWidth As Integer
        Dim test As String

        For i = 0 To 10
            Select Case i
                Case 0
                    TempWidth = 90
                Case 1
                    TempWidth = 90
                Case 2
                    TempWidth = 60
                Case 3
                    TempWidth = 30
                Case 4
                    TempWidth = 90
                Case 5
                    TempWidth = 90
                Case 6
                    TempWidth = 90
                Case 7
                    TempWidth = 90
                Case 8
                    TempWidth = 60
                Case 9
                    TempWidth = 60
                Case 10
                    TempWidth = 30
                Case 11
                    TempWidth = 30

            End Select
            DataGrid1.Columns.Item(i).ItemStyle.Width = New Unit(TempWidth)
            DataGrid1.Columns.Item(i).HeaderStyle.Width = New Unit(TempWidth)
        Next
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        pFormNo = DFormName.SelectedValue
        '
        '
        If DKeepData.SelectedValue = "0" Then
            DataList()      '--���ʦs���
        Else
            If DKeepData.SelectedValue = "1" Then
                KeepDataList()  '--�w�ʦs���
            Else
                FormDataList()  '--�����
            End If
        End If
    End Sub
    '*****************************************************************
    '**
    '**     ��Excel�@�ε{��
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '�{���O���P

        pFormNo = DFormName.SelectedValue
        DataGrid1.AllowPaging = False   '�{���O���P
        '
        If DKeepData.SelectedValue = "0" Then
            '--���ʦs���
            DataList()
        Else
            If DKeepData.SelectedValue = "1" Then
                KeepDataList()  '--�w�ʦs���
            Else
                FormDataList()  '--�����
            End If
        End If

        Response.AppendHeader("Content-Disposition", "attachment;filename=Commission_List.xls")     '�{���O���P
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        DataGrid1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        DataGrid1.AllowPaging = wAllowPaging        '�{���O���P
    End Sub

    ''*****************************************************************
    ''**
    ''**     �]�w��쪬�A
    ''**
    ''*****************************************************************

    Private Sub DMPSelect_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DMPSelect.SelectedIndexChanged
        DataGrid1.DataBind()    ''�M��DataGrid

        If DMPSelect.SelectedValue = "1" Then   ''�P�_�Ҩ�O�άO�_���
            DProgress.SelectedIndex = 2         ''�]�w���}�o����
            DSts.SelectedIndex = 1              ''�]�w��OK
            DProgress.Enabled = False
            DDivision.Enabled = False
            DNo.ReadOnly = True
            DPerson.ReadOnly = True
            DBuyer.ReadOnly = True
            DMakeMap.ReadOnly = True
            DSpec.ReadOnly = True
            DMapNo.ReadOnly = True
            DSliderCode.ReadOnly = True
            DCpsc.ReadOnly = True
        Else
            DFormName_SelectedIndexChanged(sender, e)   ''�I�s��ƥH������쪬�A
        End If
    End Sub


    Private Sub DataGrid1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class
