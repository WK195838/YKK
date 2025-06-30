Imports System.Data
Imports System.Data.OleDb

Public Class OPInqCommission
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BEDate As System.Web.UI.WebControls.Button
    Protected WithEvents DEDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSDate As System.Web.UI.WebControls.Button
    Protected WithEvents DSDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSts As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMakeMap As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCpsc As System.Web.UI.WebControls.TextBox

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
    Dim pFormNo As String
    Dim pSts As String
    Dim NowDateTime As String       '�{�b����ɶ�

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "OPInqCommission.aspx"

        SetParameter()          '�]�w�@�ΰѼ�
        If Not Me.IsPostBack Then
            SetSearchItem()
            If pFormNo <> "" And pSts <> "" Then
                DataList()
            End If
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
        pSts = Request.QueryString("pSts")
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")

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
        SQL = SQL + "  And ( (FormNo >= '000001' And FormNo <= '001000') or FormNo='800001' ) "
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

        If pSts = "" Then wSts = "OK"
        If pSts = "OK" Or pSts = "OKY" Or pSts = "OKBY" Then wSts = "OK"
        If pSts = "NG" Or pSts = "NGY" Or pSts = "NGBY" Then wSts = "NG"
        If pSts = "Cancel" Or pSts = "CancelY" Or pSts = "CancelBY" Then wSts = "����"

        For i = 0 To DSts.Items.Count - 1
            If DSts.Items.Item(i).Text = wSts Then
                DSts.Items.Item(i).Selected = True
            End If
        Next

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
        If pSts = "OK" Or pSts = "NG" Or pSts = "Cancel" Then
            DSDate.Text = CStr(DateTime.Now.Year) + "/" + CStr(DateTime.Now.Month) + "/01"
            If DateTime.Now.Month = 1 Or DateTime.Now.Month = 3 Or DateTime.Now.Month = 5 Or DateTime.Now.Month = 7 Or DateTime.Now.Month = 8 Or DateTime.Now.Month = 10 Or DateTime.Now.Month = 12 Then
                DEDate.Text = CStr(DateTime.Now.Year) + "/" + CStr(DateTime.Now.Month) + "/31"
            Else
                If DateTime.Now.Month = 2 Then
                    DEDate.Text = CStr(DateTime.Now.Year) + "/" + CStr(DateTime.Now.Month) + "/28"
                Else
                    DEDate.Text = CStr(DateTime.Now.Year) + "/" + CStr(DateTime.Now.Month) + "/30"
                End If
            End If
        End If
        If pSts = "OKY" Or pSts = "NGY" Or pSts = "CancelY" Then
            DSDate.Text = CStr(DateTime.Now.Year) + "/01/01"
            DEDate.Text = CStr(DateTime.Now.Year) + "/12/31"
        End If
        If pSts = "OKBY" Or pSts = "NGBY" Or pSts = "CancelBY" Then
            DSDate.Text = CStr(DateTime.Now.Year - 42) + "/01/01"
            DEDate.Text = CStr(DateTime.Now.Year - 1) + "/12/31"
        End If
        If pSts = "" Then
            DSDate.Text = CStr(DateAdd("d", -31, DateTime.Now.Today))
            DEDate.Text = CStr(DateTime.Now.Today)
        End If
        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DEDate');"

        pFormNo = DFormName.SelectedValue
        pSts = DSts.SelectedValue

        Search_Item_Attribute()

    End Sub

    Sub DataList()
        Dim i As Integer = 0
        Dim wTableName As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
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
            DataGrid1.Columns.Item(0).HeaderText = "�e�UNo"
            DataGrid1.Columns.Item(1).HeaderText = "�̿��"
            DataGrid1.Columns.Item(2).HeaderText = "�̿��"
            DataGrid1.Columns.Item(3).HeaderText = "������"
            SQL = "SELECT "
            'SQL = SQL + wTableName + ".No as Field1, "
            SQL = SQL + "Case " + wTableName + ".No When '' Then '���s��' Else " + wTableName + ".No End As Field1, "
            SQL = SQL + wTableName + ".Division + '--' + "
            SQL = SQL + wTableName + ".Person as Field2, "
            SQL = SQL + "V_WaitHandle_01.ApplyTime as Field3, "
            SQL = SQL + wTableName + ".CompletedTime as Field4, "
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
                DataGrid1.Columns.Item(6).HeaderText = "���Y����"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".SliderType as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If

            If pFormNo = "000012" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "�e�U�t��"
                DataGrid1.Columns.Item(6).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".SellVendor as Field6, "
                SQL = SQL + wTableName + ".MapNo as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If

            If pFormNo = "000013" Then
                DataGrid1.Columns.Item(4).HeaderText = "Slider Code"
                DataGrid1.Columns.Item(5).HeaderText = "�ت�"
                DataGrid1.Columns.Item(6).HeaderText = ""
                DataGrid1.Columns.Item(7).HeaderText = ""

                SQL = SQL + wTableName + ".SliderCode as Field5, "
                SQL = SQL + wTableName + ".Target as Field6, "
                SQL = SQL + "'' as Field7, "
                SQL = SQL + "'' as Field8, "
            End If

            If pFormNo = "000014" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-����"
                DataGrid1.Columns.Item(6).HeaderText = "���s/�~�`"
                DataGrid1.Columns.Item(7).HeaderText = "Code"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".ManufType as Field7, "
                SQL = SQL + wTableName + ".Code as Field8, "
            End If

            If pFormNo = "000015" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Code"
                DataGrid1.Columns.Item(6).HeaderText = "�e�U�t��"
                DataGrid1.Columns.Item(7).HeaderText = ""

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Code    as Field6, "
                SQL = SQL + wTableName + ".SellVendor as Field7, "
                SQL = SQL + "'' as Field8, "
            End If

            If pFormNo = "000016" Then
                DataGrid1.Columns.Item(4).HeaderText = "Slider Code"
                DataGrid1.Columns.Item(5).HeaderText = "�ϸ�"
                DataGrid1.Columns.Item(6).HeaderText = ""
                DataGrid1.Columns.Item(7).HeaderText = ""

                SQL = SQL + wTableName + ".SliderCode as Field5, "
                SQL = SQL + wTableName + ".MapNo as Field6, "
                SQL = SQL + "'' as Field7, "
                SQL = SQL + "'' as Field8, "
            End If

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
            SQL = SQL + "Where V_WaitHandle_01.Step = '999' "
            SQL = SQL + "  And V_WaitHandle_01.FormSno > '2000' "
            '���
            If DFormName.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".FormNo = '" + DFormName.SelectedValue + "'"
            End If

            '���A
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
                        SQL = SQL + " And " + wTableName + ".SliderCode Like '%" + DSliderCode.Text + "%'"
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
            '���
            If DSDate.Text <> "" And DEDate.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".CompletedTime >= '" + DSDate.Text + " 00:00:00" + "'"
                SQL = SQL + " And " + wTableName + ".CompletedTime <= '" + DEDate.Text + " 23:59:59" + "'"
            End If
            'Sort
            SQL = SQL + " Order by " + wTableName + ".No desc "

            OleDbConnection1.Open()
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "WaitHandle")
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

        '�̿��
        DPerson.Text = "�̿��"
        'Buyer
        DBuyer.Text = "Buyer"
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
            DBuyer.ReadOnly = False
            DSliderCode.ReadOnly = False
            DSpec.ReadOnly = False
            DMapNo.ReadOnly = False
            DCpsc.ReadOnly = False
        End If
        If pFormNo = "000010" Then
            DSliderCode.ReadOnly = False
        End If

    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        pFormNo = DFormName.SelectedValue
        pSts = DSts.SelectedValue
        DataList()
    End Sub


    Private Sub DFormName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged
        pFormNo = DFormName.SelectedValue
        pSts = DSts.SelectedValue
        Search_Item_Attribute()
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        pFormNo = DFormName.SelectedValue
        pSts = DSts.SelectedValue
        DataList()
    End Sub
End Class
