Imports System.Data
Imports System.Data.OleDb

Public Class HR_InqCommission
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DProgress As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSts As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BEDate As System.Web.UI.WebControls.Button
    Protected WithEvents DEDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSDate As System.Web.UI.WebControls.Button
    Protected WithEvents DSDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DName As System.Web.UI.WebControls.TextBox

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
    Dim wLevel As String = ""
    Dim wDivision As String = ""
    Dim wName As String = ""
    Dim NowDateTime As String       '�{�b����ɶ�

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "HR_InqCommission.aspx"

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

    End Sub

    Sub SetSearchItem()
        Dim i As Integer = 0
        Dim wSts As String
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        '**** �]�w ����/�m�W
        OleDbConnection1.Open()

        '���o�z���v��
        SQL = "Select * From M_Referp  "
        SQL = SQL + "Where Cat='1999'  "
        SQL = SQL + "  and DKey='" & "AUTHORITY-" & Request.QueryString("pUserID") & "' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Referp")
        If DBDataSet1.Tables("M_Referp").Rows.Count > 0 Then
            If DBDataSet1.Tables("M_Referp").Rows(0).Item("Data") = "ALL" Then
                wLevel = "ALL"
            Else
                wLevel = "DIVISION"
            End If
        Else
            wLevel = "PERSON"
        End If

        DBDataSet1.Clear()
        DDivision.Items.Clear()
        If wLevel = "PERSON" Then
            '���o�ӤH��T
            SQL = "Select * From M_Users  "
            SQL = SQL + "Where UserID='" & Request.QueryString("pUserID") & "' "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "M_Users")
            DBTable1 = DBDataSet1.Tables("M_Users")
            If DBTable1.Rows.Count > 0 Then
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(0).Item("HRWDivName")
                ListItem1.Value = DBTable1.Rows(0).Item("HRWDivName")
                DDivision.Items.Add(ListItem1)
                wName = DBTable1.Rows(0).Item("UserName")
            End If
        Else
            If wLevel = "DIVISION" Then
                '���o�ҫ��w����
                SQL = "Select * From M_Referp  "
                SQL = SQL + "Where Cat='1999'  "
                SQL = SQL + "  and DKey='" & "DIVISION-" & Request.QueryString("pUserID") & "' "
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    DDivision.Items.Add(ListItem1)
                Next
                wName = "�ӽЪ�"
            Else
                '���o��������
                If wLevel = "ALL" Then
                    SQL = "Select HRWDivName From M_Users Group by HRWDivName Order by HRWDivName "
                    DDivision.Items.Add("ALL")
                    Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter2.Fill(DBDataSet1, "M_Users")
                    DBTable1 = DBDataSet1.Tables("M_Users")
                    For i = 0 To DBTable1.Rows.Count - 1
                        Dim ListItem1 As New ListItem
                        ListItem1.Text = DBTable1.Rows(i).Item("HRWDivName")
                        ListItem1.Value = DBTable1.Rows(i).Item("HRWDivName")
                        DDivision.Items.Add(ListItem1)
                    Next
                    wName = "�ӽЪ�"
                End If
            End If
        End If

        DBDataSet1.Clear()
        '���
        SQL = "SELECT FormName, FormNo FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And FormNo >= '001001' And FormNo <= '001050' "
        SQL = SQL + "  And (InqAuthority = '0' "
        SQL = SQL + "       Or (InqAuthority = '1' "
        SQL = SQL + "           And (InqUser like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                Or InqUserName like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "               ) "
        SQL = SQL + "          ) "
        SQL = SQL + "      ) "
        SQL = SQL + "Order by FormNo "

        DFormName.Items.Clear()
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "M_Form")
        DBTable1 = DBDataSet1.Tables("M_Form")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("FormName")
            ListItem1.Value = DBTable1.Rows(i).Item("FormNo")
            If ListItem1.Value = pFormNo Then ListItem1.Selected = True
            DFormName.Items.Add(ListItem1)
        Next

        OleDbConnection1.Close()

        '���
        DSDate.Text = CStr(DateAdd("d", -180, DateTime.Now.Today))
        DEDate.Text = CStr(DateTime.Now.Today)
        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DEDate');"

        pFormNo = DFormName.SelectedValue

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
            DataGrid1.Columns.Item(0).HeaderText = "No."
            DataGrid1.Columns.Item(1).HeaderText = "���A"
            DataGrid1.Columns.Item(2).HeaderText = "�ӽЪ�"
            DataGrid1.Columns.Item(3).HeaderText = "�ӽФ�"

            SQL = "SELECT "
            SQL = SQL + "Case " + wTableName + ".No When '' Then '���s��' Else " + wTableName + ".No End As Field1, "
            SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '�֩w��' When '1' Then '����' When '2' Then '����' Else '����' End As Field2, "
            SQL = SQL + "V_WaitHandle_01.HRWDivName + '--' + "
            SQL = SQL + wTableName + ".Name as Field3, "
            SQL = SQL + "Convert(VARCHAR(10), V_WaitHandle_01.ApplyTime, 111) as Field4, "

            If pFormNo = "001001" Then
                DataGrid1.Columns.Item(4).HeaderText = "����"
                DataGrid1.Columns.Item(5).HeaderText = "�[�Z���"
                DataGrid1.Columns.Item(6).HeaderText = "�뭹"

                SQL = SQL + wTableName + ".DateType as Field5, "
                SQL = SQL + "Convert(VARCHAR(10), " + wTableName + ".OverTimeDate, 111) as Field6, "
                SQL = SQL + wTableName + ".Food as Field7, "
            End If
            If pFormNo = "001002" Then
                DataGrid1.Columns.Item(4).HeaderText = "���O"
                DataGrid1.Columns.Item(5).HeaderText = "�а����"
                DataGrid1.Columns.Item(6).HeaderText = "�Ѽ�"

                SQL = SQL + wTableName + ".Vacation as Field5, "
                SQL = SQL + "Convert(VARCHAR(10), " + wTableName + ".AStartDate, 111) + ' ' + str(AStartH) + ':00' + '~' + "
                SQL = SQL + "Convert(VARCHAR(10), " + wTableName + ".AEndDate, 111)   + ' ' + str(AEndH) + ':00' as Field6, "
                SQL = SQL + wTableName + ".ADays as Field7, "
            End If
            If pFormNo = "001003" Then
                DataGrid1.Columns.Item(4).HeaderText = "�ت��a"
                DataGrid1.Columns.Item(5).HeaderText = "�~�X���"
                DataGrid1.Columns.Item(6).HeaderText = "���"

                SQL = SQL + wTableName + ".Place as Field5, "
                SQL = SQL + "Convert(VARCHAR(10), " + wTableName + ".AStartDate, 111) + ' ' + str(AStartH) + ':00' + '~' + "
                SQL = SQL + "Convert(VARCHAR(10), " + wTableName + ".AEndDate, 111)   + ' ' + str(AEndH) + ':00' as Field6, "
                SQL = SQL + "str(" + wTableName + ".ADay) + '��' + str(" + wTableName + ".AHour) + '��' as Field7, "
            End If
            If pFormNo = "001004" Then
                DataGrid1.Columns.Item(4).HeaderText = "�ʶԤ��"
                DataGrid1.Columns.Item(5).HeaderText = "�ɤu���"
                DataGrid1.Columns.Item(6).HeaderText = "�ɼ�"
                SQL = SQL + "Convert(VARCHAR(10), " + wTableName + ".WorkDate, 111) + ' ' + str(" + wTableName + ".wStartH) +  ':' + str(" + wTableName + ".wStartM) + ':00' + '~' + str(" + wTableName + ".wEndH) +  ':' + str(" + wTableName + ".wEndM) + ':00' as Field5, "
                SQL = SQL + "Convert(VARCHAR(10), " + wTableName + ".AWorkDate, 111) + ' ' + str(" + wTableName + ".AStartH) +  ':' + str(" + wTableName + ".AStartM) + ':00' + '~' + str(" + wTableName + ".AEndH) +  ':' + str(" + wTableName + ".AEndM) + ':00' as Field6, "
                SQL = SQL + "str(" + wTableName + ".AH) + '��' + str(" + wTableName + ".AM) + '��' as Field7, "
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

            SQL = SQL + "Where V_WaitHandle_01.Step  < '10' "
            '���
            If DFormName.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".FormNo = '" + DFormName.SelectedValue + "'"
            End If

            '���A
            If DProgress.SelectedValue <> "ALL" Then
                If DProgress.SelectedValue = "1" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '0'  "
                End If
                If DProgress.SelectedValue = "2" Then
                    SQL = SQL + " And " + wTableName + ".Sts <>  '0'  "
                End If
            End If

            '�������A
            If DSts.SelectedValue <> "ALL" Then
                If DSts.SelectedValue <> "3" Then
                    SQL = SQL + " And " + wTableName + ".Sts = '" + DSts.SelectedValue + "'"
                Else
                    SQL = SQL + " And (" + wTableName + ".Sts = '2' " + " or " + wTableName + ".Sts = '3') "
                End If
            End If

            '����
            If DDivision.SelectedValue <> "ALL" Then
                SQL = SQL + " And V_WaitHandle_01.HRWDivName = '" + DDivision.SelectedValue + "'"
            End If
            '�ӽЪ�
            If DName.Text <> "�ӽЪ�" And DName.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".Name Like '%" + DName.Text + "%'"
            End If
            'No
            If DNo.Text <> "No." And DNo.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".No Like '%" + DNo.Text + "%'"
            End If
            '�ӽФ�
            SQL = SQL + " And " + wTableName + ".Date between '" + DSDate.Text + " 00:00:00' and '" + DEDate.Text + " 23:59:59'"
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
        DNo.Text = "No."
        DNo.ReadOnly = False
        If wLevel = "PERSON" Then
            DName.Text = wName
            DName.ReadOnly = True
        Else
            DName.Text = "�ӽЪ�"
            DName.ReadOnly = False
        End If
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        pFormNo = DFormName.SelectedValue
        DataList()
    End Sub


    Private Sub DFormName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged

    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        pFormNo = DFormName.SelectedValue
        DataList()
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
        DataList()                      '�{���O���P

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

End Class
