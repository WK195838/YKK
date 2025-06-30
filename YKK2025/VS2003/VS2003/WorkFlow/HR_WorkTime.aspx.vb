Imports System.Data
Imports System.Data.OleDb

Public Class HR_WorkTime
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DYear As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMonth As System.Web.UI.WebControls.DropDownList

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
    Dim wLevel As String = ""
    Dim NowDateTime As String       '�{�b����ɶ�

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "HR_ErrorCard.aspx"

        SetParameter()          '�]�w�@�ΰѼ�
        If Not Me.IsPostBack Then
            SetSearchItem()
            DataList()
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
                DName.Text = DBTable1.Rows(0).Item("UserName")
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
                DName.Text = ""
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
                    DName.Text = ""
                End If
            End If
        End If
        OleDbConnection1.Close()

        '**** �]�w �[�Z�~��
        Dim wYY As String = CStr(DateTime.Now.Year)
        Dim wMM As String = CStr(DateTime.Now.Month)
        '�[�Z�~
        DYear.Items.Clear()
        For i = 2008 To 2020
            Dim ListItem1 As New ListItem
            ListItem1.Text = CStr(i)
            ListItem1.Value = CStr(i)
            If ListItem1.Value = wYY Then ListItem1.Selected = True
            DYear.Items.Add(ListItem1)
        Next
        '�[�Z��
        If CInt(wMM) < 10 Then wMM = "0" & wMM
        DMonth.Items.Clear()
        For i = 1 To 12
            Dim ListItem1 As New ListItem
            If i < 10 Then
                ListItem1.Text = "0" & CStr(i)
                ListItem1.Value = "0" & CStr(i)
            Else
                ListItem1.Text = CStr(i)
                ListItem1.Value = CStr(i)
            End If
            If ListItem1.Value = wMM Then ListItem1.Selected = True
            DMonth.Items.Add(ListItem1)
        Next

        '**** �]�w �m�W�ݩ�
        If wLevel = "PERSON" Then
            DName.ReadOnly = True
        Else
            DName.ReadOnly = False
        End If
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        SQL = "SELECT "
        SQL = SQL + "Convert(VARCHAR(10), CDate, 111) as DateDesc, "
        SQL = SQL + "RTrim(Name)+'('+RTrim(EmpID)+')' as NameDesc, "
        SQL = SQL + "HRWDivName, JobName, TimeA, TimeB, "
        SQL = SQL + "OTDesc, OTURL, VADesc, VAURL, AWayDesc, AWayURL, AWorkDesc, AWorkURL "

        SQL = SQL + "FROM V_WorkTime_01 "
        '�~��
        SQL = SQL + "Where SelectYM = '" + DYear.SelectedValue + "/" + DMonth.SelectedValue + "' "
        '����
        If DDivision.SelectedValue <> "ALL" Then
            SQL = SQL + " And HRWDivName = '" + DDivision.SelectedValue + "'"
        End If
        '�ӽЪ�
        If DName.Text <> "" Then
            SQL = SQL + " And Name Like '%" + DName.Text + "%'"
        End If
        'Sort
        SQL = SQL + " Order by HRWDivName, EmpID, CDate Desc "

        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "WaitHandle")
        DataGrid1.DataSource = DBDataSet2
        DataGrid1.DataBind()

        OleDbConnection1.Close()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
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

        DataGrid1.AllowPaging = False   '�{���O���P
        DataList()                      '�{���O���P

        Response.AppendHeader("Content-Disposition", "attachment;filename=HR_WorkTimeList.xls")     '�{���O���P

        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
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
