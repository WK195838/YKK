Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class InqManufOutList
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DYear As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DMonth As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSOrB As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSts As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton

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
    Dim NowYear As String
    Dim NowDateTime As String       '�{�b����ɶ�

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "InqManufOutList.aspx"

        SetParameter()
        If Not Me.IsPostBack Then
            SetYear()
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
        NowYear = CStr(DateTime.Now.Year)

        ''pFormNo = Request.QueryString("pFormNo")    '��渹�X
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")

    End Sub

    '*****************************************************************
    '**
    '**     �z��~�`�e�U�ѩe�U�~��
    '**
    '*****************************************************************
    Sub SetYear()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        SQL = "SELECT CONVERT(varchar(4), CreateTime, 120) AS DYear FROM F_ManufOutSheet "
        SQL = SQL + "GROUP BY CONVERT(varchar(4), CreateTime, 120)"

        OleDbConnection1.Open()
        DYear.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ManufOutSheet")
        DBTable1 = DBDataSet1.Tables("F_ManufOutSheet")
        '�N�z��X���~���[�JDropDownList
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("DYear")
            ListItem1.Value = DBTable1.Rows(i).Item("DYear")
            If ListItem1.Value = NowYear Then ListItem1.Selected = True
            DYear.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        CreateWorkTable()
    End Sub

    Sub CreateWorkTable()
        Dim i, j As Integer
        Dim MonthlyCount, CumulativeCount As Integer             '�C����,�֭p���
        Dim SQL, SQL1, SQL2 As String
        Dim wTableName As String = "F_ManufOutSheet"
        Dim wMDate As String = DYear.SelectedValue + "-" + DMonth.SelectedValue + "-1"  '�z�����d��
        Dim wCDate As String = DYear.SelectedValue + "-1-1"                             '�֭p����d��
        Dim wMTotal, wCTotal As Integer             'wMTotal:�����`�p,wCTotal:�~����`�p(�֭p)
        Dim MonthlyRate, CumulativeRate As String   '��~�`�v,�~�~�`�v
        Dim wTColumn As String = "CreateTime"        '���e�U�η��}�o�������z�����

        Dim WorkTable As String = "Temp_InqManufOutList_" & Request.Cookies("UserID").Value

        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBTable1 As DataTable
        Dim DBDataRow As DataRow

        Dim OleDBCommand1 As New OleDbCommand
        Dim OleDbConnection1 As New OleDbConnection

        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        If DSts.SelectedValue = 1 Then  '�P�_�O�_��ܶ}�o����
            SQL2 = "and Sts = 1 "
            wTColumn = "CompletedTime"
        End If
        '�z�������
        SQL1 = "SELECT Case " + DSOrB.SelectedValue + " When '' Then '�L�W��' Else " + DSOrB.SelectedValue + " End AS SOrB, COUNT(*) AS MonthlyCount FROM F_ManufOutSheet "
        SQL1 = SQL1 + "WHERE (" + wTColumn + " > DATEADD(d, 0, '" + wMDate + "')) AND (" + wTColumn + " < DATEADD(m, 1, '" + wMDate + "')) "
        SQL1 = SQL1 + SQL2
        SQL1 = SQL1 + "GROUP BY " + DSOrB.SelectedValue
        Dim DBAdapter1 As New OleDbDataAdapter(SQL1, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Monthly")
        '�����`�p
        SQL1 = "SELECT COUNT(*) AS MTotal FROM F_ManufOutSheet "
        SQL1 = SQL1 + "WHERE (" + wTColumn + " > DATEADD(d, 0, '" + wMDate + "')) AND (" + wTColumn + " < DATEADD(m, 1, '" + wMDate + "')) "
        SQL1 = SQL1 + SQL2
        Dim DBAdapter2 As New OleDbDataAdapter(SQL1, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "Total")

        wMTotal = DBDataSet1.Tables("Total").Rows(0).Item("MTotal")
        If wMTotal = 0 Then
            DataGrid1.Visible = False
            Exit Sub
        Else
            DataGrid1.Visible = True
        End If
        '�z��~�����
        SQL1 = "SELECT Case " + DSOrB.SelectedValue + " When '' Then '�L�W��' Else " + DSOrB.SelectedValue + " End AS SOrB, COUNT(*) AS CumulativeCount FROM F_ManufOutSheet "
        SQL1 = SQL1 + "WHERE (" + wTColumn + " > DATEADD(d, 0, '" + wCDate + "')) AND (" + wTColumn + " < DATEADD(m, 1, '" + wMDate + "')) "
        SQL1 = SQL1 + SQL2
        SQL1 = SQL1 + "GROUP BY " + DSOrB.SelectedValue
        Dim DBAdapter3 As New OleDbDataAdapter(SQL1, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "Cumulative")
        '�~����`�p
        SQL1 = "SELECT COUNT(*) AS CTotal FROM F_ManufOutSheet "
        SQL1 = SQL1 + "WHERE (" + wTColumn + " > DATEADD(d, 0, '" + wCDate + "')) AND (" + wTColumn + " < DATEADD(m, 1, '" + wMDate + "')) "
        SQL1 = SQL1 + SQL2
        Dim DBAdapter4 As New OleDbDataAdapter(SQL1, OleDbConnection1)
        DBAdapter4.Fill(DBDataSet1, "CumulativeTotal")

        wCTotal = DBDataSet1.Tables("CumulativeTotal").Rows(0).Item("CTotal")

        'Call Stored Procedure Create WorkTable
        SQL = "Exec sp_Temp_InqManufOutList '" & WorkTable & "'"
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQL
        OleDBCommand1.ExecuteNonQuery()

        For i = 0 To DBDataSet1.Tables("Cumulative").Rows.Count - 1
            MonthlyCount = 0
            MonthlyRate = "0"
            CumulativeCount = CInt(DBDataSet1.Tables("Cumulative").Rows(i).Item("CumulativeCount"))
            '����ƬO�_�]�t�b�~��Ƥ�,�Y������,�h�N���ƨ��X
            For j = 0 To DBDataSet1.Tables("Monthly").Rows.Count - 1
                If DBDataSet1.Tables("Cumulative").Rows(i).Item("SOrB") = DBDataSet1.Tables("Monthly").Rows(j).Item("SOrB") Then
                    MonthlyCount = CInt(DBDataSet1.Tables("Monthly").Rows(j).Item("MonthlyCount"))
                    Exit For
                End If
            Next j
            '�p���~�`�v
            If MonthlyCount <> 0 Then MonthlyRate = Format(MonthlyCount / wMTotal * 100, "##0.00")
            '�p��~�~�`�v
            CumulativeRate = Format(CumulativeCount / wCTotal * 100, "##0.00")

            SQL = "Insert into " & WorkTable & " "
            SQL = SQL + "( "
            SQL = SQL + "SOrB, MonthlyCount, MonthlyRate, CumulativeCount, CumulativeRate, CreateUser, CreateTime "
            SQL = SQL + ")  "
            SQL = SQL + "VALUES( "
            SQL = SQL + " '" + DBDataSet1.Tables("Cumulative").Rows(i).Item("SOrB") + "', "
            SQL = SQL + " '" + CStr(MonthlyCount) + "', "
            SQL = SQL + " '" + CStr(MonthlyRate) + "', "
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Cumulative").Rows(i).Item("CumulativeCount")) + "', "
            SQL = SQL + " '" + CStr(CumulativeRate) + "', "
            SQL = SQL + " '" + Request.Cookies("UserID").Value + "', "     '�@����
            SQL = SQL + " '" + NowDateTime + "' "       '�@���ɶ�
            SQL = SQL + " ) "

            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDBCommand1.ExecuteNonQuery()
        Next i

        SQL = "SELECT SOrB, MonthlyCount, MonthlyRate, CumulativeCount, CumulativeRate "
        SQL = SQL + "FROM " + WorkTable + " "
        SQL = SQL + "Order by SOrB"

        Dim DBAdapter5 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter5.Fill(DBDataSet2, "Header")
        DBTable1 = DBDataSet2.Tables("Header")

        DBDataRow = DBTable1.NewRow()
        DBDataRow.Item(0) = DYear.SelectedValue + "�~" + DMonth.SelectedValue + "��"
        DBTable1.Rows.InsertAt(DBDataRow, 0)    '���J�~��b�Ĥ@�C

        DBDataRow = DBTable1.NewRow()
        DBDataRow.Item(0) = "�`�p"
        DBDataRow.Item(1) = wMTotal
        DBDataRow.Item(2) = "100"
        DBDataRow.Item(3) = wCTotal
        DBDataRow.Item(4) = "100"
        DBTable1.Rows.Add(DBDataRow)            '�s�W�`�p��Ʀb�̫�@�C

        DataGrid1.DataSource = DBTable1
        DataGrid1.DataBind()

        DataGrid1.Items(0).Cells(0).ColumnSpan = 5  '�N�Ĥ@�C��ܬ�����
        For i = 1 To 4
            DataGrid1.Items(0).Cells(i).Visible = False '���ë᭱4��
        Next i

        OleDbConnection1.Close()
    End Sub

    '*****************************************************************
    '**
    '**     ��Excel�@�ε{��
    '**
    '*****************************************************************
    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '�{���O���P

        ''pFormNo = DFormName.SelectedValue
        DataGrid1.AllowPaging = False   '�{���O���P
        CreateWorkTable()                      '�{���O���P

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
