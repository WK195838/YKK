Imports System.Data
Imports System.Data.OleDb

Public Class OPScheduleList
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DFlowType As System.Web.UI.WebControls.DropDownList

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region


    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim pWorkID As String
    Dim pFormNo As String
    Dim pType As String
    Dim NowDateTime As String       '現在日期時間

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "OPSchedule.aspx"

        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            SetSearchItem()
            DataList()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '現在日時
        pWorkID = Request.QueryString("pWorkID")
        pFormNo = Request.QueryString("pFormNo")    '表單號碼
        pType = Request.QueryString("pType")
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub SetSearchItem()
        Dim i As Integer = 0
        For i = 0 To DFlowType.Items.Count - 1
            If DFlowType.Items.Item(i).Value = pType Then
                DFlowType.Items.Item(i).Selected = True
            End If
        Next
    End Sub

    Sub DataList()
        Dim wTableName As String = ""
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        'OleDbConnection1.Open()
        'SQL = "SELECT FormNo, FormName, TableName1 FROM M_Form "
        'SQL = SQL + " Where Active  =  '1' "
        'SQL = SQL + "   And FormNo = '" + pFormNo + "'"
        'Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter1.Fill(DBDataSet1, "Flow")
        'If DBDataSet1.Tables("Flow").Rows.Count > 0 Then
        'wTableName = "F_" + DBDataSet1.Tables("Flow").Rows(0).Item("TableName1")
        'End If
        'OleDbConnection1.Close()

        'If wTableName <> "" Then
        SQL = "SELECT "
        SQL = SQL + "FormName + '(' + str(FormSno, Len(FormSno)) +')' as No, "
        SQL = SQL + "ApplyName as Person, "
        SQL = SQL + "StepName as StepName, "
        SQL = SQL + "DecideName as StepPerson, "
        SQL = SQL + "Convert(VarChar, BStartTime, 20) + '~' + Convert(VarChar, BEndTime, 20) As BOPTime, "
        SQL = SQL + "Convert(VarChar, AStartTime, 20) as AStartTime, "
        SQL = SQL + "Convert(VarChar, ApplyTime, 20)  as ApplyTime, "
        SQL = SQL + "'收件時間：[' + Convert(VarChar, ReceiptTime, 20) + '], ' +  "
        SQL = SQL + "'閱讀期限：[' + Convert(VarChar, ReadTimeLimit, 20) + '], ' + "
        SQL = SQL + "'首次閱讀：[' + Convert(VarChar, FirstReadTime, 20) + '], ' + "
        SQL = SQL + "'最後閱讀：[' + Convert(VarChar, LastReadTime, 20) + ']  ' as ReadTime "
        SQL = SQL + "FROM V_WaitHandle_01 "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And WorkID = '" + pWorkID + "' "
        'FlowType
        If DFlowType.SelectedValue <> "ALL" Then
            SQL = SQL + " And   FlowType = '" + DFlowType.SelectedValue + "'"
        End If
        'Sort
        SQL = SQL + " Order by BStartTime "

        'SQL = "SELECT "
        'SQL = SQL + wTableName + ".No as No, "
        'SQL = SQL + wTableName + ".Person as Person, "
        'SQL = SQL + "V_WaitHandle_01.StepName as StepName, "
        'SQL = SQL + "V_WaitHandle_01.DecideName as StepPerson, "
        'SQL = SQL + "Convert(VarChar, V_WaitHandle_01.BStartTime, 20) + '~' + Convert(VarChar, V_WaitHandle_01.BEndTime, 20) As BOPTime, "
        'SQL = SQL + "Convert(VarChar, V_WaitHandle_01.AStartTime, 20) as AStartTime, "
        'SQL = SQL + "Convert(VarChar, V_WaitHandle_01.ApplyTime, 20)  as ApplyTime, "
        'SQL = SQL + "'收件時間：[' + Convert(VarChar, V_WaitHandle_01.ReceiptTime, 20) + '], ' +  "
        'SQL = SQL + "'閱讀期限：[' + Convert(VarChar, V_WaitHandle_01.ReadTimeLimit, 20) + '], ' + "
        'SQL = SQL + "'首次閱讀：[' + Convert(VarChar, V_WaitHandle_01.FirstReadTime, 20) + '], ' + "
        'SQL = SQL + "'最後閱讀：[' + Convert(VarChar, V_WaitHandle_01.LastReadTime, 20) + ']  ' as ReadTime "
        'SQL = SQL + "FROM " + wTableName + " "
        'SQL = SQL + "Left Outer Join V_WaitHandle_01 ON " + wTableName + ".FormNo=V_WaitHandle_01.FormNo "
        'SQL = SQL + "                               And " + wTableName + ".FormSno=V_WaitHandle_01.FormSno "

        'SQL = SQL + "Where " + wTableName + ".Sts = '0' "
        'SQL = SQL + "  And V_WaitHandle_01.Active = '1' "
        'SQL = SQL + "  And V_WaitHandle_01.WorkID = '" + pWorkID + "' "
        'FlowType
        'If DFlowType.SelectedValue <> "ALL" Then
        'SQL = SQL + " And   V_WaitHandle_01.FlowType = '" + DFlowType.SelectedValue + "'"
        'End If
        'Sort
        'SQL = SQL + " Order by BStartTime "

        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "WaitHandle")
        DataGrid1.DataSource = DBDataSet2
        DataGrid1.DataBind()
        OleDbConnection1.Close()

        'End If
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        DataList()
    End Sub

    Private Sub DataGrid1_PageIndexChanged1(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged

        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataList()
    End Sub

End Class
