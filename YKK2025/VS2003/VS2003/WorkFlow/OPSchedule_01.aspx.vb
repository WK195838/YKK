Imports System.Data
Imports System.Data.OleDb

Public Class OPSchedule_01
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DFlowType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid

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
    Dim pFormNo As String
    Dim pUserID As String
    Dim pWorkID As String
    Dim NowDateTime As String       '現在日期時間

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "OPSchedule_01.aspx"

        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            DataList()
        End If
    End Sub

    Sub SetParameter()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet

        pFormNo = Request.QueryString("pFormNo")
        pUserID = Request.QueryString("pUserID")

        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        If pWorkID = "" Then
            SQL = "Select * From M_Users "
            SQL = SQL & " Where UserID =  '" & pUserID & "'"
            SQL = SQL & "   And Active = '1' "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Users")
            If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then
                pWorkID = DBDataSet1.Tables("M_Users").Rows(0).Item("WorkID")
            Else
                pWorkID = ""
            End If
        End If

        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

    Sub DataList()
        If pWorkID <> "" Then
            Dim SQL As String
            Dim DBDataSet1 As New DataSet
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

            SQL = "SELECT "
            SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
            SQL = SQL + "StsDesc, FormName, FlowTypeDesc, ApplyName, DecideName, StepNameDesc, "
            SQL = SQL + "'流程資訊' As WorkFlow, "
            SQL = SQL + "Convert(VarChar, BStartTime, 20) + '~' + Convert(VarChar, BEndTime, 20) As Description, "
            SQL = SQL + "ViewURL, "

            SQL = SQL + "'BefOPList.aspx?' + "
            SQL = SQL + "'pFormNo='   + FormNo + "
            SQL = SQL + "'&pFormSno=' + str(FormSno,Len(FormSno)) + "
            SQL = SQL + "'&pStep='    + str(Step,Len(Step)) + "
            SQL = SQL + "'&pSeqNo='   + str(SeqNo,Len(SeqNo)) + "
            SQL = SQL + "'&pApplyID=' + ApplyID "
            SQL = SQL + " As OPURL "

            SQL = SQL + "FROM V_WaitHandle_01 "
            SQL = SQL + "Where Active = '1' "
            SQL = SQL + " And   WorkID = '" + pWorkID + "'"
            'FlowType
            If DFlowType.SelectedValue <> "ALL" Then
                SQL = SQL + " And   FlowType = '" + DFlowType.SelectedValue + "'"
            End If
            SQL = SQL + " Order by BStartTime "

            OleDbConnection1.Open()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "WaitHandle")
            DataGrid1.DataSource = DBDataSet1
            DataGrid1.DataBind()
            OleDbConnection1.Close()
        End If
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        DataList()
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As System.Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs)
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataList()
    End Sub


End Class
