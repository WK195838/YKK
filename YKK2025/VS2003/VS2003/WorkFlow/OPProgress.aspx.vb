Imports System.Data
Imports System.Data.OleDb

Public Class OPProgress
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DSts As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DActive As System.Web.UI.WebControls.DropDownList
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
    Dim pFormNo As String
    Dim pSts As String
    Dim NowDateTime As String       '現在日期時間

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "OPProgress.aspx"

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
        pFormNo = Request.QueryString("pFormNo")    '表單號碼
        pSts = Request.QueryString("pSts")
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub SetSearchItem()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        '表單
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

        For i = 0 To DActive.Items.Count - 1
            If DActive.Items.Item(i).Value = "開發中" Then
                DActive.Items.Item(i).Selected = True
            End If
        Next

        For i = 0 To DSts.Items.Count - 1
            If DSts.Items.Item(i).Value = pSts Then
                DSts.Items.Item(i).Selected = True
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
            SQL = "SELECT "
            'SQL = SQL + wTableName + ".No as No, "
            SQL = SQL + "Case " + wTableName + ".No When '' Then '未編號' Else " + wTableName + ".No End As No, "
            SQL = SQL + wTableName + ".Division + '--' + "
            SQL = SQL + wTableName + ".Person as Person, "
            SQL = SQL + "V_WaitHandle_01.StepName as StepName, "
            SQL = SQL + "V_WaitHandle_01.DecideName as StepPerson, "
            SQL = SQL + "Convert(VarChar, V_WaitHandle_01.BStartTime, 20)  as BStartTime, "
            SQL = SQL + "Convert(VarChar, V_WaitHandle_01.BEndTime, 20) as BEndTime, "
            SQL = SQL + "Convert(VarChar, V_WaitHandle_01.AStartTime, 20) as AStartTime, "
            SQL = SQL + "Convert(VarChar, V_WaitHandle_01.ApplyTime, 20)  as ApplyTime, "
            SQL = SQL + "'收件時間：[' + Convert(VarChar, V_WaitHandle_01.ReceiptTime, 20) + '], ' +  "
            SQL = SQL + "'閱讀期限：[' + Convert(VarChar, V_WaitHandle_01.ReadTimeLimit, 20) + '], ' + "
            SQL = SQL + "'首次閱讀：[' + Convert(VarChar, V_WaitHandle_01.FirstReadTimeDesc, 20) + '], ' + "
            SQL = SQL + "'最後閱讀：[' + Convert(VarChar, V_WaitHandle_01.LastReadTimeDesc, 20) + ']  ' as ReadTime, "

            SQL = SQL + "'OPSchedule.aspx?' + "
            SQL = SQL + "'pWorkID='   + V_WaitHandle_01.WorkID + "
            SQL = SQL + "'&pFormNo='  + V_WaitHandle_01.FormNo + "
            SQL = SQL + "'&pType='    + str(V_WaitHandle_01.FlowType,Len(V_WaitHandle_01.FlowType)) As SCURL,  "

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
            SQL = SQL + "Where " + wTableName + ".Sts  = '0' "
            SQL = SQL + "  And V_WaitHandle_01.FormSno > '2000' "
            SQL = SQL + "  And V_WaitHandle_01.SeqNo   = '1' "
            '表單
            If DFormName.SelectedValue <> "ALL" Then
                SQL = SQL + " And   V_WaitHandle_01.FormNo = '" + DFormName.SelectedValue + "'"
            End If
            '狀態
            SQL = SQL + " And   V_WaitHandle_01.Active = '" + DActive.SelectedValue + "'"
            '延遲
            If DSts.SelectedValue = "Delay" Then
                SQL = SQL + " And   V_WaitHandle_01.Delay = '1' "
            End If
            '未閱讀
            If DSts.SelectedValue = "ReadDelay" Then
                SQL = SQL + " And   V_WaitHandle_01.ReadDelay = '1' "
            End If
            '正確
            If DSts.SelectedValue = "Normal" Then
                SQL = SQL + " And   V_WaitHandle_01.Delay = '0' "
                SQL = SQL + " And   V_WaitHandle_01.ReadDelay = '0' "
            End If
            'FlowType
            If DFlowType.SelectedValue <> "ALL" Then
                SQL = SQL + " And   V_WaitHandle_01.FlowType = '" + DFlowType.SelectedValue + "'"
            End If
            'Sort
            SQL = SQL + " Order by No desc "

            OleDbConnection1.Open()
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "WaitHandle")
            DataGrid1.DataSource = DBDataSet2
            DataGrid1.DataBind()
            OleDbConnection1.Close()
        End If
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
