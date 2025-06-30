Imports System.Data
Imports System.Data.OleDb

Public Class WaitHandle
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DSts As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFlowType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DApplyName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSortKey As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSort As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DKey1 As System.Web.UI.WebControls.TextBox

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "WaitHandle.aspx"

        If Not Me.IsPostBack Then
            CheckAuthority()
            SetSearchItem()
            WaitHandle_DataList()
        End If
    End Sub

    Sub CheckAuthority()
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
        SQL = "SELECT FormName, FormNo FROM V_WaitHandle_Approve "
        SQL = SQL + " Where Active = '1' "
        SQL = SQL + "   And FormNo <= '900000' "
        SQL = SQL + "   And (Sts = '0'  Or  Sts = '4') "
        SQL = SQL + "   And DecideID = '" + Request.QueryString("pUserID") + "'"
        SQL = SQL + " Group by FormName, FormNo "
        SQL = SQL + " Order by FormName, FormNo "

        OleDbConnection1.Open()
        DFormName.Items.Clear()
        DFormName.Items.Add("ALL")
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "WaitHandle")
        DBTable1 = DBDataSet1.Tables("WaitHandle")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("FormName")
            ListItem1.Value = DBTable1.Rows(i).Item("FormNo")
            DFormName.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()

        '申請人
        SQL = "SELECT ApplyName, ApplyID FROM V_WaitHandle_Approve "
        SQL = SQL + " Where Active = '1' "
        SQL = SQL + "   And (Sts = '0'  Or  Sts = '3') "
        SQL = SQL + "   And DecideID = '" + Request.QueryString("pUserID") + "'"
        SQL = SQL + "Group by ApplyName, ApplyID "
        SQL = SQL + "Order by ApplyName, ApplyID "

        OleDbConnection1.Open()
        DApplyName.Items.Clear()
        DApplyName.Items.Add("ALL")
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "WaitHandle_01")
        DBTable1 = DBDataSet1.Tables("WaitHandle_01")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("ApplyName")
            ListItem1.Value = DBTable1.Rows(i).Item("ApplyID")
            DApplyName.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()

        'Key1
        DKey1.Text = ""
        If DFormName.SelectedValue = "002002" Then
            DKey1.Enabled = True
        Else
            DKey1.Enabled = False
        End If
    End Sub


    Sub WaitHandle_DataList()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定


        SQL = "SELECT "
        SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
        SQL = SQL + "Case  When FormNo>='001001' and (FormNo  not in ('001051','001052','001053','003109' ) ) Then Division "
        SQL = SQL + "      When FormNo  in ('001051','001052','001053') and  ( Division = overtimeDivision  or   isnull(overtimeDivision,'') ='')   then  Division "
        SQL = SQL + "      When FormNo  in ('001051','001052','001053') Then Division+'('+ overtimeDivision  +')'  "
        SQL = SQL + "      When FormNo  ='003109' then QCInf "
        SQL = SQL + "      Else MapNo End As  MapNo, "

        SQL = SQL + "Case  When FormNo='002002' "
        SQL = SQL + "  then "
        SQL = SQL + "    Case  When No='' or No IS NULL Then '未編號' Else No + '('+ IsNull(CodeNo,'') +'/'+ IsNull(DevNo,'') +')' End "
        SQL = SQL + "  else "
        SQL = SQL + "    Case  When No='' or No IS NULL Then '未編號' Else No End "
        SQL = SQL + "End As No, "

        SQL = SQL + "Case  When FormNo='001054' "
        SQL = SQL + "  then TimeOff "
        SQL = SQL + "  else StepNameDesc "
        SQL = SQL + "End As StepNameDesc, "
        SQL = SQL + "StsDesc, FormName, FlowTypeDesc, ApplyName, AgentName, "

        'SQL = SQL + "StsDesc, FormName, FlowTypeDesc, ApplyName, StepNameDesc, AgentName, "

        SQL = SQL + "'申請時間：[' + Convert(VarChar, ApplyTime, 20) + '], ' + "
        SQL = SQL + "'收件時間：[' + Convert(VarChar, ReceiptTime, 20) + '], ' + "
        SQL = SQL + "'閱讀期限：[' + Convert(VarChar, ReadTimeLimit, 20) + '], ' + "
        SQL = SQL + "'首次閱讀：[' + FirstReadTimeDesc + '], ' + "
        SQL = SQL + "'最後閱讀：[' + LastReadTimeDesc  + '], ' + "
        SQL = SQL + "'預定開始：[' + Convert(VarChar, BStartTime, 20) + '], ' + "
        '2012/4/3 納期對應-Joy
        'Modify-Start
        'SQL = SQL + "'預定完成：[' + Convert(VarChar, BEndTime, 20) + '], ' + "
        'SQL = SQL + "'實際開始：[' + Convert(VarChar, AStartTime, 20) + '] '  "
        SQL = SQL + "'預定完成：[' + Convert(VarChar, BEndTime, 20) + '] '  "
        'Modify-End
        'SQL = SQL + " As Description, URL "
        SQL = SQL + " As Description,"

        SQL = SQL + " case When formno ='001151' then replace(URL,'ItemRegisterSheet_01','ItemRegisterSheet_03')+ '" & "&pUserID=" & Request.QueryString("pUserID") & "'"

        SQL = SQL + " else URL + '" & "&pUserID=" & Request.QueryString("pUserID") & "' end As URL "

        ' URL + '" & "&pUserID=" & Request.QueryString("pUserID") & "' As URL "

        SQL = SQL + " FROM V_WaitHandle_Approve "
        SQL = SQL + " Where Active = '1' "
        SQL = SQL + "  And (Sts = '0'  Or  Sts = '4') "
        SQL = SQL + "  And DecideID = '" + Request.QueryString("pUserID") + "'"
        '
        'ADD-Start 代理簽核
        SQL = SQL + " And FORMNO + '-' + CONVERT(VARCHAR,FORMSNO) + '-' + CONVERT(VARCHAR,STEP) "
        SQL = SQL + " Not in "
        SQL = SQL + " (select FORMNO + '-' + CONVERT(VARCHAR,FORMSNO) + '-' + CONVERT(VARCHAR,STEP) from Q_AgentApprov) "
        'ADD-End
        '
        '表單
        If DFormName.SelectedValue <> "ALL" Then
            SQL = SQL + " And   FormNo = '" + DFormName.SelectedValue + "'"
        End If
        '狀態
        If DSts.SelectedValue <> "ALL" Then
            SQL = SQL + " And   Sts = '" + DSts.SelectedValue + "'"
        End If
        '類別
        If DFlowType.SelectedValue <> "ALL" Then
            SQL = SQL + " And   FlowType = '" + DFlowType.SelectedValue + "'"
        End If
        '委託人
        If DApplyName.SelectedValue <> "ALL" Then
            SQL = SQL + " And   ApplyID = '" + DApplyName.SelectedValue + "'"
        End If
        'Key1
        If DKey1.Text <> "" Then
            SQL = SQL + " And ( CodeNo Like '%" + DKey1.Text + "%' or "
            SQL = SQL + "       DevNo Like '%" + DKey1.Text + "%' ) "
        End If

        '20221026 JESSICA fas14 no show 006001  
        If Request.QueryString("pUserID") = "fas014" Then
            SQL = SQL + " And FORMNO <> '006001'"
        End If

        'OLD對應-Start 08-03-12
        'SQL = SQL + " Order by No Desc "
        'OLD對應-End

        '表單單號
        If DSortKey.SelectedValue <> "DFormSno" Then
            SQL = SQL + " Order by " + DSortKey.SelectedValue + " " + DSort.SelectedValue
        Else
            SQL = SQL + " Order by FormNo " + DSort.SelectedValue + ", FormSno " + DSort.SelectedValue
        End If
        'DataGrid1.Columns(3).HeaderText = "加班單位"
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "WaitHandle")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()
        OleDbConnection1.Close()

        ' Response.Write(SQL)
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        WaitHandle_DataList()
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataGrid1.DataBind()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        WaitHandle_DataList()
    End Sub

    Private Sub DFormName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged
        'Key1
        DKey1.Text = ""
        If DFormName.SelectedValue = "002002" Then
            DKey1.Enabled = True
        Else
            DKey1.Enabled = False
        End If
    End Sub
End Class
