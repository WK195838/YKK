Imports System.Data
Imports System.Data.OleDb

Partial Class WaitHandle_IRW
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "WaitHandle_IRW.aspx"
        Server.ScriptTimeout = 900  ' 設定逾時時間

        If Not Me.IsPostBack Then
            CheckAuthority()
            SetSearchItem()
            'WaitHandle_DataList()
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
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        '表單
        'SQL = "SELECT FormName, FormNo FROM V_WaitHandle_01B "
        SQL = "SELECT FormName, FormNo FROM M_Form "
        SQL = SQL + " Where Active = '1' "
        SQL = SQL + "   And FormNo >= '001151' "
        SQL = SQL + "   And FormNo <= '001155' "
        SQL = SQL + " Order by FormName, FormNo "

        OleDbConnection1.Open()
        DFormName.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "WaitHandle")
        DBTable1 = DBDataSet1.Tables("WaitHandle")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("FormName")
            ListItem1.Value = DBTable1.Rows(i).Item("FormNo")
            If ListItem1.Value = "001151" Then
                ListItem1.Selected = True
            End If
            DFormName.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()
        '
        DApplyName.Text = "申請者"
        DNo.Text = "No"
    End Sub


    Sub WaitHandle_DataList()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        Dim StartDate As String = DateTime.Now.ToString("yyyyMMdd")

        SQL = "SELECT top 30 "
        SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
        SQL = SQL + "Case  When FormNo>='001151' and (FormNo  not in ('001051','001052','001053' ) ) Then Division "
        SQL = SQL + "      When FormNo  in ('001051','001052','001053') and  ( Division = overtimeDivision  or   isnull(overtimeDivision,'') ='')   then  Division "
        SQL = SQL + "      When FormNo  in ('001051','001052','001053') Then Division+'('+ overtimeDivision  +')'  "
        SQL = SQL + "      Else MapNo End As MapNo, "
        SQL = SQL + "Case  When No='' or No IS NULL Then '未編號' Else No End As No, "
        '
        SQL = SQL + "StsDesc, FormName, FlowTypeDesc, ApplyName, StepNameDesc, AgentName, "
        SQL = SQL + "'申請時間：[' + Convert(VarChar, ApplyTime, 20) + '], ' + "
        SQL = SQL + "'收件時間：[' + Convert(VarChar, ReceiptTime, 20) + '], ' + "
        SQL = SQL + "'閱讀期限：[' + Convert(VarChar, ReadTimeLimit, 20) + '], ' + "
        SQL = SQL + "'首次閱讀：[' + FirstReadTimeDesc + '], ' + "
        SQL = SQL + "'最後閱讀：[' + LastReadTimeDesc  + '], ' + "
        SQL = SQL + "'預定開始：[' + Convert(VarChar, BStartTime, 20) + '], ' + "
        SQL = SQL + "'預定完成：[' + Convert(VarChar, BEndTime, 20) + '], ' + "
        SQL = SQL + "'實際開始：[' + Convert(VarChar, AStartTime, 20) + '] '  "

        'jessica 20220729 將 URL　ItemRegisterSheet_01　取代成　ItemRegisterSheet_03
        ' 20220801以後 換到 ItemRegisterSheet_03

        'SQL = SQL + " As Description, URL + '" & "&pUserID=" & Request.QueryString("pUserID") & "' As URL "
        SQL = SQL + " As Description,  "
        If StartDate >= "20220729" Then
            SQL = SQL + " case When formno ='001151' then replace(URL,'ItemRegisterSheet_01','ItemRegisterSheet_03')+ '" & "&pUserID=" & Request.QueryString("pUserID") & "'"

            SQL = SQL + " else URL + '" & "&pUserID=" & Request.QueryString("pUserID") & "' end As URL "
        Else
            SQL = SQL + " URL+ '" & "&pUserID=" & Request.QueryString("pUserID") & "'  As URL "

        End If
        '   
        SQL = SQL + "FROM V_WaitHandle_01B "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And (Sts = '0'  Or  Sts = '4') "
        'SQL = SQL + "  And DecideID = '" + Request.QueryString("pUserID") + "'"
        SQL = SQL + "  And (DecideID = 'irw028' Or DecideID = 'mk028') "
        'Modify-End
        '表單
        If DFormName.SelectedValue <> "ALL" Then
            SQL = SQL + " And   FormNo = '" + DFormName.SelectedValue + "'"
        End If
        '類別
        If DFlowType.SelectedValue <> "ALL" Then
            If DFlowType.SelectedValue = "0" Then
                SQL = SQL + " And   FlowType = '" + DFlowType.SelectedValue + "'"
            Else
                SQL = SQL + " And   FlowType > '0' "
            End If
        End If
        '委託人
        If DApplyName.Text <> "申請者" And DApplyName.Text <> "" Then
            SQL = SQL + " And   ApplyName Like '%" + DApplyName.Text + "%'"
        End If
        'No
        If DNo.Text <> "No" And DNo.Text <> "" Then
            SQL = SQL + " And   No Like '%" + DNo.Text + "%'"
        End If
        SQL = SQL + " Order by No Desc "
        '
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "WaitHandle")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()
        OleDbConnection1.Close()
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


End Class
