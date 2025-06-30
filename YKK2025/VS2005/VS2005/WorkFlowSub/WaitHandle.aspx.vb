Imports System.Data
Imports System.Data.OleDb

Partial Class WaitHandle
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        Response.Cookies("PGM").Value = "WaitHandle.aspx"
        '待處理連結
        LWaitPage.NavigateUrl = "http://10.245.1.10/WorkFlow/WaitHandle.aspx?pUserID=" & Request.QueryString("pUserID")

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
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        If Request.QueryString("pformno") <> "" Then '新程式代表單號

            '表單
            SQL = "SELECT * FROM M_form "
            SQL = SQL + " Where formno = '" + Request.QueryString("pformno") + "'"

            OleDbConnection1.Open()
            DFormName.Items.Clear()
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
        Else

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
        End If


        If Request.QueryString("pflowtype") <> "" Then '新程式代表單號

            If Request.QueryString("pflowtype") = 0 Then
                DFlowType.SelectedIndex = 1
            ElseIf Request.QueryString("pflowtype") = 1 Then
                DFlowType.SelectedIndex = 2
            ElseIf Request.QueryString("pflowtype") = 2 Then
                DFlowType.SelectedIndex = 0
            End If


        Else

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
        End If

        If Request.QueryString("pformno") <> "" Then '新程式代表單號
            DApplyName.Items.Add("ALL")
        Else
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
        End If


      

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
        Dim StartDate As String = DateTime.Now.ToString("yyyyMMdd")
        
        SQL = "SELECT "
        SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
        SQL = SQL + "Case  When FormNo>='001001' and (FormNo  not in ('001051','001052','001053', '003109') ) Then Division "
        SQL = SQL + "      When FormNo  in ('001051','001052','001053') and  ( Division = overtimeDivision  or   isnull(overtimeDivision,'') ='')   then  Division "
        SQL = SQL + "      When FormNo  in ('001051','001052','001053') Then Division+'('+ overtimeDivision  +')'  "
        SQL = SQL + "      When FormNo  in ('003109') then QCInf "
        SQL = SQL + "      Else MapNo End As MapNo, "

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

        SQL = SQL + " As Description,  "

        'SQL = SQL +URL + '" & "&pUserID=" & Request.QueryString("pUserID") & "'  As URL 

        'jessica 20220729 將 URL　ItemRegisterSheet_01　取代成　ItemRegisterSheet_03
        ' 20220801以後 換到 ItemRegisterSheet_03
        If StartDate >= "20220729" Then
            SQL = SQL + " case When formno ='001151' then replace(URL,'ItemRegisterSheet_01','ItemRegisterSheet_03')+ '" & "&pUserID=" & Request.QueryString("pUserID") & "'"

            SQL = SQL + " else URL + '" & "&pUserID=" & Request.QueryString("pUserID") & "' end As URL "
        Else
            SQL = SQL + " URL+ '" & "&pUserID=" & Request.QueryString("pUserID") & "'  As URL "

        End If


        SQL = SQL + " FROM  V_WaitHandle_Approve "

        SQL = SQL + " Where Active = '1' "
        SQL = SQL + "  And (Sts = '0'  Or  Sts = '4') "
        SQL = SQL + "  And DecideID = '" + Request.QueryString("pUserID") + "'"
        'Modify-End
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

    Protected Sub Refresh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Refresh.Click
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
