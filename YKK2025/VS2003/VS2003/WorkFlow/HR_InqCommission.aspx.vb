Imports System.Data
Imports System.Data.OleDb

Public Class HR_InqCommission
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
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
    Dim wLevel As String = ""
    Dim wDivision As String = ""
    Dim wName As String = ""
    Dim NowDateTime As String       '現在日期時間

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "HR_InqCommission.aspx"

        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            SetSearchItem()
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
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")

    End Sub

    Sub SetSearchItem()
        Dim i As Integer = 0
        Dim wSts As String
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        '**** 設定 部門/姓名
        OleDbConnection1.Open()

        '取得篩選權限
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
            '取得個人資訊
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
                '取得所指定部門
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
                wName = "申請者"
            Else
                '取得全部部門
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
                    wName = "申請者"
                End If
            End If
        End If

        DBDataSet1.Clear()
        '表單
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

        '日期
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
            DataGrid1.Columns.Item(0).HeaderText = "No."
            DataGrid1.Columns.Item(1).HeaderText = "狀態"
            DataGrid1.Columns.Item(2).HeaderText = "申請者"
            DataGrid1.Columns.Item(3).HeaderText = "申請日"

            SQL = "SELECT "
            SQL = SQL + "Case " + wTableName + ".No When '' Then '未編號' Else " + wTableName + ".No End As Field1, "
            SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '核定中' When '1' Then '完成' When '2' Then '取消' Else '取消' End As Field2, "
            SQL = SQL + "V_WaitHandle_01.HRWDivName + '--' + "
            SQL = SQL + wTableName + ".Name as Field3, "
            SQL = SQL + "Convert(VARCHAR(10), V_WaitHandle_01.ApplyTime, 111) as Field4, "

            If pFormNo = "001001" Then
                DataGrid1.Columns.Item(4).HeaderText = "類型"
                DataGrid1.Columns.Item(5).HeaderText = "加班日期"
                DataGrid1.Columns.Item(6).HeaderText = "伙食"

                SQL = SQL + wTableName + ".DateType as Field5, "
                SQL = SQL + "Convert(VARCHAR(10), " + wTableName + ".OverTimeDate, 111) as Field6, "
                SQL = SQL + wTableName + ".Food as Field7, "
            End If
            If pFormNo = "001002" Then
                DataGrid1.Columns.Item(4).HeaderText = "假別"
                DataGrid1.Columns.Item(5).HeaderText = "請假日期"
                DataGrid1.Columns.Item(6).HeaderText = "天數"

                SQL = SQL + wTableName + ".Vacation as Field5, "
                SQL = SQL + "Convert(VARCHAR(10), " + wTableName + ".AStartDate, 111) + ' ' + str(AStartH) + ':00' + '~' + "
                SQL = SQL + "Convert(VARCHAR(10), " + wTableName + ".AEndDate, 111)   + ' ' + str(AEndH) + ':00' as Field6, "
                SQL = SQL + wTableName + ".ADays as Field7, "
            End If
            If pFormNo = "001003" Then
                DataGrid1.Columns.Item(4).HeaderText = "目的地"
                DataGrid1.Columns.Item(5).HeaderText = "外出日期"
                DataGrid1.Columns.Item(6).HeaderText = "日時"

                SQL = SQL + wTableName + ".Place as Field5, "
                SQL = SQL + "Convert(VARCHAR(10), " + wTableName + ".AStartDate, 111) + ' ' + str(AStartH) + ':00' + '~' + "
                SQL = SQL + "Convert(VARCHAR(10), " + wTableName + ".AEndDate, 111)   + ' ' + str(AEndH) + ':00' as Field6, "
                SQL = SQL + "str(" + wTableName + ".ADay) + '日' + str(" + wTableName + ".AHour) + '時' as Field7, "
            End If
            If pFormNo = "001004" Then
                DataGrid1.Columns.Item(4).HeaderText = "缺勤日期"
                DataGrid1.Columns.Item(5).HeaderText = "補工日期"
                DataGrid1.Columns.Item(6).HeaderText = "時數"
                SQL = SQL + "Convert(VARCHAR(10), " + wTableName + ".WorkDate, 111) + ' ' + str(" + wTableName + ".wStartH) +  ':' + str(" + wTableName + ".wStartM) + ':00' + '~' + str(" + wTableName + ".wEndH) +  ':' + str(" + wTableName + ".wEndM) + ':00' as Field5, "
                SQL = SQL + "Convert(VARCHAR(10), " + wTableName + ".AWorkDate, 111) + ' ' + str(" + wTableName + ".AStartH) +  ':' + str(" + wTableName + ".AStartM) + ':00' + '~' + str(" + wTableName + ".AEndH) +  ':' + str(" + wTableName + ".AEndM) + ':00' as Field6, "
                SQL = SQL + "str(" + wTableName + ".AH) + '時' + str(" + wTableName + ".AM) + '分' as Field7, "
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
            '表單
            If DFormName.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".FormNo = '" + DFormName.SelectedValue + "'"
            End If

            '狀態
            If DProgress.SelectedValue <> "ALL" Then
                If DProgress.SelectedValue = "1" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '0'  "
                End If
                If DProgress.SelectedValue = "2" Then
                    SQL = SQL + " And " + wTableName + ".Sts <>  '0'  "
                End If
            End If

            '完成狀態
            If DSts.SelectedValue <> "ALL" Then
                If DSts.SelectedValue <> "3" Then
                    SQL = SQL + " And " + wTableName + ".Sts = '" + DSts.SelectedValue + "'"
                Else
                    SQL = SQL + " And (" + wTableName + ".Sts = '2' " + " or " + wTableName + ".Sts = '3') "
                End If
            End If

            '部門
            If DDivision.SelectedValue <> "ALL" Then
                SQL = SQL + " And V_WaitHandle_01.HRWDivName = '" + DDivision.SelectedValue + "'"
            End If
            '申請者
            If DName.Text <> "申請者" And DName.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".Name Like '%" + DName.Text + "%'"
            End If
            'No
            If DNo.Text <> "No." And DNo.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".No Like '%" + DNo.Text + "%'"
            End If
            '申請日
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
            DName.Text = "申請者"
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
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '程式別不同

        pFormNo = DFormName.SelectedValue
        DataGrid1.AllowPaging = False   '程式別不同
        DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=Commission_List.xls")     '程式別不同
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        DataGrid1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        DataGrid1.AllowPaging = wAllowPaging        '程式別不同
    End Sub

End Class
