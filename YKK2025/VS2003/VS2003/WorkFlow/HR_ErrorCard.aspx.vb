Imports System.Data
Imports System.Data.OleDb

Public Class HR_ErrorCard
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DCardType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DName As System.Web.UI.WebControls.DropDownList

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
    Dim wLevel As String = ""
    Dim NowDateTime As String       '現在日期時間

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "HR_ErrorCard.aspx"

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
        DName.Items.Clear()
        If wLevel = "PERSON" Then
            '取得個人資訊
            SQL = "Select isnull(Name,'') as UserName, isnull(HRWDivName,'') as HRWDivName From V_WorkTime_01 "
            SQL = SQL + "Where UserID='" & Request.QueryString("pUserID") & "' "
            SQL = SQL + "  And Not HRWDivName is Null "
            SQL = SQL + "  And Not Name is Null "
            SQL = SQL + "Order by Name "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "M_Users")
            DBTable1 = DBDataSet1.Tables("M_Users")
            If DBTable1.Rows.Count > 0 Then
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(0).Item("HRWDivName")
                ListItem1.Value = DBTable1.Rows(0).Item("HRWDivName")
                DDivision.Items.Add(ListItem1)
                '姓名
                Dim ListItem2 As New ListItem
                ListItem2.Text = DBTable1.Rows(0).Item("UserName")
                ListItem2.Value = DBTable1.Rows(0).Item("UserName")
                DName.Items.Add(ListItem2)
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
                '取得個人資訊
                DBDataSet1.Clear()
                DName.Items.Add("ALL")
                SQL = "Select isnull(Name,'') as UserName From V_WorkTime_01 "
                SQL = SQL + "Where HRWDivName = '" & DDivision.SelectedValue & "' "
                SQL = SQL + "  And Not HRWDivName is Null "
                SQL = SQL + "  And Not Name is Null "
                SQL = SQL + "Group by Name "
                SQL = SQL + "Order by Name "
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    '姓名
                    Dim ListItem2 As New ListItem
                    ListItem2.Text = DBTable1.Rows(i).Item("UserName")
                    ListItem2.Value = DBTable1.Rows(i).Item("UserName")
                    DName.Items.Add(ListItem2)
                Next
            Else
                '取得全部部門
                If wLevel = "ALL" Then
                    DDivision.Items.Add("ALL")
                    SQL = "Select isnull(HRWDivName,'') as HRWDivName From V_WorkTime_01 "
                    SQL = SQL + "Where Not HRWDivName is Null "
                    SQL = SQL + "  And Not Name is Null "
                    SQL = SQL + "Group by HRWDivName "
                    SQL = SQL + "Order by HRWDivName "
                    Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter2.Fill(DBDataSet1, "M_Users")
                    DBTable1 = DBDataSet1.Tables("M_Users")
                    For i = 0 To DBTable1.Rows.Count - 1
                        Dim ListItem1 As New ListItem
                        ListItem1.Text = DBTable1.Rows(i).Item("HRWDivName")
                        ListItem1.Value = DBTable1.Rows(i).Item("HRWDivName")
                        If i = 0 Then ListItem1.Selected = True
                        DDivision.Items.Add(ListItem1)
                    Next
                    '姓名
                    DBDataSet1.Clear()
                    DName.Items.Add("ALL")
                    SQL = "Select isnull(Name,'') as UserName From V_WorkTime_01  "
                    SQL = SQL + "Where HRWDivName = '" & DDivision.SelectedValue & "' "
                    SQL = SQL + "  And Not HRWDivName is Null "
                    SQL = SQL + "  And Not Name is Null "
                    SQL = SQL + "Group by Name "
                    SQL = SQL + "Order by Name "
                    Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter3.Fill(DBDataSet1, "M_Users")
                    DBTable1 = DBDataSet1.Tables("M_Users")
                    For i = 0 To DBTable1.Rows.Count - 1
                        '姓名
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = DBTable1.Rows(i).Item("UserName")
                        ListItem2.Value = DBTable1.Rows(i).Item("UserName")
                        DName.Items.Add(ListItem2)
                    Next
                End If
            End If
        End If
        OleDbConnection1.Close()
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        SQL = "SELECT "
        SQL = SQL + "Convert(VARCHAR(10), CDate, 111) as DateDesc, WeekDesc, "
        SQL = SQL + "RTrim(Name)+'('+RTrim(EmpID)+')' as NameDesc, "
        SQL = SQL + "HRWDivName, JobName, TimeA, TimeB, Remark, "
        SQL = SQL + "OTDesc, OTURL, VADesc, VAURL, AWayDesc, AWayURL, AWorkDesc, AWorkURL "

        SQL = SQL + "FROM V_WorkTime_01 "
        SQL = SQL + "Where JobID  <> 'ZZ' "
        SQL = SQL + "  And pluralism = '0' "        '非兼職
        '部門
        If DDivision.SelectedValue <> "ALL" Then
            SQL = SQL + " And HRWDivName = '" + DDivision.SelectedValue + "'"
        End If
        '申請者
        If DName.SelectedValue <> "ALL" Then
            SQL = SQL + " And Name = '" + DName.SelectedValue + "'"
        End If
        '狀態
        If DCardType.SelectedValue = "1" Then
            SQL = SQL + " And (Delay1>0 or Early1>0 or OutTime1>0 or TimeA='') "
        Else
            If DCardType.SelectedValue = "2" Then
                SQL = SQL + " And (Delay1=0 And Early1=0 And OutTime1=0 And TimeA<>'') "
            End If
        End If
        SQL = SQL + " And CDate = '" + CStr(DateTime.Now.Today) + "' "
        'Sort
        SQL = SQL + " Order by HRWDivName, JobID, Name "

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
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '程式別不同

        DataGrid1.AllowPaging = False   '程式別不同
        DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=HR_CardDataList.xls")     '程式別不同

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

        DataGrid1.AllowPaging = wAllowPaging        '程式別不同
    End Sub

    Private Sub DDivision_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DDivision.SelectedIndexChanged
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        DName.Items.Clear()
        DName.Items.Add("ALL")
        OleDbConnection1.Open()

        SQL = "Select isnull(name,'') as UserName From V_WorkTime_01 "
        SQL = SQL + "Where HRWDivName = '" & DDivision.SelectedValue & "' "
        SQL = SQL + "  And Not HRWDivName is Null "
        SQL = SQL + "  And Not Name is Null "
        SQL = SQL + "Group by Name "
        SQL = SQL + "Order by Name "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            '姓名
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("UserName")
            ListItem1.Value = DBTable1.Rows(i).Item("UserName")
            DName.Items.Add(ListItem1)
        Next

        OleDbConnection1.Close()
    End Sub

End Class
