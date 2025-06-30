Imports System.Data
Imports System.Data.OleDb

Public Class MPM_InqCommission
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
    Protected WithEvents DName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BEDate As System.Web.UI.WebControls.Button
    Protected WithEvents DEDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSDate As System.Web.UI.WebControls.Button
    Protected WithEvents DSDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents Dtype2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Dtype1 As System.Web.UI.WebControls.DropDownList

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
    Dim NowDateTime As String       '現在日期時間

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "MPM_InqCommission.aspx"

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

        '表單
        SQL = "SELECT FormName, FormNo FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And FormNo >= '004001' And FormNo <= '004099' "
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

        '依賴部門
        DBDataSet1.Clear()
        SQL = "Select   dep_name  as  DivName From M_emp Group by  dep_name  Order by dep_name "
        DDivision.Items.Clear()
        DDivision.Items.Add("ALL")
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("DivName")
            ListItem1.Value = DBTable1.Rows(i).Item("DivName")
            DDivision.Items.Add(ListItem1)
        Next

        'TYPE1
        DBDataSet1.Clear()
        SQL = "select * from M_referp where cat = '4001' and dkey = 'Type1' order  by Unique_id "
        Dtype1.Items.Clear()
        Dtype1.Items.Add("ALL")
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("Data")
            ListItem1.Value = DBTable1.Rows(i).Item("Data")
            DType1.Items.Add(ListItem1)
        Next


        'TYPE2
        DBDataSet1.Clear()
        SQL = "select * from M_referp where cat = '4001' and dkey = 'Type2' order by Unique_id "
        Dtype2.Items.Clear()
        Dtype2.Items.Add("ALL")
        Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter4.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("Data")
            ListItem1.Value = DBTable1.Rows(i).Item("Data")
            Dtype2.Items.Add(ListItem1)
        Next




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

            DataGrid1.Columns.Item(0).HeaderText = "加工編號o"
            DataGrid1.Columns.Item(1).HeaderText = "開發狀態"
            DataGrid1.Columns.Item(2).HeaderText = "依賴者"
            DataGrid1.Columns.Item(3).HeaderText = "圖號"
            DataGrid1.Columns.Item(4).HeaderText = "類別1"
            DataGrid1.Columns.Item(5).HeaderText = "類別2"
            DataGrid1.Columns.Item(6).HeaderText = "收件日"
            DataGrid1.Columns.Item(7).HeaderText = "預訂完成日"

            SQL = "SELECT "
            SQL = SQL + "Case " + wTableName + ".No When '' Then '未編號' Else " + wTableName + ".No End As Field1, "
            SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '開發中' When '1' Then '開發完成' Else '開發取消' End As Field2, "
            SQL = SQL + "V_WaitHandle_01.Division+'-'+V_WaitHandle_01.Clinter as Field3, "
            SQL = SQL + "V_WaitHandle_01.MapNo as Field4, "
            SQL = SQL + wTableName + ".Type1  as Field5, "
            SQL = SQL + wTableName + ".Type2  as Field6, "
            SQL = SQL + wTableName + ".APPDate    as Field7, "
            SQL = SQL + wTableName + ".FinishDate    as Field8, "


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
        '------------------------------------
        SQL = SQL + "Where V_WaitHandle_01.Step  < '5' "
        '表單
        If DFormName.SelectedValue <> "ALL" Then
            SQL = SQL + " And " + wTableName + ".FormNo = '" + DFormName.SelectedValue + "'"
        End If
        '開發狀態
        If DProgress.SelectedValue <> "ALL" Then
            If DProgress.SelectedValue = "1" Then
                SQL = SQL + " And " + wTableName + ".Sts =  '0'  "
            End If
            If DProgress.SelectedValue = "2" Then
                SQL = SQL + " And " + wTableName + ".Sts =  '1'  "
            End If
        End If
        '開發完成狀態
        If DSts.SelectedValue <> "ALL" Then
            SQL = SQL + " And " + wTableName + ".Sts = '" + DSts.SelectedValue + "'"
        End If
        'No
        If DNo.Text <> "委託單No." And DNo.Text <> "" Then
            SQL = SQL + " And " + wTableName + ".No Like '%" + DNo.Text + "%'"
        End If
        '部門
        If DDivision.SelectedValue <> "ALL" Then
            SQL = SQL + " And " + wTableName + ".Division = '" + DDivision.SelectedValue + "'"
        End If

        If DMapNo.Text <> "圖號" Then
            SQL = SQL + " And " + wTableName + ".MapNo Like '%" + DMapNo.Text + "%'"
        End If

        If Dtype1.SelectedValue <> "ALL" Then
            SQL = SQL + " And " + wTableName + ".type1 = '" + Dtype1.SelectedValue + "'"
        End If

        If Dtype2.SelectedValue <> "ALL" Then
            SQL = SQL + " And " + wTableName + ".type2 = '" + Dtype2.SelectedValue + "'"
        End If



        'End If
        '委託No
        '   If DRno.Text <> "委託No" And DRno.Text <> "" Then
        '      SQL = SQL + " And " + wTableName + ".Rno Like '%" + DRno.Text + "%'"
        '  End If
        '開發No
        '  If DDevNo.Text <> "開發No" And DDevNo.Text <> "" Then
        '      SQL = SQL + " And " + wTableName + ".DevNo Like '%" + DDevNo.Text + "%'"
        '  End If
        ' 'Code-No
        ' If DCodeNo.Text <> "Code-No" And DCodeNo.Text <> "" Then
        '     SQL = SQL + " And " + wTableName + ".CodeNo Like '%" + DCodeNo.Text + "%'"
        ' End If
        '完成日
        SQL = SQL + " And " + wTableName + ".CompletedTime between '" + DSDate.Text + " 00:00:00' and '" + DEDate.Text + " 23:59:59'"

        'Sort
        SQL = SQL + " Order by " + wTableName + ".CreateTime desc "

        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "WaitHandle")
        DataGrid1.DataSource = DBDataSet2
        DataGrid1.DataBind()
        OleDbConnection1.Close()


    End Sub

    Sub Search_Item_Attribute()
        DNo.ReadOnly = True
   



        'No
        DNo.Text = "委託單No."
        '依賴者
        'No
        DMapNo.Text = "圖號"
        '依賴者
  


        DNo.ReadOnly = False

 
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        pFormNo = DFormName.SelectedValue
        DataList()
    End Sub


    Private Sub DFormName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged
        pFormNo = DFormName.SelectedValue

        DNo.ReadOnly = True
    
    
        DProgress.Enabled = True    ''預設為啟動
        DProgress.SelectedIndex = 0 ''預設為ALL
        DDivision.Enabled = True    ''預設為啟動
        DSts.SelectedIndex = 0      ''預設為ALL

        DataGrid1.DataBind()    ''清空DataGrid

        'No
        DNo.Text = "委託單No."
        '依賴者

        'No
        DMapNo.Text = "圖號"
        '依賴者


        DNo.ReadOnly = False


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

        Response.AppendHeader("Content-Disposition", "attachment;filename=MPM_InqCommission.xls")     '程式別不同
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
