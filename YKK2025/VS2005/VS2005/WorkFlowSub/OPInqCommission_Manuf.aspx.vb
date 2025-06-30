Imports System.Data
Imports System.Data.OleDb

Partial Class OPInqCommission_Manuf
    Inherits System.Web.UI.Page

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
        Response.Cookies("PGM").Value = "OPInqCommission_Manuf.aspx"

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
        NowDateTime = CStr(DateTime.Now.Date) + " " + _
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
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        '表單
        SQL = "SELECT FormName, FormNo FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And (FormNo = '000003' or FormNo = '000007') "
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
        SQL = "Select DivName From M_Users Group by DivName Order by DivName "
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

        pFormNo = DFormName.SelectedValue
        Search_Item_Attribute()
    End Sub

    Sub DataList()
        Dim i As Integer = 0
        Dim wTableName As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim MainUrl = System.Configuration.ConfigurationManager.AppSettings("MainUrl")  'Main System URL
        Dim ManufIn = System.Configuration.ConfigurationManager.AppSettings("Http") + System.Configuration.ConfigurationManager.AppSettings("ManufInFilePath")  'ManufInSheet Document Path
        Dim ManufOut = System.Configuration.ConfigurationManager.AppSettings("Http") + System.Configuration.ConfigurationManager.AppSettings("ManufOutFilePath")  'ManufOutSheet Document Path
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        GridView1.Columns.Item(1).HeaderText = "委託No"
        GridView1.Columns.Item(2).HeaderText = "開發狀態"
        GridView1.Columns.Item(3).HeaderText = "依賴者"
        GridView1.Columns.Item(4).HeaderText = "依賴日"
        GridView1.Columns.Item(5).HeaderText = "Buyer"
        GridView1.Columns.Item(6).HeaderText = "Size-Chain-胴體"
        GridView1.Columns.Item(7).HeaderText = "圖號"
        GridView1.Columns.Item(8).HeaderText = "Slider Code"

        GridView2.Columns.Item(0).HeaderText = "委託No"
        GridView2.Columns.Item(1).HeaderText = "開發狀態"
        GridView2.Columns.Item(2).HeaderText = "依賴者"
        GridView2.Columns.Item(3).HeaderText = "依賴日"
        GridView2.Columns.Item(4).HeaderText = "Buyer"
        GridView2.Columns.Item(5).HeaderText = "Size-Chain-胴體"
        GridView2.Columns.Item(6).HeaderText = "圖號"
        GridView2.Columns.Item(7).HeaderText = "Slider Code"

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
            SQL = SQL + "Case " + wTableName + ".No When '' Then '未編號' Else " + wTableName + ".No End As Field1, "
            SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '開發中' When '1' Then '開發完成(OK)' When '2' Then '開發完成(NG)' Else '開發取消' End As Field2, "
            SQL = SQL + wTableName + ".Division + '--' + "
            SQL = SQL + wTableName + ".Person as Field3, "
            SQL = SQL + "V_WaitHandle_01.ApplyTime as Field4, "
            SQL = SQL + wTableName + ".Buyer   as Field5, "
            SQL = SQL + wTableName + ".Spec    as Field6, "
            SQL = SQL + wTableName + ".MapNo   as Field7, "
            SQL = SQL + wTableName + ".SliderCode as Field8, "

            SQL = SQL + " '....' as WorkFlow, "
            SQL = SQL + "'" + MainUrl + "' + ViewURL as ViewURL, "

            If DFormName.SelectedValue = "000003" Then
                SQL = SQL + "'" + ManufIn + "' + MapFile as ImageURL, "
            Else
                SQL = SQL + "'" + ManufOut + "' + MapFile as ImageURL, "
            End If

            SQL = SQL + "'" + MainUrl + "' + 'BefOPList.aspx?' + "
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

            '部門
            If DDivision.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".Division = '" + DDivision.SelectedValue + "'"
            End If
            '依賴人
            If DPerson.Text <> "依賴者" And DPerson.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".Person Like '%" + DPerson.Text + "%'"
            End If
            'Buyer
            If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".Buyer Like '%" + DBuyer.Text + "%'"
            End If
            'Slider Code
            If DSliderCode.Text <> "Slider Code" And DSliderCode.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".SliderCode Like '%" + DSliderCode.Text + "%'"
            End If
            'Spec
            If DSpec.Text <> "Size-Chain-胴體" And DSpec.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".Spec Like '%" + DSpec.Text + "%'"
            End If
            'MapNo
            If DMapNo.Text <> "圖號" And DMapNo.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".MapNo Like '%" + DMapNo.Text + "%'"
            End If
            'Cpsc
            If DCpsc.Text <> "CPSC" And DCpsc.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".Cpsc Like '%" + DCpsc.Text + "%'"
            End If
            'No
            If DNo.Text <> "委託單No." And DNo.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".No Like '%" + DNo.Text + "%'"
            End If
            'Sort
            SQL = SQL + " Order by " + wTableName + ".No desc "

            OleDbConnection1.Open()
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "WaitHandle")

            'GridView-1
            GridView1.DataSource = DBDataSet2
            GridView1.DataBind()
            'GridView-2
            GridView2.DataSource = DBDataSet2
            GridView2.DataBind()

            OleDbConnection1.Close()
            GridView2.Visible = False
        End If

    End Sub

    Sub Search_Item_Attribute()
        '依賴者
        DPerson.Text = "依賴者"
        'Buyer
        DBuyer.Text = "Buyer"
        'SliderCode
        DSliderCode.Text = "Slider Code"
        'Spec
        DSpec.Text = "Size-Chain-胴體"
        'No
        DNo.Text = "委託單No."
        '圖號
        DMapNo.Text = "圖號"
        'Cpsc
        DCpsc.Text = "CPSC"

        DNo.ReadOnly = False
        DPerson.ReadOnly = False
        DBuyer.ReadOnly = False
        DSliderCode.ReadOnly = False
        DSpec.ReadOnly = False
        DMapNo.ReadOnly = False
        DCpsc.ReadOnly = False
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        GridView1.PageIndex = 0
        pFormNo = DFormName.SelectedValue
        DataList()
    End Sub


    Private Sub DFormName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged
        pFormNo = DFormName.SelectedValue

        DNo.ReadOnly = True
        DPerson.ReadOnly = True
        DBuyer.ReadOnly = True
        DSpec.ReadOnly = True
        DMapNo.ReadOnly = True
        DSliderCode.ReadOnly = True
        DCpsc.ReadOnly = True
        DProgress.Enabled = True    ''預設為啟動
        DProgress.SelectedIndex = 0 ''預設為ALL
        DDivision.Enabled = True    ''預設為啟動
        DSts.SelectedIndex = 0      ''預設為ALL

        GridView1.DataBind()    ''清空GridView

        Search_Item_Attribute()
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

        GridView2.Visible = True

        Response.AppendHeader("Content-Disposition", "attachment;filename=Commission_List_01.xls")     '程式別不同
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)

        GridView2.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
        pFormNo = DFormName.SelectedValue
        DataList()
    End Sub
End Class
