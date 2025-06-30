Imports System.Data
Imports System.Data.OleDb

Partial Class MPM_InqCommission
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
    Dim wLevel As String = ""
    Dim wDivision As String = ""
    Dim wName As String = ""
    Dim NowDateTime As String       '現在日期時間
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        Response.Cookies("PGM").Value = "HRWCL_InqCommission.aspx"

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
        NowDateTime = CStr(Now.Date) + " " + _
                      CStr(Now.Hour) + ":" + _
                      CStr(Now.Minute) + ":" + _
                      CStr(Now.Second)     '現在日時
        pFormNo = Request.QueryString("pFormNo")    '表單號碼
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")

    End Sub

    Sub SetSearchItem()
        Dim i As Integer = 0

        Dim SQL As String

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

        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        DFormName.Items.Clear()
        For i = 0 To DBAdapter1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBAdapter1.Rows(i).Item("FormName")
            ListItem1.Value = DBAdapter1.Rows(i).Item("FormNo")
            If ListItem1.Value = pFormNo Then ListItem1.Selected = True
            DFormName.Items.Add(ListItem1)
        Next


        '依賴部門

        SQL = "Select   dep_name  as  DivName From M_emp Group by  dep_name  Order by dep_name "
        DDivision.Items.Clear()
        DDivision.Items.Add("ALL")
        Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)

        For i = 0 To DBAdapter2.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBAdapter2.Rows(i).Item("DivName")
            ListItem1.Value = DBAdapter2.Rows(i).Item("DivName")
            DDivision.Items.Add(ListItem1)
        Next

        'TYPE1

        SQL = "select * from M_referp where cat = '4001' and dkey = 'Type1' order  by Unique_id "
        DType1.Items.Clear()
        DType1.Items.Add("ALL")
        Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)

        For i = 0 To DBAdapter3.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBAdapter3.Rows(i).Item("Data")
            ListItem1.Value = DBAdapter3.Rows(i).Item("Data")
            DType1.Items.Add(ListItem1)
        Next


        'TYPE2

        SQL = "select * from M_referp where cat = '4001' and dkey = 'Type2' order by Unique_id "
        Dtype2.Items.Clear()
        Dtype2.Items.Add("ALL")
        Dim DBAdapter4 As DataTable = uDataBase.GetDataTable(SQL)

        For i = 0 To DBAdapter4.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBAdapter4.Rows(i).Item("Data")
            ListItem1.Value = DBAdapter4.Rows(i).Item("Data")
            Dtype2.Items.Add(ListItem1)
        Next


        'Engine

        SQL = "select * from M_referp where dkey ='EngineSelect-單獨選取'"
        SQL = SQL + " order  by unique_id"

        DEngine.Items.Clear()
        DEngine.Items.Add("ALL")
        Dim DBAdapter5 As DataTable = uDataBase.GetDataTable(SQL)

        For i = 0 To DBAdapter5.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBAdapter5.Rows(i).Item("Data")
            ListItem1.Value = DBAdapter5.Rows(i).Item("Data")
            DEngine.Items.Add(ListItem1)
        Next



        '日期
        DSDate.Text = Format(Now.AddDays(-180), "yyyy/MM/dd")
        DEDate.Text = Format(Now, "yyyy/MM/dd")
      
        BSDate.Attributes.Add("onClick", "calendarPicker('DSDate')")
        BEDate.Attributes.Add("onClick", "calendarPicker('DEDate')")
 


        pFormNo = DFormName.SelectedValue

    End Sub

    Sub DataList()
        Dim i As Integer = 0
        Dim wTableName As String = ""
        Dim SQL As String



        SQL = "SELECT FormNo, FormName, TableName1 FROM M_Form "
        SQL = SQL + " Where Active  =  '1' "
        SQL = SQL + "   And FormNo = '" + DFormName.SelectedValue + "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            wTableName = "F_" + DBAdapter1.Rows(0).Item("TableName1")
        End If





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

        If DMapno.Text <> "圖號" Then
            SQL = SQL + " And " + wTableName + ".MapNo Like '%" + DMapno.Text + "%'"
        End If

     

        If DType1.SelectedValue <> "ALL" Then
            SQL = SQL + " And " + wTableName + ".type1 = '" + DType1.SelectedValue + "'"
        End If

        If Dtype2.SelectedValue <> "ALL" Then
            SQL = SQL + " And " + wTableName + ".type2 = '" + Dtype2.SelectedValue + "'"
        End If




        SQL = SQL + " And " + wTableName + ".CompletedTime between '" + DSDate.Text + " 00:00:00' and '" + DEDate.Text + " 23:59:59'"



  

  

        If DEngine.SelectedValue <> "ALL" Or DName.Text <> "擔當者" Then
            SQL = SQL + " and F_MPMProcessesSheet.No in (  select   distinct  no  from  F_MPMProcessesSheetdt where 1=1 "
        End If


        If DName.Text <> "擔當者" Then
            SQL = SQL + " And  starter  Like '%" + DName.Text + "%'"
        End If

        If DEngine.SelectedValue <> "ALL" Then
            '   SQL = SQL + " And Engine = '" + DEngine.SelectedValue + "'"
            SQL = SQL + " And engine = '" + DEngine.SelectedValue + "'"

        End If


        If DEngine.SelectedValue <> "ALL" Or DName.Text <> "擔當者" Then
            SQL = SQL + ")"
        End If



        'Sort
        SQL = SQL + " Order by " + wTableName + ".CreateTime desc "


        Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)
        DataGrid1.DataSource = DBAdapter2
        DataGrid1.DataBind()




    End Sub


    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        pFormNo = DFormName.SelectedValue
        DataList()
    End Sub


    Private Sub DFormName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged
        pFormNo = DFormName.SelectedValue

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

        Response.AppendHeader("Content-Disposition", "attachment;filename=HRWCL_Commission_ist.xls")     '程式別不同
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


    Protected Sub BEDate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BEDate.Click

    End Sub
End Class
