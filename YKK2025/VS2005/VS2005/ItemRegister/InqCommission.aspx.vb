Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Data.OleDb

Partial Class InqCommission
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    '
    Dim pFormNo As String
    Dim wLevel As String = ""
    Dim wDivision As String = ""
    Dim wName As String = ""
    Dim NowDateTime As String       '現在日期時間

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        Response.Cookies("PGM").Value = "InqCommission.aspx"

        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            SetSearchItem()
            DItemCode.ReadOnly = False
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
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定搜尋欄位初始化及屬性
    '**
    '*****************************************************************
    Sub SetSearchItem()
        '表單
        Dim i As Integer
        Dim Sql As String = "SELECT FormName, FormNo FROM M_FORM "
        Sql = Sql + "Where Active = '1' "
        Sql = Sql + "  And FormNo >= '001151' And FormNo <= '001199' "
        Sql = Sql + "  And (InqAuthority = '0' "
        Sql = Sql + "       Or (InqAuthority = '1' "
        Sql = Sql + "           And (InqUser like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        Sql = Sql + "                Or InqUserName like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        Sql = Sql + "               ) "
        Sql = Sql + "          ) "
        Sql = Sql + "      ) "
        Sql = Sql + "Order by FormNo "
        Dim dtForm As DataTable = uDataBase.GetDataTable(Sql)
        DFormName.Items.Clear()
        For i = 0 To dtForm.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtForm.Rows(i)("FormName").ToString.Trim
            ListItem1.Value = dtForm.Rows(i)("FormNo")
            If ListItem1.Value = pFormNo Then ListItem1.Selected = True
            DFormName.Items.Add(ListItem1)
        Next
        '日期
        DSDate.Text = Format(Now.Date, "yyyy/MM/01")
        DEDate.Text = Format(Now.Date, "yyyy/MM/dd")
        '行事曆事件
        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DEDate');"
        '其他
        DNo.Text = "No."
        'modify-start by joy 160617
        'DName.Text = "申請者"
        DName.Text = "申請者 / 部門"
        'modify-end
        DItemName1.Text = "ItemName1"
        DItemName2.Text = "ItemName2"
        DItemName3.Text = "ItemName3"
        DNo.ReadOnly = False
        DName.ReadOnly = False
        DItemName1.ReadOnly = False
        DItemName2.ReadOnly = False
        DItemName3.ReadOnly = False
        DSDate.ReadOnly = False
        DEDate.ReadOnly = False
        '表單No
        pFormNo = DFormName.SelectedValue
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     資料抽出
    '**
    '*****************************************************************
    Sub DataList()
        Dim Sql, wTableName As String
        ' Get Table Name
        SQL = "SELECT FormNo, FormName, TableName1 FROM M_Form "
        SQL = SQL + " Where Active  =  '1' "
        SQL = SQL + "   And FormNo = '" + DFormName.SelectedValue + "'"
        Dim dtForm As DataTable = uDataBase.GetDataTable(Sql)
        If dtForm.Rows.Count > 0 Then
            If pFormNo >= "001152" And pFormNo <= "001154" Then
                'IRW表單
                wTableName = "V_" + dtForm.Rows(0)("TableName1").ToString.Trim + "_02"
            Else
                '其他表單
              
                wTableName = "F_" + dtForm.Rows(0)("TableName1").ToString.Trim


            End If
        End If
        '設定GridTitle
        DataGrid1.Columns.Item(0).HeaderText = "No."
        DataGrid1.Columns.Item(1).HeaderText = "狀態"
        If pFormNo = "001171" Then
            DataGrid1.Columns.Item(2).HeaderText = "申請者/報告者"
            DataGrid1.Columns.Item(3).HeaderText = "申請日/報告日"
        Else
            DataGrid1.Columns.Item(2).HeaderText = "申請者"
            DataGrid1.Columns.Item(3).HeaderText = "申請日"
        End If


        DataGrid1.Columns.Item(4).HeaderText = "Item Code"
        DataGrid1.Columns.Item(5).HeaderText = "Item Name"
        '設定篩選條件
        Sql = "SELECT "
        If pFormNo >= "001151" And pFormNo <= "001154" Then
            'IRW表單
            Sql = Sql + "Case No When '' Then '未編號' Else No End + "
            '
            'ADD-START ISOS-2308 PJ
            If pFormNo = "001151" Then
                Sql = Sql + "Case AttachFile1 When '' Then '' Else '-@' End + "
                Sql = Sql + "CASE WHEN CHARINDEX('NOT FOUND',SPDURL)>0 THEN '..' ELSE '' END AS Field1, "
            Else
                Sql = Sql + "Case AttachFile1 When '' Then '' Else '-@' End As Field1, "
            End If
            'ADD-END
            '
        Else
            '其他表單
            Sql = Sql + "Case No When '' Then '未編號' Else No End As Field1, "
        End If
        Sql = Sql + "Case Sts When '0' Then '核定中' When '1' Then '完成' When '2' Then '取消' Else '取消' End As Field2, "
        If pFormNo = "001171" Then
            Sql = Sql + "DepName + '--' + AppName as Field3, "
        Else
            Sql = Sql + "Division + '--' + Name as Field3, "
        End If
      
        Sql = Sql + "Convert(VARCHAR(10), Date, 111) as Field4, "


        If pFormNo < "001155" Then
            Sql = Sql + "ItemName1 + ' ' + ItemName2 + ' ' + ItemName3 as Field5, "
        Else
            Sql = Sql + "' ' as Field5, "
        End If

        Sql = Sql + " '....' as WorkFlow, "
        If pFormNo = "001151" Then Sql = Sql + "'http://10.245.1.6/IRW/ItemRegisterSheet_02.aspx?' + "
        If pFormNo = "001152" Then Sql = Sql + "'http://10.245.1.6/IRW/ItemRegisterZIPSheet_02.aspx?' + "
        If pFormNo = "001153" Then Sql = Sql + "'http://10.245.1.6/IRW/ItemRegisterSLDSheet_02.aspx?' + "
        If pFormNo = "001154" Then Sql = Sql + "'http://10.245.1.6/IRW/ItemRegisterCHSheet_02.aspx?' + "
        If pFormNo = "001155" Then Sql = Sql + "'http://10.245.1.6/IRW/ItemRegisterFSLDSheet_02.aspx?' + "
        If pFormNo = "001161" Then Sql = Sql + "'http://10.245.1.6/IRW/PriceInforSheet_02.aspx?' + "
        If pFormNo = "001171" Then Sql = Sql + "'http://10.245.1.6/IRW/IRWNoOrderReportSheet_02.aspx?' + "


        Sql = Sql + "'pUserid=' + '" & Response.Cookies("UserID").Value & "' + "
        Sql = Sql + "'&pFormNo='   + FormNo + "
        Sql = Sql + "'&pFormSno=' + str(FormSno,Len(FormSno)) "
        Sql = Sql + " As ViewURL, "
        '
        If pFormNo = "001151" Then
            Sql = Sql + "'http://10.245.1.6/IRW/StatusList.aspx?' + "
            Sql = Sql + "'pFormNo='   + FormNo + "
            Sql = Sql + "'&pNo='   + No + "
            If DKeepData.SelectedValue = "0" Then
                '--未封存資料
                Sql = Sql + "'&pKeepdata='   + rtrim(ltrim(str(0))) "
            Else
                '--已封存資料
                Sql = Sql + "'&pKeepdata='   + rtrim(ltrim(str(1))) "
            End If
            Sql = Sql + " As OPURL,Code "
        ElseIf pFormNo = "001152" Or pFormNo = "001153" Or pFormNo = "001154" Then
            Sql = Sql + "'http://10.245.1.10/WorkFlow/BefOPList.aspx?' + "
            Sql = Sql + "'pFormNo='   + FormNo + "
            Sql = Sql + "'&pFormSno=' + str(FormSno,Len(FormSno)) + "
            If DKeepData.SelectedValue = "0" Then
                '--未封存資料
                Sql = Sql + "'&pKeepdata='   + rtrim(ltrim(str(0))) + "
            Else
                '--已封存資料
                Sql = Sql + "'&pKeepdata='   + rtrim(ltrim(str(1))) + "
            End If
            Sql = Sql + "'&pApplyID=aaa' "
            Sql = Sql + " As OPURL, Code "
        Else
            Sql = Sql + "'http://10.245.1.10/WorkFlow/BefOPList.aspx?' + "
            Sql = Sql + "'pFormNo='   + FormNo + "
            Sql = Sql + "'&pFormSno=' + str(FormSno,Len(FormSno)) + "
            If DKeepData.SelectedValue = "0" Then
                '--未封存資料
                Sql = Sql + "'&pKeepdata='   + rtrim(ltrim(str(0))) + "
            Else
                '--已封存資料
                Sql = Sql + "'&pKeepdata='   + rtrim(ltrim(str(1))) + "
            End If
            Sql = Sql + "'&pApplyID=aaa' "
            Sql = Sql + " As OPURL,'' as Code "

        End If
        Sql = Sql + "FROM " + wTableName + " "
        Sql = Sql + "Where Sts <> '9' "
        '狀態
        If DProgress.SelectedValue <> "ALL" Then
            If DProgress.SelectedValue = "1" Then
                Sql = Sql + " And Sts =  '0'  "
            End If
            If DProgress.SelectedValue = "2" Then
                Sql = Sql + " And Sts <>  '0'  "
            End If
        End If
        '完成狀態
        If DSts.SelectedValue <> "ALL" Then
            If DSts.SelectedValue <> "3" Then
                Sql = Sql + " And Sts = '" + DSts.SelectedValue + "'"
            Else
                Sql = Sql + " And (Sts = '2' or Sts = '3') "
            End If
        End If
        'No
        If DNo.Text <> "No." And DNo.Text <> "" Then
            Sql = Sql + " And No Like '%" + DNo.Text + "%'"
        End If
        '申請者
        'modify-start by joy 160617
        'If DName.Text <> "申請者" And DName.Text <> "" Then
        'Sql = Sql + " And Name Like '%" + DName.Text + "%'"
        'End If
        If DName.Text <> "申請者 / 部門" And DName.Text <> "" Then
            Sql = Sql + " And (Name Like '%" + DName.Text + "%' OR Division Like '%" + DName.Text + "%')"
        End If
        'modify-end
        'ItemName1
        If DItemName1.Text <> "ItemName1" And DItemName1.Text <> "" Then
            Sql = Sql + " And ItemName1 Like '%" + DItemName1.Text + "%'"
        End If
        'ItemName2
        If DItemName2.Text <> "ItemName2" And DItemName2.Text <> "" Then
            Sql = Sql + " And ItemName2 Like '%" + DItemName2.Text + "%'"
        End If
        'ItemName3
        If DItemName3.Text <> "ItemName3" And DItemName3.Text <> "" Then
            Sql = Sql + " And ItemName3 Like '%" + DItemName3.Text + "%'"
        End If
        'ItmeCode
        If DItemCode.Text <> "ItemCode" And DItemCode.Text <> "" Then
            Sql = Sql + " And code Like '%" + DItemCode.Text + "%'"
        End If

        '申請日
        If IsDate(DSDate.Text) = True And IsDate(DEDate.Text) = True Then
            Sql = Sql + " And  convert(char(10),Date,111) between '" + CDate(DSDate.Text).ToString("yyyy/MM/dd") + "' and '" + CDate(DEDate.Text).ToString("yyyy/MM/dd") + "' "
        End If
        Sql = Sql + "Order by No desc "
        '
        DataGrid1.DataSource = uDataBase.GetDataTable(Sql)
        DataGrid1.DataBind()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     按鈕(Go)
    '**
    '*****************************************************************
    Protected Sub Go_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        pFormNo = DFormName.SelectedValue
        DataList()
    End Sub
    '*****************************************************************
    '**
    '**     表單變動
    '**
    '*****************************************************************
    Protected Sub DFormName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged
        pFormNo = DFormName.SelectedValue
        'MsgBox("[" + pFormNo + "]")
        DItemCode.Text = ""
        If pFormNo = "001155" Or pFormNo = "001161" Or pFormNo = "001171" Then
            DItemName1.Text = ""
            DItemName2.Text = ""
            DItemName3.Text = ""
            DItemName1.ReadOnly = True
            DItemName2.ReadOnly = True
            DItemName3.ReadOnly = True
            
        Else
            DItemName1.ReadOnly = False
            DItemName2.ReadOnly = False
            DItemName3.ReadOnly = False
        End If
        If pFormNo = "001151" Or pFormNo = "001152" Or pFormNo = "001153" Or pFormNo = "001154" Then
            DItemCode.ReadOnly = False
        Else
            DItemCode.ReadOnly = True
        End If


    End Sub
    '*****************************************************************
    '**
    '**     DataGrid換頁
    '**
    '*****************************************************************
    Protected Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        pFormNo = DFormName.SelectedValue
        '
        DataList()
    End Sub
    '*****************************************************************
    '**
    '**     按鈕(轉Excel)
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '程式別不同

        pFormNo = DFormName.SelectedValue
        DataGrid1.AllowPaging = False   '程式別不同
        '
        DataList()
        '  Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")
        Response.AppendHeader("Content-Disposition", "attachment;filename=Commission_ist.xls")     '程式別不同
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
    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

    Protected Sub BEDate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BEDate.Click

    End Sub

    Protected Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        If pFormNo = "001151" Or pFormNo = "001152" Or pFormNo = "001153" Or pFormNo = "001154" Then
            e.Item.Cells(4).Visible = True
        ElseIf pFormNo = "001171" Then
         
            e.Item.Cells(4).Visible = False
            e.Item.Cells(5).Visible = False
        Else
            e.Item.Cells(4).Visible = False
        End If
        e.Item.Cells(0).Attributes.Add("style", "vnd.ms-excel.numberformat:@")

    End Sub
 
    
End Class
