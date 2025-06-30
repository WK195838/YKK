Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class FindItemImagePage
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim wUserID As String = ""      '申請者

    '外部Object
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim ConnectString As String = uCommon.GetAppSetting("WAVESDB")
    Dim ConnectStringCL As String = uCommon.GetAppSetting("WAVESDBCL")
    Dim EDLConnect As String = uCommon.GetAppSetting("EDLDB")

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900 ' 設定逾時時間

        SetParameter()                  ' 設定共用參數
        If Not IsPostBack Then
            SetDefaultValue()                   '設定初值
            If DCode.Text <> "" Or DName1.Text <> "" Then
                DataList()
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日時
        wUserID = Request.QueryString("pUserID")
        Response.Cookies("PGM").Value = "FindItemImagePage.aspx"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        Dim i, j, k As Integer
        Dim str, key As String
        Dim xItemName As String()
        '
        DCode.Text = ""
        DName1.Text = ""
        DName2.Text = ""
        DName3.Text = ""
        DName4.Text = ""
    End Sub
    '------------------------------------------------------
    Protected Sub BFindItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFindItem.Click
        If DCode.Text <> "" Or DName1.Text <> "" Then
            DataList()
        End If
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        DataList()
        GridView1.PageIndex = e.NewPageIndex
        GridView1.DataBind()
    End Sub

    Sub DataList()
        Dim Sql As String = ""
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        cn.ConnectionString = ConnectString
        Sql = "SELECT "
        Sql = Sql + "LTRIM(RTRIM(ITEM)) As Code, "
        Sql = Sql + "LTRIM(RTRIM(ITEM_NAME1)) As Name1, "
        Sql = Sql + "LTRIM(RTRIM(ITEM_NAME2)) As Name2, "
        Sql = Sql + "LTRIM(RTRIM(ITEM_NAME3)) As Name3, "
        '
        Sql = Sql + "'' As Season, "
        Sql = Sql + "'' As Color, "
        Sql = Sql + "'' As ColorName, "
        Sql = Sql + "LTRIM(RTRIM(SLIDER)) As Slider, "
        '
        '材料
        Sql = Sql + "'材料' as ISOS2FAS, "
        Sql = Sql + "'' as URL, "
        '
        '**BCP
        Sql = Sql + "'BCP' as BCP, "
        Sql = Sql + "'' as BCPURL, "
        '
        '**BCT
        Sql = Sql + "'BCT' as BCT, "
        Sql = Sql + "'' as BCTURL, "
        '
        '**IRW
        Sql = Sql + "'IRW' as IRW, "
        Sql = Sql + "'' as IRWURL, "
        '
        '**SPC
        Sql = Sql + "'SPC' as SPC, "
        Sql = Sql + "'' as SPCURL, "
        '
        '**SPP
        Sql = Sql + "'SPP' as SPP, "
        Sql = Sql + "'' as SPPURL, "
        '
        '**IMG
        Sql = Sql + "'IMG' as IMG, "
        Sql = Sql + "'http://10.245.0.3/ISOS/PS_INQ_OpenFile.aspx?pUserID=" & wUserID & "&pBuyer=" & "" & "&pBuyerItem=' + LTrim(RTrim(ITEM)) + '&pFun=IMG" & "' as IMGURL, "
        '
        '**COMBI
        Sql = Sql + "'COMBI' as COMBI, "
        Sql = Sql + "'' as COMBIURL, "
        '
        '**估價
        Sql = Sql + "'估價' as VAL, "
        Sql = Sql + "''  as VALURL, "
        '
        Sql = Sql + "'' as XXX "
        '
        Sql = Sql + "From MST_ITEM "
        '
        If DCode.Text <> "" Then
            Sql = Sql + "Where ITEM LIKE '%" & DCode.Text.ToUpper & "%' "
        Else
            Sql = Sql + "Where REPLACE(ITEM_NAME1+ITEM_NAME2+ITEM_NAME3,' ','') LIKE '%" & Replace(DName1.Text.ToUpper, " ", "") & "%' "
            If DName2.Text <> "" Then
                Sql = Sql + "And REPLACE(ITEM_NAME1+ITEM_NAME2+ITEM_NAME3,' ','') LIKE '%" & Replace(DName2.Text.ToUpper, " ", "") & "%' "
            End If
            If DName3.Text <> "" Then
                Sql = Sql + "And REPLACE(ITEM_NAME1+ITEM_NAME2+ITEM_NAME3,' ','') LIKE '%" & Replace(DName3.Text.ToUpper, " ", "") & "%' "
            End If
            If DName4.Text <> "" Then
                Sql = Sql + "And REPLACE(ITEM_NAME1+ITEM_NAME2+ITEM_NAME3,' ','') LIKE '%" & Replace(DName4.Text.ToUpper, " ", "") & "%' "
            End If
        End If
        Sql = Sql + "And NoDisp <> '1' "
        Sql = Sql + "ORDER BY ITEM DESC, ITEM_NAME1, ITEM_NAME2, ITEM_NAME3 "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(Sql, cn)
        DBAdapter1.Fill(ds, "FA000")
        If ds.Tables("FA000").Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = ds
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If

        cn.Close()
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim cn, cn1 As New OleDbConnection
        Dim ds As New DataSet
        Dim dc As New OleDbCommand
        Dim sql As String
        Dim i As Integer
        '
        cn.ConnectionString = EDLConnect
        '
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
            '=================================================================
            '共通
            '=================================================================
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '添加新的表头第一行
            tcl.Add(New TableHeaderCell())
            tcl(0).Text = "Item Code"
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "Item Name-1"
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "Item Name-2"
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "Item Name-3"
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "Slider Code"
            tcl.Add(New TableHeaderCell())
            tcl(5).Text = "PFAS"
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            '
            cn1.ConnectionString = ConnectStringCL
            ds.Clear()
            '
            sql = "SELECT ISNULL(IF8CA0,'') AS PFAS From MST_ITEM "
            sql &= "Where ITEM = '" & Trim(e.Row.Cells(2).Text) & "' "
            Dim DBAdapter2 As New OleDbDataAdapter(sql, cn1)
            DBAdapter2.Fill(ds, "ITEM")
            If ds.Tables("ITEM").Rows.Count > 0 Then
                If ds.Tables("ITEM").Rows(0).Item("PFAS") <> "" Then
                    e.Row.Cells(9).ForeColor = Color.Red
                    e.Row.Cells(9).BackColor = Color.Yellow
                    e.Row.Cells(9).Text = e.Row.Cells(9).Text & "[" & Trim(ds.Tables("ITEM").Rows(0).Item("PFAS")) & "]"
                End If
            End If
            '
            cn1.Close()
            '
            e.Row.Cells(0).Visible = False
            e.Row.Cells(1).Visible = False
            e.Row.Cells(7).Visible = False
            e.Row.Cells(8).Visible = False
            e.Row.Cells(10).Visible = False
            e.Row.Cells(11).Visible = False
            e.Row.Cells(12).Visible = False
            e.Row.Cells(13).Visible = False
            e.Row.Cells(14).Visible = False
            e.Row.Cells(16).Visible = False
            e.Row.Cells(17).Visible = False
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub
    '
    '*****************************************************************
    '**
    '**     轉Excel
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        '
        Response.Clear()
        Response.Buffer = True

        Response.AppendHeader("Content-Disposition", "attachment;filename=FindItemPage.xls")     '程式別不同

        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        '
        DataList()
        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()
    End Sub
    '
    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub


End Class
