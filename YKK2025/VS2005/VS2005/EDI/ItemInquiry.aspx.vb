Imports System.Data
Imports System.Data.SqlClient

Partial Class ItemInquiry
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
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期
    '
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Page_Load)
    '**     Load Page 處理
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            SetSearchItem()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(Now.Date) + " " + _
                      CStr(Now.Hour) + ":" + _
                      CStr(Now.Minute) + ":" + _
                      CStr(Now.Second)     '現在日時
        NowDate = CStr(Now.Date)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetSearchItem)
    '**     設定搜尋欄位
    '**
    '*****************************************************************
    Sub SetSearchItem()
        Dim i As Integer
        Dim Sql As String
        '
        DBuyer.Items.Clear()
        Sql = "SELECT Buyer, "
        Sql = Sql + "ISNULL( "
        Sql = Sql + "( "
        Sql = Sql + "Select Top 1 Name from M_ControlRecord where M_ITEMConvert.Buyer= M_ControlRecord.Buyer "
        Sql = Sql + "UNION "
        Sql = Sql + "Select Top 1 Name from M_FControlRecord where M_ITEMConvert.Buyer= M_FControlRecord.Buyer "
        Sql = Sql + "UNION "
        Sql = Sql + "Select 'NIKE'  where M_ITEMConvert.Buyer= '000013' "
        Sql = Sql + ") "
        Sql = Sql + ", 'NO-NAME') AS BuyerName "
        Sql = Sql + "From M_ITEMConvert "
        Sql = Sql + "Where  Buyer in (Select Buyer From M_ControlRecord) "
        Sql = Sql + "Or Buyer in (Select Buyer From M_FControlRecord) "
        Sql = Sql + "Or Buyer = '000013' "
        Sql = Sql + "Group by Buyer "
        Sql = Sql + "Order by Buyer "
        '
        Dim dt_Item As DataTable = uDataBase.GetDataTable(Sql)
        For i = 0 To dt_Item.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dt_Item.Rows(i).Item("BuyerName")
            ListItem1.Value = dt_Item.Rows(i).Item("Buyer")
            ListItem1.Selected = False
            DBuyer.Items.Add(ListItem1)
        Next
        '
        DCode.Text = ""
        DColor.Text = ""
        If DBuyer.SelectedValue = "000013" Then
            DAction.Enabled = True
        Else
            DAction.SelectedIndex = 0
            DAction.Enabled = False
        End If
        '
        DStartTime.Text = CStr(Now.Date.AddMonths(-2))
        DEndTime.Text = CStr(Now.Date.AddDays(-1))
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BuyerChange)
    '**     Buyer Change
    '**
    '*****************************************************************
    Protected Sub DBuyer_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DBuyer.SelectedIndexChanged
        If DBuyer.SelectedValue = "000013" Then
            DAction.Enabled = True
        Else
            DAction.SelectedIndex = 0
            DAction.Enabled = False
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(DataList)
    '**     資料抽出
    '**
    '*****************************************************************
    Sub DataList()
        Dim Sql As String
        '
        Sql = "SELECT "
        Sql = Sql + "Buyer, Action, Code, Description, Color1, Color2, Color3, SliderStatus, YCode, "
        Sql = Sql + "YName1+' '+YName2+' '+YName3 As YName, "

        Sql = Sql + "ISNULL( "
        Sql = Sql + "( "
        Sql = Sql + "Select Name from M_ControlRecord where M_ITEMConvert.Buyer= M_ControlRecord.Buyer "
        Sql = Sql + "UNION "
        Sql = Sql + "Select Name from M_FControlRecord where M_ITEMConvert.Buyer= M_FControlRecord.Buyer "
        Sql = Sql + "UNION "
        Sql = Sql + "Select 'NIKE'  where M_ITEMConvert.Buyer= '000013' "
        Sql = Sql + ") "
        Sql = Sql + ", 'NO-NAME') AS BuyerDesc "
        '
        Sql = Sql + "From M_ITEMConvert "
        Sql = Sql + "Where Buyer = '" + DBuyer.SelectedValue + "' "
        Sql = Sql + "  And Action = '" + DAction.SelectedValue + "' "
        'Code
        If DCode.Text <> "" Then
            Sql = Sql + "  And Code Like '%" + DCode.Text + "%' "
        End If
        'Color
        If DColor.Text <> "" Then
            Sql = Sql + "  And ( Color1 Like '%" + DColor.Text + "%' "
            Sql = Sql + "     or Color2 Like '%" + DColor.Text + "%' "
            Sql = Sql + "     or Color3 Like '%" + DColor.Text + "%' "
            Sql = Sql + "     ) "
        End If
        '期間範圍
        Sql = Sql + "  And ( "
        Sql = Sql + "       (CreateTime >= '" + DStartTime.Text + " 00:00:01" + "' And  CreateTime <= '" + DEndTime.Text + " 23:59:59" + "') "
        Sql = Sql + "       or "
        Sql = Sql + "       (ModifyTime >= '" + DStartTime.Text + " 00:00:01" + "' And  ModifyTime <= '" + DEndTime.Text + " 23:59:59" + "') "
        Sql = Sql + "      ) "
        '
        Sql = Sql + "Order by Code, Color1, Color2, Color3 "
        '
        GridView1.DataSource = uDataBase.GetDataTable(Sql)
        GridView1.DataBind()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Go)
    '**     按鈕(Go)
    '**
    '*****************************************************************
    Protected Sub Go_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Go.Click
        DataList()
    End Sub
    '*****************************************************************
    '**
    '**     按鈕(轉Excel)
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        '
        Response.AppendHeader("Content-Disposition", "attachment;filename=EDI_ItemTransferList.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)

        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()
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

End Class
