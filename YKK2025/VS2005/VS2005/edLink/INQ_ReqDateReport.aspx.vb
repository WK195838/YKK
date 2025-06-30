Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class INQ_ReqDateReport
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim UserID As String            'UserID
    Dim pStartDate As String
    Dim pEndDate As String

    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            SetParameter()                              '設定參數
            GetData()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定參數
    '**
    '*****************************************************************
    Sub SetParameter()
        '-----------------------------------------------------------------
        '-- 系統參數
        '-----------------------------------------------------------------
        Server.ScriptTimeout = 900                                  '設定逾時時間
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        UserID = Request.QueryString("pUserID")             'UserID
        If Request.QueryString("pStartDate") = "" Then
            pStartDate = DateAdd("d", -1, Now).ToString("yyyyMMdd")
        Else
            pStartDate = Request.QueryString("pStartDate")
        End If
        If Request.QueryString("pStartDate") = "" Then
            pEndDate = NowDate
        Else
            pEndDate = Request.QueryString("pEndDate")
        End If
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        DCust.Text = ""
        DSalesCode.Text = ""

        DSStartDate.Text = pStartDate
        DSEndDate.Text = pEndDate

        AtLongDate.Checked = False
        DataGrid1.Visible = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetData)
    '**     取得資料
    '**
    '*****************************************************************
    Sub GetData()
        Dim Sql As String = ""
        Sql &= "SELECT *, '['+Buyer+']' as buyer1, "

        Sql &= "CASE WHEN EDIPO='' THEN '' ELSE '[' + EDIPo + '][' + Convert(varchar(8),EDIPoDate) + '][' + EDICust + '-' + EDIBUYER + '][' + (CASE WHEN EDISMPF='' THEN 'BULK' ELSE 'SAMPLE' END) + ']' END AS EDIInf, "
        Sql &= "CASE WHEN LModUser='' THEN '' ELSE '[' + LModUser + '][' + Convert(varchar(8),LModDate) + '][' + Convert(varchar(6),LModTime) + ']' END AS MODInf, "

        Sql &= "CASE WHEN SAMPLE='' THEN BulkDays ELSE SampleDays END AS StdDays, "
        Sql &= "CASE WHEN SAMPLE='' THEN Convert(varchar(10),BulkDATE,112) ELSE Convert(varchar(10),SampleDATE,112) END AS StdDate, "

        Sql &= "'http://10.245.1.6/WorkFlowSub/INQ_S5MReport.aspx?pOrderNo=' + OrderNo AS URL "

        If AtLongDate.Checked = True Then
            Sql &= "FROM V_OPReportData_Delivey_L "
        Else
            Sql &= "FROM V_OPReportData_Delivey_E "
        End If

        Sql &= "Where CUSTOMER NOT LIKE 'ZZZZJJJ%' "

        If Not String.IsNullOrEmpty(DCust.Text.Trim) Then
            Sql &= "and CUSTOMER = '" & DCust.Text.Trim & "' "
        End If
        If Not String.IsNullOrEmpty(DSalesCode.Text.Trim) Then
            Sql &= "and SalesMan = '" & DSalesCode.Text.Trim & "' "
        End If
        If Not String.IsNullOrEmpty(DSStartDate.Text.Trim) And Not String.IsNullOrEmpty(DSEndDate.Text.Trim) Then
            Sql &= "and ENTRYDATE >= " & DSStartDate.Text.Trim & " "
            Sql &= "and ENTRYDATE <= " & DSEndDate.Text.Trim & " "
        End If
        Sql &= "order by Salesman,Customer,Buyer,EntryDate "

        Dim dt_CustomerGroup As DataTable = uDataBase.GetDataTable(Sql)
        If dt_CustomerGroup.Rows.Count > 0 Then
            DataGrid1.Visible = True
            DataGrid1.DataSource = dt_CustomerGroup
            DataGrid1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BSearch_Click)
    '**     搜尋資料
    '**
    '*****************************************************************
    Protected Sub BSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSearch.Click
        GetData()
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

        ' pFormNo = DFormName.SelectedValue
        DataGrid1.AllowPaging = False   '程式別不同
        GetData()                      '程式別不同

        Response.Clear()
        Response.Buffer = True
        Response.AppendHeader("Content-Disposition", "attachment;filename=CheckReqDateList.xls")     '程式別不同
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
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
