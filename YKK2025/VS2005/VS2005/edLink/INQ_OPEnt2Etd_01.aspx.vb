Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class INQ_OPEnt2Etd_01
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim pMonth As String
    Dim pDays As String
    Dim pAction As String

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
        pMonth = Request.QueryString("pMonth")
        pDays = Request.QueryString("pDays")
        pAction = Request.QueryString("pAction")
        '-----------------------------------------------------------------
        '-- 初值
        GridView1.Visible = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetData)
    '**     取得資料
    '**
    '*****************************************************************
    Sub GetData()
        Dim Sql As String = ""

        Sql = "select * "
        Sql &= ", convert(int, Quantity) as Qty  "
        Sql &= "FROM V_OPReportData_Ent2Etd "
        Sql &= "WHERE Sample = '' "
        '
        Sql &= "And EntryYYMM = '" & pMonth & "' "
        '
        If pAction = "LT" Then
            Sql &= "and Entry2ETDDays < '" & pDays & "' "
        End If
        If pAction = "GT" Then
            Sql &= "and Entry2ETDDays > '" & pDays & "' "
        End If
        If pAction = "EQ" Then
            Sql &= "and Entry2ETDDays = '" & pDays & "' "
        End If
        Sql &= "ORDER BY Entry2ETDDays desc, OrderNo, OrderSubNo "
        '
        Dim dt_OP As DataTable = uDataBase.GetDataTable(Sql)
        If dt_OP.Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = dt_OP
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then

        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            For i = 0 To 13
                Select Case i
                    Case 2, 5, 8
                        e.Row.Cells(i).Attributes.Add("class", "text")
                    Case Else
                End Select

            Next

        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

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

        Dim wAllowPaging As Boolean = GridView1.AllowPaging   '程式別不同

        GridView1.AllowPaging = False   '程式別不同
        GetData()                      '程式別不同

        Response.Clear()
        Response.Buffer = True
        Response.AppendHeader("Content-Disposition", "attachment;filename=OPEnt2Etd.xls")     '程式別不同
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        Dim style As String = "<style> .text { mso-number-format:\@; } </style> " '文字樣式字串

        GridView1.RenderControl(hw)
        Response.Write(style)
        Response.Write(tw.ToString())
        Response.End()

    End Sub
End Class
