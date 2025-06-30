Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing


Partial Class INQ_REQStatus01
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
    Dim pCTYC As String
    Dim pAction As String

    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                              '設定參數
        If Not IsPostBack Then
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
        pCTYC = Request.QueryString("pCTYC")
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
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        cn.ConnectionString = ConnectString
        '
        Sql = "SELECT ordn5c, CHAR(cstc5c) AS CSTC5C, CHAR(byrc5c) AS BYRC5C, rdlu5m , rdlu5c, "

        Sql &= "CASE CPON5C "
        Sql &= "  WHEN '' THEN UIDCF || '/' || CHAR(HCTUF) || '/' || CHAR(HCRTF) "
        Sql &= "  ELSE UIDCF || '/' || CHAR(HCTUF) || '/' || CHAR(HCRTF) || '/EDI[' || CPON5C || ']' "
        Sql &= "END AS FREMARK, "
        Sql &= "UIDCL || '/' || CHAR(HCTUL) || '/' || CHAR(HCRTL) AS LREMARK, "

        Sql &= "'' AS URL "
        Sql &= "FROM SONGOLIB.DELSTATUS "
        Sql &= "WHERE ordn5c <> '9999999999'"
        '
        Sql &= "And CTYC5C = '" & pCTYC & "' "

        'If pCTYC = "Y" Then
        '    Sql &= "And (CTYC5C = 'E' OR CTYC5C = 'K') "
        'Else
        '    Sql &= "And (CTYC5C = 'A' OR CTYC5C = 'H') "
        'End If
        '
        If pAction = "P" Then
            Sql &= "And rdlu5m < rdlu5c "
        End If
        If pAction = "A" Then
            Sql &= "And rdlu5m > rdlu5c "
        End If
        If pAction = "S" Then
            Sql &= "And rdlu5m = rdlu5c "
        End If

        Sql &= "And YYMM = " & pMonth & " "
        Sql &= "ORDER by ordn5c, cstc5c, byrc5c, rdlu5m , rdlu5c "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(Sql, cn)
        DBAdapter1.Fill(ds, "S5M00")
        If ds.Tables("S5M00").Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = ds
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If

        cn.Close()

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
            For i = 0 To 4
                Select Case i
                    Case 0
                        Dim h1 As New HyperLink
                        h1.Target = "_blank"
                        h1.Text = e.Row.Cells(i).Text
                        h1.NavigateUrl = "http://10.245.1.6/WorkFlowSub/INQ_S5MReport.aspx?pOrderNo=" & e.Row.Cells(0).Text
                        e.Row.Cells(i).Text = ""
                        e.Row.Cells(i).Controls.Add(h1)
                    Case 1, 2
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
        Response.AppendHeader("Content-Disposition", "attachment;filename=ReqDateList.xls")     '程式別不同
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
