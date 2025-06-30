Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Data.OleDb

Partial Class IRWNOActivityCust01
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim UserID As String            'UserID
    Dim pUID As String
    Dim pDataType As String
    Dim xCustName As String
    Dim YYMM(6) As String

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
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                              '設定參數
        SetPopupFunction()                          '設定彈出視窗事件

        If Not IsPostBack Then                      'PostBack
            SetDefaultValue()                       '設定初值
            ShowData()
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
        Response.Cookies("PGM").Value = "IRWNOActivityCust01.aspx"     '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        UserID = Request.QueryString("pUserID")             'UserID
        pUID = Request.QueryString("pUID")
        pDataType = Request.QueryString("pDataType")
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GRIDVIEW1.Visible = False
        '
        Dim Sql As String
        Sql = "SELECT * FROM W_IRWSeqNo Order By Unique_id "
        Dim dt As DataTable = uDataBase.GetDataTable(Sql)
        For i As Integer = 0 To dt.Rows.Count - 1
            YYMM(i + 1) = dt.Rows(i).Item("YYMM")
        Next
        '
        xCustName = "XXXXXX"
        Sql = "SELECT CustName "
        Sql = Sql & "From V_IRWNOActivityCust_01 "
        Sql = Sql & "Where Unique_ID = " & pUID & " "
        Dim dt1 As DataTable = uDataBase.GetDataTable(Sql)
        If dt1.Rows.Count > 0 Then
            xCustName = dt1.Rows(0).Item("CustName")
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
    End Sub

    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim Sql As String

        'On Error GoTo LBL_Error
        '
        '篩選資料
        Sql = "SELECT "
        Sql = Sql & "DepName, CustName, MM1_YES, MM1_NO, MM2_YES, MM2_NO, MM3_YES, MM3_NO, "
        Sql = Sql & "MM4_YES, MM4_NO, MM5_YES, MM5_NO, MM6_YES, MM6_NO, TOTAL_YES, TOTAL_NO, "
        Sql = Sql & "0 AS MM1_PER, 0 AS MM2_PER, 0 AS MM3_PER, 0 AS MM4_PER, 0 AS MM5_PER, 0 AS MM6_PER, 0 AS TOTAL_PER, "
        Sql = Sql & "EmpID, "

        Sql &= "'IRWNOActivityCust02.aspx?pUserID=" & UserID & "&pDepCode=' + DepCode + '&pUID=' + LTRIM(RTRIM(STR([UNIQUE_ID]))) as URL "

        Sql = Sql & "From V_IRWNOActivityCust_01 "
        Sql = Sql & "Where DataType = 'C' "
        Sql = Sql & "And CustName = '" & xCustName & "' "
        If pDataType = "SALES" Then
            Sql = Sql & "And DepName IN ('營業一課','營業二課','營業三課','營業五課','中部營業部','南部營業部') "
            Sql = Sql & "Order By MM1_PER + MM2_PER + MM3_PER  DESC "
        Else
            Sql = Sql & "And DepName IN ('企劃一課','企劃二課','企劃三課') "
            Sql = Sql & "Order By MM1_PER + MM2_PER + MM3_PER  DESC "
        End If
        Dim dt As DataTable = uDataBase.GetDataTable(Sql)
        If dt.Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = dt
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
        End If
        '
        GoTo LBL_End
        '##
LBL_Error:
        On Error GoTo -1
        uJavaScript.PopMsg(Me, "指定條件搜尋不到資料,請確認!")
        '##
LBL_End:
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        ''DataRow
        Dim i, j As Integer
        Dim str As String
        '
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            '
            For i = 0 To 22
                Select Case i
                    Case 4, 7, 10, 13, 16, 19, 22
                        If CDbl(e.Row.Cells(i - 2).Text) + CDbl(e.Row.Cells(i - 1).Text) > 0 Then
                            str = CDbl(e.Row.Cells(i - 1).Text) / ((CDbl(e.Row.Cells(i - 2).Text) + CDbl(e.Row.Cells(i - 1).Text))) * 10000
                            str = Format(CDbl(str) / 100, "###,###,###.00")
                            e.Row.Cells(i).Text = str & "%"
                        Else
                            e.Row.Cells(i).Text = ".00%"
                        End If
                    Case Else
                End Select
            Next

            'JOY-ADD
            If InStr(e.Row.Cells(23).Text, "營業") > 0 And e.Row.Cells(23).Text <> "營業業務室" And InStr(e.Row.Cells(23).Text, "企劃營業室") <= 0 Then
                For i = 2 To 10
                    e.Row.Cells(i).BackColor = Color.LightGreen
                    Select Case i
                        Case 4, 7, 10, 13, 16, 19, 22
                            If CDbl(Replace(e.Row.Cells(i).Text, "%", "")) > 15 Then
                                e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                                e.Row.Cells(i).BackColor = Color.Pink
                            End If
                        Case Else
                    End Select
                Next
            End If
            If InStr(e.Row.Cells(23).Text, "企劃") > 0 And InStr(e.Row.Cells(23).Text, "企劃營業室") <= 0 Then
                For i = 2 To 10
                    e.Row.Cells(i).BackColor = Color.Khaki
                    Select Case i
                        Case 4, 7, 10, 13, 16, 19, 22
                            If CDbl(Replace(e.Row.Cells(i).Text, "%", "")) > 15 Then
                                e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                                e.Row.Cells(i).BackColor = Color.Pink
                            End If
                        Case Else
                    End Select
                Next
            End If
            '
            e.Row.Cells(23).Visible = False
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯表頭
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCreated1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then

            Dim gv As GridView = sender
            'Dim H1row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            'Dim H2row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            ' 表頭-4行
            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = ""
            H4row.Cells.Add(H4tc0)

            Dim H4tc00 As TableCell = New TableCell
            H4tc00.Text = ""
            H4row.Cells.Add(H4tc00)

            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "YES"
            H4row.Cells.Add(H4tc1)

            Dim H4tc1A As TableCell = New TableCell
            H4tc1A.Text = "NG"
            H4row.Cells.Add(H4tc1A)

            Dim H4tc1B As TableCell = New TableCell
            H4tc1B.Text = "%"
            H4tc1B.BackColor = Color.Blue
            H4row.Cells.Add(H4tc1B)

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = "YES"
            H4row.Cells.Add(H4tc2)

            Dim H4tc2A As TableCell = New TableCell
            H4tc2A.Text = "NG"
            H4row.Cells.Add(H4tc2A)

            Dim H4tc2B As TableCell = New TableCell
            H4tc2B.Text = "%"
            H4tc2B.BackColor = Color.Blue
            H4row.Cells.Add(H4tc2B)

            Dim H4tc3 As TableCell = New TableCell
            H4tc3.Text = "YES"
            H4row.Cells.Add(H4tc3)

            Dim H4tc3A As TableCell = New TableCell
            H4tc3A.Text = "NG"
            H4row.Cells.Add(H4tc3A)

            Dim H4tc3B As TableCell = New TableCell
            H4tc3B.Text = "%"
            H4tc3B.BackColor = Color.Blue
            H4row.Cells.Add(H4tc3B)

            Dim H4tc4 As TableCell = New TableCell
            H4tc4.Text = "YES"
            H4row.Cells.Add(H4tc4)

            Dim H4tc4A As TableCell = New TableCell
            H4tc4A.Text = "NG"
            H4row.Cells.Add(H4tc4A)

            Dim H4tc4B As TableCell = New TableCell
            H4tc4B.Text = "%"
            H4tc4B.BackColor = Color.Blue
            H4row.Cells.Add(H4tc4B)

            Dim H4tc5 As TableCell = New TableCell
            H4tc5.Text = "YES"
            H4row.Cells.Add(H4tc5)

            Dim H4tc5A As TableCell = New TableCell
            H4tc5A.Text = "NG"
            H4row.Cells.Add(H4tc5A)

            Dim H4tc5B As TableCell = New TableCell
            H4tc5B.Text = "%"
            H4tc5B.BackColor = Color.Blue
            H4row.Cells.Add(H4tc5B)

            Dim H4tc6 As TableCell = New TableCell
            H4tc6.Text = "YES"
            H4row.Cells.Add(H4tc6)

            Dim H4tc6A As TableCell = New TableCell
            H4tc6A.Text = "NG"
            H4row.Cells.Add(H4tc6A)

            Dim H4tc6B As TableCell = New TableCell
            H4tc6B.Text = "%"
            H4tc6B.BackColor = Color.Blue
            H4row.Cells.Add(H4tc6B)

            Dim H4tc7 As TableCell = New TableCell
            H4tc7.Text = "YES"
            H4row.Cells.Add(H4tc7)

            Dim H4tc7A As TableCell = New TableCell
            H4tc7A.Text = "NG"
            H4row.Cells.Add(H4tc7A)

            Dim H4tc7B As TableCell = New TableCell
            H4tc7B.Text = "%"
            H4tc7B.BackColor = Color.Blue
            H4row.Cells.Add(H4tc7B)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            ' 表頭-3行
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "部門"
            H3row.Cells.Add(H3tc1)

            Dim H3tc11 As TableCell = New TableCell
            H3tc11.Text = "客戶"
            H3row.Cells.Add(H3tc11)

            Dim H3tc2 As TableCell = New TableCell
            H3tc2.Text = YYMM(1)
            H3tc2.ColumnSpan = 3
            H3row.Cells.Add(H3tc2)

            Dim H3tc3 As TableCell = New TableCell
            H3tc3.Text = YYMM(2)
            H3tc3.ColumnSpan = 3
            H3row.Cells.Add(H3tc3)

            Dim H3tc4 As TableCell = New TableCell
            H3tc4.Text = YYMM(3)
            H3tc4.ColumnSpan = 3
            H3row.Cells.Add(H3tc4)

            Dim H3tc5 As TableCell = New TableCell
            H3tc5.Text = YYMM(4)
            H3tc5.ColumnSpan = 3
            H3row.Cells.Add(H3tc5)

            Dim H3tc6 As TableCell = New TableCell
            H3tc6.Text = YYMM(5)
            H3tc6.ColumnSpan = 3
            H3row.Cells.Add(H3tc6)

            Dim H3tc7 As TableCell = New TableCell
            H3tc7.Text = YYMM(6)
            H3tc7.ColumnSpan = 3
            H3row.Cells.Add(H3tc7)

            Dim H3tc8 As TableCell = New TableCell
            H3tc8.Text = "SUM"
            H3tc8.ColumnSpan = 3
            H3row.Cells.Add(H3tc8)

            gv.Controls(0).Controls.AddAt(0, H3row)
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

        Response.AppendHeader("Content-Disposition", "attachment;filename=IRWNOActivityCustBuyerList01.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        '
        ShowData()
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
    '
    '------------------------------------------------------------------

End Class
