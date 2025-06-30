Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class IRWInqItemSA01
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim uJavaScript As New Utility.JScript
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String
    Dim UserID As String            'UserID
    Dim xCount51 As Integer
    Dim xCount52 As Integer
    Dim xCount53 As Integer
    Dim xCount54 As Integer
    Dim xCount51SA As Integer
    Dim xCount52SA As Integer
    Dim xCount53SA As Integer
    Dim xCount54SA As Integer
    Dim xCount51ND As Integer
    Dim xCount52ND As Integer
    Dim xCount53ND As Integer
    Dim xCount54ND As Integer

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "IRWInqSA01.aspx"

        If Not Me.IsPostBack Then
            SetParameter()          '設定共用參數
            DataList()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        UserID = Request.QueryString("pUserID")             'UserID
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
        GridView2.Visible = False
        xCount51 = 0
        xCount52 = 0
        xCount53 = 0
        xCount54 = 0
        xCount51SA = 0
        xCount52SA = 0
        xCount53SA = 0
        xCount54SA = 0
        xCount51ND = 0
        xCount52ND = 0
        xCount53ND = 0
        xCount54ND = 0

        DKNo.Text = ""
        DKNo.Style("left") = -500 & "px"

        DKUser.Text = UserID
        DKUser.Style("left") = -500 & "px"
        BInq.Style("left") = -500 & "px"

        If InStr("SL034/MK045/IT003/IT004/MK028/MK035/MK019", UCase(UserID)) > 0 Then
            DKNo.Style("left") = 680 & "px"

            DKUser.Style("left") = 544 & "px"
            BInq.Style("left") = 760 & "px"
        End If
        '
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '
        SQL = "SELECT FormNo, "
        SQL = SQL & "CASE WHEN FORMNO='001151' THEN '營業申請單' "
        SQL = SQL & "     WHEN FORMNO='001152' THEN 'ZIP申請單'  "
        SQL = SQL & "     WHEN FORMNO='001153' THEN 'SLD申請單'  "
        SQL = SQL & "     ELSE 'CH申請單' END AS FORM, "
        SQL = SQL & "NO, convert(varchar, [DATE], 111) As date, CODE, '' AS ITEMNAME, '' AS STATUS, '' AS NODISPLAY "
        SQL = SQL & "FROM V_IRWInqItemSA "
        SQL = SQL & "WHERE CreateUser = '" & DKUser.Text & "' "
        '
        If DKNo.Text <> "" Then
            SQL = SQL & "AND NO LIKE '%" & DKNo.Text & "%' "
        End If
        '
        SQL = SQL & "Order By FORMNO, NO "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "ITEM")
        If DBDataSet1.Tables("ITEM").Rows.Count > 0 Then
            GridView1.DataSource = DBDataSet1
            GridView1.DataBind()
            GridView1.Visible = True
            '
            SQL = "SELECT "
            SQL = SQL & "NAME, "
            SQL = SQL & "CASE WHEN FORMNO='001151' THEN '營業申請單' "
            SQL = SQL & "     WHEN FORMNO='001152' THEN 'ZIP申請單'  "
            SQL = SQL & "     WHEN FORMNO='001153' THEN 'SLD申請單'  "
            SQL = SQL & "     ELSE 'CH申請單' END AS FORM, "
            SQL = SQL & "Count(*) As AppCount, 0 As ComCount, 0 As SACount "
            SQL = SQL & "FROM V_IRWInqItemSA "
            SQL = SQL & "WHERE CreateUser = '" & DKUser.Text & "' "
            '
            If DKNo.Text <> "" Then
                SQL = SQL & "AND NO LIKE '%" & DKNo.Text & "%' "
            End If
            '
            SQL = SQL & "group by name, formno "
            SQL = SQL & "order by name, formno "
            Dim dt As DataTable = uDataBase.GetDataTable(SQL)
            If dt.Rows.Count > 0 Then
                GridView2.Visible = True
                GridView2.DataSource = dt
                GridView2.DataBind()
            End If
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
        End If
        OleDbConnection1.Close()
        '
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            ' 表頭-8行
            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "Form"
            H4row.Cells.Add(H4tc0)

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = "No."
            H4row.Cells.Add(H4tc2)

            Dim H4tc3 As TableCell = New TableCell
            H4tc3.Text = "Apply Date"
            H4row.Cells.Add(H4tc3)

            Dim H4tc4 As TableCell = New TableCell
            H4tc4.Text = "Code"
            H4row.Cells.Add(H4tc4)

            Dim H4tc5 As TableCell = New TableCell
            H4tc5.Text = "Description"
            H4row.Cells.Add(H4tc5)

            Dim H4tc6 As TableCell = New TableCell
            H4tc6.Text = "*" & "<BR>" & "[SA]"
            H4tc6.BackColor = Color.Blue
            H4row.Cells.Add(H4tc6)

            Dim H4tc7 As TableCell = New TableCell
            H4tc7.Text = "*" & "<BR>" & "[ND]"
            H4tc7.BackColor = Color.Blue
            H4row.Cells.Add(H4tc7)

            gv.Controls(0).Controls.AddAt(0, H4row)
        End If

    End Sub

    '
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim SQL As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
        '
        cn.ConnectionString = ConnectString

        If (e.Row.RowType = DataControlRowType.DataRow) Then
            SQL = "SELECT "
            SQL = SQL & "ITMCA0, IT1IA0 || ' ' || IT2IA0 || ' ' || IT3IA0 AS ITEMNAME, "
            SQL = SQL & "NDPCA0, SAVFA0 "
            SQL = SQL & "FROM WAVEDLIB.FA000 "
            SQL = SQL + "WHERE ITMCA0 = '" & e.Row.Cells(3).Text & "' "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, cn)
            DBAdapter1.Fill(ds, "FA000")
            If ds.Tables("FA000").Rows.Count > 0 Then
                e.Row.Cells(4).Text = ds.Tables("FA000").Rows(0).Item("ITEMNAME")
                '
                If ds.Tables("FA000").Rows(0).Item("SAVFA0") <> "" Then
                    e.Row.Cells(5).Text = "OK"
                    e.Row.Cells(5).Attributes.Add("style", "border:1px solid red ")
                    e.Row.Cells(5).BackColor = Color.Pink
                    '
                    If InStr(e.Row.Cells(0).Text, "營業") > 0 Then xCount51SA = xCount51SA + 1
                    If InStr(e.Row.Cells(0).Text, "ZIP") > 0 Then xCount52SA = xCount52SA + 1
                    If InStr(e.Row.Cells(0).Text, "SLD") > 0 Then xCount53SA = xCount53SA + 1
                    If InStr(e.Row.Cells(0).Text, "CH") > 0 Then xCount54SA = xCount54SA + 1
                Else
                    e.Row.Cells(5).Text = "未設定"
                End If
                '
                If ds.Tables("FA000").Rows(0).Item("NDPCA0") <> "" Then
                    e.Row.Cells(6).Text = "ND"
                    e.Row.Cells(6).Attributes.Add("style", "border:1px solid red ")
                    e.Row.Cells(6).BackColor = Color.Pink
                    '
                    If InStr(e.Row.Cells(0).Text, "營業") > 0 Then xCount51ND = xCount51ND + 1
                    If InStr(e.Row.Cells(0).Text, "ZIP") > 0 Then xCount52ND = xCount52ND + 1
                    If InStr(e.Row.Cells(0).Text, "SLD") > 0 Then xCount53ND = xCount53ND + 1
                    If InStr(e.Row.Cells(0).Text, "CH") > 0 Then xCount54ND = xCount54ND + 1
                Else
                    e.Row.Cells(6).Text = ""
                End If
                '
                If InStr(e.Row.Cells(0).Text, "營業") > 0 Then xCount51 = xCount51 + 1
                If InStr(e.Row.Cells(0).Text, "ZIP") > 0 Then xCount52 = xCount52 + 1
                If InStr(e.Row.Cells(0).Text, "SLD") > 0 Then xCount53 = xCount53 + 1
                If InStr(e.Row.Cells(0).Text, "CH") > 0 Then xCount54 = xCount54 + 1
            Else
                e.Row.Cells(5).Text = "等待中"
            End If
        End If
        '
        cn.Close()
    End Sub

    Protected Sub GridView2_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()
            '-----------------------------------------
            ' 表頭-8行
            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "Name"
            H4row.Cells.Add(H4tc0)

            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "Form"
            H4row.Cells.Add(H4tc1)

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = "申請件"
            H4row.Cells.Add(H4tc2)

            Dim H4tc3 As TableCell = New TableCell
            H4tc3.Text = "完成件"
            H4tc3.BackColor = Color.Green
            H4row.Cells.Add(H4tc3)

            Dim H4tc4 As TableCell = New TableCell
            H4tc4.Text = "[SA]" & "<BR>" & "完成件"
            H4tc4.BackColor = Color.Blue
            H4row.Cells.Add(H4tc4)

            gv.Controls(0).Controls.AddAt(0, H4row)
            '-----------------------------------------
            ' 表頭-行
            '
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "WINGS ITEM [SA] STATUS"
            H3tc1.ColumnSpan = 5
            H3tc1.BackColor = Color.Red
            H3row.Cells.Add(H3tc1)
            '
            gv.Controls(0).Controls.AddAt(0, H3row)
        End If

    End Sub

    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        Dim SQL As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
        '
        cn.ConnectionString = ConnectString

        If (e.Row.RowType = DataControlRowType.DataRow) Then
            If InStr(e.Row.Cells(1).Text, "營業") > 0 Then e.Row.Cells(3).Text = CStr(xCount51)
            If InStr(e.Row.Cells(1).Text, "ZIP") > 0 Then e.Row.Cells(3).Text = CStr(xCount52)
            If InStr(e.Row.Cells(1).Text, "SLD") > 0 Then e.Row.Cells(3).Text = CStr(xCount53)
            If InStr(e.Row.Cells(1).Text, "CH") > 0 Then e.Row.Cells(3).Text = CStr(xCount54)
            '
            If xCount51ND > 0 Then e.Row.Cells(3).Text = e.Row.Cells(3).Text & " ND(" & CStr(xCount51ND) & ")"
            If xCount52ND > 0 Then e.Row.Cells(3).Text = e.Row.Cells(3).Text & " ND(" & CStr(xCount52ND) & ")"
            If xCount53ND > 0 Then e.Row.Cells(3).Text = e.Row.Cells(3).Text & " ND(" & CStr(xCount53ND) & ")"
            If xCount54ND > 0 Then e.Row.Cells(3).Text = e.Row.Cells(3).Text & " ND(" & CStr(xCount54ND) & ")"
            '
            e.Row.Cells(3).Attributes.Add("style", "border:1px solid red ")
            e.Row.Cells(3).BackColor = Color.Pink
            '
            If InStr(e.Row.Cells(1).Text, "營業") > 0 Then e.Row.Cells(4).Text = CStr(xCount51SA)
            If InStr(e.Row.Cells(1).Text, "ZIP") > 0 Then e.Row.Cells(4).Text = CStr(xCount52SA)
            If InStr(e.Row.Cells(1).Text, "SLD") > 0 Then e.Row.Cells(4).Text = CStr(xCount53SA)
            If InStr(e.Row.Cells(1).Text, "CH") > 0 Then e.Row.Cells(4).Text = CStr(xCount54SA)
            e.Row.Cells(4).ForeColor = Color.Red
        End If

    End Sub

    Protected Sub BInq_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BInq.Click
        DataList()
    End Sub

End Class
