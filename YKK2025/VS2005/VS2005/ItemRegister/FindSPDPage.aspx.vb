Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class FindSPDPage
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim uJavaScript As New Utility.JScript
    Dim fpObj As New ForProject     '操作db的物件
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
    Dim SLDKey1 As String = ""
    Dim SLDKey2 As String = ""
    Dim FINKey1 As String = ""
    Dim FINKey2 As String = ""
    Dim UserID As String = ""

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "INQ_IRWSPDSlider.aspx"

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
        SLDKey1 = Request.QueryString("pSLDKey1")
        SLDKey2 = Request.QueryString("pSLDKey2")
        FINKey1 = Request.QueryString("pFINKey1")
        FINKey2 = Request.QueryString("pFINKey2")
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
        GridView2.Visible = False
        DSLDKey1.Text = ""
        DSLDKey2.Text = ""
        DFINKey1.Text = ""
        DFINKey2.Text = ""
        '-----------------------------------------------------------------
        '-- KEY
        '-----------------------------------------------------------------
        If SLDKey1 <> "" Then DSLDKey1.Text = SLDKey1
        If SLDKey2 <> "" Then DSLDKey2.Text = SLDKey2
        If FINKey1 <> "" Then DFINKey1.Text = FINKey1
        If FINKey2 <> "" Then DFINKey2.Text = FINKey2
        '
    End Sub

    Sub DataList()
        Dim SQL, xNoData As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '
        xNoData = "YES"
        If DSLDKey1.Text <> "" And DSLDKey2.Text <> "" And DFINKey1.Text <> "" And DFINKey2.Text <> "" Then
            OleDbConnection1.Open()
            SQL = "SELECT "
            SQL = SQL & "SYS, STATUS, NO, FURL, BUYER, SPEC, SLIDERGRCODE, PERSON, APPLYDATE "
            SQL = SQL & "FROM V_ISOSIRWSPD_03 "
            SQL = SQL & "Where SPEC <> '' "
            SQL = SQL & "And SLIDERGRCODE <> '' "
            '
            SQL = SQL & " And ( "
            SQL = SQL & "     ( REPLACE(REPLACE(REPLACE(SPEC+SLIDERGRCODE,']',''),'[',''),' ','') Like '%" & Replace(DSLDKey1.Text, "-", "") & "%'"
            SQL = SQL & "       Or REPLACE(REPLACE(REPLACE(SPEC+SLIDERGRCODE,']',''),'[',''),' ','') Like '%" & Replace(DSLDKey2.Text, "-", "") & "%') "
            SQL = SQL & " ) "
            '
            SQL = SQL & "Group By SYS, STATUS, NO, FURL, BUYER, SPEC, SLIDERGRCODE, PERSON, APPLYDATE "
            SQL = SQL & "Order By SPEC, SLIDERGRCODE "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "SLIDER")
            If DBDataSet1.Tables("SLIDER").Rows.Count > 0 Then
                xNoData = "NO"
            Else
                DBDataSet1.Clear()
                SQL = "SELECT "
                SQL = SQL & "SYS, STATUS, NO, FURL, BUYER, SPEC, SLIDERGRCODE, PERSON, APPLYDATE "
                SQL = SQL & "FROM V_ISOSIRWSPD_03 "
                SQL = SQL & "Where SPEC <> '' "
                SQL = SQL & "And SLIDERGRCODE <> '' "
                '
                SQL = SQL & " And ( "
                SQL = SQL & "     ( REPLACE(REPLACE(REPLACE(SPEC+SLIDERGRCODE,']',''),'[',''),' ','') Like '%" & Replace(DFINKey1.Text, "-", "") & "%'"
                SQL = SQL & "       Or REPLACE(REPLACE(REPLACE(SPEC+SLIDERGRCODE,']',''),'[',''),' ','') Like '%" & Replace(DFINKey2.Text, "-", "") & "%') "
                SQL = SQL & " ) "
                '
                SQL = SQL & "Group By SYS, STATUS, NO, FURL, BUYER, SPEC, SLIDERGRCODE, PERSON, APPLYDATE "
                SQL = SQL & "Order By SPEC, SLIDERGRCODE "
                '
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "SLIDER")
                If DBDataSet1.Tables("SLIDER").Rows.Count > 0 Then
                    xNoData = "NO"
                End If
            End If

            GridView1.DataSource = DBDataSet1
            GridView1.DataBind()
            GridView1.Visible = True
            OleDbConnection1.Close()
            '
            OleDbConnection1.Open()
            '
            SQL = "SELECT "
            SQL = SQL & "SYS, STATUS, NO, FURL, BUYER, SPEC, SLIDERGRCODE, PERSON, APPLYDATE "
            SQL = SQL & "FROM V_ISOSIRWSPD_03 "
            SQL = SQL & "Where SPEC <> '' "
            SQL = SQL & "And SLIDERGRCODE <> '' "
            SQL = SQL & "And ("
            SQL = SQL & " ('" & Replace(DSLDKey1.Text, " ", "") & "' like '%' + REPLACE(REPLACE(REPLACE(SPEC,']',''),'[',''),' ','') + '%' "
            SQL = SQL & "     AND '" & Replace(DSLDKey1.Text, " ", "") & "' like '%' +SUBSTRING(REPLACE(REPLACE(REPLACE(SLIDERGRCODE,']',''),'[',''),' ',''),1,4) + '%') "
            SQL = SQL & " or "
            SQL = SQL & " ('" & Replace(DSLDKey1.Text, " ", "") & "' like '%' +REPLACE(REPLACE(REPLACE(SLIDERGRCODE,']',''),'[',''),' ','') + '%') "
            SQL = SQL & " or "
            SQL = SQL & " ('" & Replace(DSLDKey1.Text, " ", "") & "' like '%' +REPLACE(REPLACE(REPLACE(SPEC,']',''),'[',''),' ','') + '%' "
            SQL = SQL & "     AND REPLACE(REPLACE(REPLACE(SLIDERGRCODE,']',''),'[',''),' ','') LIKE '%YG%') "
            SQL = SQL & "    ) "
            SQL = SQL & "Group By SYS, STATUS, NO, FURL, BUYER, SPEC, SLIDERGRCODE, PERSON, APPLYDATE "
            SQL = SQL & "Order By len(SLIDERGRCODE) desc, len(SPEC) desc "
            '
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet2, "SLIDER1")
            If DBDataSet2.Tables("SLIDER1").Rows.Count > 0 Then
                xNoData = "NO"
            End If
            GridView2.DataSource = DBDataSet2
            GridView2.DataBind()
            GridView2.Visible = True
            OleDbConnection1.Close()
        Else
            uJavaScript.PopMsg(Me, "搜尋欄位(黃色)，不可空白!")
        End If
        '
        If xNoData = "YES" Then uJavaScript.PopMsg(Me, "無資料!")
        '
    End Sub
    '
    Protected Sub BSEARCH1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSEARCH1.Click
        DataList()
    End Sub


    Protected Sub BSEARCH2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSEARCH2.Click
        DataList()
    End Sub


    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
            e.Row.Cells(9).Visible = False
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            e.Row.Cells(9).Visible = False
        End If
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Dim SQL, Cmd As String
        Dim xCode As String = GridView1.SelectedRow.Cells(9).Text
        '
        SQL = "SELECT TOP 1 "
        SQL = SQL & "NO, FURL "
        SQL = SQL & "FROM V_ISOSIRWSPD_03 "
        SQL = SQL & "Where NO = '" & xCode & "' "
        Dim dtSPD As DataTable = uDataBase.GetDataTable(SQL)
        If dtSPD.Rows.Count > 0 Then
            Cmd = "<script>" + _
              String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSPDNO", LTrim(RTrim(dtSPD.Rows(0).Item("NO")))) + _
              String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSPDNOUrl", LTrim(RTrim(dtSPD.Rows(0).Item("FURL")))) + _
                "window.close();" + _
              "</script>"
            Response.Write(Cmd)
        End If
    End Sub

    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound

        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
            e.Row.Cells(9).Visible = False
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            e.Row.Cells(9).Visible = False
        End If
    End Sub


    Protected Sub GridView2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView2.SelectedIndexChanged
        Dim SQL, Cmd As String
        Dim xCode As String = GridView2.SelectedRow.Cells(9).Text
        '
        SQL = "SELECT TOP 1 "
        SQL = SQL & "NO, FURL "
        SQL = SQL & "FROM V_ISOSIRWSPD_03 "
        SQL = SQL & "Where NO = '" & xCode & "' "
        Dim dtSPD As DataTable = uDataBase.GetDataTable(SQL)
        If dtSPD.Rows.Count > 0 Then
            Cmd = "<script>" + _
              String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSPDNO", LTrim(RTrim(dtSPD.Rows(0).Item("NO")))) + _
              String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSPDNOUrl", LTrim(RTrim(dtSPD.Rows(0).Item("FURL")))) + _
                "window.close();" + _
              "</script>"
            Response.Write(Cmd)
        End If

    End Sub
End Class
