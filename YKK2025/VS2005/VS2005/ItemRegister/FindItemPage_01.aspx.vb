Imports System.Data
Imports System.Drawing

Partial Class FindItemPage_01
    Inherits System.Web.UI.Page
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間

    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900 ' 設定逾時時間

        SetParameter()                  ' 設定共用參數
        If Not IsPostBack Then
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
        Response.Cookies("PGM").Value = "FindItemPage_01.aspx"
    End Sub

    Protected Sub BFindItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFindItem.Click
        DataList()
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        DataList()
        GridView1.PageIndex = e.NewPageIndex
        GridView1.DataBind()
    End Sub

    Sub DataList()
        Dim SQL As String
        '只顯示前50筆
        If DCode.Text <> "" Or DName1.Text <> "" Then
            SQL = "SELECT Top 50 "
            SQL = SQL + "LTRIM(RTRIM(ITMC01)) As Code, "
            SQL = SQL + "LTRIM(RTRIM(IT1I01)) As Name1, "
            SQL = SQL + "LTRIM(RTRIM(IT2I01)) As Name2, "
            SQL = SQL + "LTRIM(RTRIM(IT3I01)) As Name3 "
            SQL = SQL + "From MST_C0100 "
            If DCode.Text <> "" Then
                SQL = SQL + "Where ITMC01 LIKE '%" & DCode.Text.ToUpper & "%' "
            Else
                SQL = SQL + "Where LTRIM(RTRIM(IT1I01))+LTRIM(RTRIM(IT2I01))+LTRIM(RTRIM(IT3I01)) LIKE '%" & DName1.Text.ToUpper & "%' "
                If DName2.Text <> "" Then
                    SQL = SQL + "And LTRIM(RTRIM(IT1I01))+LTRIM(RTRIM(IT2I01))+LTRIM(RTRIM(IT3I01)) LIKE '%" & DName2.Text.ToUpper & "%' "
                End If
                If DName3.Text <> "" Then
                    SQL = SQL + "And LTRIM(RTRIM(IT1I01))+LTRIM(RTRIM(IT2I01))+LTRIM(RTRIM(IT3I01)) LIKE '%" & DName3.Text.ToUpper & "%' "
                End If
                If DName4.Text <> "" Then
                    SQL = SQL + "And LTRIM(RTRIM(IT1I01))+LTRIM(RTRIM(IT2I01))+LTRIM(RTRIM(IT3I01)) LIKE '%" & DName4.Text.ToUpper & "%' "
                End If
            End If
            SQL = SQL + "ORDER BY ITMC01 DESC, IT1I01, IT2I01, IT3I01 "
            '
            Dim dtITEM As DataTable = uDataBase.GetDataTable(SQL)
            If dtITEM.Rows.Count <= 0 Then
                uJavaScript.PopMsg(Me, "[FindItem-0001]:搜尋不到資料, 請確認篩選條件 ! ")
            Else
                GridView1.DataSource = dtITEM
                GridView1.DataBind()
            End If
        End If
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Dim xCode As String = GridView1.SelectedRow.Cells(1).Text
        Dim SQL, Cmd As String
        '
        SQL = "SELECT "
        SQL = SQL + "ITMC01, IT1I01,IT2I01, IT3I01, "
        SQL = SQL + "SIZC01, CHNC01, CLSC01, TAPC01, SLDC01, SFNC01, "
        SQL = SQL + "SL2C01, SE2C01, SF1C01, SF2C01, SF3C01, SF4C01, "
        SQL = SQL + "SF5C01, SF6C01, ST1C01, ST2C01, ST3C01, "
        SQL = SQL + "ST4C01, ST5C01, ST6C01, ST7C01, FMLC01 "
        SQL = SQL + "From MST_C0100 "
        SQL = SQL + "Where ITMC01 = '" & xCode & "' "
        Dim dtITEM As DataTable = uDataBase.GetDataTable(SQL)
        If dtITEM.Rows.Count > 0 Then
            Cmd = "<script>" + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRCode", LTrim(RTrim(dtITEM.Rows(0).Item("ITMC01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRItemName1", LTrim(RTrim(dtITEM.Rows(0).Item("IT1I01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRItemName2", LTrim(RTrim(dtITEM.Rows(0).Item("IT2I01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRItemName3", LTrim(RTrim(dtITEM.Rows(0).Item("IT3I01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSize", LTrim(RTrim(dtITEM.Rows(0).Item("SIZC01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRChain", LTrim(RTrim(dtITEM.Rows(0).Item("CHNC01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRClass", LTrim(RTrim(dtITEM.Rows(0).Item("CLSC01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRTape", LTrim(RTrim(dtITEM.Rows(0).Item("TAPC01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSlider1", LTrim(RTrim(dtITEM.Rows(0).Item("SLDC01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRFinish1", LTrim(RTrim(dtITEM.Rows(0).Item("SFNC01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSlider2", LTrim(RTrim(dtITEM.Rows(0).Item("SL2C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRFinish2", LTrim(RTrim(dtITEM.Rows(0).Item("SE2C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSRequest1", LTrim(RTrim(dtITEM.Rows(0).Item("SF1C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSRequest2", LTrim(RTrim(dtITEM.Rows(0).Item("SF2C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSRequest3", LTrim(RTrim(dtITEM.Rows(0).Item("SF3C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSRequest4", LTrim(RTrim(dtITEM.Rows(0).Item("SF4C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSRequest5", LTrim(RTrim(dtITEM.Rows(0).Item("SF5C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSRequest6", LTrim(RTrim(dtITEM.Rows(0).Item("SF6C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRFamily", LTrim(RTrim(dtITEM.Rows(0).Item("FMLC01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRST1", LTrim(RTrim(dtITEM.Rows(0).Item("ST1C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRST2", LTrim(RTrim(dtITEM.Rows(0).Item("ST2C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRST3", LTrim(RTrim(dtITEM.Rows(0).Item("ST3C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRST4", LTrim(RTrim(dtITEM.Rows(0).Item("ST4C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRST5", LTrim(RTrim(dtITEM.Rows(0).Item("ST5C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRST6", LTrim(RTrim(dtITEM.Rows(0).Item("ST6C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRST7", LTrim(RTrim(dtITEM.Rows(0).Item("ST7C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSize", LTrim(RTrim(dtITEM.Rows(0).Item("SIZC01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DChain", LTrim(RTrim(dtITEM.Rows(0).Item("CHNC01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DClass", LTrim(RTrim(dtITEM.Rows(0).Item("CLSC01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DFamily", LTrim(RTrim(dtITEM.Rows(0).Item("FMLC01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DST1", LTrim(RTrim(dtITEM.Rows(0).Item("ST1C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DST2", LTrim(RTrim(dtITEM.Rows(0).Item("ST2C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DST3", LTrim(RTrim(dtITEM.Rows(0).Item("ST3C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DST4", LTrim(RTrim(dtITEM.Rows(0).Item("ST4C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DST5", LTrim(RTrim(dtITEM.Rows(0).Item("ST5C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DST6", LTrim(RTrim(dtITEM.Rows(0).Item("ST6C01")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DST7", LTrim(RTrim(dtITEM.Rows(0).Item("ST7C01")))) + _
                        "window.close();" + _
                      "</script>"
            Response.Write(Cmd)
        End If
    End Sub
End Class
