Imports System.Data
Imports System.Drawing

Partial Class FindItemPage
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim wUserID As String = ""      '申請者

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
        wUserID = Request.QueryString("pUserID")
        Response.Cookies("PGM").Value = "FindItemPage.aspx"
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
            SQL = SQL + "LTRIM(RTRIM(ITEM)) As Code, "
            SQL = SQL + "LTRIM(RTRIM(ITEM_NAME1)) As Name1, "
            SQL = SQL + "LTRIM(RTRIM(ITEM_NAME2)) As Name2, "
            SQL = SQL + "LTRIM(RTRIM(ITEM_NAME3)) As Name3 "
            SQL = SQL + "From MST_ITEM "
            If DCode.Text <> "" Then
                SQL = SQL + "Where ITEM LIKE '%" & DCode.Text.ToUpper & "%' "
            Else
                SQL = SQL + "Where LTRIM(RTRIM(ITEM_NAME1))+LTRIM(RTRIM(ITEM_NAME2))+LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DName1.Text.ToUpper & "%' "
                If DName2.Text <> "" Then
                    SQL = SQL + "And LTRIM(RTRIM(ITEM_NAME1))+LTRIM(RTRIM(ITEM_NAME2))+LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DName2.Text.ToUpper & "%' "
                End If
                If DName3.Text <> "" Then
                    SQL = SQL + "And LTRIM(RTRIM(ITEM_NAME1))+LTRIM(RTRIM(ITEM_NAME2))+LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DName3.Text.ToUpper & "%' "
                End If
                If DName4.Text <> "" Then
                    SQL = SQL + "And LTRIM(RTRIM(ITEM_NAME1))+LTRIM(RTRIM(ITEM_NAME2))+LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DName4.Text.ToUpper & "%' "
                End If
            End If
            SQL = SQL + "And NoDisp <> '1' "
            SQL = SQL + "ORDER BY ITEM DESC, ITEM_NAME1, ITEM_NAME2, ITEM_NAME3 "
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
        SQL = SQL + "ITEM, ITEM_NAME1,ITEM_NAME2, ITEM_NAME3, "
        SQL = SQL + "SIZE, CHAINTYPE, CLASS, TAPE_, SLIDER, FINISH, "
        SQL = SQL + "SLIDER2, FINISH2, SPECIAL1, SPECIAL2, SPECIAL3, SPECIAL4, "
        SQL = SQL + "SPECIAL5, SPECIAL6, ST1CA0, ST2CA0, ST3CA0, "
        SQL = SQL + "ST4CA0, ST5CA0, ST6CA0, ST7CA0, NODISP, FMLCA0,'' as buyer "
        SQL = SQL + "From MST_ITEM "
        SQL = SQL + "Where ITEM = '" & xCode & "' "
        Dim dtITEM As DataTable = uDataBase.GetDataTable(SQL)
        If dtITEM.Rows.Count > 0 Then
            Cmd = "<script>" + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRCode", LTrim(RTrim(dtITEM.Rows(0).Item("ITEM")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRItemName1", LTrim(RTrim(dtITEM.Rows(0).Item("ITEM_NAME1")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRItemName2", LTrim(RTrim(dtITEM.Rows(0).Item("ITEM_NAME2")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRItemName3", LTrim(RTrim(dtITEM.Rows(0).Item("ITEM_NAME3")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSize", LTrim(RTrim(dtITEM.Rows(0).Item("SIZE")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRChain", LTrim(RTrim(dtITEM.Rows(0).Item("CHAINTYPE")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRClass", LTrim(RTrim(dtITEM.Rows(0).Item("CLASS")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRTape", LTrim(RTrim(dtITEM.Rows(0).Item("TAPE_")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSlider1", LTrim(RTrim(dtITEM.Rows(0).Item("SLIDER")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRFinish1", LTrim(RTrim(dtITEM.Rows(0).Item("FINISH")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSlider2", LTrim(RTrim(dtITEM.Rows(0).Item("SLIDER2")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRFinish2", LTrim(RTrim(dtITEM.Rows(0).Item("FINISH2")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSRequest1", LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL1")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSRequest2", LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL2")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSRequest3", LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL3")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSRequest4", LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL4")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSRequest5", LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL5")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRSRequest6", LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL6")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRFamily", LTrim(RTrim(dtITEM.Rows(0).Item("FMLCA0")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRST1", LTrim(RTrim(dtITEM.Rows(0).Item("ST1CA0")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRST2", LTrim(RTrim(dtITEM.Rows(0).Item("ST2CA0")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRST3", LTrim(RTrim(dtITEM.Rows(0).Item("ST3CA0")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRST4", LTrim(RTrim(dtITEM.Rows(0).Item("ST4CA0")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRST5", LTrim(RTrim(dtITEM.Rows(0).Item("ST5CA0")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRST6", LTrim(RTrim(dtITEM.Rows(0).Item("ST6CA0")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRST7", LTrim(RTrim(dtITEM.Rows(0).Item("ST7CA0")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DRNoDisplay", LTrim(RTrim(dtITEM.Rows(0).Item("NODISP")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSize", LTrim(RTrim(dtITEM.Rows(0).Item("SIZE")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DChain", LTrim(RTrim(dtITEM.Rows(0).Item("CHAINTYPE")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DClass", LTrim(RTrim(dtITEM.Rows(0).Item("CLASS")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTape", LTrim(RTrim(dtITEM.Rows(0).Item("TAPE_")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSlider1", LTrim(RTrim(dtITEM.Rows(0).Item("SLIDER")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DFinish1", LTrim(RTrim(dtITEM.Rows(0).Item("FINISH")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSlider2", LTrim(RTrim(dtITEM.Rows(0).Item("SLIDER2")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DFinish2", LTrim(RTrim(dtITEM.Rows(0).Item("FINISH2")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSRequest1", LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL1")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSRequest2", LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL2")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSRequest3", LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL3")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSRequest4", LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL4")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSRequest5", LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL5")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSRequest6", LTrim(RTrim(dtITEM.Rows(0).Item("SPECIAL6")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DFamily", LTrim(RTrim(dtITEM.Rows(0).Item("FMLCA0")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DST1", LTrim(RTrim(dtITEM.Rows(0).Item("ST1CA0")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DST2", LTrim(RTrim(dtITEM.Rows(0).Item("ST2CA0")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DST3", LTrim(RTrim(dtITEM.Rows(0).Item("ST3CA0")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DST4", LTrim(RTrim(dtITEM.Rows(0).Item("ST4CA0")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DST5", LTrim(RTrim(dtITEM.Rows(0).Item("ST5CA0")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DST6", LTrim(RTrim(dtITEM.Rows(0).Item("ST6CA0")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DST7", LTrim(RTrim(dtITEM.Rows(0).Item("ST7CA0")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DBuyer", LTrim(RTrim(dtITEM.Rows(0).Item("buyer")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DBuyerCode", LTrim(RTrim(dtITEM.Rows(0).Item("buyer")))) + _
                      String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DMemo", LTrim(RTrim(dtITEM.Rows(0).Item("buyer")))) + _
                        "window.close();" + _
                      "</script>"
            Response.Write(Cmd)


        


        End If
    End Sub
End Class
