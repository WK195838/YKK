Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Collections.Generic

Partial Class OpenNoPickerForSampleSheet
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                              '設定參數
        '
        If Not Me.IsPostBack Then   '不是PostBack
            DKey.Text = ""
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
        Server.ScriptTimeout = 900                                                                  '設定逾時時間
        Response.Cookies("PGM").Value = "OpenNoPickerForSampleSheet.aspx"                           '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")                                           '現在日時
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BKey_Click)
    '**     搜尋按鈕按下時
    '**
    '*****************************************************************
    Protected Sub BKey_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BKey.Click
        If DKey.Text <> "" Then
            DataList()
        Else
            uJavaScript.PopMsg(Me, "無輸入資料無法搜尋,請確認!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(DataList)
    '**     開發資料列表
    '**
    '*****************************************************************
    Protected Sub DataList()
        Dim SQL As String
        SQL = "Select No, CodeNo, DevNo, '...' as NoDesc, "
        SQL &= "Case Sts When 0 Then '開發中' When 1 Then '完成(OK)' When 2 Then '完成(NG)' Else '取消/中止' End As StsDesc, "
        SQL &= "'CommissionSheet_02.aspx?pFormNo=' + FormNo + '&pFormSno=' + str(FormSno, Len(FormSno)) As URL "
        SQL &= "From V_CommissionSheet_02 "
        SQL &= "Where (No Like '%" & DKey.Text & "%') "
        SQL &= "   or (CodeNo Like '%" & DKey.Text & "%') "
        SQL &= "   or (DevNo Like '%" & DKey.Text & "%') "
        SQL &= "Order by APPDate Desc"
        Dim dtCommissionSheet As DataTable = uDataBase.GetDataTable(SQL)
        If dtCommissionSheet.Rows.Count > 0 Then
            GridView1.DataSource = dtCommissionSheet
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到開發資料,請確認!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GridView1_SelectedIndexChanged)
    '**     取得所選擇開發資料
    '**
    '*****************************************************************
    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Dim xNo As String = GridView1.SelectedRow.Cells(2).Text
        Dim SQL, Cmd As String
        Dim xDATE, xTHUPCOL, xTHLOCOL As String
        '
        SQL = "Select * From FS_SampleSheet "
        SQL &= "Where No = '" & xNo & "' "
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            '開發委託
            SQL = "Select * From F_CommissionSheet "
            SQL &= "Where No = '" & xNo & "' "
            Dim dtCommissionData As DataTable = uDataBase.GetDataTable(SQL)
            If dtCommissionData.Rows.Count > 0 Then
                '縫工上線
                xTHUPCOL = ""
                If dtCommissionData.Rows(0).Item("THUPCOL") <> "" Or dtCommissionData.Rows(0).Item("THUPCOLNO") <> "" Then
                    If xTHUPCOL <> "" Then
                        xTHUPCOL &= ",上:"
                    Else
                        xTHUPCOL &= "上:"
                    End If
                    If dtCommissionData.Rows(0).Item("THUPCOL") <> "" Then xTHUPCOL &= dtCommissionData.Rows(0).Item("THUPCOL")
                    If dtCommissionData.Rows(0).Item("THUPCOLNO") <> "" Then xTHUPCOL &= dtCommissionData.Rows(0).Item("THUPCOLNO")
                End If
                If dtCommissionData.Rows(0).Item("THLUPCOL") <> "" Or dtCommissionData.Rows(0).Item("THLUPCOLNO") <> "" Then
                    If xTHUPCOL <> "" Then
                        xTHUPCOL &= ",上左:"
                    Else
                        xTHUPCOL &= "上左:"
                    End If
                    If dtCommissionData.Rows(0).Item("THLUPCOL") <> "" Then xTHUPCOL &= dtCommissionData.Rows(0).Item("THLUPCOL")
                    If dtCommissionData.Rows(0).Item("THLUPCOLNO") <> "" Then xTHUPCOL &= dtCommissionData.Rows(0).Item("THLUPCOLNO")
                End If
                If dtCommissionData.Rows(0).Item("THRUPCOL") <> "" Or dtCommissionData.Rows(0).Item("THRUPCOLNO") <> "" Then
                    If xTHUPCOL <> "" Then
                        xTHUPCOL &= ",上右:"
                    Else
                        xTHUPCOL &= "上右:"
                    End If
                    If dtCommissionData.Rows(0).Item("THRUPCOL") <> "" Then xTHUPCOL &= dtCommissionData.Rows(0).Item("THRUPCOL")
                    If dtCommissionData.Rows(0).Item("THRUPCOLNO") <> "" Then xTHUPCOL &= dtCommissionData.Rows(0).Item("THRUPCOLNO")
                End If
                '縫工下線
                xTHLOCOL = ""
                If dtCommissionData.Rows(0).Item("THLOCOL") <> "" Or dtCommissionData.Rows(0).Item("THLOCOLNO") <> "" Then
                    If xTHLOCOL <> "" Then
                        xTHLOCOL &= ",下:"
                    Else
                        xTHLOCOL &= "下:"
                    End If
                    If dtCommissionData.Rows(0).Item("THLOCOL") <> "" Then xTHLOCOL &= dtCommissionData.Rows(0).Item("THLOCOL")
                    If dtCommissionData.Rows(0).Item("THLOCOLNO") <> "" Then xTHLOCOL &= dtCommissionData.Rows(0).Item("THLOCOLNO")
                End If
                If dtCommissionData.Rows(0).Item("THLLOCOL") <> "" Or dtCommissionData.Rows(0).Item("THLLOCOLNO") <> "" Then
                    If xTHLOCOL <> "" Then
                        xTHLOCOL &= ",下左:"
                    Else
                        xTHLOCOL &= "下左:"
                    End If
                    If dtCommissionData.Rows(0).Item("THLLOCOL") <> "" Then xTHLOCOL &= dtCommissionData.Rows(0).Item("THLLOCOL")
                    If dtCommissionData.Rows(0).Item("THLLOCOLNO") <> "" Then xTHLOCOL &= dtCommissionData.Rows(0).Item("THLLOCOLNO")
                End If
                If dtCommissionData.Rows(0).Item("THRLOCOL") <> "" Or dtCommissionData.Rows(0).Item("THRLOCOLNO") <> "" Then
                    If xTHLOCOL <> "" Then
                        xTHLOCOL &= ",下右:"
                    Else
                        xTHLOCOL &= "下右:"
                    End If
                    If dtCommissionData.Rows(0).Item("THRLOCOL") <> "" Then xTHLOCOL &= dtCommissionData.Rows(0).Item("THRLOCOL")
                    If dtCommissionData.Rows(0).Item("THRLOCOLNO") <> "" Then xTHLOCOL &= dtCommissionData.Rows(0).Item("THRLOCOLNO")
                End If
                '工程
                SQL = "Select * From FS_ManufactureSheet "
                SQL &= "Where No = '" & xNo & "' "
                Dim dtManufData As DataTable = uDataBase.GetDataTable(SQL)
                If dtManufData.Rows.Count > 0 Then
                    '發行日
                    xDATE = CDate(NowDateTime).ToString("yyyy/MM/dd")
                    '傳回
                    Cmd = "<script>" + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DONO", LTrim(RTrim(xNo))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DAPPBUYER", LTrim(RTrim(dtData.Rows(0).Item("APPBUYER").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSIZENO", LTrim(RTrim(dtData.Rows(0).Item("SIZENO").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DITEM", LTrim(RTrim(dtData.Rows(0).Item("ITEM").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DCODENO", LTrim(RTrim(dtData.Rows(0).Item("CODENO").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTAWIDTH", LTrim(RTrim(dtData.Rows(0).Item("TAWIDTH").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DDEVNO", LTrim(RTrim(dtData.Rows(0).Item("DEVNO").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DDEVPRD", LTrim(RTrim(dtData.Rows(0).Item("DEVPRD").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTACOL", LTrim(RTrim(dtData.Rows(0).Item("TACOL").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTALINE", LTrim(RTrim(dtData.Rows(0).Item("TALINE").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DECOL", LTrim(RTrim(dtData.Rows(0).Item("ECOL").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DCCOL", LTrim(RTrim(dtData.Rows(0).Item("CCOL").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHUPCOL", LTrim(RTrim(xTHUPCOL))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHLOCOL", LTrim(RTrim(xTHLOCOL))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTNLITEM", LTrim(RTrim(dtData.Rows(0).Item("TNLITEM").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTNRITEM", LTrim(RTrim(dtData.Rows(0).Item("TNRITEM").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTSLITEM", LTrim(RTrim(dtData.Rows(0).Item("TSLITEM").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTSRITEM", LTrim(RTrim(dtData.Rows(0).Item("TSRITEM").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTDLITEM", LTrim(RTrim(dtData.Rows(0).Item("TDLITEM").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTDRITEM", LTrim(RTrim(dtData.Rows(0).Item("TDRITEM").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DCNITEM", LTrim(RTrim(dtData.Rows(0).Item("CNITEM").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DCSITEM", LTrim(RTrim(dtData.Rows(0).Item("CSITEM").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DCDITEM", LTrim(RTrim(dtData.Rows(0).Item("CDITEM").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DCITEM", LTrim(RTrim(dtData.Rows(0).Item("CITEM").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DO1ITEM", LTrim(RTrim(dtData.Rows(0).Item("O1ITEM").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DO2ITEM", LTrim(RTrim(dtData.Rows(0).Item("O2ITEM").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DODESCR1", LTrim(RTrim(dtData.Rows(0).Item("OTHER1").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DODESCR2", LTrim(RTrim(dtData.Rows(0).Item("OTHER2").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DOP1", LTrim(RTrim(dtManufData.Rows(0).Item("OP1").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DOP2", LTrim(RTrim(dtManufData.Rows(0).Item("OP2").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DOP3", LTrim(RTrim(dtManufData.Rows(0).Item("OP3").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DOP4", LTrim(RTrim(dtManufData.Rows(0).Item("OP4").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DOP5", LTrim(RTrim(dtManufData.Rows(0).Item("OP5").ToString))) + _
                          String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DOP6", LTrim(RTrim(dtManufData.Rows(0).Item("OP6").ToString))) + _
                            "window.close();" + _
                          "</script>"
                    Response.Write(Cmd)
                    '
                    'String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.D", dtData.Rows(0).Item("")))) + _
                End If '製造依賴
            End If '開發委員
        End If '開發見本
    End Sub
End Class
