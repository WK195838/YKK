Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Collections.Generic

Partial Class OpenNoPicker
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
        Response.Cookies("PGM").Value = "OpenNoPicker.aspx"                                         '程式名
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
        SQL = "Select No, '...' as NoDesc, "
        SQL &= "Case Sts When 0 Then '開發中' When 1 Then '完成(OK)' When 2 Then '完成(NG)' Else '取消/中止' End As StsDesc, "
        SQL &= "'CommissionSheet_02.aspx?pFormNo=' + FormNo + '&pFormSno=' + str(FormSno, Len(FormSno)) As URL "
        SQL &= "From F_CommissionSheet "
        SQL &= "Where No Like '%" & DKey.Text & "%' "
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
        '
        SQL = "Select * From F_CommissionSheet "
        SQL &= "Where No = '" & xNo & "' "
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            Cmd = "<script>" + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DNO", LTrim(RTrim(dtData.Rows(0).Item("NO")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DREFNO", LTrim(RTrim(dtData.Rows(0).Item("NO")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DAPPBUYER", LTrim(RTrim(dtData.Rows(0).Item("APPBUYER")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSellVendor", LTrim(RTrim(dtData.Rows(0).Item("SELLVENDOR")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DESYQTY", LTrim(RTrim(dtData.Rows(0).Item("ESYQTY")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DCUSTITEM", LTrim(RTrim(dtData.Rows(0).Item("CUSTITEM")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DUSAGE", LTrim(RTrim(dtData.Rows(0).Item("USAGE")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DORNO", LTrim(RTrim(dtData.Rows(0).Item("ORNO")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DNEEDMAP", LTrim(RTrim(dtData.Rows(0).Item("NEEDMAP")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DPRO", LTrim(RTrim(dtData.Rows(0).Item("PRO")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DOPPART", LTrim(RTrim(dtData.Rows(0).Item("OPPART")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DPLEN", LTrim(RTrim(dtData.Rows(0).Item("PLEN")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DPLENUN", LTrim(RTrim(dtData.Rows(0).Item("PLENUN")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DPQTY", LTrim(RTrim(dtData.Rows(0).Item("PQTY")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DPQTYUN", LTrim(RTrim(dtData.Rows(0).Item("PQTYUN")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DUPSLI", LTrim(RTrim(dtData.Rows(0).Item("UPSLI")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DLOSLI", LTrim(RTrim(dtData.Rows(0).Item("LOSLI")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DUPFIN", LTrim(RTrim(dtData.Rows(0).Item("UPFIN")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DLOFIN", LTrim(RTrim(dtData.Rows(0).Item("LOFIN")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DUPSTK", LTrim(RTrim(dtData.Rows(0).Item("UPSTK")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DLOSTK", LTrim(RTrim(dtData.Rows(0).Item("LOSTK")))) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSPSPEC", LTrim(RTrim(dtData.Rows(0).Item("SPSPEC")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DSIZENO", LTrim(RTrim(dtData.Rows(0).Item("SIZENO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DITEM", LTrim(RTrim(dtData.Rows(0).Item("ITEM")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTATYPE", LTrim(RTrim(dtData.Rows(0).Item("TATYPE")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTAWIDTH", LTrim(RTrim(dtData.Rows(0).Item("TAWIDTH")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DECOLSEL", LTrim(RTrim(dtData.Rows(0).Item("ECOLSEL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DECOL", LTrim(RTrim(dtData.Rows(0).Item("ECOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DCCOLSEL", LTrim(RTrim(dtData.Rows(0).Item("CCOLSEL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DCCOL", LTrim(RTrim(dtData.Rows(0).Item("CCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTACOL", LTrim(RTrim(dtData.Rows(0).Item("TACOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTACOLNO", LTrim(RTrim(dtData.Rows(0).Item("TACOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTAYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("TAYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTALCOL", LTrim(RTrim(dtData.Rows(0).Item("TALCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTALCOLNO", LTrim(RTrim(dtData.Rows(0).Item("TALCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTALYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("TALYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTARCOL", LTrim(RTrim(dtData.Rows(0).Item("TARCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTARCOLNO", LTrim(RTrim(dtData.Rows(0).Item("TARCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTARYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("TARYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHLUPCOL", LTrim(RTrim(dtData.Rows(0).Item("THLUPCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHLUPCOLNO", LTrim(RTrim(dtData.Rows(0).Item("THLUPCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHLUPYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("THLUPYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHLRUPCOL", LTrim(RTrim(dtData.Rows(0).Item("THLRUPCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHLRUPCOLNO", LTrim(RTrim(dtData.Rows(0).Item("THLRUPCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHLRUPYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("THLRUPYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHRUPCOL", LTrim(RTrim(dtData.Rows(0).Item("THRUPCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHRUPCOLNO", LTrim(RTrim(dtData.Rows(0).Item("THRUPCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHRUPYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("THRUPYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHRRUPCOL", LTrim(RTrim(dtData.Rows(0).Item("THRRUPCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHRRUPCOLNO", LTrim(RTrim(dtData.Rows(0).Item("THRRUPCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHRRUPYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("THRRUPYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHLLOCOL", LTrim(RTrim(dtData.Rows(0).Item("THLLOCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHLLOCOLNO", LTrim(RTrim(dtData.Rows(0).Item("THLLOCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHLLOYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("THLLOYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHLRLOCOL", LTrim(RTrim(dtData.Rows(0).Item("THLRLOCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHLRLOCOLNO", LTrim(RTrim(dtData.Rows(0).Item("THLRLOCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHLRLOYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("THLRLOYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHRLOCOL", LTrim(RTrim(dtData.Rows(0).Item("THRLOCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHRLOCOLNO", LTrim(RTrim(dtData.Rows(0).Item("THRLOCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHRLOYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("THRLOYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHRRLOCOL", LTrim(RTrim(dtData.Rows(0).Item("THRRLOCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHRRLOCOLNO", LTrim(RTrim(dtData.Rows(0).Item("THRRLOCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DTHRRLOYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("THRRLOYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DXMLEN", LTrim(RTrim(dtData.Rows(0).Item("XMLEN")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DXMCOL", LTrim(RTrim(dtData.Rows(0).Item("XMCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DXMCOLNO", LTrim(RTrim(dtData.Rows(0).Item("XMCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DXMYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("XMYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DAMLEN", LTrim(RTrim(dtData.Rows(0).Item("AMLEN")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DAMCOL", LTrim(RTrim(dtData.Rows(0).Item("AMCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DAMCOLNO", LTrim(RTrim(dtData.Rows(0).Item("AMCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DAMYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("AMYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DBMLEN", LTrim(RTrim(dtData.Rows(0).Item("BMLEN")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DBMCOL", LTrim(RTrim(dtData.Rows(0).Item("BMCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DBMCOLNO", LTrim(RTrim(dtData.Rows(0).Item("BMCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DBMYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("BMYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DCMLEN", LTrim(RTrim(dtData.Rows(0).Item("CMLEN")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DCMCOL", LTrim(RTrim(dtData.Rows(0).Item("CMCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DCMCOLNO", LTrim(RTrim(dtData.Rows(0).Item("CMCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DCMYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("CMYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DDMLEN", LTrim(RTrim(dtData.Rows(0).Item("DMLEN")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DDMCOL", LTrim(RTrim(dtData.Rows(0).Item("DMCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DDMCOLNO", LTrim(RTrim(dtData.Rows(0).Item("DMCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DDMYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("DMYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DEMLEN", LTrim(RTrim(dtData.Rows(0).Item("EMLEN")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DEMCOL", LTrim(RTrim(dtData.Rows(0).Item("EMCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DEMCOLNO", LTrim(RTrim(dtData.Rows(0).Item("EMCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DEMYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("EMYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DFMLEN", LTrim(RTrim(dtData.Rows(0).Item("FMLEN")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DFMCOL", LTrim(RTrim(dtData.Rows(0).Item("FMCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DFMCOLNO", LTrim(RTrim(dtData.Rows(0).Item("FMCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DFMYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("FMYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DGMLEN", LTrim(RTrim(dtData.Rows(0).Item("GMLEN")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DGMCOL", LTrim(RTrim(dtData.Rows(0).Item("GMCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DGMCOLNO", LTrim(RTrim(dtData.Rows(0).Item("GMCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DGMYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("GMYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DHMLEN", LTrim(RTrim(dtData.Rows(0).Item("HMLEN")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DHMCOL", LTrim(RTrim(dtData.Rows(0).Item("HMCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DHMCOLNO", LTrim(RTrim(dtData.Rows(0).Item("HMCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DHMYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("HMYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DLYLEN", LTrim(RTrim(dtData.Rows(0).Item("LYLEN")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DLYCOL", LTrim(RTrim(dtData.Rows(0).Item("LYCOL")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DLYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("LYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DLYYCOLNO", LTrim(RTrim(dtData.Rows(0).Item("LYYCOLNO")))) + _
                    String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DOTCON", LTrim(RTrim(dtData.Rows(0).Item("OTCON")))) + _
                    "window.close();" + _
                  "</script>"
            Response.Write(Cmd)
            'String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.D", LTrim(RTrim(dtData.Rows(0).Item("")))) + _
        End If
    End Sub
End Class
