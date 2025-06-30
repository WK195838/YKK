Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Data.OleDb

Partial Class SampleSheet_02
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wStep As Integer            '工程代號
    Dim wUserID As String           '簽核者
    Dim NowDateTime As String       '現在日期時間

    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900 ' 設定逾時時間
        SetParameter()                  ' 設定共用參數
        If Not IsPostBack Then
            ShowFormData()          ' 顯示表單資料
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
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號
        wStep = Request.QueryString("pStep")        '工程代號
        wUserID = Request.QueryString("pUserID")    '簽核者
        '
        Response.Cookies("PGM").Value = "SampleSheet_02.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("SampleFilePath")
        Dim SQL As String
        SQL = "Select * From F_FactorySampleSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtSampleSheet As DataTable = uDataBase.GetDataTable(SQL)
        If dtSampleSheet.Rows.Count > 0 Then
            '----基本欄位設定-------------------------------------------------
            DNO.Text = dtSampleSheet.Rows(0).Item("No")                              'No
            DAPPBUYER.Text = dtSampleSheet.Rows(0).Item("AppBuyer")                 'Customer
            DDATE.Text = dtSampleSheet.Rows(0).Item("Date")                         '發行日
            DSIZENO.Text = dtSampleSheet.Rows(0).Item("SizeNo")                     'Size
            DITEM.Text = dtSampleSheet.Rows(0).Item("Item")                         'Item
            DCODENO.Text = dtSampleSheet.Rows(0).Item("CodeNo")                     'Code No
            '實際樣品-表
            If dtSampleSheet.Rows(0).Item("SampleFile1") <> "" Then
                LSAMPLEFILE1.ImageUrl = Path & dtSampleSheet.Rows(0).Item("SampleFile1")
                LSAMPLEFILE1.Visible = True
            End If
            '實際樣品-裏
            If dtSampleSheet.Rows(0).Item("SampleFile2") <> "" Then
                LSAMPLEFILE2.ImageUrl = Path & dtSampleSheet.Rows(0).Item("SampleFile2")
                LSAMPLEFILE2.Visible = True
            End If
            '----開發規格-------------------------------------------------
            DTAWIDTH.Text = dtSampleSheet.Rows(0).Item("TAWidth")                   '布帶寬度
            DDEVNO.Text = dtSampleSheet.Rows(0).Item("DevNo")                       '開發No
            DDEVPRD.Text = dtSampleSheet.Rows(0).Item("DevPrd")                     '開發期間
            DTACOL.Text = dtSampleSheet.Rows(0).Item("TACol")                       '布帶Color
            DTALINE.Text = dtSampleSheet.Rows(0).Item("TALine")                     '條紋線Color
            DECOL.Text = dtSampleSheet.Rows(0).Item("ECol")                         '務齒
            DCCOL.Text = dtSampleSheet.Rows(0).Item("CCol")                         '丸紐
            DTHUPCOL.Text = dtSampleSheet.Rows(0).Item("THUPCol")                   '縫工上線
            DTHLOCOL.Text = dtSampleSheet.Rows(0).Item("THLOCol")                   '縫工下線
            '----生產注意-------------------------------------------------
            DMOP1.Text = dtSampleSheet.Rows(0).Item("MOP1")                         '工程-1
            DMOP2.Text = dtSampleSheet.Rows(0).Item("MOP2")                         '工程-2
            DMOP3.Text = dtSampleSheet.Rows(0).Item("MOP3")                         '工程-3
            DMOP4.Text = dtSampleSheet.Rows(0).Item("MOP4")                         '工程-4
            DMNote1.Text = dtSampleSheet.Rows(0).Item("MNote1")                     '說明-1
            DMNote2.Text = dtSampleSheet.Rows(0).Item("MNote2")                     '說明-2
            DMNote3.Text = dtSampleSheet.Rows(0).Item("MNote3")                     '說明-3
            DMNote4.Text = dtSampleSheet.Rows(0).Item("MNote4")                     '說明-4
            '----Wave's-------------------------------------------------
            DTNLITEM.Text = dtSampleSheet.Rows(0).Item("TNLItem")                   'TAPE NAT-左
            DTNRITEM.Text = dtSampleSheet.Rows(0).Item("TNRItem")                   'TAPE NAT-右
            DTSLITEM.Text = dtSampleSheet.Rows(0).Item("TSLItem")                   'TAPE SET-左
            DTSRITEM.Text = dtSampleSheet.Rows(0).Item("TSRItem")                   'TAPE SET-右
            DTDLITEM.Text = dtSampleSheet.Rows(0).Item("TDLItem")                   'TAPE DYED-左
            DTDRITEM.Text = dtSampleSheet.Rows(0).Item("TDRItem")                   'TAPE DYED-右
            DCNITEM.Text = dtSampleSheet.Rows(0).Item("CNItem")                     'CHAIN NAT
            DCSITEM.Text = dtSampleSheet.Rows(0).Item("CSItem")                     'CHAIN SET
            DCDITEM.Text = dtSampleSheet.Rows(0).Item("CDItem")                     'CHAIN DYED
            DODESCR1.Text = dtSampleSheet.Rows(0).Item("Other1")                    'Other1
            DODESCR2.Text = dtSampleSheet.Rows(0).Item("Other2")                    'Other2
            DO1ITEM.Text = dtSampleSheet.Rows(0).Item("O1Item")                     'O1Item
            DO2ITEM.Text = dtSampleSheet.Rows(0).Item("O2Item")                     'O2Item
            DCITEM.Text = dtSampleSheet.Rows(0).Item("CItem")                       'CORD
            '----FLOW-------------------------------------------------
            DOP1.Text = dtSampleSheet.Rows(0).Item("OP1")                           'OP1
            DOP2.Text = dtSampleSheet.Rows(0).Item("OP2")                           'OP2
            DOP3.Text = dtSampleSheet.Rows(0).Item("OP3")                           'OP3
            DOP4.Text = dtSampleSheet.Rows(0).Item("OP4")                           'OP4
            DOP5.Text = dtSampleSheet.Rows(0).Item("OP5")                           'OP5
            DOP6.Text = dtSampleSheet.Rows(0).Item("OP6")                           'OP6
        End If
    End Sub
End Class
