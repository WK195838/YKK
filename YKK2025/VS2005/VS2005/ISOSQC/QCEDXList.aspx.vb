Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class QCEDXList
    Inherits System.Web.UI.Page
    '
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oWaves As New WAVES.CommonService
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim UserID As String            'UserID
    Dim pBuyer As String            'Buyer
    Dim pPuller As String           'Puller
    Dim pColor As String            'Color
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                          '設定參數
        If Not IsPostBack Then                  'PostBack
            SetDefaultValue()                   '設定初值
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
        Server.ScriptTimeout = 900                          '設定逾時時間
        Response.Cookies("PGM").Value = "QCEDXList.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        UserID = Request.QueryString("pUserID")             'UserID
        pBuyer = Request.QueryString("pBuyer")
        pPuller = Request.QueryString("pPuller")
        pColor = Request.QueryString("pColor")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        '-----------------------------------------------------------------
        '-- Not IsPostBack
        '-----------------------------------------------------------------
        '
        DKPullerCode.Text = ""
        DKColor.Text = ""
        DKOther.Text = ""
        '
        If pPuller <> "" Then DKPullerCode.Text = pPuller
        If pColor <> "" Then DKColor.Text = pColor
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Go)
    '**     Find Puller
    '**
    '*****************************************************************
    Protected Sub BFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFind.Click
        ShowData()
    End Sub
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim SQL As String
        '
        SQL = "SELECT top 150 "
        SQL = SQL & "[Cat],[Size],[Family],[Body],[Puller],[Color],[Finish],[Supplier],  convert(varchar, CreateTime, 111) as CreateTime "
        SQL = SQL & "FROM M_EDX "
        '
        SQL = SQL & "where Cat in ('QC','HAND') "
        ' Puller
        If DKPullerCode.Text <> "" Then
            SQL = SQL & "and ( "
            SQL = SQL & "    puller = '" & DKPullerCode.Text & "' "
            SQL = SQL & " or puller = '" & DKPullerCode.Text & "-B' "
            SQL = SQL & " or puller = '" & DKPullerCode.Text & "SK' "
            SQL = SQL & ") "
        End If
        ' Color
        If DKColor.Text <> "" Then
            SQL = SQL & "and Color = '" & DKColor.Text & "' "
        End If
        ' Other
        If DKOther.Text <> "" Then
            SQL = SQL & "and [Cat]+[Size]+[Family]+[Body]+[Puller]+[Color]+[Finish]+[Supplier] like '%" & DKOther.Text & "%' "
        End If
        '
        SQL = SQL & "Order By [Cat] DESC,[Size],[Family],[Body] "
        '
        GridView1.Visible = True
        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()
        '
    End Sub
End Class
