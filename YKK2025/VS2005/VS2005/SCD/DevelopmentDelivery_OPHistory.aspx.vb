Imports System.Data

Partial Class DevelopmentDelivery_OPHistory
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wStep As String             '工程代碼
    Dim wCode As String             'OP-Code
    '
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.Cookies("PGM").Value = "DevelopmentDelivery_History.aspx"
        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
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

        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號
        wStep = Request.QueryString("pStep")        '工程代碼
        wCode = Request.QueryString("pOPCode")      'OP-Code

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選資料
    '**
    '*****************************************************************
    Sub DataList()
        Dim SQL As String
        SQL = "select OP,OPPER,OPBTIME,OPBHOURS,OPATIME,OPAHOURS,OPDELAYC1,OPDELAYC2,OPREM "
        SQL &= "from FS_ManufactureHistory "
        SQL &= "where formno = '" + wFormNo + "' "
        SQL &= "  and formsno ='" + CStr(wFormSno) + "' "
        SQL &= "  and step = '" + CStr(wStep) + "' "
        SQL &= "  and OPCode = '" + wCode + "' "
        SQL &= "Order by Unique_ID "
        Dim dt_ManufactureHistory As DataTable = uDataBase.GetDataTable(SQL)
        '
        GridView1.DataSource = dt_ManufactureHistory
        GridView1.DataBind()
    End Sub

End Class
