Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Data.OleDb
Partial Class YKKGroupCopyList
    Inherits System.Web.UI.Page
    Dim InputData As String

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wCode As String             'Slider Group Code 
    Dim wFun As String              'Function
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()
        If Not Me.IsPostBack Then   '不是PostBack
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
        wCode = Request.QueryString("pCode")        'Slider Group Code
        wFun = Request.QueryString("pFun")          'Function
        '   wFormNo = "000003"
        '  wFormSno = "61"

    End Sub


    Sub DataList()
        Dim SQL As String
        SQL = "SELECT "
        SQL = SQL + "Case Sts When 0 then '未結' When 1 then '已結OK' When 2 then '已結NG' else '抽單' end as StsDesc, "
        SQL = SQL + "No, Convert(Varchar, Date,111) as DateDesc, SliderCode, MapNo, "
        SQL = SQL + "OFormNo+'-'+str(OFormSno,Len(OFormSno)) as OFormDesc,Formsno,YKKGroup,Buyer "
        SQL = SQL + "FROM F_YKKGroupCopySheet "
        SQL = SQL + "Where OFormNo = '" & wFormNo & "'"
        SQL = SQL + "  and OFormSno = '" & CStr(wFormSno) & "'"
        Dim dt As DataTable = uDataBase.GetDataTable(SQL)
        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()

    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim SFormsNo As String = ""
        SFormsNo = e.Row.Cells(6).Text
        If e.Row.RowType <> DataControlRowType.Header Then
            Dim h1 As New HyperLink
            h1.Text = e.Row.Cells(1).Text
            h1.Target = "_blank"
            h1.NavigateUrl = ("SPD_YKKGroupCopySheet_02.aspx?pFormNo=007001" & "&pFormSno=" & SFormsNo)
            e.Row.Cells(1).Controls.Add(h1)
        End If
        e.Row.Cells(6).Visible = False
    End Sub
End Class
