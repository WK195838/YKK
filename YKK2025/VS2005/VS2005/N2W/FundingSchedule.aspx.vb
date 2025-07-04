Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class FundingSchedule
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
    Dim pFormNo As String
    Dim wLevel As String = ""
    Dim wDivision As String = ""
    Dim wName As String = ""
    Dim NowDateTime As String       '現在日期時間
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim ItemCount As Integer = 6  '預先定義欄位數量 





    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '設定共用參數
        SetMainMenu()         '設定主畫面
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
        NowDateTime = CStr(Now.Date) + " " + _
                          CStr(Now.Hour) + ":" + _
                          CStr(Now.Minute) + ":" + _
                          CStr(Now.Second)     '現在日時
        pFormNo = Request.QueryString("pFormNo")    '表單號碼
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")

    End Sub



    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub




    Sub DataList()
        Dim sql As String
        sql = "  select * from ("
        sql = sql + " select a.step,a.stepname,"
        sql = sql + " case when allcase is null then 0 else allcase end as allcase,"
        sql = sql + " case when allcase is null or allcase =0 then '' else '...' end as allcaseURL,"
        sql = sql + " case when delaycase is null then 0 else delaycase end as delaycase,"
        sql = sql + " case when delaycase is null or delaycase =0 then ''  else '...'  end as delaycaseURL   "
        sql = sql + " from ("
        sql = sql + " select distinct step,substring(stepname,6,len(stepname)-1)stepname from M_flow"
        sql = sql + " where formno = 003110"
        sql = sql + " and step not in (999,1)"
        sql = sql + " )a left join ("
        sql = sql + " select step,stepname,allcase,'...' as allcaseURL,delaycase,'...' as delayacaseURL    from (  "
        sql = sql + " select step,stepname,sum(allcase)allcase,sum(delaycase)delaycase  from ( "
        sql = sql + " select step,substring(stepname,6,len(stepname)-1)stepname,1 as allcase, case when delaysts ='正常' then  0 else 1 end as delaycase  from "
        sql = sql + " V_WaitHandle_01 where formno = '003110' and flowtype ='1' and active ='1'  )a group by step,stepname )a "
        sql = sql + " )b on a.step =b.step"
        sql = sql + " )a order by step"
        Dim dtData As DataTable = uDataBase.GetDataTable(sql)
        GridView1.DataSource = dtData
        GridView1.DataBind()
    End Sub





    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(0).Text = "合計"
        End If


        If e.Row.RowType = DataControlRowType.DataRow Then


            Dim h1 As New HyperLink
            Dim h2 As New HyperLink

            Dim wStep As String
            wStep = e.Row.Cells(5).Text


            If e.Row.Cells(1).Text <> "0" Then
                h1.Target = "_blank"
                h1.Text = e.Row.Cells(2).Text
                ' 連結到待處理LIST
                ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep
                'h1.NavigateUrl = "FundingScheduleAll.aspx?&pStep=" + wStep
                h1.NavigateUrl = "FundingScheduleAll.aspx?&pStep=" + wStep + "&pUserID=" + Request.QueryString("pUserID")

                e.Row.Cells(2).Text = ""
                e.Row.Cells(2).Controls.Add(h1)

            End If


            If e.Row.Cells(3).Text <> "0" Then
                h2.Target = "_blank"
                h2.Text = e.Row.Cells(4).Text
                ' 連結到延遲
                'h2.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep + "&pDelay=1"
                'h2.NavigateUrl = "FundingScheduleAll.aspx?&pStep=" + wStep + "&pDelay=1"
                h2.NavigateUrl = "FundingScheduleAll.aspx?&pStep=" + wStep + "&pDelay=1" + "&pUserID=" + Request.QueryString("pUserID")

                e.Row.Cells(4).Text = ""
                e.Row.Cells(4).Controls.Add(h2)
            End If
           
        End If
        '連結

        e.Row.Cells(5).Visible = False

    End Sub

    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBound
        Dim formatNu As Integer

        formatNu = 0

        Dim intTotal As Decimal = 0
        Dim intTotal1 As Decimal = 0
        For Each gvr As GridViewRow In GridView1.Rows


            intTotal += gvr.Cells(1).Text

            intTotal1 += gvr.Cells(3).Text

            gvr.Cells(1).Text = FormatNumber(gvr.Cells(1).Text, formatNu, TriState.True, TriState.False, TriState.True)
            'gvr.Cells(1).HorizontalAlign = HorizontalAlign.Center
            'gvr.Cells(2).HorizontalAlign = HorizontalAlign.Center
            GridView1.FooterRow.Cells(1).Text = FormatNumber(intTotal, formatNu, TriState.True, TriState.False, TriState.True)
            GridView1.FooterRow.Cells(1).HorizontalAlign = HorizontalAlign.Center


            gvr.Cells(3).Text = FormatNumber(gvr.Cells(3).Text, formatNu, TriState.True, TriState.False, TriState.True)
            'gvr.Cells(3).HorizontalAlign = HorizontalAlign.Center
            'gvr.Cells(4).HorizontalAlign = HorizontalAlign.Center
            GridView1.FooterRow.Cells(3).Text = FormatNumber(intTotal1, formatNu, TriState.True, TriState.False, TriState.True)
            GridView1.FooterRow.Cells(3).HorizontalAlign = HorizontalAlign.Center
        Next



    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定菜單各程式 
    '**
    '*****************************************************************
    Private Sub SetMainMenu()
        Dim xStep As Integer = 1
        Dim SQL As String = ""


        LFun01.Enabled = True
        LFun01.NavigateUrl = "FundingSheet_01.aspx?pFormNo=003110" & _
                                                          "&pFormSno=0" & _
                                                          "&pStep=" & xStep & _
                                                          "&pSeqNo=0" & _
                                                          "&pUserID=" & Request.QueryString("pUserID") & _
                                                          "&pApplyID=" & Request.QueryString("pUserID")
        LFun01.Target = "_blank"

        LFun02.Enabled = True
        LFun02.NavigateUrl = "FundinginqCommission.aspx?pFormNo=003110" & _
                                                          "&pUserID=" & Request.QueryString("pUserID")


        '  LFun02.Target = "_blank"

    End Sub


    Protected Sub LFun01_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun01.Init
        LFun01.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun01.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LFun02_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun02.Init
        LFun02.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun02.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub
End Class

