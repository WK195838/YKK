Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class EApprovalRDSchedule
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

        'sql = "select * from ( "
        'sql = sql + "select a.step,a.stepname,isnull(allcase,0)allcase, case when allcase is null or allcase =0 then '' else '...' end as allcaseURL, "
        'sql = sql + "isnull(delaycase,0)delaycase, case when delaycase is null or delaycase =0 then ''  else '...'  end as delaycaseURL  from ("
        'sql = sql + "select  step,substring(stepname,6,len(stepname)-1)stepname,Hours/8/60  AS hours  from M_leadtime where formno ='003105'  )a left join "
        'sql = sql + "( select step,sum(allcase)allcase,sum(delaycase)delaycase from( select formno,step,substring(stepname,11,len(stepname)-1)stepname,1 as allcase, "
        'sql = sql + "case when delaysts ='正常' then  0 else 1 end as delaycase,formsno  from   V_WaitHandle_01  where formno = '003105' and flowtype ='1' and active ='1' "
        'sql = sql + "and seqno=1  )a,( select formno,formsno  from F_ExpenseSheet where sts =0)b "
        'sql = sql + "where a.formsno =b.formsno and b.formno = '003105' group by step )b on a.step =b.step)a where step in (10,20,30,500) order by step "

        sql = "select * from ( "
        sql = sql + "select a.step,a.stepname,isnull(allcase,0)allcase, case when allcase is null or allcase =0 then '' else '...' end as allcaseURL, "
        sql = sql + "isnull(delaycase,0)delaycase, case when delaycase is null or delaycase =0 then ''  else '...'  end as delaycaseURL  from ( "
        sql = sql + "select  step,substring(stepname,7,len(stepname)-1)stepname from M_leadtime where formno ='007003'  )a left join "
        sql = sql + "( select step,sum(allcase)allcase,sum(delaycase)delaycase from( select formno,step,substring(stepname,11,len(stepname)-1)stepname,1 as allcase, "
        sql = sql + "case when delaysts ='正常' then  0 else 1 end as delaycase,formsno  from   V_WaitHandle_01  where formno = '007003' and flowtype ='1' and active ='1'  "
        sql = sql + "and seqno=1  "
        sql = sql + ")a,( select formno,formsno  from F_EApprovalRDSheet "
        sql = sql + "where sts =0  )b  where a.formsno =b.formsno and b.formno = '007003' group by step )b on a.step =b.step )a where step in (10,20,31,32,33,34,41,42,43,44,50,60,70,80,500 ) order by step"

        'sql = "SELECT A.step,A.stepname,CASE WHEN A.CASECOUNT is null THEN 0 ELSE A.CASECOUNT END AS CASECOUNT, CASE WHEN A.CASECOUNT=0 THEN '' ELSE '...' END AS CASEURL , "
        'sql = sql + "CASE WHEN B.DECASECOUNT is null THEN 0 ELSE B.DECASECOUNT END AS DECASECOUNT, CASE WHEN B.DECASECOUNT is null or B.DECASECOUNT=0 THEN '' ELSE '...' END AS DECASEURL FROM "
        'sql = sql + "(SELECT M_Flow.step,substring(M_Flow.stepname,6,len(M_Flow.stepname)-1)stepname,count(*) CASECOUNT FROM  M_Flow "
        'sql = sql + "LEFT JOIN  V_WaitHandle_01 "
        'sql = sql + "ON M_Flow.step = V_WaitHandle_01.step and M_Flow.FormNo = V_WaitHandle_01.FormNo "
        'sql = sql + "WHERE  V_WaitHandle_01.FormNo = '003105' and M_Flow.Action=0 and V_WaitHandle_01.active =1 and M_Flow.FlowType =1 "
        'sql = sql + "GROUP BY  M_Flow.step,M_Flow.stepname) A "
        'sql = sql + "Left Join "
        'sql = sql + "(SELECT M_Flow.step,substring(M_Flow.stepname,6,len(M_Flow.stepname)-1)stepname,count(*) DECASECOUNT FROM  M_Flow "
        'sql = sql + "LEFT JOIN  V_WaitHandle_01 "
        'sql = sql + "ON M_Flow.step = V_WaitHandle_01.step and M_Flow.FormNo = V_WaitHandle_01.FormNo "
        'sql = sql + "WHERE  V_WaitHandle_01.FormNo = '003105' and M_Flow.Action=0 and V_WaitHandle_01.active =1 and M_Flow.FlowType =1 and delaysts ='延遲' "
        'sql = sql + "GROUP BY  M_Flow.step,M_Flow.stepname) B "
        'sql = sql + "ON A.Step = B.Step "
        'sql = sql + "ORDER BY A.Step "


        Dim dtData As DataTable = uDataBase.GetDataTable(sql)
        GridView1.DataSource = dtData
        GridView1.DataBind()

    End Sub




    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(0).Text = "合計"
        End If


        If (e.Row.RowType = DataControlRowType.DataRow) <> (e.Row.RowType = DataControlRowType.Footer) Then


            Dim h1 As New HyperLink
            Dim h2 As New HyperLink

            Dim wStep As String
            wStep = e.Row.Cells(5).Text

            If e.Row.Cells(0).Text <> "合計" Then
                If Trim(e.Row.Cells(1).Text) <> "0" Then
                    h1.Target = "_blank"
                    h1.Text = e.Row.Cells(2).Text
                    ' 連結到待處理LIST
                    ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep
                    'h1.NavigateUrl = "FundingScheduleAll.aspx?&pStep=" + wStep
                    h1.NavigateUrl = "EApprovalRDScheduleAll.aspx?&pStep=" + wStep + "&pUserID=" + Request.QueryString("pUserID")

                    ' e.Row.Cells(3).Text = ""
                    e.Row.Cells(2).Controls.Add(h1)

                End If

                If Trim(e.Row.Cells(3).Text) <> "0" Then
                    h2.Target = "_blank"
                    h2.Text = e.Row.Cells(4).Text
                    ' 連結到待處理LIST
                    ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep
                    'h1.NavigateUrl = "FundingScheduleAll.aspx?&pStep=" + wStep
                    h2.NavigateUrl = "EApprovalRDScheduleAll.aspx?&pStep=" + wStep + "&pUserID=" + Request.QueryString("pUserID")

                    ' e.Row.Cells(4).Text = ""
                    e.Row.Cells(4).Controls.Add(h2)

                End If


            End If

            e.Row.Cells(5).Visible = False


            '連結
        End If


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
        LFun02.Enabled = True


        LFun01.NavigateUrl = "EApprovalRDSheet_01.aspx?pFormNo=007003" & _
                                                                            "&pFormSno=0" & _
                                                                            "&pStep=" & xStep & _
                                                                            "&pSeqNo=0" & _
                                                                            "&pUserID=" & Request.QueryString("pUserID") & _
                                                                            "&pApplyID=" & Request.QueryString("pUserID")
        LFun02.NavigateUrl = "EApprovalRDinqCommission.aspx?pFormNo=007003" & _
                                                         "&pUserID=" & Request.QueryString("pUserID")


        LFun01.Target = "_blank"



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


    Protected Sub LFun03_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun01.Init
        LFun01.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun01.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub


    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then
            ' 建立自訂的標題 

            Dim gv As GridView = DirectCast(sender, GridView)
            Dim gvRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

            Dim tc1 As New TableCell()
            tc1.Text = "工程"
            tc1.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc1)

            Dim tc2 As New TableCell()
            tc2.Text = "待處理件數"
            tc2.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc2)

            Dim tc3 As New TableCell()
            tc3.Text = "待處理連結"
            tc3.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc3)


            Dim tc4 As New TableCell()
            tc4.Text = "延遲件數"
            tc4.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc4)

            Dim tc5 As New TableCell()
            tc5.Text = "延遲連結"
            tc5.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc5)



            ' Dim tc6 As New TableCell()
            'tc6.Text = "STEP"
            'tc6.HorizontalAlign = HorizontalAlign.Center
            'gvRow.Cells.Add(tc6)

            e.Row.Cells.Clear()
            gv.Controls(0).Controls.AddAt(0, gvRow)

        End If

    End Sub
End Class

