Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class FundingScheduleTC
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


        sql = "  SELECT * FROM ("

        sql = sql + " SELECT A.STEP, CASE WHEN a.STEP=1 THEN '公司別'  else a.stepname end as stepname,"
        sql = sql + " CASE WHEN a.STEP=1 THEN '台北'  else  A.ALLCASE end as TALLCASE,"
        sql = sql + " CASE WHEN a.STEP=1 THEN '中壢'  else  B.ALLCASE end as CALLCASE,"
        sql = sql + " CASE WHEN a.STEP=1 THEN '台北'  else  A.ALLCASEURL end as TALLCASEURL,"
        sql = sql + " CASE WHEN a.STEP=1 THEN '中壢'  else  B.ALLCASEURL end as CALLCASEURL,"
        sql = sql + " CASE WHEN a.STEP=1 THEN '台北'  else  A.DELAYCASE end as  TDELAYCASE,"
        sql = sql + " CASE WHEN a.STEP=1 THEN '中壢'  else  B.DELAYCASE end as CDELAYCASE,"
        sql = sql + " CASE WHEN a.STEP=1 THEN '台北'  else  A.DELAYCASEURL end as  TDELAYCASEURL,"
        sql = sql + " CASE WHEN a.STEP=1 THEN '中壢'  else  B.DELAYCASEURL end as CDELAYCASEURL"
        sql = sql + " FROM ("
        sql = sql + " select a.step,a.stepname,"
        sql = sql + " case when allcase is null then '0' else str(allcase) end as allcase,"
        sql = sql + " case when allcase is null or allcase =0 then '' else 'TP' end as allcaseURL,"
        sql = sql + " case when delaycase is null then '0' else str(delaycase) end as delaycase,"
        sql = sql + " case when delaycase is null or delaycase =0 then ''  else 'TP'  end as delaycaseURL   "
        sql = sql + " from ("
        sql = sql + " select distinct step,substring(stepname,6,len(stepname)-1)stepname from M_flow"
        sql = sql + " where formno = 003110"
        sql = sql + " and step not in (999)"
        sql = sql + " )a left join ("
        sql = sql + " select step,stepname,allcase,'TP' as allcaseURL,delaycase,'TP' as delayacaseURL    from (  "
        sql = sql + " select step,stepname,sum(allcase)allcase,sum(delaycase)delaycase  from ( "
        sql = sql + " select step,substring(stepname,6,len(stepname)-1)stepname,1 as allcase, case when delaysts ='正常' then  '0' else 1 end as delaycase  from "
        sql = sql + " V_WaitHandle_01 where formno = '003110' and flowtype ='1' and active ='1'  )a group by step,stepname )a "
        sql = sql + " )b on a.step =b.step"
        sql = sql + " )A,("
        sql = sql + " select a.step,a.stepname,"
        sql = sql + " case when allcase is null then '0' else str(allcase) end as allcase,"
        sql = sql + " case when allcase is null or allcase =0 then '' else 'CL' end as allcaseURL,"
        sql = sql + " case when delaycase is null then '0' else str(delaycase) end as delaycase,"
        sql = sql + " case when delaycase is null or delaycase =0 then ''  else 'CL'  end as delaycaseURL   "
        sql = sql + " from ("
        sql = sql + " select distinct step,substring(stepname,6,len(stepname)-1)stepname from M_flow"
        sql = sql + " where formno = 003111"
        sql = sql + " and step not in (999)"
        sql = sql + " )a left join ("
        sql = sql + " select step,stepname,allcase,'CL' as allcaseURL,delaycase,'CL' as delayacaseURL    from (  "
        sql = sql + " select step,stepname,sum(allcase)allcase,sum(delaycase)delaycase  from ( "
        sql = sql + " select step,substring(stepname,6,len(stepname)-1)stepname,1 as allcase, case when delaysts ='正常' then  0 else 1 end as delaycase  from "
        sql = sql + " V_WaitHandle_01 where formno = '003111' and flowtype ='1' and active ='1'  )a group by step,stepname )a "
        sql = sql + " )b on a.step =b.step"
        sql = sql + " )B WHERE A.STEP =B.STEP "
        sql = sql + " )A ORDER BY STEP"



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

            Dim h3 As New HyperLink
            Dim h4 As New HyperLink

            Dim wStep As String
            wStep = e.Row.Cells(9).Text

            If e.Row.Cells(0).Text <> "公司別" And e.Row.Cells(0).Text <> "合計" Then
                If Trim(e.Row.Cells(1).Text) <> "0" Then
                    h1.Target = "_blank"
                    h1.Text = e.Row.Cells(3).Text
                    ' 連結到待處理LIST
                    ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep
                    'h1.NavigateUrl = "FundingScheduleAll.aspx?&pStep=" + wStep
                
                    h1.NavigateUrl = "FundingScheduleAllTC.aspx?&pStep=" + wStep + "&pUserID=" + Request.QueryString("pUserID") + "&pDepo=TP"


                    ' e.Row.Cells(3).Text = ""
                    e.Row.Cells(3).Controls.Add(h1)

                End If

                If Trim(e.Row.Cells(2).Text) <> "0" Then
                    h2.Target = "_blank"
                    h2.Text = e.Row.Cells(4).Text
                    ' 連結到待處理LIST
                    ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep
                    'h1.NavigateUrl = "FundingScheduleAll.aspx?&pStep=" + wStep
                    h2.NavigateUrl = "FundingScheduleAllTC.aspx?&pStep=" + wStep + "&pUserID=" + Request.QueryString("pUserID") + "&pDepo=CL"

                    ' e.Row.Cells(4).Text = ""
                    e.Row.Cells(4).Controls.Add(h2)

                End If



                If Trim(e.Row.Cells(5).Text) <> "0" Then
                    h3.Target = "_blank"
                    h3.Text = e.Row.Cells(7).Text
                    ' 連結到延遲
                    'h2.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep + "&pDelay=1"
                    'h2.NavigateUrl = "FundingScheduleAll.aspx?&pStep=" + wStep + "&pDelay=1"
                    h3.NavigateUrl = "FundingScheduleAllTC.aspx?&pStep=" + wStep + "&pDelay=1" + "&pUserID=" + Request.QueryString("pUserID") + "&pDepo=TP"

                    ' e.Row.Cells(7).Text = ""
                    e.Row.Cells(7).Controls.Add(h3)
                End If

                If Trim(e.Row.Cells(6).Text) <> "0" Then
                    h4.Target = "_blank"
                    h4.Text = e.Row.Cells(8).Text
                    ' 連結到延遲
                    'h2.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep + "&pDelay=1"
                    'h2.NavigateUrl = "FundingScheduleAll.aspx?&pStep=" + wStep + "&pDelay=1"
                    h4.NavigateUrl = "FundingScheduleAllTC.aspx?&pStep=" + wStep + "&pDelay=1" + "&pUserID=" + Request.QueryString("pUserID") + "&pDepo=CL"

                    ' e.Row.Cells(8).Text = ""
                    e.Row.Cells(8).Controls.Add(h4)
                End If

            End If

            e.Row.Cells(9).Visible = False



        End If
        '連結

       
      
    End Sub

    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBound
        Dim formatNu As Integer

        formatNu = 0

        Dim intTotalT As Decimal = 0
        Dim intTotal1T As Decimal = 0
        Dim intTotalC As Decimal = 0
        Dim intTotal1C As Decimal = 0

        For Each gvr As GridViewRow In GridView1.Rows

            If gvr.Cells(0).Text <> "公司別" Then
                intTotalT += gvr.Cells(1).Text
                intTotalC += gvr.Cells(2).Text

                intTotal1T += gvr.Cells(5).Text
                intTotal1C += gvr.Cells(6).Text

                gvr.Cells(1).Text = FormatNumber(gvr.Cells(1).Text, formatNu, TriState.True, TriState.False, TriState.True)
                'gvr.Cells(1).HorizontalAlign = HorizontalAlign.Center
                'gvr.Cells(2).HorizontalAlign = HorizontalAlign.Center
                GridView1.FooterRow.Cells(1).Text = FormatNumber(intTotalT, formatNu, TriState.True, TriState.False, TriState.True)
                GridView1.FooterRow.Cells(1).HorizontalAlign = HorizontalAlign.Center

                gvr.Cells(2).Text = FormatNumber(gvr.Cells(2).Text, formatNu, TriState.True, TriState.False, TriState.True)
                'gvr.Cells(1).HorizontalAlign = HorizontalAlign.Center
                'gvr.Cells(2).HorizontalAlign = HorizontalAlign.Center
                GridView1.FooterRow.Cells(2).Text = FormatNumber(intTotalC, formatNu, TriState.True, TriState.False, TriState.True)
                GridView1.FooterRow.Cells(2).HorizontalAlign = HorizontalAlign.Center

                'gvr.Cells(3).HorizontalAlign = HorizontalAlign.Center
                'gvr.Cells(4).HorizontalAlign = HorizontalAlign.Center

                gvr.Cells(5).Text = FormatNumber(gvr.Cells(5).Text, formatNu, TriState.True, TriState.False, TriState.True)
                'gvr.Cells(3).HorizontalAlign = HorizontalAlign.Center
                'gvr.Cells(4).HorizontalAlign = HorizontalAlign.Center
                GridView1.FooterRow.Cells(5).Text = FormatNumber(intTotal1T, formatNu, TriState.True, TriState.False, TriState.True)
                GridView1.FooterRow.Cells(5).HorizontalAlign = HorizontalAlign.Center

                gvr.Cells(6).Text = FormatNumber(gvr.Cells(6).Text, formatNu, TriState.True, TriState.False, TriState.True)
                'gvr.Cells(3).HorizontalAlign = HorizontalAlign.Center
                'gvr.Cells(4).HorizontalAlign = HorizontalAlign.Center
                GridView1.FooterRow.Cells(6).Text = FormatNumber(intTotal1C, formatNu, TriState.True, TriState.False, TriState.True)
                GridView1.FooterRow.Cells(6).HorizontalAlign = HorizontalAlign.Center
            End If
           
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

        SQL = "  select  deponame,left(divid,2) as  DepoID  from M_users"
        SQL = SQL + " where Userid ='" + Request.QueryString("pUserID") + "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)



        If Now.ToString("yyyy/MM/ddH") >= "2025/04/02" And Now.ToString("yyyy/MM/ddH") <= "2025/04/06" Then
            LFun02.NavigateUrl = "Notice.aspx?"

            If dtData.Rows(0).Item("DepoName") = "中壢" Or dtData.Rows(0).Item("DepoID") = "10" Then
                LFun01.NavigateUrl = "Notice.aspx?"
                LFun02.NavigateUrl = "FundinginqCommissionTC.aspx?pFormNo=003111" & _
                                                                 "&pUserID=" & Request.QueryString("pUserID")
            Else
                LFun01.NavigateUrl = "Notice.aspx?"
                LFun02.NavigateUrl = "FundinginqCommissionTC.aspx?pFormNo=003110" & _
                                                                  "&pUserID=" & Request.QueryString("pUserID")


            End If


        Else


            If dtData.Rows(0).Item("DepoName") = "中壢" Or dtData.Rows(0).Item("DepoID") = "10" Then
                LFun01.NavigateUrl = "FundingSheetCL_01.aspx?pFormNo=003111" & _
                                                                                    "&pFormSno=0" & _
                                                                                    "&pStep=" & xStep & _
                                                                                    "&pSeqNo=0" & _
                                                                                    "&pUserID=" & Request.QueryString("pUserID") & _
                                                                                    "&pApplyID=" & Request.QueryString("pUserID")
                LFun02.NavigateUrl = "FundinginqCommissionTC.aspx?pFormNo=003111" & _
                                                                 "&pUserID=" & Request.QueryString("pUserID")
            Else
                LFun01.NavigateUrl = "FundingSheet_01.aspx?pFormNo=003110" & _
                                                                        "&pFormSno=0" & _
                                                                        "&pStep=" & xStep & _
                                                                        "&pSeqNo=0" & _
                                                                        "&pUserID=" & Request.QueryString("pUserID") & _
                                                                        "&pApplyID=" & Request.QueryString("pUserID")



                LFun02.NavigateUrl = "FundinginqCommissionTC.aspx?pFormNo=003110" & _
                                                                  "&pUserID=" & Request.QueryString("pUserID")


            End If
        End If





        LFun01.Target = "_blank"

        

        LFun02.Target = "_blank"

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
            tc2.ColumnSpan = 2   ' 跨二欄
            tc2.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc2)

            Dim tc3 As New TableCell()
            tc3.Text = "待處理連結"
            tc3.ColumnSpan = 2   ' 跨二欄
            tc3.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc3)


            Dim tc4 As New TableCell()
            tc4.Text = "延遲件數"
            tc4.ColumnSpan = 2   ' 跨二欄
            tc4.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc4)

            Dim tc5 As New TableCell()
            tc5.Text = "延遲連結"
            tc5.ColumnSpan = 2   ' 跨二欄
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

