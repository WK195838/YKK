Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class ComplaintInScheduleAll
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
    Dim wStep, wDelay, wDep, wType As String
    Dim NowDateTime As String       '現在日期時間
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript





    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '設定共用參數

        If Not Me.IsPostBack Then
            SetFieldData()
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
        wStep = Request.QueryString("pStep")
        wDelay = Request.QueryString("pDelay")
        wDep = Request.QueryString("pDep")
        wType = Request.QueryString("pType")
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
        Dim str999 As String = ""
        Dim StepStr As String
        If wStep = "999" Then
            StepStr = " and b.sts =1 and step = '" + wStep + "'"
            str999 = " union all "
            str999 = str999 + " select a.*,qcno,ACCDEPNAME,spec,specname,STATUS,orqty,ERRORQTY,CUSTOMER,COMNAME from ("
            str999 = str999 + " select * from V_WaitHandle_01 where formno = '003108' "
            str999 = str999 + " )a,F_ComplaintInSheet b where   a.formno =b.formno"
            str999 = str999 + " and a.formsno = b.formsno   and b.sts =1 and step = '999'"
        Else
            StepStr = " and step = '" + wStep + "'"
        End If

        Dim DelayStr As String
        If wDelay <> "" Then
            'DelayStr = " where delaycase = '" + wDelay + "'"
            DelayStr = " where  delaysts = '" + "延遲" + "'"
        Else
            DelayStr = ""
        End If
        Dim Accdep1 As String
        If DACCDEP1.SelectedValue = "全部" Then
            Accdep1 = ""
        Else
            Accdep1 = "and ACCDEPNAME='" + DACCDEP1.SelectedValue + "'"
        End If

        Dim TypeStr As String = ""

        If wType <> "" Then
            If wType = 0 Then
                TypeStr = " and  COMINTYPE ='廠內客訴'"
            ElseIf wType = 1 Then
                TypeStr = " and  COMINTYPE ='最終檢驗'"
            Else
                TypeStr = " and  COMINTYPE ='特殊全檢'"
            End If

        Else
            TypeStr = ""
        End If


        Dim sql As String

        'If wStep = 1 Or wStep = 500 Then
        'sql = "  SELECT no 委託NO,spec 型別,STATUS 客訴內容,orqty ORQTY,ERRORQTY 客訴數量,CUSTOMER 顧客名稱,buyer BUYER名,COMNAME 投訴擔當者,"
        'Else
        sql = "  SELECT no 委託NO, qcno 品管受理編號,ACCDEPNAME 被訴部門,spec ITEMCODE,specname 型別 ,STATUS 客訴內容,orqty ORQTY,ERRORQTY 全檢後不良數量,CUSTOMER 顧客名稱,buyer BUYER名,COMNAME  投訴擔當者, "
        'End If

        sql = sql + " '履歷' as 履歷"
        sql = sql + " ,Delaysts as '正常/延遲',"
        ' sql = sql + "'申請時間：[' + Convert(VarChar, ApplyTime, 20) + '], '+"
        ' sql = sql + "'收件時間：[' + Convert(VarChar, ReceiptTime, 20) + '], ' +"
        ' sql = sql + "'閱讀期限：[' + Convert(VarChar, ReadTimeLimit, 20) + '], ' +"
        ' sql = sql + "'首次閱讀：[' + FirstReadTimeDesc + '], ' +"
        ' sql = sql + "'最後閱讀：[' + LastReadTimeDesc  + '], ' +"
        ' sql = sql + "'預定開始：[' + Convert(VarChar, BStartTime, 20) + '], ' +"
        ' sql = sql + "'預定完成：[' + Convert(VarChar, BEndTime, 20) + '], ' +"
        ' sql = sql + "'實際開始：[' + Convert(VarChar, AStartTime, 20) + '] ' +"
        ' sql = sql + "'預定完成：[' + Convert(VarChar, BEndTime, 20) + '] '  "
        ' sql = sql + " As 參考資料,
        sql = sql + " stsdesc 狀態,formsno "

        'sql = sql + " FormNo, FormSno, Step, SeqNo,Division,No,substring(stepnamedesc,9,"
        'sql = sql + " StsDesc, FormName, FlowTypeDesc, ApplyName, AgentName,"
        'sql = sql + "'申請時間：[' + Convert(VarChar, ApplyTime, 20) + '], '+"
        'sql = sql + "'收件時間：[' + Convert(VarChar, ReceiptTime, 20) + '], ' +"
        'sql = sql + "'閱讀期限：[' + Convert(VarChar, ReadTimeLimit, 20) + '], ' +"
        'sql = sql + "'首次閱讀：[' + FirstReadTimeDesc + '], ' +"
        'sql = sql + "'最後閱讀：[' + LastReadTimeDesc  + '], ' +"
        'sql = sql + "'預定開始：[' + Convert(VarChar, BStartTime, 20) + '], ' +"
        'sql = sql + "'預定完成：[' + Convert(VarChar, BEndTime, 20) + '], ' +"
        'sql = sql + "'實際開始：[' + Convert(VarChar, AStartTime, 20) + '] ' +"
        'sql = sql + "'預定完成：[' + Convert(VarChar, BEndTime, 20) + '] '  "
        'sql = sql + " As Description, URL ,Delaysts,Accdep1 "


        sql = sql + " FROM ( "
        sql = sql + " select a.no,Delaysts,stsdesc,a.formsno,qcno,ACCDEPNAME,spec,specname,STATUS ,orqty,ERRORQTY,CUSTOMER,COMNAME,b.buyer  "
        sql = sql + " from ("
        sql = sql + " select *from V_WaitHandle_01"
        sql = sql + " where formno = '003108'"
        sql = sql + " )a,F_ComplaintInSheet b where   a.formno =b.formno "
        sql = sql + " and a.formsno = b.formsno"
        sql = sql + " and Active = '1'"
        sql = sql + " and flowtype =1 and seqno =1 " + StepStr + Accdep1 + TypeStr
        sql = sql + str999 + Accdep1
        sql = sql + " )a " + DelayStr
        sql = sql + " order by no "
        Dim dtData As DataTable = uDataBase.GetDataTable(sql)
        GridView1.DataSource = dtData
        GridView1.DataBind()
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

        '       Dim tempFolderPath As String
        '
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim h1, h2 As New HyperLink

            ' '檢查是否有附檔
            ' tempFolderPath = "\\10.245.1.6\wfs$\N2W\003108\" + e.Row.Cells(0).Text + "\品管最終確認"

            'Dim dirInfo As New System.IO.DirectoryInfo(tempFolderPath)
            'Dim FileDir As Integer  '資料夾



            'If System.IO.Directory.Exists(tempFolderPath) Then
            '    FileDir = dirInfo.GetDirectories("*").Length
            '    Dim FileCount As Integer '檔案
            '    FileCount = dirInfo.GetFiles("*.*").Length


            '    If FileCount > 0 Or FileDir > 0 Then
            '        If wStep = 10 Or wStep = 500 Then
            '            e.Row.Cells(13).Text = "V"
            '        Else
            '            e.Row.Cells(17).Text = "V"
            '        End If

            '        '   DChkData2.Text = Str(FileCount) + "件"
            '    Else
            '        If wStep = 10 Or wStep = 500 Then
            '            e.Row.Cells(13).Text = ""
            '        Else
            '            e.Row.Cells(17).Text = ""
            '        End If
            '    End If
            'End If




            Dim wFormsno As String
            
            wFormsno = e.Row.Cells(14).Text
            h1.Text = e.Row.Cells(0).Text
            ' 連結到待處理LIST
            ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep
            h1.NavigateUrl = "ComplaintInSheet_02.aspx?&pFormno=003108&pFormsno=" + wFormsno
            h1.Target = "_blank"
            e.Row.Cells(0).Text = ""
            e.Row.Cells(0).Controls.Add(h1)


            h2.Text = e.Row.Cells(11).Text
            ' 連結到待處理LIST
            ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep
            h2.NavigateUrl = "http://10.245.1.10/WorkFlow/BefOPList.aspx?pFormNo=003108&pFormSno=" + wFormsno + "&pStep=1&pSeqNo=1"
            h2.Target = "_blank"
            e.Row.Cells(11).Text = "履歷"
            e.Row.Cells(11).Controls.Add(h2)
            e.Row.Cells(0).Width = 100
            e.Row.Cells(1).Width = 100
            e.Row.Cells(2).Width = 100
            e.Row.Cells(3).Width = 100
            e.Row.Cells(4).Width = 100
            e.Row.Cells(5).Width = 100
            e.Row.Cells(6).Width = 100
            e.Row.Cells(7).Width = 100
            e.Row.Cells(8).Width = 100
            e.Row.Cells(9).Width = 120
            e.Row.Cells(10).Width = 100
            e.Row.Cells(11).Width = 100
            e.Row.Cells(12).Width = 100
            e.Row.Cells(13).Width = 100
            e.Row.Cells(14).Width = 100
            '  e.Row.Cells(12).Visible = False
       
        End If
        e.Row.Cells(14).Visible = False
        '連結
    End Sub

    Protected Sub BExcel_Click1(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click

        Dim wAllowPaging As Boolean = GridView1.AllowPaging   '程式別不同

        ' pFormNo = DFormName.SelectedValue
        GridView1.AllowPaging = False   '程式別不同
        DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=ComPlainOutScheduleALL.xls")     '程式別不同
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        GridView1.AllowPaging = wAllowPaging        '程式別不同


    End Sub

    Sub SetFieldData()
        Dim Sql As String
        Dim i As Integer
        Sql = "  Select  distinct substring(data,1,CHARINDEX('-',data)-1)Data  from M_referp"
        Sql = Sql & " where  cat = '3108'"
        Sql = Sql & " and dkey = 'COMDEPNAME'  "

        Dim dtReferp As DataTable = uDataBase.GetDataTable(Sql)
        DACCDEP1.Items.Clear()
        DACCDEP1.Items.Add("全部")
        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("Data")
            ListItem1.Value = dtReferp.Rows(i).Item("Data")
            DACCDEP1.Items.Add(ListItem1)
            DACCDEP1.SelectedValue = wDep
        Next
        dtReferp.Clear()

    End Sub

    Protected Sub DACCDEP1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DACCDEP1.SelectedIndexChanged
        DataList()
    End Sub
End Class

