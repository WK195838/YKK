Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

Partial Class BatchQAApprov
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
    Dim wUserID As String = ""
    Dim wName As String = ""
    Dim NowDateTime As String       '現在日期時間
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim ItemCount As Integer = 3 '預先定義欄位數量
    Dim wUserIP As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Response.Cookies("PGM").Value = "BatchApprovStandard.aspx"
        ' SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then

            '單一細項通知
            RunResultSend()
            '  自動承認
            RunBatch()


            Response.Write("<script   language=javascript>  window.opener=null;    window.open('','_self');  window.close();</script> ")
            'IE11
            Response.Write("<script>window.open('', '_self', ''); window.close();</script>")
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
        wUserID = Request.QueryString("pUserID")    'UserID
        wUserIP = Request.ServerVariables("REMOTE_ADDR")
    End Sub

    Sub CheckAuthority()
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     自動承認
    '**
    '*****************************************************************

    Sub RunBatch()
        Dim i As Integer
        Dim wError As Integer = 0
        Dim Message As String = ""
        Dim wNo As String = ""
        Dim ErrCode As Integer = 0
        '
        Dim wFun As String = ""
        Dim wAction As Integer = 0
        Dim wSts As String = ""
        Dim wStsDes As String = ""
        Dim wFormNo As String = ""
        Dim wFormSno As Integer = 0
        Dim wStep As Integer = 0
        Dim wSeqNo As Integer = 0
        Dim wDecideCalendar As String = ""
        Dim wDecideID As String = ""
        Dim wApplyID As String = ""
        Dim wAgentID As String = ""
        Dim wAllocteID As String = ""
        Dim wDecideDesc As String = ""
        Dim wTableName As String = ""
        '

        Dim SQL As String = ""
        '檢查 EDX判定是否全部OK，如果是就將流程結束-- EDX判定

        SQL = "   select a.no,a.formno,a.formsno,'20' as step,1 as seqno,'OK' as Fun,0 as action,1 as sts,'CL1' as decidecalendar, decideID,a.createuser,'OK' as DecideDesc,'F_QASheet' as TableName   from ("
        SQL = SQL & " select  formno,formsno,a.no,a.createuser,count(*)cuna  from F_QASheet a,F_QASheetDT b"
        SQL = SQL & "  where a.no = b.no "
        SQL = SQL & "  and sts =0"
        SQL = SQL & " group by  formno,a.no,formsno,a.createuser"
        SQL = SQL & " )a,"
        SQL = SQL & " ("
        SQL = SQL & " select  formsno,a.no,count(*)cunb  from F_QASheet a,F_QASheetDT b"
        SQL = SQL & " where a.no =b.no "
        SQL = SQL & " and sts =0 and ( result1 in ('OK','取消') or result2 ='OK') "
        SQL = SQL & " group by a.no,formsno"
        SQL = SQL & " )b,t_waithandle c  where a.no =b.no and cuna =cunb and a.formno =c.formno and a.formsno =c.formsno and c.step =20"
        Dim dt As DataTable = uDataBase.GetDataTable(SQL)
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                wFun = "OK"
                wAction = 0
                wSts = "1"
                wStsDes = "OK"

                'pFormNo	    表單代號
                'pFormSno	    單號
                'pStep		    工程代號
                'pSeqNo		    序號

                wFormNo = dt.Rows(i)("FormNo").ToString

                wFormSno = dt.Rows(i)("FormsNo").ToString
                wStep = dt.Rows(i)("step").ToString
                wSeqNo = dt.Rows(i)("seqno").ToString
                'pDecideID	    簽核者	
                'pDecideCalendar行事曆
                'pApplyID	    申請者
                'pAgentID	    被代理者
                'pAllocteID	    指定簽核者
                'pDecideDesc    承認說明
                'pTableName     表單Table
                wDecideID = dt.Rows(i)("decideID").ToString
                wDecideCalendar = oCommon.GetCalendarGroup(wDecideID)
                wApplyID = dt.Rows(i)("createuser").ToString
                wAgentID = ""
                wAllocteID = ""

                wDecideDesc = dt.Rows(i)("DecideDesc").ToString
                wTableName = dt.Rows(i)("TableName").ToString
                '
                'pNo    表單No
                wNo = dt.Rows(i)("No").ToString

                '採用AgentApprov暫存檔
                ErrCode = oCommon.RunAgentApprov(wFun, _
                                                    wAction, _
                                                    wSts, _
                                                    wStsDes, _
                                                    wFormNo, _
                                                    wFormSno, _
                                                    wStep, _
                                                    wSeqNo, _
                                                    wDecideCalendar, _
                                                    wDecideID, _
                                                    wApplyID, _
                                                    wAgentID, _
                                                    wAllocteID, _
                                                    wDecideDesc, _
                                                    wTableName, _
                                                    wUserIP)

                '更新已寄信的Flag  
                SQL = " update F_QASheetDT"
                SQL = SQL + " set result1send = 1"
                SQL = SQL + " where no = '" + dt.Rows(i)("No").ToString + "'"
                SQL = SQL + " and result1 ='OK'"
                uDataBase.ExecuteNonQuery(SQL)



                '更新已寄信的Flag
                SQL = " update F_QASheetDT"
                SQL = SQL + " set result2send = 1"
                SQL = SQL + " where no = '" + dt.Rows(i)("No").ToString + "'"
                SQL = SQL + " and result2 ='OK'"
                uDataBase.ExecuteNonQuery(SQL)

                '更新分析完了日
                SQL = " update F_QASheet"
                SQL = SQL + " set QCDate =convert(char(10),getdate(),111)  "
                SQL = SQL + " where no = '" + dt.Rows(i)("No").ToString + "'"
                uDataBase.ExecuteNonQuery(SQL)


            Next
           

        End If

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     單一細項通知
    '**
    '*****************************************************************
    Sub RunResultSend()
        Dim SQL As String = ""
        Dim SQL1 As String = ""
        Dim SQL2 As String = ""
        Dim SQL3 As String = ""
        Dim i As Integer
        Dim j As Integer

        '判段是否有Result1 =OK 但未通和的信
        SQL = "     select  a.formno,a.formsno  from F_QASheet a,F_QASheetDT b "
        SQL = SQL + " where a.no = b.no"
        SQL = SQL + "  and sts =0 and result1 ='OK' and Result1Send =0 "
        Dim dt As DataTable = uDataBase.GetDataTable(SQL)
        If dt.Rows.Count > 0 Then
            SQL1 = "   select  decideid,b.applyid,b.applyid,a.formno,a.formsno,step     from ("
            SQL1 = SQL1 + "   select  distinct a.formno,a.formsno   from F_QASheet a,F_QASheetDT b "
            SQL1 = SQL1 + " where a.no = b.no"
            SQL1 = SQL1 + " and sts =0 and  result1 = 'OK' and Result1Send =0 "
            SQL1 = SQL1 + " )a,t_waithandle b"
            SQL1 = SQL1 + " where a.formno = b.formno and a.formsno =b.formsno and step =20"
            Dim dt1 As DataTable = uDataBase.GetDataTable(SQL1)
            If dt1.Rows.Count > 0 Then

                For i = 0 To dt1.Rows.Count - 1
                    oCommon.Send(dt1.Rows(i)("decideid"), dt1.Rows(i)("applyid"), dt1.Rows(i)("applyid"), dt1.Rows(i)("formno"), dt1.Rows(i)("formsno"), "21", "FLOW")
                    '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別        

                    ' 生管Master CC 
                    SQL3 = " select *from M_group  where  groupid = '008002-1'"
                    Dim dt3 As DataTable = uDataBase.GetDataTable(SQL3)
                    If dt3.Rows.Count > 0 Then
                        For j = 0 To dt3.Rows.Count - 1
                            oCommon.Send(dt1.Rows(i)("decideid"), dt3.Rows(j)("userid"), dt3.Rows(j)("userid"), dt1.Rows(i)("formno"), dt1.Rows(i)("formsno"), "21", "FLOW")
                        Next
                    End If

                Next
            End If

            '更新已寄信的Flag
            SQL = "update b   "
            SQL = SQL + " set Result1Send=1  "
            SQL = SQL + " from F_QASheet a,F_QASheetDT b "
            SQL = SQL + " where a.no = b.no"
            SQL = SQL + " and sts =0 and result1 ='OK' and Result1Send =0 "
            uDataBase.ExecuteNonQuery(SQL)

        End If

        '判段是否有Result2 =OK 但未通和的信
        SQL = "     select  a.formno,a.formsno  from F_QASheet a,F_QASheetDT b "
        SQL = SQL + " where a.no = b.no"
        SQL = SQL + "  and sts =0 and result2 in('OK','取消')  and Result2Send =0 "
        Dim dt2 As DataTable = uDataBase.GetDataTable(SQL)
        If dt2.Rows.Count > 0 Then
            SQL1 = "   select  decideid,b.applyid,b.applyid,a.formno,a.formsno,step     from ("
            SQL1 = SQL1 + "   select  distinct a.formno,a.formsno   from F_QASheet a,F_QASheetDT b "
            SQL1 = SQL1 + " where a.no = b.no"
            SQL1 = SQL1 + " and sts =0 and result2 ='OK' and Result2Send =0 "
            SQL1 = SQL1 + " )a,t_waithandle b"
            SQL1 = SQL1 + " where a.formno = b.formno and a.formsno =b.formsno and step =20"
            Dim dt1 As DataTable = uDataBase.GetDataTable(SQL1)
            If dt1.Rows.Count > 0 Then

                For i = 0 To dt1.Rows.Count - 1
                    oCommon.Send(dt1.Rows(i)("decideid"), dt1.Rows(i)("applyid"), dt1.Rows(i)("applyid"), dt1.Rows(i)("formno"), dt1.Rows(i)("formsno"), "21", "FLOW")
                    '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別        

                    ' 生管Master CC 
                    SQL3 = " select *from M_group  where  groupid = '008002-1'"
                    Dim dt3 As DataTable = uDataBase.GetDataTable(SQL3)
                    If dt3.Rows.Count > 0 Then
                        For j = 0 To dt3.Rows.Count - 1
                            oCommon.Send(dt1.Rows(i)("decideid"), dt3.Rows(i)("userid"), dt3.Rows(i)("userid"), dt1.Rows(i)("formno"), dt1.Rows(i)("formsno"), "21", "FLOW")
                        Next
                    End If

                Next
            End If

            '更新已寄信的Flag
            SQL = "update b   "
            SQL = SQL + " set Result2Send=1  "
            SQL = SQL + " from F_QASheet a,F_QASheetDT b "
            SQL = SQL + " where a.no = b.no"
            SQL = SQL + " and sts =0 and result2 ='OK' and Result2Send =0 "
            uDataBase.ExecuteNonQuery(SQL)

        End If



    End Sub

End Class
