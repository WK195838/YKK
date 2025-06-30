Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
Imports System.IO



Partial Class DispoalSend
    Inherits System.Web.UI.Page
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Server.ScriptTimeout = 900      ' 設定逾時時間
        If Not IsPostBack Then
            SendMail30()
            ' SendMail80()
            SendMail130()
            SendMail140()


            'IE8 可以自行關閉網頁
            Response.Write("<script   language=javascript>  window.opener=null;    window.open('','_self');  window.close();</script> ")
            'IE11
            Response.Write("<script>window.open('', '_self', ''); window.close();</script>")


        End If

      
        
    End Sub
    Sub SendMail80()
        Dim SQL, SQL1 As String
        Dim wStep, wToID, wApplyID, wFormSno As String

        ' 先判斷今天是否是預設要執行報廢簽核的時間
        SQL = " select  substring(dkey,10,len(dkey)-1)step ,data  from M_referp"
        SQL = SQL + " where cat = 6001 "
        SQL = SQL + " and left(dkey,14) = 'SendTime-admin'"
        SQL = SQL + " and data = convert(char(10),getdate(),111) "
        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL)
        If dtFlow.Rows.Count > 0 Then
            wStep = "80"

            SQL1 = " SELECT  top 1 workid,applyid,formno,formsno,step  FROM  t_waithandle"
            SQL1 = SQL1 + " where active=1 and  formno ='006001'"
            SQL1 = SQL1 + " and step  =" + wStep + ""
            SQL1 = SQL1 + " and sts <>5 "
            Dim dtSendData As DataTable = uDataBase.GetDataTable(SQL1)
            If dtSendData.Rows.Count > 0 Then
                wToID = dtSendData.Rows(0)("workid")
                wApplyID = dtSendData.Rows(0)("Applyid")
                wFormSno = dtSendData.Rows(0)("FormSno")
                oCommon.Send("System", wToID, wApplyID, "006001", wFormSno, wStep, "FLOW")
                '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別

            End If
            
        End If
       
    End Sub

    Sub SendMail30()
        Dim SQL, SQL1 As String
        Dim wStep, wToID, wApplyID, wFormSno As String

        ' 先判斷今天是否是預設要執行報廢簽核的時間
        SQL = " select  substring(dkey,10,len(dkey)-1)step ,data  from M_referp"
        SQL = SQL + " where cat = 6001 "
        SQL = SQL + " and left(dkey,12) = 'SendTime-pcs'"
        SQL = SQL + " and data = convert(char(10),getdate(),111) "
        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL)
        If dtFlow.Rows.Count > 0 Then
            wStep = "30"

            SQL1 = " SELECT  top 1 workid,applyid,formno,formsno,step  FROM  t_waithandle"
            SQL1 = SQL1 + " where active=1 and  formno ='006001'"
            SQL1 = SQL1 + " and step  in(" + wStep + ")"
            SQL1 = SQL1 + " and sts <>5"
            Dim dtSendData As DataTable = uDataBase.GetDataTable(SQL1)
            If dtSendData.Rows.Count > 0 Then
                wToID = dtSendData.Rows(0)("workid")
                wApplyID = dtSendData.Rows(0)("Applyid")
                wFormSno = dtSendData.Rows(0)("FormSno")
                oCommon.Send("System", wToID, wApplyID, "006001", wFormSno, wStep, "FLOW")

                '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別

            End If

        End If

    End Sub

    Sub SendMail130()
        Dim SQL, SQL1, SQL2 As String
        Dim wStep, wToID, wApplyID, wFormSno As String

        ' 檢查給130工程的件數是否全到了

        SQL = " select * from ("
        SQL = SQL + " select disposalym,count(*)cun from ("
        SQL = SQL + " select formsno from t_waithandle "
        SQL = SQL + " where formno ='006001' and step = 130 and active =1"
        SQL = SQL + " )a,f_disposalsheet b "
        SQL = SQL + " where a.formsno = b.formsno"
        SQL = SQL + " group by disposalym"
        SQL = SQL + " )a,"
        SQL = SQL + " ("
        SQL = SQL + " select disposalym,count(*)cun  from f_disposalsheet "
        SQL = SQL + " where sts not in (2,3) and stype ='月次'"
        SQL = SQL + " group by disposalym"
        SQL = SQL + " )b where a.disposalym = b.disposalym"
        SQL = SQL + " and a.cun =b.cun"


        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL)

        If dtFlow.Rows.Count > 0 Then
            Dim disposalym = dtFlow.Rows(0)("disposalym")
            wStep = "130"
            ' 判斷是否寄過信了，若沒有才寄
            SQL1 = " select workid,applyid,formno,formsno,step,count(*)cun  from ("
            SQL1 = SQL1 + "  SELECT top 1 workid,applyid,formno,formsno,step,substring(convert(char(10),bstarttime,111),1,7)BStartDate  FROM  t_waithandle  "
            SQL1 = SQL1 + " where  formno ='006001' and step = 130 and active =1 and substring(convert(char(10),bstarttime,111),1,7) ='" + disposalym + "' order by formsno desc "
            SQL1 = SQL1 + " )a where formsno  in ("
            SQL1 = SQL1 + " select  formsno  from  Q_SendOff  where formno = '006001' and step  =130 "
            SQL1 = SQL1 + " and  msgname = '核定')"
            SQL1 = SQL1 + " group by workid,applyid,formno,formsno,step"
            Dim dtSendData As DataTable = uDataBase.GetDataTable(SQL1)
            If dtSendData.Rows.Count = 0 Then
                SQL2 = "  SELECT top 1 workid,applyid,formno,formsno,step,substring(convert(char(10),bstarttime,111),1,7)BStartDate  FROM  t_waithandle  "
                SQL2 = SQL2 + " where  formno ='006001' and step = 130 and active =1 and substring(convert(char(10),bstarttime,111),1,7) ='" + disposalym + "' order by formsno desc "
                Dim dtFlow1 As DataTable = uDataBase.GetDataTable(SQL2)

                wToID = dtFlow1.Rows(0)("workid")
                wApplyID = dtFlow1.Rows(0)("Applyid")
                wFormSno = dtFlow1.Rows(0)("FormSno")
                oCommon.Send("System", wToID, wApplyID, "006001", wFormSno, wStep, "FLOW")
            End If
        End If
    End Sub

    Sub SendMail140()
        Dim SQL, SQL1, SQL2 As String
        Dim wStep, wToID, wApplyID, wFormSno As String

        ' 檢查給140工程的件數是否全到了

        SQL = " select * from ("
        SQL = SQL + " select disposalym,count(*)cun from ("
        SQL = SQL + " select formsno from t_waithandle "
        SQL = SQL + " where formno ='006001' and step = 140 and active =1"
        SQL = SQL + " )a,f_disposalsheet b "
        SQL = SQL + " where a.formsno = b.formsno"
        SQL = SQL + " group by disposalym"
        SQL = SQL + " )a,"
        SQL = SQL + " ("
        SQL = SQL + " select disposalym,count(*)cun  from f_disposalsheet "
        SQL = SQL + " where sts not in (2,3) and stype ='月次'"
        SQL = SQL + " group by disposalym"
        SQL = SQL + " )b where a.disposalym = b.disposalym"
        SQL = SQL + " and a.cun =b.cun"

        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL)

        If dtFlow.Rows.Count > 0 Then
            Dim disposalym = dtFlow.Rows(0)("disposalym")
            wStep = "140"
            ' 判斷是否寄過信了，若沒有才寄
            SQL1 = " select workid,applyid,formno,formsno,step,count(*)cun  from ("
            SQL1 = SQL1 + "  SELECT top 1 workid,applyid,formno,formsno,step,substring(convert(char(10),bstarttime,111),1,7)BStartDate  FROM  t_waithandle  "
            SQL1 = SQL1 + " where  formno ='006001' and step = 140 and active =1 and substring(convert(char(10),bstarttime,111),1,7) ='" + disposalym + "' order by formsno desc "
            SQL1 = SQL1 + " )a where formsno  in ("
            SQL1 = SQL1 + " select  formsno  from  Q_SendOff  where formno = '006001' and step  =140 "
            SQL1 = SQL1 + " and  msgname = '核定')"
            SQL1 = SQL1 + " group by workid,applyid,formno,formsno,step"
            Dim dtSendData As DataTable = uDataBase.GetDataTable(SQL1)
            If dtSendData.Rows.Count = 0 Then
                SQL2 = "  SELECT top 1 workid,applyid,formno,formsno,step,substring(convert(char(10),bstarttime,111),1,7)BStartDate  FROM  t_waithandle  "
                SQL2 = SQL2 + " where  formno ='006001' and step = 140 and active =1 and substring(convert(char(10),bstarttime,111),1,7) ='" + disposalym + "' order by formsno desc "
                Dim dtFlow1 As DataTable = uDataBase.GetDataTable(SQL2)

                wToID = dtFlow1.Rows(0)("workid")
                wApplyID = dtFlow1.Rows(0)("Applyid")
                wFormSno = dtFlow1.Rows(0)("FormSno")
                oCommon.Send("System", wToID, wApplyID, "006001", wFormSno, wStep, "FLOW")

            End If
        End If
    End Sub

End Class
