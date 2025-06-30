Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System

Partial Class AddItemListACC
    Inherits System.Web.UI.Page

    Dim FieldName(60) As String     '各欄位
    Dim Attribute(60) As Integer    '各欄位屬性    
    Dim Top As String               '預設的控制項位置
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wStep As Integer            '工程代碼
    Dim wbFormSno As Integer        '連續起單-表單流水號
    Dim wbStep As Integer           '連續起單-工程代碼
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者ID
    Dim wAgentID As String          '被代理人ID
    Dim NowDateTime As String       '現在日期時間
    Dim wNextGate As String         '下一關人
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    Dim wUserName As String = ""            '姓名代理用
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    Dim remark1, remark2 As String
    Dim DTNo As String
    Dim Days As String
    Dim AEDate As String
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '*****************************************************************

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Errcode As Integer = 0
        Dim Message As String = ""

        BReference.Attributes("onclick") = "GetReference();"


        If Request.QueryString("pDays") <> "" Then
 


            NowDateTime = Now.ToString("yyMMddHHmmss")
            'DTNo = NowDateTime
            ' DTNo = "23030908302001"

            Days = Request.QueryString("pDays")
            wStep = Request.QueryString("pwStep")
            DTNo = Request.QueryString("pDTNo")
            AEDate = Request.QueryString("pAEDate")

            If Not Me.IsPostBack Then   '不是PostBack
                If wStep = 30 Or wStep = 50 Then
                    BEDIT.Visible = True
                    'DCurrency1.Visible = True
                    'DRate1.Visible = True
                    'BChange.Visible = True
                    DCurrency1.Visible = False
                    DRate1.Visible = False
                    BChange.Visible = False
                    GridView1.AutoGenerateSelectButton = True
                    GridView1.Columns(8).HeaderText = "小計金額（台幣)"
                Else
                    BEDIT.Visible = False
                    DCurrency1.Visible = False
                    DRate1.Visible = False
                    BChange.Visible = False
                    GridView1.AutoGenerateSelectButton = False
                    GridView1.Columns(8).HeaderText = "小計金額"
                End If




                SetFieldData()
                SetFieldAttribute("New")
                GetData()
            Else


            End If


        End If
        '取經費NO

        GetExpenesNo()


        '
    End Sub


    Protected Sub BADD_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BADD.ServerClick
       
   


        Dim money As String = ""
        '預支款自動代入負數
        '預支款自動代入負數
        If DType.SelectedItem.Text = "預支款沖銷" Then
            money = "-" + DMoney.Text
        Else
            money = DMoney.Text
        End If


        Dim SQL As String
        GetRate()

        If InputDataOK(0) Then
            SQL = " Insert into F_CloseAccountSheetDTTemp (No,Typeno ,Type,SDate,Pay,Currency,Money,Days,Rate,SumAmt,Remark,ExpenseNo,Createuser,CreateTime,ModifyUser,ModifyTime) "
            SQL = SQL & "VALUES( "
            SQL = SQL & " '" & DTNo & "', "
            SQL = SQL & " '" & DType.SelectedValue & "', "
            SQL = SQL & " '" & DType.SelectedItem.Text & "', "
            SQL = SQL & " '" & DSDate.Text & "', "
            SQL = SQL & " '" & DPay.SelectedValue & "', "
            SQL = SQL & " '" & DCurrency.SelectedItem.Text & "', "
            SQL = SQL & " '" & money & "', "
            SQL = SQL & " '" & DDays.Text & "', "
            If DPay.SelectedValue = "公司卡" Then
                SQL = SQL & " 0, "
            Else
                SQL = SQL & " '" & DRate.Text & "', "
            End If

            SQL = SQL & " 0, "
            SQL = SQL & " N'" & DRemark.Text & "', "
            SQL = SQL & " '" & DExpenseNo.Text & "', "
            SQL = SQL & " '" & Request.QueryString("pUserID") & "', "
            SQL = SQL & " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "', "
            SQL = SQL & " '" & Request.QueryString("pUserID") & "', "
            SQL = SQL & " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "') "
            uDataBase.ExecuteNonQuery(SQL)

        End If

        '恢復未選取
        DID.Text = ""
        GridView1.SelectedIndex = -1
        'If wStep = 30 Or wStep = 50 Then
        '    SQL = " select *, convert(decimal(9,0),money*days*rate) as SumAmt from ("
        'Else
        '    SQL = " select *, cast(round(money*days*rate ,2) as numeric(10,2)) as SumAmt from ("
        'End If

        'SQL = SQL + " select Unique_ID,type,typeno,currency,convert(char(10),sdate,111)sdate,pay,case when money = '' then '0' else  convert(decimal(9,2),money) end as money,"
        'SQL = SQL + "  case when days='' then '0' else convert(decimal(9,0),days) end as days"
        'SQL = SQL + " ,case when pay = '公司卡' then 0 when currency ='TWD' then 1 when Rate ='' then 1  else  convert(decimal(9,6),rate) end as rate,remark,expenseno"
        'SQL = SQL + " from  F_CloseAccountSheetDTTemp "
        'SQL = SQL + "  where 1=1 and No='" + DTNo + "')a "
        'SQL = SQL + " order by  typeno,sdate "


        SQL = "  select  UNIQUE_ID,TYPE,TYPENO,CURRENCY,SDATE,PAY,MONEY,DAYS,case when pay ='公司卡' then 0 else  AVERAGERATE end AS RATE,REMARK,EXPENSENO,case when pay ='公司卡' then 0 else  round(money*days*AVERAGERATE,0) end  as SumAmt from ("
        SQL = SQL + " Select* from ("
        SQL = SQL + " select Unique_ID,type,typeno,currency,convert(char(10),sdate,111)sdate,pay,case when money = '' then '0' else  convert(decimal(9,2),money) end as money,"
        SQL = SQL + "  case when days='' then '0' else convert(decimal(9,0),days) end as days"
        SQL = SQL + " ,case when pay = '公司卡' then 0 when currency ='TWD' then 1 when Rate ='' then 1  else  convert(decimal(9,6),rate) end as rate,remark,expenseno"
        SQL = SQL + " from  F_CloseAccountSheetDTTemp "
        SQL = SQL + "  where 1=1 and No='" + DTNo + "' "
        SQL = SQL + " )a,("
        SQL = SQL + " select currencycode,syear+'/'+smonth+'/'+substring(ltrim(sday),1,2) as StrD,syear+'/'+smonth+'/'+substring(ltrim(sday),4,2) as EndD,AVERAGERATE  from F_exchangerate"
        SQL = SQL + "  )b where a.currency =b.currencycode "
        SQL = SQL + " and '" + AEDate + "' between strD and EndD )a  "
        SQL = SQL + " order by  typeno,sdate "



        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()


    End Sub

    Sub GetData()
        Dim SQL As String

        If wStep <> 1 Then
            SQL = " delete from  F_CloseAccountSheetDTTemp where no='" + DTNo + "'"
            uDataBase.ExecuteNonQuery(SQL)

            SQL = " insert into F_CloseAccountSheetdttemp (No,TypeNo,Type,Sdate,Pay,Currency,Money,Days,Rate,Sumamt,ExpenseNo,remark,CreateUser, CreateTime, ModifyUser, ModifyTime) "
            'SQL = SQL + " select no,"
            'SQL = SQL + " TypeNo,Type,Sdate,Pay,Currency,Money,Days,Rate, convert(decimal(9,0),money*days*rate) as SumAmt,ExpenseNo,Remark,'" + Request.QueryString("pUserID") + "',"
            'SQL = SQL + " getdate(),'" + Request.QueryString("pUserID") + "',getdate() from ("
            'SQL = SQL + " select no,Unique_ID,type,typeno,currency,convert(char(10),sdate,111)sdate,pay,case when money = '' then '0' else  convert(decimal(9,2),money) end as money,"
            'SQL = SQL + "  case when days='' then '0' else convert(decimal(9,0),days) end as days"
            'SQL = SQL + " ,case when pay = '公司卡' then 0 when currency ='TWD' then 1 when Rate ='' then 1  else  convert(decimal(9,6),rate) end as rate,remark,expenseno"
            'SQL = SQL + " from  F_CloseAccountSheetDT "
            'SQL = SQL + "  where 1=1 and No='" + DTNo + "')a "
            'SQL = SQL + " order by typeno,sdate  "

            SQL = SQL + " select no,TYPEno,TYPE,SDATE,PAY,CURRENCY,MONEY,DAYS,AVERAGERATE AS RATE,cast(round(money*days*AVERAGERATE ,2) as numeric(10,2)) as SumAmt,EXPENSENO,REMARK,'" + Request.QueryString("pUserID") + "', "
            SQL = SQL + " getdate(),'" + Request.QueryString("pUserID") + "',getdate() from ("
            SQL = SQL + " select * from ("
            SQL = SQL + " select no,Unique_ID,type,typeno,currency,convert(char(10),sdate,111)sdate,pay,case when money = '' then '0' else  convert(decimal(9,2),money) end as money,"
            SQL = SQL + "  case when days='' then '0' else convert(decimal(9,0),days) end as days"
            SQL = SQL + " ,case when pay = '公司卡' then 0 when currency ='TWD' then 1 when Rate ='' then 1  else  convert(decimal(9,6),rate) end as rate,remark,expenseno"
            SQL = SQL + " from  F_CloseAccountSheetDT "
            SQL = SQL + "  where 1=1 and No='" + DTNo + "')a "
            SQL = SQL + " ,("
            SQL = SQL + " select currencycode,syear+'/'+smonth+'/'+substring(ltrim(sday),1,2) as StrD,syear+'/'+smonth+'/'+substring(ltrim(sday),4,2) as EndD,AVERAGERATE  from F_exchangerate"
            SQL = SQL + " )b where a.currency =b.currencycode "
            SQL = SQL + " and '" + AEDate + "' between strD and EndD )a  "
            SQL = SQL + " order by typeno,sdate  "
            uDataBase.ExecuteNonQuery(SQL)
        End If



       



   

        SQL = "  select  UNIQUE_ID,TYPE,TYPENO,CURRENCY,SDATE,PAY,MONEY,DAYS,case when pay ='公司卡' then 0 else  AVERAGERATE end AS RATE,REMARK,EXPENSENO,case when pay ='公司卡' then 0 else  round(money*days*AVERAGERATE,0) end  as SumAmt from ("
        SQL = SQL + " Select* from ("
        SQL = SQL + " select Unique_ID,type,typeno,currency,convert(char(10),sdate,111)sdate,pay,case when money = '' then '0' else  convert(decimal(9,2),money) end as money,"
        SQL = SQL + "  case when days='' then '0' else convert(decimal(9,0),days) end as days"
        SQL = SQL + " ,case when pay = '公司卡' then 0 when currency ='TWD' then 1 when Rate ='' then 1  else  convert(decimal(9,6),rate) end as rate,remark,expenseno"
        SQL = SQL + " from  F_CloseAccountSheetDTTemp "
        SQL = SQL + "  where 1=1 and No='" + DTNo + "' "
        SQL = SQL + " )a,("
        SQL = SQL + " select currencycode,syear+'/'+smonth+'/'+substring(ltrim(sday),1,2) as StrD,syear+'/'+smonth+'/'+substring(ltrim(sday),4,2) as EndD,AVERAGERATE  from F_exchangerate"
        SQL = SQL + "  )b where a.currency =b.currencycode "
        SQL = SQL + " and '" + AEDate + "' between strD and EndD )a  "
        SQL = SQL + " order by  typeno,sdate "



        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()

        If wStep = 30 Or wStep = 50 Then
            GridView1.AutoGenerateSelectButton = True
        Else
            GridView1.AutoGenerateSelectButton = False
        End If


    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        e.Row.Cells(1).Visible = False
        'If wStep = 30 Or wStep = 50 Then
        '    e.Row.Cells(8).Visible = True
        'Else

        '    e.Row.Cells(8).Visible = False
        'End If
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Dim SQL As String

        Dim id As String = GridView1.DataKeys(e.RowIndex).Value

        SQL = " delete  from   F_CloseAccountSheetDTTemp "
        SQL = SQL & " where unique_id = " & id & " "
        uDataBase.ExecuteNonQuery(SQL)
        '
        DID.Text = ""

        SQL = " select *, convert(decimal(9,0),money*days*rate) as SumAmt from ("
        SQL = SQL + " select Unique_ID,type,typeno,currency,convert(char(10),sdate,111)sdate,pay,case when money = '' then '0' else  convert(decimal(9,2),money) end as money,"
        SQL = SQL + "  case when days='' then '0' else convert(decimal(9,0),days) end as days"
        SQL = SQL + " ,case when pay = '公司卡' then 0 when currency ='TWD' then 1 when Rate ='' then 1  else  convert(decimal(9,6),rate) end as rate,remark,expenseno"
        SQL = SQL + " from  F_CloseAccountSheetDTTemp "
        SQL = SQL + "  where 1=1 and No='" + DTNo + "')a "
        SQL = SQL + " order by  typeno,sdate "



        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()

    End Sub


    Protected Sub BClose_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BClose.ServerClick
        Dim sql As String

        Sql = " delete from  F_CloseAccountSheetDT where no='" + DTNo + "'"
        uDataBase.ExecuteNonQuery(Sql)

        Sql = " insert into F_CloseAccountSheetdt (No,TypeNo,Type,Sdate,Pay,Currency,Money,Days,Rate,Sumamt,ExpenseNo,remark,CreateUser, CreateTime, ModifyUser, ModifyTime) "

        sql = sql + " select no,TYPEno,TYPE,SDATE,PAY,CURRENCY,MONEY,DAYS,AVERAGERATE AS RATE,round(money*days*AVERAGERATE,0) as SumAmt,EXPENSENO,REMARK,'" + Request.QueryString("pUserID") + "', "
        sql = sql + " getdate(),'" + Request.QueryString("pUserID") + "',getdate() from ("
        sql = sql + " select * from ("
        sql = sql + " select no,Unique_ID,type,typeno,currency,convert(char(10),sdate,111)sdate,pay,case when money = '' then '0' else  convert(decimal(9,2),money) end as money,"
        sql = sql + "  case when days='' then '0' else convert(decimal(9,0),days) end as days"
        sql = sql + " ,case when pay = '公司卡' then 0 when currency ='TWD' then 1 when Rate ='' then 1  else  convert(decimal(9,6),rate) end as rate,remark,expenseno"
        sql = sql + " from  F_CloseAccountSheetDTTemp "
        sql = sql + "  where 1=1 and No='" + DTNo + "')a "
        sql = sql + " ,("
        sql = sql + " select currencycode,syear+'/'+smonth+'/'+substring(ltrim(sday),1,2) as StrD,syear+'/'+smonth+'/'+substring(ltrim(sday),4,2) as EndD,AVERAGERATE  from F_exchangerate"
        sql = sql + " )b where a.currency =b.currencycode "
        sql = sql + " and '" + AEDate + "' between strD and EndD )a  "
        sql = sql + " order by typeno,sdate  "

        uDataBase.ExecuteNonQuery(Sql)



        Dim js As String = ""
        js &= "window.opener.document.all.DUpdate.value = '1';"
        ' js &= "window.close();"
        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        '
        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

        ' 

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(FindFieldInf)
    '**     尋找表單欄位屬性
    '**
    '*****************************************************************
    Function FindFieldInf(ByVal pFieldName As String) As Integer
        Dim Run As Boolean
        Dim i As Integer
        Run = True
        FindFieldInf = 9
        While i <= 60 And Run
            If FieldName(i) = pFieldName Then
                FindFieldInf = Attribute(i)
                Run = False
            End If
            i = i + 1
        End While
    End Function


    '*****************************************************************
    '**(SetFieldAttribute)
    '**     表單各欄位屬性及欄位輸入檢查等設定
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BExpense.Attributes.Add("onClick", "AddExpense()") '交際

    End Sub

    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData()

        Dim SQL As String

        '類別

        Dim i As Integer
        SQL = "  SELECT  substring(data,1,2) seqno, substring(data,4,len(data)-1)Data1,data    from M_referp  "
        SQL = SQL + " where cat = '3115' and dkey ='Type' order by data    "

        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
        DType.Items.Clear()
        DType.Items.Add("")
        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("data1")
            ListItem1.Value = dtReferp.Rows(i).Item("seqno")
            DType.Items.Add(ListItem1)
        Next
        dtReferp.Clear()

        '支付
        SQL = "  SELECT   Data   from M_referp  "
        SQL = SQL + " where cat = '3115' and dkey ='Pay' order by Data   "
        Dim dtReferp1 As DataTable = uDataBase.GetDataTable(SQL)
        DPay.Items.Clear()
        DPay.Items.Add("")
        For i = 0 To dtReferp1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp1.Rows(i).Item("Data")
            ListItem1.Value = dtReferp1.Rows(i).Item("Data")
            DPay.Items.Add(ListItem1)
        Next
        dtReferp1.Clear()






        '幣別
        SQL = "  SELECT  substring(data,4,len(data)-1)Data1,data  from M_referp    "
        SQL = SQL + " where cat = '3114' and dkey ='Currency' order by Data   "
        Dim dtReferp4 As DataTable = uDataBase.GetDataTable(SQL)
        DCurrency.Items.Clear()
        DCurrency.Items.Add("")

        DCurrency1.Items.Clear()
        DCurrency1.Items.Add("")

        For i = 0 To dtReferp4.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp4.Rows(i).Item("Data1")
            ListItem1.Value = dtReferp4.Rows(i).Item("Data1")
            DCurrency.Items.Add(ListItem1)
            DCurrency1.Items.Add(ListItem1)
        Next
        dtReferp4.Clear()



        '日當
        SQL = "  SELECT  substring(data,4,len(data)-1)Data1,data  from M_referp    "
        SQL = SQL + " where cat = '3115' and dkey ='DayMoney' order by Data   "
        Dim dtReferp2 As DataTable = uDataBase.GetDataTable(SQL)
        DDayMoney.Items.Clear()
        DDayMoney.Items.Add("")

        For i = 0 To dtReferp2.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp2.Rows(i).Item("Data1")
            ListItem1.Value = dtReferp2.Rows(i).Item("Data1")
            DDayMoney.Items.Add(ListItem1)

        Next
        dtReferp2.Clear()






    End Sub

    '*****************************************************************
    '**(ShowRequiredFieldValidator)
    '**     建立表單需輸入欄位
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator

        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.Text = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", Top + 20 & "px")
        rqdVal.Style.Add("Left", "8px")
        rqdVal.Style.Add("Height", "20px")
        rqdVal.Style.Add("Width", "250px")
        rqdVal.Style.Add("POSITION", "absolute")
        Me.Form.Controls.Add(rqdVal)
    End Sub

    Function InputDataOK(ByVal pAction As Integer) As Boolean

        Dim Message As String = ""
        Dim isOK As Boolean = False
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False


        If ErrCode = 0 Then
            isOK = IsNumeric(DMoney.Text)
            If isOK = False Then
                ErrCode = 12
            End If

        End If



        If ErrCode = 0 Then
            isOK = IsNumeric(DDays.Text)
            If isOK = False Then
                ErrCode = 13
            End If

        End If



        If ErrCode = 0 Then
            If DType.SelectedItem.Text = "" Then
                ErrCode = 4
            End If
        End If





        '檢查日期是否有輸入
        If ErrCode = 0 Then
            If DSDate.Text = "" Then
                ErrCode = 2
            End If
        End If

        '        If DPay.SelectedValue <> "公司卡" Then
        If ErrCode = 0 Then
            If DPay.SelectedValue = "" Then
                ErrCode = 7
            End If
        End If


        If ErrCode = 0 Then
            If DMoney.Text = "" Then
                ErrCode = 5
            End If
        End If

        If ErrCode = 0 Then
            If DDays.Text = "" Then
                ErrCode = 9
            End If
        End If


        If ErrCode = 0 Then
            If DCurrency.SelectedItem.Text = "" Then
                ErrCode = 6
            End If
        End If


        If DType.SelectedItem.Text = "膳雜費-日當" Then
            If DDayMoney.SelectedValue = "" Then
                ErrCode = 10
            End If
        End If


        If ErrCode = 0 Then
            If DRemark.Text = "" Then
                ErrCode = 11
            End If
        End If

        If Mid(DType.SelectedItem.Text, 1, 3) = "交際費" Then
            If DExpenseNo.Text = "" Then
                ErrCode = 14
            End If
        End If
 
        '       End If
        '
        Dim CheckSQL As String


        '預定
        '開始日期
        Dim ASdate As Date
        Dim ASdate1 As String = ""
        If DSDate.Text <> "" Then
            ASdate = DSDate.Text
            ASdate1 = Format(ASdate, "yyyy/MM/dd")

        End If

    


        '2024/08/13 增加檢查匯率
        If DSDate.Text <> "" Then
            CheckSQL = " select * from ("
            CheckSQL = CheckSQL + " select currencycode,syear+'/'+smonth+'/'+substring(ltrim(sday),1,2) as StrD,syear+'/'+smonth+'/'+substring(ltrim(sday),4,2) as EndD,AVERAGERATE  from F_exchangeratE"
            CheckSQL = CheckSQL + "  )a WHERE '" + ASdate1 + "' BETWEEN STRD AND ENDD "
            CheckSQL = CheckSQL + " AND currencycode ='" + DCurrency.SelectedItem.Text + "'"
            Dim dtData As DataTable = uDataBase.GetDataTable(CheckSQL)
            If dtData.Rows.Count = 0 Then
                ErrCode = 15
            End If
        End If



        If ErrCode = 1 Then Message = "異常：[到達/退房]日期必需大於[出發/入住]日期"
        If ErrCode = 2 Then Message = "異常：需輸入日期"
        If ErrCode = 3 Then Message = "異常：需輸入[到達/退房]日期"
        If ErrCode = 4 Then Message = "異常：需輸入類別"
        If ErrCode = 5 Then Message = "異常：需輸入金額"
        If ErrCode = 6 Then Message = "異常：需輸入幣別"
        If ErrCode = 7 Then Message = "異常：需輸入支付"
        If ErrCode = 8 Then Message = "異常：需輸入內容"
        If ErrCode = 9 Then Message = "異常：需輸入次數"
        If ErrCode = 10 Then Message = "異常：需選擇職等"
        If ErrCode = 11 Then Message = "異常：需輸入備註"
        If ErrCode = 12 Then Message = "異常：金額請輸入數字"
        If ErrCode = 13 Then Message = "異常：數量/次數請輸入數字"
        If ErrCode = 14 Then Message = "異常：需輸入交際費"
        If ErrCode = 15 Then Message = "異常：請確定匯率日期是否正確？"

        If ErrCode <> 0 Then
            isOK = False
            uJavaScript.PopMsg(Me, Message)

        Else
            isOK = True
        End If

        Return isOK


    End Function



    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        DRate.Enabled = True
        Dim SQL As String

        Dim id As String = GridView1.SelectedValue

        SQL = " select* from  F_CloseAccountSheetDTTemp "
        SQL = SQL & " where unique_id = " & id & " "
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        '

        DType.SelectedItem.Text = dtData.Rows(0).Item("Type")
        DSDate.Text = dtData.Rows(0).Item("SDate")
        DPay.SelectedValue = dtData.Rows(0).Item("Pay")
        DCurrency.SelectedItem.Text = dtData.Rows(0).Item("Currency")
        DMoney.Text = Math.Abs(Convert.ToDouble(dtData.Rows(0).Item("Money"))).ToString
        DDays.Text = dtData.Rows(0).Item("Days")
        DRate.Text = dtData.Rows(0).Item("Rate")
        DExpenseNo.Text = dtData.Rows(0).Item("ExpenseNo")
        DRemark.Text = dtData.Rows(0).Item("Remark")
        DRate.BackColor = Color.GreenYellow

        DID.Text = id

        BADD.Visible = False


        If DType.SelectedItem.Text = "膳雜費-日當" Then
            DCurrency.SelectedItem.Text = "USD"
            DCurrency.Enabled = False
            DDayMoney.Visible = True
            If wStep = 30 Then
                DMoney.Enabled = True
            Else
                DMoney.Enabled = False
            End If

            DDays.Enabled = True
        Else
            DDayMoney.Visible = False
            DCurrency.Enabled = True
            DMoney.Enabled = True
            DDays.Enabled = True
        End If


    End Sub

    Protected Sub BEDIT_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BEDIT.ServerClick

        Dim Message As String = ""
        '預支款自動代入負數

        Dim money As String = ""
        '預支款自動代入負數
        If DType.SelectedItem.Text = "預支款沖銷" Then
            money = "-" + DMoney.Text
        Else
            money = DMoney.Text
        End If

        If DRate.Text <> "" Then
            Dim rate As Decimal
            rate = Convert.ToDecimal(DRate.Text)
            Dim ARate As Double
            ARate = Math.Round(rate, 4)
            DRate.Text = Str(ARate)
        End If
        Dim SQL As String

        '20231012jessica 增加檢查typeno
        Dim CheckSQL As String = ""
        Dim typeno As String = ""
        If DType.SelectedValue = "" Then
            CheckSQL = " select left(data,2)Data  from M_referp"
            CheckSQL = CheckSQL + " where cat = 3115 "
            CheckSQL = CheckSQL + " and dkey = 'type'"
            CheckSQL = CheckSQL + " and data like '%" + DType.SelectedItem.Text + "%'"
            Dim dtData As DataTable = uDataBase.GetDataTable(CheckSQL)
            If dtData.Rows.Count > 0 Then
                typeno = dtData.Rows(0).Item("Data")
            End If
        Else
            typeno = DType.SelectedValue
        End If



        Try

            If DID.Text <> "" Then
                If InputDataOK(0) Then
                    SQL = " delete  from   F_CloseAccountSheetDTTemp "
                    SQL = SQL & " where unique_id = " & DID.Text & " "
                    uDataBase.ExecuteNonQuery(SQL)

                    SQL = " Insert into F_CloseAccountSheetDTTemp (No,Typeno ,Type,SDate,Pay,Currency,Money,Days,Rate,SumAmt,Remark,ExpenseNo,Createuser,CreateTime,ModifyUser,ModifyTime) "
                    SQL = SQL & "VALUES( "
                    SQL = SQL & " '" & DTNo & "', "
                    SQL = SQL & " '" & typeno & "', "
                    SQL = SQL & " '" & DType.SelectedItem.Text & "', "
                    SQL = SQL & " '" & DSDate.Text & "', "
                    SQL = SQL & " '" & DPay.SelectedValue & "', "
                    SQL = SQL & " '" & DCurrency.SelectedItem.Text & "', "
                    SQL = SQL & " '" & money & "', "
                    SQL = SQL & " '" & DDays.Text & "', "
                    SQL = SQL & " '" & DRate.Text & "', "
                    SQL = SQL & " 0, "
                    SQL = SQL & " N'" & DRemark.Text & "', "
                    SQL = SQL & " '" & DExpenseNo.Text & "', "
                    SQL = SQL & " '" & Request.QueryString("pUserID") & "', "
                    SQL = SQL & " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "', "
                    SQL = SQL & " '" & Request.QueryString("pUserID") & "', "
                    SQL = SQL & " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "') "
                    uDataBase.ExecuteNonQuery(SQL)


                    Message = "資料已修改!"
                End If
            Else
                Message = "未選取資料!"
            End If


        Catch ex As Exception
            Message = "異常,請連絡系統人員!"

        End Try


        uJavaScript.PopMsg(Me, Message)

        '恢復未選取
        DID.Text = ""
        GridView1.SelectedIndex = -1

        'SQL = " select *, convert(decimal(9,0),money*days*rate) as SumAmt from ("
        'SQL = SQL + " select Unique_ID,type,typeno,currency,convert(char(10),sdate,111)sdate,pay,case when money = '' then '0' else  convert(decimal(9,2),money) end as money,"
        'SQL = SQL + "  case when days='' then '0' else convert(decimal(9,0),days) end as days"
        'SQL = SQL + " ,case when pay = '公司卡' then 0 when currency ='TWD' then 1 when Rate ='' then 1  else  convert(decimal(9,6),rate) end as rate,remark,expenseno"
        'SQL = SQL + " from  F_CloseAccountSheetDTTemp "
        'SQL = SQL + "  where 1=1 and No='" + DTNo + "')a "
        'SQL = SQL + " order by  typeno,sdate "



        SQL = " select UNIQUE_ID,TYPE,TYPENO,CURRENCY,SDATE,PAY,MONEY,DAYS,AVERAGERATE AS RATE,REMARK,EXPENSENO,cast(round(money*days*AVERAGERATE ,2) as numeric(10,2)) as SumAmt from ("
        SQL = SQL + " select * from ("
        SQL = SQL + " select Unique_ID,type,typeno,currency,convert(char(10),sdate,111)sdate,pay,case when money = '' then '0' else  convert(decimal(9,2),money) end as money,"
        SQL = SQL + "  case when days='' then '0' else convert(decimal(9,0),days) end as days"
        SQL = SQL + " ,case when pay = '公司卡' then 0 when currency ='TWD' then 1 when Rate ='' then 1  else  convert(decimal(9,6),rate) end as rate,remark,expenseno"
        SQL = SQL + " from  F_CloseAccountSheetDTTemp "
        SQL = SQL + "  where 1=1 and No='" + DTNo + "')a "
        SQL = SQL + " ,("
        SQL = SQL + " select currencycode,syear+'/'+smonth+'/'+substring(ltrim(sday),1,2) as StrD,syear+'/'+smonth+'/'+substring(ltrim(sday),4,2) as EndD,AVERAGERATE  from F_exchangerate"
        SQL = SQL + " )b where a.currency =b.currencycode "
        SQL = SQL + " and '" + AEDate + "' between strD and EndD )a  "
        SQL = SQL + " order by typeno,sdate  "



        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()

        BADD.Visible = True
        SetFieldData()
    End Sub

  
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim h1 As New HyperLink
        Dim sql As String
        If e.Row.RowType = DataControlRowType.DataRow Then
            h1.Text = e.Row.Cells(10).Text
            ' 連結到待處理LIST
            ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep

            If h1.Text <> "&nbsp;" Then
                sql = " select * from F_ExpenseSheet"
                sql = sql + "  where no ='" & DExpenseNo.Text & "'"
                Dim Ddata2 As DataTable = uDataBase.GetDataTable(sql)
                If Ddata2.Rows.Count > 0 Then

                    h1.NavigateUrl = "http://10.245.1.10/N2W/ExpenseSheet_02.aspx?pFormNo=003105&pFormSno=" + Trim(Ddata2.Rows(0).Item("formsno"))

                End If
                h1.Target = "_blank"
                e.Row.Cells(10).Text = ""
                e.Row.Cells(10).Controls.Add(h1)
            End If

         

        End If
    End Sub

    Protected Sub DType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DType.SelectedIndexChanged

        SetFieldData1()

        DSDate.Text = ""
        DPay.Text = ""
        DCurrency.Text = ""
        DMoney.Text = ""
        DDays.Text = ""
        DRate.Text = ""
        DExpenseNo.Text = ""
        DRemark.Text = ""
        DENo.Text = ""



        If Mid(DType.SelectedItem.Text, 1, 3) = "交際費" Then
            DExpenseNo.Enabled = True
            DExpenseNo.BackColor = Color.GreenYellow
            BExpense.Visible = True
            DExpenseNo.ReadOnly = True
        Else
            DExpenseNo.Enabled = False
            DExpenseNo.BackColor = Color.LightGray
            BExpense.Visible = False
        End If

        If DType.SelectedItem.Text = "膳雜費-日當" Then
            DCurrency.SelectedItem.Text = "USD"
            DCurrency.Enabled = False
            DDayMoney.Visible = True
            If wStep = 30 Then
                DMoney.Enabled = True
            Else
                DMoney.Enabled = False
            End If

            DDays.Enabled = True
            DRemark.Text = "膳雜費-日當"
            DPay.SelectedItem.Text = "公司"
        Else
            DDayMoney.Visible = False
            DCurrency.Enabled = True
            DMoney.Enabled = True
            DDays.Enabled = True
        End If

        If DType.SelectedItem.Text = "膳雜費-日當" Or DType.SelectedItem.Text = "住宿費-宿泊" Then

            DDays.Text = Days
            If DType.SelectedItem.Text = "住宿費-宿泊" Then
                If DDays.Text Then
                    DDays.Text = CStr(CInt(Days) - 1)
                End If

            End If

        Else

            DDays.Text = "1"

        End If



    End Sub

    Protected Sub DDayMoney_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDayMoney.SelectedIndexChanged
        Dim Str As String
        Str = DDayMoney.SelectedValue
        Dim I As Integer
        I = Str.IndexOf("-")
     
        DMoney.Text = Mid(Str, I + 2, Len(Str) - 1)

    End Sub


    Sub GetDays()
        Dim Errcode As Integer = 0
        Dim Message As String = ""
        Dim ASDate, AEDate As String
        Dim result As Integer
        ASDate = Request.QueryString("pASDate")
        AEDate = Request.QueryString("pAEDate")

        If ASDate <> "" And AEDate <> "" Then
            Dim d1 As DateTime
            Dim d2 As DateTime
            Dim dTimeSpan As TimeSpan
            d1 = ASDate
            d2 = AEDate
            dTimeSpan = d2.Subtract(d1)

            ' result = DateTime.Compare(d1, d2)  '比日期大小
            result = dTimeSpan.Days
            DDays.Text = dTimeSpan.Days
            If Int(result) < 1 Then
                Errcode = 1
            End If
        Else
            Errcode = 2

        End If

        If Errcode <> 0 Then
            If Errcode = 1 Then Message = "異常：[到達/退房]日期必需大於[出發/入住]日期"
            If Errcode = 2 Then Message = "請先輸入實際[到達/退房]日期及[出發/入住]日期"

            uJavaScript.PopMsg(Me, Message)

            Dim js As String = ""
            js &= "window.close();"
            '
            Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)
        End If

    End Sub

    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData1()


        Dim SQL As String
        Dim i As Integer
      

        '支付
        SQL = "  SELECT   Data   from M_referp  "
        SQL = SQL + " where cat = '3115' and dkey ='Pay' order by Data   "
        Dim dtReferp1 As DataTable = uDataBase.GetDataTable(SQL)
        DPay.Items.Clear()
        DPay.Items.Add("")
        For i = 0 To dtReferp1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp1.Rows(i).Item("Data")
            ListItem1.Value = dtReferp1.Rows(i).Item("Data")
            DPay.Items.Add(ListItem1)
        Next
        dtReferp1.Clear()



        '幣別
        SQL = "  select distinct currencycode as Data1 from F_exchangerate order by  currencycode   "

        Dim dtReferp4 As DataTable = uDataBase.GetDataTable(SQL)
        DCurrency.Items.Clear()
        DCurrency.Items.Add("")

        DCurrency1.Items.Clear()
        DCurrency1.Items.Add("")

        For i = 0 To dtReferp4.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp4.Rows(i).Item("Data1")
            ListItem1.Value = dtReferp4.Rows(i).Item("Data1")
            DCurrency.Items.Add(ListItem1)
            DCurrency1.Items.Add(ListItem1)
        Next
        dtReferp4.Clear()

        '日當
        SQL = "  SELECT  substring(data,4,len(data)-1)Data1,data  from M_referp    "
        SQL = SQL + " where cat = '3115' and dkey ='DayMoney' order by Data   "
        Dim dtReferp2 As DataTable = uDataBase.GetDataTable(SQL)
        DDayMoney.Items.Clear()
        DDayMoney.Items.Add("")

        For i = 0 To dtReferp2.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp2.Rows(i).Item("Data1")
            ListItem1.Value = dtReferp2.Rows(i).Item("Data1")
            DDayMoney.Items.Add(ListItem1)

        Next
        dtReferp2.Clear()


 





    End Sub

    Protected Sub BChange_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BChange.ServerClick

        Dim SQL As String
        SQL = "     UPDATE  F_CloseAccountSheetDTTemp "
        SQL = SQL + " SET RATE = '" + DRate1.Text + "'"
        SQL = SQL + "  where no = '" + DTNo + "'"
        SQL = SQL + "  and currency ='" + Trim(DCurrency1.SelectedItem.Text) + "'"
        uDataBase.ExecuteNonQuery(SQL)


        SQL = " select *, convert(decimal(9,0),money*days*rate) as SumAmt from ("
        SQL = SQL + " select Unique_ID,type,typeno,currency,convert(char(10),sdate,111)sdate,pay,case when money = '' then '0' else  convert(decimal(9,2),money) end as money,"
        SQL = SQL + "  case when days='' then '0' else convert(decimal(9,0),days) end as days"
        SQL = SQL + " ,case when pay = '公司卡' then 0 when currency ='TWD' then 1 when Rate ='' then 1  else  convert(decimal(9,6),rate) end as rate,remark,expenseno"
        SQL = SQL + " from  F_CloseAccountSheetDTTemp "
        SQL = SQL + "  where 1=1 and No='" + DTNo + "')a "
        SQL = SQL + " order by  typeno,sdate "



        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()

        If wStep = 30 Or wStep = 50 Then
            GridView1.AutoGenerateSelectButton = True
        Else
            GridView1.AutoGenerateSelectButton = False
        End If

    End Sub

    Sub GetExpenesNo()
        Dim SQL As String

        If DENo.Text <> "" Then
            SQL = " select no,appname,freason,formno,formsno from F_ExpenseSheet  where sts =1 and no ='" + DENo.Text + "'"
            Dim dtENo As DataTable = uDataBase.GetDataTable(SQL)
            If dtENo.Rows.Count > 0 Then
                DExpenseNo.Text = dtENo.Rows(0).Item("No")
            End If
        End If

    End Sub


    '取匯率
    Sub GetRate()
        Dim SQL As String

        If DPay.SelectedItem.Text = "公司卡" Then
            DRate.Text = 0
        Else

            If DCurrency.SelectedValue <> "" Then
                SQL = " select AVERAGERATE  from ("
                SQL = SQL + "select currencycode,syear+'/'+smonth+'/'+substring(ltrim(sday),1,2) as StrD,syear+'/'+smonth+'/'+substring(ltrim(sday),4,2) as EndD,AVERAGERATE  from F_exchangerate"
                SQL = SQL + " )a where currencycode = '" + DCurrency.SelectedValue + "' and   '" + AEDate + "' between strD and EndD"

                Dim dtRate As DataTable = uDataBase.GetDataTable(SQL)
                If dtRate.Rows.Count > 0 Then
                    DRate.Text = dtRate.Rows(0).Item("AVERAGERATE")
                End If
            End If
        End If


    End Sub

 
End Class
