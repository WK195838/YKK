Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Collections.Generic

Partial Class AddItemListBC
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

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '*****************************************************************

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        NowDateTime = Now.ToString("yyyyMMddHHmmss")
        ' DTNo = NowDateTime
       

        DTNo = Request.QueryString("pDTNo")


        If Not Me.IsPostBack Then   '不是PostBack
            SetFieldData()
            SetFieldAttribute("New")
            GetData()
        Else

        End If





        '
    End Sub


    Protected Sub BADD_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BADD.ServerClick

        If DSDate.Text <> "" And DEDate.Text <> "" Then
            Dim dtone As DateTime = Convert.ToDateTime(DEDate.Text)
            Dim dtwo As DateTime = Convert.ToDateTime(DSDate.Text)
            Dim span As TimeSpan = dtone.Subtract(dtwo)
            DDays.Text = span.Days
        End If

        Dim sql As String = ""
        If InputDataOK(0) Then
            sql = " Insert into F_BusinessTripSheetDTTemp (No,typeno,Type, Appoint,SDate,EDate,STime1,STime2,ETime1,ETime2,Days,Money,Currency,Pay, FlyInf,SFly,EFly,HotelInf, Remark,Createuser,CreateTime,ModifyUser,ModifyTime) "
            sql = sql & "VALUES( "
            sql = sql & " '" & DTNo & "', "
            sql = sql & " '" & DType.SelectedValue & "', "
            sql = sql & " '" & DType.SelectedItem.Text & "', "
            sql = sql & " '" & DAppoint.SelectedValue & "', "
            sql = sql & " '" & DSDate.Text & "', "
            sql = sql & " '" & DEDate.Text & "', "
            sql = sql & " '" & DSTime1.Text & "', "
            sql = sql & " '" & DSTime2.Text & "', "
            sql = sql & " '" & DETime1.Text & "', "
            sql = sql & " '" & DETime2.Text & "', "
            sql = sql & " '" & DDays.Text & "', "
            sql = sql & " '" & DMoney.Text & "', "
            sql = sql & " '" & DCurrency.SelectedValue & "', "
            sql = sql & " '', "
            sql = sql & " N'" & DFlyInf.Text & "', "
            sql = sql & " N'" & DSFly.Text & "', "
            sql = sql & " N'" & DEFly.Text & "', "
            sql = sql & " N'" & DHotelInf.Text & "', "
            sql = sql & " N'" & DRemark.Text & "', "
            sql = sql & " '" & Request.QueryString("pUserID") & "', "
            sql = sql & " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "', "
            sql = sql & " '" & Request.QueryString("pUserID") & "', "
            sql = sql & " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "') "
            uDataBase.ExecuteNonQuery(sql)

        End If

        '
        GridView1.SelectedIndex = -1

        sql = " Select  Unique_ID,No, Type, Appoint,convert(char(10),SDate,111)+' '+stime1+':'+stime2 as SDate,"
        sql = sql + " convert(char(10),EDate,111)+' '+etime1+':'+etime2 as EDate,"
        sql = sql + " Days,Money,Currency, FlyInf,SFly,Efly,HotelInf, Remark  from  F_BusinessTripSheetDTTemp where 1=1 "
        sql = sql + " and NO = '" + DTNo + "'"
        sql = sql + " order by unique_id "


        GridView1.DataSource = uDataBase.GetDataTable(sql)
        GridView1.DataBind()
 




    End Sub

    Sub GetData()
        Dim SQL As String


        If wStep <> 1 Then
            SQL = " delete from  F_BusinessTripSheetdtTemp  where no='" + DTNo + "'"
            uDataBase.ExecuteNonQuery(SQL)

            SQL = " insert into F_BusinessTripSheetdtTemp (No,TypeNo,Type,Appoint,Sdate,Edate,STime1,STime2,ETime1,ETime2,Days,Money,Currency,Pay,FlyInf,SFly,Efly,HotelInf,Remark,CreateUser, CreateTime, ModifyUser, ModifyTime) "
            SQL = SQL + " select no,"
            SQL = SQL + " TypeNo,Type,Appoint,Sdate,Edate,STime1,STime2,ETime1,ETime2,Days,Money,Currency,Pay,FlyInf,SFly,Efly,HotelInf,Remark,'" + Request.QueryString("pUserID") + "',"
            SQL = SQL + " getdate(),'" + Request.QueryString("pUserID") + "',getdate() from  F_BusinessTripSheetDT "
            SQL = SQL + "  where 1=1 and No='" + DTNo + "'"
            uDataBase.ExecuteNonQuery(SQL)

        End If







        SQL = " Select  Unique_ID,No, Type, Appoint,convert(char(10),SDate,111)+' '+stime1+':'+stime2 as SDate,"
        SQL = SQL + " convert(char(10),EDate,111)+' '+etime1+':'+etime2 as EDate,"
        SQL = SQL + " Days,Money,Currency, FlyInf,SFly,Efly,HotelInf,Remark  from  F_BusinessTripSheetDTTemp where 1=1 "
        SQL = SQL + " and NO = '" + DTNo + "'"
        SQL = SQL + " order by unique_id "


        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()

       



    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        e.Row.Cells(1).Visible = False
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Dim SQL As String
 
        Dim id As String = GridView1.DataKeys(e.RowIndex).Value

        SQL = " delete  from   F_BusinessTripSheetDTTemp "
        SQL = SQL & " where unique_id = " & id & " "
        uDataBase.ExecuteNonQuery(SQL)
        '
        SQL = " Select  Unique_ID,No, Type, Appoint,convert(char(10),SDate,111)+' '+stime1+':'+stime2 as SDate,"
        SQL = SQL + " convert(char(10),EDate,111)+' '+etime1+':'+etime2 as EDate,"
        SQL = SQL + " Days,Money,Currency, FlyInf,SFly,Efly,HotelInf, Remark  from  F_BusinessTripSheetDTTemp where 1=1 "
        SQL = SQL + " and NO = '" + DTNo + "'"
        SQL = SQL + " order by unique_id "


        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()

    End Sub


    Protected Sub BClose_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BClose.ServerClick

        Dim sql As String

        '先檢查是否有KEY 住宿

        sql = sql + " select No,"
        sql = sql + " TypeNo,Type,Appoint,Sdate,Edate,STime1,STime2,ETime1,ETime2,Days,Money,Currency,Pay,FlyInf,SFly,Efly,HotelInf,Remark,'" + Request.QueryString("pUserID") + "',"
        sql = sql + " getdate(),'" + Request.QueryString("pUserID") + "',getdate() from  F_BusinessTripSheetDTTemp"
        sql = sql + "  where 1=1 and No='" + DTNo + "' and type = '住宿' "
        Dim dtData As DataTable = uDataBase.GetDataTable(sql)
        If dtData.Rows.Count > 0 Then
            sql = " delete from  F_BusinessTripSheetdt  where no='" + DTNo + "'"
            uDataBase.ExecuteNonQuery(sql)

            sql = " insert into F_BusinessTripSheetdt (No,TypeNo,Type,Appoint,Sdate,Edate,STime1,STime2,ETime1,ETime2,Days,Money,Currency,Pay,FlyInf,SFly,Efly,HotelInf,Remark,CreateUser, CreateTime, ModifyUser, ModifyTime) "
            sql = sql + " select No,"
            sql = sql + " TypeNo,Type,Appoint,Sdate,Edate,STime1,STime2,ETime1,ETime2,Days,Money,Currency,Pay,FlyInf,SFly,Efly,HotelInf,Remark,'" + Request.QueryString("pUserID") + "',"
            sql = sql + " getdate(),'" + Request.QueryString("pUserID") + "',getdate() from  F_BusinessTripSheetDTTemp"
            sql = sql + "  where 1=1 and No='" + DTNo + "'"
            uDataBase.ExecuteNonQuery(sql)



            Dim js As String = ""
            js &= "window.opener.document.all.DUpdate.value = '1';"
            ' js &= "window.close();"
            js &= "window.opener.document.forms[0].submit();"
            js &= "window.close();"


            '
            Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

        Else
            uJavaScript.PopMsg(Me, "請輸入住宿資訊！若不確定飯店資訊請輸入無")

        End If


      

    End Sub



    '*****************************************************************
    '**(SetFieldAttribute)
    '**     表單各欄位屬性及欄位輸入檢查等設定
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DEDate');"
       
    End Sub

    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData()

        '類別
        Dim SQL As String
        Dim i As Integer
        SQL = "  SELECT  substring(data,1,1)seqno, substring(data,3,len(data)-1)Data1   from M_referp  "
        SQL = SQL + " where cat = '3114' and dkey ='Type' order by  data    "

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

        '預約
        SQL = "  SELECT   Data   from M_referp  "
        SQL = SQL + " where cat = '3114' and dkey ='Appoint' order by Data   "
        Dim dtReferp1 As DataTable = uDataBase.GetDataTable(SQL)
        DAppoint.Items.Clear()
        DAppoint.Items.Add("")
        For i = 0 To dtReferp1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp1.Rows(i).Item("Data")
            ListItem1.Value = dtReferp1.Rows(i).Item("Data")
            DAppoint.Items.Add(ListItem1)
        Next
        dtReferp1.Clear()

        '時
        SQL = "  SELECT   Data   from M_referp  "
        SQL = SQL + " where cat = '3114' and dkey ='Hour' order by Data   "
        Dim dtReferp2 As DataTable = uDataBase.GetDataTable(SQL)
        DSTime1.Items.Clear()
        DSTime1.Items.Add("00")
        DETime1.Items.Clear()
        DETime1.Items.Add("00")
        For i = 0 To dtReferp2.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp2.Rows(i).Item("Data")
            ListItem1.Value = dtReferp2.Rows(i).Item("Data")
            DSTime1.Items.Add(ListItem1)
            DETime1.Items.Add(ListItem1)
        Next
        dtReferp2.Clear()

        '分
        SQL = "  SELECT   Data   from M_referp  "
        SQL = SQL + " where cat = '3114' and dkey ='Min' order by Data   "
        Dim dtReferp3 As DataTable = uDataBase.GetDataTable(SQL)
        DSTime2.Items.Clear()
        DSTime2.Items.Add("00")
        DETime2.Items.Clear()
        DETime2.Items.Add("00")
        For i = 0 To dtReferp3.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp3.Rows(i).Item("Data")
            ListItem1.Value = dtReferp3.Rows(i).Item("Data")
            DSTime2.Items.Add(ListItem1)
            DETime2.Items.Add(ListItem1)
        Next
        dtReferp3.Clear()



        '幣別
        SQL = "  select distinct currencycode as Data1 from F_exchangerate order by  currencycode"

        Dim dtReferp4 As DataTable = uDataBase.GetDataTable(SQL)
        DCurrency.Items.Clear()
        DCurrency.Items.Add("")

        For i = 0 To dtReferp4.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp4.Rows(i).Item("Data1")
            ListItem1.Value = dtReferp4.Rows(i).Item("Data1")
            DCurrency.Items.Add(ListItem1)

        Next
        dtReferp4.Clear()





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
            If DType.SelectedValue = "" Then
                ErrCode = 4
            End If
        End If

        If ErrCode = 0 Then
            If DAppoint.SelectedValue = "" Then
                ErrCode = 5
            End If
        End If


        '檢查日期是否有輸入
        If ErrCode = 0 Then
            If DSDate.Text = "" Then
                ErrCode = 2
            End If
        End If


        If ErrCode = 0 Then
            If DEDate.Text = "" Then
                ErrCode = 3
            End If
        End If



        'If ErrCode = 0 Then
        '    If DDays.Text = "" Then
        '        ErrCode = 9
        '    End If
        'End If

        If DType.SelectedItem.Text = "住宿" Then
            If ErrCode = 0 Then
                If DCurrency.SelectedValue = "" Then
                    ErrCode = 6
                End If
            End If

            If ErrCode = 0 Then
                If DMoney.Text = "" Then
                    ErrCode = 7
                End If
            End If


        End If

        If DType.SelectedItem.Text = "航班" Then
            '航班
            If ErrCode = 0 Then
                If DFlyInf.Text = "" Then
                    ErrCode = 8
                End If
            End If

            '起點
            If ErrCode = 0 Then
                If DSFly.Text = "" Then
                    ErrCode = 10
                End If
            End If

            '迄點
            If ErrCode = 0 Then
                If DEFly.Text = "" Then
                    ErrCode = 11
                End If
            End If

        Else

            If ErrCode = 0 Then
                If DHotelInf.Text = "" Then
                    ErrCode = 12
                End If
            End If

        End If

 









        '檢查日期[到達/退房]日期必需大於[出發/入住]日期
        If ErrCode = 0 Then

            Dim d1 As DateTime
            Dim d2 As DateTime
            Dim result As Integer
            d1 = DSDate.Text
            d2 = DEDate.Text()
            result = DateTime.Compare(d1, d2)  '比日期大小
            If result = 1 Then
                ErrCode = 1
            End If

        End If


        If ErrCode = 1 Then Message = "異常：[到達/退房]日期必需大於[出發/入住]日期"
        If ErrCode = 2 Then Message = "異常：需輸入[出發/入住]日期"
        If ErrCode = 3 Then Message = "異常：需輸入[到達/退房]日期"
        If ErrCode = 4 Then Message = "異常：需輸入類別"
        If ErrCode = 5 Then Message = "異常：需輸入預約"
        If ErrCode = 6 Then Message = "異常：需輸入幣別"
        If ErrCode = 7 Then Message = "異常：需輸入金額"
        If ErrCode = 8 Then Message = "異常：需輸入航班資訊"
        If ErrCode = 9 Then Message = "異常：需輸入天數"
        If ErrCode = 10 Then Message = "異常：需輸入起點"
        If ErrCode = 11 Then Message = "異常：需輸入迄點"
        If ErrCode = 12 Then Message = "異常：需輸入飯店資訊"
 

        If ErrCode <> 0 Then
            isOK = False
            uJavaScript.PopMsg(Me, Message)

        Else
            isOK = True
        End If

        Return isOK


    End Function



    Protected Sub DEDate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub DType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DType.SelectedIndexChanged
        DCurrency.Text = ""
        DMoney.Text = ""
        DDays.Text = ""
        DFlyInf.Text = ""
        DSFly.Text = ""
        DEFly.Text = ""
        DHotelInf.Text = ""

        If DType.SelectedItem.Text = "航班" Then
            DCurrency.Enabled = False
            DMoney.Enabled = False
            DCurrency.BackColor = Color.LightGray
            DMoney.BackColor = Color.LightGray
            DFlyInf.Visible = True
            DSFly.Visible = True
            DEFly.Visible = True
            DHotelInf.Visible = False

        Else
            DCurrency.Enabled = True
            DMoney.Enabled = True
            DCurrency.BackColor = Color.GreenYellow
            DMoney.BackColor = Color.GreenYellow
            DHotelInf.Visible = True
            DFlyInf.Visible = False
            DSFly.Visible = False
            DEFly.Visible = False
        End If
    End Sub
End Class
