Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Collections.Generic

Partial Class AddItemListFly
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
            BEDIT.Visible = False
            GetData()
        Else


        End If







        '
    End Sub


    Protected Sub BADD_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BADD.ServerClick

   
        Dim sql As String = ""
        If InputDataOK(0) Then
            sql = " Insert into F_BusinessFlyDTTemp (No,FlyType,Pay, Money1,Currency1,Remark1,SDate,SNo,Createuser,CreateTime,ModifyUser,ModifyTime) "
            sql = sql & "VALUES( "
            sql = sql & " '" & DTNo & "', "
            sql = sql & " '" & DFlyType.SelectedValue & "', "
            sql = sql & " '" & DPay.SelectedValue & "', "
            sql = sql & " '" & DMoney1.Text & "', "
            sql = sql & " '" & DCurrency1.SelectedItem.Text & "', "
            sql = sql & " N'" & DRemark1.Text & "', "
            sql = sql & " '" & DSDate.Text & "', "
            sql = sql & " '" & DSNo.Text & "', "
            sql = sql & " '" & Request.QueryString("pUserID") & "', "
            sql = sql & " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "', "
            sql = sql & " '" & Request.QueryString("pUserID") & "', "
            sql = sql & " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "') "
            uDataBase.ExecuteNonQuery(sql)

        End If

        '


        sql = " Select Unique_ID,No,FlyType,Pay, Money1,Currency1,Remark1,convert(char(10),SDate,112)Sdate,SNo "
        sql = sql + "   from  F_BusinessFlyDTTemp  where 1=1 "
        sql = sql + " and NO ='" + DTNo + "'"


        GridView1.DataSource = uDataBase.GetDataTable(sql)
        GridView1.DataBind()








    End Sub

    Sub GetData()
        Dim SQL As String


        If wStep <> 1 Then
            SQL = " delete from   F_BusinessFlyDTTemp  where no='" + DTNo + "'"
            uDataBase.ExecuteNonQuery(SQL)


            ''存入機票款

            SQL = " Insert into F_BusinessFlyDTTemp(No,FlyType,Pay, Money1,Currency1,Remark1,sdate,sno,Createuser,CreateTime,ModifyUser,ModifyTime) "
            SQL = SQL + " Select No,FlyType,Pay, Money1,Currency1,Remark1,convert(char(10),SDate,112)Sdate,SNo,'" + Request.QueryString("pUserID") + "',"
            SQL = SQL + " getdate(),'" + Request.QueryString("pUserID") + "',getdate()  "
            SQL = SQL + "   from  F_BusinessFlyDT  where 1=1 "
            SQL = SQL + " and NO ='" + DTNo + "'"

            uDataBase.ExecuteNonQuery(SQL)

        End If





        SQL = " Select Unique_ID,No,FlyType,Pay, Money1,Currency1,Remark1,convert(char(10),SDate,112)Sdate,SNo "
        SQL = SQL + "   from  F_BusinessFlyDTTemp  where 1=1 "
        SQL = SQL + " and NO ='" + DTNo + "'"


        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()




    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        e.Row.Cells(1).Visible = False
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Dim SQL As String
        Dim id As String = GridView1.DataKeys(e.RowIndex).Value

        SQL = " delete  from   F_BusinessFlyDTTemp "
        SQL = SQL & " where unique_id = " & id & " "
        uDataBase.ExecuteNonQuery(SQL)
        '

        SQL = " Select Unique_ID,No,FlyType,Pay, Money1,Currency1,Remark1,convert(char(10),SDate,112)Sdate,SNo "
        SQL = SQL + "   from  F_BusinessFlyDTTemp  where 1=1 "
        SQL = SQL + " and NO ='" + DTNo + "'"



        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()



    End Sub


    Protected Sub BClose_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BClose.ServerClick

        Dim Sql As String



        Sql = " delete from  F_BusinessFlyDT where no='" + DTNo + "'"
        uDataBase.ExecuteNonQuery(Sql)


        ''存入機票款
        Sql = " Insert into F_BusinessFlyDT(No,FlyType,Pay, Money1,Currency1,Remark1,sdate,sno,Createuser,CreateTime,ModifyUser,ModifyTime) "
        Sql = Sql + " Select No,FlyType,Pay, Money1,Currency1,Remark1,convert(char(10),SDate,112)Sdate,SNo,'" + Request.QueryString("pUserID") + "',"
        Sql = Sql + " getdate(),'" + Request.QueryString("pUserID") + "',getdate()  "
        Sql = Sql + "   from  F_BusinessFlyDTTemp  where 1=1 "
        Sql = Sql + " and NO ='" + DTNo + "'"
        uDataBase.ExecuteNonQuery(Sql)

        Dim js As String = ""
        js &= "window.opener.document.all.DUpdate.value = '1';"
        ' js &= "window.close();"
        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        '
        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)



    End Sub



    '*****************************************************************
    '**(SetFieldAttribute)
    '**     表單各欄位屬性及欄位輸入檢查等設定
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"


    End Sub

    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData()
        Dim sql As String
        Dim i As Integer


        '機票款類別
        sql = "  SELECT substring(data,4,len(data)-1)Data1,data from M_referp  "
        sql = sql + " where cat = '3114' and dkey ='FlyType' order by Data   "
        Dim dtReferp5 As DataTable = uDataBase.GetDataTable(sql)
        DFlyType.Items.Clear()
        DFlyType.Items.Add("")
        For i = 0 To dtReferp5.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp5.Rows(i).Item("Data1")
            ListItem1.Value = dtReferp5.Rows(i).Item("Data1")
            DFlyType.Items.Add(ListItem1)
        Next
        dtReferp5.Clear()




        '支付
        Sql = "  SELECT   Data   from M_referp  "
        sql = sql + " where cat = '3114' and dkey ='Pay' order by Unique_ID   "
        Dim dtReferp1 As DataTable = uDataBase.GetDataTable(Sql)
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
        Sql = "  SELECT  substring(data,4,len(data)-1)Data1,data  from M_referp    "
        Sql = Sql + " where cat = '3114' and dkey ='Currency' order by Data   "
        Dim dtReferp4 As DataTable = uDataBase.GetDataTable(Sql)
        DCurrency1.Items.Clear()
        DCurrency1.Items.Add("")

        For i = 0 To dtReferp4.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp4.Rows(i).Item("Data1")
            ListItem1.Value = dtReferp4.Rows(i).Item("Data1")
            DCurrency1.Items.Add(ListItem1)

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
            If DFlyType.SelectedValue = "" Then
                ErrCode = 1
            End If
        End If



        If ErrCode = 0 Then
            If DPay.SelectedValue = "" Then
                ErrCode = 2
            End If
        End If


        If ErrCode = 0 Then
            If DCurrency1.SelectedValue = "" Then
                ErrCode = 3
            End If
        End If



        If ErrCode = 0 Then
            If DMoney1.Text = "" Then
                ErrCode = 4
            End If
        End If




 


        '檢查日期是否有輸入
        If ErrCode = 0 Then
            If DSDate.Text = "" Then
                ErrCode = 5
            End If
        End If





   
 

 
        If ErrCode = 1 Then Message = "異常：需輸入類別"
        If ErrCode = 2 Then Message = "異常：需輸入支付"
        If ErrCode = 3 Then Message = "異常：需輸入幣別"
        If ErrCode = 4 Then Message = "異常：需輸入金額"
        If ErrCode = 5 Then Message = "異常：需輸入收據日期"
        If ErrCode = 9 Then Message = "異常：需輸入天數"


        If ErrCode <> 0 Then
            isOK = False
            uJavaScript.PopMsg(Me, Message)

        Else
            isOK = True
        End If

        Return isOK


    End Function
 
    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        Dim SQL As String

        Dim id As String = GridView1.SelectedValue

        SQL = " select* from F_BusinessFlyDTTemp  "
        SQL = SQL & " where unique_id = " & id & " "
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        '
        DFlyType.SelectedValue = dtData.Rows(0).Item("FlyType")
        DPay.SelectedValue = dtData.Rows(0).Item("Pay")
        DCurrency1.SelectedValue = dtData.Rows(0).Item("Currency1")
        DMoney1.Text = Math.Abs(Convert.ToDouble(dtData.Rows(0).Item("Money1"))).ToString
        DSDate.Text = dtData.Rows(0).Item("SDate")
        DSNo.Text = dtData.Rows(0).Item("SNo")
        DRemark1.Text = dtData.Rows(0).Item("Remark1")
        DID.Text = id
        BADD.Visible = False

        BEDIT.Visible = True
    End Sub
 
    Protected Sub BEDIT_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BEDIT.ServerClick
        Dim SQL As String
        Dim Message As String = ""
        Try

            If DID.Text <> "" Then
                If InputDataOK(0) Then
                    SQL = " delete  from   F_BusinessFlyDTTemp "
                    SQL = SQL & " where unique_id = " & DID.Text & " "
                    uDataBase.ExecuteNonQuery(SQL)

                    SQL = " Insert into F_BusinessFlyDTTemp (No,FlyType,Pay, Money1,Currency1,Remark1,SDate,SNo,Createuser,CreateTime,ModifyUser,ModifyTime) "
                    SQL = SQL & "VALUES( "
                    SQL = SQL & " '" & DTNo & "', "
                    SQL = SQL & " '" & DFlyType.SelectedValue & "', "
                    SQL = SQL & " '" & DPay.SelectedValue & "', "
                    SQL = SQL & " '" & DMoney1.Text & "', "
                    SQL = SQL & " '" & DCurrency1.SelectedItem.Text & "', "
                    SQL = SQL & " N'" & DRemark1.Text & "', "
                    SQL = SQL & " '" & DSDate.Text & "', "
                    SQL = SQL & " '" & DSNo.Text & "', "
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


        SQL = " Select Unique_ID,No,FlyType,Pay, Money1,Currency1,Remark1,convert(char(10),SDate,112)Sdate,SNo "
        SQL = SQL + "   from  F_BusinessFlyDTTemp  where 1=1 "
        SQL = SQL + " and NO ='" + DTNo + "'"




        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()

        BADD.Visible = True

        DFlyType.SelectedValue = ""
        DPay.SelectedValue = ""
        DCurrency1.SelectedValue = ""
        DMoney1.Text = ""
        DSDate.Text = ""
        DSNo.Text = ""
        DRemark1.Text = ""
        DID.Text = ""


        SetFieldData()
        BEDIT.Visible = False
    End Sub

End Class
