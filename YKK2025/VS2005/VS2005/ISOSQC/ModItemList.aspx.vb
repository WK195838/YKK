Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Collections.Generic

Partial Class ModItemList
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
    Dim DTNo, DNo As String
    Dim Days As String
    Dim DTTable As String
    Dim DTTempTable As String

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '*****************************************************************

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Errcode As Integer = 0
        Dim Message As String = ""



        BApproveNo.Attributes("onclick") = "GetApproveNo();"

        BOldToNewNo.Attributes("onclick") = "GetOldToNewNo();"

        'DTNo = NowDateTime
        ' DTNo = "23030908302001"

        Days = Request.QueryString("pDays")
        wStep = Request.QueryString("pwStep")
        DTNo = Request.QueryString("pDTNo")
        DNo = Request.QueryString("pDNo")
        wFormNo = Request.QueryString("pwFormNo")

        If wFormNo = "008002" Then
            DTTable = "F_QASheetDT"
            DTTempTable = "F_QASheetDTTemp"
        Else
            DTTable = "F_QAModSheetDT"
            DTTempTable = "F_QAModSheetDTTemp"
        End If


        BEDIT.Visible = False

        If wStep = 20 Then
            DResult1.Visible = True
            DResult2.Visible = True
            DResult1.BackColor = Color.Yellow
            DElement.Visible = True
            DRemark.Visible = True
            GridView1.AutoGenerateSelectButton = True
        Else
            DResult1.Visible = False
            DResult2.Visible = False
            DResult1.BackColor = Color.GreenYellow
            DElement.Visible = False
            DRemark.Visible = False
            GridView1.AutoGenerateSelectButton = False

        End If

        If Not Me.IsPostBack Then   '不是PostBack


            SetFieldData()

            GetData()
        Else


        End If






        '
    End Sub


    Protected Sub BADD_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BADD.ServerClick

        Dim SQL As String
        '先檢查Seqno 有沒有重覆
        SQL = " select * from  " + DTTempTable + " where no ='" + DNo + "'"
        SQL = SQL + " and seqno = '" + DSeqNo.Text.ToUpper + "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count = 0 Then
            If InputDataOK(0) Then
                SQL = " Insert into " + DTTempTable + " (No,ModNo,SeqNo,Type,Supplier,Code,Size,Family,Body,Puller,COLOR,Finish,ApproveNo,OldToNewNo,OldToNew,Remark,Element,Result1,Result2,Createuser,CreateTime,ModifyUser,ModifyTime) "
                SQL = SQL & "VALUES( "
                SQL = SQL & " '" & DTNo & "', "
                SQL = SQL & " '" & DNo & "', "
                SQL = SQL & " N'" & DSeqNo.Text.ToUpper & "', "
                SQL = SQL & " '" & DType.SelectedItem.Text & "', "
                SQL = SQL & " '" & DSupplier.SelectedItem.Text & "', "
                SQL = SQL & " N'" & DCode.Text & "', "
                SQL = SQL & " N'" & DSize.SelectedItem.Text & "', "
                SQL = SQL & " N'" & Trim(DFamily.Text.ToUpper) & "', "
                SQL = SQL & " N'" & Trim(DBODY.Text.ToUpper) & "', "
                SQL = SQL & " N'" & Trim(DPuller.Text.ToUpper) & "', "
                SQL = SQL & " N'" & Trim(DColor.Text.ToUpper) & "', "
                SQL = SQL & " N'" & Trim(DFINISH.Text.ToUpper) & "', "
                SQL = SQL & " N'" & Trim(DApproveNo.Text.ToUpper) & "', "
                SQL = SQL & " N'" & Trim(DOldToNewNo.Text.ToUpper) & "', "
                If DOldToNew.Checked = True Then
                    SQL = SQL & " 'Yes', "
                Else
                    SQL = SQL & " 'No', "
                End If
                SQL = SQL & " N'" & Trim(DRemark.Text) & "', "
                SQL = SQL & " N'" & Trim(DElement.Text) & "', "
                SQL = SQL & " '" & DResult1.SelectedItem.Text & "', "
                SQL = SQL & " '" & DResult2.SelectedItem.Text & "', "
                SQL = SQL & " '" & Request.QueryString("pUserID") & "', "
                SQL = SQL & " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "', "
                SQL = SQL & " '" & Request.QueryString("pUserID") & "', "
                SQL = SQL & " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "') "
                uDataBase.ExecuteNonQuery(SQL)
            End If



            SQL = " Select  *   FROM " + DTTempTable + "  where 1=1 "
            SQL = SQL + " and ModNO = '" + DNo + "'"
            SQL = SQL + " order by unique_id "


            GridView1.DataSource = uDataBase.GetDataTable(SQL)
            GridView1.DataBind()

        Else
            uJavaScript.PopMsg(Me, "No重覆")
        End If



    End Sub

    Sub GetData()
        Dim SQL As String


        SQL = " select * from " + DTTempTable
        SQL = SQL + "  where 1=1 and ModNo='" + DNo + "' AND No='" + DTNo + "' "
        SQL = SQL + " order by seqno "

        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()



        'If wStep <> 1 Then
        '    SQL = " delete from  " + DTTempTable + "  where no='" + DNo + "'"
        '    uDataBase.ExecuteNonQuery(SQL)

        '    SQL = " Insert into " + DTTempTable + " (No,SeqNo,Type,Supplier,Code,Size,Family,Body,Puller,COLOR,Finish,ApproveNo,OldToNewNo,OldToNew,Remark,Result1,Result2,Createuser,CreateTime,ModifyUser,ModifyTime) "
        '    SQL = SQL + " select no,SeqNo,Type,Supplier,Code,Size,Family,Body,Puller,COLOR,Finish,ApproveNo,OldToNewNo,OldToNew,Remark,Result1,Result2,"
        '    SQL = SQL + " '" + Request.QueryString("pUserID") + "',"
        '    SQL = SQL + " getdate(),'" + Request.QueryString("pUserID") + "',getdate() "
        '    SQL = SQL + " from  " + DTTable
        '    SQL = SQL + "  where 1=1 and No='" + DNo + "' "
        '    SQL = SQL + " order by SeqNo "
        '    uDataBase.ExecuteNonQuery(SQL)


        'End If







        SQL = " Select  *   FROM " + DTTempTable + "  where 1=1 "
        SQL = SQL + " and ModNO = '" + DNo + "' AND No='" + DTNo + "' "
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

        SQL = " delete  from   " + DTTempTable
        SQL = SQL & " where unique_id = " & id & " "
        uDataBase.ExecuteNonQuery(SQL)
        '
        DID.Text = ""

        GetData()
    End Sub


    Protected Sub BClose_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BClose.ServerClick
        Dim sql As String

        'sql = " delete from  " + DTTable + " where Modno='" + DNo + "'"
        'uDataBase.ExecuteNonQuery(sql)

        sql = " Insert into " + DTTable + " (No,Modno,SeqNo,Type,Supplier,Code,Size,Family,Body,Puller,COLOR,Finish,ApproveNo,OldToNewNo,OldToNew,Remark,Result1,Result2,Createuser,CreateTime,ModifyUser,ModifyTime) "
        sql = sql + " select '" + DTNo + "',Modno,SeqNo,Type,Supplier,Code,Size,Family,Body,Puller,COLOR,Finish,ApproveNo,OldToNewNo,OldToNew,Remark,Result1,Result2,"
        sql = sql + " '" + Request.QueryString("pUserID") + "',"
        sql = sql + " getdate(),'" + Request.QueryString("pUserID") + "',getdate() "
        sql = sql + " from  " + DTTempTable
        sql = sql + "  where 1=1 and Modno='" + DNo + "' AND No='" + DTNo + "' "
        sql = sql + " order by SeqNo "
        uDataBase.ExecuteNonQuery(sql)



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



    Function InputDataOK(ByVal pAction As Integer) As Boolean

        Dim Message As String = ""
        Dim isOK As Boolean = False
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False



        ''        If DPay.SelectedValue <> "公司卡" Then
        'If ErrCode = 0 Then
        '    If DResult1.SelectedValue = "" Then
        '        ErrCode = 7
        '    End If
        'End If


        'If ErrCode = 0 Then
        '    If DOldToNew.Text = "" Then
        '        ErrCode = 5
        '    End If
        'End If

        'If ErrCode = 0 Then
        '    If DSize.Text = "" Then
        '        ErrCode = 9
        '    End If
        'End If


        'If ErrCode = 0 Then
        '    If DCurrency.SelectedItem.Text = "" Then
        '        ErrCode = 6
        '    End If
        'End If


        'If DType.SelectedItem.Text = "膳雜費-日當" Then
        '    If DDayMoney.SelectedValue = "" Then
        '        ErrCode = 10
        '    End If
        'End If


        'If ErrCode = 0 Then
        '    If DRemark.Text = "" Then
        '        ErrCode = 11
        '    End If
        'End If



        '       End If
        '




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

        SQL = " select * from  " + DTTempTable
        SQL = SQL & " where unique_id = " & id & " "
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        ''
        DSeqNo.Text = dtData.Rows(0).Item("seqno")
        DType.SelectedItem.Text = dtData.Rows(0).Item("Type")
        DSupplier.SelectedValue = dtData.Rows(0).Item("Supplier")
        DCode.Text = dtData.Rows(0).Item("Code")
        DSize.SelectedValue = dtData.Rows(0).Item("Size")
        DFamily.Text = dtData.Rows(0).Item("Family")
        DFINISH.Text = dtData.Rows(0).Item("Body")
        DColor.Text = dtData.Rows(0).Item("Color")
        DBODY.Text = dtData.Rows(0).Item("Finish")
        DApproveNo.Text = dtData.Rows(0).Item("ApproveNo")
        DOldToNewNo.Text = dtData.Rows(0).Item("OldToNewNo")
        If dtData.Rows(0).Item("OldToNew") = "Yes" Then
            DOldToNew.Checked = True
        Else
            DOldToNew.Checked = False
        End If
        DRemark.Text = dtData.Rows(0).Item("Remark")
        DElement.Text = dtData.Rows(0).Item("Remark")


        DResult1.SelectedValue = dtData.Rows(0).Item("Result1")
        DResult2.SelectedValue = dtData.Rows(0).Item("Result2")

        BADD.Visible = False
        DID.Text = id

    End Sub

    Protected Sub BEDIT_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BEDIT.ServerClick

        Dim Message As String = ""
        ''預支款自動代入負數


        Dim SQL As String
        Try

            If DID.Text <> "" Then
                If InputDataOK(0) Then
                    SQL = " delete  from   + DTTempTable"
                    SQL = SQL & " where unique_id = " & DID.Text & " "
                    uDataBase.ExecuteNonQuery(SQL)

                    SQL = " Insert into " + DTTempTable + " (No,SeqNo,Type,Supplier,Code,Size,Family,Body,Puller,COLOR,Finish,ApproveNo,OldToNewNo,OldToNew,Remark,Element,Result1,Result2,Createuser,CreateTime,ModifyUser,ModifyTime) "
                    SQL = SQL & "VALUES( "
                    SQL = SQL & " '" & DNo & "', "
                    SQL = SQL & " N'" & DSeqNo.Text.ToUpper & "', "
                    SQL = SQL & " '" & DType.SelectedItem.Text & "', "
                    SQL = SQL & " '" & DSupplier.SelectedItem.Text & "', "
                    SQL = SQL & " N'" & DCode.Text & "', "
                    SQL = SQL & " N'" & DSize.SelectedItem.Text & "', "
                    SQL = SQL & " N'" & DFamily.Text & "', "
                    SQL = SQL & " N'" & DFINISH.Text & "', "
                    SQL = SQL & " N'" & Trim(DPuller.Text.ToUpper) & "', "
                    SQL = SQL & " N'" & Trim(DColor.Text.ToUpper) & "', "
                    SQL = SQL & " N'" & Trim(DBODY.Text.ToUpper) & "', "
                    SQL = SQL & " N'" & Trim(DApproveNo.Text.ToUpper) & "', "
                    SQL = SQL & " N'" & Trim(DOldToNewNo.Text.ToUpper) & "', "
                    If DOldToNew.Checked = True Then
                        SQL = SQL & " 'Yes', "
                    Else
                        SQL = SQL & " 'No', "
                    End If
                    SQL = SQL & " N'" & Trim(DRemark.Text) & "', "
                    SQL = SQL & " N'" & Trim(DElement.Text) & "', "
                    SQL = SQL & " '" & DResult1.SelectedItem.Text & "', "
                    SQL = SQL & " '" & DResult2.SelectedItem.Text & "', "
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

            GridView1.SelectedIndex = -1
            GetData()

        Catch ex As Exception
            Message = "異常,請連絡系統人員!"

        End Try

        BADD.Visible = True
        SetFieldData()

    End Sub

    Protected Sub DPay_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DResult1.SelectedIndexChanged
        If DResult1.SelectedValue = "公司卡" Then
            '  DCurrency.Enabled = False
            '  DDays.Enabled = False
            ' DMoney.Enabled = False

            ' DCurrency.BackColor = Color.LightGray
            ' DDays.BackColor = Color.LightGray
            'DMoney.BackColor = Color.LightGray
            'DRate.BackColor = Color.LightGray
        Else
            'DCurrency.Enabled = True
            'DDays.Enabled = True
            'DMoney.Enabled = True

            'DCurrency.BackColor = Color.GreenYellow
            'DDays.BackColor = Color.GreenYellow
            'DMoney.BackColor = Color.GreenYellow
            ' DRate.BackColor = Color.GreenYellow
        End If
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        'Dim h1 As New HyperLink
        'Dim sql As String
        'If e.Row.RowType = DataControlRowType.DataRow Then
        '    h1.Text = e.Row.Cells(10).Text
        '    ' 連結到待處理LIST
        '    ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep

        '    If h1.Text <> "&nbsp;" Then
        '        sql = " select * from F_ExpenseSheet"
        '        sql = sql + "  where no ='" & DExpenseNo.Text & "'"
        '        Dim Ddata2 As DataTable = uDataBase.GetDataTable(sql)
        '        If Ddata2.Rows.Count > 0 Then

        '            h1.NavigateUrl = "http://10.245.1.10/N2W/ExpenseSheet_02.aspx?pFormNo=003105&pFormSno=" + Trim(Ddata2.Rows(0).Item("formsno"))

        '        End If
        '        h1.Target = "_blank"
        '        e.Row.Cells(10).Text = ""
        '        e.Row.Cells(10).Controls.Add(h1)
        '    End If



        'End If
    End Sub





    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData()


        Dim SQL As String
        Dim i As Integer


        '類別
        SQL = "  SELECT   Data   from M_referp  "
        SQL = SQL + " where cat = '8002' and dkey ='Type' order by unique_id    "
        Dim dtReferp1 As DataTable = uDataBase.GetDataTable(SQL)
        DType.Items.Clear()
        DType.Items.Add("")
        For i = 0 To dtReferp1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp1.Rows(i).Item("Data")
            ListItem1.Value = dtReferp1.Rows(i).Item("Data")
            DType.Items.Add(ListItem1)
        Next

        DType.Items(3).Selected = True
        dtReferp1.Clear()




        '廠商
        SQL = "  SELECT  Data from M_referp    "
        SQL = SQL + " where cat = '8002' and dkey ='Supplier' order by unique_id    "
        Dim dtReferp4 As DataTable = uDataBase.GetDataTable(SQL)
        DSupplier.Items.Clear()
        DSupplier.Items.Add("")


        For i = 0 To dtReferp4.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp4.Rows(i).Item("Data")
            ListItem1.Value = dtReferp4.Rows(i).Item("Data")
            DSupplier.Items.Add(ListItem1)

        Next
        dtReferp4.Clear()



        '判定
        SQL = "  SELECT Data  from M_referp    "
        SQL = SQL + " where cat = '8002' and dkey ='Result1' order by unique_id   "
        Dim dtReferp2 As DataTable = uDataBase.GetDataTable(SQL)
        DResult1.Items.Clear()
        DResult1.Items.Add("")

        For i = 0 To dtReferp2.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp2.Rows(i).Item("Data")
            ListItem1.Value = dtReferp2.Rows(i).Item("Data")
            DResult1.Items.Add(ListItem1)

        Next
        dtReferp2.Clear()




        '判定
        SQL = "  SELECT Data  from M_referp    "
        SQL = SQL + " where cat = '8002' and dkey ='Result2' order by unique_id   "
        Dim dtReferp3 As DataTable = uDataBase.GetDataTable(SQL)
        DResult2.Items.Clear()
        DResult2.Items.Add("")

        For i = 0 To dtReferp3.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp3.Rows(i).Item("Data")
            ListItem1.Value = dtReferp3.Rows(i).Item("Data")
            DResult2.Items.Add(ListItem1)

        Next
        dtReferp3.Clear()





        'size
        SQL = "  select distinct size data from mst_item "
        SQL = SQL + "    order by size "
        Dim dtReferp5 As DataTable = uDataBase.GetDataTable(SQL)
        DSize.Items.Clear()
        DSize.Items.Add("")

        For i = 0 To dtReferp5.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp5.Rows(i).Item("Data")
            ListItem1.Value = dtReferp5.Rows(i).Item("Data")
            DSize.Items.Add(ListItem1)

        Next
        dtReferp5.Clear()







    End Sub


End Class
