Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DTMW_NewColorSheet02_LULU
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim FieldName(60) As String     '各欄位
    Dim Attribute(60) As Integer    '各欄位屬性    
    Dim Top As String               '預設的控制項位置
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wbFormSno As Integer        '連續起單-表單流水號
    Dim wbStep As Integer           '連續起單-工程代碼
    Dim wStep As Integer            '工程代碼
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者ID
    Dim wUserID As String          '使用者ID
    Dim wAgentID As String          '被代理人ID
    Dim NowDateTime As String       '現在日期時間
    Dim wNextGate As String         '下一關人
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    'Modify-Start by Joy 2009/11/20(2010行事曆對應)
    'Dim wDepo As String = "CL"      '台北行事曆(CL->中壢, TP->台北)
    '
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    'Modify-End

    Dim wUserName As String = ""    '姓名代理用
    Dim HolidayList As New List(Of Integer) '用以記錄假日的欄位索引值
    Dim indexList As New List(Of Integer)   '用以記錄不屬於選取月份的欄位索引值
    Dim DateList As New List(Of String)     '用以記錄所選取的一周日期

    ''' <summary>
    ''' 以下為共用函式的宣告
    ''' </summary>
    ''' <remarks></remarks>
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    Dim aAgentID As String
    Dim Makemap1, Makemap2, Makemap3, Makemap4, Makemap5, Makemap6 As Integer
    Dim AQty, DevNo, Manufout As String
    Dim isOK As Boolean = True
    Dim Message As String = ""
    Dim LastStep As String






    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '設定共用參數


        If Not Me.IsPostBack Then   '不是PostBack

            ShowFormData()      '顯示表單資料

        End If


    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日時
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號


        LSheetName.Text = "拉鏈ZIPPER CHAIN (LULU)" '表單名稱

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub UpdateTranFile()
        'Modify-Start by Joy 2009/11/20(2010行事曆對應)
        'oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDepo,wUserID)
        '
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, wUserID)
        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者
        'Modify-End
    End Sub

    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()


        Dim SQL As String

        SQL = "Select * From F_NewColorLULU "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then


            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            '   DBarCode.ImageUrl = "imga.aspx?mycode=" & DNo.Text  BARCODE
            DDate.Text = DBAdapter1.Rows(0).Item("Date")                         'No 
            DDepoName.Text = DBAdapter1.Rows(0).Item("DepoName")                         'No
            DName.Text = DBAdapter1.Rows(0).Item("Name")                         'No
            DCustomer.Text = DBAdapter1.Rows(0).Item("Customer")                         'No
            DCustomerCode.Text = DBAdapter1.Rows(0).Item("CustomerCode")                         'No
            DBuyer.Text = DBAdapter1.Rows(0).Item("Buyer")                         'No
            DBuyerCode.Text = DBAdapter1.Rows(0).Item("BuyerCode")
            DCustomerColor.Text = DBAdapter1.Rows(0).Item("CustomerColor")                         'No3
            DCustomerColorCode.Text = DBAdapter1.Rows(0).Item("CustomerColorCode")
            DOverseaYKKCode.Text = DBAdapter1.Rows(0).Item("OverseaYKKCode")

            DPANTONECode.Text = DBAdapter1.Rows(0).Item("PANTONECode")
            SetFieldData("DevYear", DBAdapter1.Rows(0).Item("DevYear"))    '年
            SetFieldData("DevSeason", DBAdapter1.Rows(0).Item("DevSeason"))    '季

            If Mid(DBAdapter1.Rows(0).Item("ReceiveDate").ToString, 1, 4) = "1900" Then
                DReceiveDate.Text = ""
            Else
                DReceiveDate.Text = DBAdapter1.Rows(0).Item("ReceiveDate")               '下單時間
            End If

            SetFieldData("ColorLight1", DBAdapter1.Rows(0).Item("ColorLight1"))    '類別1
            SetFieldData("ColorLight2", DBAdapter1.Rows(0).Item("ColorLight2"))    '類別2
            SetFieldData("ColorLight3", DBAdapter1.Rows(0).Item("ColorLight3"))    '類別2
            SetFieldData("NewOldColor", DBAdapter1.Rows(0).Item("NewOldColor"))    '類別2

            If Mid(DBAdapter1.Rows(0).Item("DeliveryDate").ToString, 1, 4) = "1900" Then
                DDeliveryDate.Value = ""
            Else
                DDeliveryDate.Value = DBAdapter1.Rows(0).Item("DeliveryDate")               '下單時間
            End If

            If wStep = 500 Then  '(NG再送出希望交期需重新選擇)
                DDeliveryDate.Value = ""
            End If


            DDuplicateNo.Text = DBAdapter1.Rows(0).Item("DuplicateNo")
            If DBAdapter1.Rows(0).Item("NOCCS") = 1 Then
                DNOCCS.Checked = True
            End If
            DNOCCSReason.Text = DBAdapter1.Rows(0).Item("NOCCSReason")
            DCustomerNGColor.Text = DBAdapter1.Rows(0).Item("CustomerNGColor")
            DRemark.Text = DBAdapter1.Rows(0).Item("Remark")
            SetFieldData("CustomerSample", DBAdapter1.Rows(0).Item("CustomerSample"))    '類別2

            DColorSystem.Text = DBAdapter1.Rows(0).Item("ColorSystem")

            DCheckNo.Text = DBAdapter1.Rows(0).Item("CheckNo")
            DPFBWire.Value = DBAdapter1.Rows(0).Item("PFBWire")
            DPFOpenParts.Value = DBAdapter1.Rows(0).Item("PFOpenParts")

            SetFieldData("YKKColorType", DBAdapter1.Rows(0).Item("YKKColorType"))    '類別2

            DYKKColorCode.Value = DBAdapter1.Rows(0).Item("YKKColorCode")
            DYKKColorCodeVF.Value = DBAdapter1.Rows(0).Item("YKKColorCodeVF")
            DYKKColorCodeSLD.Value = DBAdapter1.Rows(0).Item("YKKColorCodeSLD")


            Dim Version As Integer

            '計算是第幾版
            SQL = " select  count(*)cun from  t_waithandle"
            SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & " and  ( "
            SQL = SQL & "  ( step in (525,25)  and sts =3)"
            SQL = SQL & " or "
            SQL = SQL & "  ( step in (20,520)  and sts =1)"
            SQL = SQL & " )"

            Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
            If DBAdapter3.Rows.Count > 0 Then
                If (wStep = 520 Or wStep = 525) Then
                    Version = DBAdapter3.Rows(0).Item("cun") + 1
                Else
                    Version = DBAdapter3.Rows(0).Item("cun")
                End If

            Else
                Version = 1
            End If

            DVersion.Text = Version
            DSLDColor.Text = DBAdapter1.Rows(0).Item("SLDColor")

            SetFieldData("ColorType", DBAdapter1.Rows(0).Item("Again"))    '類別2
            If DBAdapter1.Rows(0).Item("again") = 1 Then
                DAgain.SelectedValue = "淡色"
            ElseIf DBAdapter1.Rows(0).Item("again") = 2 Then
                DAgain.SelectedValue = "濃色"
            End If

            DVFColor.Text = DBAdapter1.Rows(0).Item("VFColor")
            DVFColor1.Text = DBAdapter1.Rows(0).Item("VFColor1")

            If DBAdapter1.Rows(0).Item("Complete") = 1 Then       '有新色依賴完成表
                LComplete.NavigateUrl = "NewColoCompleteList.aspx?pNo=" & DNo.Text
                LComplete.Visible = True
            Else
                LComplete.Visible = False
            End If

            'If DBAdapter1.Rows(0).Item("DTNO") = 1 Then       '有追加核可單
            ' LDTSheet.NavigateUrl = "NewColorDTSheetList.aspx?pNo=" & DNo.Text
            ' LDTSheet.Visible = True
            'Else
            '   LDTSheet.Visible = False
            'End If

        End If


        '核定履歷資料
        SQL = "SELECT "
        SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
        SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
        SQL = SQL + "AEndTimeDesc As Description, "
        SQL = SQL + "URL "
        SQL = SQL + "FROM V_WaitHandle_01 "
        SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
        SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
        SQL = SQL + "Order by Unique_ID Desc "
        Dim dtWaitHandle1 As DataTable = uDataBase.GetDataTable(SQL)

        'Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter4.Fill(DBDataSet9, "DecideHistory")
        GridView2.DataSource = dtWaitHandle1
        GridView2.DataBind()

        'DB連結關閉
        'OleDbConnection1.Close()

    End Sub


    '*****************************************************************
    '**(ShowSheetField)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim DBDataSet1 As New DataSet

        Dim SQL As String
        Dim idx As Integer
        Dim i As Integer



        idx = FindFieldInf(pFieldName)



        '光源
        If pFieldName = "ColorLight1" Then
            DColorLight1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DColorLight1.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'ColorLight'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DColorLight1.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DColorLight1.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        '光源
        If pFieldName = "ColorLight2" Then
            DColorLight2.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DColorLight2.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  data from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'ColorLight2' "
                SQL = SQL & " order by createtime desc "



                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DColorLight2.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DColorLight2.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        'YKK色別
        If pFieldName = "YKKColorType" Then
            DYKKColorType.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DYKKColorType.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'YKKColorType'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DYKKColorType.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DYKKColorType.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'YKK色別
        If pFieldName = "ColorType" Then
            DAgain.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAgain.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'ColorType'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DAgain.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAgain.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'YKK色別
        If pFieldName = "CustomerSample" Then

            DCustomerSample.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCustomerSample.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'CustomerSample'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DCustomerSample.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCustomerSample.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        '開發年
        If pFieldName = "DevYear" Then

            DDevYear.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDevYear.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'DevYear' order by data"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DDevYear.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDevYear.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        '開發季
        If pFieldName = "DevSeason" Then

            DDevSeason.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDevSeason.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'DevSeason'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DDevSeason.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDevSeason.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        '新色舊色
        If pFieldName = "NewOldColor" Then
            DNewOldColor.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DNewOldColor.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'NewOldColor'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DNewOldColor.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DNewOldColor.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
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



    Protected Sub DYKKColorCode_ServerChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles DYKKColorCode.ServerChange

    End Sub
End Class

