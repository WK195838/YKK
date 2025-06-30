Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DTMW_NewColorAgain_01
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




    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '設定共用參數


        If Not Me.IsPostBack Then   '不是PostBack
            SetFieldAttribute("Post")
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
        wStep = Request.QueryString("pStep")        '工程代碼      
        wSeqNo = Request.QueryString("pSeqNo")      '序號
        wApplyID = Request.QueryString("pApplyID")  '申請者ID
        wAgentID = Request.QueryString("pAgentID")  '被代理人ID
        wUserID = Request.QueryString("pUserID")  '被代理人ID
        'Add-Start by Joy 2009/11/20(2010行事曆對應)
        '申請者-群組行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        wDecideCalendar = oCommon.GetCalendarGroup(wUserID)
        Response.Cookies("PGM").Value = "DTMW_NewColorComplete.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼

    End Sub



    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Table As String = ""

        Dim SQL As String
        SQL = "Select * From M_referp "
        SQL = SQL & " Where Cat='5001' "
        SQL = SQL & " and dkey  ='Sheet-" + wFormNo + "'"
        Dim DBTable As DataTable = uDataBase.GetDataTable(SQL)
        If DBTable.Rows.Count > 0 Then
            LSheetName.Text = DBTable.Rows(0).Item("Data")
        End If

        SQL = "Select * From v_NewColor "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then

            If DBAdapter1.Rows(0).Item("CustomerCheck") = 1 Then
                DCustomerCheck.Checked = True
            End If
            If DBAdapter1.Rows(0).Item("FactoryCheck") = 1 Then
                DFactoryCheck.Checked = True
            End If
            If DBAdapter1.Rows(0).Item("VCACheck") = 1 Then
                DVCACheck.Checked = True
            End If

            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            '  DBarCode.ImageUrl = "imga.aspx?mycode=" & DNo.Text  BARCODE
            DDate.Text = DBAdapter1.Rows(0).Item("Date")                         'No 
            DDepoName.Text = DBAdapter1.Rows(0).Item("DepoName")                         'No
            DName.Text = DBAdapter1.Rows(0).Item("Name")                         'No
            DCustomer.Text = DBAdapter1.Rows(0).Item("Customer")                         'No
            DCustomerCode.Text = DBAdapter1.Rows(0).Item("CustomerCode")                         'No
            DBuyer.Text = DBAdapter1.Rows(0).Item("Buyer")                         'No
            DCustomerColor.Text = DBAdapter1.Rows(0).Item("CustomerColor")                         'No3
            DCustomerColorCode.Text = DBAdapter1.Rows(0).Item("CustomerColorCode")
            DOverseaYKKCode.Text = DBAdapter1.Rows(0).Item("OverseaYKKCode")
            DPANTONECode.Text = DBAdapter1.Rows(0).Item("PANTONECode")
            DDevYear.Text = DBAdapter1.Rows(0).Item("DevYear")
            DDevSeason.Text = DBAdapter1.Rows(0).Item("DevSeason")

            If Mid(DBAdapter1.Rows(0).Item("ReceiveDate").ToString, 1, 4) = "1900" Then
                DDReceiveDate.Text = ""
            Else
                DDReceiveDate.Text = DBAdapter1.Rows(0).Item("ReceiveDate")               '下單時間
            End If
            DColorLight1.Text = DBAdapter1.Rows(0).Item("ColorLight1")
            DColorLight1.Text = DBAdapter1.Rows(0).Item("ColorLight2")
            DYKKColorType.Text = DBAdapter1.Rows(0).Item("YKKColorType")
            DYKKColorCode.Text = DBAdapter1.Rows(0).Item("YKKColorCode")

        End If

        'DTM開始日
        SQL = " select  convert(char(10), max(astarttime),111) astarttime from t_waithandle"
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & " And FormSno =  '" & CStr(wFormSno) & "'"
        SQL = SQL & " and step =60 "
        Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter2.Rows.Count > 0 Then
            DAgainSdate.Text = DBAdapter2.Rows(0).Item("Astarttime")
        End If


        'DTM完成日
        SQL = " select  convert(char(10),max(AEndtime),111) AEndtime from t_waithandle "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & " And FormSno =  '" & CStr(wFormSno) & "'"
        SQL = SQL & " and step= 60  "
        Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter3.Rows.Count > 0 Then
            DAgainEDate.Text = DBAdapter3.Rows(0).Item("AEndtime")
        End If

        '計算DTM天數
        Dim date1 As DateTime
        Dim date2 As DateTime
        date1 = DAgainSdate.Text
        date2 = DAgainEDate.Text
        Dim ts As TimeSpan
        ts = date2 - date1
        Dim iDays As Integer
        iDays = ts.Days
        DAgainDays.Text = iDays

        SetFieldData("CheckType", "ZZZZZZ")
        SetFieldData("Color", "ZZZZZZ")




    End Sub

    '*****************************************************************
    '**(ShowSheetField)
    '**     建立表單需輸入欄位
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator

        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.ErrorMessage = pMessage
        rqdVal.Display = ValidatorDisplay.None              ' 因在頁面上加入ValidationSummary , 故驗證控制項統一顯示
        rqdVal.Style.Add("Top", Top + 300 & "px")
        rqdVal.Style.Add("Left", "8")
        rqdVal.Style.Add("Height", "20")
        rqdVal.Style.Add("Width", "250")
        rqdVal.Style.Add("POSITION", "absolute")
        Me.Form.Controls.Add(rqdVal)
    End Sub

    '*****************************************************************
    '**(ShowSheetField)
    '**     表單各欄位屬性及欄位輸入檢查等設定
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        '表單各欄位屬性及欄位輸入檢查等設定


        ShowRequiredFieldValidator("DDyeTimesRqd", "DDyeTimes", "異常：需輸入染色次數")

    End Sub



    '*****************************************************************
    '**(ShowSheetField)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)




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


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     新增表單資料
    '**
    '*****************************************************************
    Sub AppendData()


        Dim SQl As String

        SQl = "Select * From M_referp "
        SQl = SQl & " Where Cat='5001' "
        SQl = SQl & "and dkey  ='Table-" + wFormNo + "'"
        Dim DBTable1 As DataTable = uDataBase.GetDataTable(SQl)
        Dim wTable As String = ""
        If DBTable1.Rows.Count > 0 Then
            wTable = DBTable1.Rows(0).Item("Data")
        End If


        SQl = "Insert into F_NewColorAgain "
        SQl = SQl + "( "
        SQl = SQl + "CompletedTime, FormNo, FormSno,"
        SQl = SQl + "No, AgainSDate,AgainEDate,AgainDays,DyeTimes,"  '1~5                
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime) "
        SQl = SQl + "VALUES( "
        SQl = SQl + " '" + NowDateTime + "', "        '結案日
        SQl = SQl + " '" + CStr(wFormNo) + "', "   '表單流水號
        SQl = SQl + " '" + CStr(wFormSno) + "', "   '表單流水號
        SQl = SQl + " N'" + DNo.Text + "', "   'NO  1     
        SQl = SQl + " '" + DAgainSdate.Text + "', "                '客戶8
        SQl = SQl + " '" + DAgainEDate.Text + "', "                '客戶9
        SQl = SQl + " '" + DAgainDays.Text + "', "                '客戶9
        SQl = SQl + " '" + DDyeTimes.Text + "', "                'BUYER11   
        SQl = SQl + " '" + wUserID + "', "     '作成者
        SQl = SQl + " '" + NowDateTime + "', "       '作成時間
        SQl = SQl + " '" + "" + "', "                       '修改者
        SQl = SQl + " '" + NowDateTime + "' "       '修改時間
        SQl = SQl + " ) "
        uDataBase.ExecuteNonQuery(SQl)

   
        SQl = " Update  " + wTable
        SQl = SQl & " Set Again = 1"
        SQl = SQl & " Where FormNo =  '" & wFormNo & "'"
        SQl = SQl & " And FormSno =  '" & CStr(wFormSno) & "'"
        uDataBase.ExecuteNonQuery(SQl)



    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click


        Dim SQL As String  '檢查是否有相同版本的新色依賴完成書
        SQL = " select  *  from F_NewColorAgain "
        SQL = SQL & " where No='" + DNo.Text + "'"
        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL)
        If dtFlow.Rows.Count = 0 Then '如果沒有才可以新增
            AppendData()  'Insert
            Dim URL As String
            URL = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                                                              "&pUserID=" & wUserID & "&pApplyID=" & wApplyID
            Response.Redirect(URL)
        Else
            uJavaScript.PopMsg(Me, "單號：" + DNo.Text + "-新色再現完成表已完成!")
        End If




    End Sub
End Class

