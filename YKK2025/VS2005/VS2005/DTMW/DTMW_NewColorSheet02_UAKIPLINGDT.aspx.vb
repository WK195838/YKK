Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DTMW_NewColorSheet02_UAKIPLINGDT
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
        wStep = Request.QueryString("pStep")        '工程代碼
        wbFormSno = Request.QueryString("pFormSno")  '連續起單用流水號
        wbStep = Request.QueryString("pStep")        '連續起單用工程代碼

        wSeqNo = Request.QueryString("pSeqNo")      '序號
        wApplyID = Request.QueryString("pApplyID")  '申請者ID

        wAgentID = Request.QueryString("pAgentID")  '被代理人ID
        'Add-Start by Joy 2009/11/20(2010行事曆對應)
        '申請者-群組行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)

        wUserID = Request.QueryString("pUserID")

        'wUserID = Request.QueryString("UserID")
        wDecideCalendar = oCommon.GetCalendarGroup(wUserID)

        Response.Cookies("PGM").Value = "DTMW_NewColorSheet01_03CFP12.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼

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

        SQL = "Select * From F_NewColorUAKIPLINGDT "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then



            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            '   DBarCode.ImageUrl = "imga.aspx?mycode=" & DNo.Text  BARCODE
            DDate.Text = DBAdapter1.Rows(0).Item("Date")                         'No 
            DDepoName.Text = DBAdapter1.Rows(0).Item("DepoName")                         'No
            DName.Text = DBAdapter1.Rows(0).Item("Name")                         'No
            DNO1.Text = DBAdapter1.Rows(0).Item("No1")

            DColorSystem.Text = DBAdapter1.Rows(0).Item("ColorSystem")

            DPFBWire.Text = DBAdapter1.Rows(0).Item("PFBWire")
            DPFOpenParts.Text = DBAdapter1.Rows(0).Item("PFOpenParts")

            SetFieldData("YKKColorType", DBAdapter1.Rows(0).Item("YKKColorType"))

            DYKKColorCode.Value = DBAdapter1.Rows(0).Item("YKKColorCode")

            SetFieldData("NewOldColor", DBAdapter1.Rows(0).Item("NewOldColor"))

            DSLDColor.Text = DBAdapter1.Rows(0).Item("SLDColor")
            DVFColor.Text = DBAdapter1.Rows(0).Item("VFColor")

            SetFieldData("DTSheet", DBAdapter1.Rows(0).Item("DTSheet"))


        End If


        '帶入最新屢歷
        Dim sqlhistory As String
        sqlhistory = " SELECT"
        sqlhistory = sqlhistory + " a.No  As Field1,"
        sqlhistory = sqlhistory + " case b.Sts When '0' Then '核定中' When '1' Then '完成' When '2' Then '取消' Else '取消' End As Field2,"
        sqlhistory = sqlhistory + " Convert(VARCHAR(10), a.ApplyTime, 111) as Field3,"
        sqlhistory = sqlhistory + " case when b.sts ='0' then '' else Convert(VARCHAR(10), b.completedtime, 111)  end as completedate,"
        sqlhistory = sqlhistory + " a.FormName as Field4,"
        sqlhistory = sqlhistory + " a.Division As Field5,a.ApplyName as Field6,b.Customer As Field7,b.Buyer As Field8,"
        sqlhistory = sqlhistory + " YKKColorType As Field9,YKKColorCode as Field10,SLDColor As Field11,VFColor As Field12,NewOldColor,"
        sqlhistory = sqlhistory + " '....' as WorkFlow, ViewURL,"
        sqlhistory = sqlhistory + " 'http://10.245.1.10/WorkFlow/BefOPList.aspx?' +"
        sqlhistory = sqlhistory + " 'pFormNo='   + a.FormNo +"
        sqlhistory = sqlhistory + " '&pFormSno=' + str(a.FormSno,Len(a.FormSno)) +"
        sqlhistory = sqlhistory + " '&pStep='    + str(a.Step,Len(a.Step)) +"
        sqlhistory = sqlhistory + " '&pSeqNo='   + str(a.SeqNo,Len(a.SeqNo)) +"
        sqlhistory = sqlhistory + " '&pApplyID=' + a.ApplyID"
        sqlhistory = sqlhistory + " As OPURL, "
        sqlhistory = sqlhistory + " Case when b.sts = 1 then convert(char(10),b.completedTime,111) else '' end as completedDate,customerColor,"
        sqlhistory = sqlhistory + " customerColorCode,overSeaYkkCode,pantonecode,substring(stepnamedesc,7,len(stepnamedesc )-1)stepnamedesc,a.FormSno "
        sqlhistory = sqlhistory + " from V_WaitHandle_01 a,V_NewColor b"
        sqlhistory = sqlhistory + " Where a.formno=b.formno and a.formsno =b.formsno and active  = '1' "
        sqlhistory = sqlhistory + " and a.no ='" & DNO1.Text & "'"
        Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(sqlhistory)
        Dim i As Integer

        If DBAdapter3.Rows.Count > 0 Then
            HyperLink1.NavigateUrl = DBAdapter3.Rows(0).Item("OPURL")  '請假證明
            HyperLink1.Visible = True


            MNOSts.Text = ""

            For i = 0 To DBAdapter3.Rows.Count - 1
                If MNOSts.Text = "" Then
                    MNOSts.Text = DBAdapter3.Rows(i).Item("stepnamedesc")
                Else
                    MNOSts.Text = MNOSts.Text + "," + DBAdapter3.Rows(i).Item("stepnamedesc")
                End If

            Next
            '    MNOSts.Text = DBAdapter3.Rows(0).Item("stepnamedesc")
            MNOSts.Visible = True
            DOFormSno.Text = DBAdapter3.Rows(0).Item("FormSno")

        End If




    End Sub

    '*****************************************************************
    '**(ShowSheetField)
    '**     表單各欄位屬性及欄位輸入檢查等設定
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        '表單各欄位屬性及欄位輸入檢查等設定


        'Dim SQL As String


        'MNo
        Select Case FindFieldInf("NO1")
            Case 0  '顯示
                DNO1.BackColor = Color.LightGray
                DNO1.ReadOnly = True
                DNO1.Visible = True

            Case 1  '修改+檢查
                DNO1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNO1Rqd", "DNO1", "異常：需輸入Ｎｏ")
                DNO1.Visible = True

            Case 2  '修改
                DNO1.BackColor = Color.Yellow
                DNO1.Visible = True

            Case Else   '隱藏
                DNO1.Visible = False

        End Select
        If pPost = "New" Then DNO1.Text = ""





        'No
        Select Case FindFieldInf("No")
            Case 0  '顯示
                DNo.BackColor = Color.LightGray
                DNo.ReadOnly = True
                DNo.Visible = True
            Case 1  '修改+檢查
                DNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DnoRqd", "Dno", "異常：需輸入Ｎｏ")
                DNo.Visible = True
            Case 2  '修改
                DNo.BackColor = Color.Yellow
                DNo.Visible = True
            Case Else   '隱藏
                DNo.Visible = False
        End Select
        If pPost = "New" Then DNo.Text = ""



        '日期
        Select Case FindFieldInf("Date")
            Case 0  '顯示
                DDate.BackColor = Color.LightGray
                DDate.ReadOnly = True
                DDate.Visible = True

            Case 1  '修改+檢查
                DDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDateRqd", "DDate", "異常：需輸入日期")
                DDate.Visible = True

            Case 2  '修改
                DDate.BackColor = Color.Yellow
                DDate.Visible = True

            Case Else   '隱藏
                DDate.Visible = False

        End Select
        If pPost = "New" Then DDate.Text = Now.ToString("yyyy/MM/dd") '現在日時

        'DepoName
        Select Case FindFieldInf("DepoName")
            Case 0  '顯示
                DDepoName.BackColor = Color.LightGray
                DDepoName.ReadOnly = True
                DDepoName.Visible = True
            Case 1  '修改+檢查
                DDepoName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDepoNameRqd", "DDepoName", "異常：需輸入部門")
                DDepoName.Visible = True
            Case 2  '修改
                DDepoName.BackColor = Color.Yellow
                DDepoName.Visible = True
            Case Else   '隱藏
                DDepoName.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DepoName", "ZZZZZZ")


        'Name
        Select Case FindFieldInf("Name")
            Case 0  '顯示
                DName.BackColor = Color.LightGray
                DName.ReadOnly = True
                DName.Visible = True
            Case 1  '修改+檢查
                DName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNameRqd", "DName", "異常：需輸入姓名")
                DName.Visible = True
            Case 2  '修改
                DName.BackColor = Color.Yellow
                DName.Visible = True
            Case Else   '隱藏
                DName.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Name", "ZZZZZZ")





        'SLDColor
        Select Case FindFieldInf("SLDColor")
            Case 0  '顯示
                DSLDColor.BackColor = Color.LightGray
                DSLDColor.ReadOnly = True
                DSLDColor.Visible = True

            Case 1  '修改+檢查
                DSLDColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSLDColorRqd", "DSLDColor", "異常：需輸入確認拉頭兼用色")
                DSLDColor.Visible = True

            Case 2  '修改
                DSLDColor.BackColor = Color.Yellow
                DSLDColor.Visible = True
            Case Else   '隱藏
                DSLDColor.Visible = False

        End Select
        If pPost = "New" Then SetFieldData("SLDColor", "ZZZZZZ")

        'VFColor 
        Select Case FindFieldInf("VFColor")
            Case 0  '顯示
                DVFColor.BackColor = Color.LightGray
                DVFColor.ReadOnly = True
                DVFColor.Visible = True
            Case 1  '修改+檢查
                DVFColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVFColorRqd", "DVFColor", "異常：需輸入確認VF兼用色")
                DVFColor.Visible = True

            Case 2  '修改
                DVFColor.BackColor = Color.Yellow
                DVFColor.Visible = True
            Case Else   '隱藏
                DVFColor.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("VFColor", "ZZZZZZ")







        'YKK色別
        Select Case FindFieldInf("YKKColorType")
            Case 0  '顯示
                DYKKColorType.BackColor = Color.LightGray
                DYKKColorType.Visible = True

            Case 1  '修改+檢查
                DYKKColorType.BackColor = Color.GreenYellow
                '      ShowRequiredFieldValidator("DYKKColorTypeRqd", "DYKKColorType", "異常：需輸YKK色別")
                DYKKColorType.Visible = True
            Case 2  '修改
                DYKKColorType.BackColor = Color.Yellow
                DYKKColorType.Visible = True
            Case Else   '隱藏
                DYKKColorType.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("YKKColorType", "ZZZZZZ")




        'YKKColorCode
        Select Case FindFieldInf("YKKColorCode")
            Case 0  '顯示
                DYKKColorCode.Visible = True
                DYKKColorCode.Style.Add("background-color", "lightgrey")
                DYKKColorCode.Attributes.Add("readonly", "true")
            
            Case 1  '修改+檢查
                DYKKColorCode.Visible = True
                DYKKColorCode.Style.Add("background-color", "greenyellow")
                '  DYKKColorCode.Attributes.Add("readonly", "true")
                '  ShowRequiredFieldValidator("DYKKColorCodeRqd", "DYKKColorCode", "異常：需輸入YKK色號")


            Case 2  '修改
                DYKKColorCode.Visible = True
                DYKKColorCode.Style.Add("background-color", "yellow")
                DYKKColorCode.Attributes.Add("readonly", "true")

            Case Else   '隱藏
                DYKKColorCode.Visible = False


        End Select
        If pPost = "New" Then DYKKColorCode.Value = ""



        'PFBWire


        'PFBWire 
        Select Case FindFieldInf("PFBWire")
            Case 0  '顯示
                DPFBWire.BackColor = Color.LightGray
                DPFBWire.ReadOnly = True
                DPFBWire.Visible = True
            Case 1  '修改+檢查
                DPFBWire.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPFBWireRqd", "DPFBWire", "異常：需輸入確認VF兼用色")
                DPFBWire.Visible = True

            Case 2  '修改
                DPFBWire.BackColor = Color.Yellow
                DPFBWire.Visible = True
            Case Else   '隱藏
                DPFBWire.Visible = False
        End Select
        If pPost = "New" Then DPFBWire.Text = ""



        'DPFOpenParts 
        Select Case FindFieldInf("PFOpenParts")
            Case 0  '顯示
                DPFOpenParts.BackColor = Color.LightGray
                DPFOpenParts.ReadOnly = True
                DPFOpenParts.Visible = True
            Case 1  '修改+檢查
                DPFOpenParts.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPFOpenPartsRqd", "DPFOpenParts", "異常：需輸入確認VF兼用色")
                DPFOpenParts.Visible = True

            Case 2  '修改
                DPFOpenParts.BackColor = Color.Yellow
                DPFOpenParts.Visible = True
            Case Else   '隱藏
                DPFOpenParts.Visible = False
        End Select
        If pPost = "New" Then DPFOpenParts.Text = ""




        'YKK色別
        Select Case FindFieldInf("ColorSystem")
            Case 0  '顯示
                DColorSystem.BackColor = Color.LightGray
                DColorSystem.Visible = True

            Case 1  '修改+檢查
                DColorSystem.BackColor = Color.GreenYellow
                '     ShowRequiredFieldValidator("DColorSystemRqd", "DColorSystem", "異常：需輸色系")
                DColorSystem.Visible = True
            Case 2  '修改
                DColorSystem.BackColor = Color.Yellow
                DColorSystem.Visible = True
            Case Else   '隱藏
                DColorSystem.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("ColorSystem", "ZZZZZZ")




        '新舊色
        Select Case FindFieldInf("NewOldColor")
            Case 0  '顯示
                DNewOldColor.BackColor = Color.LightGray
                DNewOldColor.Visible = True

            Case 1  '修改+檢查
                DNewOldColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNewOldColorRqd", "DNewOldColor", "異常：需輸新舊色")
                DNewOldColor.Visible = True
            Case 2  '修改
                DNewOldColor.BackColor = Color.Yellow
                DNewOldColor.Visible = True
            Case Else   '隱藏
                DNewOldColor.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("NewOldColor", "ZZZZZZ")


      

        '核可單種類
        Select Case FindFieldInf("DTSheet")
            Case 0  '顯示
                DDTSheet.BackColor = Color.LightGray
                DDTSheet.Visible = True

            Case 1  '修改+檢查
                DDTSheet.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDTSheetRqd", "DDTSheet", "異常：需輸核可單種類")
                DDTSheet.Visible = True
            Case 2  '修改
                DDTSheet.BackColor = Color.Yellow
                DDTSheet.Visible = True
            Case Else   '隱藏
                DDTSheet.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("DTSheet", "ZZZZZZ")


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
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim DBDataSet1 As New DataSet

        Dim SQL As String
        Dim idx As Integer
        Dim i As Integer



        idx = FindFieldInf(pFieldName)

      


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


        '核可單種類
        If pFieldName = "DTSheet" Then
            DDTSheet.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDTSheet.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'DTSheet'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DDTSheet.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDTSheet.Items.Add(ListItem1)
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

 
End Class

