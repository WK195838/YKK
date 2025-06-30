Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb

Partial Class FASSheet_02
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
    Dim Count As Integer
    Dim intTotal As Decimal = 0
    Dim AMT As Decimal = 0
    Dim INVQTY As Decimal = 0
    Dim INVAMT As Decimal = 0
    Dim SCount As Decimal = 0
    Dim DetailSQL As String


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '設定共用參數


        If Not Me.IsPostBack Then   '不是PostBack

            ShowFormData()      '顯示表單資料

        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     按鈕及RequestedField的Top位置
    '**
    '*****************************************************************
    Sub TopPosition()
        Top = 500
    End Sub



    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetControlPosition)
    '**    設定按鈕,說明,延遲,核定履歷的位置
    '**
    '*****************************************************************
    Sub SetControlPosition()
        If GridView3.Rows.Count > 0 Then
            Dim GVTop As Integer = GridView3.Style("top").Replace("px", "")
            Dim GVHeight As Integer = GridView3.Rows.Count * 40       ' 55是列高

            Dim ControlTop As Integer = (GVTop + GVHeight)
           

        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日時
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號    
        wApplyID = Request.QueryString("pApplyID")  '申請者ID
     
        '
        Response.Cookies("PGM").Value = "FASSheet_01.aspx"
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
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, Request.QueryString("pUserID"))
        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者
    End Sub
   
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
    '**     表單欄位顯示及欄位輸入檢查
    '**
    '*****************************************************************
    Sub ShowSheetField(ByVal pPost As String)
        '建置欄位及屬性陣列
        oCommon.GetFieldAttribute(wFormNo, wStep, FieldName, Attribute)
        '表單號碼,工程關卡號碼,欄位名,欄位屬性


    End Sub
  
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()

        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("FASFilePath")
        Dim SQL As String
        SQL = "Select * From F_FASSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dt As DataTable = uDataBase.GetDataTable(SQL)
        If dt.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dt.Rows(0).Item("No")
            DApper.Text = dt.Rows(0).Item("Apper")
            DAppDate.Text = dt.Rows(0).Item("AppDate")
            DDivision.Text = dt.Rows(0).Item("Division")
            DTYPE.Text = dt.Rows(0).Item("TYPE")
            DRemark.Text = dt.Rows(0).Item("Remark")

        End If


        SQL = "Select a.sts,Buyer,ITEMCODE,qty,UnitPrice AMT,invqty,convert(int,invamt)invamt,attachfile From F_FASSheet a,F_FASSheetDT b "
        SQL &= " Where a.FormNO= b.FormNo and a.formsno = b.formsno and a.FormNo =  '" & wFormNo & "'"
        SQL &= "   And a.FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtDttail As DataTable = uDataBase.GetDataTable(SQL)
        If dtDttail.Rows.Count > 0 Then
            '細項內容
            If dtDttail.Rows(0).Item("sts") = "1" Then
                LFCDataList.NavigateUrl = "FCDataList.aspx?Userid=" & Request.QueryString("pUserID") & "&pFormSno=" & CStr(wFormSno) & "&pSts=1"
            Else
                LFCDataList.NavigateUrl = "FCDataList.aspx?Userid=" & Request.QueryString("pUserID") & "&pFormSno=" & CStr(wFormSno) & "&pSts=0"
            End If

            Count = dtDttail.Rows.Count
            GridView3.DataSource = dtDttail
            GridView3.DataBind()
        End If


        SetControlPosition()    ' 設定控制項位置

    End Sub
    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)

    End Sub

    '*****************************************************************
    '**(ShowRequiredFieldValidator)
    '**     建立表單需輸入欄位
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator
        '
        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.Text = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", Top & "px")
        rqdVal.Style.Add("Left", "8px")
        rqdVal.Style.Add("Height", "20px")
        rqdVal.Style.Add("Width", "250px")
        rqdVal.Style.Add("POSITION", "absolute")
        Me.Form.Controls.Add(rqdVal)
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

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()

    End Sub
   

    Protected Sub GridView3_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView3.RowDataBound
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("FASFilePath")
        Dim formatNu As Integer

        '附檔連結
        If e.Row.RowType <> DataControlRowType.Footer And e.Row.RowType <> DataControlRowType.Header Then
            If Trim(e.Row.Cells(6).Text) <> "&nbsp;" Then '不是空白
                Dim h1 As New HyperLink
                ' h1.Text = e.Row.Cells(3).Text
                h1.Text = "附檔"
                h1.Target = "_blank"
                h1.NavigateUrl = Path & e.Row.Cells(6).Text  '上傳檔案
                ' e.Row.Cells(3).Text = ""
                e.Row.Cells(6).Controls.Add(h1)
            End If
            intTotal += e.Row.Cells(2).Text
            AMT += e.Row.Cells(3).Text
            INVQTY += e.Row.Cells(4).Text
            INVAMT += e.Row.Cells(5).Text
            SCount = SCount + 1
        End If



        If e.Row.RowType <> DataControlRowType.Header Then
            e.Row.Cells(2).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(3).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(4).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(5).HorizontalAlign = HorizontalAlign.Right
        End If



        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(1).Text = FormatNumber(SCount, formatNu, TriState.True, TriState.False, TriState.True) + "件"
            e.Row.Cells(2).Text = FormatNumber(intTotal, formatNu, TriState.True, TriState.False, TriState.True)
            e.Row.Cells(3).Text = FormatNumber(AMT, formatNu, TriState.True, TriState.False, TriState.True)
            e.Row.Cells(4).Text = FormatNumber(INVQTY, formatNu, TriState.True, TriState.False, TriState.True)
            e.Row.Cells(5).Text = FormatNumber(INVAMT, formatNu, TriState.True, TriState.False, TriState.True)

            e.Row.Cells(0).Text = "合計"

        End If


    End Sub


     
End Class
