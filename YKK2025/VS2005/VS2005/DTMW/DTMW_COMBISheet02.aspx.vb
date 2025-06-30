Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DTMW_COMBISheet02
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
    Dim LastStep As String




    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '設定共用參數

        ShowFormData()      '顯示表單資料
        ' UpdateTranFile()    '更新交易資料
        ' SetPopupFunction()      '設定彈出視窗事件
        ' ChangeVisible()
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
      
    End Sub



    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()



        Dim SQL As String

        SQL = "Select * From F_COMBISheet "
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




            DDuplicateNo.Text = DBAdapter1.Rows(0).Item("DuplicateNo")


            DYKKColorType.Text = DBAdapter1.Rows(0).Item("YKKColorType")



            DYKKColorCode.Text = DBAdapter1.Rows(0).Item("YKKColorCode")
            DCOMBIItem.Text = DBAdapter1.Rows(0).Item("COMBIITem")

            If DCOMBIItem.Text = "塑鋼 VS (一般)" Or DCOMBIItem.Text = "塑鋼VT810 / VT108" Then
                DITEMNAME1.Text = DCOMBIItem.Text
            End If

            'VF 
            DVFLTape.Text = DBAdapter1.Rows(0).Item("VFLTape")
            DVFLChain.Text = DBAdapter1.Rows(0).Item("VFLChain")
            DVFRChain.Text = DBAdapter1.Rows(0).Item("VFRChain")
            DVFRTape.Text = DBAdapter1.Rows(0).Item("VFRTape")

            'VF & MF
            DVFMLTape.Text = DBAdapter1.Rows(0).Item("VFMLTape")
            DVFMLChain.Text = DBAdapter1.Rows(0).Item("VFMLChain")
            DVFMRChain.Text = DBAdapter1.Rows(0).Item("VFMRChain")
            DVFMRTape.Text = DBAdapter1.Rows(0).Item("VFMRTape")

            'VF & PF
            DPFMFLTape.Text = DBAdapter1.Rows(0).Item("PFMFLTape")
            DPFMFRTape.Text = DBAdapter1.Rows(0).Item("PFMFRTape")
        End If





    End Sub
 

 



End Class

