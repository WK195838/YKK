Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DTMW_NewColorSheet02_05CNLSBS16
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
        LFormSno.Text = "單號:" + CStr(wFormSno)



        If Now.ToString("yyyy/MM/dd") < "2025/04/01" Then
            LSheetName.Text = "拉鏈ZIPPER CHAIN (05 CNL SBS16)" '表單名稱"
        Else
            LSheetName.Text = "型別轉換-05CNPBS16" '表單名稱"
        End If

    End Sub



    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim SQL As String

        SQL = "Select * From F_NewColor05CNLSBS16 "
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
            DBuyerCode.Text = DBAdapter1.Rows(0).Item("BuyerCode")
            DCustomerColor.Text = DBAdapter1.Rows(0).Item("CustomerColor")                         'No3
            DCustomerColorCode.Text = DBAdapter1.Rows(0).Item("CustomerColorCode")
            DOverseaYKKCode.Text = DBAdapter1.Rows(0).Item("OverseaYKKCode")
            DReColorCode.Text = DBAdapter1.Rows(0).Item("RecolorCode")

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
            SetFieldData("NewOldColor", DBAdapter1.Rows(0).Item("NewOldColor"))    '類別2

            DDeliveryDate.Text = DBAdapter1.Rows(0).Item("DeliveryDate")
            DDuplicateNo.Text = DBAdapter1.Rows(0).Item("DuplicateNo")
            If DBAdapter1.Rows(0).Item("NOCCS") = 1 Then
                DNOCCS.Checked = True
            End If
            DNOCCSReason.Text = DBAdapter1.Rows(0).Item("NOCCSReason")
            DCustomerNGColor.Text = DBAdapter1.Rows(0).Item("CustomerNGColor")
            DRemark.Text = DBAdapter1.Rows(0).Item("Remark")
            SetFieldData("CustomerSample", DBAdapter1.Rows(0).Item("CustomerSample"))    '類別2

            DColorSystem.Text = DBAdapter1.Rows(0).Item("ColorSystem")
            DPFBWire.Text = DBAdapter1.Rows(0).Item("PFBWire")
            DPFOpenParts.Text = DBAdapter1.Rows(0).Item("PFOpenParts")
            SetFieldData("YKKColorType", DBAdapter1.Rows(0).Item("YKKColorType"))    '類別2
            DYKKColorCode.Text = DBAdapter1.Rows(0).Item("YKKColorCode")
            DVersion.Text = DBAdapter1.Rows(0).Item("Version")
            DSLDColor.Text = DBAdapter1.Rows(0).Item("SLDColor")
            SetFieldData("ColorType", DBAdapter1.Rows(0).Item("Again"))    '類別2
            If DBAdapter1.Rows(0).Item("again") = 1 Then
                DAgain.SelectedValue = "淡色"
            ElseIf DBAdapter1.Rows(0).Item("again") = 2 Then
                DAgain.SelectedValue = "濃色"
            End If

            DVFColor.Text = DBAdapter1.Rows(0).Item("VFColor")
            DCheckNo.Text = DBAdapter1.Rows(0).Item("CheckNo")
            If DBAdapter1.Rows(0).Item("Complete") = 1 Then       '有新色依賴完成表
                LComplete.NavigateUrl = "NewColoCompleteList.aspx?pNo=" & DNo.Text
                LComplete.Visible = True
            Else
                LComplete.Visible = False
            End If



            'If DBAdapter1.Rows(0).Item("Again") = 1 Then      '有新色用現完成表
            '    LAgain.NavigateUrl = "NewColorAgainList.aspx?pNo=" & DNo.Text
            '    LAgain.Visible = True
            'Else
            '    LAgain.Visible = False
            'End If

        End If







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
                SQL = SQL & " and dkey = 'DevYear'"


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


End Class

