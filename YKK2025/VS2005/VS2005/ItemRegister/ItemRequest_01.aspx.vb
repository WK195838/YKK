Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb

Partial Class ItemRequest_01
    Inherits System.Web.UI.Page
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim uErr As Boolean
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page_Load
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()              '設定共用參數
        SetButtonFunction()         '設定功能按鈕
        If Not IsPostBack Then
            GetWaitRequest()        '取得待處理資料
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        Response.Cookies("PGM").Value = "ItemRequest_01.aspx"
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日時
        uErr = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetButtonFunction)
    '**     功能按鈕顯示
    '**
    '*****************************************************************
    Sub SetButtonFunction()
        '展開
        BExpandRefItem.Attributes("onclick") = "var ok=window.confirm('" + "前次資料將會清除是否執行[" + BExpandRefItem.Text + "]作業嗎？\n\r請確認...." + "');if(!ok){return false;}"
        '上傳
        BUpload.Attributes("onclick") = "var ok=window.confirm('" + "將會將資料轉入Waves是否執行[" + BUpload.Text + "]作業嗎？\n\r請確認...." + "');if(!ok){return false;}"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetWaitRequest)
    '**     取得待處理資料
    '**
    '*****************************************************************
    Sub GetWaitRequest()
        Dim SQL As String = ""
        Dim xFormNo As String = ""
        Dim xNo As String = ""
        '
        SQL = "Select "
        SQL = SQL & "Case FormNo When '001152' Then 'ZIP-' When '001153' Then 'SLD-' When '001154' Then 'CH -' Else 'SLDF-' End + No As DataType, "
        SQL = SQL & "No "
        SQL = SQL & "From W_ItemRequest "
        SQL = SQL & "Group by FormNo, No "
        SQL = SQL & "Order by FormNo, No "
        Dim dt_WaitRequest As DataTable = uDataBase.GetDataTable(SQL)
        DWaitRequest.Items.Clear()
        If dt_WaitRequest.Rows.Count > 0 Then
            For i As Integer = 0 To dt_WaitRequest.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = dt_WaitRequest.Rows(i)("DataType").ToString
                ListItem1.Value = dt_WaitRequest.Rows(i)("No").ToString
                DWaitRequest.Items.Add(ListItem1)
            Next
            '
            If InStr(DWaitRequest.SelectedItem.Text, "ZIP") > 0 Then
                xFormNo = "001152"
            Else
                If InStr(DWaitRequest.SelectedItem.Text, "SLDF") > 0 Then
                    xFormNo = "001155"
                Else
                    If InStr(DWaitRequest.SelectedItem.Text, "CH ") > 0 Then
                        xFormNo = "001154"
                    Else
                        xFormNo = "001153"
                    End If
                End If
            End If
            xNo = DWaitRequest.SelectedValue
        End If
        '
        ShowItemRequest(xFormNo, xNo)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BExpandRefItem)
    '**     展開
    '**
    '*****************************************************************
    Protected Sub BExpandRefItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BExpandRefItem.Click
        Dim SQL As String = ""
        Dim xFormNo As String = ""
        Dim xNo As String = ""
        uErr = False
        '
        If InStr(DWaitRequest.SelectedItem.Text, "ZIP") > 0 Then
            xFormNo = "001152"
        Else
            If InStr(DWaitRequest.SelectedItem.Text, "SLDF") > 0 Then
                xFormNo = "001155"
            Else
                If InStr(DWaitRequest.SelectedItem.Text, "CH ") > 0 Then
                    xFormNo = "001154"
                Else
                    xFormNo = "001153"
                End If
            End If
        End If
        xNo = DWaitRequest.SelectedValue
        '
        Dim xChainClassSizeSldFun, xSlider1, xSlider2, xFinish1, xFinish2, xTape, xOther1 As String
        Dim xItemName As String()
        Dim xSpecialRequest As String()
        Dim xPartType, xSpecial, xStr, xName, xCode As String
        Dim k As Integer = 0
        '
        If DItemName.Text <> "" Then
            SQL = "Select * "
            SQL = SQL & "From W_ItemRequest "
            SQL = SQL + "Where FormNo = '" & xFormNo & "' "
            SQL = SQL + "  And No = '" & xNo & "' "
            If UCase(DItemName.Text) <> "ALL" Then
                SQL = SQL & "  And IT1I16 + ' ' + IT2I16 + ' ' + IT3I16 Like '%" & DItemName.Text & "%' "
            End If
            SQL = SQL & "Order by IT1I16, IT2I16, IT3I16 "
            Dim dt_WaitRequest As DataTable = uDataBase.GetDataTable(SQL)
            For i As Integer = 0 To dt_WaitRequest.Rows.Count - 1
                'Default Value
                xChainClassSizeSldFun = ""
                xSlider1 = ""
                xSlider2 = ""
                xFinish1 = ""
                xFinish2 = ""
                xTape = ""
                xOther1 = ""
                xPartType = ""
                '
                '****  解析ItemName1 ********************************************************************************************* 
                xItemName = dt_WaitRequest.Rows(i)("IT1I16").ToString.Trim.Split(" ")
                '
                '製品Type
                If xItemName(0).Length > 2 Then
                    xPartType = "ZIP"
                Else
                    If InStr(dt_WaitRequest.Rows(i)("IT1I16").ToString.Trim, "PARTS") > 0 Then
                        xPartType = "SLIDER"
                    Else
                        If InStr(dt_WaitRequest.Rows(i)("IT1I16").ToString.Trim, "CHAIN") > 0 Then
                            xPartType = "CHAIN"
                        Else
                            If InStr(dt_WaitRequest.Rows(i)("IT1I16").ToString.Trim, "TAPE") > 0 Then
                                xPartType = "TAPE"
                            Else
                                xPartType = "PULLER"
                            End If
                        End If
                    End If
                End If
                '
                'ZIP------------------------------------------
                If xPartType = "ZIP" Then
                    k = 0
                    If InStr(xItemName(0), "*") > 0 And InStr(dt_WaitRequest.Rows(i)("IT1I16").ToString.Trim, "/") > 0 Then
                        '雙拉頭
                        '   CFML-3* DSBN*73FM V7/DAV3LH6 V7 P12
                        '   1111111 222222222 3333333333 44 555
                        For j As Integer = 0 To xItemName.Length - 1
                            Select Case k
                                Case 0
                                    xChainClassSizeSldFun = xItemName(j)
                                Case 1
                                    xSlider1 = xItemName(j)
                                Case 2
                                    xFinish1 = Mid(xItemName(j), 1, InStr(xItemName(j), "/") - 1)
                                    xSlider2 = Mid(xItemName(j), InStr(xItemName(j), "/") + 1)
                                Case 3
                                    xFinish2 = xItemName(j)
                                Case Else
                                    xTape = xItemName(j)
                            End Select
                            k = k + 1
                        Next
                    Else
                        '單拉頭
                        '   CFC-39 DSBLUY09BK H6 P12
                        '   111111 2222222222 33 444
                        For j As Integer = 0 To xItemName.Length - 1
                            Select Case k
                                Case 0
                                    xChainClassSizeSldFun = xItemName(j)
                                Case 1
                                    xSlider1 = xItemName(j)
                                Case 2
                                    xFinish1 = xItemName(j)
                                Case Else
                                    xTape = xItemName(j)
                            End Select
                            k = k + 1
                        Next
                    End If
                Else
                    '
                    'SLIDER-----------------------------------------
                    If xPartType = "SLIDER" Then
                        '   05 CN DFYON5 N PARTS                
                        '   11 22 333333 4 55555
                        xChainClassSizeSldFun = xItemName(0) + " " + xItemName(1)
                        xSlider1 = xItemName(2)
                        xFinish1 = xItemName(3)
                    Else
                        '
                        'CHAIN------------------------------------------
                        If xPartType = "CHAIN" Then
                            '   05 CNL CHAIN PF16 NAT              
                            '   11 222 33333 4444 555
                            xChainClassSizeSldFun = xItemName(0) + " " + xItemName(1)
                            xTape = xItemName(3)
                            xOther1 = xItemName(4)
                        Else
                            '
                            'TAPE------------------------------------------
                            If xPartType = "TAPE" Then
                                '   05 CN PF16 TAPE NAT             
                                '   11 22 3333 4444 555
                                xChainClassSizeSldFun = xItemName(0) + " " + xItemName(1)
                                xTape = xItemName(2)
                                xOther1 = xItemName(4)
                            Else
                                '
                                'PULLER------------------------------------------
                                '   05 CN DATSS04-370 C5
                                '   11 22 33333       4444
                                xChainClassSizeSldFun = xItemName(0) + " " + xItemName(1)
                                xSlider1 = xItemName(2)
                                xFinish1 = xItemName(3)
                            End If
                        End If
                    End If
                End If
                '
                '****  解析ItemName2 and ItemName3 ********************************************************************************************* 
                '
                xSpecial = ""
                xItemName = (dt_WaitRequest.Rows(i)("IT2I16").ToString.Trim + " " + dt_WaitRequest.Rows(i)("IT3I16").ToString.Trim).Split(" ")
                For j As Integer = 0 To xItemName.Length - 1
                    If j = 0 Then
                        If CheckData(xItemName(j)) = True Then
                            xTape = xTape + xItemName(j)
                        Else
                            If xItemName(j) <> "" Then
                                If xSpecial = "" Then
                                    xSpecial = xItemName(j)
                                Else
                                    xSpecial = xSpecial + " " + xItemName(j)
                                End If
                            End If
                        End If
                    Else
                        If xItemName(j) <> "" Then
                            If xSpecial = "" Then
                                xSpecial = xItemName(j)
                            Else
                                xSpecial = xSpecial + " " + xItemName(j)
                            End If
                        End If
                    End If
                Next
                xSpecialRequest = xSpecial.Split(" ")
                '
                '尋找參考ITEM----------------------------------------------------------------------
                'CFML-3* DSBN*73FM V7/DAV3LH6 V7 P12
                'CFC-39 DSBLUY09BK H6 P12
                SQL = "SELECT Top 1 * From MST_C0100 "
                SQL = SQL + "WHERE IT1I01 Like '%" & dt_WaitRequest.Rows(i)("IT1I16").ToString.Trim & "%' "
                Dim dtITEM As DataTable = uDataBase.GetDataTable(SQL)
                If dtITEM.Rows.Count > 0 Then
                    xStr = dt_WaitRequest.Rows(i)("IT1I16").ToString.Trim
                    xCode = GetRefItem(xStr, xSpecialRequest)
                Else
                    'ZIP------------------------------------------
                    If xPartType = "ZIP" Then
                        xStr = xChainClassSizeSldFun + " " + xSlider1
                        xName = xChainClassSizeSldFun + " " + GetRefItemName("SLDC01", xStr, xFinish1, "")
                        '
                        xStr = xName + " " + xFinish1
                        xName = xName + " " + GetRefItemName("SFNC01", xStr, xFinish1, "")
                        '
                        If xSlider2 <> "" Then
                            xStr = xName + "/" + xSlider2
                            xName = xName + "/" + GetRefItemName("SL2C01", xStr, xFinish1, "")
                            '
                            xStr = xName + " " + xFinish2
                            xName = xName + " " + GetRefItemName("SE2C01", xStr, xFinish1, "")
                        End If
                        xStr = xName + " " + xTape
                        xName = xName + " " + GetRefItemName("TAPC01", xStr, xFinish1, "")
                        '
                        xCode = GetRefItem(xName, xSpecialRequest)
                    Else
                        'SLIDER-----------------------------------------
                        If xPartType = "SLIDER" Then
                            '   05 CN DFYON5 N PARTS               
                            xStr = xChainClassSizeSldFun + " " + xSlider1
                            xName = xChainClassSizeSldFun + " " + GetRefItemName("SLDC01", xStr, xFinish1, "PARTS")
                            'MsgBox("NAME1=[" + xName + "]")
                            '
                            xStr = xName + " " + xFinish1
                            xName = xName + " " + GetRefItemName("SFNC01", xStr, xFinish1, "PARTS")
                            'MsgBox("NAME2=[" + xName + "]")
                            '
                            xStr = xName + " " + "PARTS"
                            xName = GetRefItemName("PARTS", xStr, xFinish1, "PARTS")
                            'MsgBox("NAME3=[" + xName + "]")
                            '
                            xCode = GetRefItem(xName, xSpecialRequest)
                        Else
                            'CHAIN------------------------------------------
                            If xPartType = "CHAIN" Then
                                '05 CNL CHAIN PF16 NAT
                                xStr = xChainClassSizeSldFun + " " + "CHAIN" + " " + xTape
                                xName = xChainClassSizeSldFun + " " + "CHAIN" + " " + GetRefItemName("TAPC01", xStr, "", "CHAIN")
                                '
                                xStr = xName + " " + xOther1
                                xName = GetRefItemName("NAT", xStr, "", "CHAIN")
                                '
                                xCode = GetRefItem(xName, xSpecialRequest)
                            Else
                                'TAPE------------------------------------------
                                If xPartType = "TAPE" Then
                                    '   05 CN PF16 TAPE NAT             
                                    xStr = xChainClassSizeSldFun + " " + xTape
                                    xName = xChainClassSizeSldFun + " " + GetRefItemName("TAPC01", xStr, "", "TAPE")
                                    '
                                    xStr = xName + " " + "TAPE" + " " + xOther1
                                    xName = GetRefItemName("NAT", xStr, "", "TAPE")
                                    '
                                    xCode = GetRefItem(xName, xSpecialRequest)
                                Else
                                    'PULLER------------------------------------------
                                    If xPartType = "PULLER" Then
                                        '   05 CN DATSS04-370 C5
                                        xStr = xChainClassSizeSldFun + " " + xSlider1
                                        xName = xChainClassSizeSldFun + " " + GetRefItemName("SLDC01", xStr, xFinish1, "-")
                                        '
                                        xStr = xName + " " + xFinish1
                                        xName = xName + " " + GetRefItemName("SFNC01", xStr, xFinish1, "-")
                                        '
                                        xCode = GetRefItem(xName, xSpecialRequest)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                '
                '更新ItemRequest--------------------------------------------------------------------
                xStr = xChainClassSizeSldFun + "!" + _
                       xSlider1 + "!" + _
                       xSlider2 + "!" + _
                       xFinish1 + "!" + _
                       xFinish2 + "!" + _
                       xTape

                If xCode = "" Then uErr = True
                '
                UpdateItemRequest(dt_WaitRequest.Rows(i)("Unique_ID"), xCode, xStr, xSpecialRequest)
            Next
            '
            If uErr Then DSts.SelectedIndex = 2
            ShowItemRequest(xFormNo, xNo)
            '
            If uErr Then uJavaScript.PopMsg(Me, "有搜尋不到的參考ITEM,請確認 ! ")
            '
        Else
            uJavaScript.PopMsg(Me, "需輸入Item Name才能展開,請確認 ! ")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetRefItemName)
    '**     取得參考ITEM-NAME
    '**
    '*****************************************************************
    Public Function GetRefItemName(ByVal pField As String, ByVal pStr As String, ByVal pFinish As String, ByVal pPartType As String) As String
        Dim RtnString As String = ""
        Dim sql As String = ""
        Dim xStr As String = pStr + " "
        Dim i, xFinish As Integer
        '
        '03 C DSPO061 H3 PARTS
        '       03 C DSPO061 H6 PARTS
        i = 1
        Do While i < Len(pStr)
            '
            sql = "SELECT * From MST_C0100 "
            sql = sql + "WHERE IT1I01 Like '%" & xStr & "%' "
            sql = sql + "  And IT1I01 Like '%" & pPartType & "%' "
            sql = sql + "ORDER BY IT1I01, IT2I01, IT3I01 "
            Dim dtITEM As DataTable = uDataBase.GetDataTable(sql)
            If dtITEM.Rows.Count > 0 Then
                xFinish = 0
                For j As Integer = 0 To dtITEM.Rows.Count - 1
                    i = 999
                    '
                    If dtITEM.Rows(j)("SFNC01").ToString.Trim = pFinish Then
                        If pField = "PARTS" Or pField = "NAT" Then
                            RtnString = xStr
                        Else
                            RtnString = dtITEM.Rows(j)(pField).ToString.Trim
                            'SLDC01 and 引手
                            If pField = "SLDC01" And dtITEM.Rows(j)("CLSC01").ToString.Trim = "SP" Then
                                RtnString = dtITEM.Rows(j)(pField).ToString.Trim + "-" + dtITEM.Rows(j)("TAPC01").ToString.Trim
                            End If
                        End If
                        xFinish = 3
                        Exit For
                    Else
                        If InStr(dtITEM.Rows(j)("SFNC01").ToString.Trim, Mid(pFinish, 1, 2)) > 0 Or _
                           InStr(dtITEM.Rows(j)("SFNC01").ToString.Trim, Mid(pFinish, 2, 2)) > 0 Then
                            If xFinish <= 2 Then
                                If pField = "PARTS" Or pField = "NAT" Then
                                    RtnString = xStr
                                Else
                                    RtnString = dtITEM.Rows(j)(pField).ToString.Trim
                                    'SLDC01 and 引手
                                    If pField = "SLDC01" And dtITEM.Rows(j)("CLSC01").ToString.Trim = "SP" Then
                                        RtnString = dtITEM.Rows(j)(pField).ToString.Trim + "-" + dtITEM.Rows(j)("TAPC01").ToString.Trim
                                    End If
                                End If
                                xFinish = 2
                            End If
                        Else
                            If InStr(dtITEM.Rows(j)("SFNC01").ToString.Trim, Mid(pFinish, 1, 1)) > 0 Or _
                               InStr(dtITEM.Rows(j)("SFNC01").ToString.Trim, Mid(pFinish, 2, 1)) > 0 Or _
                               InStr(dtITEM.Rows(j)("SFNC01").ToString.Trim, Mid(pFinish, 3, 1)) > 0 Then
                                If xFinish <= 1 Then
                                    If pField = "PARTS" Or pField = "NAT" Then
                                        RtnString = xStr
                                    Else
                                        RtnString = dtITEM.Rows(j)(pField).ToString.Trim
                                        'SLDC01 and 引手
                                        If pField = "SLDC01" And dtITEM.Rows(j)("CLSC01").ToString.Trim = "SP" Then
                                            RtnString = dtITEM.Rows(j)(pField).ToString.Trim + "-" + dtITEM.Rows(j)("TAPC01").ToString.Trim
                                        End If
                                    End If
                                    xFinish = 1
                                End If
                            Else
                                If xFinish <= 0 Then
                                    If pField = "PARTS" Or pField = "NAT" Then
                                        RtnString = xStr
                                    Else
                                        RtnString = dtITEM.Rows(j)(pField).ToString.Trim
                                        'SLDC01 and 引手
                                        If pField = "SLDC01" And dtITEM.Rows(j)("CLSC01").ToString.Trim = "SP" Then
                                            RtnString = dtITEM.Rows(j)(pField).ToString.Trim + "-" + dtITEM.Rows(j)("TAPC01").ToString.Trim
                                        End If
                                    End If
                                    xFinish = 0
                                End If
                            End If
                        End If
                    End If
                Next
            Else
                xStr = Mid(xStr, 1, Len(xStr) - 1)
            End If
        Loop
        '
        Return RtnString
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetRefItem)
    '**     取得參考ITEM
    '**
    '*****************************************************************
    Public Function GetRefItem(ByVal pI1 As String, ByVal pSpecial As String()) As String
        Dim RtnString As String = ""
        Dim xCnt1, xCnt2 As Integer
        Dim sql As String = ""
        '
        sql = "SELECT * From MST_C0100 "
        sql = sql + "WHERE IT1I01+IT2I01 Like '%" & pI1 & "%' "
        sql = sql + "Order by IT2I01, IT3I01 "
        Dim dtITEM As DataTable = uDataBase.GetDataTable(sql)
        xCnt1 = 0
        For i As Integer = 0 To dtITEM.Rows.Count - 1
            xCnt2 = 0
            For j As Integer = 0 To pSpecial.Length - 1
                If InStr(dtITEM.Rows(i)("IT2I01").ToString.Trim + dtITEM.Rows(i)("IT3I01").ToString.Trim, pSpecial(j)) > 0 Then
                    xCnt2 = xCnt2 + 1
                End If
            Next
            If xCnt2 >= xCnt1 Then
                xCnt1 = xCnt2
                RtnString = dtITEM.Rows(i)("ITMC01").ToString.Trim
            End If
        Next
        '
        Return RtnString
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(UpdateItemRequest)
    '**     更新ItemRequest
    '**
    '*****************************************************************
    Protected Sub UpdateItemRequest(ByVal pID As Integer, ByVal pCode As String, ByVal pItemName As String, ByVal pSpecial As String())
        Dim sql As String = ""
        Dim xItemName As String() = pItemName.Split("!")
        '
        Try
            If pCode <> "" Then
                sql = "SELECT * From MST_C0100 "
                sql = sql + "WHERE ITMC01 = '" & pCode & "' "
                Dim dtITEM As DataTable = uDataBase.GetDataTable(sql)
                If dtITEM.Rows.Count > 0 Then
                    sql = "Update W_ItemRequest Set "
                    sql &= " STS = '" & "0" & "',"                  'Status(0:正常/1:異常)
                    sql &= " SEMI16 = '" & "" & "',"                'EMPLOYEE NAME (SECOND LANG  
                    sql &= " SSNN16 = '" & "" & "',"                'SESSION NO.

                    sql &= " YNXC16 = '" & dtITEM.Rows(0)("YNXC01").ToString.Trim & "',"                'YNX PRODUCT CODE
                    sql &= " ST1C16 = '" & dtITEM.Rows(0)("ST1C01").ToString.Trim & "',"                  'STATISTICS CODE 1 
                    sql &= " ST2C16 = '" & dtITEM.Rows(0)("ST2C01").ToString.Trim & "',"                  'STATISTICS CODE 2
                    sql &= " ST3C16 = '" & dtITEM.Rows(0)("ST3C01").ToString.Trim & "',"                  'STATISTICS CODE 3
                    sql &= " ST4C16 = '" & dtITEM.Rows(0)("ST4C01").ToString.Trim & "',"                  'STATISTICS CODE 4
                    '
                    sql &= " ST5C16 = '" & dtITEM.Rows(0)("ST5C01").ToString.Trim & "',"                  'STATISTICS CODE 5
                    sql &= " ST6C16 = '" & dtITEM.Rows(0)("ST6C01").ToString.Trim & "',"                  'STATISTICS CODE 6
                    sql &= " ST7C16 = '" & dtITEM.Rows(0)("ST7C01").ToString.Trim & "',"                  'STATISTICS CODE 7
                    sql &= " SIZC16 = '" & dtITEM.Rows(0)("SIZC01").ToString.Trim & "',"                    'SIZE 
                    sql &= " CHNC16 = '" & dtITEM.Rows(0)("CHNC01").ToString.Trim & "',"               'CHAIN CODE
                    '
                    sql &= " CLSC16 = '" & dtITEM.Rows(0)("CLSC01").ToString.Trim & "',"                   'CLASSIFICATION CODE 
                    sql &= " SIMF16 = '" & dtITEM.Rows(0)("SIMF01").ToString.Trim & "',"                                     'SET ITEM FLAG
                    '
                    If dtITEM.Rows(0)("CLSC01").ToString.Trim = "SP" Then
                        sql &= " SLDC16 = '" & Mid(xItemName(1).ToString, 1, InStr(xItemName(1).ToString, "-") - 1) & "'," 'SLIDER CODE
                        sql &= " SL2C16 = '" & xItemName(2).ToString & "'," 'SLIDER CODE 2
                        sql &= " SFNC16 = '" & xItemName(3).ToString & "'," 'SLIDER FINISH CODE
                        sql &= " SE2C16 = '" & xItemName(4).ToString & "'," 'SLIDER FINISH CODE 2
                        sql &= " TAPC16 = '" & Mid(xItemName(1).ToString, InStr(xItemName(1).ToString, "-") + 1) & "'," 'TAPE CODE
                    Else
                        sql &= " SLDC16 = '" & xItemName(1).ToString & "'," 'SLIDER CODE
                        sql &= " SL2C16 = '" & xItemName(2).ToString & "'," 'SLIDER CODE 2
                        sql &= " SFNC16 = '" & xItemName(3).ToString & "'," 'SLIDER FINISH CODE
                        sql &= " SE2C16 = '" & xItemName(4).ToString & "'," 'SLIDER FINISH CODE 2
                        sql &= " TAPC16 = '" & xItemName(5).ToString & "'," 'TAPE CODE
                    End If
                    '
                    Select Case pSpecial.Length - 1
                        Case 0
                            sql &= " SF1C16 = '" & pSpecial(0) & "',"                  'SPECIAL FEATURE CODE 1
                        Case 1
                            sql &= " SF1C16 = '" & pSpecial(0) & "',"                  'SPECIAL FEATURE CODE 1
                            sql &= " SF2C16 = '" & pSpecial(1) & "',"                  'SPECIAL FEATURE CODE 2
                        Case 2
                            sql &= " SF1C16 = '" & pSpecial(0) & "',"                  'SPECIAL FEATURE CODE 1
                            sql &= " SF2C16 = '" & pSpecial(1) & "',"                  'SPECIAL FEATURE CODE 2
                            sql &= " SF3C16 = '" & pSpecial(2) & "',"                  'SPECIAL FEATURE CODE 3
                        Case 3
                            sql &= " SF1C16 = '" & pSpecial(0) & "',"                  'SPECIAL FEATURE CODE 1
                            sql &= " SF2C16 = '" & pSpecial(1) & "',"                  'SPECIAL FEATURE CODE 2
                            sql &= " SF3C16 = '" & pSpecial(2) & "',"                  'SPECIAL FEATURE CODE 3
                            sql &= " SF4C16 = '" & pSpecial(3) & "',"                  'SPECIAL FEATURE CODE 4
                        Case 4
                            sql &= " SF1C16 = '" & pSpecial(0) & "',"                  'SPECIAL FEATURE CODE 1
                            sql &= " SF2C16 = '" & pSpecial(1) & "',"                  'SPECIAL FEATURE CODE 2
                            sql &= " SF3C16 = '" & pSpecial(2) & "',"                  'SPECIAL FEATURE CODE 3
                            sql &= " SF4C16 = '" & pSpecial(3) & "',"                  'SPECIAL FEATURE CODE 4
                            sql &= " SF5C16 = '" & pSpecial(4) & "',"                  'SPECIAL FEATURE CODE 5
                        Case Else
                            sql &= " SF1C16 = '" & pSpecial(0) & "',"                  'SPECIAL FEATURE CODE 1
                            sql &= " SF2C16 = '" & pSpecial(1) & "',"                  'SPECIAL FEATURE CODE 2
                            sql &= " SF3C16 = '" & pSpecial(2) & "',"                  'SPECIAL FEATURE CODE 3
                            sql &= " SF4C16 = '" & pSpecial(3) & "',"                  'SPECIAL FEATURE CODE 4
                            sql &= " SF5C16 = '" & pSpecial(4) & "',"                  'SPECIAL FEATURE CODE 5
                            sql &= " SF6C16 = '" & pSpecial(5) & "',"                  'SPECIAL FEATURE CODE 6
                    End Select
                    '
                    sql &= " FMLC16 = '" & dtITEM.Rows(0)("FMLC01").ToString.Trim & "',"    'FAMILY CODE
                    sql &= " QUNC16 = '" & dtITEM.Rows(0)("QUNC01").ToString.Trim & "',"    'QUANTITY UNIT CODE
                    '
                    sql &= " PO1C16 = '" & "" & "',"                  'PLACE OF ORIGIN CODE 1 
                    sql &= " PO2C16 = '" & "" & "',"                  'PLACE OF ORIGIN CODE 2 
                    sql &= " PO3C16 = '" & "" & "',"                  'PLACE OF ORIGIN CODE 3
                    sql &= " PO4C16 = '" & "" & "',"                  'PLACE OF ORIGIN CODE 4
                    sql &= " PO5C16 = '" & "" & "',"                  'PLACE OF ORIGIN CODE 5
                    '
                    sql &= " PO6C16 = '" & "" & "',"                  'PLACE OF ORIGIN CODE 6
                    sql &= " PO7C16 = '" & "" & "',"                  'PLACE OF ORIGIN CODE 7
                    sql &= " ICRX16 = '" & "" & "',"                  'ITEM CODE REQUEST COMMENT
                    sql &= " RFCC16 = '" & "000034" & "',"            'COMPANY CODE(REQUEST FROM)
                    sql &= " RTCC16 = '" & "000100" & "',"            'COMPANY CODE(REQUEST TO)
                    '
                    sql &= " ITMC16 = '" & "" & "',"                  'ITEM CODE
                    'sql &= " ICQC16 = '" & dtITEM.Rows(0)("ICQC01").ToString.Trim & "',"                  'ITEM CODE REQUEST CODE
                    sql &= " ICQC16 = '" & "1" & "',"                                                     'ITEM CODE REQUEST CODE
                    sql &= " MAIC16 = '" & dtITEM.Rows(0)("MAIC01").ToString.Trim & "',"                  'ITEM CODE(ALTERNATIVE ITEM)
                    sql &= " MPTC16 = '" & dtITEM.Rows(0)("MPTC01").ToString.Trim & "',"                  'MACHINE PARTS TYPE CODE
                    '------------------------------
                    sql &= " RSEMI16 = '" & "" & "',"                 'EMPLOYEE NAME (SECOND LANG  
                    sql &= " RSSNN16 = '" & "" & "',"                 'SESSION NO.
                    sql &= " RIT1I16 = '" & dtITEM.Rows(0)("IT1I01").ToString.Trim & "',"               'ITEM NAME 1
                    sql &= " RIT2I16 = '" & dtITEM.Rows(0)("IT2I01").ToString.Trim & "',"               'ITEM NAME 2
                    sql &= " RIT3I16 = '" & dtITEM.Rows(0)("IT3I01").ToString.Trim & "',"               'ITEM NAME 3
                    '
                    sql &= " RYNXC16 = '" & dtITEM.Rows(0)("YNXC01").ToString.Trim & "',"                 'YNX PRODUCT CODE
                    sql &= " RST1C16 = '" & dtITEM.Rows(0)("ST1C01").ToString.Trim & "',"                  'STATISTICS CODE 1 
                    sql &= " RST2C16 = '" & dtITEM.Rows(0)("ST2C01").ToString.Trim & "',"                  'STATISTICS CODE 2
                    sql &= " RST3C16 = '" & dtITEM.Rows(0)("ST3C01").ToString.Trim & "',"                  'STATISTICS CODE 3
                    sql &= " RST4C16 = '" & dtITEM.Rows(0)("ST4C01").ToString.Trim & "',"                  'STATISTICS CODE 4
                    '
                    sql &= " RST5C16 = '" & dtITEM.Rows(0)("ST5C01").ToString.Trim & "',"                  'STATISTICS CODE 5
                    sql &= " RST6C16 = '" & dtITEM.Rows(0)("ST6C01").ToString.Trim & "',"                  'STATISTICS CODE 6
                    sql &= " RST7C16 = '" & dtITEM.Rows(0)("ST7C01").ToString.Trim & "',"                  'STATISTICS CODE 7
                    sql &= " RSIZC16 = '" & dtITEM.Rows(0)("SIZC01").ToString.Trim & "',"                    'SIZE 
                    sql &= " RCHNC16 = '" & dtITEM.Rows(0)("CHNC01").ToString.Trim & "',"               'CHAIN CODE
                    '
                    sql &= " RCLSC16 = '" & dtITEM.Rows(0)("CLSC01").ToString.Trim & "',"                   'CLASSIFICATION CODE 
                    sql &= " RTAPC16 = '" & dtITEM.Rows(0)("TAPC01").ToString.Trim & "',"                   'TAPE CODE
                    sql &= " RSIMF16 = '" & dtITEM.Rows(0)("SIMF01").ToString.Trim & "',"                                                 'SET ITEM FLAG
                    sql &= " RSLDC16 = '" & dtITEM.Rows(0)("SLDC01").ToString.Trim & "',"                  'SLIDER CODE
                    sql &= " RSFNC16 = '" & dtITEM.Rows(0)("SFNC01").ToString.Trim & "',"                  'SLIDER FINISH CODE
                    '   
                    sql &= " RSL2C16 = '" & dtITEM.Rows(0)("SL2C01").ToString.Trim & "',"                 'SLIDER CODE 2
                    sql &= " RSE2C16 = '" & dtITEM.Rows(0)("SE2C01").ToString.Trim & "',"                 'SLIDER FINISH CODE 2
                    sql &= " RSF1C16 = '" & dtITEM.Rows(0)("SF1C01").ToString.Trim & "',"        'SPECIAL FEATURE CODE 1
                    sql &= " RSF2C16 = '" & dtITEM.Rows(0)("SF2C01").ToString.Trim & "',"        'SPECIAL FEATURE CODE 2
                    sql &= " RSF3C16 = '" & dtITEM.Rows(0)("SF3C01").ToString.Trim & "',"        'SPECIAL FEATURE CODE 3
                    '
                    sql &= " RSF4C16 = '" & dtITEM.Rows(0)("SF4C01").ToString.Trim & "',"        'SPECIAL FEATURE CODE 4
                    sql &= " RSF5C16 = '" & dtITEM.Rows(0)("SF5C01").ToString.Trim & "',"        'SPECIAL FEATURE CODE 5
                    sql &= " RSF6C16 = '" & dtITEM.Rows(0)("SF6C01").ToString.Trim & "',"        'SPECIAL FEATURE CODE 6
                    sql &= " RFMLC16 = '" & dtITEM.Rows(0)("FMLC01").ToString.Trim & "',"    'FAMILY CODE
                    sql &= " RQUNC16 = '" & dtITEM.Rows(0)("QUNC01").ToString.Trim & "',"    'QUANTITY UNIT CODE
                    '
                    sql &= " RPO1C16 = '" & "" & "',"                  'PLACE OF ORIGIN CODE 1 
                    sql &= " RPO2C16 = '" & "" & "',"                  'PLACE OF ORIGIN CODE 2 
                    sql &= " RPO3C16 = '" & "" & "',"                  'PLACE OF ORIGIN CODE 3
                    sql &= " RPO4C16 = '" & "" & "',"                  'PLACE OF ORIGIN CODE 4
                    sql &= " RPO5C16 = '" & "" & "',"                  'PLACE OF ORIGIN CODE 5
                    '
                    sql &= " RPO6C16 = '" & "" & "',"                  'PLACE OF ORIGIN CODE 6
                    sql &= " RPO7C16 = '" & "" & "',"                  'PLACE OF ORIGIN CODE 7
                    sql &= " RICRX16 = '" & "" & "',"                  'ITEM CODE REQUEST COMMENT
                    sql &= " RRFCC16 = '" & "000034" & "',"            'COMPANY CODE(REQUEST FROM)
                    sql &= " RRTCC16 = '" & "000100" & "',"            'COMPANY CODE(REQUEST TO)
                    '
                    sql &= " RITMC16 = '" & dtITEM.Rows(0)("ITMC01").ToString.Trim & "',"                  'ITEM CODE
                    sql &= " RICQC16 = '" & dtITEM.Rows(0)("ICQC01").ToString.Trim & "',"                  'ITEM CODE REQUEST CODE
                    sql &= " RMAIC16 = '" & dtITEM.Rows(0)("MAIC01").ToString.Trim & "',"                  'ITEM CODE(ALTERNATIVE ITEM)
                    sql &= " RMPTC16 = '" & dtITEM.Rows(0)("MPTC01").ToString.Trim & "',"                  'MACHINE PARTS TYPE CODE
                    '------------------------------
                    sql &= " ModifyUser = '" & "ItemRequest01" & "',"
                    sql &= " ModifyTime = '" & NowDateTime & "' "
                    sql &= " Where Unique_ID  =  '" & CStr(pID) & "'"
                    '
                    uDataBase.ExecuteNonQuery(sql)
                End If
            Else
                sql = "Update W_ItemRequest Set "
                sql &= " STS = '" & "1" & "',"                  'Status(0:正常/1:異常)
                sql &= " TAPC16 = '" & xItemName(5).ToString.Trim & "',"                  'TAPE CODE
                sql &= " SLDC16 = '" & xItemName(1).ToString.Trim & "',"                  'SLIDER CODE
                sql &= " SFNC16 = '" & xItemName(3).ToString.Trim & "',"                  'SLIDER FINISH CODE
                sql &= " SL2C16 = '" & xItemName(2).ToString.Trim & "',"                  'SLIDER CODE 2
                sql &= " SE2C16 = '" & xItemName(4).ToString.Trim & "',"                  'SLIDER FINISH CODE 2
                Select Case pSpecial.Length - 1
                    Case 0
                        sql &= " SF1C16 = '" & pSpecial(0) & "',"                  'SPECIAL FEATURE CODE 1
                    Case 1
                        sql &= " SF1C16 = '" & pSpecial(0) & "',"                  'SPECIAL FEATURE CODE 1
                        sql &= " SF2C16 = '" & pSpecial(1) & "',"                  'SPECIAL FEATURE CODE 2
                    Case 2
                        sql &= " SF1C16 = '" & pSpecial(0) & "',"                  'SPECIAL FEATURE CODE 1
                        sql &= " SF2C16 = '" & pSpecial(1) & "',"                  'SPECIAL FEATURE CODE 2
                        sql &= " SF3C16 = '" & pSpecial(2) & "',"                  'SPECIAL FEATURE CODE 3
                    Case 3
                        sql &= " SF1C16 = '" & pSpecial(0) & "',"                  'SPECIAL FEATURE CODE 1
                        sql &= " SF2C16 = '" & pSpecial(1) & "',"                  'SPECIAL FEATURE CODE 2
                        sql &= " SF3C16 = '" & pSpecial(2) & "',"                  'SPECIAL FEATURE CODE 3
                        sql &= " SF4C16 = '" & pSpecial(3) & "',"                  'SPECIAL FEATURE CODE 4
                    Case 4
                        sql &= " SF1C16 = '" & pSpecial(0) & "',"                  'SPECIAL FEATURE CODE 1
                        sql &= " SF2C16 = '" & pSpecial(1) & "',"                  'SPECIAL FEATURE CODE 2
                        sql &= " SF3C16 = '" & pSpecial(2) & "',"                  'SPECIAL FEATURE CODE 3
                        sql &= " SF4C16 = '" & pSpecial(3) & "',"                  'SPECIAL FEATURE CODE 4
                        sql &= " SF5C16 = '" & pSpecial(4) & "',"                  'SPECIAL FEATURE CODE 5
                    Case Else
                        sql &= " SF1C16 = '" & pSpecial(0) & "',"                  'SPECIAL FEATURE CODE 1
                        sql &= " SF2C16 = '" & pSpecial(1) & "',"                  'SPECIAL FEATURE CODE 2
                        sql &= " SF3C16 = '" & pSpecial(2) & "',"                  'SPECIAL FEATURE CODE 3
                        sql &= " SF4C16 = '" & pSpecial(3) & "',"                  'SPECIAL FEATURE CODE 4
                        sql &= " SF5C16 = '" & pSpecial(4) & "',"                  'SPECIAL FEATURE CODE 5
                        sql &= " SF6C16 = '" & pSpecial(5) & "',"                  'SPECIAL FEATURE CODE 6
                End Select
                '
                sql &= " ModifyUser = '" & "ItemRequest01" & "',"
                sql &= " ModifyTime = '" & NowDateTime & "' "
                sql &= " Where Unique_ID  =  '" & CStr(pID) & "'"
                '
                uDataBase.ExecuteNonQuery(sql)
            End If
        Catch ex As Exception
            sql = "Update W_ItemRequest Set "
            sql &= " STS = '" & "1" & "',"                  'Status(0:正常/1:異常)
            '
            sql &= " ModifyUser = '" & "ItemRequest01" & "',"
            sql &= " ModifyTime = '" & NowDateTime & "' "
            sql &= " Where Unique_ID  =  '" & CStr(pID) & "'"
            '
            uDataBase.ExecuteNonQuery(sql)
        End Try
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Upload)
    '**     上傳
    '**
    '*****************************************************************
    Protected Sub BUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BUpload.Click
        Dim SQL As String = ""
        Dim Upload_ID As String = Now.ToString("yyyyMMddHHmmss")
        Dim xFormNo As String = ""
        Dim xNo As String = ""
        '
        If InStr(DWaitRequest.SelectedItem.Text, "ZIP") > 0 Then
            xFormNo = "001152"
        Else
            If InStr(DWaitRequest.SelectedItem.Text, "SLDF") > 0 Then
                xFormNo = "001155"
            Else
                If InStr(DWaitRequest.SelectedItem.Text, "CH ") > 0 Then
                    xFormNo = "001154"
                Else
                    xFormNo = "001153"
                End If
            End If
        End If
        xNo = DWaitRequest.SelectedValue
        '
        SQL = "Select * From W_ItemRequest "
        SQL = SQL + "Where FormNo = '" & xFormNo & "' "
        SQL = SQL + "  And No = '" & xNo & "' "
        SQL = SQL + "  And Sts = '2' "
        Dim dt_WaitRequestx As DataTable = uDataBase.GetDataTable(SQL)
        If dt_WaitRequestx.Rows.Count <= 0 Then
            uJavaScript.PopMsg(Me, "[" + DWaitRequest.SelectedItem.Text + "]:無上傳資料,請確認 ! ")
        Else
            SQL = "Select * From W_ItemRequest "
            SQL = SQL + "Where FormNo = '" & xFormNo & "' "
            SQL = SQL + "  And No = '" & xNo & "' "
            SQL = SQL + "  And Sts = '2' "
            SQL = SQL + "Order by IT1I16, IT2I16, IT3I16 "
            Dim dt_WaitRequest As DataTable = uDataBase.GetDataTable(SQL)
            For i As Integer = 0 To dt_WaitRequest.Rows.Count - 1
                'Add Wave's WAVEALIB.TWM420
                Dim cn As New OleDbConnection
                Dim cmd As New OleDbCommand
                cn.ConnectionString = uCommon.GetAppSetting("WAVESOLEDB")
                cmd.Connection = cn
                cn.Open()
                '
                SQL = "INSERT INTO WAVEALIB.TWM420 "
                SQL &= "( "
                SQL &= "SEMI16, SSNN16, IT1I16, IT2I16, IT3I16, "
                SQL &= "YNXC16, ST1C16, ST2C16, ST3C16, ST4C16, "
                SQL &= "ST5C16, ST6C16, ST7C16, SIZC16, CHNC16, "
                SQL &= "CLSC16, TAPC16, SIMF16, SLDC16, SFNC16, "
                SQL &= "SL2C16, SE2C16, SF1C16, SF2C16, SF3C16, "
                SQL &= "SF4C16, SF5C16, SF6C16, FMLC16, QUNC16, "
                SQL &= "PO1C16, PO2C16, PO3C16, PO4C16, PO5C16, "
                SQL &= "PO6C16, PO7C16, ICRX16, RFCC16, RTCC16, "
                SQL &= "ITMC16, ICQC16, MAIC16, MPTC16, "
                SQL &= "UIDC16, PRGC16, DEVC16, RADU16, RADT16, "
                SQL &= "RUPU16, RUPT16, RCMC16 "
                SQL &= " ) "
                SQL &= "VALUES( "
                SQL &= "'" & dt_WaitRequest.Rows(i)("SEMI16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("SSNN16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("IT1I16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("IT2I16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("IT3I16").ToString.Trim & "' ,"
                '
                SQL &= "'" & dt_WaitRequest.Rows(i)("YNXC16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("ST1C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("ST2C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("ST3C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("ST4C16").ToString.Trim & "' ,"
                '
                SQL &= "'" & dt_WaitRequest.Rows(i)("ST5C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("ST6C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("ST7C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("SIZC16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("CHNC16").ToString.Trim & "' ,"
                '
                SQL &= "'" & dt_WaitRequest.Rows(i)("CLSC16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("TAPC16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("SIMF16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("SLDC16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("SFNC16").ToString.Trim & "' ,"
                '
                SQL &= "'" & dt_WaitRequest.Rows(i)("SL2C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("SE2C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("SF1C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("SF2C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("SF3C16").ToString.Trim & "' ,"
                '
                SQL &= "'" & dt_WaitRequest.Rows(i)("SF4C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("SF5C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("SF6C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("FMLC16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("QUNC16").ToString.Trim & "' ,"
                '
                SQL &= "'" & dt_WaitRequest.Rows(i)("PO1C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("PO2C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("PO3C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("PO4C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("PO5C16").ToString.Trim & "' ,"
                '
                SQL &= "'" & dt_WaitRequest.Rows(i)("PO6C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("PO7C16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("ICRX16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("RFCC16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("RTCC16").ToString.Trim & "' ,"
                '
                SQL &= "'" & dt_WaitRequest.Rows(i)("ITMC16").ToString.Trim & "' ,"
                'SQL &= "'" & dt_WaitRequest.Rows(i)("ICQC16").ToString.Trim & "' ,"
                SQL &= "'" & "1" & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("MAIC16").ToString.Trim & "' ,"
                SQL &= "'" & dt_WaitRequest.Rows(i)("MPTC16").ToString.Trim & "' ,"
                '----------------------------------------------------------------------
                SQL &= "'" & "IT004" & "' ,"
                SQL &= "'" & "ITEMREQ" & "' ,"
                SQL &= "'" & "IRW" & "' ,"
                SQL &= "'" & Now.ToString("yyyyMMdd") & "' ,"
                SQL &= "'" & Now.ToString("HHmmss") & "' ,"
                '
                SQL &= "'" & "0" & "' ,"
                SQL &= "'" & "0" & "' ,"
                SQL &= "'" & "000034" & "' "
                SQL &= " ) "
                cmd.CommandText = SQL
                cmd.ExecuteNonQuery()
                '
                'Add L_ItemRequest 
                SQL = "Insert Into L_ItemRequest "
                SQL = SQL & "Select "
                SQL = SQL & "'" & Upload_ID & "', "
                SQL = SQL & "Sts, FormNo, No, FNo, "
                '
                SQL = SQL & "SEMI16, SSNN16, IT1I16, IT2I16, IT3I16, "
                SQL = SQL & "YNXC16, ST1C16, ST2C16, ST3C16, ST4C16, "
                SQL = SQL & "ST5C16, ST6C16, ST7C16, SIZC16, CHNC16, "
                SQL = SQL & "CLSC16, TAPC16, SIMF16, SLDC16, SFNC16, "
                SQL = SQL & "SL2C16, SE2C16, SF1C16, SF2C16, SF3C16, "
                SQL = SQL & "SF4C16, SF5C16, SF6C16, FMLC16, QUNC16, "
                SQL = SQL & "PO1C16, PO2C16, PO3C16, PO4C16, PO5C16, "
                SQL = SQL & "PO6C16, PO7C16, ICRX16, RFCC16, RTCC16, "
                SQL = SQL & "ITMC16, ICQC16, MAIC16, MPTC16, "
                '
                SQL = SQL & "RSEMI16, RSSNN16, RIT1I16, RIT2I16, RIT3I16, "
                SQL = SQL & "RYNXC16, RST1C16, RST2C16, RST3C16, RST4C16, "
                SQL = SQL & "RST5C16, RST6C16, RST7C16, RSIZC16, RCHNC16, "
                SQL = SQL & "RCLSC16, RTAPC16, RSIMF16, RSLDC16, RSFNC16, "
                SQL = SQL & "RSL2C16, RSE2C16, RSF1C16, RSF2C16, RSF3C16, "
                SQL = SQL & "RSF4C16, RSF5C16, RSF6C16, RFMLC16, RQUNC16, "
                SQL = SQL & "RPO1C16, RPO2C16, RPO3C16, RPO4C16, RPO5C16, "
                SQL = SQL & "RPO6C16, RPO7C16, RICRX16, RRFCC16, RRTCC16, "
                SQL = SQL & "RITMC16, RICQC16, RMAIC16, RMPTC16, "
                '
                SQL = SQL & "CreateUser, CreateTime, ModifyUser, ModifyTime "
                SQL = SQL & "From W_ItemRequest "
                SQL = SQL & " Where Unique_ID = '" & dt_WaitRequest.Rows(i)("Unique_ID").ToString.Trim & "'"
                uDataBase.ExecuteNonQuery(SQL)
                '
                'Delete W_ItemRequest 
                SQL = "Delete From W_ItemRequest "
                SQL = SQL & " Where Unique_ID = '" & dt_WaitRequest.Rows(i)("Unique_ID").ToString.Trim & "'"
                uDataBase.ExecuteNonQuery(SQL)
            Next
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(CheckData)
    '**     檢查輸入資料
    '**
    '*****************************************************************
    '*********************************************************************
    Public Function CheckData(ByVal pData As String) As Boolean
        Dim uError As Boolean = False
        Dim SQL As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
        '
        If pData <> "" Then
            SQL = "SELECT CN1I09 FROM WAVEDLIB.C0900 "
            SQL = SQL + "WHERE DGRC09 = 'SF1C' AND DDTC09 = '" & pData & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, cn)
            DBAdapter1.Fill(ds, "C0900")
            If ds.Tables("C0900").Rows.Count <= 0 Then uError = True
        End If
        '
        Return uError
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowItemRequest)
    '**     顯示ItemRequest資料
    '**
    '*****************************************************************
    Protected Sub ShowItemRequest(ByVal pFormNo As String, ByVal pNo As String)
        Dim SQL As String
        '
        SQL = "Select "
        SQL = SQL & "'編輯' As Field1, "
        SQL = SQL + "'ItemRequest_02.aspx?' + "
        SQL = SQL + "'pID='   + LTrim(str(Unique_ID)) As UpdateURL, "
        SQL = SQL & "Case Sts When 1 Then 'Ｘ' When 2 Then 'Ｗ' Else 'Ｏ' End As StsD, "
        SQL = SQL & "IT1I16 + ' ' + IT2I16 + ' ' + IT3I16 As ItemName, "
        SQL = SQL & "RITMC16 as RCode, "
        SQL = SQL & "RIT1I16 + ' ' + RIT2I16 + ' ' + RIT3I16 As RItemName "
        SQL = SQL & "From W_ItemRequest "
        SQL = SQL & "Where FormNo = '" & pFormNo & "' "
        SQL = SQL & "  And No = '" & pNo & "' "
        If DSts.SelectedValue <> "ALL" Then
            SQL = SQL & "  And Sts = '" & DSts.SelectedValue & "' "
        End If
        If DItemName.Text <> "" Then
            SQL = SQL & "  And IT1I16 + ' ' + IT2I16 + ' ' + IT3I16 Like '%" & DItemName.Text & "%' "
        End If
        SQL = SQL & "Order by IT1I16, IT2I16, IT3I16 "
        Dim dt_WaitRequest As DataTable = uDataBase.GetDataTable(SQL)
        GridView1.DataSource = dt_WaitRequest
        GridView1.DataBind()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(RefreshItemRequest)
    '**     更新ItemRequest資料
    '**
    '*****************************************************************
    Protected Sub BRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BRefresh.Click
        Dim xFormNo As String = ""
        Dim xNo As String = ""
        '
        If InStr(DWaitRequest.SelectedItem.Text, "ZIP") > 0 Then
            xFormNo = "001152"
        Else
            If InStr(DWaitRequest.SelectedItem.Text, "SLDF") > 0 Then
                xFormNo = "001155"
            Else
                If InStr(DWaitRequest.SelectedItem.Text, "CH ") > 0 Then
                    xFormNo = "001154"
                Else
                    xFormNo = "001153"
                End If
            End If
        End If
        xNo = DWaitRequest.SelectedValue
        '
        ShowItemRequest(xFormNo, xNo)
    End Sub
End Class
