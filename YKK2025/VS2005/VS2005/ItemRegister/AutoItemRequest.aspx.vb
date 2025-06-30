Imports System.Data
Imports System.Data.OleDb

Partial Class AutoItemRequest
    Inherits System.Web.UI.Page

    'iexplore.exe "http://10.245.1.10/IRW/AutoItemRequest.aspx"
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim oCommon As New Common.CommonService
    Dim NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")
    Dim NowDate = Now.ToString("yyyy/MM/dd")
    '
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900      ' 設定逾時時間
        If Not IsPostBack Then
            DSProgressBar.Text = Now.ToString("yyyy/MM/dd HH:mm:ss") & "    自動配對處理中.........."
            SearchItem()
            DEProgressBar.Text = Now.ToString("yyyy/MM/dd HH:mm:ss") & "    自動配對處理完成........"
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SearchItem)
    '**     自動配對
    '**
    '*****************************************************************
    Sub SearchItem()
        Dim wCount As Integer = 0
        Dim wRCode As Integer = 0
        Dim SQL As String = ""
        '
        '查詢自動簽核暫存檔
        SQL = "Select Top 1 * "
        SQL = SQL & "From Q_WaitAutoApprove "
        'SQL = SQL & "Where FormNo in ('001152', '001153', '001154', '001155') "
        SQL = SQL & "Where FormNo = '001153' "
        SQL = SQL & "Order by FormNo "
        Dim dt_WaitAutoApprove As DataTable = uDataBase.GetDataTable(SQL)
        For i As Integer = 0 To dt_WaitAutoApprove.Rows.Count - 1
            '
            SQL = "Select No "
            SQL = SQL & "From V_WaitHandle_01 "
            SQL = SQL & "where FormNo = '" & dt_WaitAutoApprove.Rows(i)("FormNo").ToString.Trim & "' "
            SQL = SQL & "  and FormSno = '" & dt_WaitAutoApprove.Rows(i)("FormSno").ToString.Trim & "' "
            Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
            If dt_WaitHandle.Rows.Count > 0 Then
                '
                DataPump(dt_WaitAutoApprove.Rows(i)("FormNo").ToString.Trim, dt_WaitHandle.Rows(0)("No").ToString.Trim)
                '
                wRCode = DataExpand(dt_WaitAutoApprove.Rows(i)("FormNo").ToString.Trim, dt_WaitHandle.Rows(0)("No").ToString.Trim)
                '
                If wRCode = 0 Then
                    '自動簽核後刪除暫存table資料
                    SQL = "Select * "
                    SQL = SQL & "From W_ItemRequest "
                    SQL = SQL + "Where FormNo = '" & dt_WaitAutoApprove.Rows(i)("FormNo").ToString.Trim & "' "
                    SQL = SQL + "  And No = '" & dt_WaitHandle.Rows(0)("No").ToString.Trim & "' "
                    SQL = SQL + "  And Sts = '" & "0" & "' "
                    SQL = SQL + "  And RIT1I16 + RIT2I16 + RIT3I16 = '" & "" & "' "
                    Dim dt_WaitRequest As DataTable = uDataBase.GetDataTable(SQL)
                    If dt_WaitRequest.Rows.Count <= 0 Then
                        SQL = "delete  From Q_WaitAutoApprove "
                        SQL = SQL + " where Unique_ID = '" + dt_WaitAutoApprove.Rows(i)("Unique_ID").ToString.Trim + "' "
                        uDataBase.ExecuteNonQuery(SQL)
                    End If
                End If
            End If
            '
            wCount = wCount + 1
        Next
        '
        DCount.Text = "處理筆數 ( " & CStr(wCount) & " )"

        '限IE11
        Response.Write("<script>window.open('', '_self', ''); window.close();</script>")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(DataExpand)
    '**
    '*****************************************************************
    Public Function DataExpand(ByVal pFormNo As String, ByVal pNo As String) As Integer
        Dim SQL As String = ""
        Dim RtnCode As Integer = 0
        '
        Dim xChainClassSizeSldFun, xSlider1, xSlider2, xFinish1, xFinish2, xTape, xOther1 As String
        Dim xItemName As String()
        Dim xSpecialRequest As String()
        Dim xPartType, xSpecial, xStr, xName, xCode As String
        Dim k As Integer = 0
        '
        Server.ScriptTimeout = 900 * 1000   ' 設定逾時時間
        '無需配對資料
        SQL = "Select Top 200 * "
        SQL = SQL & "From W_ItemRequest "
        SQL = SQL + "Where FormNo = '" & pFormNo & "' "
        SQL = SQL + "  And No = '" & pNo & "' "
        SQL = SQL + "  And Sts = '" & "0" & "' "
        SQL = SQL + "  And RIT1I16 + RIT2I16 + RIT3I16 = '" & "" & "' "
        SQL = SQL & "Order by IT1I16, IT2I16, IT3I16 "
        Dim dt_WaitRequest As DataTable = uDataBase.GetDataTable(SQL)
        If dt_WaitRequest.Rows.Count <= 0 Then
            RtnCode = 1
        End If
        '
        '有需配對資料
        If RtnCode = 0 Then
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
                    SQL = "select * from V_ItemRegisterSheet_02 "
                    SQL = SQL & "WHERE DTFormNo = '" & pFormNo & "'"
                    SQL = SQL & "  AND DTNo = '" & pNo & "'"
                    SQL = SQL & "  AND DTItemName1 = '" & dt_WaitRequest.Rows(i)("IT1I16").ToString.Trim & "'"
                    SQL = SQL & "  AND DTItemName2 = '" & dt_WaitRequest.Rows(i)("IT2I16").ToString.Trim & "'"
                    SQL = SQL & "  AND DTItemName3 = '" & dt_WaitRequest.Rows(i)("IT3I16").ToString.Trim & "'"
                    SQL = SQL & "  AND ZIPApply = '1' "
                    SQL = SQL & "  AND DTApplyType = '0' "
                    Dim dt_ItemRegisterSheet As DataTable = uDataBase.GetDataTable(SQL)
                    If dt_ItemRegisterSheet.Rows.Count > 0 Then
                        '   CFML-3* DSBN*73FM V7/DAV3LH6 V7 P12     雙拉頭
                        '   CFC-39 DSBLUY09BK H6 P12                單拉頭
                        '   CFDT1C-3 P12                            無拉頭
                        '
                        xChainClassSizeSldFun = xItemName(0)
                        xSlider1 = dt_ItemRegisterSheet.Rows(0)("SLIDER1").ToString.Trim
                        xFinish1 = dt_ItemRegisterSheet.Rows(0)("FINISH1").ToString.Trim
                        xSlider2 = dt_ItemRegisterSheet.Rows(0)("SLIDER2").ToString.Trim
                        xFinish2 = dt_ItemRegisterSheet.Rows(0)("FINISH2").ToString.Trim
                        xTape = dt_ItemRegisterSheet.Rows(0)("TAPE").ToString.Trim
                        'MsgBox("xSlider1=[" + xSlider1 + "][" + xFinish1 + "]" + Chr(13) + _
                        '       "xSlider2=[" + xSlider2 + "][" + xFinish2 + "]" + Chr(13) + _
                        '       "xTape=[" + xTape + "]")
                        '
                        xSpecial = ""
                        If dt_ItemRegisterSheet.Rows(0)("SREQUEST1").ToString.Trim <> "" Then
                            If xSpecial = "" Then
                                xSpecial = dt_ItemRegisterSheet.Rows(0)("SREQUEST1").ToString.Trim
                            Else
                                xSpecial = xSpecial + " " + dt_ItemRegisterSheet.Rows(0)("SREQUEST1").ToString.Trim
                            End If
                        End If
                        If dt_ItemRegisterSheet.Rows(0)("SREQUEST2").ToString.Trim <> "" Then
                            If xSpecial = "" Then
                                xSpecial = dt_ItemRegisterSheet.Rows(0)("SREQUEST2").ToString.Trim
                            Else
                                xSpecial = xSpecial + " " + dt_ItemRegisterSheet.Rows(0)("SREQUEST2").ToString.Trim
                            End If
                        End If
                        If dt_ItemRegisterSheet.Rows(0)("SREQUEST3").ToString.Trim <> "" Then
                            If xSpecial = "" Then
                                xSpecial = dt_ItemRegisterSheet.Rows(0)("SREQUEST3").ToString.Trim
                            Else
                                xSpecial = xSpecial + " " + dt_ItemRegisterSheet.Rows(0)("SREQUEST3").ToString.Trim
                            End If
                        End If
                        If dt_ItemRegisterSheet.Rows(0)("SREQUEST4").ToString.Trim <> "" Then
                            If xSpecial = "" Then
                                xSpecial = dt_ItemRegisterSheet.Rows(0)("SREQUEST4").ToString.Trim
                            Else
                                xSpecial = xSpecial + " " + dt_ItemRegisterSheet.Rows(0)("SREQUEST4").ToString.Trim
                            End If
                        End If
                        If dt_ItemRegisterSheet.Rows(0)("SREQUEST5").ToString.Trim <> "" Then
                            If xSpecial = "" Then
                                xSpecial = dt_ItemRegisterSheet.Rows(0)("SREQUEST5").ToString.Trim
                            Else
                                xSpecial = xSpecial + " " + dt_ItemRegisterSheet.Rows(0)("SREQUEST5").ToString.Trim
                            End If
                        End If
                        If dt_ItemRegisterSheet.Rows(0)("SREQUEST6").ToString.Trim <> "" Then
                            If xSpecial = "" Then
                                xSpecial = dt_ItemRegisterSheet.Rows(0)("SREQUEST6").ToString.Trim
                            Else
                                xSpecial = xSpecial + " " + dt_ItemRegisterSheet.Rows(0)("SREQUEST6").ToString.Trim
                            End If
                        End If
                        '
                        xSpecialRequest = xSpecial.Split(" ")
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
                If xPartType <> "ZIP" Then
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
                End If
                '
                '****  尋找參考ITEM ********************************************************************************************* 
                '
                'CFML-3* DSBN*73FM V7/DAV3LH6 V7 P12
                'CFC-39 DSBLUY09BK H6 P12
                'CFDT1C-3 P12
                '
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
                            '
                            xStr = xName + " " + xFinish1
                            xName = xName + " " + GetRefItemName("SFNC01", xStr, xFinish1, "PARTS")
                            '
                            xStr = xName + " " + "PARTS"
                            xName = GetRefItemName("PARTS", xStr, xFinish1, "PARTS")
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
                '
                UpdateItemRequest(dt_WaitRequest.Rows(i)("Unique_ID"), xCode, xStr, xSpecialRequest)
            Next
        End If
        '
        Return RtnCode
    End Function
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
                    'sql &= " ICQC16 = '" & dtITEM.Rows(0)("ICQC01").ToString.Trim & "',"                 'ITEM CODE REQUEST CODE
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
        cn.ConnectionString = ConnectString
        '
        If pData <> "" Then
            SQL = "SELECT CN1I09 FROM WAVEDLIB.C0900 "
            SQL = SQL + "WHERE DGRC09 = 'SF1C' AND DDTC09 = '" & pData & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, cn)
            DBAdapter1.Fill(ds, "C0900")
            If ds.Tables("C0900").Rows.Count <= 0 Then uError = True
            cn.Close()
        End If
        '
        Return uError
    End Function


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(DataPump)
    '**
    '*****************************************************************
    Protected Sub DataPump(ByVal pFormNo As String, ByVal pNo As String)
        Dim SQL As String = ""
        '
        SQL = "Select Top 1 * From W_ItemRequest "
        SQL = SQL + "Where FormNo = '" & pFormNo & "' "
        SQL = SQL + "  And No = '" & pNo & "' "
        Dim dt_WaitRequest As DataTable = uDataBase.GetDataTable(SQL)
        If dt_WaitRequest.Rows.Count <= 0 Then
            '
            SQL = "Insert Into W_ItemRequest "
            SQL = SQL & "Select 0, FormNo, No, '', "                            ' FormNo, No, FNo,
            '-----------------------------------------------------------------------------------------------------------
            SQL = SQL & "       '', '', RTRIM(ItemName1), RTRIM(ItemName2), RTRIM(ItemName3), "      ' SEMI16, SSNN16, IT1I16, IT2I16, IT3I16, 
            SQL = SQL & "       '', '', '', '', '', "                           ' YNXC16, ST1C16, ST2C16, ST3C16, ST4C16,
            SQL = SQL & "       '', '', '', '', '', "                           ' ST5C16, ST6C16, ST7C16, SIZC16, CHNC16,
            SQL = SQL & "       '', '', '', '', '', "                           ' CLSC16, TAPC16, SIMF16, SLDC16, SFNC16,
            SQL = SQL & "       '', '', '', '', '', "                           ' SL2C16, SE2C16, SF1C16, SF2C16, SF3C16,
            SQL = SQL & "       '', '', '', '', '', "                           ' SF4C16, SF5C16, SF6C16, FMLC16, QUNC16,
            SQL = SQL & "       '', '', '', '', '', "                           ' PO1C16, PO2C16, PO3C16, PO4C16, PO5C16,
            SQL = SQL & "       '', '', '', '', '', "                           ' PO6C16, PO7C16, ICRX16, RFCC16, RTCC16,
            SQL = SQL & "       '', '1', '', '', "                              ' ITMC16, ICQC16, MAIC16, MPTC16, 
            '-----------------------------------------------------------------------------------------------------------
            SQL = SQL & "       '', '', '', '', '', "                           ' RSEMI16, RSSNN16, RIT1I16, RIT2I16, RIT3I16, 
            SQL = SQL & "       '', '', '', '', '', "                           ' RYNXC16, RST1C16, RST2C16, RST3C16, RST4C16,
            SQL = SQL & "       '', '', '', '', '', "                           ' RST5C16, RST6C16, RST7C16, RSIZC16, RCHNC16,
            SQL = SQL & "       '', '', '', '', '', "                           ' RCLSC16, RTAPC16, RSIMF16, RSLDC16, RSFNC16,
            SQL = SQL & "       '', '', '', '', '', "                           ' RSL2C16, RSE2C16, RSF1C16, RSF2C16, RSF3C16,
            SQL = SQL & "       '', '', '', '', '', "                           ' RSF4C16, RSF5C16, RSF6C16, RFMLC16, RQUNC16,
            SQL = SQL & "       '', '', '', '', '', "                           ' RPO1C16, RPO2C16, RPO3C16, RPO4C16, RPO5C16,
            SQL = SQL & "       '', '', '', '', '', "                           ' RPO6C16, RPO7C16, RICRX16, RRFCC16, RRTCC16,
            SQL = SQL & "       '', '', '', '', "                               ' RITMC16, RICQC16, RMAIC16, RMPTC16, 
            '-----------------------------------------------------------------------------------------------------------
            SQL = SQL & "       'ItemRequest01', "                              ' CreateUser, CreateTime, ModifyUser, ModifyTime
            SQL = SQL & "       '" & NowDateTime & "', "
            SQL = SQL & "       '', "
            SQL = SQL & "       '" & NowDateTime & "' "
            SQL = SQL & "From F_ItemRegisterSheetDT "
            SQL = SQL & " Where FormNo = '" & pFormNo & "' "
            SQL = SQL & "   And No     = '" & pNo & "' "
            SQL = SQL & "   And ApplyType = '" & "0" & "' "
            'test
            'SQL = SQL & "   And Unique_ID = '" & "80934" & "' "
            '
            SQL = SQL & "Group by FormNo, No, RTRIM(ItemName1), RTRIM(ItemName2), RTRIM(ItemName3) "
            SQL = SQL & "Order by FormNo, No, RTRIM(ItemName1), RTRIM(ItemName2), RTRIM(ItemName3) "
            uDataBase.ExecuteNonQuery(SQL)
            '
        End If
    End Sub

End Class
