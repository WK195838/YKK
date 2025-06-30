Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'** Database
'***************************************************************************************************
'Database-Start
Imports System.Data             'SQL
Imports System.Data.SqlClient
'Database-End

<WebService(Namespace:="http://tempuri.org/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class MappingService
     Inherits System.Web.Services.WebService

    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '** 全域變數
    '***********************************************************************************************
    '全域變數-Start
    Dim oDB As New ForProject
    Dim oDataBase As Utility.DataBase = oDB.GetDataBaseObj()
    Dim oWaves As New Waves.CommonService
    Dim oEDICommon As New CommonService
    Dim xRuleList(200) As String            ' Rule資料
    Dim xFieldList(200) As String           ' 欄位名
    Dim xDataTypeList(200) As String        ' 資料型態
    Dim xDataLengthList(200) As String      ' 資料長度
    Dim xDataList(200) As String            ' 解析後資料
    Dim xLogID, xBuyer, xUserID As String   ' LogID / Buyer / UserID
    Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")     ' 現在日時               
    '
    '全域變數-End
    '-----------------------------------------------------------------------------------------------
    '[000]-Rule2Data								                                | 解析Rule將Excel欄位轉入Waves格式檔 
    '[010]-Rule2DataEOES    						                                | 解析Rule將Excel欄位轉入Waves格式檔(EOES) 170525
    '	+									                                        |
    '	+----[100]-GetListInformation						                        | 取得各項定義之Rule資訊
    '	+	+								                                        |
    '	+	+----[200]-ClearListInformation					                        | 清除指定Rule資訊
    '	+									                                        |
    '	+----[200]-ClearListInformation						                        | 清除指定Rule資訊
    '	+									                                        |
    '	+----[300]-GetCommand							                            | 分解Rule變成Command
    '	+									                                        |
    '	+----[400]-GetData							                                | 解析Command後將資料組合
    '	+	    +								                                    |
    '	+	    +----[410]-Command2Data						                        | 解析Command	
    '	+       +               +							                        |
    '	+	    +               +----[411]-(#)DefaultData	                        | 初值	
    '	+	    +               +							                        |
    '	+	    +               +----[412]-(@)SystemData	                        | 系統值	
    '	+	    +               +							                        |
    '	+	    +               +----[413]-(CE)CellData		                        | Excel Cell	
    '	+	    +               +							                        |
    '	+	    +               +----[414]-(FD)FieldData	                        | DataField	
    '	+	    +               +							                        |
    '	+	    +               +----[415]-(MID)MidData		                        | 截取資料-1	
    '	+	    +               +							                        |
    '	+	    +               +----[416]-(MIDSTR)MidStrData                       | 截取資料-2	
    '	+	    +               +							                        |
    '	+	    +               +----[417]-(IF)IFData                               | IF判斷	
    '	+	    +               +							                        |
    '	+	    +               +----[418]-(ITEMCONVERT)ItemConvertData             | 轉換ITEM	
    '	+	    +               +							                        |
    '	+	    +               +----[419]-(COLORCONVERT)ColorConvertData           | 轉換COLOR	
    '	+	    +               +							                        |
    '	+	    +               +----[41A]-(GETKEEPCODE)GetKeepCodeData             | 取得KeepCode(未使用)	
    '	+	    +               +							                        |
    '	+	    +               +----[41B]-(GETDELIVERY)GetDeliveryDate             | 取得納期	
    '	+	    +               +							                        |
    '	+	    +               +----[41C]-(GETQTYUNITCODE)GetQtyUnitCode           | 取得Qty Unit Code	
    '	+	    +								                                    | 
    '	+	    +----[420]-Calculate						                        | 解析Command及組合資料
    '	+									                                        |
    '	+----[500]-ReplaceString						                            | 解析後資料置換關鍵字 
    '	+									                                        |
    '	+----[600]-AccessLog							                            | 解析結果寫入記錄檔
    '	+									                                        |
    '	+----[700]-CheckDataError						                            | 檢測是否有資料解析異常
    '	+									                                        |
    '	+----[800]-WriteDataFile						                            | 寫入Waves格式檔
    '	+									                                        |
    '	+----[810]-WriteDataFileEOES					                            | 寫入Waves格式檔(EOES)
    '	+									                                        |
    '	+----									                                    |
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([000]-Rule2Data)
    '**     解析Rule將Excel欄位轉入Waves格式檔 
    '***********************************************************************************************
    'Rule2Data-Start
    <WebMethod()> _
    Public Function Rule2Data(ByVal pLogID As String, _
                              ByVal pBuyer As String, _
                              ByVal pUserID As String, _
                              ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL, xData As String
        Dim i, j As Integer
        Dim xTempList(10) As String                                                 ' TempList
        '
        xLogID = pLogID                                                             ' Global LogID / Buyer / User
        xBuyer = pBuyer
        xUserID = pUserID
        GetListInformation(pBuyer)                                                  ' 取得Buyer Rule / Field / DataType
        '
        Try
            SQL = "SELECT Unique_ID FROM E_InputSheet "                             ' 讀取Excel Sheet Data
            SQL &= "Where Buyer = '" & pBuyer & "' "
            SQL &= "Order by Unique_ID "
            Dim dt_InputSheet As DataTable = oDataBase.GetDataTable(SQL)
            For i = 0 To dt_InputSheet.Rows.Count - 1
                ClearListInformation(xDataList, 200)                                ' 清空Data List資料
                '
                For j = 0 To 199
                    If xRuleList(j) <> "" Then
                        ClearListInformation(xTempList, 10)                         ' 清空Temp List資料
                        xTempList = GetCommand(xRuleList(j), xFieldList(j))         ' Rule分解成各Command
                        xData = GetData(pBuyer, xTempList, xFieldList(j), xDataTypeList(j), dt_InputSheet.Rows(i).Item("Unique_ID"))    '取得及組合各Command資料
                        xDataList(j) = ReplaceString(xData)                         ' 置換關鍵字
                        '
                        ' 檢查解析後內容
                        If InStr(UCase(xData), "ERROR") > 0 Then
                            ' 內容異常
                            oDB.AccessLog(xLogID, xBuyer, "Rule2Data", NowDateTime, 1, CStr(i + 1), xFieldList(j), xRuleList(j), xDataList(j), "", xUserID, "")
                        Else
                            If Len(xData) > CInt(xDataLengthList(j)) Then
                                ' 內容字串長度異常
                                oDB.AccessLog(xLogID, xBuyer, "Rule2Data", NowDateTime, 1, CStr(i + 1), xFieldList(j), xRuleList(j), xDataList(j), "Length Error:[" + CStr(Len(xData)) + "]<>[" + xDataLengthList(j) + "]", xUserID, "")
                            Else
                                ' 內容正常
                                oDB.AccessLog(xLogID, xBuyer, "Rule2Data", NowDateTime, 0, CStr(i + 1), xFieldList(j), xRuleList(j), xDataList(j), "", xUserID, "")
                            End If
                        End If
                        '
                    End If
                Next
                '
                RtnCode = CheckDataError(pLogID)
                If RtnCode = 0 Then
                    RtnCode = WriteDataFile(pBuyer, pGRBuyer, dt_InputSheet.Rows(i).Item("Unique_ID"))
                    If RtnCode <> 0 Then Exit For
                Else
                    Exit For
                End If
            Next
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(xLogID, xBuyer, "Rule2Data", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", xUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'Rule2Data-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([010]-Rule2DataEOES)
    '**     解析Rule將Excel欄位轉入Waves格式檔(EOES) 170523
    '***********************************************************************************************
    'Rule2DataEOSE-Start
    <WebMethod()> _
    Public Function Rule2DataEOES(ByVal pLogID As String, _
                                  ByVal pBuyer As String, _
                                  ByVal pUserID As String, _
                                  ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL, xData As String
        Dim i, j As Integer
        Dim xTempList(10) As String                                                 ' TempList
        ' 
        Dim eBuyer As String = ""
        Dim eGRBuyer As String = ""
        '
        xLogID = pLogID                                                             ' Global LogID / Buyer / User
        xBuyer = pBuyer
        xUserID = pUserID
        '
        Try
            '
            ' JJJ-START 20/02/19 (JJJ展開)
            ' 2/19前PGM
            'SQL = "SELECT Unique_ID, B1 + '-' + C1 As CustBuyer FROM E_InputSheet "     ' 讀取Excel Sheet Data
            ' 新PGM
            SQL = "SELECT Unique_ID, "
            SQL &= "Case SUBSTRING(B1, 1, 4) WHEN 'ZJJJ' THEN B1 + '-' + '000999' ELSE B1 + '-' + C1 END  As CustBuyer "
            SQL &= "FROM E_InputSheet "
            ' JJJ-END
            '
            SQL &= "Where Buyer = '" & pBuyer & "' "
            SQL &= "Order by Unique_ID "
            Dim dt_InputSheet As DataTable = oDataBase.GetDataTable(SQL)
            For i = 0 To dt_InputSheet.Rows.Count - 1
                ' 取得 BUYERGROUP
                eBuyer = dt_InputSheet.Rows(i).Item("CustBuyer")
                SQL = "Select BuyerGroup From M_ControlRecord "
                SQL = SQL & " Where Buyer = '" & eBuyer & "'"
                Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(SQL)
                If dt_ControlRecord.Rows.Count > 0 Then eGRBuyer = dt_ControlRecord.Rows(0).Item("BuyerGroup")

                GetListInformation(eBuyer)                                          ' 取得Buyer Rule / Field / DataType
                ClearListInformation(xDataList, 200)                                ' 清空Data List資料
                '
                For j = 0 To 199
                    If xRuleList(j) <> "" Then
                        ClearListInformation(xTempList, 10)                         ' 清空Temp List資料
                        xTempList = GetCommand(xRuleList(j), xFieldList(j))         ' Rule分解成各Command
                        xData = GetData(eBuyer, xTempList, xFieldList(j), xDataTypeList(j), dt_InputSheet.Rows(i).Item("Unique_ID"))    '取得及組合各Command資料
                        xDataList(j) = ReplaceString(xData)                         ' 置換關鍵字
                        '
                        ' 檢查解析後內容
                        If InStr(UCase(xData), "ERROR") > 0 Then
                            ' 內容異常
                            oDB.AccessLog(xLogID, xBuyer, "Rule2Data", NowDateTime, 1, CStr(i + 1), xFieldList(j), xRuleList(j), xDataList(j), "", xUserID, "")
                        Else
                            If Len(xData) > CInt(xDataLengthList(j)) Then
                                ' 內容字串長度異常
                                oDB.AccessLog(xLogID, xBuyer, "Rule2Data", NowDateTime, 1, CStr(i + 1), xFieldList(j), xRuleList(j), xDataList(j), "Length Error:[" + CStr(Len(xData)) + "]<>[" + xDataLengthList(j) + "]", xUserID, "")
                            Else
                                ' 內容正常
                                oDB.AccessLog(xLogID, xBuyer, "Rule2Data", NowDateTime, 0, CStr(i + 1), xFieldList(j), xRuleList(j), xDataList(j), "", xUserID, "")
                            End If
                        End If
                        '
                    End If
                Next
                '
                RtnCode = CheckDataError(pLogID)
                If RtnCode = 0 Then
                    RtnCode = WriteDataFileEOES(pBuyer, eBuyer, eGRBuyer, dt_InputSheet.Rows(i).Item("Unique_ID"))
                    If RtnCode <> 0 Then Exit For
                Else
                    Exit For
                End If
            Next
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(xLogID, xBuyer, "Rule2Data", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", xUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'Rule2DataEOES-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([100]-GetListInformation)
    '**     取得各項定義之Rule資訊(Rule / Field / DataType / DataLength) 
    '***********************************************************************************************
    'GetListInformation-Start
    Sub GetListInformation(ByVal pBuyer As String)
        Dim SQL As String
        Dim i As Integer
        Try
            ClearListInformation(xRuleList, 200)        ' 清空 Rule / Field / DataType / DataLength 等List資料
            ClearListInformation(xFieldList, 200)
            ClearListInformation(xDataTypeList, 200)
            ClearListInformation(xDataLengthList, 200)
            '
            SQL = "SELECT * FROM M_ImportRule "         ' 取得Buyer Rule / Field / DataType 等資料
            SQL &= "Where Buyer = '" & pBuyer & "' "
            SQL &= "  And Sts   = '1' "
            SQL &= "Order by SeqNo "
            Dim dt_List As DataTable = oDataBase.GetDataTable(SQL)
            For i = 0 To dt_List.Rows.Count - 1
                xRuleList(i) = dt_List.Rows(i).Item("DataRule").ToString
                xFieldList(i) = dt_List.Rows(i).Item("Field").ToString
                xDataTypeList(i) = dt_List.Rows(i).Item("DataType").ToString
            Next
            '
            SQL = "SELECT * FROM M_ImportRule "         ' 取得DataLength資料
            SQL &= "Where Buyer = '" & "xxxxxx" & "' "
            SQL &= "  And Sts   = '1' "
            SQL &= "Order by SeqNo "
            Dim dt_List1 As DataTable = oDataBase.GetDataTable(SQL)
            For i = 0 To dt_List1.Rows.Count - 1
                If dt_List1.Rows(i).Item("DataType").ToString = "C" Then
                    xDataLengthList(i) = dt_List1.Rows(i).Item("DataRule").ToString
                Else
                    xDataLengthList(i) = "999"
                End If
            Next
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "GetListInformation", NowDateTime, 1, "Failed", "Code", "=", "9", "", xUserID, "")
        End Try
    End Sub
    'GetListInformation-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([200]-ClearListInformation)
    '**     清除指定Rule資訊
    '***********************************************************************************************
    'ClearListInformation-Start
    Sub ClearListInformation(ByVal pList() As String, ByVal Count As Integer)
        Try
            For i As Integer = 0 To Count - 1
                pList(i) = ""
            Next
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "ClearListInformation", NowDateTime, 1, "Failed", "Code", "=", "9", "", xUserID, "")
        End Try
    End Sub
    'ClearListInformation-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([300]-GetCommand)
    '**     分解Rule變成Command
    '***********************************************************************************************
    'GetCommand-Start
    Function GetCommand(ByVal pRule As String, ByVal pField As String) As String()
        Dim xTemp(10) As String
        Dim i, skip As Integer
        Dim str As String = pRule
        '
        Try
            For i = 0 To 9
                xTemp(i) = ""
            Next
            '
            i = 0
            While (InStr(str, "+") > 0 And InStr(str, "(+)") <= 0) Or _
                  (InStr(str, "-") > 0 And InStr(str, "(-)") <= 0) Or _
                  (InStr(str, "*") > 0 And InStr(str, "(*)") <= 0) Or _
                  (InStr(str, "/") > 0 And InStr(str, "(/)") <= 0) Or _
                  (InStr(str, "|") > 0)
                skip = 0
                ' [ + ]
                If InStr(str, "+") > 0 Then
                    If skip = 0 Then
                        xTemp(i) = Mid(str, 1, InStr(str, "+") - 1)
                        xTemp(i + 1) = "+"
                        i = i + 2
                        str = Mid(str, InStr(str, "+") + 1)
                        skip = 1
                    End If
                End If
                ' [ - ]
                If InStr(str, "-") > 0 Then
                    If skip = 0 Then
                        xTemp(i) = Mid(str, 1, InStr(str, "-") - 1)
                        xTemp(i + 1) = "-"
                        i = i + 2
                        str = Mid(str, InStr(str, "-") + 1)
                        skip = 1
                    End If
                End If
                ' [ * ]
                If InStr(str, "*") > 0 Then
                    If skip = 0 Then
                        xTemp(i) = Mid(str, 1, InStr(str, "*") - 1)
                        xTemp(i + 1) = "*"
                        i = i + 2
                        str = Mid(str, InStr(str, "*") + 1)
                        skip = 1
                    End If
                End If
                ' [ / ]
                If InStr(str, "/") > 0 Then
                    If skip = 0 Then
                        xTemp(i) = Mid(str, 1, InStr(str, "/") - 1)
                        xTemp(i + 1) = "/"
                        i = i + 2
                        str = Mid(str, InStr(str, "/") + 1)
                        skip = 1
                    End If
                End If
                ' [ | ]
                If InStr(str, "|") > 0 Then
                    If skip = 0 Then
                        xTemp(i) = Mid(str, 1, InStr(str, "|") - 1)
                        xTemp(i + 1) = "|"
                        i = i + 2
                        str = Mid(str, InStr(str, "|") + 1)
                        skip = 1
                    End If
                End If
            End While
            '
            If str <> "" Then xTemp(i) = str
            '
        Catch ex As Exception
            For i = 0 To 9
                xTemp(i) = "#(Error)"
            Next
            oDB.AccessLog(xLogID, xBuyer, "GetCommand", NowDateTime, 1, "Failed", pField, pRule, "", "", xUserID, "")
        End Try
        '
        Return xTemp
    End Function
    'GetCommand-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([400]-GetData)
    '**     解析Command後將資料組合
    '***********************************************************************************************
    'GetData-Start
    Function GetData(ByVal pBuyer As String, _
                     ByVal pTempList As String(), _
                     ByVal pField As String, _
                     ByVal pDataType As String, _
                     ByVal pUID As Integer) As String
        Dim xData As String = ""
        Dim xMark As String = ""
        Dim i As Integer
        '
        Try
            '解析Command及組合資料
            For i = 0 To 9
                If pTempList(i) <> "" Then
                    If pTempList(i) <> "+" And pTempList(i) <> "-" And pTempList(i) <> "*" And pTempList(i) <> "/" And pTempList(i) <> "|" Then
                        If xData = "" Then
                            xData = Command2Data(pBuyer, pTempList(i), pUID)
                        Else
                            xData = Calculate(xData, xMark, Command2Data(pBuyer, pTempList(i), pUID), pTempList(0))
                        End If
                    Else
                        xMark = pTempList(i)
                    End If
                End If
            Next
            '檢測Field是數字欄位(I)/(N)
            If pDataType = "I" Then
                If xData = "" Then xData = "0"
                If Not IsNumeric(xData) Then xData = xData + "-->Error"
            End If
            If pDataType = "N" Then
                If xData = "" Then xData = "0.0"
                If Not IsNumeric(xData) Then xData = xData + "-->Error"
            End If
        Catch ex As Exception
            xData = "-->Error"
            oDB.AccessLog(xLogID, xBuyer, "GetData", NowDateTime, 1, "Failed", pField, "=", xData, "", xUserID, "")
        End Try
        '
        Return xData
    End Function
    'GetData-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([410]-Command2Data)
    '**     解析Command
    '***********************************************************************************************
    'Command2Data-Start
    Function Command2Data(ByVal pBuyer As String, ByVal pRule As String, ByVal pUID As Integer) As String
        Dim RtnString As String = ""
        Dim Fun As String = Mid(UCase(pRule), 1, InStr(pRule, "(") - 1)
        '
        Try
            Select Case Fun
                Case "#"                                                    ' 初值
                    RtnString = DefaultData(pRule)
                Case "@"                                                    ' 系統值
                    RtnString = SystemData(pRule)
                Case "CE"                                                   ' Excel Cell
                    RtnString = CellData(pRule, pUID)
                Case "FD"                                                   ' DataField
                    RtnString = FieldData(pRule)
                Case "MID"                                                  ' 截取資料-1
                    RtnString = MidData(pBuyer, pRule, pUID)
                Case "MIDSTR"                                               ' 截取資料-2
                    RtnString = MidStrData(pBuyer, pRule, pUID)
                Case "IF"                                                   ' IF判斷
                    RtnString = IFData(pBuyer, pRule, pUID)
                Case "ITEMCONVERT"                                          ' 轉換ITEM
                    RtnString = ItemConvertData(pBuyer, pRule, pUID)
                Case "COLORCONVERT"                                         ' 轉換COLOR
                    RtnString = ColorConvertData(pBuyer, pRule, pUID)
                Case "GETKEEPCODE"                                          ' 取得KeepCode(未使用)
                    RtnString = GetKeepCodeData(pBuyer, pRule, pUID)
                Case "GETDELIVERY"                                          ' 取得納期
                    RtnString = GetDeliveryDate(pBuyer, pRule, pUID)
                Case "GETQTYUNITCODE"                                       ' 取得Qty Unit Code
                    RtnString = GetQtyUnitCode(pBuyer, pRule, pUID)
                Case Else
                    RtnString = "Error"
            End Select

        Catch ex As Exception
            RtnString = "Error"
            oDB.AccessLog(xLogID, xBuyer, "Command2Data", NowDateTime, 1, "Failed", "Code", "=", "9", pRule, xUserID, "")
        End Try
        '
        Return RtnString
    End Function
    'Command2Data-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([411]-DefaultData)
    '**     初值
    '***********************************************************************************************
    'DefaultData-Start
    Function DefaultData(ByVal pRule As String) As String
        Dim RtnString As String = ""
        Dim start As Integer = InStr(pRule, "(") + 1
        Dim length As Integer = InStr(pRule, ")") - 1 - start + 1
        '
        Try
            If InStr(pRule, "(SP)") > 0 Then
                RtnString = ""
            Else
                RtnString = Mid(pRule, start, length)
            End If
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "DefaultData", NowDateTime, 1, "Failed", "Code", "=", "9", pRule, xUserID, "")
        End Try
        '
        Return RtnString
    End Function
    'DefaultData-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([412]-SystemData)
    '**     系統值
    '***********************************************************************************************
    'SystemData-Start
    Function SystemData(ByVal pRule As String) As String
        Dim RtnString As String = ""
        Dim start As Integer = InStr(pRule, "(") + 1
        Dim length As Integer = InStr(pRule, ")") - 1 - start + 1
        Dim Value As String = Mid(UCase(pRule), start, length)
        '
        Try
            If Value = "NOW.DATE" Then RtnString = Now.ToString("yyyyMMdd") ' 現在日期
            If Value = "NOW.TIME" Then RtnString = Now.ToString("HHmmss") ' 現在時間
            If Value = "NOW.MONTH" Then RtnString = Now.Month.ToString
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "SystemData", NowDateTime, 1, "Failed", "Code", "=", "9", pRule, xUserID, "")
        End Try
        '
        Return RtnString
    End Function
    'SystemData-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([413]-CellData)
    '**     Excel Cell
    '***********************************************************************************************
    'CellData-Start
    Function CellData(ByVal pRule As String, ByVal pUID As Integer) As String
        Dim RtnString As String = ""
        Dim start As Integer = InStr(pRule, "(") + 1
        Dim length As Integer = InStr(pRule, ")") - 1 - start + 1
        Dim Value As String = Mid(UCase(pRule), start, length)
        Dim SQL As String
        '
        Try
            SQL = "SELECT * FROM E_InputSheet "
            SQL &= "Where Unique_ID = '" & CStr(pUID) & "' "
            Dim dt_Cell As DataTable = oDataBase.GetDataTable(SQL)
            If dt_Cell.Rows.Count > 0 Then RtnString = dt_Cell.Rows(0).Item(Value)
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "CellData", NowDateTime, 1, "Failed", "Code", "=", "9", pRule, xUserID, "")
        End Try
        '
        Return RtnString
    End Function
    'CellData-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([414]-FieldData)
    '**     DataField
    '***********************************************************************************************
    'FieldData-Start
    Function FieldData(ByVal pRule As String) As String
        Dim RtnString As String = ""
        Dim start As Integer = InStr(pRule, "(") + 1
        Dim length As Integer = InStr(pRule, ")") - 1 - start + 1
        Dim Value As String = Mid(UCase(pRule), start, length)
        '
        Try
            For i As Integer = 0 To 199
                If xFieldList(i) = Value Then RtnString = xDataList(i)
            Next
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "FieldData", NowDateTime, 1, "Failed", "Code", "=", "9", pRule, xUserID, "")
        End Try
        '
        Return RtnString
    End Function
    'FieldData-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([415]-MidData)
    '**     截取資料-1
    '***********************************************************************************************
    'MidData-Start
    Function MidData(ByVal pBuyer As String, ByVal pRule As String, ByVal pUID As Integer) As String
        Dim RtnString As String = ""
        '
        Try
            'MID(CE(H1),2,1)==>CE(H1)
            Dim start As Integer = InStr(pRule, "(") + 1
            Dim length As Integer = InStr(pRule, ")") - start + 1
            Dim Value As String = Mid(UCase(pRule), start, length)
            Value = Command2Data(pBuyer, Value, pUID)
            'MID(CE(H1),2,1)==>2,1
            start = InStr(pRule, ",") + 1
            length = Len(pRule) - InStr(pRule, ",") + 1
            Dim str As String = Mid(UCase(pRule), start, length)
            '2,1==>2
            start = 1
            length = InStr(str, ",") - 1
            Dim xStart As Integer = CInt(Mid(str, start, length))
            '2,1==>1
            start = InStr(str, ",") + 1
            length = Len(str) - InStr(str, ",") - 1
            Dim xLength As Integer = CInt(Mid(str, start, length))
            '
            RtnString = Mid(Value, xStart, xLength)
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "MidData", NowDateTime, 1, "Failed", "Code", "=", "9", pRule, xUserID, "")
        End Try
        '
        Return RtnString
    End Function
    'MidData-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([416]-MidStrData)
    '**     截取資料-1
    '***********************************************************************************************
    'MidStrData-Start
    Function MidStrData(ByVal pBuyer As String, ByVal pRule As String, ByVal pUID As Integer) As String
        Dim RtnString As String = ""
        '
        Try
            'MIDSTR(CE(A1),#(EX),#(CEL))==>CE(A1)
            Dim start As Integer = InStr(pRule, "(") + 1
            Dim length As Integer = InStr(pRule, ")") - start + 1
            Dim Value As String = Mid(UCase(pRule), start, length)
            Value = Command2Data(pBuyer, Value, pUID)
            'MIDSTR(CE(A1),#(EX),#(CEL))==>#(EX),#(CEL)
            start = InStr(pRule, ",") + 1
            length = Len(pRule) - InStr(pRule, ",") - 1
            Dim str As String = Mid(UCase(pRule), start, length)
            '#(EX),#(CEL)==>#(EX)
            start = 1
            length = InStr(str, ",") - 1
            Dim StartV As String = Mid(str, start, length)
            StartV = Command2Data(pBuyer, StartV, pUID)
            If StartV = "~" Then StartV = "-"
            '#(EX),#(CEL)==>#(CEL)
            start = InStr(str, ",") + 1
            length = Len(str) - start + 1
            Dim EndV As String = Mid(str, start, length)
            EndV = Command2Data(pBuyer, EndV, pUID)
            '
            If IsNumeric(StartV) Then
                start = CInt(StartV)
            Else
                start = InStr(Value, StartV) + Len(StartV)
            End If
            If IsNumeric(EndV) Then
                length = CInt(EndV)
            Else
                length = InStr(Value, EndV) - start
            End If
            '
            RtnString = Mid(Value, start, length)
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "MidStrData", NowDateTime, 1, "Failed", "Code", "=", "9", pRule, xUserID, "")
        End Try
        '
        Return RtnString
    End Function
    'MidStrData-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([417]-IFData)
    '**     IF判斷
    '***********************************************************************************************
    'IFData-Start
    Function IFData(ByVal pBuyer As String, ByVal pRule As String, ByVal pUID As Integer) As String
        Dim RtnString As String = ""
        '
        Try
            'IF(CE(A1)=FD(DDD),#(OK),#(NO))==>CE(A1)
            Dim start As Integer = InStr(pRule, "(") + 1
            Dim length As Integer = InStr(pRule, ")=") - start + 1
            Dim Key1 As String = Mid(UCase(pRule), start, length)
            Key1 = Command2Data(pBuyer, Key1, pUID)
            '判斷符號(=, !=, >=, <=, >, <)
            Dim Mark As String = ""
            If InStr(pRule, "=") > 0 Then
                Mark = "="
                If InStr(pRule, "!=") > 0 Then Mark = "!="
                If InStr(pRule, ">=") > 0 Then Mark = ">="
                If InStr(pRule, "<=") > 0 Then Mark = "<="
            Else
                Mark = ">"
                If InStr(pRule, "<") > 0 Then Mark = "<"
            End If
            'IF(CE(A1)=FD(DDD),#(OK),#(NO))==>FD(DDD)
            start = InStr(pRule, Mark) + Len(Mark)
            If InStr(pRule, ",#") Then length = InStr(pRule, ",#") - start
            If InStr(pRule, ",CE") Then length = InStr(pRule, ",CE") - start
            If InStr(pRule, ",FD") Then length = InStr(pRule, ",FD") - start
            Dim Key2 As String = Mid(UCase(pRule), start, length)
            Key2 = Command2Data(pBuyer, Key2, pUID)
            'IF(CE(A1)=FD(DDD),#(OK),#(NO))==>#(OK),#(NO)
            If InStr(pRule, ",#") Then start = InStr(pRule, ",#") + 1
            If InStr(pRule, ",CE") Then start = InStr(pRule, ",CE") + 1
            If InStr(pRule, ",FD") Then start = InStr(pRule, ",FD") + 1
            length = Len(pRule) - start
            Dim str As String = Mid(UCase(pRule), start, length)
            '#(OK),#(NO)==>#(OK)
            start = 1
            length = InStr(str, ",") - 1
            Dim TrueValue As String = Mid(str, start, length)
            TrueValue = Command2Data(pBuyer, TrueValue, pUID)
            '#(OK),#(NO)==>#(NO)
            start = InStr(str, ",") + 1
            length = Len(str) - start + 1
            Dim FalseValue As String = Mid(str, start, length)
            FalseValue = Command2Data(pBuyer, FalseValue, pUID)
            '
            Select Case Mark
                Case "="
                    If Key1 = Key2 Then
                        RtnString = TrueValue
                    Else
                        RtnString = FalseValue
                    End If
                Case "!="
                    If Key1 <> Key2 Then
                        RtnString = TrueValue
                    Else
                        RtnString = FalseValue
                    End If
                Case ">"
                    If Key1 > Key2 Then
                        RtnString = TrueValue
                    Else
                        RtnString = FalseValue
                    End If
                Case "<"
                    If Key1 < Key2 Then
                        RtnString = TrueValue
                    Else
                        RtnString = FalseValue
                    End If
                Case ">="
                    If Key1 >= Key2 Then
                        RtnString = TrueValue
                    Else
                        RtnString = FalseValue
                    End If
                Case "<="
                    If Key1 <= Key2 Then
                        RtnString = TrueValue
                    Else
                        RtnString = FalseValue
                    End If
                Case Else
                    RtnString = ""
            End Select
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "IFData", NowDateTime, 1, "Failed", "Code", "=", "9", pRule, xUserID, "")
        End Try
        '
        Return RtnString
    End Function
    'IFData-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([418]-ItemConvertData)
    '**     轉換ITEM
    '***********************************************************************************************
    'ItemConvertData-Start
    Function ItemConvertData(ByVal pBuyer As String, ByVal pRule As String, ByVal pUID As Integer) As String
        Dim RtnString As String = ""
        '
        Try
            'ITEMCONVERT(#(A4760~000001),CE(A1),CE(B1),#(SP))==>CE(A1),CE(B1),#(SP)
            Dim start As Integer = InStr(pRule, ",") + 1
            Dim length As Integer = Len(pRule) - start
            Dim str As String = Mid(UCase(pRule), start, length)
            'CE(A1),CE(B1),#(SP)==>CE(A1)
            start = 1
            length = InStr(str, ",") - 1
            Dim xCode As String = Mid(str, start, length)
            xCode = Command2Data(pBuyer, xCode, pUID)
            'CE(A1),CE(B1),#(SP)==>CE(B1),#(SP)
            start = InStr(str, ",") + 1
            length = Len(str) - start + 1
            str = Mid(str, start, length)
            'CE(B1),#(SP)==>CE(B1)
            start = 1
            length = InStr(str, ",") - 1
            Dim xColor As String = Mid(str, start, length)
            xColor = Command2Data(pBuyer, xColor, pUID)
            'CE(B1),#(SP)==>#(SP)
            start = InStr(str, ",") + 1
            length = Len(str) - start + 1
            Dim xSlider As String = Mid(str, start, length)
            xSlider = Command2Data(pBuyer, xSlider, pUID)
            '
            '搜尋Item Master
            '取得指定Buyer and Color     ITEMCONVERT(#(Buyer),CE(A1),CE(B1),#(SP))==>#(Buyer)
            start = InStr(pRule, "(") + 1
            length = InStr(pRule, ",") - start
            Dim xBuyer = Mid(UCase(pRule), start, length)
            xBuyer = Replace(Command2Data(pBuyer, xBuyer, pUID), "~", "-")
            ' Color(xxx/xxx/xxxxxx/xxx/xxx)
            Dim Color As Object
            If xColor <> "" Then
                Color = Split(xColor + "/", "/")
            Else
                Color = Split(xColor + "/////", "/")
            End If
            '
            Dim sql As String
            sql = "SELECT YCode FROM M_ItemConvert "
            sql &= "Where Buyer = '" & xBuyer & "' "
            '益鏈(K0790-000999)支援大小寫客戶ITEM
            If pBuyer <> "K0790-000999" Then
                sql &= "  And Code  = '" & xCode & "' "
            Else
                sql &= "  And Code  = '" & xCode & "' Collate SQL_Latin1_General_CP1_CS_AS "
            End If
            sql &= "  And Color1 = '" & Color(2) & "' "
            sql &= "  And Color2 = '" & Color(3) & "' "
            sql &= "  And Color3 = '" & Color(4) & "' "
            sql &= "  And SliderStatus = '" & xSlider & "' "
            Dim dt_Item As DataTable = oDataBase.GetDataTable(sql)
            If dt_Item.Rows.Count > 0 Then
                RtnString = dt_Item.Rows(0).Item("YCode")
            Else
                RtnString = "Error"
            End If
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "ItemConvertData", NowDateTime, 1, "Failed", "Code", "=", "9", pRule, xUserID, "")
        End Try
        '
        Return RtnString
    End Function
    'ItemConvertData-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([419]-ColorConvertData)
    '**     轉換COLOR
    '***********************************************************************************************
    'ColorConvertData-Start
    Function ColorConvertData(ByVal pBuyer As String, ByVal pRule As String, ByVal pUID As Integer) As String
        Dim RtnString As String = ""
        '
        Try
            'Season, Color(xxx/xxx/xxxxxx/xxx/xxx), Item
            'COLORCONVERT(#(01),CE(C1),CE(G1),FD(ITMC5K))==>CE(C1),CE(G1),FD(ITMC5K)
            Dim start As Integer = InStr(pRule, ",") + 1
            Dim length As Integer = Len(pRule) - start
            Dim str As String = Mid(UCase(pRule), start, length)

            'CE(C1),CE(G1),FD(ITMC5K)==>CE(C1)
            start = 1
            length = InStr(str, ",") - 1
            Dim xSeason As String = Mid(str, start, length)
            xSeason = Command2Data(pBuyer, xSeason, pUID)

            'CE(C1),CE(G1),FD(ITMC5K)==>CE(G1),FD(ITMC5K)
            start = InStr(str, ",") + 1
            length = Len(str) - start + 1
            str = Mid(str, start, length)

            'CE(G1),FD(ITMC5K)==>CE(G1)
            start = 1
            length = InStr(str, ",") - 1
            Dim xColor As String = Mid(str, start, length)
            xColor = Command2Data(pBuyer, xColor, pUID)

            'CE(G1),FD(ITMC5K)==>FD(ITMC5K)
            start = InStr(str, ",") + 1
            length = Len(str) - start + 1
            Dim xCode As String = Mid(str, start, length)
            xCode = Command2Data(pBuyer, xCode, pUID)

            '
            '搜尋Color Master
            '   取得指定Buyer and Color     COLORCONVERT(#(01),CE(C1),CE(G1),FD(ITMC5K))==>#(Buyer)
            start = InStr(pRule, "(") + 1
            length = InStr(pRule, ",") - start
            Dim xBuyer = Mid(UCase(pRule), start, length)
            xBuyer = Replace(Command2Data(pBuyer, xBuyer, pUID), "~", "-")
            Dim Color As Object = Split(xColor + "/", "/")         ' Color(xxx/xxx/xxxxxx/xxx/xxx)
            '
            '取得ITEM GROUP
            Dim xPartType As String = ""
            If Len(xCode) = 7 And IsNumeric(xCode) = True Then
                oWaves.Timeout = 900 * 1000
                xPartType = oWaves.GetPartType(xCode)    ' 採用ITEM決定 (G=環保 / IW=進口防水)
            Else
                xPartType = xCode                        ' 採用FCT-SYSTEM或自定
            End If
            '
            Dim sql As String
            sql = "SELECT YColor FROM M_ColorConvert "
            sql &= "Where Buyer  = '" & xBuyer & "' "
            sql &= "  And Season = '" & xSeason & "' "
            sql &= "  And Color1 = '" & Color(0) & "' "
            sql &= "  And Color2 = '" & Color(1) & "' "
            sql &= "  And Green  = '" & xPartType & "' "
            Dim dt_Color As DataTable = oDataBase.GetDataTable(sql)
            If dt_Color.Rows.Count > 0 Then
                RtnString = dt_Color.Rows(0).Item("YColor")
            Else
                RtnString = "Error"
            End If
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "ColorConvertData", NowDateTime, 1, "Failed", "Code", "=", "9", pRule, xUserID, "")
        End Try
        '
        Return RtnString
    End Function
    'ColorConvertData-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([41A]-GetKeepCodeData)
    '**     取得KeepCode
    '***********************************************************************************************
    'GetKeepCodeData-Start
    Function GetKeepCodeData(ByVal pBuyer As String, ByVal pRule As String, ByVal pUID As Integer) As String
        Dim RtnString As String = ""
        '
        Try
            'KeepCode-Type=01
            If InStr(pRule, "01") > 0 Then
                '
                'GETKEEPCODE(#(01),CE(E1))==>CE(E1)
                Dim start As Integer = InStr(pRule, ",") + 1
                Dim length As Integer = Len(pRule) - start
                Dim Value As String = Mid(UCase(pRule), start, length)
                Value = Command2Data(pBuyer, Value, pUID)
                '
                If Mid(Value, 1, 1) = "Y" Then RtnString = "NK"
                If Mid(Value, 1, 1) = "S" Then RtnString = "NK-SS"
            End If
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "GetKeepCodeData", NowDateTime, 1, "Failed", "Code", "=", "9", pRule, xUserID, "")
        End Try
        '
        Return RtnString
    End Function
    'GetKeepCodeData-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([41B]-GetDeliveryDate)
    '**     取得納期
    '***********************************************************************************************
    'GetDeliveryDate-Start
    '@@@@@@@
    Function GetDeliveryDate(ByVal pBuyer As String, ByVal pRule As String, ByVal pUID As Integer) As String
        Dim RtnString As String = ""
        Dim i As Integer
        '
        Try
            'FD(ACRNW1):客戶ITEM, FD(SMPF5K):SAMPLE FLAG, CE(T1):客戶納期 or #(SP):系統納期
            'GETDELIVERY(FD(ACRNW1),FD(SMPF5K),CE(T1))==>FD(ACRNW1),FD(SMPF5K),CE(T1)
            Dim start As Integer = InStr(pRule, "(") + 1
            Dim length As Integer = Len(pRule) - start
            Dim str As String = Mid(UCase(pRule), start, length)
            'GET FD(ACRNW1)
            start = 1
            length = InStr(str, ",") - 1
            Dim xITEM As String = Mid(str, start, length)
            xITEM = Command2Data(pBuyer, xITEM, pUID)
            xITEM = Trim(xITEM)

            'FD(ACRNW1),FD(SMPF5K),CE(T1)==>FD(SMPF5K),CE(T1)
            start = InStr(str, ",") + 1
            length = Len(str) - start + 1
            str = Mid(str, start, length)
            'GET FD(SMPF5K)
            start = 1
            length = InStr(str, ",") - 1
            Dim xSMPF As String = Mid(str, start, length)
            xSMPF = Command2Data(pBuyer, xSMPF, pUID)
            xSMPF = Trim(xSMPF)

            'FD(SMPF5K),CE(T1)==>CE(T1)
            start = InStr(str, ",") + 1
            length = Len(str) - start + 1
            str = Mid(str, start, length)
            'GET CE(T1)
            Dim xDeliveryDate As String = str
            xDeliveryDate = Command2Data(pBuyer, xDeliveryDate, pUID)
            xDeliveryDate = Trim(xDeliveryDate)
            '
            '納期計算
            '   參數:客戶料號, 樣品Flag, 客戶指定納期
            '   @----@納期基準
            '           @----0:系統
            '           @       #----特殊ITEM納期
            '           @               #----0:無
            '           @               #       %----支援種類
            '           @               #               %----0:大貨+樣品
            '           @               #                       +----"大貨納期" 或 "樣品納期"
            '           @               #               %----1:大貨
            '           @               #                       +----"大貨納期"
            '           @               #               %----2:樣品
            '           @               #                       +----"樣品納期"
            '           @               #----1:有
            '           @                       %----支援種類
            '           @                               %----0:大貨+樣品
            '           @                                       +----"特殊ITEM納期"-->"大貨納期" 或 "樣品納期"
            '           @                               %----1:大貨
            '           @                                       +----"特殊ITEM納期"-->"大貨納期"
            '           @                               %----2:樣品
            '           @                                       +----"樣品納期"
            '           @
            '           @----1:客戶
            '           @       #----特殊ITEM納期
            '           @               #----0:無
            '           @               #       %----支援種類
            '           @               #               %----0:大貨+樣品
            '           @               #                       +----客戶指定的 "大貨納期" 或 "樣品納期"
            '           @               #               %----1:大貨
            '           @               #                       +----客戶指定的 "大貨納期"
            '           @               #               %----2:樣品
            '           @               #                       +----客戶指定的 "樣品納期"
            '           @               #----1:有
            '           @                       %----支援種類
            '           @                               %----0:大貨+樣品
            '           @                                       +----"特殊ITEM納期"-->客戶指定的 "大貨納期" 或 "樣品納期"
            '           @                               %----1:大貨
            '           @                                       +----"特殊ITEM納期"-->客戶指定的 "大貨納期"
            '           @                               %----2:樣品
            '           @                                       +----客戶指定的 "樣品納期"
            '       
            '@@@@@@@
            Dim sql As String
            sql = "SELECT * FROM M_DeliveryDate "
            sql &= "Where Buyer  = '" & pBuyer & "' "
            sql &= "  And Active = 1 "
            sql &= "Order by DeliveryFlag, SpecialItem, SpecialItemBuyer, Action "
            Dim dt_Delivery As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_Delivery.Rows.Count - 1

                Select Case dt_Delivery.Rows(i).Item("DeliveryFlag")            ' @納期基準
                    '  -------------------------------------------------------------------------------------------------------
                    Case 0                                                      ' @0:系統

                        Select Case dt_Delivery.Rows(i).Item("SpecialItem")     ' #特殊ITEM納期
                            Case 0                                              ' #0:無
                                Select Case dt_Delivery.Rows(i).Item("Action")  ' %支援種類 大貨, 樣品
                                    Case 0                                      ' %0:Bulk & Sample
                                        If xSMPF = "" Then
                                            RtnString = DateAdd("d", dt_Delivery.Rows(i).Item("BulkDays"), Now).ToString("yyyyMMdd")
                                        Else
                                            RtnString = DateAdd("d", dt_Delivery.Rows(i).Item("SampleDays"), Now).ToString("yyyyMMdd")
                                        End If
                                    Case 1                                      ' %1:Bulk
                                        If xSMPF = "" Then
                                            RtnString = DateAdd("d", dt_Delivery.Rows(i).Item("BulkDays"), Now).ToString("yyyyMMdd")
                                        End If
                                    Case 2                                      ' %2:Sample
                                        If xSMPF = "" Then
                                        Else
                                            RtnString = DateAdd("d", dt_Delivery.Rows(i).Item("SampleDays"), Now).ToString("yyyyMMdd")
                                        End If
                                    Case Else                                   ' %其他
                                End Select

                            Case 1                                              ' #1:有

                                Select Case dt_Delivery.Rows(i).Item("Action")  ' %支援種類 大貨, 樣品
                                    Case 0                                      ' %0:Bulk & Sample
                                        If xSMPF = "" Then
                                            sql = "SELECT * FROM M_BothAgreeItem "
                                            sql &= "Where Buyer  = '" & dt_Delivery.Rows(i).Item("SpecialItemBuyer") & "' "
                                            sql &= "And   Season = '" & "ALL" & "' "
                                            sql &= "And   Code = '" & xITEM & "' "
                                            sql &= "And   Active = 1 "
                                            Dim dt_SpecialItem As DataTable = oDataBase.GetDataTable(sql)
                                            If dt_SpecialItem.Rows.Count > 0 Then
                                                RtnString = DateAdd("d", dt_SpecialItem.Rows(i).Item("Days"), Now).ToString("yyyyMMdd")
                                            Else
                                                RtnString = DateAdd("d", dt_Delivery.Rows(i).Item("BulkDays"), Now).ToString("yyyyMMdd")
                                            End If
                                        Else
                                            RtnString = DateAdd("d", dt_Delivery.Rows(i).Item("SampleDays"), Now).ToString("yyyyMMdd")
                                        End If
                                    Case 1                                      ' %1:Bulk
                                        If xSMPF = "" Then
                                            sql = "SELECT * FROM M_BothAgreeItem "
                                            sql &= "Where Buyer  = '" & dt_Delivery.Rows(i).Item("SpecialItemBuyer") & "' "
                                            sql &= "And   Season = '" & "ALL" & "' "
                                            sql &= "And   Code = '" & xITEM & "' "
                                            sql &= "And   Active = 1 "
                                            Dim dt_SpecialItem As DataTable = oDataBase.GetDataTable(sql)
                                            If dt_SpecialItem.Rows.Count > 0 Then
                                                RtnString = DateAdd("d", dt_SpecialItem.Rows(i).Item("Days"), Now).ToString("yyyyMMdd")
                                            Else
                                                RtnString = DateAdd("d", dt_Delivery.Rows(i).Item("BulkDays"), Now).ToString("yyyyMMdd")
                                            End If
                                        End If
                                    Case 2                                      ' %2:Sample
                                        If xSMPF = "" Then
                                        Else
                                            RtnString = DateAdd("d", dt_Delivery.Rows(i).Item("SampleDays"), Now).ToString("yyyyMMdd")
                                        End If
                                    Case Else                                   ' %其他
                                End Select
                            Case Else                                           ' #其他
                        End Select
                        '  -------------------------------------------------------------------------------------------------------
                    Case 1                                                      ' @1:客戶

                        Select Case dt_Delivery.Rows(i).Item("SpecialItem")     ' #特殊ITEM納期
                            Case 0                                              ' #0:無

                                Select Case dt_Delivery.Rows(i).Item("Action")  ' %支援種類 大貨, 樣品
                                    Case 0                                      ' %0:Bulk & Sample
                                        RtnString = xDeliveryDate
                                    Case 1                                      ' %1:Bulk
                                        If xSMPF = "" Then
                                            RtnString = xDeliveryDate
                                        End If
                                    Case 2                                      ' %2:Sample
                                        If xSMPF = "" Then
                                        Else
                                            RtnString = xDeliveryDate
                                        End If
                                    Case Else                                   ' %其他
                                End Select

                            Case 1                                              ' #1:有

                                Select Case dt_Delivery.Rows(i).Item("Action")  ' %支援種類 大貨, 樣品
                                    Case 0                                      ' %0:Bulk & Sample
                                        If xSMPF = "" Then
                                            sql = "SELECT * FROM M_BothAgreeItem "
                                            sql &= "Where Buyer  = '" & dt_Delivery.Rows(i).Item("SpecialItemBuyer") & "' "
                                            sql &= "And   Season = '" & "ALL" & "' "
                                            sql &= "And   Code = '" & xITEM & "' "
                                            sql &= "And   Active = 1 "
                                            Dim dt_SpecialItem As DataTable = oDataBase.GetDataTable(sql)
                                            If dt_SpecialItem.Rows.Count > 0 Then
                                                RtnString = DateAdd("d", dt_SpecialItem.Rows(i).Item("Days"), Now).ToString("yyyyMMdd")
                                            Else
                                                RtnString = xDeliveryDate
                                            End If
                                        Else
                                            RtnString = xDeliveryDate
                                        End If
                                    Case 1                                      ' %1:Bulk
                                        If xSMPF = "" Then
                                            sql = "SELECT * FROM M_BothAgreeItem "
                                            sql &= "Where Buyer  = '" & dt_Delivery.Rows(i).Item("SpecialItemBuyer") & "' "
                                            sql &= "And   Season = '" & "ALL" & "' "
                                            sql &= "And   Code = '" & xITEM & "' "
                                            sql &= "And   Active = 1 "
                                            Dim dt_SpecialItem As DataTable = oDataBase.GetDataTable(sql)
                                            If dt_SpecialItem.Rows.Count > 0 Then
                                                RtnString = DateAdd("d", dt_SpecialItem.Rows(i).Item("Days"), Now).ToString("yyyyMMdd")
                                            Else
                                                RtnString = xDeliveryDate
                                            End If
                                        End If
                                    Case 2                                      ' %2:Sample
                                        If xSMPF = "" Then
                                        Else
                                            RtnString = xDeliveryDate
                                        End If
                                    Case Else                                   ' %其他
                                End Select
                            Case Else                                           ' #其他
                        End Select
                        '  -------------------------------------------------------------------------------------------------------
                    Case Else                                                   ' @其他
                End Select
            Next
            '
            'ADD-START BY JOY 20220114 
            '*  JOY 20220602中止 
            '1/17 受注殘太高,強制調整納期 30 --> 55  
            'If Now.ToString("yyyyMMdd") >= "20220116" Then
            If Now.ToString("yyyyMMdd") >= "ZZZZZZZZ" Then
                sql = "SELECT "
                sql &= "*, (Case When BulkDays=30 then 55 else BulkDays End) As BulkDays220117 "
                sql &= "FROM M_DeliveryDate "
                sql &= "Where Buyer  = '" & pBuyer & "' "
                sql &= "  And Active = 1 "
                sql &= "Order by DeliveryFlag, SpecialItem, SpecialItemBuyer, Action "
                Dim dt_LongDelivery As DataTable = oDataBase.GetDataTable(sql)
                For i = 0 To dt_LongDelivery.Rows.Count - 1

                    Select Case dt_LongDelivery.Rows(i).Item("DeliveryFlag")            ' @納期基準
                        '  -------------------------------------------------------------------------------------------------------
                        Case 0                                                      ' @0:系統
                            Select Case dt_LongDelivery.Rows(i).Item("SpecialItem")     ' #特殊ITEM納期
                                Case 0                                              ' #0:無

                                    Select Case dt_LongDelivery.Rows(i).Item("Action")  ' %支援種類 大貨, 樣品
                                        Case 0                                      ' %0:Bulk & Sample
                                            If xSMPF = "" Then
                                                RtnString = DateAdd("d", dt_LongDelivery.Rows(i).Item("BulkDays220117"), Now).ToString("yyyyMMdd")
                                            Else
                                            End If
                                        Case 1                                      ' %1:Bulk
                                            If xSMPF = "" Then
                                                RtnString = DateAdd("d", dt_LongDelivery.Rows(i).Item("BulkDays220117"), Now).ToString("yyyyMMdd")
                                            End If
                                        Case Else                                   ' %其他
                                    End Select
                                Case Else                                           ' #其他
                            End Select
                        Case Else                                                   ' @其他
                    End Select
                Next
            End If
            'ADD-END
            '
            If RtnString = "" Then
                RtnString = "Error"
            Else
                If CDbl(RtnString) <= CDbl(Now.ToString("yyyyMMdd")) Or CDbl(RtnString) >= CDbl(DateAdd("d", 365, Now).ToString("yyyyMMdd")) Then
                    RtnString = "Error"
                End If
            End If
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "GetDeliveryDate", NowDateTime, 1, "Failed", "Code", "=", "9", pRule, xUserID, "")
        End Try
        '
        Return RtnString
    End Function
    'GetDeliveryDate-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([41C]-GetQtyUnitCode)
    '**     Get Unit Code
    '***********************************************************************************************
    'GetQtyUnitCode-Start
    Function GetQtyUnitCode(ByVal pBuyer As String, ByVal pRule As String, ByVal pUID As Integer) As String
        Dim RtnString As String = ""
        '
        Try
            'GetQtyUnitCode(FD(ITMC5K))==>FD(ITMC5K)
            Dim start As Integer = InStr(pRule, "(") + 1
            Dim length As Integer = Len(pRule) - start + 1
            Dim xCode As String = Mid(UCase(pRule), start, length)
            xCode = Command2Data(pBuyer, xCode, pUID)
            '
            '取得QTY Unit Code
            If xCode <> "" Then
                oWaves.Timeout = 900 * 1000
                RtnString = oWaves.GetItemQtyUnit(xCode)
            Else
                RtnString = "Error"
            End If
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "GetQtyUnitCode", NowDateTime, 1, "Failed", "Code", "=", "9", pRule, xUserID, "")
        End Try
        '
        Return RtnString
    End Function
    'GetQtyUnitCode-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([420]-Calculate)
    '**     解析Command及組合資料
    '***********************************************************************************************
    'Calculate-Start
    Function Calculate(ByVal pData As String, ByVal pMark As String, ByVal pData1 As String, ByVal pAction As String) As String
        Dim RtnString As String = ""
        '
        Try
            Select Case pMark
                Case "+"
                    If InStr(pAction, "DATE") > 0 Then
                        RtnString = DateAdd("d", CInt(pData1), Now).ToString("yyyyMMdd")
                    Else
                        RtnString = CStr(CInt(pData) + CInt(pData1))
                    End If
                Case "-"
                    RtnString = CStr(CInt(pData) - CInt(pData1))
                Case "*"
                    RtnString = CStr(CInt(pData) * CInt(pData1))
                Case "/"
                    RtnString = CStr(CInt(pData) / CInt(pData1))
                Case "|"
                    RtnString = pData + pData1
                Case Else
                    RtnString = "Error"
            End Select
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "Calculate", NowDateTime, 1, "Failed", "Code", "=", "9", "[" + pData + "][" + pData1 + "]", xUserID, "")
        End Try
        '
        Return RtnString
    End Function
    'Calculate-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([500]-ReplaceString)
    '**     解析後資料置換關鍵字
    '***********************************************************************************************
    'ReplaceString-Start
    Function ReplaceString(ByVal pData As String) As String
        Dim RtnString As String = pData
        '
        Try
            ' [~] → [-]
            If InStr(pData, "~") > 0 Then RtnString = Replace(RtnString, "~", "-")
            ' [{] → [(]
            If InStr(pData, "{") > 0 Then RtnString = Replace(RtnString, "{", "(")
            ' [}] → [)]
            If InStr(pData, "}") > 0 Then RtnString = Replace(RtnString, "}", ")")
            ' [^] → [/]
            If InStr(pData, "^") > 0 Then RtnString = Replace(RtnString, "^", "/")
            ' [!] → [+]
            If InStr(pData, "!") > 0 Then RtnString = Replace(RtnString, "!", "+")
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "ReplaceString", NowDateTime, 1, "Failed", "Code", "=", "9", pData, xUserID, "")
        End Try
        '
        Return RtnString
    End Function
    'ReplaceString-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([800]-WriteDataFile)
    '**     寫入Waves格式檔
    '***********************************************************************************************
    'WriteDataFile-Start
    Function WriteDataFile(ByVal pBuyer As String, ByVal pGRBuyer As String, ByVal pUID As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim xPODN As String = ""
        Dim xORFN As String = ""
        Dim xPUCN As String = ""
        Dim xITMC As String = ""
        Dim SQL, Str As String
        Dim i As Integer
        '------------------------------------------
        ' 2019/1/11 TNF 舊曆過年-希望納期變更
        Dim xDate As Date
        Dim xRADU As String = ""
        Dim xPRDU As String = ""
        Dim xRFCC As String = ""
        '------------------------------------------
        ' 2018/1/26 LINETYPE 0 --> 4 (詹寬宏)
        Dim xRCMC As String = ""
        Dim xLNGV As String = ""
        '------------------------------------------
        ' 2019/11/18 USAGE CODE --> 空白 (外林+蔡逸星)
        Dim xUSGC As String = ""
        '------------------------------------------
        ' 2020/7/03 SSO對應 AKEEP CODE變更
        Dim xAKPC As String = ""
        '------------------------------------------
        ' 2020/7/28 AWAS對應 DK2C5K DELIVERYCODE-2變更
        Dim xSQL As String = ""
        Dim xDK1C As String = ""
        Dim xDK2C As String = ""
        '------------------------------------------
        '
        Try
            '決定客戶Reference
            For i = 0 To 199
                'ITMC5K
                If xFieldList(i) = "ITMC5K" Then xITMC = xDataList(i)
                'PODN5K
                If xFieldList(i) = "PODN5K" Then xPODN = xDataList(i)
                'ORFN5K
                If xFieldList(i) = "ORFN5K" Then xORFN = xDataList(i)
                'PUCN5K
                If xFieldList(i) = "PUCN5K" Then xPUCN = xDataList(i)
                '------------------------------------------
                ' 2019/1/11 TNF 舊曆過年-希望納期變更
                'RADU5K
                If xFieldList(i) = "RADU5K" Then xRADU = xDataList(i)
                'PRDU5K
                If xFieldList(i) = "PRDU5K" Then xPRDU = xDataList(i)
                'RFCC5K
                If xFieldList(i) = "RFCC5K" Then xRFCC = xDataList(i)
                '------------------------------------------
                ' 2018/1/26 LINETYPE 0 --> 4 (詹寬宏)
                'RCMC5K
                If xFieldList(i) = "RCMC5K" Then xRCMC = Mid(xDataList(i), 1, 6)
                'LNGV5K
                If xFieldList(i) = "LNGV5K" Then xLNGV = xDataList(i)
                '------------------------------------------
                ' 2020/7/03 SSO對應 AKEEP CODE變更
                'AKPC5K
                If xFieldList(i) = "AKPC5K" Then xAKPC = xDataList(i)
                '------------------------------------------
                ' 2020/7/28 AWAS對應 DK2C5K DELIVERYCODE-2變更
                'DK1C5K
                If xFieldList(i) = "DK1C5K" Then xDK1C = xDataList(i)
                'DK2C5K
                If xFieldList(i) = "DK2C5K" Then xDK2C = xDataList(i)
                '------------------------------------------
            Next
            '[S5K00]
            SQL = "Insert into S5K00 "
            SQL &= "( "
            SQL &= "ITMC5K, "
            SQL &= "CLRC5K, "
            SQL &= "DPDC5K, "
            SQL &= "OTYC5K, "
            SQL &= "VNDC5K, "
            SQL &= "POSN5K, "
            SQL &= "PODN5K, "
            SQL &= "GRPC5K, "
            SQL &= "STLN5K, "
            SQL &= "PFMC5K, "
            SQL &= "PUCN5K, "
            SQL &= "CORN5K, "
            SQL &= "ORFN5K, "
            SQL &= "SBKN5K, "
            SQL &= "PODU5K, "
            SQL &= "POEU5K, "
            '@@@@@@@
            'SQL &= "PRDU5K, "

            SQL &= "PRDC5K, "
            SQL &= "RDRV5K, "
            SQL &= "CDLF5K, "
            SQL &= "MCNC5K, "
            SQL &= "MARC5K, "
            SQL &= "TRNC5K, "
            SQL &= "TTRC5K, "
            SQL &= "TCRC5K, "
            SQL &= "IDPR5K, "
            SQL &= "PMTC5K, "
            SQL &= "PTYC5K, "
            SQL &= "PD1V5K, "
            SQL &= "PD2V5K, "
            SQL &= "PD3V5K, "
            SQL &= "US1V5K, "
            SQL &= "US2V5K, "
            SQL &= "US3V5K, "
            SQL &= "PDOC5K, "
            SQL &= "RDPC5K, "
            SQL &= "PDDF5K, "
            SQL &= "DK1C5K, "
            SQL &= "DK2C5K, "
            SQL &= "RACC5K, "
            SQL &= "RVNC5K, "
            SQL &= "LUNC5K, "
            SQL &= "LNGV5K, "
            SQL &= "POOC5K, "
            SQL &= "AKPC5K, "
            SQL &= "ABKN5K, "
            SQL &= "BSBN5K, "
            SQL &= "PODQ5K, "
            SQL &= "PQUC5K, "
            SQL &= "PPVC5K, "
            SQL &= "PRTC5K, "
            SQL &= "PUNP5K, "
            SQL &= "TTYC5K, "
            SQL &= "SUNP5K, "
            SQL &= "DTCX5K, "
            SQL &= "NCMF5K, "
            SQL &= "SMPF5K, "
            SQL &= "USGC5K, "
            SQL &= "RFCC5K, "
            SQL &= "PRBF5K, "
            SQL &= "ERRC5K, "
            SQL &= "UIDC5K, "
            SQL &= "PRGC5K, "
            SQL &= "DEVC5K, "
            SQL &= "RADU5K, "
            SQL &= "RADT5K, "
            SQL &= "RUPU5K, "
            SQL &= "RUPT5K, "
            SQL &= "RCMC5K, "
            '@@@@@@@
            SQL &= "PRDU5K, "

            SQL &= "BUYER, "
            SQL &= "ID "
            SQL &= " ) "
            SQL &= "VALUES( "
            For i = 0 To 199
                If InStr(xFieldList(i), "5K") > 0 Then
                    '
                    '設定值變更
                    '------------------------------------------
                    ' 2014/7/1 WAVE'S變更
                    '   銷樣增加樣品區分
                    If xFieldList(i) = "SMPF5K" Then
                        If Trim(xDataList(i)) <> "" Then
                            SQL &= " '" & "H" & "', "
                        Else
                            '不變更
                            SQL &= " '" & xDataList(i) & "', "
                        End If
                    Else
                        If xFieldList(i) = "PRDU5K" Then
                            '@
                            '------------------------------------------
                            ' 2021/9/7 BP-納期管控-長納期對應
                            ' 對象:ALL
                            ' 期間:2021/9/7
                            '
                            ' 第2世代不支援
                            '@
                            '------------------------------------------(ADD-START BY JOY 2024/1/19)
                            '目的 : adidas 舊曆過年- EDI希望納期變更 (14D --> 21D)
                            'BUYER：adidas(000001) & reebok(000016) 
                            '客戶：   adidas(000001)  reebok(000016) 下單的成衣廠
                            'ORDER種類: 受注ORDER
                            'KEYIN 期間：2025/01/13~2025/1/24
                            '處理內容: KEY IN日＋21日（行事曆日） 
                            '20250113~20250124
                            '#
                            If (InStr(xRCMC, "000001") > 0) Or (InStr(xRCMC, "000016") > 0) Then
                                If InStr(pBuyer, "FCT-BY") <= 0 Then
                                    If xRADU >= "20250113" And xRADU <= "20250124" Then
                                        ' KEYIN日+21日
                                        xDate = CDate(Mid(xRADU, 1, 4) + "/" + Mid(xRADU, 5, 2) + "/" + Mid(xRADU, 7, 2))
                                        SQL &= " '" & DateAdd("d", 21, xDate).ToString("yyyyMMdd") & "', "
                                    Else
                                        SQL &= " '" & xDataList(i) & "', "      '不變更
                                    End If
                                Else
                                    '#
                                    SQL &= " '" & xDataList(i) & "', "          '不變更
                                End If
                            Else
                                '@
                                SQL &= " '" & xDataList(i) & "', "              '不變更
                            End If
                        Else
                            '------------------------------------------
                            ' 2016/1/13 半日化對應 PRDC5K(REQUEST DELIVERY DATE PUR) = 固定2
                            If xFieldList(i) = "PRDC5K" Then
                                SQL &= " '" & "2" & "', "
                            Else
                                '------------------------------------------
                                ' 2017/5/9 LINETYPE=4 改成 2
                                If xFieldList(i) = "PDOC5K" Then
                                    If Trim(xDataList(i)) = "4" Then
                                        SQL &= " '" & "2" & "', "
                                    Else
                                        '------------------------------------------
                                        ' 2018/1/26 MODIFY (詹寬宏)
                                        ' 對象:000001, 000016, 000062, TW0054, 000021
                                        ' 期間:2018/1/29~2018/6/30
                                        ' 內容:LINETYPE=0 改成 4
                                        '
                                        ' 2018/3/2  MODIFY (葉筱婷)
                                        ' 追加條件  LENGTH > 0 (LNGV5K) 
                                        '   4/3 BUG 追加條件  LENGTH <> 空白 (LNGV5K) 
                                        '
                                        ' 2018/6/29 MODIFY (詹寬宏)
                                        ' 延長有效期間至 2018/12/31
                                        '
                                        ' 2018/8/9 MODIFY (詹寬宏)
                                        ' 變更有效期間至 2018/08/14
                                        '
                                        If (Trim(xRCMC) = "000001" Or Trim(xRCMC) = "000016" Or Trim(xRCMC) = "000062" Or Trim(xRCMC) = "TW0054" Or Trim(xRCMC) = "000021") And _
                                           (xRADU >= "20180129" And xRADU <= "20180814") And _
                                           (Trim(xDataList(i)) = "0") And _
                                           (xLNGV > "0" And xLNGV <> "") Then
                                            SQL &= " '" & "4" & "', "
                                        Else
                                            '不變更
                                            SQL &= " '" & xDataList(i) & "', "
                                        End If
                                    End If
                                Else
                                    '------------------------------------------
                                    ' 2018/1/2 QTY只允許整數
                                    If xFieldList(i) = "PODQ5K" Then
                                        SQL &= " '" & CStr(CInt(xDataList(i))) & "', "
                                    Else
                                        '------------------------------------------
                                        ' 2019/11/18 USAGE CODE --> 空白 (外林+蔡逸星)
                                        ' 期間:2019/11/19 ~
                                        If xFieldList(i) = "USGC5K" And xRADU >= "20191119" Then
                                            SQL &= " '" & xUSGC & "', "
                                        Else
                                            '------------------------------------------
                                            ' 2020/7/03 SSO對應 AKEEP-CODE變更成空白 (高中島)
                                            ' 對象:000001, 000016
                                            ' 期間:2020/7/3 ~
                                            If xFieldList(i) = "AKPC5K" And (Trim(xRCMC) = "000001" Or Trim(xRCMC) = "000016") And xRADU >= "20200703" Then
                                                SQL &= " '" & "" & "', "
                                            Else
                                                '------------------------------------------
                                                ' 2020/7/28 AWAS對應 DK2C5K DELIVERYCODE-2變更 (YATO)
                                                ' 對象:ALL
                                                ' 期間:9999/99/99 ~
                                                If xFieldList(i) = "DK2C5K" And Trim(xRCMC) <> "ZZZZZZ" And xRADU >= "99999999" Then
                                                    xSQL = "Select DELIVERYCODE1, DELIVERYCODE2 From M_DELIVERY "
                                                    xSQL &= "Where DELIVERYCODE1 = '" & xDK1C & "' "
                                                    xSQL &= "And   DELIVERYCODE2 = '" & xDK2C & "' "
                                                    Dim dt_DELIVERY As DataTable = oDataBase.GetDataTable(SQL)
                                                    If dt_DELIVERY.Rows.Count > 0 Then
                                                        SQL &= " '" & "999" & "', "
                                                    Else
                                                        Str = Replace(xDataList(i), Chr(10), "")
                                                        Str = Replace(Str, Chr(13), "")
                                                        SQL &= " '" & Str & "', "
                                                    End If
                                                Else
                                                    '------------------------------------------
                                                    ' 2021/12/16 FALCON對應
                                                    ' 對象:ALL
                                                    ' 期間:20211220 ~
                                                    ' RACC5K (REQUEST ALLOCATE CONTROL)= SET 空白
                                                    If xFieldList(i) = "RACC5K" Then
                                                        SQL &= " '" & "" & "', "
                                                    Else
                                                        '------------------------------------------
                                                        '標準設有值
                                                        ' 對應解決 CR-LF CHR(13)-CHR(10) 
                                                        'SQL &= " '" & xDataList(i) & "', "
                                                        Str = Replace(xDataList(i), Chr(10), "")
                                                        Str = Replace(Str, Chr(13), "")
                                                        SQL &= " '" & Str & "', "
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            SQL &= " '" & pBuyer & "', "
            SQL &= " '" & CStr(pUID) & "' "
            SQL &= " ) "
            oDataBase.ExecuteNonQuery(SQL)
            '
            '[S5L00]
            SQL = "SELECT * FROM S5L00 "
            SQL &= "Where Buyer = '" & pBuyer & "' "
            If oDB.GetFunctionCode(pGRBuyer, 2) = "P" Then        '製作採購號碼 --> GRBuyer(2)=P 
                If oDB.GetFunctionCode(pGRBuyer, 5) = "R" Then    '[R]→PUCN5K + ITMC5K / [P]→ PUCN5K / [X]→ORFN5K+ITMC5K / [Y]→ORFN5K
                    SQL &= "  And ORFN   = '" & xPUCN + xITMC & "' "
                Else
                    If oDB.GetFunctionCode(pGRBuyer, 5) = "P" Then
                        SQL &= "  And ORFN   = '" & xPUCN & "' "
                    Else
                        ' MOD-START 2017/2/8
                        'SQL &= "  And ORFN   = '" & xORFN + xITMC & "' "
                        If oDB.GetFunctionCode(pGRBuyer, 5) = "X" Then
                            SQL &= "  And ORFN   = '" & xORFN + xITMC & "' "
                        Else
                            SQL &= "  And ORFN   = '" & xORFN & "' "
                        End If
                        ' MOD-END
                    End If
                End If
            Else
                SQL &= "  And PODN5L = '" & xPODN & "' "
            End If

            Dim dt_S5L00 As DataTable = oDataBase.GetDataTable(SQL)
            If dt_S5L00.Rows.Count <= 0 Then
                SQL = "Insert into S5L00 "
                SQL &= "( "
                SQL &= "DPDC5L, "
                SQL &= "PODN5L, "
                SQL &= "PO1X5L, "
                SQL &= "PO2X5L, "
                SQL &= "PO3X5L, "
                SQL &= "PO4X5L, "
                SQL &= "FN1I5L, "
                SQL &= "FN2I5L, "
                SQL &= "DL1A5L, "
                SQL &= "DL2A5L, "
                SQL &= "DL3A5L, "
                SQL &= "FD1I5L, "
                SQL &= "FD2I5L, "
                SQL &= "FD3I5L, "
                SQL &= "FM1X5L, "
                SQL &= "CM1C5L, "
                SQL &= "FM2X5L, "
                SQL &= "CM2C5L, "
                SQL &= "FM3X5L, "
                SQL &= "CM3C5L, "
                SQL &= "FM4X5L, "
                SQL &= "CM4C5L, "
                SQL &= "FM5X5L, "
                SQL &= "CM5C5L, "
                SQL &= "FM6X5L, "
                SQL &= "CM6C5L, "
                SQL &= "FM7X5L, "
                SQL &= "CM7C5L, "
                SQL &= "FM8X5L, "
                SQL &= "CM8C5L, "
                SQL &= "FM9X5L, "
                SQL &= "CM9C5L, "
                SQL &= "FMAX5L, "
                SQL &= "CMAC5L, "
                SQL &= "PCSN5L, "
                SQL &= "RFCC5L, "
                SQL &= "PRBF5L, "
                SQL &= "UIDC5L, "
                SQL &= "PRGC5L, "
                SQL &= "DEVC5L, "
                SQL &= "RADU5L, "
                SQL &= "RADT5L, "
                SQL &= "RUPU5L, "
                SQL &= "RUPT5L, "
                SQL &= "RCMC5L, "
                SQL &= "BUYER, "
                SQL &= "ORFN "
                SQL &= " ) "
                SQL &= "VALUES( "
                For i = 0 To 199
                    '**2021/12/14 MOD-START
                    '222 / 777 / ### / CB / 尺寸嚴 / 色嚴
                    'If InStr(xFieldList(i), "5L") > 0 Then SQL &= " '" & xDataList(i) & "', "
                    '
                    If InStr(xFieldList(i), "5L") > 0 Then
                        '
                        If xFieldList(i) = "PO1X5L" Or xFieldList(i) = "PO2X5L" Or xFieldList(i) = "PO3X5L" Or xFieldList(i) = "PO4X5L" Then
                            SQL &= " '" & ReplaceComment(xDataList(i)) & "', "
                        Else
                            SQL &= " '" & xDataList(i) & "', "
                        End If
                    End If
                    'MOD-END
                Next
                SQL &= " '" & pBuyer & "', "
                '
                If oDB.GetFunctionCode(pGRBuyer, 2) = "P" Then          '製作採購號碼 --> GRBuyer(2)=P 
                    If oDB.GetFunctionCode(pGRBuyer, 5) = "R" Then      '[R]→PUCN5K + ITMC5K / [P]→ PUCN5K / [X]→ORFN5K+ITMC5K / [Y]→ORFN5K
                        SQL &= " '" & xPUCN + xITMC & "' "
                    Else
                        If oDB.GetFunctionCode(pGRBuyer, 5) = "P" Then
                            SQL &= " '" & xPUCN & "' "
                        Else
                            ' MOD-START 2017/2/8
                            'SQL &= " '" & xORFN + xITMC & "' "
                            If oDB.GetFunctionCode(pGRBuyer, 5) = "X" Then
                                SQL &= " '" & xORFN + xITMC & "' "
                            Else
                                SQL &= " '" & xORFN & "' "
                            End If
                            ' MOD-END
                        End If
                    End If
                Else
                    SQL &= " '" & xPODN & "' "
                End If
                SQL &= " ) "
                oDataBase.ExecuteNonQuery(SQL)
            End If
            '
            '[SC760W1] NIKE VDP 專用
            If oDB.GetFunctionCode(pGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                SQL = "Insert into SC760W1 "
                SQL &= "( "
                SQL &= "ORDNW1, "
                SQL &= "OSBNW1, "
                SQL &= "CSTCW1, "
                SQL &= "SNCDW1, "
                SQL &= "SNYRW1, "
                SQL &= "SMTCW1, "
                SQL &= "ACRNW1, "
                SQL &= "ACR2W1, "
                SQL &= "ACR3W1, "
                SQL &= "ADCDW1, "
                SQL &= "BYMHW1, "
                SQL &= "BYMSW1, "
                SQL &= "NFCDW1, "
                SQL &= "NSNOW1, "
                SQL &= "NINOW1, "
                SQL &= "NCCDW1, "
                SQL &= "NCN1W1, "
                SQL &= "NCN2W1, "
                SQL &= "NCN3W1, "
                SQL &= "NCN4W1, "
                SQL &= "REFNW1, "
                SQL &= "CMCDW1, "
                SQL &= "CMMTW1, "
                SQL &= "NOICW1, "
                SQL &= "DVSCW1, "
                SQL &= "YBF1W1, "
                SQL &= "YBF2W1, "
                SQL &= "YBF3W1, "
                SQL &= "YBF4W1, "
                SQL &= "YBF5W1, "
                SQL &= "YBD1W1, "
                SQL &= "YBD2W1, "
                SQL &= "YBD3W1, "
                SQL &= "UIDCW1, "
                SQL &= "PRGCW1, "
                SQL &= "DEVCW1, "
                SQL &= "RADUW1, "
                SQL &= "RADTW1, "
                SQL &= "RUPUW1, "
                SQL &= "RUPTW1, "
                SQL &= "RCMCW1, "
                SQL &= "GRPCW1, "
                SQL &= "BUYER, "
                SQL &= "ORFN, "
                SQL &= "ID "
                SQL &= " ) "
                SQL &= "VALUES( "
                For i = 0 To 199
                    If InStr(xFieldList(i), "W1") > 0 Then
                        SQL &= " '" & xDataList(i) & "', "
                        'If xFieldList(i) = "REFNW1" Then
                        '    SQL &= " '" & xORFN & "', "
                        'Else
                        '    SQL &= " '" & xDataList(i) & "', "
                        'End If
                    End If
                Next
                SQL &= " '" & pBuyer & "', "
                '
                If oDB.GetFunctionCode(pGRBuyer, 2) = "P" Then        '製作採購號碼 --> GRBuyer(2)=P
                    If oDB.GetFunctionCode(pGRBuyer, 5) = "R" Then       '[R]→PUCN5K + ITMC5K / [P]→ PUCN5K / [X]→ORFN5K+ITMC5K / [Y]→ORFN5K
                        SQL &= " '" & xPUCN + xITMC & "', "
                    Else
                        If oDB.GetFunctionCode(pGRBuyer, 5) = "P" Then
                            SQL &= " '" & xPUCN & "', "
                        Else
                            ' MOD-START 2017/2/8
                            'SQL &= " '" & xORFN + xITMC & "', "
                            If oDB.GetFunctionCode(pGRBuyer, 5) = "X" Then
                                SQL &= " '" & xORFN + xITMC & "', "
                            Else
                                SQL &= " '" & xORFN & "', "
                            End If
                            ' MOD-END
                        End If
                    End If
                Else
                    SQL &= " '" & xPODN & "', "
                End If
                '
                SQL &= " '" & CStr(pUID) & "' "
                SQL &= " ) "
                oDataBase.ExecuteNonQuery(SQL)
            End If
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "WriteDataFile", NowDateTime, 1, "Failed", "Code", "=", "9", "", xUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'WriteDataFile-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([810]-WriteDataFileEOES)
    '**     寫入Waves格式檔
    '***********************************************************************************************
    'WriteDataFileEOES-Start
    Function WriteDataFileEOES(ByVal pBuyer As String, ByVal eBuyer As String, ByVal eGRBuyer As String, ByVal pUID As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim xPODN As String = ""
        Dim xORFN As String = ""
        Dim xPUCN As String = ""
        Dim xITMC As String = ""
        Dim SQL, Str As String
        Dim i As Integer
        '------------------------------------------
        ' 2019/1/11 TNF 舊曆過年-希望納期變更
        Dim xDate As Date
        Dim xRADU As String = ""
        Dim xPRDU As String = ""
        Dim xRFCC As String = ""
        '------------------------------------------
        ' 2018/1/26 LINETYPE 0 --> 4 (詹寬宏)
        Dim xRCMC As String = ""
        Dim xLNGV As String = ""
        '------------------------------------------
        ' 2019/11/18 USAGE CODE --> 空白 (外林+蔡逸星)
        Dim xUSGC As String = ""
        '------------------------------------------
        ' 2019/1/11 TNF 舊曆過年-希望納期變更
        ' 2020/7/03 SSO對應 AKEEP CODE變更
        Dim xAKPC As String = ""
        '------------------------------------------
        ' 2020/7/28 AWAS對應 DK2C5K DELIVERYCODE-2變更
        Dim xSQL As String = ""
        Dim xDK1C As String = ""
        Dim xDK2C As String = ""
        '------------------------------------------        '
        '------------------------------------------
        ' 2021/9/7 BP-納期管控-長納期對應
        Dim xUIDC As String = ""
        '------------------------------------------
        Try
            '決定客戶Reference
            For i = 0 To 199
                'ITMC5K
                If xFieldList(i) = "ITMC5K" Then xITMC = xDataList(i)
                'PODN5K
                If xFieldList(i) = "PODN5K" Then xPODN = xDataList(i)
                'ORFN5K
                If xFieldList(i) = "ORFN5K" Then xORFN = xDataList(i)
                'PUCN5K
                If xFieldList(i) = "PUCN5K" Then xPUCN = xDataList(i)
                '------------------------------------------
                ' 2019/1/11 TNF 舊曆過年-希望納期變更
                'RADU5K
                If xFieldList(i) = "RADU5K" Then xRADU = xDataList(i)
                'PRDU5K
                If xFieldList(i) = "PRDU5K" Then xPRDU = xDataList(i)
                'RFCC5K
                If xFieldList(i) = "RFCC5K" Then xRFCC = xDataList(i)
                '------------------------------------------
                ' 2018/1/26 LINETYPE 0 --> 4 (詹寬宏)
                'RCMC5K
                If xFieldList(i) = "RCMC5K" Then xRCMC = Mid(xDataList(i), 1, 6)
                'LNGV5K
                If xFieldList(i) = "LNGV5K" Then xLNGV = xDataList(i)
                '------------------------------------------
                ' 2020/7/03 SSO對應 AKEEP CODE變更
                'AKPC5K
                If xFieldList(i) = "AKPC5K" Then xAKPC = xDataList(i)
                '------------------------------------------
                ' 2020/7/28 AWAS對應 DK2C5K DELIVERYCODE-2變更
                'DK1C5K
                If xFieldList(i) = "DK1C5K" Then xDK1C = xDataList(i)
                'DK2C5K
                If xFieldList(i) = "DK2C5K" Then xDK2C = xDataList(i)
                '------------------------------------------
                ' 2021/9/7 BP-納期管控-長納期對應
                'UIDC5K
                If xFieldList(i) = "UIDC5K" Then xUIDC = xDataList(i)
                '------------------------------------------
            Next
            '[S5K00]
            SQL = "Insert into S5K00 "
            SQL &= "( "
            SQL &= "ITMC5K, "
            SQL &= "CLRC5K, "
            SQL &= "DPDC5K, "
            SQL &= "OTYC5K, "
            SQL &= "VNDC5K, "
            SQL &= "POSN5K, "
            SQL &= "PODN5K, "
            SQL &= "GRPC5K, "
            SQL &= "STLN5K, "
            SQL &= "PFMC5K, "
            SQL &= "PUCN5K, "
            SQL &= "CORN5K, "
            SQL &= "ORFN5K, "
            SQL &= "SBKN5K, "
            SQL &= "PODU5K, "
            SQL &= "POEU5K, "
            '@@@@@@@
            'SQL &= "PRDU5K, "

            SQL &= "PRDC5K, "
            SQL &= "RDRV5K, "
            SQL &= "CDLF5K, "
            SQL &= "MCNC5K, "
            SQL &= "MARC5K, "
            SQL &= "TRNC5K, "
            SQL &= "TTRC5K, "
            SQL &= "TCRC5K, "
            SQL &= "IDPR5K, "
            SQL &= "PMTC5K, "
            SQL &= "PTYC5K, "
            SQL &= "PD1V5K, "
            SQL &= "PD2V5K, "
            SQL &= "PD3V5K, "
            SQL &= "US1V5K, "
            SQL &= "US2V5K, "
            SQL &= "US3V5K, "
            SQL &= "PDOC5K, "
            SQL &= "RDPC5K, "
            SQL &= "PDDF5K, "
            SQL &= "DK1C5K, "
            SQL &= "DK2C5K, "
            SQL &= "RACC5K, "
            SQL &= "RVNC5K, "
            SQL &= "LUNC5K, "
            SQL &= "LNGV5K, "
            SQL &= "POOC5K, "
            SQL &= "AKPC5K, "
            SQL &= "ABKN5K, "
            SQL &= "BSBN5K, "
            SQL &= "PODQ5K, "
            SQL &= "PQUC5K, "
            SQL &= "PPVC5K, "
            SQL &= "PRTC5K, "
            SQL &= "PUNP5K, "
            SQL &= "TTYC5K, "
            SQL &= "SUNP5K, "
            SQL &= "DTCX5K, "
            SQL &= "NCMF5K, "
            SQL &= "SMPF5K, "
            SQL &= "USGC5K, "
            SQL &= "RFCC5K, "
            SQL &= "PRBF5K, "
            SQL &= "ERRC5K, "
            SQL &= "UIDC5K, "
            SQL &= "PRGC5K, "
            SQL &= "DEVC5K, "
            SQL &= "RADU5K, "
            SQL &= "RADT5K, "
            SQL &= "RUPU5K, "
            SQL &= "RUPT5K, "
            SQL &= "RCMC5K, "
            '@@@@@@@
            SQL &= "PRDU5K, "

            SQL &= "BUYER, "
            SQL &= "ID "
            SQL &= " ) "
            SQL &= "VALUES( "
            For i = 0 To 199
                If InStr(xFieldList(i), "5K") > 0 Then
                    '
                    '設定值變更
                    '------------------------------------------
                    ' 2014/7/1 WAVE'S變更
                    '   銷樣增加樣品區分
                    If xFieldList(i) = "SMPF5K" Then
                        If Trim(xDataList(i)) <> "" Then
                            SQL &= " '" & "H" & "', "
                        Else
                            '不變更
                            SQL &= " '" & xDataList(i) & "', "
                        End If
                    Else
                        If xFieldList(i) = "PRDU5K" Then
                            '@
                            '------------------------------------------
                            ' 2021/9/7 BP-納期管控-長納期對應
                            ' 對象:ALL
                            ' 期間:2021/9/7
                            '#
                            If xUIDC <> "" And IsNumeric(xUIDC) = True Then
                                'MsgBox(CStr(xUIDC) & ":" & CStr(xPRDU))
                                If xUIDC > xPRDU Then
                                    SQL &= " '" & xUIDC & "', "
                                Else
                                    SQL &= " '" & xDataList(i) & "', "      '不變更
                                End If
                            Else
                                '@
                                '------------------------------------------(ADD-START BY JOY 2024/1/19)
                                '目的 : adidas 舊曆過年- EDI希望納期變更 (14D --> 21D)
                                'BUYER：adidas(000001) & reebok(000016) 
                                '客戶：   adidas(000001)  reebok(000016) 下單的成衣廠
                                'ORDER種類: 受注ORDER
                                'KEYIN 期間：2025/01/13~2025/1/24
                                '處理內容: KEY IN日＋21日（行事曆日） 
                                '20250113~20250124
                                '#
                                If (InStr(eBuyer, "000001") > 0) Or (InStr(eBuyer, "000016") > 0) Then
                                    '#
                                    If InStr(pBuyer, "FCT-BY") <= 0 Then
                                        If xRADU >= "20250113" And xRADU <= "20250124" Then
                                            ' KEYIN日+21日
                                            xDate = CDate(Mid(xRADU, 1, 4) + "/" + Mid(xRADU, 5, 2) + "/" + Mid(xRADU, 7, 2))
                                            SQL &= " '" & DateAdd("d", 21, xDate).ToString("yyyyMMdd") & "', "
                                        Else
                                            SQL &= " '" & xDataList(i) & "', "      '不變更
                                        End If
                                    Else
                                        '#
                                        SQL &= " '" & xDataList(i) & "', "          '不變更
                                    End If
                                Else
                                    '@
                                    SQL &= " '" & xDataList(i) & "', "              '不變更
                                End If
                            End If
                        Else
                            '------------------------------------------
                            ' 2016/1/13 半日化對應 PRDC5K(REQUEST DELIVERY DATE PUR) = 固定2
                            If xFieldList(i) = "PRDC5K" Then
                                SQL &= " '" & "2" & "', "
                            Else
                                '------------------------------------------
                                ' 2017/5/9 LINETYPE=4 改成 2
                                If xFieldList(i) = "PDOC5K" Then
                                    If Trim(xDataList(i)) = "4" Then
                                        SQL &= " '" & "2" & "', "
                                    Else
                                        '------------------------------------------
                                        ' 2018/1/26 MODIFY (詹寬宏)
                                        ' 對象:000001, 000016, 000062, TW0054, 000021
                                        ' 期間:2018/1/29~2018/6/30
                                        ' 內容:LINETYPE=0 改成 4
                                        ' 2018/3/2  MODIFY (葉筱婷)
                                        ' 追加條件  LENGTH > 0 (LNGV5K) 
                                        '   4/3 BUG 追加條件  LENGTH <> 空白 (LNGV5K) 
                                        If (Trim(xRCMC) = "000001" Or Trim(xRCMC) = "000016" Or Trim(xRCMC) = "000062" Or Trim(xRCMC) = "TW0054" Or Trim(xRCMC) = "000021") And _
                                           (xRADU >= "20180129" And xRADU <= "20180630") And _
                                           (Trim(xDataList(i)) = "0") And _
                                           (xLNGV > "0" And xLNGV <> "") Then
                                            SQL &= " '" & "4" & "', "
                                        Else
                                            '不變更
                                            SQL &= " '" & xDataList(i) & "', "
                                        End If
                                    End If
                                Else
                                    '------------------------------------------
                                    ' 2018/1/2 QTY只允許整數
                                    If xFieldList(i) = "PODQ5K" Then
                                        SQL &= " '" & CStr(CInt(xDataList(i))) & "', "
                                    Else
                                        '------------------------------------------
                                        ' 2019/11/18 USAGE CODE --> 空白 (外林+蔡逸星)
                                        ' 期間:2019/11/19 ~
                                        If xFieldList(i) = "USGC5K" And xRADU >= "20191119" Then
                                            SQL &= " '" & xUSGC & "', "
                                        Else
                                            '------------------------------------------
                                            ' 2020/7/03 SSO對應 AKEEP-CODE變更成空白 (高中島)
                                            ' 對象:000001, 000016
                                            ' 期間:2020/7/3 ~
                                            If xFieldList(i) = "AKPC5K" And (Trim(xRCMC) = "000001" Or Trim(xRCMC) = "000016") And xRADU >= "20200703" Then
                                                SQL &= " '" & "" & "', "
                                            Else
                                                '------------------------------------------
                                                ' 2020/7/28 AWAS對應 DK2C5K DELIVERYCODE-2變更 (YATO)
                                                ' 對象:ALL
                                                ' 期間:9999/99/99 ~
                                                If xFieldList(i) = "DK2C5K" And Trim(xRCMC) <> "ZZZZZZ" And xRADU >= "99999999" Then
                                                    xSQL = "Select DELIVERYCODE1, DELIVERYCODE2 From M_DELIVERY "
                                                    xSQL &= "Where DELIVERYCODE1 = '" & xDK1C & "' "
                                                    xSQL &= "And   DELIVERYCODE2 = '" & xDK2C & "' "
                                                    Dim dt_DELIVERY As DataTable = oDataBase.GetDataTable(SQL)
                                                    If dt_DELIVERY.Rows.Count > 0 Then
                                                        SQL &= " '" & "999" & "', "
                                                    Else
                                                        Str = Replace(xDataList(i), Chr(10), "")
                                                        Str = Replace(Str, Chr(13), "")
                                                        SQL &= " '" & Str & "', "
                                                    End If
                                                Else
                                                    '------------------------------------------
                                                    ' 2021/12/16 FALCON對應
                                                    ' 對象:ALL
                                                    ' 期間:20211220 ~
                                                    ' RACC5K (REQUEST ALLOCATE CONTROL)= SET 空白
                                                    If xFieldList(i) = "RACC5K" Then
                                                        SQL &= " '" & "" & "', "
                                                    Else
                                                        '------------------------------------------
                                                        '標準設有值
                                                        ' 對應解決 CR-LF CHR(13)-CHR(10) 
                                                        'SQL &= " '" & xDataList(i) & "', "
                                                        Str = Replace(xDataList(i), Chr(10), "")
                                                        Str = Replace(Str, Chr(13), "")
                                                        SQL &= " '" & Str & "', "
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            SQL &= " '" & pBuyer & "', "
            SQL &= " '" & CStr(pUID) & "' "
            SQL &= " ) "
            oDataBase.ExecuteNonQuery(SQL)
            '
            '[S5L00]
            SQL = "SELECT * FROM S5L00 "
            SQL &= "Where Buyer = '" & pBuyer & "' "
            If oDB.GetFunctionCode(eGRBuyer, 2) = "P" Then        '製作採購號碼 --> GRBuyer(2)=P 
                If oDB.GetFunctionCode(eGRBuyer, 5) = "R" Then    '[R]→PUCN5K + ITMC5K / [P]→ PUCN5K / [X]→ORFN5K+ITMC5K / [Y]→ORFN5K
                    SQL &= "  And ORFN   = '" & xPUCN + xITMC & "' "
                Else
                    If oDB.GetFunctionCode(eGRBuyer, 5) = "P" Then
                        SQL &= "  And ORFN   = '" & xPUCN & "' "
                    Else
                        ' MOD-START 2017/2/8
                        'SQL &= "  And ORFN   = '" & xORFN + xITMC & "' "
                        If oDB.GetFunctionCode(eGRBuyer, 5) = "X" Then
                            SQL &= "  And ORFN   = '" & xORFN + xITMC & "' "
                        Else
                            SQL &= "  And ORFN   = '" & xORFN & "' "
                        End If
                        ' MOD-END
                    End If
                End If
            Else
                SQL &= "  And PODN5L = '" & xPODN & "' "
            End If

            Dim dt_S5L00 As DataTable = oDataBase.GetDataTable(SQL)
            If dt_S5L00.Rows.Count <= 0 Then
                SQL = "Insert into S5L00 "
                SQL &= "( "
                SQL &= "DPDC5L, "
                SQL &= "PODN5L, "
                SQL &= "PO1X5L, "
                SQL &= "PO2X5L, "
                SQL &= "PO3X5L, "
                SQL &= "PO4X5L, "
                SQL &= "FN1I5L, "
                SQL &= "FN2I5L, "
                SQL &= "DL1A5L, "
                SQL &= "DL2A5L, "
                SQL &= "DL3A5L, "
                SQL &= "FD1I5L, "
                SQL &= "FD2I5L, "
                SQL &= "FD3I5L, "
                SQL &= "FM1X5L, "
                SQL &= "CM1C5L, "
                SQL &= "FM2X5L, "
                SQL &= "CM2C5L, "
                SQL &= "FM3X5L, "
                SQL &= "CM3C5L, "
                SQL &= "FM4X5L, "
                SQL &= "CM4C5L, "
                SQL &= "FM5X5L, "
                SQL &= "CM5C5L, "
                SQL &= "FM6X5L, "
                SQL &= "CM6C5L, "
                SQL &= "FM7X5L, "
                SQL &= "CM7C5L, "
                SQL &= "FM8X5L, "
                SQL &= "CM8C5L, "
                SQL &= "FM9X5L, "
                SQL &= "CM9C5L, "
                SQL &= "FMAX5L, "
                SQL &= "CMAC5L, "
                SQL &= "PCSN5L, "
                SQL &= "RFCC5L, "
                SQL &= "PRBF5L, "
                SQL &= "UIDC5L, "
                SQL &= "PRGC5L, "
                SQL &= "DEVC5L, "
                SQL &= "RADU5L, "
                SQL &= "RADT5L, "
                SQL &= "RUPU5L, "
                SQL &= "RUPT5L, "
                SQL &= "RCMC5L, "
                SQL &= "BUYER, "
                SQL &= "ORFN "
                SQL &= " ) "
                SQL &= "VALUES( "
                For i = 0 To 199
                    '**2021/12/14 MOD-START
                    '222 / 777 / ### / CB / 尺寸嚴 / 色嚴
                    'If InStr(xFieldList(i), "5L") > 0 Then SQL &= " '" & xDataList(i) & "', "
                    '
                    If InStr(xFieldList(i), "5L") > 0 Then
                        '
                        If xFieldList(i) = "PO1X5L" Or xFieldList(i) = "PO2X5L" Or xFieldList(i) = "PO3X5L" Or xFieldList(i) = "PO4X5L" Then
                            SQL &= " '" & ReplaceComment(xDataList(i)) & "', "
                        Else
                            SQL &= " '" & xDataList(i) & "', "
                        End If
                    End If
                    'MOD-END
                Next
                SQL &= " '" & pBuyer & "', "
                '
                If oDB.GetFunctionCode(eGRBuyer, 2) = "P" Then          '製作採購號碼 --> GRBuyer(2)=P 
                    If oDB.GetFunctionCode(eGRBuyer, 5) = "R" Then      '[R]→PUCN5K + ITMC5K / [P]→ PUCN5K / [X]→ORFN5K+ITMC5K / [Y]→ORFN5K
                        SQL &= " '" & xPUCN + xITMC & "' "
                    Else
                        If oDB.GetFunctionCode(eGRBuyer, 5) = "P" Then
                            SQL &= " '" & xPUCN & "' "
                        Else
                            ' MOD-START 2017/2/8
                            'SQL &= " '" & xORFN + xITMC & "' "
                            If oDB.GetFunctionCode(eGRBuyer, 5) = "X" Then
                                SQL &= " '" & xORFN + xITMC & "' "
                            Else
                                SQL &= " '" & xORFN & "' "
                            End If
                            ' MOD-END
                        End If
                    End If
                Else
                    SQL &= " '" & xPODN & "' "
                End If
                SQL &= " ) "
                oDataBase.ExecuteNonQuery(SQL)
            End If
            '
            '[SC760W1] NIKE VDP 專用
            If oDB.GetFunctionCode(eGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                SQL = "Insert into SC760W1 "
                SQL &= "( "
                SQL &= "ORDNW1, "
                SQL &= "OSBNW1, "
                SQL &= "CSTCW1, "
                SQL &= "SNCDW1, "
                SQL &= "SNYRW1, "
                SQL &= "SMTCW1, "
                SQL &= "ACRNW1, "
                SQL &= "ACR2W1, "
                SQL &= "ACR3W1, "
                SQL &= "ADCDW1, "
                SQL &= "BYMHW1, "
                SQL &= "BYMSW1, "
                SQL &= "NFCDW1, "
                SQL &= "NSNOW1, "
                SQL &= "NINOW1, "
                SQL &= "NCCDW1, "
                SQL &= "NCN1W1, "
                SQL &= "NCN2W1, "
                SQL &= "NCN3W1, "
                SQL &= "NCN4W1, "
                SQL &= "REFNW1, "
                SQL &= "CMCDW1, "
                SQL &= "CMMTW1, "
                SQL &= "NOICW1, "
                SQL &= "DVSCW1, "
                SQL &= "YBF1W1, "
                SQL &= "YBF2W1, "
                SQL &= "YBF3W1, "
                SQL &= "YBF4W1, "
                SQL &= "YBF5W1, "
                SQL &= "YBD1W1, "
                SQL &= "YBD2W1, "
                SQL &= "YBD3W1, "
                SQL &= "UIDCW1, "
                SQL &= "PRGCW1, "
                SQL &= "DEVCW1, "
                SQL &= "RADUW1, "
                SQL &= "RADTW1, "
                SQL &= "RUPUW1, "
                SQL &= "RUPTW1, "
                SQL &= "RCMCW1, "
                SQL &= "GRPCW1, "
                SQL &= "BUYER, "
                SQL &= "ORFN, "
                SQL &= "ID "
                SQL &= " ) "
                SQL &= "VALUES( "
                For i = 0 To 199
                    If InStr(xFieldList(i), "W1") > 0 Then
                        SQL &= " '" & xDataList(i) & "', "
                        'If xFieldList(i) = "REFNW1" Then
                        '    SQL &= " '" & xORFN & "', "
                        'Else
                        '    SQL &= " '" & xDataList(i) & "', "
                        'End If
                    End If
                Next
                SQL &= " '" & pBuyer & "', "
                '
                If oDB.GetFunctionCode(eGRBuyer, 2) = "P" Then        '製作採購號碼 --> GRBuyer(2)=P
                    If oDB.GetFunctionCode(eGRBuyer, 5) = "R" Then       '[R]→PUCN5K + ITMC5K / [P]→ PUCN5K / [X]→ORFN5K+ITMC5K / [Y]→ORFN5K
                        SQL &= " '" & xPUCN + xITMC & "', "
                    Else
                        If oDB.GetFunctionCode(eGRBuyer, 5) = "P" Then
                            SQL &= " '" & xPUCN & "', "
                        Else
                            ' MOD-START 2017/2/8
                            'SQL &= " '" & xORFN + xITMC & "', "
                            If oDB.GetFunctionCode(eGRBuyer, 5) = "X" Then
                                SQL &= " '" & xORFN + xITMC & "', "
                            Else
                                SQL &= " '" & xORFN & "', "
                            End If
                            ' MOD-END
                        End If
                    End If
                Else
                    SQL &= " '" & xPODN & "', "
                End If
                '
                SQL &= " '" & CStr(pUID) & "' "
                SQL &= " ) "
                oDataBase.ExecuteNonQuery(SQL)
            End If
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "WriteDataFile", NowDateTime, 1, "Failed", "Code", "=", "9", "", xUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'WriteDataFileEOES-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([700]-CheckDataError)
    '**     檢測是否有資料解析異常
    '***********************************************************************************************
    'CheckDataError-Start
    Function CheckDataError(ByVal pLogID As String) As Integer
        Dim SQL As String
        Dim RtnCode As Integer = 0
        Try
            SQL = "SELECT * FROM L_ActionHistory "
            SQL &= "Where LogID = '" & pLogID & "' "
            SQL &= "  And Error = '1' "
            Dim dt_Error As DataTable = oDataBase.GetDataTable(SQL)
            If dt_Error.Rows.Count > 0 Then RtnCode = 1
        Catch ex As Exception
        End Try
        '
        Return RtnCode
    End Function
    'CheckDataError-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([710]-ReplaceComment)
    '**      OrderComment置換特別字串
    '**      2021/12/14 ADD
    '***********************************************************************************************
    'ReplaceComment-Start
    Function ReplaceComment(ByVal pComment As String) As String
        Dim RtnString As String = pComment
        Dim str As String = pComment
        '
        Try
            '222 / 777 / ### / CB / 尺寸嚴 / 色嚴 / 尺吋嚴 / 品質嚴 / 呎吋嚴
            '
            str = Replace(str, "#CB 尺寸嚴 色嚴", "")
            str = Replace(str, "#CB尺寸嚴 色嚴", "")
            str = Replace(str, "#CB尺寸嚴色嚴", "")
            str = Replace(str, "尺寸嚴 色嚴", "")
            str = Replace(str, "尺寸嚴色嚴", "")
            str = Replace(str, "色嚴 尺寸嚴", "")
            str = Replace(str, "色嚴尺寸嚴", "")
            str = Replace(str, "尺吋嚴", "")
            str = Replace(str, "品質嚴", "")
            str = Replace(str, "呎吋嚴", "")
            str = Replace(str, "尺寸嚴", "")
            str = Replace(str, "色嚴", "")
            '
            str = Replace(str, "222", "")
            str = Replace(str, "777", "")
            str = Replace(str, "###", "")
            str = Replace(str, "CB ", "")
            '
            RtnString = str
        Catch ex As Exception
        End Try
        '
        Return RtnString
    End Function
    'ReplaceComment-End
    '
    '===================================================================================================
    ' MEMO
    ' 
    ' EOES - PRDU5K
    ''@
    ''------------------------------------------
    '' 目的 : TNF 舊曆過年-希望納期變更
    '' BUYER：TNF  ( 000021 )
    '' 客戶：QMI  (W9520, W9523)
    '' ORDER種類： 受注ORDER 
    '' KEYIN 期間：2019/1/24~2019/1/31
    '' 處理內容：KEYIN日＋２1日（行事曆日）
    'If (Trim(xRFCC) = "W9520" And InStr(eBuyer, "000021") > 0) Or (Trim(xRFCC) = "W9523" And InStr(eBuyer, "000021") > 0) Then
    '    '#
    '    If InStr(eBuyer, "FCT-BY") <= 0 Then
    '        If xRADU >= "20190124" And xRADU <= "20190131" Then
    '            ' KEYIN日+21日
    '            xDate = CDate(Mid(xRADU, 1, 4) + "/" + Mid(xRADU, 5, 2) + "/" + Mid(xRADU, 7, 2))
    '            SQL &= " '" & DateAdd("d", 21, xDate).ToString("yyyyMMdd") & "', "
    '        Else
    '            SQL &= " '" & xDataList(i) & "', "      '不變更
    '        End If
    '    Else
    '        '#
    '        SQL &= " '" & xDataList(i) & "', "          '不變更
    '    End If
    'Else
    '    '@
    '    '------------------------------------------
    '    ' 目的 : ADIDAS 舊曆過年-希望納期變更
    '    ' BUYER：ADIDAS(  000001)  REEBOK  (000016)
    '    ' 客戶： ALL
    '    ' ORDER種類： 受注ORDER 
    '    ' KEYIN 期間：2021/2/1~2021/2/9
    '    ' 處理內容:KEYIN日＋２1日（行事曆日）
    '    If (InStr(eBuyer, "000001") > 0) Or (InStr(eBuyer, "000016") > 0) Then
    '        '#
    '        If InStr(pBuyer, "FCT-BY") <= 0 Then
    '            If xRADU >= "20210201" And xRADU <= "20210209" Then
    '                ' KEYIN日+21日
    '                xDate = CDate(Mid(xRADU, 1, 4) + "/" + Mid(xRADU, 5, 2) + "/" + Mid(xRADU, 7, 2))
    '                SQL &= " '" & DateAdd("d", 21, xDate).ToString("yyyyMMdd") & "', "
    '            Else
    '                SQL &= " '" & xDataList(i) & "', "      '不變更
    '            End If
    '        Else
    '            '#
    '            SQL &= " '" & xDataList(i) & "', "          '不變更
    '        End If
    '    Else
    '        '@
    '        SQL &= " '" & xDataList(i) & "', "              '不變更
    '    End If
    '    ''@
    '    ''------------------------------------------
    '    '' 目的 : ADIDAS 舊曆過年-希望納期變更
    '    '' BUYER：ADIDAS(  000001)  REEBOK  (000016)
    '    '' 客戶： ALL
    '    '' ORDER種類： 受注ORDER 
    '    '' KEYIN 期間：2020/1/15~2020/1/22
    '    '' 處理內容:KEYIN日＋２1日（行事曆日）
    '    'If (InStr(eBuyer, "000001") > 0) Or (InStr(eBuyer, "000016") > 0) Then
    '    '    '#
    '    '    If InStr(pBuyer, "FCT-BY") <= 0 Then
    '    '        If xRADU >= "20200115" And xRADU <= "20200122" Then
    '    '            ' KEYIN日+21日
    '    '            xDate = CDate(Mid(xRADU, 1, 4) + "/" + Mid(xRADU, 5, 2) + "/" + Mid(xRADU, 7, 2))
    '    '            SQL &= " '" & DateAdd("d", 21, xDate).ToString("yyyyMMdd") & "', "
    '    '        Else
    '    '            SQL &= " '" & xDataList(i) & "', "      '不變更
    '    '        End If
    '    '    Else
    '    '        '#
    '    '        SQL &= " '" & xDataList(i) & "', "          '不變更
    '    '    End If
    '    'Else
    '    '    '@
    '    '    SQL &= " '" & xDataList(i) & "', "              '不變更
    '    'End If
    'End If
    '
    ' 
    ' 2世代 - PRDU5K
    ''@
    ''------------------------------------------
    '' 目的 : TNF 舊曆過年-希望納期變更
    '' BUYER：TNF  ( 000021 )
    '' 客戶：QMI  (W9520, W9523)
    '' ORDER種類： 受注ORDER 
    '' KEYIN 期間：2019/1/24~2019/1/31
    '' 處理內容：KEYIN日＋２1日（行事曆日）
    'If (Trim(xRFCC) = "W9520" And Trim(xRCMC) = "000021") Or (Trim(xRFCC) = "W9523" And Trim(xRCMC) = "000021") Then
    '    '#
    '    If InStr(pBuyer, "FCT-BY") <= 0 Then
    '        If xRADU >= "20190124" And xRADU <= "20190131" Then
    '            ' KEYIN日+21日
    '            xDate = CDate(Mid(xRADU, 1, 4) + "/" + Mid(xRADU, 5, 2) + "/" + Mid(xRADU, 7, 2))
    '            SQL &= " '" & DateAdd("d", 21, xDate).ToString("yyyyMMdd") & "', "
    '        Else
    '            SQL &= " '" & xDataList(i) & "', "      '不變更
    '        End If
    '    Else
    '        '#
    '        SQL &= " '" & xDataList(i) & "', "          '不變更
    '    End If
    'Else
    '    '@
    '    '------------------------------------------
    '    ' 目的 : ADIDAS 舊曆過年-希望納期變更
    '    ' BUYER：ADIDAS(  000001)  REEBOK  (000016)
    '    ' 客戶： ALL
    '    ' ORDER種類： 受注ORDER 
    '    ' KEYIN 期間：2021/2/1~2021/2/9
    '    ' 處理內容:KEYIN日＋２1日（行事曆日）
    '    'If (Trim(xRCMC) = "000001") Or (Trim(xRCMC) = "000016") Then
    '    If (InStr(xRCMC, "000001") > 0) Or (InStr(xRCMC, "000016") > 0) Then

    '        '#
    '        If InStr(pBuyer, "FCT-BY") <= 0 Then
    '            If xRADU >= "20210201" And xRADU <= "20210209" Then
    '                ' KEYIN日+21日
    '                xDate = CDate(Mid(xRADU, 1, 4) + "/" + Mid(xRADU, 5, 2) + "/" + Mid(xRADU, 7, 2))
    '                SQL &= " '" & DateAdd("d", 21, xDate).ToString("yyyyMMdd") & "', "
    '            Else
    '                SQL &= " '" & xDataList(i) & "', "      '不變更
    '            End If
    '        Else
    '            '#
    '            SQL &= " '" & xDataList(i) & "', "          '不變更
    '        End If
    '    Else
    '        '@
    '        SQL &= " '" & xDataList(i) & "', "              '不變更
    '    End If
    'End If
    '
    '
End Class
