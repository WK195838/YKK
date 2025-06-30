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
Public Class FMappingService
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
    Dim oFASCommon As New FCommonService
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
    '	+	    +               +----[41A]-(GETKEEPCODE)GetKeepCodeData             | 取得KeepCode	
    '	+	    +               +							                        |
    '	+	    +               +----[41B]-(GETITEMNAME)GetItemName                 | 取得ItemName	
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
                              ByVal pFunList As String) As Integer
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
                    RtnCode = WriteDataFile(pBuyer, pFunList, dt_InputSheet.Rows(i).Item("Unique_ID"))
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
            SQL &= "Where Buyer = '" & pBuyer & "-1' "
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
            SQL &= "Where Buyer = '" & "YYYYYY" & "' "
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
                Case "GETKEEPCODE"                                          ' 取得KeepCode
                    RtnString = GetKeepCodeData(pBuyer, pRule, pUID)
                Case "GETITEMNAME"                                          ' 取得ItemName
                    RtnString = GetItemName(pBuyer, pRule, pUID)
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
            If InStr(pRule, "SP") > 0 Then
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
            'MsgBox(pRule)
            start = InStr(pRule, "(") + 1
            length = InStr(pRule, ",") - start
            Dim xBuyer = Mid(UCase(pRule), start, length)
            xBuyer = Replace(Command2Data(pBuyer, xBuyer, pUID), "~", "-")
            Dim Color As Object = Split(xColor + "/", "/")        ' Color(xxx/xxx/xxxxxx/xxx/xxx)
            '
            'MsgBox(xBuyer)
            Dim sql As String
            sql = "SELECT YCode FROM M_ItemConvert "
            sql &= "Where Buyer = '" & xBuyer & "' "
            sql &= "  And Code  = '" & xCode & "' "
            sql &= "  And Color1 = '" & Color(2) & "' "
            sql &= "  And Color2 = '" & Color(3) & "' "
            sql &= "  And Color3 = '" & Color(4) & "' "
            sql &= "  And SliderStatus = '" & xSlider & "' "
            '
            'MsgBox(sql)
            '
            Dim dt_Item As DataTable = oDataBase.GetDataTable(sql)
            If dt_Item.Rows.Count > 0 Then
                If oFASCommon.CheckItemNoDisplay(dt_Item.Rows(0).Item("YCode")) = 0 Then
                    RtnString = dt_Item.Rows(0).Item("YCode")
                    'MsgBox(RtnString)
                Else
                    RtnString = "Error"
                End If
            Else
                'JOY-DEBUG-START TNF 2014/10/20 
                'RtnString = "Error"
                'If InStr(pBuyer, "000021") > 0 Then
                '    RtnString = "1579172"
                'End If
                'ORI-SOURCE
                RtnString = "Error"
                'JOY-DEBUG-END TNF
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
            'Dim xChain As String = ""
            'sql = "Select * From M_Referp "
            'sql = sql & " Where Cat  = '" & "113" & "' "
            'sql = sql & "   And DKey = '" & "COLOR" & "' "
            'Dim dt_Referp As DataTable = oDataBase.GetDataTable(sql)
            'If dt_Referp.Rows.Count > 0 Then
            '    xChain = dt_Referp.Rows(0).Item("Data")
            'End If
            ''
            'oWaves.Timeout = 900 * 1000
            'Dim xPartType As String = oWaves.GetSpecialChain(xCode, xChain)
            '
            Dim sql As String
            sql = "SELECT YColor FROM M_ColorConvert "
            sql &= "Where Buyer  = '" & xBuyer & "' "
            sql &= "  And Season = '" & xSeason & "' "
            sql &= "  And Color1 = '" & Color(0) & "' "
            sql &= "  And Color2 = '" & Color(1) & "' "
            sql &= "  And Green  = '" & xCode & "' "
            Dim dt_Color As DataTable = oDataBase.GetDataTable(sql)
            If dt_Color.Rows.Count > 0 Then
                If oFASCommon.CheckColorNotFound(dt_Color.Rows(0).Item("YColor")) = 0 Then
                    RtnString = dt_Color.Rows(0).Item("YColor")
                Else
                    RtnString = "Error"
                End If
            Else
                'JOY-DEBUG-START TNF 2014/10/20 
                'RtnString = "Error"
                'If InStr(pBuyer, "000021") > 0 Then
                '    RtnString = "  580"
                'End If
                'ORI-SOURCE
                RtnString = "Error"
                'JOY-DEBUG-END TNF
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
    '**([41B]-GetItemName)
    '**     取得ItemName
    '***********************************************************************************************
    'GetItemName-Start
    Function GetItemName(ByVal pBuyer As String, ByVal pRule As String, ByVal pUID As Integer) As String
        Dim RtnString As String = ""
        '
        Try
            'GETITEMNAME(FD(Y_ItemCode),1)==>FD(Y_ItemCode),1
            Dim start As Integer = InStr(pRule, "(") + 1
            Dim length As Integer = Len(pRule) - start
            Dim str As String = Mid(UCase(pRule), start, length)

            'FD(Y_ItemCode),1 ==>FD(Y_ItemCode)
            start = 1
            length = InStr(str, ",") - 1
            Dim xCode As String = Mid(str, start, length)
            xCode = Command2Data(pBuyer, xCode, pUID)

            'FD(Y_ItemCode),1 ==>1
            start = InStr(str, ",") + 1
            length = Len(str) - start + 1
            Dim idx As Integer = CInt(Mid(str, start, length))

            oWaves.Timeout = 900 * 1000
            oWaves.GetItemName001(xCode, idx, RtnString)

        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "GetItemName", NowDateTime, 1, "Failed", "Code", "=", "9", pRule, xUserID, "")
        End Try
        '
        Return RtnString
    End Function
    'GetItemName-End
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
    Function WriteDataFile(ByVal pBuyer As String, ByVal pFunList As String, ByVal pUID As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim SQL As String
        Dim i As Integer
        '
        Try
            '
            '[FCT PLAN]
            SQL = "Insert into ForcastPlan "
            SQL &= "( "
            SQL &= "BuyerGroup, "
            SQL &= "FCTNo, "
            SQL &= "FCTSubNo, "
            SQL &= "Version, "
            SQL &= "C_Code, "
            SQL &= "C_Color, "
            SQL &= "C_SpecialRequest, "
            SQL &= "C_Season, "
            SQL &= "C_ShortenLT, "
            SQL &= "C_Style, "
            SQL &= "C_A1, "
            SQL &= "C_B1, "
            SQL &= "C_C1, "
            SQL &= "C_D1, "
            SQL &= "C_E1, "
            SQL &= "C_F1, "
            SQL &= "C_G1, "
            SQL &= "C_H1, "
            SQL &= "C_I1, "
            SQL &= "C_J1, "
            SQL &= "C_K1, "
            SQL &= "C_L1, "
            SQL &= "C_M1, "
            SQL &= "C_N1, "
            SQL &= "C_O1, "

            SQL &= "Y_LEVEL, "
            SQL &= "Y_ItemCode, "
            SQL &= "Y_ItemName1, "
            SQL &= "Y_ItemName2, "
            SQL &= "Y_ItemName3, "
            SQL &= "Y_Color, "
            SQL &= "Y_A1, "
            SQL &= "Y_B1, "
            SQL &= "Y_C1, "
            SQL &= "Y_D1, "
            SQL &= "Y_E1, "
            SQL &= "Y_F1, "
            SQL &= "Y_G1, "
            SQL &= "Y_H1, "
            SQL &= "Y_I1, "
            SQL &= "Y_J1, "

            SQL &= "Total, "
            SQL &= "N_F, "
            SQL &= "N1_F, "
            SQL &= "N2_F, "
            SQL &= "N3_F, "
            SQL &= "N4_F, "
            SQL &= "N5_F, "
            SQL &= "N6_F, "
            SQL &= "N7_F, "
            SQL &= "N8_F, "
            SQL &= "N9_F, "
            SQL &= "N10_F, "
            SQL &= "N11_F, "
            SQL &= "N12_F, "

            SQL &= "BUYER, "
            SQL &= "ID, "
            SQL &= "CreateUser, "
            SQL &= "CreateTime "
            SQL &= " ) "
            SQL &= "VALUES( "
            For i = 0 To 199
                If xFieldList(i) <> "" Then
                    SQL &= " '" & xDataList(i) & "', "
                End If
            Next
            SQL &= " '" & pBuyer & "', "
            SQL &= " '" & CStr(pUID) & "', "
            SQL &= " '" & "Convert" & "', "
            SQL &= " '" & NowDateTime & "' "
            SQL &= " ) "
            oDataBase.ExecuteNonQuery(SQL)
            '
        Catch ex As Exception
            oDB.AccessLog(xLogID, xBuyer, "WriteDataFile", NowDateTime, 1, "Failed", "Code", "=", "9", "", xUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'WriteDataFile-End

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

End Class
