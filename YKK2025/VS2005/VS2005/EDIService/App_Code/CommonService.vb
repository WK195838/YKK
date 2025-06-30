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
Public Class CommonService
     Inherits System.Web.Services.WebService

    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '** 全域變數
    '***********************************************************************************************
    '全域變數-Start
    Dim oDB As New ForProject
    Dim oDataBase As Utility.DataBase = oDB.GetDataBaseObj()
    Dim oMapping As MappingService
    Dim oWavesEDI As New Waves.EDIService
    Dim oWaves As New Waves.CommonService
    Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")     ' 現在日時    
    '	+-----------------------------------------------------------------------------------------------------------------
    '[000]-Service				    				                                |  
    '	+									                                        |
    '	+----[100]-MakePONO         						                        | 製作採購號碼
    '	+									                                        |
    '	+----[110]-CheckPONO        						                        | 檢測採購號碼
    '	+									                                        |
    '	+----[120]-MakeGRPC         						                        | 製作Group Code
    '	+									                                        |
    '	+----[130]-CheckGRPC									                    | 檢測Group Code
    '	+									                                        |
    '	+----[135]-CheckSampleQty   							                    | 檢測Sample Qty (不可SampleQty>600)
    '	+									                                        |
    '	+----[140]-CheckCompanyCode								                    | 檢測Company Code	
    '	+									                                        |
    '	+----[150]-CheckKeepCode								                    | 檢測Keep Code	
    '	+									                                        |
    '	+----[160]-MakeColorCode       						                        | 製作Color Code(對象:數字Color Code)
    '	+									                                        |
    '	+----[170]-CheckColorCode								                    | 檢測Color Code
    '	+									                                        |
    '	+----[180]-CheckItemCode								                    | 檢測Item Code (含PBR ITEM切替)
    '	+									                                        |
    '	+----[190]-CheckDuplicateData						                        | 檢測Duplicate EDI-Data(S5K00)
    '	+									                                        |
    '	+----[200]-CheckNikeVDP     						                        | 檢測NIKE VDP
    '	+									                                        |
    '	+----[210]-CheckPOStructure  						                        | 檢測客戶PO結構資料
    '	+									                                        |
    '	+----[220]-EDI2Waves        						                        | EDI Data To Waves System
    '	+									                                        |
    '	+----[230]-EDI2B2B         						                            | EDI Data To B2B Link File
    '	+									                                        |
    '	+----[240]-PriceList       						                            | 展開 Price List 資料
    '	+									                                        |
    '	+----[250]-UpdateOrderProgress					                            | 更新 Order Progress 資料
    '	+									                                        |
    '	+----[260]-UpdatePackingList   					                            | 更新 Packing List 資料
    '	+									                                        |
    '	+-----------------------------------------------------------------------------------------------------------------
    '	+									                                        |
    '	+----[500]-UpdateControlRecord      	                                    | 更新客戶控制檔-Action Flag
    '	+									                                        |
    '	+----[510]-CheckControlRecord									            | 檢查客戶控制檔-Action Flag
    '	+									                                        |
    '	+----[520]-DeleteHistory 									                | 刪除執行履歷
    '	+									                                        |
    '	+----[530]-DeleteActionHistory 									            | 刪除執行履歷(指定Action)
    '	+									                                        |
    '	+----[540]-DeleteS5K00Data 	    								            | 刪除S5K00資料
    '	+								                                            |
    '	+----[550]-DeleteS5L00Data 			    						            | 刪除S5L00資料
    '	+									                                        |
    '	+----[560]-DeleteSC760W1Data 									            | 刪除SC760W1資料
    '	+									                                        |
    '	+----[561]-DeletePriceListData 									            | 刪除PriceList資料
    '	+									                                        |
    '	+----[570]-GetPOSeqNo        									            | 取得PO SeqNo
    '	+									                                        |
    '	+----[580]-UpdateNextPONO								                    | 更新下一次可使用PONO
    '	+									                                        |
    '	+----									                                    |
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([100]-MakePONO)
    '**     製作採購號碼 
    '***********************************************************************************************
    'MakePONO-Start
    <WebMethod()> _
    Public Function MakePONO(ByVal pLogID As String, _
                             ByVal pBuyer As String, _
                             ByVal pUserID As String, _
                             ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim xPO As String = "X"
        Dim xPONO As String = ""
        Dim xPOSeqString As String = ""
        Dim xPOSeqNo, xSeqNo, i, j As Integer
        '------------------------------------------
        ' 2017/5/23 EOES-多客戶多BUYER
        Dim eBuyer As String = ""
        Dim eGRBuyer As String = ""
        Dim k As Integer
        '------------------------------------------
        ' 2020/6/12 COVID-19 RECOVERY CAMPAIGN
        'Dim str As String = ""
        '
        Try
            '[EOES使用]
            If InStr(pBuyer, "EOES-") > 0 Then
                ' **EOES-START
                'JJJ-對應
                'sql = "Select RFCC5K + '-' + RCMC5K as CustBuyer From S5K00 "
                sql = "Select Case SUBSTRING(RFCC5K, 1, 4) WHEN 'ZJJJ' THEN RFCC5K + '-' + '000999' ELSE RFCC5K + '-' + RCMC5K END  As CustBuyer "
                sql &= "From S5K00 "
                '<<<
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "Group By RFCC5K, RCMC5K "
                sql &= "Order By RFCC5K, + RCMC5K "
                Dim dt_EOES_Buyer As DataTable = oDataBase.GetDataTable(sql)
                For k = 0 To dt_EOES_Buyer.Rows.Count - 1
                    '
                    ' BUYER處理
                    eBuyer = dt_EOES_Buyer.Rows(k).Item("CustBuyer")
                    '
                    ' PONO=客戶代表Code + 年月日 + PO流水號(xPOSeqNo) 
                    ' SeqNo=細項流水號(xSeqNo) 
                    ' 取得客戶代表Code
                    sql = "Select BuyerGroup, POCode From M_ControlRecord "
                    sql = sql & " Where Buyer = '" & eBuyer & "'"
                    Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
                    If dt_ControlRecord.Rows.Count > 0 Then
                        eGRBuyer = dt_ControlRecord.Rows(0).Item("BuyerGroup")
                        xPO = dt_ControlRecord.Rows(0).Item("POCode")
                    End If
                    ' 年月日 
                    Select Case Month(Now)
                        Case 10
                            If Day(Now) < 10 Then
                                xPONO = xPO + CStr(Year(Now) - 2000) + "A" + "0" + CStr(Day(Now))
                            Else
                                xPONO = xPO + CStr(Year(Now) - 2000) + "A" + CStr(Day(Now))
                            End If
                        Case 11
                            If Day(Now) < 10 Then
                                xPONO = xPO + CStr(Year(Now) - 2000) + "B" + "0" + CStr(Day(Now))
                            Else
                                xPONO = xPO + CStr(Year(Now) - 2000) + "B" + CStr(Day(Now))
                            End If
                        Case 12
                            If Day(Now) < 10 Then
                                xPONO = xPO + CStr(Year(Now) - 2000) + "C" + "0" + CStr(Day(Now))
                            Else
                                xPONO = xPO + CStr(Year(Now) - 2000) + "C" + CStr(Day(Now))
                            End If
                        Case Else
                            If Day(Now) < 10 Then
                                xPONO = xPO + CStr(Year(Now) - 2000) + CStr(Month(Now)) + "0" + CStr(Day(Now))
                            Else
                                xPONO = xPO + CStr(Year(Now) - 2000) + CStr(Month(Now)) + CStr(Day(Now))
                            End If
                    End Select
                    '
                    ' 以PUCN5K+ITMC5K編碼 --> GRBuyer(4)=P
                    If oDB.GetFunctionCode(eGRBuyer, 4) = "P" Then
                        xPOSeqNo = GetPOSeqNo(eBuyer, xPONO)        ' PO流水號
                        sql = "SELECT PUCN5K+ITMC5K AS POKEY FROM S5K00 "
                        sql &= "Where Buyer = '" & pBuyer & "' "

                        'JJJ-對應
                        'sql &= "And RFCC5K + '-' + RCMC5K = '" & eBuyer & "' "
                        If InStr(eBuyer, "ZJJJ") > 0 Then
                            sql &= "And SUBSTRING(RFCC5K, 1, 4) = '" & "ZJJJ" & "' "
                        Else
                            sql &= "And RFCC5K + '-' + RCMC5K = '" & eBuyer & "' "
                        End If
                        '<<<

                        sql &= "Group by PUCN5K+ITMC5K "
                        sql &= "Order by PUCN5K+ITMC5K "
                        Dim dt_S5K00A As DataTable = oDataBase.GetDataTable(sql)
                        For i = 0 To dt_S5K00A.Rows.Count - 1
                            xSeqNo = 1              ' 細項流水號
                            sql = "SELECT *, PUCN5K+ITMC5K AS POKEY FROM S5K00 "
                            sql &= "Where Buyer = '" & pBuyer & "' "

                            'JJJ-對應
                            'sql &= "And RFCC5K + '-' + RCMC5K = '" & eBuyer & "' "
                            If InStr(eBuyer, "ZJJJ") > 0 Then
                                sql &= "And SUBSTRING(RFCC5K, 1, 4) = '" & "ZJJJ" & "' "
                            Else
                                sql &= "And RFCC5K + '-' + RCMC5K = '" & eBuyer & "' "
                            End If
                            '<<<

                            sql &= "  And PUCN5K+ITMC5K = '" & dt_S5K00A.Rows(i).Item("POKEY") & "' "
                            sql &= "Order by Unique_ID "
                            Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                            For j = 0 To dt_S5K00.Rows.Count - 1
                                ' 製作PO流水號(不足3位補0)
                                xPOSeqString = CStr(xPOSeqNo)
                                If Len(xPOSeqString) < 2 Then
                                    xPOSeqString = "00" + CStr(xPOSeqNo)
                                Else
                                    If Len(xPOSeqString) < 3 Then
                                        xPOSeqString = "0" + CStr(xPOSeqNo)
                                    End If
                                End If
                                '
                                ' S5K00-更新PONO
                                '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                                If oDB.GetFunctionCode(eGRBuyer, 3) = "P" Then
                                    sql = "Update S5K00 Set "
                                    sql = sql + " PODN5K = '" & xPONO + xPOSeqString & "', "
                                    sql = sql + " POSN5K = '" & CStr(xSeqNo) & "' "
                                Else
                                    sql = "Update S5K00 Set "
                                    sql = sql + " PODN5K = '" & xPONO + xPOSeqString & "'  "
                                End If
                                sql = sql + " Where Unique_ID = '" & dt_S5K00.Rows(j).Item("Unique_ID").ToString & "'"
                                oDataBase.ExecuteNonQuery(sql)
                                '
                                ' S5L00-更新PONO
                                sql = "Update S5L00 Set "
                                sql = sql + " PODN5L = '" & xPONO + xPOSeqString & "' "
                                sql = sql + "Where Buyer = '" & pBuyer & "' "
                                sql = sql + "And ORFN = '" & dt_S5K00.Rows(j).Item("POKEY") & "' "
                                oDataBase.ExecuteNonQuery(sql)
                                '
                                ' SC760W1-更新PONO
                                If oDB.GetFunctionCode(eGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                                    '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                                    If oDB.GetFunctionCode(eGRBuyer, 3) = "P" Then
                                        sql = "Update SC760W1 Set "
                                        sql = sql + " ORDNW1 = '" & xPONO + xPOSeqString & "', "
                                        sql = sql + " OSBNW1 = '" & CStr(xSeqNo) & "' "
                                    Else
                                        sql = "Update SC760W1 Set "
                                        sql = sql + " ORDNW1 = '" & xPONO + xPOSeqString & "'  "
                                    End If
                                    sql = sql + "Where Buyer = '" & pBuyer & "' "
                                    sql = sql + "  And ORFN = '" & dt_S5K00.Rows(j).Item("POKEY") & "' "
                                    sql = sql + "  And ID = '" & dt_S5K00.Rows(j).Item("ID").ToString & "' "
                                    oDataBase.ExecuteNonQuery(sql)
                                End If
                                '
                                xSeqNo = xSeqNo + 1
                            Next

                            xPOSeqNo = xPOSeqNo + 1
                        Next
                    End If
                    '
                    ' 以ORFN5K編碼 --> GRBuyer(4)=R
                    If oDB.GetFunctionCode(eGRBuyer, 4) = "R" Then
                        xPOSeqNo = GetPOSeqNo(eBuyer, xPONO)        ' PO流水號
                        sql = "SELECT ORFN5K+ITMC5K AS POKEY FROM S5K00 "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "And RFCC5K + '-' + RCMC5K = '" & eBuyer & "' "
                        sql &= "Group by ORFN5K+ITMC5K "
                        sql &= "Order by ORFN5K+ITMC5K "
                        Dim dt_S5K00A As DataTable = oDataBase.GetDataTable(sql)
                        For i = 0 To dt_S5K00A.Rows.Count - 1
                            xSeqNo = 1              ' 細項流水號
                            sql = "SELECT *,ORFN5K+ITMC5K AS POKEY FROM S5K00 "
                            sql &= "Where Buyer = '" & pBuyer & "' "
                            sql &= "And RFCC5K + '-' + RCMC5K = '" & eBuyer & "' "
                            sql &= "  And ORFN5K+ITMC5K = '" & dt_S5K00A.Rows(i).Item("POKEY") & "' "
                            sql &= "Order by Unique_ID "
                            Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                            For j = 0 To dt_S5K00.Rows.Count - 1
                                ' 製作PO流水號(不足3位補0)
                                xPOSeqString = CStr(xPOSeqNo)
                                If Len(xPOSeqString) < 2 Then
                                    xPOSeqString = "00" + CStr(xPOSeqNo)
                                Else
                                    If Len(xPOSeqString) < 3 Then
                                        xPOSeqString = "0" + CStr(xPOSeqNo)
                                    End If
                                End If
                                '
                                ' S5K00-更新PONO
                                '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                                If oDB.GetFunctionCode(eGRBuyer, 3) = "P" Then
                                    sql = "Update S5K00 Set "
                                    sql = sql + " PODN5K = '" & xPONO + xPOSeqString & "', "
                                    sql = sql + " POSN5K = '" & CStr(xSeqNo) & "' "
                                Else
                                    sql = "Update S5K00 Set "
                                    sql = sql + " PODN5K = '" & xPONO + xPOSeqString & "'  "
                                End If
                                sql = sql + " Where Unique_ID = '" & dt_S5K00.Rows(j).Item("Unique_ID").ToString & "'"
                                oDataBase.ExecuteNonQuery(sql)
                                '
                                ' S5L00-更新PONO
                                sql = "Update S5L00 Set "
                                sql = sql + " PODN5L = '" & xPONO + xPOSeqString & "' "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql = sql + "  And ORFN = '" & dt_S5K00.Rows(j).Item("POKEY") & "' "
                                oDataBase.ExecuteNonQuery(sql)
                                '
                                ' SC760W1-更新PONO
                                If oDB.GetFunctionCode(eGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                                    '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                                    If oDB.GetFunctionCode(eGRBuyer, 3) = "P" Then
                                        sql = "Update SC760W1 Set "
                                        sql = sql + " ORDNW1 = '" & xPONO + xPOSeqString & "', "
                                        sql = sql + " OSBNW1 = '" & CStr(xSeqNo) & "' "
                                    Else
                                        sql = "Update SC760W1 Set "
                                        sql = sql + " ORDNW1 = '" & xPONO + xPOSeqString & "'  "
                                    End If
                                    sql &= "Where Buyer = '" & pBuyer & "' "
                                    sql = sql + "  And ORFN = '" & dt_S5K00.Rows(j).Item("POKEY") & "' "
                                    sql = sql + "  And ID = '" & dt_S5K00.Rows(j).Item("ID").ToString & "' "
                                    oDataBase.ExecuteNonQuery(sql)
                                End If
                                '
                                xSeqNo = xSeqNo + 1
                            Next
                            xPOSeqNo = xPOSeqNo + 1
                        Next
                    End If
                    '
                    ' 以PUCN5K編碼 --> GRBuyer(4)=X
                    If oDB.GetFunctionCode(eGRBuyer, 4) = "X" Then
                        xPOSeqNo = GetPOSeqNo(eBuyer, xPONO)        ' PO流水號
                        '
                        sql = "SELECT PUCN5K AS POKEY FROM S5K00 "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "And RFCC5K + '-' + RCMC5K = '" & eBuyer & "' "
                        sql &= "Group by PUCN5K "
                        sql &= "Order by PUCN5K "
                        Dim dt_S5K00A As DataTable = oDataBase.GetDataTable(sql)
                        For i = 0 To dt_S5K00A.Rows.Count - 1
                            xSeqNo = 1              ' 細項流水號
                            sql = "SELECT *, PUCN5K AS POKEY FROM S5K00 "
                            sql &= "Where Buyer = '" & pBuyer & "' "
                            sql &= "And RFCC5K + '-' + RCMC5K = '" & eBuyer & "' "
                            sql &= "  And PUCN5K = '" & dt_S5K00A.Rows(i).Item("POKEY") & "' "
                            sql &= "Order by Unique_ID "
                            Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                            For j = 0 To dt_S5K00.Rows.Count - 1
                                '
                                ' 製作PO流水號(不足3位補0)
                                xPOSeqString = CStr(xPOSeqNo)
                                If Len(xPOSeqString) < 2 Then
                                    xPOSeqString = "00" + CStr(xPOSeqNo)
                                Else
                                    If Len(xPOSeqString) < 3 Then
                                        xPOSeqString = "0" + CStr(xPOSeqNo)
                                    End If
                                End If
                                '
                                ' S5K00-更新PONO
                                '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                                If oDB.GetFunctionCode(eGRBuyer, 3) = "P" Then
                                    sql = "Update S5K00 Set "
                                    sql = sql + " PODN5K = '" & xPONO + xPOSeqString & "', "
                                    sql = sql + " POSN5K = '" & CStr(xSeqNo) & "' "
                                Else
                                    sql = "Update S5K00 Set "
                                    sql = sql + " PODN5K = '" & xPONO + xPOSeqString & "'  "
                                End If
                                sql = sql + " Where Unique_ID = '" & dt_S5K00.Rows(j).Item("Unique_ID").ToString & "'"
                                oDataBase.ExecuteNonQuery(sql)
                                '
                                ' S5L00-更新PONO
                                sql = "Update S5L00 Set "
                                sql = sql + " PODN5L = '" & xPONO + xPOSeqString & "' "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql = sql + "  And ORFN = '" & dt_S5K00.Rows(j).Item("POKEY") & "' "
                                oDataBase.ExecuteNonQuery(sql)
                                '
                                ' SC760W1-更新PONO
                                If oDB.GetFunctionCode(eGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                                    '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                                    If oDB.GetFunctionCode(eGRBuyer, 3) = "P" Then
                                        sql = "Update SC760W1 Set "
                                        sql = sql + " ORDNW1 = '" & xPONO + xPOSeqString & "', "
                                        sql = sql + " OSBNW1 = '" & CStr(xSeqNo) & "' "
                                    Else
                                        sql = "Update SC760W1 Set "
                                        sql = sql + " ORDNW1 = '" & xPONO + xPOSeqString & "'  "
                                    End If
                                    sql &= "Where Buyer = '" & pBuyer & "' "
                                    sql = sql + "  And ORFN = '" & dt_S5K00.Rows(j).Item("POKEY") & "' "
                                    sql = sql + "  And ID = '" & dt_S5K00.Rows(j).Item("ID").ToString & "' "
                                    oDataBase.ExecuteNonQuery(sql)
                                End If
                                '
                                xSeqNo = xSeqNo + 1
                            Next

                            xPOSeqNo = xPOSeqNo + 1
                        Next
                    End If
                    ' ADD-START 2017/2/8
                    '
                    ' 以PUCN5K編碼 --> GRBuyer(4)=Y
                    If oDB.GetFunctionCode(eGRBuyer, 4) = "Y" Then
                        xPOSeqNo = GetPOSeqNo(eBuyer, xPONO)        ' PO流水號
                        sql = "SELECT ORFN5K AS POKEY FROM S5K00 "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "And RFCC5K + '-' + RCMC5K = '" & eBuyer & "' "
                        sql &= "Group by ORFN5K "
                        sql &= "Order by ORFN5K "
                        Dim dt_S5K00A As DataTable = oDataBase.GetDataTable(sql)
                        For i = 0 To dt_S5K00A.Rows.Count - 1
                            xSeqNo = 1              ' 細項流水號
                            sql = "SELECT *, ORFN5K AS POKEY FROM S5K00 "
                            sql &= "Where Buyer = '" & pBuyer & "' "
                            sql &= "And RFCC5K + '-' + RCMC5K = '" & eBuyer & "' "
                            sql &= "  And ORFN5K = '" & dt_S5K00A.Rows(i).Item("POKEY") & "' "
                            sql &= "Order by Unique_ID "
                            Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                            For j = 0 To dt_S5K00.Rows.Count - 1
                                ' 製作PO流水號(不足3位補0)
                                xPOSeqString = CStr(xPOSeqNo)
                                If Len(xPOSeqString) < 2 Then
                                    xPOSeqString = "00" + CStr(xPOSeqNo)
                                Else
                                    If Len(xPOSeqString) < 3 Then
                                        xPOSeqString = "0" + CStr(xPOSeqNo)
                                    End If
                                End If
                                '
                                ' S5K00-更新PONO
                                '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                                If oDB.GetFunctionCode(eGRBuyer, 3) = "P" Then
                                    sql = "Update S5K00 Set "
                                    sql = sql + " PODN5K = '" & xPONO + xPOSeqString & "', "
                                    sql = sql + " POSN5K = '" & CStr(xSeqNo) & "' "
                                Else
                                    sql = "Update S5K00 Set "
                                    sql = sql + " PODN5K = '" & xPONO + xPOSeqString & "'  "
                                End If
                                sql = sql + " Where Unique_ID = '" & dt_S5K00.Rows(j).Item("Unique_ID").ToString & "'"
                                oDataBase.ExecuteNonQuery(sql)
                                '
                                ' S5L00-更新PONO
                                sql = "Update S5L00 Set "
                                sql = sql + " PODN5L = '" & xPONO + xPOSeqString & "' "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql = sql + "  And ORFN = '" & dt_S5K00.Rows(j).Item("POKEY") & "' "
                                oDataBase.ExecuteNonQuery(sql)
                                '
                                ' SC760W1-更新PONO
                                If oDB.GetFunctionCode(eGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                                    '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                                    If oDB.GetFunctionCode(eGRBuyer, 3) = "P" Then
                                        sql = "Update SC760W1 Set "
                                        sql = sql + " ORDNW1 = '" & xPONO + xPOSeqString & "', "
                                        sql = sql + " OSBNW1 = '" & CStr(xSeqNo) & "' "
                                    Else
                                        sql = "Update SC760W1 Set "
                                        sql = sql + " ORDNW1 = '" & xPONO + xPOSeqString & "'  "
                                    End If
                                    sql &= "Where Buyer = '" & pBuyer & "' "
                                    sql = sql + "  And ORFN = '" & dt_S5K00.Rows(j).Item("POKEY") & "' "
                                    sql = sql + "  And ID = '" & dt_S5K00.Rows(j).Item("ID").ToString & "' "
                                    oDataBase.ExecuteNonQuery(sql)
                                End If
                                '
                                xSeqNo = xSeqNo + 1
                            Next

                            xPOSeqNo = xPOSeqNo + 1
                        Next
                    End If
                    ' ADD-END
                    '
                    ' 製作PO流水號(不足3位補0)
                    xPOSeqString = CStr(xPOSeqNo)
                    If Len(xPOSeqString) < 2 Then
                        xPOSeqString = "00" + CStr(xPOSeqNo)
                    Else
                        If Len(xPOSeqString) < 3 Then
                            xPOSeqString = "0" + CStr(xPOSeqNo)
                        End If
                    End If
                    ' 更新下一次可使用PONO
                    UpdateNextPONO(eBuyer, xPONO + xPOSeqString)
                Next
                ' **EOES-END
            Else
                ' **原程式
                ' PONO=客戶代表Code + 年月日 + PO流水號(xPOSeqNo) 
                ' SeqNo=細項流水號(xSeqNo) 
                ' 取得客戶代表Code
                sql = "SELECT POCode FROM M_ControlRecord "
                sql &= "Where Buyer = '" & pBuyer & "' "
                Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
                If dt_ControlRecord.Rows.Count > 0 Then xPO = dt_ControlRecord.Rows(0).Item("POCode")
                ' 年月日 
                Select Case Month(Now)
                    Case 10
                        If Day(Now) < 10 Then
                            xPONO = xPO + CStr(Year(Now) - 2000) + "A" + "0" + CStr(Day(Now))
                        Else
                            xPONO = xPO + CStr(Year(Now) - 2000) + "A" + CStr(Day(Now))
                        End If
                    Case 11
                        If Day(Now) < 10 Then
                            xPONO = xPO + CStr(Year(Now) - 2000) + "B" + "0" + CStr(Day(Now))
                        Else
                            xPONO = xPO + CStr(Year(Now) - 2000) + "B" + CStr(Day(Now))
                        End If
                    Case 12
                        If Day(Now) < 10 Then
                            xPONO = xPO + CStr(Year(Now) - 2000) + "C" + "0" + CStr(Day(Now))
                        Else
                            xPONO = xPO + CStr(Year(Now) - 2000) + "C" + CStr(Day(Now))
                        End If
                    Case Else
                        If Day(Now) < 10 Then
                            xPONO = xPO + CStr(Year(Now) - 2000) + CStr(Month(Now)) + "0" + CStr(Day(Now))
                        Else
                            xPONO = xPO + CStr(Year(Now) - 2000) + CStr(Month(Now)) + CStr(Day(Now))
                        End If
                End Select
                '
                ' 以PUCN5K+ITMC5K編碼 --> GRBuyer(4)=P
                If oDB.GetFunctionCode(pGRBuyer, 4) = "P" Then
                    xPOSeqNo = GetPOSeqNo(pBuyer, xPONO)        ' PO流水號
                    sql = "SELECT PUCN5K+ITMC5K AS POKEY FROM S5K00 "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "Group by PUCN5K+ITMC5K "
                    sql &= "Order by PUCN5K+ITMC5K "
                    Dim dt_S5K00A As DataTable = oDataBase.GetDataTable(sql)
                    For i = 0 To dt_S5K00A.Rows.Count - 1
                        xSeqNo = 1              ' 細項流水號
                        sql = "SELECT *, PUCN5K+ITMC5K AS POKEY FROM S5K00 "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "  And PUCN5K+ITMC5K = '" & dt_S5K00A.Rows(i).Item("POKEY") & "' "
                        sql &= "Order by Unique_ID "
                        Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                        For j = 0 To dt_S5K00.Rows.Count - 1
                            ' 製作PO流水號(不足3位補0)
                            xPOSeqString = CStr(xPOSeqNo)
                            If Len(xPOSeqString) < 2 Then
                                xPOSeqString = "00" + CStr(xPOSeqNo)
                            Else
                                If Len(xPOSeqString) < 3 Then
                                    xPOSeqString = "0" + CStr(xPOSeqNo)
                                End If
                            End If
                            '
                            ' S5K00-更新PONO
                            '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                            If oDB.GetFunctionCode(pGRBuyer, 3) = "P" Then
                                sql = "Update S5K00 Set "
                                sql = sql + " PODN5K = '" & xPONO + xPOSeqString & "', "
                                sql = sql + " POSN5K = '" & CStr(xSeqNo) & "' "
                            Else
                                sql = "Update S5K00 Set "
                                sql = sql + " PODN5K = '" & xPONO + xPOSeqString & "'  "
                            End If
                            sql = sql + " Where Unique_ID = '" & dt_S5K00.Rows(j).Item("Unique_ID").ToString & "'"
                            oDataBase.ExecuteNonQuery(sql)
                            '
                            ' S5L00-更新PONO
                            sql = "Update S5L00 Set "
                            sql = sql + " PODN5L = '" & xPONO + xPOSeqString & "' "
                            sql = sql + "Where Buyer = '" & pBuyer & "' "
                            sql = sql + "  And ORFN = '" & dt_S5K00.Rows(j).Item("POKEY") & "' "
                            oDataBase.ExecuteNonQuery(sql)
                            '
                            ' SC760W1-更新PONO
                            If oDB.GetFunctionCode(pGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                                '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                                If oDB.GetFunctionCode(pGRBuyer, 3) = "P" Then
                                    sql = "Update SC760W1 Set "
                                    sql = sql + " ORDNW1 = '" & xPONO + xPOSeqString & "', "
                                    sql = sql + " OSBNW1 = '" & CStr(xSeqNo) & "' "
                                Else
                                    sql = "Update SC760W1 Set "
                                    sql = sql + " ORDNW1 = '" & xPONO + xPOSeqString & "'  "
                                End If
                                sql = sql + "Where Buyer = '" & pBuyer & "' "
                                sql = sql + "  And ORFN = '" & dt_S5K00.Rows(j).Item("POKEY") & "' "
                                sql = sql + "  And ID = '" & dt_S5K00.Rows(j).Item("ID").ToString & "' "
                                oDataBase.ExecuteNonQuery(sql)
                            End If
                            '
                            xSeqNo = xSeqNo + 1
                        Next

                        xPOSeqNo = xPOSeqNo + 1
                    Next
                End If
                '
                ' 以ORFN5K編碼 --> GRBuyer(4)=R
                If oDB.GetFunctionCode(pGRBuyer, 4) = "R" Then
                    xPOSeqNo = GetPOSeqNo(pBuyer, xPONO)        ' PO流水號
                    sql = "SELECT ORFN5K+ITMC5K AS POKEY FROM S5K00 "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "Group by ORFN5K+ITMC5K "
                    sql &= "Order by ORFN5K+ITMC5K "
                    Dim dt_S5K00A As DataTable = oDataBase.GetDataTable(sql)
                    For i = 0 To dt_S5K00A.Rows.Count - 1
                        xSeqNo = 1              ' 細項流水號
                        sql = "SELECT *,ORFN5K+ITMC5K AS POKEY FROM S5K00 "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "  And ORFN5K+ITMC5K = '" & dt_S5K00A.Rows(i).Item("POKEY") & "' "
                        sql &= "Order by Unique_ID "
                        Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                        For j = 0 To dt_S5K00.Rows.Count - 1
                            ' 製作PO流水號(不足3位補0)
                            xPOSeqString = CStr(xPOSeqNo)
                            If Len(xPOSeqString) < 2 Then
                                xPOSeqString = "00" + CStr(xPOSeqNo)
                            Else
                                If Len(xPOSeqString) < 3 Then
                                    xPOSeqString = "0" + CStr(xPOSeqNo)
                                End If
                            End If
                            '
                            ' S5K00-更新PONO
                            '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                            If oDB.GetFunctionCode(pGRBuyer, 3) = "P" Then
                                sql = "Update S5K00 Set "
                                sql = sql + " PODN5K = '" & xPONO + xPOSeqString & "', "
                                sql = sql + " POSN5K = '" & CStr(xSeqNo) & "' "
                            Else
                                sql = "Update S5K00 Set "
                                sql = sql + " PODN5K = '" & xPONO + xPOSeqString & "'  "
                            End If
                            sql = sql + " Where Unique_ID = '" & dt_S5K00.Rows(j).Item("Unique_ID").ToString & "'"
                            oDataBase.ExecuteNonQuery(sql)
                            '
                            ' S5L00-更新PONO
                            sql = "Update S5L00 Set "
                            sql = sql + " PODN5L = '" & xPONO + xPOSeqString & "' "
                            sql = sql + "Where Buyer = '" & pBuyer & "' "
                            sql = sql + "  And ORFN = '" & dt_S5K00.Rows(j).Item("POKEY") & "' "
                            oDataBase.ExecuteNonQuery(sql)
                            '
                            ' SC760W1-更新PONO
                            If oDB.GetFunctionCode(pGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                                '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                                If oDB.GetFunctionCode(pGRBuyer, 3) = "P" Then
                                    sql = "Update SC760W1 Set "
                                    sql = sql + " ORDNW1 = '" & xPONO + xPOSeqString & "', "
                                    sql = sql + " OSBNW1 = '" & CStr(xSeqNo) & "' "
                                Else
                                    sql = "Update SC760W1 Set "
                                    sql = sql + " ORDNW1 = '" & xPONO + xPOSeqString & "'  "
                                End If
                                sql = sql + "Where Buyer = '" & pBuyer & "' "
                                sql = sql + "  And ORFN = '" & dt_S5K00.Rows(j).Item("POKEY") & "' "
                                sql = sql + "  And ID = '" & dt_S5K00.Rows(j).Item("ID").ToString & "' "
                                oDataBase.ExecuteNonQuery(sql)
                            End If
                            '
                            xSeqNo = xSeqNo + 1
                        Next
                        xPOSeqNo = xPOSeqNo + 1
                    Next
                End If
                '
                ' 以PUCN5K編碼 --> GRBuyer(4)=X
                If oDB.GetFunctionCode(pGRBuyer, 4) = "X" Then
                    xPOSeqNo = GetPOSeqNo(pBuyer, xPONO)        ' PO流水號
                    sql = "SELECT PUCN5K AS POKEY FROM S5K00 "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "Group by PUCN5K "
                    sql &= "Order by PUCN5K "
                    Dim dt_S5K00A As DataTable = oDataBase.GetDataTable(sql)
                    For i = 0 To dt_S5K00A.Rows.Count - 1
                        xSeqNo = 1              ' 細項流水號
                        sql = "SELECT *, PUCN5K AS POKEY FROM S5K00 "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "  And PUCN5K = '" & dt_S5K00A.Rows(i).Item("POKEY") & "' "
                        sql &= "Order by Unique_ID "
                        Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                        For j = 0 To dt_S5K00.Rows.Count - 1
                            ' 製作PO流水號(不足3位補0)
                            xPOSeqString = CStr(xPOSeqNo)
                            If Len(xPOSeqString) < 2 Then
                                xPOSeqString = "00" + CStr(xPOSeqNo)
                            Else
                                If Len(xPOSeqString) < 3 Then
                                    xPOSeqString = "0" + CStr(xPOSeqNo)
                                End If
                            End If
                            '
                            ' S5K00-更新PONO
                            '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                            If oDB.GetFunctionCode(pGRBuyer, 3) = "P" Then
                                sql = "Update S5K00 Set "
                                sql = sql + " PODN5K = '" & xPONO + xPOSeqString & "', "
                                sql = sql + " POSN5K = '" & CStr(xSeqNo) & "' "
                            Else
                                sql = "Update S5K00 Set "
                                sql = sql + " PODN5K = '" & xPONO + xPOSeqString & "'  "
                            End If
                            sql = sql + " Where Unique_ID = '" & dt_S5K00.Rows(j).Item("Unique_ID").ToString & "'"
                            oDataBase.ExecuteNonQuery(sql)
                            '
                            ' S5L00-更新PONO
                            sql = "Update S5L00 Set "
                            sql = sql + " PODN5L = '" & xPONO + xPOSeqString & "' "
                            sql = sql + "Where Buyer = '" & pBuyer & "' "
                            sql = sql + "  And ORFN = '" & dt_S5K00.Rows(j).Item("POKEY") & "' "
                            oDataBase.ExecuteNonQuery(sql)
                            '
                            ' SC760W1-更新PONO
                            If oDB.GetFunctionCode(pGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                                '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                                If oDB.GetFunctionCode(pGRBuyer, 3) = "P" Then
                                    sql = "Update SC760W1 Set "
                                    sql = sql + " ORDNW1 = '" & xPONO + xPOSeqString & "', "
                                    sql = sql + " OSBNW1 = '" & CStr(xSeqNo) & "' "
                                Else
                                    sql = "Update SC760W1 Set "
                                    sql = sql + " ORDNW1 = '" & xPONO + xPOSeqString & "'  "
                                End If
                                sql = sql + "Where Buyer = '" & pBuyer & "' "
                                sql = sql + "  And ORFN = '" & dt_S5K00.Rows(j).Item("POKEY") & "' "
                                sql = sql + "  And ID = '" & dt_S5K00.Rows(j).Item("ID").ToString & "' "
                                oDataBase.ExecuteNonQuery(sql)
                            End If
                            '
                            xSeqNo = xSeqNo + 1
                        Next

                        xPOSeqNo = xPOSeqNo + 1
                    Next
                End If
                ' ADD-START 2017/2/8
                '
                ' 以PUCN5K編碼 --> GRBuyer(4)=Y
                If oDB.GetFunctionCode(pGRBuyer, 4) = "Y" Then
                    xPOSeqNo = GetPOSeqNo(pBuyer, xPONO)        ' PO流水號
                    sql = "SELECT ORFN5K AS POKEY FROM S5K00 "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "Group by ORFN5K "
                    sql &= "Order by ORFN5K "
                    Dim dt_S5K00A As DataTable = oDataBase.GetDataTable(sql)
                    For i = 0 To dt_S5K00A.Rows.Count - 1
                        xSeqNo = 1              ' 細項流水號
                        sql = "SELECT *, ORFN5K AS POKEY FROM S5K00 "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "  And ORFN5K = '" & dt_S5K00A.Rows(i).Item("POKEY") & "' "
                        sql &= "Order by Unique_ID "
                        Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                        For j = 0 To dt_S5K00.Rows.Count - 1
                            ' 製作PO流水號(不足3位補0)
                            xPOSeqString = CStr(xPOSeqNo)
                            If Len(xPOSeqString) < 2 Then
                                xPOSeqString = "00" + CStr(xPOSeqNo)
                            Else
                                If Len(xPOSeqString) < 3 Then
                                    xPOSeqString = "0" + CStr(xPOSeqNo)
                                End If
                            End If
                            '
                            ' S5K00-更新PONO
                            '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                            If oDB.GetFunctionCode(pGRBuyer, 3) = "P" Then
                                sql = "Update S5K00 Set "
                                sql = sql + " PODN5K = '" & xPONO + xPOSeqString & "', "
                                sql = sql + " POSN5K = '" & CStr(xSeqNo) & "' "
                            Else
                                sql = "Update S5K00 Set "
                                sql = sql + " PODN5K = '" & xPONO + xPOSeqString & "'  "
                            End If
                            sql = sql + " Where Unique_ID = '" & dt_S5K00.Rows(j).Item("Unique_ID").ToString & "'"
                            oDataBase.ExecuteNonQuery(sql)
                            '
                            ' S5L00-更新PONO
                            sql = "Update S5L00 Set "
                            sql = sql + " PODN5L = '" & xPONO + xPOSeqString & "' "
                            sql = sql + "Where Buyer = '" & pBuyer & "' "
                            sql = sql + "  And ORFN = '" & dt_S5K00.Rows(j).Item("POKEY") & "' "
                            oDataBase.ExecuteNonQuery(sql)
                            '
                            ' SC760W1-更新PONO
                            If oDB.GetFunctionCode(pGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                                '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                                If oDB.GetFunctionCode(pGRBuyer, 3) = "P" Then
                                    sql = "Update SC760W1 Set "
                                    sql = sql + " ORDNW1 = '" & xPONO + xPOSeqString & "', "
                                    sql = sql + " OSBNW1 = '" & CStr(xSeqNo) & "' "
                                Else
                                    sql = "Update SC760W1 Set "
                                    sql = sql + " ORDNW1 = '" & xPONO + xPOSeqString & "'  "
                                End If
                                sql = sql + "Where Buyer = '" & pBuyer & "' "
                                sql = sql + "  And ORFN = '" & dt_S5K00.Rows(j).Item("POKEY") & "' "
                                sql = sql + "  And ID = '" & dt_S5K00.Rows(j).Item("ID").ToString & "' "
                                oDataBase.ExecuteNonQuery(sql)
                            End If
                            '
                            xSeqNo = xSeqNo + 1
                        Next

                        xPOSeqNo = xPOSeqNo + 1
                    Next
                End If
                ' ADD-END
                '
                ' 製作PO流水號(不足3位補0)
                xPOSeqString = CStr(xPOSeqNo)
                If Len(xPOSeqString) < 2 Then
                    xPOSeqString = "00" + CStr(xPOSeqNo)
                Else
                    If Len(xPOSeqString) < 3 Then
                        xPOSeqString = "0" + CStr(xPOSeqNo)
                    End If
                End If
                ' 更新下一次可使用PONO
                UpdateNextPONO(pBuyer, xPONO + xPOSeqString)
                ' **原程式-END
            End If
            '
            '[COVID-19 RECOVERY CAMPAIGN]
            '2020/6/16 START
            'sql = "Select PODN5L, PO1X5L + PO2X5L + PO3X5L + PO4X5L AS COMMENT "
            'sql &= "From S5L00 "
            'sql &= "Where Buyer = '" & pBuyer & "' "
            'sql &= "And RFCC5L IN (SELECT CODE FROM M_CUSTCOVID19) "
            'sql &= "Order By PODN5L, PO1X5L, PO2X5L, PO3X5L, PO4X5L "
            'Dim dt_COVID19 As DataTable = oDataBase.GetDataTable(sql)
            'For k = 0 To dt_COVID19.Rows.Count - 1
            '    str = Replace(dt_COVID19.Rows(k).Item("COMMENT"), " ", "")
            '    If InStr(str, "COVID-19RECOVERYCAMPAIGN") <= 0 Then
            '        RtnCode = 9
            '        oDB.AccessLog(pLogID, pBuyer, "MakePONO", NowDateTime, 1, "COVID-19 COMMENT", "Code", "=", CStr(RtnCode), "", pUserID, "")
            '    End If
            'Next
            'END
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "MakePONO", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'MakePONO-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([110]-CheckPONO)
    '**     檢測採購號碼 
    '***********************************************************************************************
    'CheckPONO-Start
    <WebMethod()> _
    Public Function CheckPONO(ByVal pLogID As String, _
                              ByVal pBuyer As String, _
                              ByVal pUserID As String, _
                              ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim i, Row As Integer
        Dim err As Boolean = False
        Dim msg As String = ""
        ' ActionLog-Key欄
        Dim eCORN5K As String = ""
        Dim eGRPC5K As Integer = 0
        '------------------------------------------
        ' 2017/5/23 EOES-多客戶多BUYER
        Dim eBuyer As String = ""
        Dim eGRBuyer As String = ""
        '
        Try
            '[EOES使用]
            If InStr(pBuyer, "EOES-") > 0 Then
                ' **EOES-START
                Row = 1
                'JJJ-對應
                'sql = "SELECT * FROM S5K00 "
                sql = "Select *, Case SUBSTRING(RFCC5K, 1, 4) WHEN 'ZJJJ' THEN RFCC5K + '-' + '000999' ELSE RFCC5K + '-' + RCMC5K END  As CustBuyer "
                sql &= "From S5K00 "
                '<<<
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "Order by Unique_ID "
                Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                For i = 0 To dt_S5K00.Rows.Count - 1
                    ' ActionLog-Key欄
                    eCORN5K = dt_S5K00.Rows(i).Item("CORN5K").ToString
                    eGRPC5K = dt_S5K00.Rows(i).Item("GRPC5K")
                    'JJJ-對應
                    'eBuyer = dt_S5K00.Rows(i).Item("RFCC5K").ToString + "-" + dt_S5K00.Rows(i).Item("RCMC5K").ToString
                    eBuyer = dt_S5K00.Rows(i).Item("CustBuyer").ToString
                    '<<<

                    ' 取得各客戶-BUYER BUYERGROUP
                    sql = "Select BuyerGroup From M_ControlRecord "
                    sql = sql & " Where Buyer = '" & eBuyer & "'"
                    Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
                    If dt_ControlRecord.Rows.Count > 0 Then eGRBuyer = dt_ControlRecord.Rows(0).Item("BuyerGroup")
                    ' [Reference檢測]
                    If Not err Then
                        ' [R]→ORFN5K+ITMC5K / [Y]→ORFN5K / [P]→ PUCN5K+ITMC5K / [X]→ PUCN5K
                        If oDB.GetFunctionCode(eGRBuyer, 4) = "P" Or oDB.GetFunctionCode(eGRBuyer, 4) = "X" Then
                            If dt_S5K00.Rows(i).Item("PUCN5K") = "" Then
                                err = True
                                msg = "Check-PONO:[PUCN5K = Blank !]"
                            End If
                        Else
                            If dt_S5K00.Rows(i).Item("ORFN5K") = "" Then
                                err = True
                                msg = "Check-PONO:[ORFN5K = Blank !]"
                            End If
                        End If
                    End If
                    ' [GRPC5K檢測]
                    If Not err Then
                        If dt_S5K00.Rows(i).Item("GRPC5K") = 0 Then
                            err = True
                            msg = "Check-PONO:[GRPC5K = 0 !]"
                        End If
                    End If
                    ' [PONO檢測]
                    If Not err Then
                        If dt_S5K00.Rows(i).Item("PODN5K") = "" Then
                            err = True
                            msg = "Check-PONO:[PODN5K = Blank !]"
                        End If
                    End If
                    ' [PONO-SEQ檢測]
                    If Not err Then
                        If dt_S5K00.Rows(i).Item("POSN5K") = 0 Then
                            err = True
                            msg = "Check-PONO:[POSN5K = 0 !]"
                        End If
                    End If
                    ' [檢測異常處理]
                    If err Then
                        RtnCode = 1
                        oDB.AccessLog(pLogID, pBuyer, "CheckPONO", NowDateTime, 1, oDB.GetAccessLogLine(pLogID, pBuyer, eCORN5K, eGRPC5K), msg, "=", CStr(RtnCode), "", pUserID, "")
                        Exit For
                    Else
                        Row = Row + 1
                    End If
                Next
                '
                ' S5L00-檢測PONO
                If Not err Then
                    Row = 1
                    sql = "SELECT * FROM S5L00 "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "Order by Unique_ID "
                    Dim dt_S5L00 As DataTable = oDataBase.GetDataTable(sql)
                    For i = 0 To dt_S5L00.Rows.Count - 1
                        ' [PONO檢測]
                        If Not err Then
                            If dt_S5L00.Rows(i).Item("PODN5L") = "" Then
                                err = True
                                msg = "Check-PONO:[PODN5L = Blank !]"
                            End If
                        End If
                        ' [檢測異常處理]
                        If err Then
                            RtnCode = 1
                            oDB.AccessLog(pLogID, pBuyer, "CheckPONO", NowDateTime, 1, CStr(Row), msg, "=", CStr(RtnCode), "", pUserID, "")
                            Exit For
                        Else
                            Row = Row + 1
                        End If
                    Next
                End If
                '
                ' SC760W1-檢測PONO
                If Not err Then
                    Row = 1
                    sql = "SELECT * FROM SC760W1 "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "Order by Unique_ID "
                    Dim dt_SC760W1 As DataTable = oDataBase.GetDataTable(sql)
                    For i = 0 To dt_SC760W1.Rows.Count - 1
                        ' 取得S5K00 客戶-BUYER
                        'JJJ-對應
                        'sql = "SELECT RFCC5K + '-' + RCMC5K as CustBuyer FROM S5K00 "
                        sql = "Select Case SUBSTRING(RFCC5K, 1, 4) WHEN 'ZJJJ' THEN RFCC5K + '-' + '000999' ELSE RFCC5K + '-' + RCMC5K END  As CustBuyer "
                        sql &= "From S5K00 "
                        '<<<
                        sql &= "Where ID = '" & dt_SC760W1.Rows(i).Item("ID").ToString & "' "
                        Dim dt_S5K00A As DataTable = oDataBase.GetDataTable(sql)
                        If dt_S5K00A.Rows.Count > 0 Then eBuyer = dt_S5K00A.Rows(0).Item("CustBuyer")
                        ' 取得 BUYERGROUP
                        sql = "Select BuyerGroup From M_ControlRecord "
                        sql = sql & " Where Buyer = '" & eBuyer & "'"
                        Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
                        If dt_ControlRecord.Rows.Count > 0 Then eGRBuyer = dt_ControlRecord.Rows(0).Item("BuyerGroup")
                        '
                        If oDB.GetFunctionCode(eGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                            ' [Reference檢測]
                            If Not err Then
                                ' [R]→ORFN5K+ITMC5K / [Y]→ORFN5K / [P]→ PUCN5K+ITMC5K / [X]→ PUCN5K
                                If oDB.GetFunctionCode(eGRBuyer, 4) = "P" Or oDB.GetFunctionCode(eGRBuyer, 4) = "X" Then
                                Else
                                    If dt_SC760W1.Rows(i).Item("REFNW1") = "" Then
                                        err = True
                                        msg = "Check-PONO:[REFNW1 = Blank !]"
                                    End If
                                End If
                            End If
                            ' [GRPCW1檢測]
                            If Not err Then
                                If dt_SC760W1.Rows(i).Item("GRPCW1") = 0 Then
                                    err = True
                                    msg = "Check-PONO:[GRPCW1 = 0 !]"
                                End If
                            End If
                            ' [PONO檢測]
                            If Not err Then
                                If dt_SC760W1.Rows(i).Item("ORDNW1") = "" Then
                                    err = True
                                    msg = "Check-PONO:[ORDNW1 = Blank !]"
                                End If
                            End If
                            ' [PONO-SEQ檢測]
                            If Not err Then
                                If dt_SC760W1.Rows(i).Item("OSBNW1") = 0 Then
                                    err = True
                                    msg = "Check-PONO:[OSBNW1 = 0 !]"
                                End If
                            End If
                            ' [檢測異常處理]
                            If err Then
                                RtnCode = 1
                                oDB.AccessLog(pLogID, pBuyer, "CheckPONO", NowDateTime, 1, CStr(Row), msg, "=", CStr(RtnCode), "", pUserID, "")
                                Exit For
                            End If
                        End If
                        Row = Row + 1
                    Next
                End If
                ' **EOES-END
            Else
                ' **原程式-START
                ' S5K00-檢測PONO
                Row = 1
                sql = "SELECT * FROM S5K00 "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "Order by Unique_ID "
                Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                For i = 0 To dt_S5K00.Rows.Count - 1
                    ' ActionLog-Key欄
                    eCORN5K = dt_S5K00.Rows(i).Item("CORN5K").ToString
                    eGRPC5K = dt_S5K00.Rows(i).Item("GRPC5K")

                    ' [Reference檢測]
                    If Not err Then
                        ' [R]→ORFN5K+ITMC5K / [Y]→ORFN5K / [P]→ PUCN5K+ITMC5K / [X]→ PUCN5K
                        If oDB.GetFunctionCode(pGRBuyer, 4) = "P" Or oDB.GetFunctionCode(pGRBuyer, 4) = "X" Then
                            If dt_S5K00.Rows(i).Item("PUCN5K") = "" Then
                                err = True
                                msg = "Check-PONO:[PUCN5K = Blank !]"
                            End If
                        Else
                            If dt_S5K00.Rows(i).Item("ORFN5K") = "" Then
                                err = True
                                msg = "Check-PONO:[ORFN5K = Blank !]"
                            End If
                        End If
                    End If
                    ' [GRPC5K檢測]
                    If Not err Then
                        If dt_S5K00.Rows(i).Item("GRPC5K") = 0 Then
                            err = True
                            msg = "Check-PONO:[GRPC5K = 0 !]"
                        End If
                    End If
                    ' [PONO檢測]
                    If Not err Then
                        If dt_S5K00.Rows(i).Item("PODN5K") = "" Then
                            err = True
                            msg = "Check-PONO:[PODN5K = Blank !]"
                        End If
                    End If
                    ' [PONO-SEQ檢測]
                    If Not err Then
                        If dt_S5K00.Rows(i).Item("POSN5K") = 0 Then
                            err = True
                            msg = "Check-PONO:[POSN5K = 0 !]"
                        End If
                    End If
                    ' [檢測異常處理]
                    If err Then
                        RtnCode = 1
                        oDB.AccessLog(pLogID, pBuyer, "CheckPONO", NowDateTime, 1, oDB.GetAccessLogLine(pLogID, pBuyer, eCORN5K, eGRPC5K), msg, "=", CStr(RtnCode), "", pUserID, "")
                        Exit For
                    Else
                        Row = Row + 1
                    End If
                Next
                '
                ' S5L00-檢測PONO
                If Not err Then
                    Row = 1
                    sql = "SELECT * FROM S5L00 "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "Order by Unique_ID "
                    Dim dt_S5L00 As DataTable = oDataBase.GetDataTable(sql)
                    For i = 0 To dt_S5L00.Rows.Count - 1
                        ' [PONO檢測]
                        If Not err Then
                            If dt_S5L00.Rows(i).Item("PODN5L") = "" Then
                                err = True
                                msg = "Check-PONO:[PODN5L = Blank !]"
                            End If
                        End If
                        ' [檢測異常處理]
                        If err Then
                            RtnCode = 1
                            oDB.AccessLog(pLogID, pBuyer, "CheckPONO", NowDateTime, 1, CStr(Row), msg, "=", CStr(RtnCode), "", pUserID, "")
                            Exit For
                        Else
                            Row = Row + 1
                        End If
                    Next
                End If
                '
                ' SC760W1-檢測PONO
                If Not err Then
                    If oDB.GetFunctionCode(pGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                        Row = 1
                        sql = "SELECT * FROM SC760W1 "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "Order by Unique_ID "
                        Dim dt_SC760W1 As DataTable = oDataBase.GetDataTable(sql)
                        For i = 0 To dt_SC760W1.Rows.Count - 1
                            ' [Reference檢測]
                            If Not err Then
                                ' [R]→ORFN5K+ITMC5K / [Y]→ORFN5K / [P]→ PUCN5K+ITMC5K / [X]→ PUCN5K
                                If oDB.GetFunctionCode(pGRBuyer, 4) = "P" Or oDB.GetFunctionCode(pGRBuyer, 4) = "X" Then
                                Else
                                    If dt_SC760W1.Rows(i).Item("REFNW1") = "" Then
                                        err = True
                                        msg = "Check-PONO:[REFNW1 = Blank !]"
                                    End If
                                End If
                            End If
                            ' [GRPCW1檢測]
                            If Not err Then
                                If dt_SC760W1.Rows(i).Item("GRPCW1") = 0 Then
                                    err = True
                                    msg = "Check-PONO:[GRPCW1 = 0 !]"
                                End If
                            End If
                            ' [PONO檢測]
                            If Not err Then
                                If dt_SC760W1.Rows(i).Item("ORDNW1") = "" Then
                                    err = True
                                    msg = "Check-PONO:[ORDNW1 = Blank !]"
                                End If
                            End If
                            ' [PONO-SEQ檢測]
                            If Not err Then
                                If dt_SC760W1.Rows(i).Item("OSBNW1") = 0 Then
                                    err = True
                                    msg = "Check-PONO:[OSBNW1 = 0 !]"
                                End If
                            End If
                            ' [檢測異常處理]
                            If err Then
                                RtnCode = 1
                                oDB.AccessLog(pLogID, pBuyer, "CheckPONO", NowDateTime, 1, CStr(Row), msg, "=", CStr(RtnCode), "", pUserID, "")
                                Exit For
                            Else
                                Row = Row + 1
                            End If
                        Next
                    End If
                End If
                ' **原程式-END
            End If
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "CheckPONO", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'CheckPONO-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([120]-MakeGRPC)
    '**     製作Group Code 
    '***********************************************************************************************
    'MakeGRPC-Start
    <WebMethod()> _
    Public Function MakeGRPC(ByVal pLogID As String, _
                             ByVal pBuyer As String, _
                             ByVal pUserID As String, _
                             ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim i, j As Integer
        '------------------------------------------
        ' 2017/5/23 EOES-多客戶多BUYER
        Dim eBuyer As String = ""
        Dim eGRBuyer As String = ""
        '
        Try
            '[EOES使用]
            If InStr(pBuyer, "EOES-") > 0 Then
                ' **EOES-START
                sql = "SELECT PODN5K, ITMC5K FROM S5K00 "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "Group by PODN5K, ITMC5K  "
                sql &= "Order by PODN5K, ITMC5K  "
                Dim dt_S5K00A As DataTable = oDataBase.GetDataTable(sql)
                For i = 0 To dt_S5K00A.Rows.Count - 1
                    '
                    'JJJ-對應
                    'sql = "SELECT * FROM S5K00 "
                    sql = "Select *, Case SUBSTRING(RFCC5K, 1, 4) WHEN 'ZJJJ' THEN RFCC5K + '-' + '000999' ELSE RFCC5K + '-' + RCMC5K END  As CustBuyer "
                    sql &= "From S5K00 "
                    '<<<
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "  And PODN5K = '" & dt_S5K00A.Rows(i).Item("PODN5K") & "' "
                    sql &= "  And ITMC5K = '" & dt_S5K00A.Rows(i).Item("ITMC5K") & "' "
                    sql &= "Order by Unique_ID  "
                    Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                    For j = 0 To dt_S5K00.Rows.Count - 1
                        '
                        ' 取得 BUYERGROUP
                        'JJJ-對應
                        'eBuyer = dt_S5K00.Rows(j).Item("RFCC5K") + "-" + dt_S5K00.Rows(j).Item("RCMC5K")
                        eBuyer = dt_S5K00.Rows(j).Item("CustBuyer")
                        '<<<

                        sql = "Select BuyerGroup From M_ControlRecord "
                        sql = sql & " Where Buyer = '" & eBuyer & "'"
                        Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
                        If dt_ControlRecord.Rows.Count > 0 Then eGRBuyer = dt_ControlRecord.Rows(0).Item("BuyerGroup")
                        '
                        ' S5K00-更新
                        '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                        If oDB.GetFunctionCode(eGRBuyer, 3) = "P" Then
                            sql = "Update S5K00 Set "
                            sql = sql + " GRPC5K = '" & CStr(j + 1) & "' "
                        Else
                            '-MODIFY-START 201208 (失效不處理)
                            '-MODIFY-START 240926 (再復活)
                            sql = "Update S5K00 Set "
                            sql = sql + " GRPC5K = POSN5K "
                            'sql = sql + " GRPC5K = GRPC5K "
                            '-END
                        End If
                        sql = sql + " Where Unique_ID = '" & dt_S5K00.Rows(j).Item("Unique_ID").ToString & "'"
                        oDataBase.ExecuteNonQuery(sql)
                        '
                        ' SC760W1-更新
                        If oDB.GetFunctionCode(eGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                            '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                            If oDB.GetFunctionCode(eGRBuyer, 3) = "P" Then
                                sql = "Update SC760W1 Set "
                                sql = sql + " GRPCW1 = '" & CStr(j + 1) & "' "
                            Else
                                '-MODIFY-START 201208 (失效不處理)
                                '-MODIFY-START 240926 (再復活)
                                sql = "Update SC760W1 Set "
                                sql = sql + " GRPCW1 = OSBNW1 "
                                'sql = sql + " GRPCW1 = GRPCW1 "
                                '-END
                            End If
                            sql = sql + "Where Buyer = '" & pBuyer & "' "
                            sql = sql + "  And ORDNW1 = '" & dt_S5K00.Rows(j).Item("PODN5K") & "' "
                            sql = sql + "  And ID = '" & dt_S5K00.Rows(j).Item("ID").ToString & "' "
                            oDataBase.ExecuteNonQuery(sql)
                        End If
                    Next
                Next
                ' **EOES-END
            Else
                ' **原程式-START
                sql = "SELECT PODN5K, ITMC5K FROM S5K00 "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "Group by PODN5K, ITMC5K  "
                sql &= "Order by PODN5K, ITMC5K  "
                Dim dt_S5K00A As DataTable = oDataBase.GetDataTable(sql)
                For i = 0 To dt_S5K00A.Rows.Count - 1
                    '
                    sql = "SELECT * FROM S5K00 "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "  And PODN5K = '" & dt_S5K00A.Rows(i).Item("PODN5K") & "' "
                    sql &= "  And ITMC5K = '" & dt_S5K00A.Rows(i).Item("ITMC5K") & "' "
                    sql &= "Order by Unique_ID  "
                    Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                    For j = 0 To dt_S5K00.Rows.Count - 1
                        '
                        ' S5K00-更新
                        '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                        If oDB.GetFunctionCode(pGRBuyer, 3) = "P" Then
                            sql = "Update S5K00 Set "
                            sql = sql + " GRPC5K = '" & CStr(j + 1) & "' "
                        Else
                            sql = "Update S5K00 Set "
                            sql = sql + " GRPC5K = POSN5K "
                        End If
                        sql = sql + " Where Unique_ID = '" & dt_S5K00.Rows(j).Item("Unique_ID").ToString & "'"
                        oDataBase.ExecuteNonQuery(sql)
                        '
                        ' SC760W1-更新
                        If oDB.GetFunctionCode(pGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                            '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                            If oDB.GetFunctionCode(pGRBuyer, 3) = "P" Then
                                sql = "Update SC760W1 Set "
                                sql = sql + " GRPCW1 = '" & CStr(j + 1) & "' "
                            Else
                                sql = "Update SC760W1 Set "
                                sql = sql + " GRPCW1 = OSBNW1 "
                            End If
                            sql = sql + "Where Buyer = '" & pBuyer & "' "
                            sql = sql + "  And ORDNW1 = '" & dt_S5K00.Rows(j).Item("PODN5K") & "' "
                            sql = sql + "  And ID = '" & dt_S5K00.Rows(j).Item("ID").ToString & "' "
                            oDataBase.ExecuteNonQuery(sql)
                        End If
                    Next
                Next
                ' **原程式-END
            End If
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "MakeGRPC", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'MakeGRPC-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([130]-CheckGRPC)
    '**     檢測Group Code
    '***********************************************************************************************
    'CheckGRPC-Start
    <WebMethod()> _
    Public Function CheckGRPC(ByVal pLogID As String, _
                              ByVal pBuyer As String, _
                              ByVal pUserID As String, _
                              ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql, xPO, xItem As String
        Dim i, Row, xGRPC, rGRPC As Integer
        Dim err As Boolean = False
        Dim msg As String = ""
        ' ActionLog-Key欄
        Dim eCORN5K As String = ""
        Dim eGRPC5K As Integer = 0
        '------------------------------------------
        ' 2017/5/23 EOES-多客戶多BUYER
        Dim eBuyer As String = ""
        Dim eGRBuyer As String = ""
        '
        Try
            '[EOES使用]
            If InStr(pBuyer, "EOES-") > 0 Then
                ' **EOES-START
                ' S5K00-檢測GRPC
                Row = 1
                xGRPC = 1
                'JJJ-對應
                'sql = "SELECT * FROM S5K00 "
                sql = "Select *, Case SUBSTRING(RFCC5K, 1, 4) WHEN 'ZJJJ' THEN RFCC5K + '-' + '000999' ELSE RFCC5K + '-' + RCMC5K END  As CustBuyer "
                sql &= "From S5K00 "
                '<<<
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "Order by PODN5K, ITMC5K, GRPC5K "
                Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                If dt_S5K00.Rows.Count > 0 Then
                    xPO = dt_S5K00.Rows(0).Item("PODN5K")
                    xItem = dt_S5K00.Rows(0).Item("ITMC5K")
                End If
                For i = 0 To dt_S5K00.Rows.Count - 1
                    ' ActionLog-Key欄
                    eCORN5K = dt_S5K00.Rows(i).Item("CORN5K").ToString
                    eGRPC5K = dt_S5K00.Rows(i).Item("GRPC5K")

                    rGRPC = dt_S5K00.Rows(i).Item("GRPC5K")
                    ' [GPRC檢測-1]
                    '-MODIFY-START 201208 (失效不處理)
                    '-MODIFY-START 20240926 (再復活)
                    If Not err Then
                        If dt_S5K00.Rows(i).Item("POSN5K") <> dt_S5K00.Rows(i).Item("GRPC5K") Then
                            err = True
                            msg = "Check-GRPC:[POSN5K <> GPRC5K (PODN5K/ITMC5K) !]"
                        End If
                    End If
                    '-END
                    ' [GPRC檢測-2]
                    If Not err Then
                        '
                        ' 取得 BUYERGROUP
                        'JJJ-對應
                        'eBuyer = dt_S5K00.Rows(i).Item("RFCC5K") + "-" + dt_S5K00.Rows(i).Item("RCMC5K")
                        eBuyer = dt_S5K00.Rows(i).Item("CustBuyer")
                        '<<<
                        sql = "Select BuyerGroup From M_ControlRecord "
                        sql = sql & " Where Buyer = '" & eBuyer & "'"
                        Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
                        If dt_ControlRecord.Rows.Count > 0 Then eGRBuyer = dt_ControlRecord.Rows(0).Item("BuyerGroup")
                        '
                        '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                        If oDB.GetFunctionCode(eGRBuyer, 3) = "P" Then
                            If xPO <> dt_S5K00.Rows(i).Item("PODN5K") Or xItem <> dt_S5K00.Rows(i).Item("ITMC5K") Then
                                xPO = dt_S5K00.Rows(i).Item("PODN5K")
                                xItem = dt_S5K00.Rows(i).Item("ITMC5K")
                                xGRPC = 1
                            End If
                            If rGRPC <> xGRPC Then
                                err = True
                                msg = "Check-GRPC:[GPRC5K Error (PODN5K/ITMC5K) !]"
                            End If
                        End If
                    End If
                    ' [檢測異常處理]
                    If err Then
                        RtnCode = 1
                        oDB.AccessLog(pLogID, pBuyer, "CheckGRPC", NowDateTime, 1, oDB.GetAccessLogLine(pLogID, pBuyer, eCORN5K, eGRPC5K), msg, "=", CStr(RtnCode), "", pUserID, "")
                        Exit For
                    Else
                        xGRPC = xGRPC + 1
                        Row = Row + 1
                    End If
                Next
                '
                ' SC760W1-檢測GRPC
                If Not err Then
                    Row = 1
                    'JJJ-對應
                    'sql = "SELECT PODN5K, POSN5K, GRPC5K, RFCC5K, RCMC5K FROM S5K00 "
                    sql = "Select PODN5K, POSN5K, GRPC5K, RFCC5K, RCMC5K, "
                    sql &= "Case SUBSTRING(RFCC5K, 1, 4) WHEN 'ZJJJ' THEN RFCC5K + '-' + '000999' ELSE RFCC5K + '-' + RCMC5K END  As CustBuyer "
                    sql &= "From S5K00 "
                    '<<<
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "Order by PODN5K, POSN5K, GRPC5K "
                    Dim dt_S5K00x As DataTable = oDataBase.GetDataTable(sql)
                    For i = 0 To dt_S5K00x.Rows.Count - 1
                        '
                        ' 取得 BUYERGROUP
                        'JJJ-對應
                        'eBuyer = dt_S5K00x.Rows(i).Item("RFCC5K") + "-" + dt_S5K00x.Rows(i).Item("RCMC5K")
                        eBuyer = dt_S5K00.Rows(i).Item("CustBuyer")
                        '<<<

                        sql = "Select BuyerGroup From M_ControlRecord "
                        sql = sql & " Where Buyer = '" & eBuyer & "' "
                        Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
                        If dt_ControlRecord.Rows.Count > 0 Then eGRBuyer = dt_ControlRecord.Rows(0).Item("BuyerGroup")
                        '
                        ' [GPRC檢測-1]
                        If oDB.GetFunctionCode(eGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                            If Not err Then
                                sql = "SELECT * FROM SC760W1 "
                                sql &= "Where ORDNW1 = '" & dt_S5K00x.Rows(i).Item("PODN5K") & "' "
                                sql &= "  And OSBNW1 = '" & dt_S5K00x.Rows(i).Item("POSN5K").ToString & "' "
                                sql &= "  And GRPCW1 = '" & dt_S5K00x.Rows(i).Item("GRPC5K").ToString & "' "
                                Dim dt_SC760W1 As DataTable = oDataBase.GetDataTable(sql)
                                If dt_SC760W1.Rows.Count <= 0 Then
                                    err = True
                                    msg = "Check-GRPC:[GPRCW1 (S5K00 <> SC760W1) !]"
                                End If
                            End If
                            ' [檢測異常處理]
                            If err Then
                                RtnCode = 1
                                oDB.AccessLog(pLogID, pBuyer, "CheckGRPC", NowDateTime, 1, CStr(Row), msg, "=", CStr(RtnCode), "", pUserID, "")
                                Exit For
                            End If
                        End If
                        Row = Row + 1
                    Next
                End If
                ' **EOES-END
            Else
                ' **原程式-START
                ' S5K00-檢測GRPC
                Row = 1
                xGRPC = 1
                sql = "SELECT * FROM S5K00 "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "Order by PODN5K, ITMC5K, GRPC5K "
                Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                If dt_S5K00.Rows.Count > 0 Then
                    xPO = dt_S5K00.Rows(0).Item("PODN5K")
                    xItem = dt_S5K00.Rows(0).Item("ITMC5K")
                End If
                For i = 0 To dt_S5K00.Rows.Count - 1
                    ' ActionLog-Key欄
                    eCORN5K = dt_S5K00.Rows(i).Item("CORN5K").ToString
                    eGRPC5K = dt_S5K00.Rows(i).Item("GRPC5K")

                    rGRPC = dt_S5K00.Rows(i).Item("GRPC5K")
                    ' [GPRC檢測-1]
                    If Not err Then
                        If dt_S5K00.Rows(i).Item("POSN5K") <> dt_S5K00.Rows(i).Item("GRPC5K") Then
                            err = True
                            msg = "Check-GRPC:[POSN5K <> GPRC5K (PODN5K/ITMC5K) !]"
                        End If
                    End If
                    ' [GPRC檢測-2]
                    If Not err Then
                        '使用自動PO(-SUBNO)  [X]→客戶提供 / [P]→系統自動產生
                        If oDB.GetFunctionCode(pGRBuyer, 3) = "P" Then
                            If xPO <> dt_S5K00.Rows(i).Item("PODN5K") Or xItem <> dt_S5K00.Rows(i).Item("ITMC5K") Then
                                xPO = dt_S5K00.Rows(i).Item("PODN5K")
                                xItem = dt_S5K00.Rows(i).Item("ITMC5K")
                                xGRPC = 1
                            End If
                            If rGRPC <> xGRPC Then
                                err = True
                                msg = "Check-GRPC:[GPRC5K Error (PODN5K/ITMC5K) !]"
                            End If
                        End If
                    End If
                    ' [檢測異常處理]
                    If err Then
                        RtnCode = 1
                        oDB.AccessLog(pLogID, pBuyer, "CheckGRPC", NowDateTime, 1, oDB.GetAccessLogLine(pLogID, pBuyer, eCORN5K, eGRPC5K), msg, "=", CStr(RtnCode), "", pUserID, "")
                        Exit For
                    Else
                        xGRPC = xGRPC + 1
                        Row = Row + 1
                    End If
                Next
                '
                ' SC760W1-檢測GRPC
                If Not err Then
                    If oDB.GetFunctionCode(pGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                        Row = 1
                        sql = "SELECT PODN5K, POSN5K, GRPC5K FROM S5K00 "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "Order by Unique_ID "
                        Dim dt_S5K00x As DataTable = oDataBase.GetDataTable(sql)
                        For i = 0 To dt_S5K00x.Rows.Count - 1
                            ' [GPRC檢測-1]
                            If Not err Then
                                sql = "SELECT * FROM SC760W1 "
                                sql &= "Where ORDNW1 = '" & dt_S5K00x.Rows(i).Item("PODN5K") & "' "
                                sql &= "  And OSBNW1 = '" & dt_S5K00x.Rows(i).Item("POSN5K").ToString & "' "
                                sql &= "  And GRPCW1 = '" & dt_S5K00x.Rows(i).Item("GRPC5K").ToString & "' "
                                Dim dt_SC760W1 As DataTable = oDataBase.GetDataTable(sql)
                                If dt_SC760W1.Rows.Count <= 0 Then
                                    err = True
                                    msg = "Check-GRPC:[GPRCW1 (S5K00 <> SC760W1) !]"
                                End If
                            End If
                            ' [檢測異常處理]
                            If err Then
                                RtnCode = 1
                                oDB.AccessLog(pLogID, pBuyer, "CheckGRPC", NowDateTime, 1, CStr(Row), msg, "=", CStr(RtnCode), "", pUserID, "")
                                Exit For
                            Else
                                Row = Row + 1
                            End If
                        Next
                    End If
                End If
                ' **原程式-END
            End If
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "CheckGRPC", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'CheckGRPC-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**[135]-CheckSampleQty 
    '**     檢測Sample Qty (不可SampleQty>600)
    '***********************************************************************************************
    'CheckSampleQty-Start
    <WebMethod()> _
    Public Function CheckSampleQty(ByVal pLogID As String, _
                                   ByVal pBuyer As String, _
                                   ByVal pUserID As String, _
                                   ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql, msg, msgcontent As String
        Dim i As Integer
        '
        Try
            If pBuyer <> "EOES-OUTSOURCE" Then
                '
                ' S5K00-檢測 SampleQty>600
                sql = "SELECT PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K, PODQ5K FROM S5K00 "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "  And SMPF5K = '" & "H" & "' "
                sql &= "  And PODQ5K > " & "600" & " "
                sql &= "Order by PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K "
                Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                If dt_S5K00.Rows.Count > 0 Then
                    msg = "Check-SampleQty:[樣品數>600 !]"
                    msgcontent = dt_S5K00.Rows(0).Item("PODN5K").ToString + "/" + _
                                 dt_S5K00.Rows(0).Item("POSN5K").ToString + "/" + _
                                 dt_S5K00.Rows(0).Item("PUCN5K").ToString + "/" + _
                                 dt_S5K00.Rows(0).Item("ORFN5K").ToString + "/" + _
                                 dt_S5K00.Rows(0).Item("ITMC5K").ToString + "/" + _
                                 dt_S5K00.Rows(0).Item("CLRC5K").ToString + "/" + _
                                 dt_S5K00.Rows(0).Item("PODQ5K").ToString
                    '    
                    RtnCode = 1
                    oDB.AccessLog(pLogID, pBuyer, "CheckSampleQty", NowDateTime, 1, msg, msgcontent, "=", CStr(RtnCode), "", pUserID, "")
                End If
                '
                ' 20230311-ADD-START BY JOY
                ' S5K00-檢測 [非SAMPLE使用 SAMPLE LINE]  (NISHIGUCHI)
                If RtnCode = 0 Then
                    sql = "SELECT PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K, PDOC5K, SMPF5K FROM S5K00 "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "  And SMPF5K = '" & "" & "' "
                    sql &= "  And PDOC5K = '" & "2" & "' "
                    sql &= "Order by PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K "
                    Dim dt_S5K00A As DataTable = oDataBase.GetDataTable(sql)
                    If dt_S5K00A.Rows.Count > 0 Then
                        msg = "Check-NOT SANPLE 不可使用 SAMPLE LINE !]"
                        msgcontent = dt_S5K00A.Rows(0).Item("PODN5K").ToString + "/" + _
                                     dt_S5K00A.Rows(0).Item("POSN5K").ToString + "/" + _
                                     dt_S5K00A.Rows(0).Item("PUCN5K").ToString + "/" + _
                                     dt_S5K00A.Rows(0).Item("ORFN5K").ToString + "/" + _
                                     dt_S5K00A.Rows(0).Item("ITMC5K").ToString + "/" + _
                                     dt_S5K00A.Rows(0).Item("CLRC5K").ToString + "/" + _
                                     dt_S5K00A.Rows(0).Item("PDOC5K").ToString
                        '    
                        RtnCode = 1
                        oDB.AccessLog(pLogID, pBuyer, "CheckSampleLine", NowDateTime, 1, msg, msgcontent, "=", CStr(RtnCode), "", pUserID, "")
                    End If
                End If
                ' S5K00-檢測 [LINE TYPE='1' EXPRESS LINE (NISHIGUCHI)
                If RtnCode = 0 Then
                    sql = "SELECT PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K, PDOC5K, SMPF5K FROM S5K00 "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "  And PDOC5K = '" & "1" & "' "
                    sql &= "Order by PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K "
                    Dim dt_S5K00B As DataTable = oDataBase.GetDataTable(sql)
                    If dt_S5K00B.Rows.Count > 0 Then
                        msg = "不可使用 LINE TYPE=1 EXPRESS LINE !]"
                        msgcontent = dt_S5K00B.Rows(0).Item("PODN5K").ToString + "/" + _
                                     dt_S5K00B.Rows(0).Item("POSN5K").ToString + "/" + _
                                     dt_S5K00B.Rows(0).Item("PUCN5K").ToString + "/" + _
                                     dt_S5K00B.Rows(0).Item("ORFN5K").ToString + "/" + _
                                     dt_S5K00B.Rows(0).Item("ITMC5K").ToString + "/" + _
                                     dt_S5K00B.Rows(0).Item("CLRC5K").ToString + "/" + _
                                     dt_S5K00B.Rows(0).Item("PDOC5K").ToString
                        '    
                        RtnCode = 1
                        oDB.AccessLog(pLogID, pBuyer, "CheckLineType", NowDateTime, 1, msg, msgcontent, "=", CStr(RtnCode), "", pUserID, "")
                    End If
                End If
                ' 20230311-ADD-END
                '
                ' ADD-START JOY 2023/4/11 
                ' **********************************************
                ' NO COMMERCIAL CHECK
                ' S5K00-檢測 (V_S5K00_ITEM)
                ' ST1CA0=1 (SF)
                '   SAMPLE ORDER
                '   240807 QTY = 600 --> 300
                If RtnCode = 0 Then
                    sql = "SELECT PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K, PODQ5K FROM V_S5K00_ITEM "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "  And ST1CA0 = '" & "1" & "' "
                    sql &= "  And NCMF5K = '" & "1" & "' "
                    sql &= "  And SMPF5K = '" & "H" & "' "
                    sql &= "  And PODQ5K > " & "300" & " "
                    sql &= "Order by PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K "
                    Dim dt_VS5K001 As DataTable = oDataBase.GetDataTable(sql)
                    If dt_VS5K001.Rows.Count > 0 Then
                        msg = "Check-NoComQty:[SF Sample無償數>300 !]"
                        msgcontent = dt_VS5K001.Rows(0).Item("PODN5K").ToString + "/" + _
                                     dt_VS5K001.Rows(0).Item("POSN5K").ToString + "/" + _
                                     dt_VS5K001.Rows(0).Item("PUCN5K").ToString + "/" + _
                                     dt_VS5K001.Rows(0).Item("ORFN5K").ToString + "/" + _
                                     dt_VS5K001.Rows(0).Item("ITMC5K").ToString + "/" + _
                                     dt_VS5K001.Rows(0).Item("CLRC5K").ToString + "/" + _
                                     dt_VS5K001.Rows(0).Item("PODQ5K").ToString
                        '    
                        RtnCode = 1
                        oDB.AccessLog(pLogID, pBuyer, "CheckNoComQty", NowDateTime, 1, msg, msgcontent, "=", CStr(RtnCode), "", pUserID, "")
                    End If
                End If
                '
                ' ST1CA0=1 (SF)
                '   BULK ORDER
                If RtnCode = 0 Then
                    sql = "SELECT PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K, PODQ5K FROM V_S5K00_ITEM "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "  And ST1CA0 = '" & "1" & "' "
                    sql &= "  And NCMF5K = '" & "1" & "' "
                    sql &= "  And SMPF5K = '" & "" & "' "
                    sql &= "  And PODQ5K > " & "0" & " "
                    sql &= "Order by PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K "
                    Dim dt_VS5K002 As DataTable = oDataBase.GetDataTable(sql)
                    If dt_VS5K002.Rows.Count > 0 Then
                        msg = "Check-NoComQty:[SF Bulk無償數>0 !]"
                        msgcontent = dt_VS5K002.Rows(0).Item("PODN5K").ToString + "/" + _
                                     dt_VS5K002.Rows(0).Item("POSN5K").ToString + "/" + _
                                     dt_VS5K002.Rows(0).Item("PUCN5K").ToString + "/" + _
                                     dt_VS5K002.Rows(0).Item("ORFN5K").ToString + "/" + _
                                     dt_VS5K002.Rows(0).Item("ITMC5K").ToString + "/" + _
                                     dt_VS5K002.Rows(0).Item("CLRC5K").ToString + "/" + _
                                     dt_VS5K002.Rows(0).Item("PODQ5K").ToString
                        '    
                        RtnCode = 1
                        oDB.AccessLog(pLogID, pBuyer, "CheckNoComQty", NowDateTime, 1, msg, msgcontent, "=", CStr(RtnCode), "", pUserID, "")
                    End If
                End If
                ' -------------------------------------------------------------------------------
                ' ST1CA0 + ST2CA0 = 32 (T&P)
                '   SAMPLE ORDER
                If RtnCode = 0 Then
                    sql = "SELECT PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K, PODQ5K FROM V_S5K00_ITEM "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "  And ST1CA0 + ST2CA0 = '" & "32" & "' "
                    sql &= "  And NCMF5K = '" & "1" & "' "
                    sql &= "  And SMPF5K = '" & "H" & "' "
                    sql &= "  And PODQ5K > " & "100" & " "
                    sql &= "Order by PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K "
                    Dim dt_VS5K003 As DataTable = oDataBase.GetDataTable(sql)
                    If dt_VS5K003.Rows.Count > 0 Then
                        msg = "Check-NoComQty:[TP Sample無償數>100 !]"
                        msgcontent = dt_VS5K003.Rows(0).Item("PODN5K").ToString + "/" + _
                                     dt_VS5K003.Rows(0).Item("POSN5K").ToString + "/" + _
                                     dt_VS5K003.Rows(0).Item("PUCN5K").ToString + "/" + _
                                     dt_VS5K003.Rows(0).Item("ORFN5K").ToString + "/" + _
                                     dt_VS5K003.Rows(0).Item("ITMC5K").ToString + "/" + _
                                     dt_VS5K003.Rows(0).Item("CLRC5K").ToString + "/" + _
                                     dt_VS5K003.Rows(0).Item("PODQ5K").ToString
                        '    
                        RtnCode = 1
                        oDB.AccessLog(pLogID, pBuyer, "CheckNoComQty", NowDateTime, 1, msg, msgcontent, "=", CStr(RtnCode), "", pUserID, "")
                    End If
                End If
                '
                ' ST1CA0 + ST2CA0 = 32 (T&P)
                '   BULK ORDER
                If RtnCode = 0 Then
                    sql = "SELECT PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K, PODQ5K FROM V_S5K00_ITEM "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "  And ST1CA0 + ST2CA0 = '" & "32" & "' "
                    sql &= "  And NCMF5K = '" & "1" & "' "
                    sql &= "  And SMPF5K = '" & "" & "' "
                    sql &= "  And PODQ5K >= " & "0" & " "
                    sql &= "Order by PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K "
                    Dim dt_VS5K004 As DataTable = oDataBase.GetDataTable(sql)
                    If dt_VS5K004.Rows.Count > 0 Then
                        msg = "Check-NoComQty:[TP Bulk無償數>0 !]"
                        msgcontent = dt_VS5K004.Rows(0).Item("PODN5K").ToString + "/" + _
                                     dt_VS5K004.Rows(0).Item("POSN5K").ToString + "/" + _
                                     dt_VS5K004.Rows(0).Item("PUCN5K").ToString + "/" + _
                                     dt_VS5K004.Rows(0).Item("ORFN5K").ToString + "/" + _
                                     dt_VS5K004.Rows(0).Item("ITMC5K").ToString + "/" + _
                                     dt_VS5K004.Rows(0).Item("CLRC5K").ToString + "/" + _
                                     dt_VS5K004.Rows(0).Item("PODQ5K").ToString
                        '    
                        RtnCode = 1
                        oDB.AccessLog(pLogID, pBuyer, "CheckNoComQty", NowDateTime, 1, msg, msgcontent, "=", CStr(RtnCode), "", pUserID, "")
                    End If
                End If
                ' -------------------------------------------------------------------------------
                ' ST1CA0 + ST2CA0 + ST3CA0 = 31L (S&B)
                '   SAMPLE ORDER
                If RtnCode = 0 Then
                    sql = "SELECT PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K, PODQ5K FROM V_S5K00_ITEM "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "  And ST1CA0 + ST2CA0 + ST3CA0 = '" & "31L" & "' "
                    sql &= "  And NCMF5K = '" & "1" & "' "
                    sql &= "  And SMPF5K = '" & "H" & "' "
                    sql &= "  And PODQ5K > " & "10" & " "
                    sql &= "Order by PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K "
                    Dim dt_VS5K005 As DataTable = oDataBase.GetDataTable(sql)
                    If dt_VS5K005.Rows.Count > 0 Then
                        msg = "Check-NoComQty:[SB Sample無償數>10 !]"
                        msgcontent = dt_VS5K005.Rows(0).Item("PODN5K").ToString + "/" + _
                                     dt_VS5K005.Rows(0).Item("POSN5K").ToString + "/" + _
                                     dt_VS5K005.Rows(0).Item("PUCN5K").ToString + "/" + _
                                     dt_VS5K005.Rows(0).Item("ORFN5K").ToString + "/" + _
                                     dt_VS5K005.Rows(0).Item("ITMC5K").ToString + "/" + _
                                     dt_VS5K005.Rows(0).Item("CLRC5K").ToString + "/" + _
                                     dt_VS5K005.Rows(0).Item("PODQ5K").ToString
                        '    
                        RtnCode = 1
                        oDB.AccessLog(pLogID, pBuyer, "CheckNoComQty", NowDateTime, 1, msg, msgcontent, "=", CStr(RtnCode), "", pUserID, "")
                    End If
                End If
                '
                ' ST1CA0 + ST2CA0 + ST3CA0 = 31L (S&B)
                '   BULK ORDER
                If RtnCode = 0 Then
                    sql = "SELECT PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K, PODQ5K FROM V_S5K00_ITEM "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "  And ST1CA0 + ST2CA0 + ST3CA0 = '" & "31L" & "' "
                    sql &= "  And NCMF5K = '" & "1" & "' "
                    sql &= "  And SMPF5K = '" & "" & "' "
                    sql &= "  And PODQ5K >= " & "0" & " "
                    sql &= "Order by PODN5K, POSN5K, PUCN5K, ORFN5K, ITMC5K, CLRC5K "
                    Dim dt_VS5K006 As DataTable = oDataBase.GetDataTable(sql)
                    If dt_VS5K006.Rows.Count > 0 Then
                        msg = "Check-NoComQty:[SB Bulk無償數>0 !]"
                        msgcontent = dt_VS5K006.Rows(0).Item("PODN5K").ToString + "/" + _
                                     dt_VS5K006.Rows(0).Item("POSN5K").ToString + "/" + _
                                     dt_VS5K006.Rows(0).Item("PUCN5K").ToString + "/" + _
                                     dt_VS5K006.Rows(0).Item("ORFN5K").ToString + "/" + _
                                     dt_VS5K006.Rows(0).Item("ITMC5K").ToString + "/" + _
                                     dt_VS5K006.Rows(0).Item("CLRC5K").ToString + "/" + _
                                     dt_VS5K006.Rows(0).Item("PODQ5K").ToString
                        '    
                        RtnCode = 1
                        oDB.AccessLog(pLogID, pBuyer, "CheckNoComQty", NowDateTime, 1, msg, msgcontent, "=", CStr(RtnCode), "", pUserID, "")
                    End If
                End If
                '
                ' ADD-END JOY 2023/4/11 
            End If
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "CheckSampleQty", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'CheckSampleQty-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([140]-CheckCompanyCode)
    '**     檢測Company Code 
    '***********************************************************************************************
    'CheckCompanyCode-Start
    <WebMethod()> _
    Public Function CheckCompanyCode(ByVal pLogID As String, _
                                     ByVal pBuyer As String, _
                                     ByVal pUserID As String, _
                                     ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim i, Row As Integer
        Dim err As Boolean = False
        Dim msg As String = ""
        ' ActionLog-Key欄
        Dim eCORN5K As String = ""
        Dim eGRPC5K As Integer = 0
        '
        Try
            ' LS Order時不列入檢測
            ' ---------------------------------------------
            If Mid(pBuyer, 1, 3) <> "FCT" Then
                ' S5K00-檢測
                Row = 1
                sql = "SELECT * FROM S5K00 "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "Order by Unique_ID "
                Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                For i = 0 To dt_S5K00.Rows.Count - 1
                    ' ActionLog-Key欄
                    eCORN5K = dt_S5K00.Rows(i).Item("CORN5K").ToString
                    eGRPC5K = dt_S5K00.Rows(i).Item("GRPC5K")

                    ' [CompanyCode檢測-1]
                    If Not err Then
                        If dt_S5K00.Rows(i).Item("DK1C5K") = "" Then
                            err = True
                            msg = "Check-CompanyCode:[DK1C5K=Blank !]"
                        End If
                    End If
                    ' [CompanyCode檢測-2]
                    If Not err Then
                        If dt_S5K00.Rows(i).Item("RFCC5K") = "" Then
                            err = True
                            msg = "Check-CompanyCode:[RFCC5K=Blank !]"
                        End If
                    End If
                    ' [CompanyCode檢測-3]
                    If Not err Then
                        'If dt_S5K00.Rows(i).Item("DK1C5K") <> dt_S5K00.Rows(i).Item("RFCC5K") Then
                        '    err = True
                        '    msg = "Check-CompanyCode:[DK1C5K<>RFCC5K !]"
                        'End If
                    End If
                    ' [檢測異常處理]
                    If err Then
                        RtnCode = 1
                        oDB.AccessLog(pLogID, pBuyer, "CheckCompanyCode", NowDateTime, 1, oDB.GetAccessLogLine(pLogID, pBuyer, eCORN5K, eGRPC5K), msg, "=", CStr(RtnCode), "", pUserID, "")
                        Exit For
                    Else
                        Row = Row + 1
                    End If
                Next
                '
                ' S5L00-檢測
                If Not err Then
                    Row = 1
                    sql = "SELECT * FROM S5L00 "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "Order by Unique_ID "
                    Dim dt_S5L00 As DataTable = oDataBase.GetDataTable(sql)
                    For i = 0 To dt_S5L00.Rows.Count - 1
                        ' [CompanyCode檢測-1]
                        If Not err Then
                            If dt_S5L00.Rows(i).Item("RFCC5L") = "" Then
                                err = True
                                msg = "Check-CompanyCode:[RFCC5L=Blank !]"
                            End If
                        End If
                        ' [檢測異常處理]
                        If err Then
                            RtnCode = 1
                            oDB.AccessLog(pLogID, pBuyer, "CheckCompanyCode", NowDateTime, 1, CStr(Row), msg, "=", CStr(RtnCode), "", pUserID, "")
                            Exit For
                        Else
                            Row = Row + 1
                        End If
                    Next
                End If
                ''
                '' SC760W1-檢測  ADD-BY 2017/8/8 CHECK NIKE-ORDERTYPE
                'If Not err Then
                '    If oDB.GetFunctionCode(pGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N
                '        Row = 1
                '        ' 取得S5K00 NIKE
                '        sql = "SELECT * FROM S5K00 "
                '        sql &= "Where Buyer = '" & pBuyer & "' "
                '        sql &= "And   SUBSTRING(RCMC5K,1,6) = '" & "000013" & "' "
                '        Dim dt_S5K00A As DataTable = oDataBase.GetDataTable(sql)
                '        For i = 0 To dt_S5K00A.Rows.Count - 1
                '            sql = "SELECT * FROM SC760W1 "
                '            sql &= "Where ORDNW1 = '" & dt_S5K00A.Rows(i).Item("PODN5K") & "' "
                '            sql &= "  And OSBNW1 = '" & dt_S5K00A.Rows(i).Item("POSN5K").ToString & "' "
                '            sql &= "  And GRPCW1 = '" & dt_S5K00A.Rows(i).Item("GRPC5K").ToString & "' "
                '            sql &= "Order by Unique_ID "
                '            Dim dt_SC760W1 As DataTable = oDataBase.GetDataTable(sql)
                '            If dt_SC760W1.Rows.Count > 0 Then
                '                If dt_SC760W1.Rows(0).Item("NOICW1") = "" And dt_SC760W1.Rows(0).Item("SMTCW1") = "" Then
                '                    err = True
                '                    msg = "Check-NIKE-ORDERTYPE:[SMTCW1 = Blank !]"
                '                End If
                '                ' [檢測異常處理]
                '                If err Then
                '                    RtnCode = 1
                '                    oDB.AccessLog(pLogID, pBuyer, "Check-NK-ORDERTYPE", NowDateTime, 1, CStr(Row), msg, "=", CStr(RtnCode), "", pUserID, "")
                '                    Exit For
                '                Else
                '                    Row = Row + 1
                '                End If
                '            End If
                '        Next
                '    End If
                'End If
            End If
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "CheckCompanyCode", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'CheckCompanyCode-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([150]-CheckKeepCode)
    '**     檢測Keep Code 
    '***********************************************************************************************
    'CheckKeepCode-Start
    <WebMethod()> _
    Public Function CheckKeepCode(ByVal pLogID As String, _
                                  ByVal pBuyer As String, _
                                  ByVal pUserID As String, _
                                  ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            oWavesEDI.Timeout = 900 * 1000
            RtnCode = oWavesEDI.EDICheckKeepCode(pLogID, pBuyer, pUserID, pGRBuyer)
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "CheckKeepCode", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'CheckKeepCode-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([160]-MakeColorCode)
    '**     製作Color Code(對象:數字Color Code)
    '***********************************************************************************************
    'MakeColorCode-Start
    <WebMethod()> _
    Public Function MakeColorCode(ByVal pLogID As String, _
                                  ByVal pBuyer As String, _
                                  ByVal pUserID As String, _
                                  ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql, xColor As String
        Dim i As Integer
        '
        Try
            ' S5K00-檢測
            sql = "SELECT CLRC5K, Unique_ID FROM S5K00 "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "Order by Unique_ID "
            Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_S5K00.Rows.Count - 1
                xColor = dt_S5K00.Rows(i).Item("CLRC5K")
                If IsNumeric(Trim(xColor)) Then
                    If Len(Trim(xColor)) = 3 Then
                        xColor = "  " + Trim(xColor)
                    Else
                        If Len(Trim(xColor)) = 4 Then
                            xColor = " " + Trim(xColor)
                        End If
                    End If
                End If
                '
                ' S5K00-更新Color
                sql = "Update S5K00 Set "
                sql = sql + " CLRC5K = '" & xColor & "' "
                'sql = sql + " CLRC5K = '" & xColor & "', "
                'sql = sql + " DEVC5K = '" & "MakeColor" & "' "
                sql = sql + " Where Unique_ID = '" & dt_S5K00.Rows(i).Item("Unique_ID").ToString & "'"
                oDataBase.ExecuteNonQuery(sql)
            Next
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "MakeColorCode", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), CStr(i), pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'MakeColorCode-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([170]-CheckColorCode)
    '**     檢測Color Code 
    '***********************************************************************************************
    'CheckColorCode-Start
    <WebMethod()> _
    Public Function CheckColorCode(ByVal pLogID As String, _
                                   ByVal pBuyer As String, _
                                   ByVal pUserID As String, _
                                   ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            oWavesEDI.Timeout = 900 * 1000
            RtnCode = oWavesEDI.EDICheckColorCode(pLogID, pBuyer, pUserID, pGRBuyer)
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "CheckColorCode", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'CheckColorCode-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([180]-CheckItemCode)
    '**     檢測Item Code  (含PBR ITEM切替)
    '***********************************************************************************************
    'CheckItemCode-Start
    <WebMethod()> _
    Public Function CheckItemCode(ByVal pLogID As String, _
                                  ByVal pBuyer As String, _
                                  ByVal pUserID As String, _
                                  ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i As Integer
        Dim msg, msgcontent As String
        Dim sql As String
        '
        Try
            '* MODIFY-START 2021/6/25  PBR切替
            'oWavesEDI.Timeout = 900 * 1000
            'RtnCode = oWavesEDI.EDICheckItemCode(pLogID, pBuyer, pUserID, pGRBuyer)
            '
            sql = "SELECT ITMC5K, RCMC5K FROM S5K00 "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "GROUP BY ITMC5K, RCMC5K "
            sql &= "Order by ITMC5K, RCMC5K "
            Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_S5K00.Rows.Count - 1
                '
                sql = "SELECT Unique_ID, Cat, YCode, PCode FROM M_PBRItemConvert "
                sql &= "Where Cat IN ('PBR', 'CZT', 'NW2P', 'NATL', 'PFAS') "
                sql &= "And   YCode = '" & dt_S5K00.Rows(i).Item("ITMC5K") & "' "
                Dim dt_PBRITEM As DataTable = oDataBase.GetDataTable(sql)
                If dt_PBRITEM.Rows.Count > 0 Then
                    '
                    'MODIFY 231016 直接CHECK WINGS
                    '
                    If dt_PBRITEM.Rows(0).Item("Cat").ToString = "PFAS" Then
                        If oWaves.GetPFASBuyer(Mid(dt_S5K00.Rows(i).Item("RCMC5K"), 1, 6)) = 1 Then
                            msg = "NOT [0901 PFAS ITEM]"
                            msgcontent = dt_PBRITEM.Rows(0).Item("YCode").ToString
                            RtnCode = 1
                            oDB.AccessLog(pLogID, pBuyer, "Check-Item", NowDateTime, 1, msg, msgcontent, "=", CStr(RtnCode), "", pUserID, "")
                            Exit For
                        End If
                    End If
                    '
                    'CAT= PFAS  230826
                    'If dt_PBRITEM.Rows(0).Item("Cat").ToString = "PFAS" Then
                    '    sql = "SELECT Unique_ID, Buyer FROM M_T8T10Buyer "
                    '    sql &= "Where Cat = '" & dt_PBRITEM.Rows(0).Item("Cat") & "' "
                    '    sql &= "And   Buyer = '" & Mid(dt_S5K00.Rows(i).Item("RCMC5K"), 1, 6) & "' "
                    '    Dim dt_T8T10BUYER As DataTable = oDataBase.GetDataTable(sql)
                    '    If dt_T8T10BUYER.Rows.Count <= 0 Then
                    '        msg = "NOT [0901 PFAS ITEM]"
                    '        msgcontent = dt_PBRITEM.Rows(0).Item("YCode").ToString
                    '        RtnCode = 1
                    '        oDB.AccessLog(pLogID, pBuyer, "Check-Item", NowDateTime, 1, msg, msgcontent, "=", CStr(RtnCode), "", pUserID, "")
                    '        Exit For
                    '    End If
                    'End If
                    '
                    'CAT= NATULON
                    If dt_PBRITEM.Rows(0).Item("Cat").ToString = "NATL" Then
                        sql = "SELECT Unique_ID, Buyer FROM M_T8T10Buyer "
                        sql &= "Where Cat = '" & dt_PBRITEM.Rows(0).Item("Cat") & "' "
                        sql &= "And   Buyer = '" & Mid(dt_S5K00.Rows(i).Item("RCMC5K"), 1, 6) & "' "
                        Dim dt_T8T10BUYER As DataTable = oDataBase.GetDataTable(sql)
                        If dt_T8T10BUYER.Rows.Count <= 0 Then
                            msg = "NOT [0722 NATULON ITEM]"
                            msgcontent = dt_PBRITEM.Rows(0).Item("YCode").ToString
                            RtnCode = 1
                            oDB.AccessLog(pLogID, pBuyer, "Check-Item", NowDateTime, 1, msg, msgcontent, "=", CStr(RtnCode), "", pUserID, "")
                            Exit For
                        End If
                    End If
                    '
                    'CAT= NO(PBR)
                    If dt_PBRITEM.Rows(0).Item("Cat").ToString = "PBR" Then
                        sql = "SELECT Unique_ID, Buyer FROM M_T8T10Buyer "
                        sql &= "Where Cat = '" & dt_PBRITEM.Rows(0).Item("Cat") & "' "
                        sql &= "And   Buyer = '" & Mid(dt_S5K00.Rows(i).Item("RCMC5K"), 1, 6) & "' "
                        Dim dt_T8T10BUYER As DataTable = oDataBase.GetDataTable(sql)
                        If dt_T8T10BUYER.Rows.Count <= 0 Then
                            msg = "NOT [PBR ITEM]"
                            msgcontent = dt_PBRITEM.Rows(0).Item("YCode").ToString
                            RtnCode = 1
                            oDB.AccessLog(pLogID, pBuyer, "Check-Item", NowDateTime, 1, msg, msgcontent, "=", CStr(RtnCode), "", pUserID, "")
                            Exit For
                        End If
                    End If
                    '
                    'CAT= T8/T8T10 (CZT)
                    If dt_PBRITEM.Rows(0).Item("Cat").ToString = "CZT" Then
                        sql = "SELECT Unique_ID, Buyer FROM M_T8T10Buyer "
                        sql &= "Where Cat = '" & dt_PBRITEM.Rows(0).Item("Cat") & "' "
                        sql &= "And   Buyer = '" & Mid(dt_S5K00.Rows(i).Item("RCMC5K"), 1, 6) & "' "
                        Dim dt_T8T10BUYER As DataTable = oDataBase.GetDataTable(sql)
                        If dt_T8T10BUYER.Rows.Count <= 0 Then
                            msg = "NOT [CZT ITEM]"
                            msgcontent = dt_PBRITEM.Rows(0).Item("YCode").ToString
                            RtnCode = 1
                            oDB.AccessLog(pLogID, pBuyer, "Check-Item", NowDateTime, 1, msg, msgcontent, "=", CStr(RtnCode), "", pUserID, "")
                            Exit For
                        End If
                    End If
                    '
                    'CAT= NW2P
                    If dt_PBRITEM.Rows(0).Item("Cat").ToString = "NW2P" Then
                        sql = "SELECT Unique_ID, Buyer FROM M_T8T10Buyer "
                        sql &= "Where Cat = '" & dt_PBRITEM.Rows(0).Item("Cat") & "' "
                        sql &= "And   Buyer = '" & Mid(dt_S5K00.Rows(i).Item("RCMC5K"), 1, 6) & "' "
                        Dim dt_T8T10BUYER As DataTable = oDataBase.GetDataTable(sql)
                        If dt_T8T10BUYER.Rows.Count <= 0 Then
                            msg = "NOT [OP-NW2P ITEM]"
                            msgcontent = dt_PBRITEM.Rows(0).Item("YCode").ToString
                            RtnCode = 1
                            oDB.AccessLog(pLogID, pBuyer, "Check-Item", NowDateTime, 1, msg, msgcontent, "=", CStr(RtnCode), "", pUserID, "")
                            Exit For
                        End If
                    End If
                End If
            Next
            '
            '* ADD-START 2022/12/07  RC CHAIN --> RCA CHAIN 切替
            sql = "SELECT ITMC5K, CLRC5K, RFCC5K, RCMC5K FROM S5K00 "
            sql &= "Where Buyer = '" & pBuyer & "' "
            '
            sql &= "And ITMC5K+CLRC5K in (Select YCODE+COLOR FROM V_RCF2RCACHAIN) "
            'sql &= "And ( CLRC5K = '  580' or CLRC5K = 'TNA12' ) "
            '
            sql &= "GROUP BY ITMC5K, CLRC5K, RFCC5K, RCMC5K "
            sql &= "Order by ITMC5K, CLRC5K, RFCC5K, RCMC5K "
            Dim dt_S5K00A As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_S5K00A.Rows.Count - 1
                '
                sql = "SELECT Unique_ID, Cat, YCode, PCode FROM M_PBRItemConvert "
                sql &= "Where Cat = 'RCA' "
                sql &= "And YCode = '" & dt_S5K00A.Rows(i).Item("ITMC5K") & "' "
                Dim dt_RCAITEM As DataTable = oDataBase.GetDataTable(sql)
                If dt_RCAITEM.Rows.Count > 0 Then
                    '
                    '不變更BUYER
                    sql = "SELECT Unique_ID, Buyer FROM M_T8T10Buyer "
                    sql &= "Where Cat = '" & dt_RCAITEM.Rows(0).Item("Cat") & "' "
                    sql &= "And   Buyer = '" & Mid(dt_S5K00A.Rows(i).Item("RCMC5K"), 1, 6) & "' "
                    Dim dt_RCABUYER As DataTable = oDataBase.GetDataTable(sql)
                    If dt_RCABUYER.Rows.Count > 0 Then
                        '
                        '不變更BUYER中 - 除外CUST
                        sql = "SELECT Unique_ID, Customer FROM M_ExcludeCust "
                        sql &= "Where Cat = '" & dt_RCAITEM.Rows(0).Item("Cat") & "' "
                        sql &= "And   Buyer = '" & Mid(dt_S5K00A.Rows(i).Item("RCMC5K"), 1, 6) & "' "
                        sql &= "And   Customer = '" & dt_S5K00A.Rows(i).Item("RFCC5K") & "' "
                        Dim dt_RCACUST As DataTable = oDataBase.GetDataTable(sql)
                        If dt_RCACUST.Rows.Count > 0 Then
                            '變更 RCF TO RCA (除外CUST)
                            'UPDATE S5K00 ITMC5K, CLRC5K
                            sql = "Update S5K00 Set "
                            sql = sql + " ITMC5K = '" & dt_RCAITEM.Rows(0).Item("PCode") & "', "
                            sql = sql + " CLRC5K = '" & "" & "', "
                            sql = sql + " PRGC5K = ITMC5K "
                            sql = sql + " Where Buyer = '" & pBuyer & "' "
                            sql = sql + " And ITMC5K = '" & dt_S5K00A.Rows(i).Item("ITMC5K") & "'"
                            sql = sql + " And CLRC5K = '" & dt_S5K00A.Rows(i).Item("CLRC5K") & "'"
                            sql = sql + " And RFCC5K = '" & dt_S5K00A.Rows(i).Item("RFCC5K") & "'"
                            sql = sql + " And RCMC5K = '" & dt_S5K00A.Rows(i).Item("RCMC5K") & "'"
                            oDataBase.ExecuteNonQuery(sql)
                        End If
                    Else
                        '變更 RCF TO RCA (變更BUYER)
                        'UPDATE S5K00 ITMC5K, CLRC5K
                        sql = "Update S5K00 Set "
                        sql = sql + " ITMC5K = '" & dt_RCAITEM.Rows(0).Item("PCode") & "', "
                        sql = sql + " CLRC5K = '" & "" & "', "
                        sql = sql + " PRGC5K = ITMC5K "
                        sql = sql + " Where Buyer = '" & pBuyer & "' "
                        sql = sql + " And ITMC5K = '" & dt_S5K00A.Rows(i).Item("ITMC5K") & "'"
                        sql = sql + " And CLRC5K = '" & dt_S5K00A.Rows(i).Item("CLRC5K") & "'"
                        sql = sql + " And RFCC5K = '" & dt_S5K00A.Rows(i).Item("RFCC5K") & "'"
                        sql = sql + " And RCMC5K = '" & dt_S5K00A.Rows(i).Item("RCMC5K") & "'"
                        oDataBase.ExecuteNonQuery(sql)
                    End If
                End If
            Next
            '* ADD-END
            '
            If RtnCode = 0 Then
                oWavesEDI.Timeout = 900 * 1000
                RtnCode = oWavesEDI.EDICheckItemCode(pLogID, pBuyer, pUserID, pGRBuyer)
            End If
            '* MODIFY-END  
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "CheckItemCode", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'CheckItemCode-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([190]-CheckDuplicateData)
    '**     檢測Duplicate EDI-Data(S5K00)
    '***********************************************************************************************
    'CheckDuplicateData-Start
    <WebMethod()> _
    Public Function CheckDuplicateData(ByVal pLogID As String, _
                                       ByVal pBuyer As String, _
                                       ByVal pUserID As String, _
                                       ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql, xPODN, xPOSN, msg As String
        Dim i As Integer
        '
        Try
            oWavesEDI.Timeout = 900 * 1000
            RtnCode = oWavesEDI.EDICheckDuplicateData(pLogID, pBuyer, pUserID, pGRBuyer)
            '
            ' Joy-Add 2017/6/29
            If RtnCode = 0 Then
                sql = "SELECT PODN5K, POSN5K FROM S5K00 "
                Sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "Group by PODN5K, POSN5K "
                sql &= "Order by PODN5K, POSN5K "
                Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(Sql)
                For i = 0 To dt_S5K00.Rows.Count - 1
                    xPODN = dt_S5K00.Rows(i).Item("PODN5K").ToString
                    xPOSN = dt_S5K00.Rows(i).Item("POSN5K").ToString

                    sql = "SELECT PODN5K, POSN5K FROM S5K00 "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "  And PODN5K = '" & xPODN & "' "
                    sql &= "  And POSN5K = '" & xPOSN & "' "
                    Dim dt_S5K00x As DataTable = oDataBase.GetDataTable(sql)
                    If dt_S5K00x.Rows.Count <> 1 Then
                        RtnCode = 1
                        msg = "Check-DuplicateData[Data Duplication (PODN5K=" + xPODN + " / POSN5K=" + xPOSN + ") !"
                        oDB.AccessLog(pLogID, pBuyer, "Check-DuplicateData", NowDateTime, 1, "Failed", msg, "=", CStr(RtnCode), "", pUserID, "")
                        Exit For
                    End If
                Next
            End If
            ' -----------------
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "CheckDuplicateData", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'CheckDuplicateData-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([200]-CheckNikeVDP)
    '**     檢測NIKE VDP
    '***********************************************************************************************
    'CheckNikeVDP-Start
    <WebMethod()> _
    Public Function CheckNikeVDP(ByVal pLogID As String, _
                                 ByVal pBuyer As String, _
                                 ByVal pUserID As String, _
                                 ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim i, Row As Integer
        Dim err As Boolean = False
        Dim msg As String = ""
        '
        Try
            ' SC760W1-檢測(限定NIKE)
            Row = 1
            sql = "SELECT * FROM SC760W1 "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "Order by Unique_ID "
            Dim dt_SC760W1 As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_SC760W1.Rows.Count - 1
                ' [SNCDW1 檢測]
                If Not err Then
                    If dt_SC760W1.Rows(i).Item("SNCDW1").ToString = "" Or dt_SC760W1.Rows(i).Item("SNCDW1").ToString = "0" Then
                        err = True
                        msg = "Check-NIKEVDP:[SNCDW1=Blank !]"
                    End If
                End If
                ' [SNYRW1 檢測]
                If Not err Then
                    If dt_SC760W1.Rows(i).Item("SNYRW1").ToString = "" Or dt_SC760W1.Rows(i).Item("SNYRW1").ToString = "0" Then
                        err = True
                        msg = "Check-NIKEVDP:[SNYRW1=Blank !]"
                    End If
                End If
                ' [BYMHW1 檢測]
                If Not err Then
                    If dt_SC760W1.Rows(i).Item("BYMHW1").ToString = "" Or dt_SC760W1.Rows(i).Item("BYMHW1").ToString = "0" Then
                        err = True
                        msg = "Check-NIKEVDP:[BYMHW1=Blank !]"
                    End If
                End If
                ' [NFCDW1 檢測]
                If Not err Then
                    If dt_SC760W1.Rows(i).Item("NFCDW1").ToString = "" Or dt_SC760W1.Rows(i).Item("NFCDW1").ToString = "0" Then
                        err = True
                        msg = "Check-NIKEVDP:[NFCDW1=Blank !]"
                    End If
                End If
                ' [NSNOW1 檢測]
                If Not err Then
                    If dt_SC760W1.Rows(i).Item("NSNOW1").ToString = "" Or dt_SC760W1.Rows(i).Item("NSNOW1").ToString = "0" Then
                        err = True
                        msg = "Check-NIKEVDP:[NSNOW1=Blank !]"
                    End If
                End If
                ' [NINOW1 檢測]
                If Not err Then
                    If dt_SC760W1.Rows(i).Item("NINOW1").ToString = "" Or dt_SC760W1.Rows(i).Item("NINOW1").ToString = "0" Then
                        err = True
                        msg = "Check-NIKEVDP:[NINOW1=Blank !]"
                    End If
                End If
                ' [NCCDW1 檢測]
                If Not err Then
                    If dt_SC760W1.Rows(i).Item("NCCDW1").ToString = "" Or dt_SC760W1.Rows(i).Item("NCCDW1").ToString = "0" Then
                        err = True
                        msg = "Check-NIKEVDP:[NCCDW1=Blank !]"
                    End If
                End If
                ' [檢測異常處理]
                If err Then
                    RtnCode = 1
                    oDB.AccessLog(pLogID, pBuyer, "CheckNikeVDP", NowDateTime, 1, CStr(Row), msg, "=", CStr(RtnCode), "", pUserID, "")
                    Exit For
                Else
                    Row = Row + 1
                End If
            Next
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "CheckNikeVDP", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'CheckNikeVDP-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([210]-CheckPOStructure)
    '**     檢測客戶PO結構資料
    '***********************************************************************************************
    'CheckPOStructure-Start
    <WebMethod()> _
    Public Function CheckPOStructure(ByVal pLogID As String, _
                                     ByVal pBuyer As String, _
                                     ByVal pUserID As String, _
                                     ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim i, Row As Integer
        Dim err As Boolean = False
        Dim msg As String = ""
        ' ActionLog-Key欄
        Dim eCORN5K As String = ""
        Dim eGRPC5K As Integer = 0
        '------------------------------------------
        ' 2017/5/23 EOES-多客戶多BUYER
        Dim eBuyer As String = ""
        Dim eGRBuyer As String = ""
        '
        Try
            ' S5K00
            Row = 1
            'JJJ-對應
            'sql = "SELECT * FROM S5K00 "
            sql = "Select *, Case SUBSTRING(RFCC5K, 1, 4) WHEN 'ZJJJ' THEN RFCC5K + '-' + '000999' ELSE RFCC5K + '-' + RCMC5K END  As CustBuyer "
            sql &= "From S5K00 "
            '<<<
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "Order by Unique_ID "
            Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_S5K00.Rows.Count - 1
                ' ActionLog-Key欄
                eCORN5K = dt_S5K00.Rows(i).Item("CORN5K").ToString
                eGRPC5K = dt_S5K00.Rows(i).Item("GRPC5K")

                ' [S5L00 檢測]
                If Not err Then
                    sql = "SELECT * FROM S5L00 "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "  And PODN5L = '" & dt_S5K00.Rows(i).Item("PODN5K") & "' "
                    Dim dt_S5L00 As DataTable = oDataBase.GetDataTable(sql)
                    If dt_S5L00.Rows.Count <= 0 Then
                        err = True
                        msg = "Check-POStructure:[PODN5L(" + dt_S5K00.Rows(i).Item("PODN5K") + ") Not Found !]"
                    End If
                End If
                '
                ' [SC760W1 檢測]
                If Not err Then
                    ' 2017/5/23 EOES-ADD-START
                    ' 取得 BUYERGROUP
                    'JJJ-對應
                    'eBuyer = dt_S5K00.Rows(i).Item("RFCC5K") + "-" + dt_S5K00.Rows(i).Item("RCMC5K")
                    eBuyer = dt_S5K00.Rows(i).Item("CustBuyer")
                    '<<<
                    sql = "Select BuyerGroup From M_ControlRecord "
                    sql = sql & " Where Buyer = '" & eBuyer & "'"
                    Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
                    If dt_ControlRecord.Rows.Count > 0 Then eGRBuyer = dt_ControlRecord.Rows(0).Item("BuyerGroup")
                    '
                    ' 2017/5/23 EOES-MODIFY-START
                    'If oDB.GetFunctionCode(pGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                    If oDB.GetFunctionCode(eGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                        sql = "SELECT * FROM SC760W1 "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "  And ORDNW1 = '" & dt_S5K00.Rows(i).Item("PODN5K") & "' "
                        sql &= "  And OSBNW1 = '" & dt_S5K00.Rows(i).Item("POSN5K").ToString & "' "
                        Dim dt_SC760W1 As DataTable = oDataBase.GetDataTable(sql)
                        If dt_SC760W1.Rows.Count <= 0 Then
                            err = True
                            msg = "Check-POStructure:[ORDNW1(" + dt_S5K00.Rows(i).Item("PODN5K") + ")-" + _
                                                     "ORDNW1(" + dt_S5K00.Rows(i).Item("POSN5K").ToString + ") Not Found !]"
                        End If
                    End If
                End If
                ' ADD-START BY JOY 2017/9/25
                ' [S5K00 PUCN5K	CORN5K 中文檢測]
                Dim j, k As Long
                Dim CT As Boolean = False
                ' PUCN5K
                j = Len(dt_S5K00.Rows(i).Item("PUCN5K").ToString)
                For k = 1 To j
                    If Asc(Mid(dt_S5K00.Rows(i).Item("PUCN5K").ToString, k, 1)) < 0 Then
                        CT = True
                        Exit For
                    End If
                Next
                ' CORN5K
                If CT = False Then
                    j = Len(dt_S5K00.Rows(i).Item("CORN5K").ToString)
                    For k = 1 To j
                        If Asc(Mid(dt_S5K00.Rows(i).Item("CORN5K").ToString, k, 1)) < 0 Then
                            CT = True
                            Exit For
                        End If
                    Next
                End If
                If CT = True Then
                    err = True
                    msg = "Check-POStructure:[PUCN5K(" + dt_S5K00.Rows(i).Item("PUCN5K") + ")-" + _
                                             "CORN5K(" + dt_S5K00.Rows(i).Item("CORN5K").ToString + ") 有中文字元 !]"
                End If
                ' ADD-END
                ' [檢測異常處理]
                If err Then
                    RtnCode = 1
                    oDB.AccessLog(pLogID, pBuyer, "Check-POStructure", NowDateTime, 1, oDB.GetAccessLogLine(pLogID, pBuyer, eCORN5K, eGRPC5K), msg, "=", CStr(RtnCode), "", pUserID, "")
                    Exit For
                Else
                    Row = Row + 1
                End If
            Next
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "Check-POStructure", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'CheckPOStructure-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([220]-EDI2Waves)
    '**     EDI Data To Waves System
    '***********************************************************************************************
    'EDI2Waves-Start
    <WebMethod()> _
    Public Function EDI2Waves(ByVal pLogID As String, _
                              ByVal pBuyer As String, _
                              ByVal pUserID As String, _
                              ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            oWavesEDI.Timeout = 900 * 1000
            RtnCode = oWavesEDI.EDIEDI2Waves(pLogID, pBuyer, pUserID, pGRBuyer)
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "EDI2Waves", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'EDI2Waves-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([230]-EDI2B2B)
    '**     EDI Data To B2B File
    '***********************************************************************************************
    'EDI2B2B-Start
    <WebMethod()> _
    Public Function EDI2B2B(ByVal pLogID As String, _
                            ByVal pBuyer As String, _
                            ByVal pUserID As String, _
                            ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql, xOption As String
        '
        Try
            ' B2B-Option --> GRBuyer(6)
            '[T] --> [A]+[B] / [A] --> PRICE / [B] --> SHIP / [X] --> 不使用
            xOption = oDB.GetFunctionCode(pGRBuyer, 6)
            '
            ' Insert Into B2B Link File
            sql = "Insert Into B_CustomerRequest "
            sql &= "Select "
            sql &= "Buyer, "
            sql &= "A1, B1, C1, D1, E1, F1, G1, H1, I1, J1, K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, U1, V1, W1, X1, Y1, Z1, "
            sql &= "AA1, AB1, AC1, AD1, AE1, AF1, AG1, AH1, AI1, AJ1, AK1, AL1, AM1, AN1, AO1, AP1, AQ1, AR1, AS1, AT1, AU1, AV1, AW1, AX1, AY1, AZ1, "
            sql &= "BA1, BB1, BC1, BD1, BE1, BF1, BG1, BH1, BI1, BJ1, BK1, BL1, BM1, BN1, BO1, BP1, BQ1, BR1, BS1, BT1, BU1, BV1, BW1, BX1, BY1, BZ1, "
            sql &= "CA1, CB1, CC1, CD1, CE1, CF1, CG1, CH1, CI1, CJ1, CK1, CL1, CM1, CN1, CO1, CP1, CQ1, CR1, CS1, CT1, CU1, CV1, CW1, CX1, CY1, CZ1, "
            sql &= "DA1, DB1, DC1, DD1, DE1, DF1, DG1, DH1, DI1, DJ1, DK1, DL1, DM1, DN1, DO1, DP1, DQ1, DR1, DS1, DT1, DU1, DV1, DW1, DX1, DY1, DZ1, "
            sql &= "EA1, EB1, EC1, ED1, EE1, EF1, EG1, EH1, EI1, EJ1, EK1, EL1, EM1, EN1, EO1, EP1, EQ1, ER1, ES1, ET1, EU1, EV1, EW1, EX1, EY1, EZ1, "
            '
            sql &= "'" & xOption & "', "
            sql &= "0, "
            sql &= "0, "
            '
            sql &= "CreateUser, CreateTime, "
            sql &= "'', '' "

            sql &= "From E_InputSheet "
            sql &= "Where Buyer = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "EDI2B2B", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'EDI2B2B-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([240]-PriceList)
    '**     展開 Price List 資料
    '***********************************************************************************************
    'PriceList-Start
    <WebMethod()> _
    Public Function PriceList(ByVal pLogID As String, _
                              ByVal pBuyer As String, _
                              ByVal pUserID As String, _
                              ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql, xItemStatistics, xItemSize, xItemChain, xPriceList, xPriceListRemark, xLastTime, msg As String
        Dim xOrderNo, xItem, xColor, xSalesPrice As String
        Dim xPCode As Integer = 0
        ' ActionLog-Key欄
        Dim eCORN5K As String = ""
        Dim eGRPC5K As Integer = 0
        '
        Try
            '刪除上回資料後再計算
            If DeletePriceListData(pBuyer) = 0 Then
                '
                ' [S5K00]
                sql = "SELECT * FROM S5K00 "
                sql &= "Where Buyer = '" & pBuyer & "' "
                'sql &= "  And CORN5K = '" & "K120009919" & "' "
                sql &= "Order by Unique_ID "
                Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                For i As Integer = 0 To dt_S5K00.Rows.Count - 1
                    '
                    xOrderNo = ""
                    xItem = ""
                    xColor = ""
                    xSalesPrice = "0"
                    '
                    xItemStatistics = ""
                    xItemSize = ""
                    xItemChain = ""
                    xPriceList = ""
                    xPriceListRemark = ""
                    xLastTime = ""
                    ' ActionLog-Key欄
                    eCORN5K = dt_S5K00.Rows(i).Item("CORN5K").ToString
                    eGRPC5K = dt_S5K00.Rows(i).Item("GRPC5K")
                    '
                    ' [標準單價]
                    xPCode = oWaves.GetStandardPrice1("A", _
                                                      dt_S5K00.Rows(i).Item("PPVC5K").ToString, _
                                                      dt_S5K00.Rows(i).Item("TCRC5K").ToString, _
                                                      dt_S5K00.Rows(i).Item("ITMC5K").ToString, _
                                                      xLastTime)
                    If xPCode = 0 Then
                        xPriceList = "標準單價"
                        If xPriceListRemark = "" Then
                            xPriceListRemark = "標準單價(S4100)"
                        Else
                            xPriceListRemark = xPriceListRemark + "→" + "標準單價(S4100)"
                        End If
                    End If
                    '
                    ' [客戶折扣]
                    xPCode = oWaves.GetItemStatistics(dt_S5K00.Rows(i).Item("ITMC5K").ToString, xItemStatistics)
                    xPCode = oWaves.GetItemSize(dt_S5K00.Rows(i).Item("ITMC5K").ToString, xItemSize)
                    xPCode = oWaves.GetItemChain(dt_S5K00.Rows(i).Item("ITMC5K").ToString, xItemChain)
                    xPCode = oWaves.GetPriceDiscount1(dt_S5K00.Rows(i).Item("PPVC5K").ToString, _
                                                      dt_S5K00.Rows(i).Item("RFCC5K").ToString, _
                                                      xItemStatistics, _
                                                      xItemSize, _
                                                      xItemChain, _
                                                      xLastTime)
                    If xPCode = 0 Then
                        xPriceList = "客戶折扣"
                        If xPriceListRemark = "" Then
                            xPriceListRemark = "客戶折扣(S3100)"
                        Else
                            xPriceListRemark = xPriceListRemark + "→" + "客戶折扣(S3100)"
                        End If
                    End If
                    '
                    ' [BuyerPrice]
                    xPCode = oWaves.GetBuyerPrice(Mid(dt_S5K00.Rows(i).Item("RCMC5K").ToString, 1, 6), _
                                                  dt_S5K00.Rows(i).Item("PPVC5K").ToString, _
                                                  dt_S5K00.Rows(i).Item("TCRC5K").ToString, _
                                                  "F", _
                                                  dt_S5K00.Rows(i).Item("ITMC5K").ToString, _
                                                  dt_S5K00.Rows(i).Item("CLRC5K").ToString, _
                                                  xLastTime)
                    If xPCode = 0 Then
                        xPriceList = "BuyerPrice"
                        If xPriceListRemark = "" Then
                            xPriceListRemark = "BuyerPrice(S3L00)"
                        Else
                            xPriceListRemark = xPriceListRemark + "→" + "BuyerPrice(S3L00)"
                        End If
                    End If
                    '
                    ' [CustomerPrice]
                    xPCode = oWaves.GetCustomerPrice(dt_S5K00.Rows(i).Item("RFCC5K").ToString, _
                                                     dt_S5K00.Rows(i).Item("PPVC5K").ToString, _
                                                     dt_S5K00.Rows(i).Item("TCRC5K").ToString, _
                                                     "F", _
                                                     dt_S5K00.Rows(i).Item("ITMC5K").ToString, _
                                                     dt_S5K00.Rows(i).Item("CLRC5K").ToString, _
                                                     xLastTime)
                    If xPCode = 0 Then
                        xPriceList = "CustomerPrice"
                        If xPriceListRemark = "" Then
                            xPriceListRemark = "CustomerPrice(S3000)"
                        Else
                            xPriceListRemark = xPriceListRemark + "→" + "CustomerPrice(S3000)"
                        End If
                    End If
                    '
                    ' [CustomerBuyerPrice]
                    xPCode = oWaves.GetCustomerBuyerPrice(dt_S5K00.Rows(i).Item("RFCC5K").ToString, _
                                                          Mid(dt_S5K00.Rows(i).Item("RCMC5K").ToString, 1, 6), _
                                                          dt_S5K00.Rows(i).Item("PPVC5K").ToString, _
                                                          dt_S5K00.Rows(i).Item("TCRC5K").ToString, _
                                                          "F", _
                                                          dt_S5K00.Rows(i).Item("ITMC5K").ToString, _
                                                          dt_S5K00.Rows(i).Item("CLRC5K").ToString, _
                                                          xLastTime)
                    If xPCode = 0 Then
                        xPriceList = "CustomerBuyerPrice"
                        If xPriceListRemark = "" Then
                            xPriceListRemark = "CustomerBuyerPrice(S3M00)"
                        Else
                            xPriceListRemark = xPriceListRemark + "→" + "CustomerBuyerPrice(S3M00)"
                        End If
                    End If
                    '
                    ' [OrderPrice]
                    xPCode = oWaves.GetOrderPrice(dt_S5K00.Rows(i).Item("RFCC5K").ToString, _
                                                  Mid(dt_S5K00.Rows(i).Item("RCMC5K").ToString, 1, 6), _
                                                  dt_S5K00.Rows(i).Item("CORN5K").ToString, _
                                                  dt_S5K00.Rows(i).Item("POSN5K"), _
                                                  xOrderNo, _
                                                  xItem, _
                                                  xColor, _
                                                  xSalesPrice)
                    '
                    ' [Insert into B_SalesPrice]
                    If xPCode = 0 Then
                        sql = "Insert into W_SalesPrice "
                        sql &= "( "
                        sql &= "CustomerBuyer, PO, Seqno, Customer, Buyer, "
                        sql &= "OrderNo, Item, Color, SalesPrice, "
                        sql &= "PriceList, PriceListRemark, RegisterTime, PriceVersion, "
                        sql &= "CreateTime "
                        sql &= " ) "
                        sql &= "VALUES( "
                        sql &= " '" & pBuyer & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("CORN5K").ToString & "', "
                        sql &= " " & CStr(CInt(dt_S5K00.Rows(i).Item("POSN5K"))) & ", "
                        sql &= " '" & dt_S5K00.Rows(i).Item("RFCC5K").ToString & "', "
                        sql &= " '" & Mid(dt_S5K00.Rows(i).Item("RCMC5K").ToString, 1, 6) & "', "
                        '
                        sql &= " '" & xOrderNo & "', "
                        sql &= " '" & xItem & "', "
                        sql &= " '" & xColor & "', "
                        sql &= " " & CStr(CDbl(xSalesPrice) / 1000000) & ", "
                        '
                        sql &= " '" & xPriceList & "', "
                        sql &= " '" & xPriceListRemark & "', "
                        sql &= " '" & xLastTime & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PPVC5K").ToString & "', "
                        '
                        sql &= " '" & NowDateTime & "' "
                        sql &= " ) "
                        oDataBase.ExecuteNonQuery(sql)
                    Else
                        RtnCode = xPCode
                        msg = "[" + dt_S5K00.Rows(i).Item("RFCC5K").ToString + "]" + _
                              "[" + Mid(dt_S5K00.Rows(i).Item("RCMC5K").ToString, 1, 6) + "]" + _
                              "[" + dt_S5K00.Rows(i).Item("CORN5K").ToString + "]" + _
                              "[" + dt_S5K00.Rows(i).Item("POSN5K").ToString + "]"
                        oDB.AccessLog(pLogID, pBuyer, "GetSalesPrice", NowDateTime, 1, oDB.GetAccessLogLine(pLogID, pBuyer, eCORN5K, eGRPC5K), msg, "=", CStr(RtnCode), "", pUserID, "")
                    End If
                Next
            End If
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "GetSalesPrice", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'PriceList-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([250]-UpdateOrderProgress)
    '**     更新 Order Progress 資料
    '***********************************************************************************************
    'UpdateOrderProgress-Start
    <WebMethod()> _
    Public Function UpdateOrderProgress(ByVal pLogID As String, _
                                        ByVal pBuyer As String, _
                                        ByVal pUserID As String, _
                                        ByVal pPO As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i As Integer
        Dim sql As String
        '
        Dim xPriceVersion, xOrderQty, xListPrice, xSalesPrice, xSalesAmount, xStatus, xStsDesc, xPackqty, xDeliveryQty, xPlanDate, xInvoiceQty, xCompletedDate As String
        xPriceVersion = ""
        xOrderQty = ""
        xListPrice = ""
        xSalesPrice = ""
        xSalesAmount = ""
        xStatus = ""
        xStsDesc = ""
        xPackqty = ""
        xDeliveryQty = ""
        xPlanDate = ""
        xInvoiceQty = ""
        xCompletedDate = ""
        '
        Try
            sql = "SELECT * From B_OrderProgress "
            sql &= "Where PO = '" & pPO & "' "
            sql &= "Order by PO, SeqNo "
            Dim dt_OrderProgress As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_OrderProgress.Rows.Count - 1
                '
                'Get Order Price
                If RtnCode = 0 Then
                    RtnCode = oWaves.GetOrderPriceDetail(dt_OrderProgress.Rows(i).Item("OrderNo").ToString, _
                                                         dt_OrderProgress.Rows(i).Item("OrderSubNo"), _
                                                         xPriceVersion, _
                                                         xOrderQty, _
                                                         xListPrice, _
                                                         xSalesPrice, _
                                                         xSalesAmount)
                End If
                '
                'Get Order Progress
                If RtnCode = 0 Then
                    RtnCode = oWaves.GetOrderProgress(dt_OrderProgress.Rows(i).Item("OrderNo").ToString, _
                                                      dt_OrderProgress.Rows(i).Item("OrderSubNo"), _
                                                      xStatus, _
                                                      xPackqty, _
                                                      xDeliveryQty, _
                                                      xPlanDate, _
                                                      xInvoiceQty, _
                                                      xCompletedDate)

                    xPlanDate = CStr(CDate(Mid(xPlanDate, 1, 4) + "/" + Mid(xPlanDate, 5, 2) + "/" + Mid(xPlanDate, 7, 2)))

                    If xCompletedDate = "0" Then
                        xCompletedDate = "1900/1/1"
                    Else
                        xCompletedDate = CStr(CDate(Mid(xCompletedDate, 1, 4) + "/" + Mid(xCompletedDate, 5, 2) + "/" + Mid(xCompletedDate, 7, 2)))
                    End If
                End If
                '
                'Get Status說明
                If RtnCode = 0 Then
                    oWaves.GetDescriptionMaster("PSTC", xStatus, xStsDesc)
                End If
                '
                'Update B_OrderProgress
                If RtnCode = 0 Then
                    sql = "Update B_OrderProgress Set "
                    sql &= " PriceVersion = '" & xPriceVersion & "', "
                    sql &= " ListPrice = " & CStr(CDbl(xListPrice) / 1000000) & ", "
                    sql &= " SalesPrice = " & CStr(CDbl(xSalesPrice) / 1000000) & ", "
                    sql &= " SalesAmount = " & CStr(CDbl(xSalesAmount) / 100) & ", "
                    sql &= " OrderQty = " & CStr(CDbl(xOrderQty) / 10000000) & ", "
                    sql &= " Status = '" & xStsDesc & "', "
                    sql &= " PackQty = " & CStr(CDbl(xPackqty) / 10000000) & ", "
                    sql &= " DeliveryQty = " & CStr(CDbl(xDeliveryQty) / 10000000) & ", "
                    sql &= " PlanDate = '" & xPlanDate & "', "
                    sql &= " InvoiceQty = " & CStr(CDbl(xInvoiceQty) / 10000000) & ", "
                    sql &= " CompletedDate = '" & xCompletedDate & "', "
                    sql &= " CreateTime = '" & NowDateTime & "' "
                    sql &= " Where PO = '" & dt_OrderProgress.Rows(i).Item("PO").ToString & "' "
                    sql &= "   And SeqNo = " & dt_OrderProgress.Rows(i).Item("SeqNo").ToString & " "
                    oDataBase.ExecuteNonQuery(sql)
                End If
            Next
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "UpdateOrderProgress", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'UpdateOrderProgress-End


    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([260]-UpdatePackingList)
    '**     更新 Packing List 資料
    '***********************************************************************************************
    'UpdatePackingList-Start
    <WebMethod()> _
    Public Function UpdatePackingList(ByVal pLogID As String, _
                                      ByVal pBuyer As String, _
                                      ByVal pOrder As String, _
                                      ByVal pUserID As String) As Integer

        Dim RtnCode As Integer = 0
        '
        Try
            oWavesEDI.Timeout = 900 * 1000
            RtnCode = oWaves.GetPackingListInf(pLogID, pBuyer, pOrder, pUserID)
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "UpdatePackingList", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'UpdatePackingList-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([500]-UpdateControlRecord)
    '**     更新客戶控制檔-Action Flag
    '**     pAction=動作名稱    pShow=顯示或不顯示
    '***********************************************************************************************
    'UpdateControlRecord-Start
    <WebMethod()> _
    Public Function UpdateControlRecord(ByVal pBuyer As String, _
                                        ByVal pUserID As String, _
                                        ByVal pAction1 As String, _
                                        ByVal pStatus1 As Integer, _
                                        ByVal pAction2 As String, _
                                        ByVal pStatus2 As Integer) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            Dim sql As String
            Select Case pAction1
                Case "Start"
                    sql = "Update M_ControlRecord Set "
                    sql = sql + " Active = '1', "
                    sql = sql + " Customer = '2', "
                    sql = sql + " DataPrepare = '1', "
                    sql = sql + " DataImport = '0', "
                    sql = sql + " CreatePONO = '0', "
                    sql = sql + " OrderNo = '0', "
                    sql = sql + " GRPC = '0', "
                    sql = sql + " Company = '0', "
                    sql = sql + " Color = '0', "
                    sql = sql + " Item = '0', "
                    sql = sql + " Duplication = '0', "
                    sql = sql + " DataBalance = '0', "
                    sql = sql + " ToWaves = '0', "
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where Buyer = '" & pBuyer & "' "
                    oDataBase.ExecuteNonQuery(sql)
                Case "End"
                    sql = "Update M_ControlRecord Set "
                    sql = sql + " Active = '0', "
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where Buyer = '" & pBuyer & "' "
                    oDataBase.ExecuteNonQuery(sql)
                Case "Reset"
                    sql = "Update M_ControlRecord Set "
                    sql = sql + " Active = '0', "
                    sql = sql + " Customer = '0', "
                    sql = sql + " DataPrepare = '0', "
                    sql = sql + " DataImport = '0', "
                    sql = sql + " CreatePONO = '0', "
                    sql = sql + " OrderNo = '0', "
                    sql = sql + " GRPC = '0', "
                    sql = sql + " Company = '0', "
                    sql = sql + " Color = '0', "
                    sql = sql + " Item = '0', "
                    sql = sql + " Duplication = '0', "
                    sql = sql + " DataBalance = '0', "
                    sql = sql + " ToWaves = '0', "
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where Buyer = '" & pBuyer & "' "
                    oDataBase.ExecuteNonQuery(sql)
                Case "Action2"
                    sql = "Update M_ControlRecord Set "
                    sql = sql + pAction2 + " = '" & CStr(pStatus2) & "', "
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where Buyer = '" & pBuyer & "' "
                    oDataBase.ExecuteNonQuery(sql)
                Case Else
                    sql = "Update M_ControlRecord Set "
                    sql = sql + pAction1 + " = '" & CStr(pStatus1) & "', "
                    sql = sql + pAction2 + " = '" & CStr(pStatus2) & "', "
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where Buyer = '" & pBuyer & "' "
                    oDataBase.ExecuteNonQuery(sql)
            End Select
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'UpdateControlRecord-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([510]-CheckControlRecord)
    '**     檢查客戶控制檔-Action Flag--判斷是否完成作業
    '***********************************************************************************************
    'CheckControlRecord-Start
    <WebMethod()> _
    Public Function CheckControlRecord(ByVal pBuyer As String, ByVal pAction As String, ByVal pValue As Integer) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Select * From M_ControlRecord "
            sql = sql & " Where Buyer = '" & pBuyer & "'"
            Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
            If dt_ControlRecord.Rows.Count > 0 Then
                If dt_ControlRecord.Rows(0).Item(pAction) <> pValue Then
                    RtnCode = 1
                End If
            Else
                RtnCode = 1
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'CheckControlRecord-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([520]-DeleteHistory)
    '**     刪除執行履歷
    '***********************************************************************************************
    'DeleteHistory-Start
    <WebMethod()> _
    Public Function DeleteHistory(ByVal pBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Delete From L_ActionHistory "
            sql = sql + "Where Buyer = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'DeleteHistory-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([530]-DeleteActionHistory)
    '**     刪除執行履歷(指定Action)
    '***********************************************************************************************
    'DeleteActionHistory-Start
    <WebMethod()> _
    Public Function DeleteActionHistory(ByVal pLogID As String, ByVal pBuyer As String, ByVal pAction As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Delete From L_ActionHistory "
            sql = sql + "Where LogID = '" & pLogID & "' "
            sql = sql + "  And Buyer = '" & pBuyer & "' "
            sql = sql + "  And Action = '" & pAction & "' "
            oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'DeleteActionHistory-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([540]-DeleteS5K00Data)
    '**     刪除S5K00資料
    '***********************************************************************************************
    'DeleteS5K00Data-Start
    <WebMethod()> _
    Public Function DeleteS5K00Data(ByVal pBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Delete From S5K00 "
            sql = sql + "Where Buyer = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'DeleteS5K00Data-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([550]-DeleteS5L00Data)
    '**     刪除S5L00資料
    '***********************************************************************************************
    'DeleteS5L00Data-Start
    <WebMethod()> _
    Public Function DeleteS5L00Data(ByVal pBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Delete From S5L00 "
            sql = sql + "Where Buyer = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'DeleteS5L00Data-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([560]-DeleteSC760W1Data)
    '**     刪除S5L00資料
    '***********************************************************************************************
    'DeleteSC760W1Data-Start
    <WebMethod()> _
    Public Function DeleteSC760W1Data(ByVal pBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Delete From SC760W1 "
            sql = sql + "Where Buyer = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'DeleteSC760W1Data-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([561]-DeletePriceList)
    '**     刪除PriceList資料
    '***********************************************************************************************
    'DeletePriceListData-Start
    <WebMethod()> _
    Public Function DeletePriceListData(ByVal pBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Delete From W_SalesPrice "
            sql = sql + "Where CustomerBuyer = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'DeletePriceListData-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([570]-GetPOSeqNo)
    '**     取得PO SeqNo
    '***********************************************************************************************
    'GetPOSeqNo-Start
    <WebMethod()> _
    Public Function GetPOSeqNo(ByVal pBuyer As String, ByVal pPONo As String) As Integer
        Dim RtnCode As Integer = 1
        '
        Dim sql As String
        Try
            sql = "Select NextPONO From M_ControlRecord "
            sql = sql + "Where Buyer = '" & pBuyer & "' "
            sql = sql + "  And NextPONO Like '" & pPONo & "%' "
            Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
            If dt_ControlRecord.Rows.Count > 0 Then
                RtnCode = CInt(Mid(dt_ControlRecord.Rows(0).Item("NextPONO"), 8, 3))
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'GetPOSeqNo-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([580]-UpdateNextPONO)
    '**     更新下一次可使用PONO
    '***********************************************************************************************
    'UpdateNextPONO-Start
    <WebMethod()> _
    Public Function UpdateNextPONO(ByVal pBuyer As String, ByVal pPONo As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            '因 [BUYER PRICE] 判別用之 [PO-CODE] 不足, 所以允許不同客戶共同使用 [PO-CODE]
            '但相同 [PO-NO] 時轉 WAVE'S ORDER 實會合併至別客戶 ORDER下細項
            '所以本來 依客戶更新[PO-NO] 改成 依共同[PO-CODE]更新[PO-NO]
            sql = "Update M_ControlRecord Set "
            sql = sql + "NextPONO = '" & pPONo & "', "
            sql = sql + "POModifyTime = '" & NowDateTime & "' "
            sql = sql + "Where Buyer = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)

            sql = "Update M_ControlRecord Set "
            sql = sql + "NextPONO = '" & pPONo & "' "
            sql = sql + "Where POCODE = '" & Mid(pPONo, 1, 2) & "' "
            oDataBase.ExecuteNonQuery(sql)
            '-----------------------------------------------
            'sql = "Update M_ControlRecord Set "
            'sql = sql + "NextPONO = '" & pPONo & "' "
            'sql = sql + "Where Buyer = '" & pBuyer & "' "
            'oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'UpdateNextPONO-End

End Class
