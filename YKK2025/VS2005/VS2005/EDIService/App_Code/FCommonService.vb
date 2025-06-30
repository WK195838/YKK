Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Threading
Imports System.Data.OleDb
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
Public Class FCommonService
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
    Dim uCommon As New Utility.Common
    Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")

    '	+-----------------------------------------------------------------------------------------------------------------
    '[000]-Service				    				                                |  
    '	+									                                        |
    '	+----[100]-MakeForcastNo      						                        | 建置 Forcast No
    '	+									                                        |
    '	+----[110]-ForcastPlan      						                        | Forcast Plan展開 (Step-1)
    '	+									                                        |
    '	+----[112]-NewForcastPlan      						                        | Forcast Plan展開 (Step-2)
    '	+									                                        |
    '	+----[113]-VendorForcastPlan      						                    | Forcast Plan展開 (Vendor專用)
    '	+									                                        |
    '	+----[114]-TPForcastPlan      						                        | Forcast Plan展開 (YKK扣具專用)
    '	+									                                        |
    '	+----[115]-MatForcastPlan      						                        | Forcast Plan展開 (材料專用)
    '	+									                                        |
    '	+----[120]-LocalStockPlan      						                        | Local Stock Plan展開 (Step-1)
    '	+									                                        |
    '	+----[122]-NewLocalStockPlan      						                    | Local Stock Plan展開 (Step-2)
    '	+									                                        |
    '	+----[125]-LocalStockPlanProdInf      					                    | Local Stock Plan-Prod Inf.
    '	+									                                        |
    '	+----[130]-BuyerLocalStockPlan      						                | Buyer Local Stock Plan展開
    '	+									                                        |
    '	+----[140]-LastFinalLocalStockPlan    						                | Buyer Local Stock Plan最終確定
    '	+									                                        |
    '	+----[150]-BYFCTDataCheck    						                        | Buyer FCT Data Check
    '	+									                                        |
    '	+----[160]-NewBYFCTImport 						                            | New Buyer FCT Data Import
    '	+									                                        |
    '	+----[170]-ConvertToBYFCT 						                            | 轉換成FAS使用FCT DATA
    '	+									                                        |
    '	+----[180]-EDITransfer   						                            | 轉換EDI-DATA
    '	+									                                        |
    '	+-----------------------------------------------------------------------------------------------------------------
    '	+									                                        |
    '	+----[500]-UpdateControlRecord      	                                    | 更新客戶控制檔-Action Flag
    '	+									                                        |
    '	+----[510]-CheckControlRecord									            | 檢查客戶控制檔-Action Flag
    '	+									                                        |
    '	+----[520]-DeleteHistory        									        | 刪除執行履歷
    '	+									                                        |
    '	+----[530]-DeleteActionHistory 									            | 刪除執行履歷(指定Action)
    '	+									                                        |
    '	+----[540]-DeleteFCTData 	    								            | 刪除 FCT Plan 資料
    '	+									                                        |
    '	+----[541]-DeleteFCTLevelData   								            | 刪除 FCT Plan 資料(對象Level<>0)
    '	+								                                            |
    '	+----[550]-DeleteLSData 			    						            | 刪除 LS Plan 資料
    '	+									                                        |
    '	+----[551]-DeleteBuyerLSData 		    						            | 刪除 Buyer LS Plan 資料
    '	+									                                        |
    '	+----[552]-DeleteLSProdInfData 		    						            | 刪除 LS Plan Prod Inf 資料
    '	+								                                            |
    '	+----[555]-DeleteLS2EDIInterface	    						            | 刪除 EDI Interface 資料
    '	+									                                        |
    '	+----[560]-GetLSSeqNo        									            | 取得LS SeqNo
    '	+									                                        |
    '	+----[561]-GetBuyerLSSeqNo        									        | 取得Buyer LS SeqNo
    '	+									                                        |
    '	+----[570]-GetFCTSeqNo        									            | 取得FCT SeqNo
    '	+									                                        |
    '	+----[575]-GetLastPlanVersion  									            | 取得最後Final Plan Version
    '	+									                                        |
    '	+----[580]-UpdateNextFCTNo								                    | 更新下一次可使用FCTNo
    '	+									                                        |
    '	+----[590]-UpdateNextLSNo								                    | 更新下一次可使用LSNo
    '	+									                                        |
    '	+----[591]-UpdateNextBuyerLSNo								                | 更新下一次可使用BuyerLSNo
    '	+									                                        |
    '	+----[600]-WriteFCTPlan								                        | 寫入Forcast Plan Data
    '	+									                                        |
    '	+----[610]-CanRunLFLocalStockPlan    						                | 判斷是否可執行 LS Plan 最終確定
    '	+									                                        |
    '	+----[620]-GetLastUniqueID    						                        | 取得最後1筆資料
    '	+									                                        |
    '	+----[630]-ResetFCTNo								                        | 重新設定下一次可使用FCTNo
    '	+									                                        |
    '	+----[631]-ResetPlanFCTNo   						                        | 重新設定 FCT Plan FCTNo
    '	+									                                        |
    '	+----[640]-ResetLSNo								                        | 重新設定下一次可使用LSNo
    '	+									                                        |
    '	+----[650]-ResetBuyerLSNo							                        | 重新設定下一次可使用Buyer LSNo
    '	+									                                        |
    '	+----[660]-UpdateBYFCTStatus        					                    | BY FCT 資料同步化(ADIDAS REF + WORKING + ARTICLE)
    '	+									                                        |
    '	+----[670]-GetNativeVendorCode    									        | 取得Native Vendor Code
    '	+									                                        |
    '	+----[680]-GetColor             									        | 取得Color Master String
    '	+									                                        |
    '	+----[690]-CheckItemNoDisplay							                    | 檢測Item NODISPLAY
    '	+									                                        |
    '	+----[695]-CheckColorNotFound     							                | 檢測Color Code  Not Found
    '	+									                                        |
    '	+----[700]-GetProdQtyUnit     							                    | 取得LS中PROD的數量單位
    '	+									                                        |
    '	+----[705]-GetZipper        							                    | 判斷是否為ZIPPER
    '	+									                                        |
    '	+----[710]-GetFilterRuleNo    							                    | 尋找適合備料基準-從過濾箱中(M_LSOrderRule)取得適合備料基準
    '	+									                                        |
    '	+----[800]-MaterialExpansion   							                    | 材料分析BOM展開 
    '	+						                                                    |
    '	+----[850]-CanRunVendorFCFinal  [VENDOR FC]					                | 判斷是否可執行Vendor FC最終確定
    '	+						                                                    |
    '	+----[855]-VendorFCFinal        [VENDOR FC]			                        | Vendor FC 最終確定
    '	+									                                        |
    '	+-----------------------------------------------------------------------------------------------------------------
    '	+									                                        |
    '	+----[900]-KPIExpansion         [eKPI]			                            | eKPI Delay Reason 判定 
    '	+									                                        |
    '	+-----------------------------------------------------------------------------------------------------------------
    '	+ SPS			    				                                        |
    '	+									                                        |
    '	+----[905]-DeleteActionData           	                                    | 刪除 Action Plan 資料
    '	+									                                        |
    '	+----[910]-SPUpdateControlRecord      	                                    | 更新客戶控制檔-Action Flag
    '	+									                                        |
    '	+----[911]-SPNo2Import                	                                    | 轉換SP-NO
    '	+									                                        |
    '	+----[912]-SPCheckControlRecord									            | 檢查客戶控制檔-Action Flag
    '	+									                                        |
    '	+----[913]-SPMakeForcastNo      					                        | 建置 Forcast No
    '	+									                                        |
    '	+----[920]-SPForcastPlan      						                        | SHOPPING Forcast Plan展開
    '	+									                                        |
    '	+----[925]-SPWriteNewFCTPlan      					                        | Write Forcast Plan
    '	+									                                        |
    '	+----[930]-SPLocalStockPlan      						                    | SHOPPING Local Stock Plan展開
    '	+									                                        |
    '	+----[935]-SPActionkPlan          						                    | SHOPPING Action Plan展開
    '	+									                                        |
    '	+----[940]-SPKPInterface          						                    | SHOPPING KP I/F
    '	+									                                        |
    '	+----[945]-SPFinalPlan          						                    | SHOPPING Plan Final
    '	+									                                        |
    '	+----[950]-SPPendingPlan          						                    | SHOPPING Plan Pending
    '	+									                                        |
    '	+----[955]-SPAutoFinalPlan      						                    | SHOPPING Plan Pending Auto Final
    '	+									                                        |
    '	+----									                                    |
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    ' SPS
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([905]-DeleteActionData)
    '**     刪除 Action Plan 資料
    '***********************************************************************************************
    'DeleteActionData-Start
    <WebMethod()> _
    Public Function DeleteActionData(ByVal pBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Delete From SPActionPlan "
            sql = sql + "Where Customer = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'DeleteActionData-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([910]-SPUpdateControlRecord)
    '**     更新客戶控制檔-Action Flag
    '**     pAction=動作名稱    pShow=顯示或不顯示
    '***********************************************************************************************
    'SPUpdateControlRecord-Start
    <WebMethod()> _
    Public Function SPUpdateControlRecord(ByVal pBuyer As String, _
                                          ByVal pUserID As String, _
                                          ByVal pAction1 As String, _
                                          ByVal pStatus1 As Integer, _
                                          ByVal pAction2 As String, _
                                          ByVal pStatus2 As Integer) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            Dim sql As String
            '
            'Update-Before
            sql = "Insert into L_SPControlRecord "
            sql &= "Select "
            sql &= " '" & pUserID & "','" & NowDateTime & "', 'BEF', "
            sql &= "Active, Code, Name, Customer, Import, Demand, ActPlan, ImpActPlan, RspActPlan, "
            sql &= "KPInterface, RspWINGS, PILSheet, Final, ChgFinal, Progress, YOBI1, YOBI4 "
            sql &= " FROM M_SPControlRecord "
            sql &= " Where Code = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
            '
            Select Case pAction1
                Case "Start"
                    sql = "Update M_SPControlRecord Set "
                    sql = sql + " Active = '1', "
                    sql = sql + " Customer = '2', "
                    sql = sql + " Import = '1', "
                    sql = sql + " Demand = '0', "
                    sql = sql + " ActPlan = '0', "
                    sql = sql + " ImpActPlan = '0', "
                    sql = sql + " RspActPlan = '0', "
                    sql = sql + " KPInterface = '0', "
                    sql = sql + " RspWINGS = '0', "
                    sql = sql + " PILSheet = '0', "
                    sql = sql + " Final = '0', "
                    sql = sql + " ChgFinal = '0', "
                    sql = sql + " Progress = '0', "
                    sql = sql + " Yobi4 = 'Start', "
                    '
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where Code = '" & pBuyer & "' "
                    oDataBase.ExecuteNonQuery(sql)
                Case "End"
                    sql = "Update M_SPControlRecord Set "
                    sql = sql + " Active = '0', "
                    sql = sql + " Customer = '0', "
                    sql = sql + " Import = '0', "
                    sql = sql + " Demand = '0', "
                    sql = sql + " ActPlan = '0', "
                    sql = sql + " ImpActPlan = '0', "
                    sql = sql + " RspActPlan = '0', "
                    sql = sql + " KPInterface = '0', "
                    sql = sql + " RspWINGS = '0', "
                    sql = sql + " PILSheet = '0', "
                    sql = sql + " Final = '0', "
                    sql = sql + " ChgFinal = '0', "
                    sql = sql + " Progress = '0', "
                    sql = sql + " Yobi4 = 'End', "
                    '
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where Code = '" & pBuyer & "' "
                    oDataBase.ExecuteNonQuery(sql)
                Case "Reset"
                    sql = "Update M_SPControlRecord Set "
                    sql = sql + " Active = '0', "
                    sql = sql + " Customer = '0', "
                    sql = sql + " Import = '0', "
                    sql = sql + " Demand = '0', "
                    sql = sql + " ActPlan = '0', "
                    sql = sql + " ImpActPlan = '0', "
                    sql = sql + " RspActPlan = '0', "
                    sql = sql + " KPInterface = '0', "
                    sql = sql + " RspWINGS = '0', "
                    sql = sql + " PILSheet = '0', "
                    sql = sql + " Final = '0', "
                    sql = sql + " ChgFinal = '0', "
                    sql = sql + " Progress = '0', "
                    sql = sql + " Yobi4 = 'Reset', "
                    '
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where Code = '" & pBuyer & "' "
                    oDataBase.ExecuteNonQuery(sql)
                Case "Action2"
                    sql = "Update M_SPControlRecord Set "
                    sql = sql + pAction2 + " = '" & CStr(pStatus2) & "', "
                    sql = sql + " Yobi4 = '" & pAction2 & "', "
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where Code = '" & pBuyer & "' "
                    oDataBase.ExecuteNonQuery(sql)
                Case Else
                    sql = "Update M_SPControlRecord Set "
                    sql = sql + pAction1 + " = '" & CStr(pStatus1) & "', "
                    sql = sql + pAction2 + " = '" & CStr(pStatus2) & "', "
                    sql = sql + " Yobi4 = '" & pAction1 & "-->" & pAction2 & "', "
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where Code = '" & pBuyer & "' "
                    oDataBase.ExecuteNonQuery(sql)
            End Select
            '
            'Update-After
            sql = "Insert into L_SPControlRecord "
            sql &= "Select "
            sql &= " '" & pUserID & "','" & NowDateTime & "', 'AFT', "
            sql &= "Active, Code, Name, Customer, Import, Demand, ActPlan, ImpActPlan, RspActPlan, "
            sql &= "KPInterface, RspWINGS, PILSheet, Final, ChgFinal, Progress, YOBI1, YOBI4 "
            sql &= " FROM M_SPControlRecord "
            sql &= " Where Code = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
            '
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'SPUpdateControlRecord-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([911]-SPNo2Import)
    '**       SP-NO置入IMPORT FILE(E_InputSheet) SPNO-->BK1-->Y_J1
    '***********************************************************************************************
    'SPNo2Import-Start
    <WebMethod()> _
    Public Function SPNo2Import(ByVal pLogID As String, _
                                ByVal pBuyer As String, _
                                ByVal pUserID As String, _
                                ByVal pGRBuyer As String, _
                                ByVal pFunList As String, _
                                ByVal pLastUniqueID As Integer) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            Dim sql As String
            '
            ' 更新IMPORT FILE 
            sql = "Update E_InputSheet Set "
            sql = sql + " BK1 = '" & "SP" & pLogID & "' "
            sql = sql + "Where Buyer = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
            '
            ' 更新SP-NO
            sql = "Update M_SPControlRecord Set "
            sql = sql + "SPNo = '" & "SP" & pLogID & "' "
            sql = sql + "Where Code = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
            '
            ' ------------
            'BACKUP PROC (前回DATA)
            Dim Conn As New SqlClient.SqlConnection("Data Source=" & "10.245.0.112" & ";" & _
                                                    "Initial Catalog=" & "EDI" & ";" & _
                                                    "User ID=" & "sa" & ";" & _
                                                    "Password=;")
            Dim cmd As SqlCommand = Conn.CreateCommand
            '
            Conn.Open()
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandTimeout = 180
            cmd.Parameters.Add(New SqlParameter("@Customer", pBuyer))
            cmd.CommandText = "sp_SPPlan2History"
            cmd.ExecuteNonQuery()
            Conn.Close()
            '
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'SPNo2Import-End
    '
    '***********************************************************************************************
    '**([912]]-SPCheckControlRecord)
    '**     檢查客戶控制檔-Action Flag--判斷是否完成作業
    '***********************************************************************************************
    'SPCheckControlRecord-Start
    <WebMethod()> _
    Public Function SPCheckControlRecord(ByVal pBuyer As String, ByVal pAction As String, ByVal pValue As Integer) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Select * From M_SPControlRecord "
            sql = sql & " Where Code = '" & pBuyer & "'"
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
    'SPCheckSPControlRecord-End
    '***********************************************************************************************
    '**([913]-SPMakeForcastNo)
    '**       建置 Forcast No 
    '***********************************************************************************************
    'SPMakeForcastNo-Start
    <WebMethod()> _
    Public Function SPMakeForcastNo(ByVal pLogID As String, _
                                    ByVal pBuyer As String, _
                                    ByVal pUserID As String, _
                                    ByVal pGRBuyer As String, _
                                    ByVal pFunList As String, _
                                    ByVal pLastUniqueID As Integer) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            ' ***********************************************************************************
            ' 變數定義及設定初值
            ' ***********************************************************************************
            Dim sql As String
            Dim xFCTNo, xSeqNoString, xGroup As String
            Dim xSeqNo, i As Integer
            xGroup = "XX"
            ' ***********************************************************************************
            ' 取得客戶FCT FCTNO = GroupCode + "F"  + 年月 + 流水號(6碼)   SHAAD + F  + 247 + 000001
            ' ***********************************************************************************
            ' 取得GroupCode
            sql = "SELECT GrCode FROM M_SPControlRecord "
            sql &= "Where Code = '" & pBuyer & "' "
            Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
            If dt_ControlRecord.Rows.Count > 0 Then xGroup = dt_ControlRecord.Rows(0).Item("GrCode")
            ' 取得年月 
            Select Case Month(Now)
                Case 10
                    xFCTNo = xGroup + "F" + CStr(Year(Now) - 2000) + "A"
                Case 11
                    xFCTNo = xGroup + "F" + CStr(Year(Now) - 2000) + "B"
                Case 12
                    xFCTNo = xGroup + "F" + CStr(Year(Now) - 2000) + "C"
                Case Else
                    xFCTNo = xGroup + "F" + CStr(Year(Now) - 2000) + CStr(Month(Now))
            End Select
            ' 流水號(下一可使用No)
            xSeqNo = 1
            sql = "Select FCTNo From M_SPControlRecord "
            sql = sql + "Where Code = '" & pBuyer & "' "
            sql = sql + "  And FCTNo Like '" & xFCTNo & "%' "
            Dim dt_FCTNo As DataTable = oDataBase.GetDataTable(sql)
            If dt_FCTNo.Rows.Count > 0 Then
                xSeqNo = CInt(Mid(dt_FCTNo.Rows(0).Item("FCTNo"), 11, 5))
            End If
            '
            ' ***********************************************************************************
            ' 更新 Forcast Plan FCTNo
            ' ***********************************************************************************
            sql = "SELECT Unique_ID, FCTNo FROM ForcastPlan "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And Unique_ID > " & pLastUniqueID & " "
            sql &= "  And Y_Level = 0 "
            sql &= "Order by Unique_ID "
            Dim dt_FCT As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_FCT.Rows.Count - 1
                '
                ' 判斷客戶FCT
                If dt_FCT.Rows(i).Item("FCTNo").ToString = "" Then
                    ' 客戶FCT
                    ' 製作SeqNo(不足5位補0)
                    xSeqNoString = CStr(xSeqNo)
                    If Len(xSeqNoString) < 2 Then
                        xSeqNoString = "0000" + CStr(xSeqNo)
                    Else
                        If Len(xSeqNoString) < 3 Then
                            xSeqNoString = "000" + CStr(xSeqNo)
                        Else
                            If Len(xSeqNoString) < 4 Then
                                xSeqNoString = "00" + CStr(xSeqNo)
                            Else
                                If Len(xSeqNoString) < 5 Then
                                    xSeqNoString = "0" + CStr(xSeqNo)
                                End If
                            End If
                        End If
                    End If
                    '
                    ' 更新 Forcast Plan
                    sql = "Update ForcastPlan Set "
                    sql = sql + " FCTNo = '" & xFCTNo + xSeqNoString & "' "
                    sql = sql + " Where Unique_ID = '" & dt_FCT.Rows(i).Item("Unique_ID").ToString & "' "
                    oDataBase.ExecuteNonQuery(sql)
                    '
                    xSeqNo = xSeqNo + 1
                End If
            Next
            '
            ' ***********************************************************************************
            ' 更新下一個可使用No
            ' ***********************************************************************************
            ' 客戶FCT
            ' 製作SeqNo(不足6位補0)
            xSeqNoString = CStr(xSeqNo)
            If Len(xSeqNoString) < 2 Then
                xSeqNoString = "0000" + CStr(xSeqNo)
            Else
                If Len(xSeqNoString) < 3 Then
                    xSeqNoString = "000" + CStr(xSeqNo)
                Else
                    If Len(xSeqNoString) < 4 Then
                        xSeqNoString = "00" + CStr(xSeqNo)
                    Else
                        If Len(xSeqNoString) < 5 Then
                            xSeqNoString = "0" + CStr(xSeqNo)
                        End If
                    End If
                End If
            End If
            ' 更新下一次可使用PONO
            sql = "Update M_SPControlRecord Set "
            sql = sql + "FCTNo = '" & xFCTNo + xSeqNoString & "' "
            sql = sql + "Where Code = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
            '
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "FCTPLAN", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "SPMakeForcastNo", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'SPMakeForcastNo-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([920]-SPForcastPlan)
    '**     SHOPPING Forcast Plan展開 
    '***********************************************************************************************
    'SPForcastPlan-Start
    <WebMethod()> _
    Public Function SPForcastPlan(ByVal pLogID As String, _
                                  ByVal pBuyer As String, _
                                  ByVal pUserID As String, _
                                  ByVal pGRBuyer As String, _
                                  ByVal pFunList As String, _
                                  ByVal pLastUniqueID As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        '
        Try
            '##
            'MsgBox("IN")
            '
            oDB.AccessLog(pLogID, pBuyer, "SPS_ForcastPlan", NowDateTime, 0, "999", "ForcastPlan-Start", "PGM=SPForcastPlan", "", "", pUserID, "")
            '
            ' 變數定義及設定初值
            oWaves.Timeout = Timeout.Infinite
            Dim xItemProduct, xLTType, xPartType, xClass, xQtyMeter, xColor, xBuyerConvert, xBuyerRule As String
            Dim xSubNo As Integer
            Dim FCTWrite As Boolean = False
            ' 構成展開專用變數
            Dim xItem(5), xItemName(5), xQty(5) As String
            Dim xItem1(5), xItemName1(5), xQty1(5) As String
            Dim xCount, xCount1 As Integer
            ' SearchItem專用變數
            Dim xSearchItem(10) As String
            ' 過濾箱中取得的資料變數
            Dim xRuleNo(5), xAction(5), xObjectType(5), xObjectProduct(5) As String
            Dim xRuleCount As Integer
            '
            ' 讀取-ForcastPlan Data
            sql = "SELECT C_Season As Season, Y_ItemCode AS Item, Y_Color AS Color, C_Color AS CSTColor, C_ShortenLT AS KeepCode FROM ForcastPlan "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And Unique_ID > " & pLastUniqueID & " "
            sql &= "  And Y_Level = 0 "
            sql &= "Group by C_Season, Y_ItemCode, Y_Color, C_Color, C_ShortenLT "
            sql &= "Order by C_Season, Y_ItemCode, Y_Color, C_Color, C_ShortenLT "
            Dim dt_FCTItem As DataTable = oDataBase.GetDataTable(sql)
            For i As Integer = 0 To dt_FCTItem.Rows.Count - 1
                FCTWrite = False
                xSubNo = 2
                ' 決定備料使用 BUYER & LT種類(LTType)
                xLTType = "LLT"
                'Get BuyerRule
                sql = "Select * From M_Referp "
                sql = sql & " Where Cat  = '" & "260" & "' "
                sql = sql & "   And DKey = 'MATERIAL-" & pBuyer & "' "
                Dim dt_MATERIAL As DataTable = oDataBase.GetDataTable(sql)
                If dt_MATERIAL.Rows.Count > 0 Then
                    xBuyerRule = Mid(dt_MATERIAL.Rows(0).Item("DATA"), 1, InStr(dt_MATERIAL.Rows(0).Item("DATA"), "/") - 1)
                    xLTType = Mid(dt_MATERIAL.Rows(0).Item("DATA"), InStr(dt_MATERIAL.Rows(0).Item("DATA"), "/") + 1)
                    xLTType = Replace(xLTType, "/", "")
                End If
                'Get BuyerConvert
                sql = "Select * From M_Referp "
                sql = sql & " Where Cat  = '" & "260" & "' "
                sql = sql & "   And DKey = 'CONVERT-" & pBuyer & "' "
                Dim dt_Convert As DataTable = oDataBase.GetDataTable(sql)
                If dt_Convert.Rows.Count > 0 Then
                    xBuyerConvert = dt_Convert.Rows(0).Item("DATA")
                End If
                ' 決定CST PULLER COLOR
                Dim xCSTPullerColor As String = ""
                If InStr(pBuyer, "AD") > 0 Or InStr(xBuyerConvert, "000001") > 0 Then
                    ' ADIDAS
                    xCSTPullerColor = Mid(dt_FCTItem.Rows(i).Item("CSTColor") & "///", InStr(dt_FCTItem.Rows(i).Item("CSTColor"), "/") + 1)
                    xCSTPullerColor = Mid(xCSTPullerColor, InStr(xCSTPullerColor, "/") + 1)
                    xCSTPullerColor = Mid(xCSTPullerColor, 1, 4)
                Else
                    ' Other BUYER
                    xCSTPullerColor = Mid(dt_FCTItem.Rows(i).Item("CSTColor") & "///", InStr(dt_FCTItem.Rows(i).Item("CSTColor"), "/") + 1)
                    xCSTPullerColor = Mid(xCSTPullerColor, InStr(xCSTPullerColor, "/") + 1)
                    xCSTPullerColor = Mid(xCSTPullerColor, 1, InStr(xCSTPullerColor, "/") - 1)
                End If
                'MsgBox("RULE: BUYER=[" & xBuyerConvert & "],LTTYPE=[" & xLTType & "]")
                '
                ' CHAIN 處理 (不需CHECK LINE-LINE ITEM)
                '--------------------------------------------------------------------------------
                ' 取得Meter換算基準 (取得ItemClass)
                'MsgBox("CH")

                oWaves.GetItemClass(dt_FCTItem.Rows(i).Item("Item"), xClass)
                sql = "Select * From M_Referp "
                sql = sql & " Where Cat  = '" & "110" & "' "
                sql = sql & "   And DKey = '" & pBuyer + "-" + xClass & "' "
                Dim dt_Referp As DataTable = oDataBase.GetDataTable(sql)
                If dt_Referp.Rows.Count > 0 Then
                    xQtyMeter = CStr(CDbl(dt_Referp.Rows(0).Item("Data")) * 100)
                Else
                    xQtyMeter = "100"
                End If
                'MsgBox(xQtyMeter)
                ' 取得完成品CHAIN(CF/VF) 或 GAP-CHAIN(MF)     (MF=金屬/CF=樹脂/VF=塑鋼)
                xPartType = "CH"                                                            ' 決定材料種類(PartTye)
                oWaves.GetItemProduct(dt_FCTItem.Rows(i).Item("Item"), xItemProduct)        ' 取得製品區分 
                '
                ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)
                ' 20250106-MODIFY-START
                'If GetZipper(xBuyerConvert, dt_FCTItem.Rows(i).Item("Item")) = 1 And Right(pBuyer, 2) <> "-Y" Then
                ' MODIFY-END
                '
                If GetZipper(xBuyerConvert, dt_FCTItem.Rows(i).Item("Item")) = 1 Then
                    If xItemProduct = "MF" Then
                        oWaves.GetChildItemStructure("01", "CH-GAP", dt_FCTItem.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                    Else
                        oWaves.GetChildItemStructure("01", "CH-DYED", dt_FCTItem.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                    End If
                Else
                    xItem(1) = dt_FCTItem.Rows(i).Item("Item")
                    xItemName(1) = ""
                    xQty(1) = CStr(1 * 10000000)
                    xCount = 1
                    For j As Integer = 2 To 5
                        xItem(j) = ""
                        xItemName(j) = ""
                        xQty(j) = "0"
                    Next
                End If

                ' 展開ITEM構成取得所指定ITEM
                If GetZipper(xBuyerConvert, dt_FCTItem.Rows(i).Item("Item")) <> 3 Then
                    For ItemIndex As Integer = 1 To xCount
                        ' 1. 取得完成品CHAIN兼用色
                        oWaves.GetChangeColor("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCTItem.Rows(i).Item("Color"), xColor)
                        ' 2. 準備備料基準相關SearchItem
                        oWaves.GetSearchItem(xPartType, xItem(ItemIndex), xCSTPullerColor, xSearchItem)
                        ' 3. 尋找適合備料基準-從過濾箱中(M_LSOrderRule)取得適合備料基準
                        GetFilterRuleNo(xBuyerRule, xLTType, xPartType, xSearchItem, xRuleNo, xAction, xObjectType, xObjectProduct, xRuleCount)
                        ' 4. 依備料基準展開結構
                        For RuleIndex As Integer = 1 To xRuleCount
                            FCTWrite = False
                            '
                            'MsgBox(xItem(1) & "--" & xItemName(ItemIndex) & xObjectProduct(RuleIndex))

                            ' 第一階結構資料是否OK
                            If InStr(xItemName(ItemIndex), xObjectProduct(RuleIndex)) > 0 Or xObjectProduct(RuleIndex) = "FINISH" Then
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    SPWriteNewFCTPlan(pBuyer, xBuyerRule, xLTType, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "CH", xQtyMeter, xRuleNo(RuleIndex))
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                            ' 搜尋下一階結構資料
                            If Not FCTWrite Then
                                '
                                oWaves.GetChildItemStructure("01", xObjectType(RuleIndex) + "-" + xObjectProduct(RuleIndex), xItem(ItemIndex), xItem1, xItemName1, xQty1, xCount1)
                                '
                                For ItemIndex1 As Integer = 1 To xCount1
                                    If InStr(xItemName1(ItemIndex1), xObjectProduct(RuleIndex)) > 0 Then
                                        ' 寫入 ForcastPlan
                                        sql = "SELECT * FROM ForcastPlan "
                                        sql &= "Where Buyer = '" & pBuyer & "' "
                                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                        sql &= "  And Y_Level = 0 "
                                        sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                        sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                        sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                        sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                        sql &= "Order by FCTNo, FCTSubNo "
                                        Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                        For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                            SPWriteNewFCTPlan(pBuyer, xBuyerRule, xLTType, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, xItem(ItemIndex), xItem1(ItemIndex1), xColor, "CH", xQtyMeter, xRuleNo(RuleIndex))
                                        Next
                                        xSubNo = xSubNo + 1
                                        FCTWrite = True
                                    End If
                                Next
                            End If
                            ' 搜尋不到任何結構資料(Set Error Rule [9900-990-99])
                            If Not FCTWrite Then
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    SPWriteNewFCTPlan(pBuyer, xBuyerRule, xLTType, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "CH", xQtyMeter, xRuleNo(RuleIndex) + "/" + "9900-990-00")
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                        Next
                    Next
                End If

                '##
                'MsgBox("SLD")
                '
                ' SLIDER 處理  (需CHECK LINE-LINE ITEM)
                '--------------------------------------------------------------------------------
                xQtyMeter = "100"                       ' 取得Meter換算基準 (取得ItemClass)
                xPartType = "SLD"                       ' 決定材料種類(PartTye)
                '
                ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)
                If GetZipper(xBuyerConvert, dt_FCTItem.Rows(i).Item("Item")) = 1 Then
                    oWaves.GetChildItemStructure("01", "SLD-FINISH", dt_FCTItem.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                Else
                    xItem(1) = dt_FCTItem.Rows(i).Item("Item")
                    xItemName(1) = ""
                    xQty(1) = CStr(1 * 10000000)
                    xCount = 1
                    For j As Integer = 2 To 5
                        xItem(j) = ""
                        xItemName(j) = ""
                        xQty(j) = "0"
                    Next
                End If
                ' 展開ITEM構成取得所指定ITEM
                ' 20250106-MODIFY-START
                'If GetZipper(xBuyerConvert, dt_FCTItem.Rows(i).Item("Item")) <> 2 Then
                'If GetZipper(xBuyerConvert, dt_FCTItem.Rows(i).Item("Item")) = 3 Then
                ' MODIFY-END
                '
                If GetZipper(xBuyerConvert, dt_FCTItem.Rows(i).Item("Item")) = 1 Or GetZipper(xBuyerConvert, dt_FCTItem.Rows(i).Item("Item")) = 3 Then
                    '
                    For ItemIndex As Integer = 1 To xCount
                        ' 1. 取得完成品SLD兼用色
                        oWaves.GetChangeColor("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCTItem.Rows(i).Item("Color"), xColor)
                        ' 2. 準備備料基準相關SearchItem

                        '##
                        'MsgBox(xPartType & "-" & xItem(ItemIndex))

                        oWaves.GetSearchItem(xPartType, xItem(ItemIndex), xCSTPullerColor, xSearchItem)
                        '##
                        'MsgBox(xSearchItem(1) & "-" & xSearchItem(2) & "-" & xSearchItem(3) & "-" & xSearchItem(4) & "-" & xSearchItem(5) & "-" & xSearchItem(6) & "-" & xSearchItem(7) & "-" & xSearchItem(8) & "-" & xSearchItem(9) & "-" & xSearchItem(10))

                        ' 3. 尋找適合備料基準-從過濾箱中(M_LSOrderRule)取得適合備料基準
                        GetFilterRuleNo(xBuyerRule, xLTType, xPartType, xSearchItem, xRuleNo, xAction, xObjectType, xObjectProduct, xRuleCount)

                        ' 4. 依備料基準展開結構
                        For RuleIndex As Integer = 1 To xRuleCount
                            '##
                            'MsgBox(xRuleNo(RuleIndex))

                            FCTWrite = False
                            ' 第一階結構資料是否OK
                            If (xObjectProduct(RuleIndex) = "FINISH" Or _
                               (xObjectProduct(RuleIndex) = "E" And InStr(xItemName(ItemIndex), xObjectProduct(RuleIndex)) > 0)) And _
                               oWaves.GetItemLineLine("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex)) = 0 Then         ' 不是LINE-LINE ITEM ((0=不是/1=是)
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    SPWriteNewFCTPlan(pBuyer, xBuyerRule, xLTType, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "PS", xQtyMeter, xRuleNo(RuleIndex))
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                            ' 搜尋下一階結構資料
                            If Not FCTWrite Then
                                '##
                                'MSGBOX("[" & xItem(ItemIndex) & "][" & xObjectType(RuleIndex) + "-" + xObjectProduct(RuleIndex) & "]")

                                oWaves.GetChildItemStructure("01", xObjectType(RuleIndex) + "-" + xObjectProduct(RuleIndex), xItem(ItemIndex), xItem1, xItemName1, xQty1, xCount1)

                                '##
                                'MSGBOX("[" & xItem1(1) & "][" & xItemName1(1) & "]")

                                '
                                For ItemIndex1 As Integer = 1 To xCount1
                                    If oWaves.GetItemLineLine("01", dt_FCTItem.Rows(i).Item("Item"), xItem1(ItemIndex1)) = 0 Then          ' 不是LINE-LINE ITEM ((0=不是/1=是)
                                        ' 寫入 ForcastPlan
                                        sql = "SELECT * FROM ForcastPlan "
                                        sql &= "Where Buyer = '" & pBuyer & "' "
                                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                        sql &= "  And Y_Level = 0 "
                                        sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                        sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                        sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                        sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                        sql &= "Order by FCTNo, FCTSubNo "
                                        Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                        For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                            SPWriteNewFCTPlan(pBuyer, xBuyerRule, xLTType, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, xItem(ItemIndex), xItem1(ItemIndex1), xColor, "PS", xQtyMeter, xRuleNo(RuleIndex))
                                        Next
                                        xSubNo = xSubNo + 1
                                        FCTWrite = True
                                    End If
                                Next
                            End If
                            ' 搜尋不到任何結構資料(Set Error Rule [9900-990-99])
                            If Not FCTWrite Then
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    SPWriteNewFCTPlan(pBuyer, xBuyerRule, xLTType, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "PS", xQtyMeter, xRuleNo(RuleIndex) + "/9900-990-00")
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                        Next
                    Next
                End If
                ' ZIP是否為採購品  (0=不是/1=是)
                If Not FCTWrite Then
                    If oWaves.GetPurchaseItem(dt_FCTItem.Rows(i).Item("Item")) = 1 Then
                        ' 寫入 ForcastPlan
                        sql = "SELECT * FROM ForcastPlan "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                        sql &= "  And Y_Level = 0 "
                        sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                        sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                        sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                        sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                        sql &= "Order by FCTNo, FCTSubNo "
                        Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                        For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                            SPWriteNewFCTPlan(pBuyer, xBuyerRule, xLTType, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), dt_FCTItem.Rows(i).Item("Item"), dt_FCT2.Rows(j).Item("Y_Color").ToString, "ZIP", 100, "1000-010-00")
                        Next
                        xSubNo = xSubNo + 1
                        FCTWrite = True
                    End If
                End If
            Next
            '
            ' Recuild Forcastplan
            Dim Conn As New SqlClient.SqlConnection("Data Source=" & "10.245.0.112" & ";" & _
                                                    "Initial Catalog=" & "EDI" & ";" & _
                                                    "User ID=" & "sa" & ";" & _
                                                    "Password=;")
            Dim cmd As SqlCommand = Conn.CreateCommand
            '
            Conn.Open()
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandTimeout = 180
            cmd.Parameters.Add(New SqlParameter("@Customer", pBuyer))
            cmd.CommandText = "sp_ForcastPlanRebuild"
            cmd.ExecuteNonQuery()
            Conn.Close()
            '
            oDB.AccessLog(pLogID, pBuyer, "SPS_ForcastPlan", NowDateTime, 0, "999", "ForcastPlan-End", "PGM=SPForcastPlan", "", "", pUserID, "")
            '
            '##
            'MsgBox("OUT")
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "SPS_ForcastPlan", NowDateTime, 1, "999", "ForcastPlan-Error", "PGM=SPForcastPlan", "", "", pUserID, "")
            '
        End Try
        '
        Return RtnCode
    End Function
    'SPForcastPlan-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([925]-SPWriteNewFCTPlan)
    '**     SHOPPING Write Forcast Plan
    '***********************************************************************************************
    'SPWriteNewFCTPlan-Start
    <WebMethod()> _
    Public Function SPWriteNewFCTPlan(ByVal pBuyer As String, _
                                      ByVal pRefBuyer As String, _
                                      ByVal pRefLTType As String, _
                                      ByVal pID As Integer, _
                                      ByVal pSubNo As Integer, _
                                      ByVal pLevel As Integer, _
                                      ByVal pFatherItem As String, _
                                      ByVal pItem As String, _
                                      ByVal pColor As String, _
                                      ByVal pClass As String, _
                                      ByVal pQty As String, _
                                      ByVal pRuleNo As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql, xsql, str As String
        Dim Qty, QtyTotal As Double
        '
        Try
            oWaves.Timeout = Timeout.Infinite

            sql = "SELECT * FROM ForcastPlan "
            sql &= "Where Unique_ID = '" & CStr(pID) & "' "
            Dim dt_FCTData As DataTable = oDataBase.GetDataTable(sql)
            If dt_FCTData.Rows.Count > 0 Then
                '
                sql = "Insert into ForcastPlan "
                sql &= "( "
                sql &= "BuyerGroup, "
                sql &= "FCTNo, "
                sql &= "FCTSubNo, "
                sql &= "Version, "
                sql &= "C_Code, "
                sql &= "C_Color, "
                sql &= "C_SpecialRequest, "
                sql &= "C_Season, "
                sql &= "C_ShortenLT, "
                sql &= "C_Style, "
                sql &= "C_A1, "
                sql &= "C_B1, "
                sql &= "C_C1, "
                sql &= "C_D1, "
                sql &= "C_E1, "
                sql &= "C_F1, "
                sql &= "C_G1, "
                sql &= "C_H1, "
                sql &= "C_I1, "
                sql &= "C_J1, "
                sql &= "C_K1, "
                sql &= "C_L1, "
                sql &= "C_M1, "
                sql &= "C_N1, "
                sql &= "C_O1, "

                sql &= "Y_LEVEL, "
                sql &= "Y_ItemCode, "
                sql &= "Y_ItemName1, "
                sql &= "Y_ItemName2, "
                sql &= "Y_ItemName3, "
                sql &= "Y_Color, "
                sql &= "Y_A1, "
                sql &= "Y_B1, "
                sql &= "Y_C1, "
                sql &= "Y_D1, "
                sql &= "Y_E1, "
                sql &= "Y_F1, "
                sql &= "Y_G1, "
                sql &= "Y_H1, "
                sql &= "Y_I1, "
                sql &= "Y_J1, "

                sql &= "N_F, "
                sql &= "N1_F, "
                sql &= "N2_F, "
                sql &= "N3_F, "
                sql &= "N4_F, "
                sql &= "N5_F, "
                sql &= "N6_F, "
                sql &= "N7_F, "
                sql &= "N8_F, "
                sql &= "N9_F, "
                sql &= "N10_F, "
                sql &= "N11_F, "
                sql &= "N12_F, "
                sql &= "Total, "

                sql &= "BUYER, "
                sql &= "ID, "
                sql &= "CreateUser, "
                sql &= "CreateTime "
                sql &= " ) "
                '
                ' 取得 備料MASTER-INF (JOYJOY)
                ' ------------------------------------------------------
                Dim xLTType, xPartType, xRuleCode, xRuleNo, xRuleSubno, xKeep As String
                Dim xRatio1, xRatio2, xRatio3, xRatio4 As Integer
                ' 決定LT種類(LTType)
                xLTType = pRefLTType
                ' 決定Part Type
                xPartType = "CH"
                If pClass = "PS" Then xPartType = "SLD" ' PARTTYPE      
                If pClass = "ZIP" Then xPartType = "ZIP" ' PARTTYPE      
                ' 初值
                xRatio1 = 100
                xRatio2 = 100
                xRatio3 = 100
                xRatio4 = 100
                xKeep = ""
                '
                str = pRuleNo   ' RULE-CODE, RULE-N0, RULESUBNO
                If InStr(str, "/") > 0 Then str = Mid(str, InStr(str, "/") + 1)
                xRuleCode = Mid(str, 1, InStr(str, "-") - 1)
                str = Mid(str, InStr(str, "-") + 1)
                xRuleNo = Mid(str, 1, InStr(str, "-") - 1)
                str = Mid(str, InStr(str, "-") + 1)
                xRuleSubno = str
                '
                xsql = "SELECT Action, Ratio1, Ratio2, Ratio3, Ratio4, Keep, "
                xsql &= "str(Ratio1,3,0) + '/' + str(Ratio2,3,0) + '/' + str(Ratio3,3,0) + '/' + str(Ratio4,3,0) + '/' As Ratio, "
                xsql &= "SearchItem1 + '/' + SearchItem2 + '/' + SearchItem3 + '/' +SearchItem4 + '/' + SearchItem5 + '/' + "
                xsql &= "SearchItem6 + '/' + SearchItem7 + '/' + SearchItem8 + '/' +SearchItem9 + '/' + SearchItem10 + '/' As SearchItem "
                xsql &= "FROM M_LSOrderRule "
                xsql &= "Where Active = '1' "
                xsql &= "  And Buyer = '" & pRefBuyer & "' "
                xsql &= "  And LTType = '" & xLTType & "' "
                xsql &= "  And PartType = '" & xPartType & "' "
                xsql &= "  And RuleCode = '" & xRuleCode & "' "
                xsql &= "  And RuleNo = '" & xRuleNo & "' "
                xsql &= "  And RuleSubno = '" & xRuleSubno & "' "
                xsql &= "Order by RuleCode, RuleNo, RuleSubno "
                Dim dt_LSOrderRule As DataTable = oDataBase.GetDataTable(xsql)
                If dt_LSOrderRule.Rows.Count > 0 Then
                    ' SHOPPING MODIFY-START
                    ' N ~ N+4  ALL = 100%
                    'xRatio1 = dt_LSOrderRule.Rows(0).Item("Ratio1")
                    'xRatio2 = dt_LSOrderRule.Rows(0).Item("Ratio2")
                    'xRatio3 = dt_LSOrderRule.Rows(0).Item("Ratio3")
                    'xRatio4 = dt_LSOrderRule.Rows(0).Item("Ratio4")
                    xRatio1 = 100
                    xRatio2 = 100
                    xRatio3 = 100
                    xRatio4 = 100
                    ' SHOPPING MODIFY-END
                    xKeep = dt_LSOrderRule.Rows(0).Item("Keep").ToString
                End If
                '
                ' 製作 FORECAST
                ' ---------------------------------------------------------
                sql &= "VALUES( "
                sql &= " '" & dt_FCTData.Rows(0).Item("BuyerGroup").ToString & "', "                ' BuyerGroup
                sql &= " '" & dt_FCTData.Rows(0).Item("FCTNo").ToString & "', "                     ' FCTNo
                sql &= " " & CStr(pSubNo) & ", "                                                    ' FCTSubNo
                sql &= " " & dt_FCTData.Rows(0).Item("Version").ToString & ", "                     ' Version
                sql &= " '" & dt_FCTData.Rows(0).Item("C_Code").ToString & "', "                    ' C_Code
                sql &= " '" & dt_FCTData.Rows(0).Item("C_Color").ToString & "', "                   ' C_Color
                sql &= " '" & dt_FCTData.Rows(0).Item("C_SpecialRequest").ToString & "', "          ' C_SpecialRequest
                sql &= " '" & dt_FCTData.Rows(0).Item("C_Season").ToString & "', "                  ' C_Season

                If xKeep = "Y" Then                                                                 ' C_ShortenLT    
                    sql &= " '" & dt_FCTData.Rows(0).Item("C_ShortenLT").ToString & "', "
                Else
                    sql &= " '" & "" & "', "
                End If
                sql &= " '" & dt_FCTData.Rows(0).Item("C_Style").ToString & "', "                   ' C_Style
                sql &= " '" & dt_FCTData.Rows(0).Item("C_A1").ToString & "', "                      ' C_A1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_B1").ToString & "', "                      ' C_B1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_C1").ToString & "', "                      ' C_C1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_D1").ToString & "', "                      ' C_D1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_E1").ToString & "', "                      ' C_E1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_F1").ToString & "', "                      ' C_F1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_G1").ToString & "', "                      ' C_G1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_H1").ToString & "', "                      ' C_H1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_I1").ToString & "', "                      ' C_I1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_J1").ToString & "', "                      ' C_J1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_K1").ToString & "', "                      ' C_K1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_L1").ToString & "', "                      ' C_L1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_M1").ToString & "', "                      ' C_M1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_N1").ToString & "', "                      ' C_N1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_O1").ToString & "', "                      ' C_O1
                ' ------------------------------------------------------------------
                sql &= " " & CStr(pLevel) & ", "                                                    ' Y_LEVEL
                sql &= " '" & pItem & "', "                                                         ' Y_ItemCode
                oWaves.GetItemName001(pItem, 1, str)                                                ' Y_ItemName1
                sql &= " '" & str & "', "
                oWaves.GetItemName001(pItem, 2, str)                                                ' Y_ItemName2
                sql &= " '" & str & "', "
                oWaves.GetItemName001(pItem, 3, str)                                                ' Y_ItemName3
                sql &= " '" & str & "', "

                If xPartType = "ZIP" Then                                                           ' Y_Color(考慮兼用色)(ZIP時不考慮)
                    str = pColor
                Else
                    oWaves.GetChangeColor("01", pFatherItem, pItem, pColor, str)
                End If
                sql &= " '" & str & "', "

                sql &= " '" & pClass & "', "                                                        ' Y_A1 (Item Class)
                If InStr(pRuleNo, "/") > 0 Then                                                     ' 異常
                    sql &= " '" & Mid(pRuleNo, 1, InStr(pRuleNo, "/") - 1) & "', "                  ' Y_B1 (Rule Inf-原LS-RuleNo)
                    sql &= " '" & Mid(pRuleNo, InStr(pRuleNo, "/") + 1) & "', "                     ' Y_C1 (Rule Inf-採用LS-RulenNo)
                Else                                                                                ' 正常
                    sql &= " '" & pRuleNo & "', "
                    sql &= " '" & pRuleNo & "', "
                End If

                If dt_LSOrderRule.Rows.Count > 0 Then
                    sql &= " '" & dt_LSOrderRule.Rows(0).Item("Action").ToString & "', "            ' Y_D1 (Rule Inf-Action)
                    sql &= " '" & dt_LSOrderRule.Rows(0).Item("SearchItem").ToString & "', "        ' Y_E1 (SearchItem)
                    sql &= " '" & dt_LSOrderRule.Rows(0).Item("Keep").ToString & "', "              ' Y_F1 (Keep)
                    ' SHOPPING MODIFY-START
                    ' N ~ N+4  ALL = 100%
                    'sql &= " '" & dt_LSOrderRule.Rows(0).Item("Ratio").ToString & "', "             ' Y_G1 (N+1 ~ N+4)
                    sql &= " '" & "100/100/100/100/100/" & "', "                                     ' Y_G1 (N+1 ~ N+4)
                    ' SHOPPING MODIFY-END
                Else
                    sql &= " '" & dt_FCTData.Rows(0).Item("Y_D1").ToString & "', "                  ' Y_D1
                    sql &= " '" & dt_FCTData.Rows(0).Item("Y_E1").ToString & "', "                  ' Y_E1
                    sql &= " '" & dt_FCTData.Rows(0).Item("Y_F1").ToString & "', "                  ' Y_F1
                    sql &= " '" & dt_FCTData.Rows(0).Item("Y_G1").ToString & "', "                  ' Y_G1
                End If
                sql &= " '" & dt_FCTData.Rows(0).Item("Y_H1").ToString & "', "                      ' Y_H1
                sql &= " '" & dt_FCTData.Rows(0).Item("Y_I1").ToString & "', "                      ' Y_I1
                sql &= " '" & dt_FCTData.Rows(0).Item("Y_J1").ToString & "', "                      ' Y_J1

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N_F"))                             ' N_F
                QtyTotal = Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N1_F"))                            ' N1_F
                If pClass = "PS" Then
                    Qty = Fix(Qty * (xRatio1 / 100) + 0.9)
                Else
                    Qty = Qty * (xRatio1 / 100)
                End If
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N2_F"))                            ' N2_F
                If pClass = "PS" Then
                    Qty = Fix(Qty * (xRatio2 / 100) + 0.9)
                Else
                    Qty = Qty * (xRatio2 / 100)
                End If
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N3_F"))                            ' N3_F
                If pClass = "PS" Then
                    Qty = Fix(Qty * (xRatio3 / 100) + 0.9)
                Else
                    Qty = Qty * (xRatio3 / 100)
                End If
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N4_F"))                            ' N4_F
                If pClass = "PS" Then
                    Qty = Fix(Qty * (xRatio4 / 100) + 0.9)
                Else
                    Qty = Qty * (xRatio4 / 100)
                End If
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N5_F"))                            ' N5_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N6_F"))                            ' N6_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N7_F"))                            ' N7_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N8_F"))                            ' N8_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N9_F"))                            ' N9_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N10_F"))                           ' N10_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N11_F"))                           ' N11_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N12_F"))                           ' N12_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                sql &= " " & CStr(QtyTotal) & ", "                                                  ' Total

                sql &= " '" & dt_FCTData.Rows(0).Item("BUYER").ToString & "', "                     ' BUYER
                sql &= " " & dt_FCTData.Rows(0).Item("ID").ToString & ", "                          ' ID
                sql &= " '" & "FCTPlan" & "', "
                sql &= " '" & NowDateTime & "' "
                sql &= " ) "
                oDataBase.ExecuteNonQuery(sql)
                '
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'SPWriteNewFCTPlan-End
    '***********************************************************************************************
    '**([930]-SPLocalStockPlan)  
    '**     SHOPPING Local Stock Plan展開 
    '***********************************************************************************************
    'SPLocalStockPlan-Start
    <WebMethod()> _
    Public Function SPLocalStockPlan(ByVal pLogID As String, _
                                     ByVal pBuyer As String, _
                                     ByVal pUserID As String, _
                                     ByVal pGRBuyer As String, _
                                     ByVal pFunList As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        '
        Try
            '
            oDB.AccessLog(pLogID, pBuyer, "SPS_LocalStockPlan", NowDateTime, 0, "999", "LocalStockPlan-Start", "PGM=SPLocalStockPlan", "", "", pUserID, "")
            '
            ' ***********************************************************************************
            ' 變數定義及設定初值
            ' ***********************************************************************************
            Dim i, j As Integer
            Dim xLastVersion As Integer = 99    ' 最終版PLAN
            Dim xSystemItem As String = ""
            Dim xGroupItem As String = ""
            Dim xSelectField As String = ""
            Dim xFCTString, xLSString, xDescr, str As String
            Dim xFCTGroupField(), xLSGroupField() As String
            Dim xKeepCode As String = ""
            Dim xQty As String = "0"
            Dim xSumQty As Double = 0           ' N月生產+在庫合計
            Dim xKeepQty As Double = 0          ' SCHE-PROD + ON-PROD + KEEP-INV
            Dim xFreeQty As Double = 0          ' FREE-INV
            Dim xFCTQty As Double = 0           ' ORDER ZONE FCT QTY
            Dim xFCTSumQty As Double = 0        ' 合計FCT QTY 
            Dim xProdQty As Double = 0          ' 需生產合計
            Dim xMonScheProd(6), xMonOnProd(6) As String        ' N1 ~ N4 SCHE_PRODQTY, ON_PRODQTY
            Dim xPurchaseItem As Boolean        ' 採購品
            Dim xPriceA As String = "0"         ' A單價
            Dim xPriceB As String = "0"         ' B單價
            Dim xPrice As Double = 0            ' 單價
            Dim xMetterLength As Double = 0     ' Length(公呎單位)

            oWaves.Timeout = Timeout.Infinite
            ' ---------------------------------------------------------------------------------
            ' 取得 SystemItem 初值
            sql = "SELECT * FROM M_ImportRule "
            sql &= "Where Buyer = '" & pBuyer & "-2' "
            sql &= "  And Sts   = '1' "
            sql &= "  And DataRule Like '" & "SY_" & "%' "
            sql &= "Order by DataRule "
            Dim dt_List As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_List.Rows.Count - 1
                ' SystemItem
                If xSystemItem = "" Then
                    xSystemItem = dt_List.Rows(i).Item("Field").ToString
                Else
                    xSystemItem = xSystemItem & "," & dt_List.Rows(i).Item("Field").ToString
                End If
                ' SelectField
                If xSelectField = "" Then
                    xSelectField = dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("Field").ToString
                Else
                    xSelectField = xSelectField & ", " & dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("Field").ToString
                End If
            Next
            ' 取得FCT & LS Group Field
            xFCTString = xSystemItem
            xLSString = xSystemItem
            ' ---------------------------------------------------------------------------------
            ' 取得 GroupItem 初值
            dt_List.Clear()
            sql = "SELECT * FROM M_ImportRule "
            sql &= "Where Buyer = '" & pBuyer + "-2" & "' "
            sql &= "  And Sts   = '1' "
            sql &= "  And DataRule Like '" & "GR_" & "%' "
            sql &= "Order by DataRule "
            dt_List = oDataBase.GetDataTable(sql)
            For i = 0 To dt_List.Rows.Count - 1
                ' GroupItem
                If xGroupItem = "" Then
                    xGroupItem = dt_List.Rows(i).Item("Field").ToString
                Else
                    xGroupItem = xGroupItem & "," & dt_List.Rows(i).Item("Field").ToString
                End If
                ' SelectField
                If xSelectField = "" Then
                    xSelectField = dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("DataRule").ToString
                Else
                    xSelectField = xSelectField & ", " & dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("DataRule").ToString
                End If
                ' 取得FCT & LS Group Field
                xFCTString = xFCTString & "," & dt_List.Rows(i).Item("Field").ToString
                xLSString = xLSString & "," & dt_List.Rows(i).Item("DataRule").ToString
            Next
            '
            ' 取得FCT & LS Group Field
            xFCTGroupField = xFCTString.Split(",")
            xLSGroupField = xLSString.Split(",")
            '
            ' GroupItem 不足10個時 ==> SelectField 需補至 10 個
            For j = i + 1 To 10
                If j < 10 Then
                    xSelectField = xSelectField & ", " & "'' AS GR_0" & CStr(j)
                Else
                    xSelectField = xSelectField & ", " & "'' AS GR_10"
                End If
            Next
            ' ---------------------------------------------------------------------------------
            ' 取得 SummaryItem 初值
            dt_List.Clear()
            sql = "SELECT * FROM M_ImportRule "
            sql &= "Where Buyer = '" & pBuyer + "-2" & "' "
            sql &= "  And Sts   = '1' "
            sql &= "  And DataRule Like '" & "FS_" & "%' "
            sql &= "Order by DataRule "
            dt_List = oDataBase.GetDataTable(sql)
            For i = 0 To dt_List.Rows.Count - 1
                ' SelectField
                If xSelectField = "" Then
                    xSelectField = "Sum(" & dt_List.Rows(i).Item("Field").ToString & ") AS " & dt_List.Rows(i).Item("DataRule").ToString
                Else
                    xSelectField = xSelectField & ", " & "Sum(" & dt_List.Rows(i).Item("Field").ToString & ") AS " & dt_List.Rows(i).Item("DataRule").ToString
                End If
            Next
            ' ***********************************************************************************
            ' 取得 LSNO = GroupCode + "L" + 年月 + 流水號(4碼)   SA + L + 135 + 0001
            ' 變數定義及設定初值
            ' ***********************************************************************************
            Dim xLSNo, xSeqNoString, xGroup, xItem As String
            Dim xSeqNo, xSubNo As Integer
            xGroup = "XX"
            xItem = ""
            ' 取得GroupCode
            sql = "SELECT GrCode FROM M_SPControlRecord "
            sql &= "Where Code = '" & pBuyer & "' "
            Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
            If dt_ControlRecord.Rows.Count > 0 Then xGroup = dt_ControlRecord.Rows(0).Item("GrCode")
            ' 取得年月 
            Select Case Month(Now)
                Case 10
                    xLSNo = xGroup + "L" + CStr(Year(Now) - 2000) + "A"
                Case 11
                    xLSNo = xGroup + "L" + CStr(Year(Now) - 2000) + "B"
                Case 12
                    xLSNo = xGroup + "L" + CStr(Year(Now) - 2000) + "C"
                Case Else
                    xLSNo = xGroup + "L" + CStr(Year(Now) - 2000) + CStr(Month(Now))
            End Select
            ' 流水號(下一可使用No)及SubNo
            xSeqNo = 1
            xSubNo = 1
            sql = "Select LSNo From M_SPControlRecord "
            sql = sql + "Where Code = '" & pBuyer & "' "
            sql = sql + "  And LSNo Like '" & xLSNo & "%' "
            Dim dt_LSNo As DataTable = oDataBase.GetDataTable(sql)
            If dt_LSNo.Rows.Count > 0 Then
                xSeqNo = CInt(Mid(dt_LSNo.Rows(0).Item("LSNo"), 11, 5))
            End If
            '
            ' ***********************************************************************************
            ' Local Stock Plan展開
            ' ***********************************************************************************
            sql = "SELECT "
            sql &= xSelectField
            sql &= " FROM ForcastPlan "
            sql &= " Where Buyer = '" & pBuyer & "' "
            sql &= "   And Version = " & xLastVersion & " "
            sql &= "   And Y_Level > 0 "
            '
            sql &= "   And CHARINDEX( "
            sql &= "         case when C_G1='SLIDER ONLY' then 'PS'"
            sql &= "              when C_G1='CHAIN ONLY' then 'CH' "
            sql &= "              else 'ZIP' "
            sql &= "         end, "
            sql &= "         'ZIP-' + Y_A1 "
            sql &= "       ) > 0 "
            '
            sql &= " Group by " & xSystemItem & ", " & xGroupItem
            sql &= " Order by " & xSystemItem & ", " & xGroupItem
            Dim dt_FCT As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_FCT.Rows.Count - 1
                ' Insert LocalStock Plan Data
                sql = "Insert into LocalStockPlan "
                sql &= "( "
                sql &= "BUYER, "
                sql &= "BUYERGROUP, "
                sql &= "LSNO, "
                sql &= "LSSUBNO, "
                sql &= "VERSION, "

                sql &= "GR_01, "
                sql &= "GR_02, "
                sql &= "GR_03, "
                sql &= "GR_04, "
                sql &= "GR_05, "
                sql &= "GR_06, "
                sql &= "GR_07, "
                sql &= "GR_08, "
                sql &= "GR_09, "
                sql &= "GR_10, "

                sql &= "MINIMUMSTOCK, "
                sql &= "N_SCHEPROD, "
                sql &= "N_ONPROD, "
                sql &= "N_FREEINV, "
                sql &= "N_KEEPINV, "
                sql &= "N_TOTAL, "

                sql &= "SP_00, "        'ADD-START 13/11/11
                sql &= "OP_00, "        'ADD-START 13/11/11
                sql &= "FS_00, "        'ADD-START 13/12/4
                sql &= "PS_00, "        'ADD-START 13/12/4
                sql &= "IS_00, "        'ADD-START 13/12/4
                sql &= "SP_01, "        'ADD-START 13/11/11
                sql &= "OP_01, "        'ADD-START 13/11/11
                sql &= "FS_01, "
                sql &= "PS_01, "
                sql &= "IS_01, "
                sql &= "SP_02, "        'ADD-START 13/11/11
                sql &= "OP_02, "        'ADD-START 13/11/11
                sql &= "FS_02, "
                sql &= "PS_02, "
                sql &= "IS_02, "
                sql &= "SP_03, "        'ADD-START 13/11/11
                sql &= "OP_03, "        'ADD-START 13/11/11
                sql &= "FS_03, "
                sql &= "PS_03, "
                sql &= "IS_03, "
                sql &= "SP_04, "        'ADD-START 13/11/11
                sql &= "OP_04, "        'ADD-START 13/11/11
                sql &= "FS_04, "
                sql &= "PS_04, "
                sql &= "IS_04, "
                sql &= "SP_05, "        'ADD-START 13/11/11
                sql &= "OP_05, "        'ADD-START 13/11/11
                sql &= "FS_05, "
                sql &= "PS_05, "
                sql &= "IS_05, "
                sql &= "FS_06, "
                sql &= "PS_06, "
                sql &= "IS_06, "
                sql &= "FS_07, "
                sql &= "PS_07, "
                sql &= "IS_07, "
                sql &= "FS_08, "
                sql &= "PS_08, "
                sql &= "IS_08, "
                sql &= "FS_09, "
                sql &= "PS_09, "
                sql &= "IS_09, "
                sql &= "FS_10, "
                sql &= "PS_10, "
                sql &= "IS_10, "
                sql &= "FS_11, "
                sql &= "PS_11, "
                sql &= "IS_11, "
                sql &= "FS_12, "
                sql &= "PS_12, "
                sql &= "IS_12, "

                sql &= "FS_13, "
                sql &= "PS_13, "
                sql &= "FREEALLOC, "
                sql &= "DESCRIPTION, "
                sql &= "BULSNo, "

                sql &= "CreateUser, "
                sql &= "CreateTime "
                sql &= " ) "
                sql &= "VALUES( "
                sql &= " '" & dt_FCT.Rows(i).Item("Buyer").ToString & "', "                 ' Buyer
                sql &= " '" & dt_FCT.Rows(i).Item("BuyerGroup").ToString & "', "            ' BuyerGroup
                ' ---------------------------------------------------------------------------------
                ' 設定 LSNo, SubNo
                ' Item Code是否相同 Counter LSNo, SubNo
                If dt_FCT.Rows(i).Item("GR_03").ToString <> xItem Then
                    If i > 0 Then
                        xSeqNo = xSeqNo + 1
                        xSubNo = 1
                    End If
                    xItem = dt_FCT.Rows(i).Item("GR_03").ToString
                Else
                    If i > 0 Then xSubNo = xSubNo + 1
                End If
                ' 製作SeqNo(不足5位補0)
                xSeqNoString = CStr(xSeqNo)
                If Len(xSeqNoString) < 2 Then
                    xSeqNoString = "0000" + CStr(xSeqNo)
                Else
                    If Len(xSeqNoString) < 3 Then
                        xSeqNoString = "000" + CStr(xSeqNo)
                    Else
                        If Len(xSeqNoString) < 4 Then
                            xSeqNoString = "00" + CStr(xSeqNo)
                        Else
                            If Len(xSeqNoString) < 5 Then
                                xSeqNoString = "0" + CStr(xSeqNo)
                            End If
                        End If
                    End If
                End If

                '設定 LSNo, SubNo
                sql &= " '" & xLSNo + xSeqNoString & "', "                                  ' LSNO
                sql &= " " & CStr(xSubNo) & ", "                                            ' LSSUBNO
                ' ---------------------------------------------------------------------------------
                sql &= " " & CStr(xLastVersion) & ", "                                      ' Version

                sql &= " '" & dt_FCT.Rows(i).Item("GR_01").ToString & "', "                 ' GR_01
                sql &= " '" & dt_FCT.Rows(i).Item("GR_02").ToString & "', "                 ' GR_02
                sql &= " '" & dt_FCT.Rows(i).Item("GR_03").ToString & "', "                 ' GR_03
                sql &= " '" & dt_FCT.Rows(i).Item("GR_04").ToString & "', "                 ' GR_04
                sql &= " '" & dt_FCT.Rows(i).Item("GR_05").ToString & "', "                 ' GR_05
                sql &= " '" & dt_FCT.Rows(i).Item("GR_06").ToString & "', "                 ' GR_06
                sql &= " '" & dt_FCT.Rows(i).Item("GR_07").ToString & "', "                 ' GR_07
                sql &= " '" & dt_FCT.Rows(i).Item("GR_08").ToString & "', "                 ' GR_08
                sql &= " '" & dt_FCT.Rows(i).Item("GR_09").ToString & "', "                 ' GR_09
                sql &= " '" & dt_FCT.Rows(i).Item("GR_10").ToString & "', "                 ' GR_10
                ' ---------------------------------------------------------------------------------

                If dt_FCT.Rows(i).Item("GR_08").ToString <> "ZIP" Then
                    '** PS / CH
                    ' ---------------------------------------------------------------------------------
                    ' 取得 MinimumStock Qty
                    oWaves.GetMininumStock("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xQty)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                    Else
                        sql &= " " & "0" & ", "
                    End If
                    ' ---------------------------------------------------------------------------------
                    ' 歸零作業
                    xSumQty = 0
                    xProdQty = 0
                    xKeepQty = 0            ' SCHE-PROD + ON-PROD + KEEP-INV
                    xFreeQty = 0            ' FREE-INV
                    xFCTQty = 0             ' ORDER ZONE FCT QTY
                    xDescr = ""             ' Free Qty備註說明
                    For j = 1 To 6          ' N1~N4 PROD-QTY
                        xMonScheProd(j) = 0
                        xMonOnProd(j) = 0
                    Next
                    ' 判斷是否採購品(0 = 不是 / 1 = 是)
                    xPurchaseItem = False
                    If oWaves.GetPurchaseItem(dt_FCT.Rows(i).Item("GR_03").ToString) = 1 Then
                        xPurchaseItem = True
                    End If
                    ' ---------------------------------------------------------------------------------
                    If xPurchaseItem Then
                        ' 採購品
                        ' 不使用 (N_SCHEPROD)
                        sql &= " " & "0" & ", "
                        '
                        ' 採購-ON Qty (N_ONPROD)
                        xKeepCode = dt_FCT.Rows(i).Item("GR_02").ToString

                        oWaves.GetPurchaseQty("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xKeepCode, xQty, xMonOnProd)
                        If CDbl(xQty) > 0 Then
                            sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                            '
                            If Trim(dt_FCT.Rows(i).Item("GR_11").ToString) = "YES" Then
                                xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                                xSumQty = xSumQty + CDbl(xQty) / 10000000
                            End If
                        Else
                            sql &= " " & "0" & ", "
                        End If
                    Else
                        ' 製造品
                        ' 生產-SCHE Qty (N_SCHEPROD)
                        xKeepCode = dt_FCT.Rows(i).Item("GR_02").ToString

                        oWaves.GetProductionQty("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xKeepCode, 0, xQty, xMonScheProd)
                        If CDbl(xQty) > 0 Then
                            sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                            '
                            If Trim(dt_FCT.Rows(i).Item("GR_11").ToString) = "YES" Then
                                xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                                xSumQty = xSumQty + CDbl(xQty) / 10000000
                            End If
                        Else
                            sql &= " " & "0" & ", "
                        End If
                        '
                        ' 生產-ON Qty (N_ONPROD)
                        oWaves.GetProductionQty("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xKeepCode, 1, xQty, xMonOnProd)
                        If CDbl(xQty) > 0 Then
                            sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                            '
                            If Trim(dt_FCT.Rows(i).Item("GR_11").ToString) = "YES" Then
                                xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                                xSumQty = xSumQty + CDbl(xQty) / 10000000
                            End If
                        Else
                            sql &= " " & "0" & ", "
                        End If
                    End If
                    '
                    ' 在庫-Free Qty (N_FREEINV)
                    oWaves.GetFreeInventory("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xQty)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xFreeQty = xFreeQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                    '
                    ' 在庫-Keep Qty (N_KEEPINV)
                    oWaves.GetKeepCodeInventory("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xKeepCode, xQty)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        '
                        If Trim(dt_FCT.Rows(i).Item("GR_11").ToString) = "YES" Then
                            xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                            xSumQty = xSumQty + CDbl(xQty) / 10000000
                        End If
                    Else
                        sql &= " " & "0" & ", "
                    End If
                    '
                    ' 合計=(生產-SCHE Qty + 生產-ON Qty + 在庫-Free Qty +在庫-Keep Qty) (N_TOTAL)
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    '** VNDOR-FC ADD 17/11/24
                    '** ZIP 
                    ' ---------------------------------------------------------------------------------
                    ' 取得 MinimumStock Qty
                    oWaves.GetMininumStockZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xQty)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                    Else
                        sql &= " " & "0" & ", "
                    End If
                    ' ---------------------------------------------------------------------------------
                    ' 歸零作業
                    xSumQty = 0
                    xProdQty = 0
                    xKeepQty = 0            ' SCHE-PROD + ON-PROD + KEEP-INV
                    xFreeQty = 0            ' FREE-INV
                    xFCTQty = 0             ' ORDER ZONE FCT QTY
                    xDescr = ""             ' Free Qty備註說明
                    For j = 1 To 6          ' N1~N4 PROD-QTY
                        xMonScheProd(j) = 0
                        xMonOnProd(j) = 0
                    Next
                    ' 判斷是否採購品(0 = 不是 / 1 = 是)
                    xPurchaseItem = False

                    If oWaves.GetPurchaseItem(dt_FCT.Rows(i).Item("GR_03").ToString) = 1 Then
                        xPurchaseItem = True
                    End If
                    ' ---------------------------------------------------------------------------------
                    If xPurchaseItem Then
                        ' 採購品
                        ' 不使用 (N_SCHEPROD)
                        sql &= " " & "0" & ", "
                        '
                        ' 採購-ON Qty (N_ONPROD)
                        xKeepCode = dt_FCT.Rows(i).Item("GR_02").ToString
                        oWaves.GetPurchaseQtyZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xKeepCode, xQty, xMonOnProd)
                        If CDbl(xQty) > 0 Then
                            sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                            '
                            If Trim(dt_FCT.Rows(i).Item("GR_11").ToString) = "YES" Then
                                xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                                xSumQty = xSumQty + CDbl(xQty) / 10000000
                            End If
                        Else
                            sql &= " " & "0" & ", "
                        End If
                    Else
                        ' 製造品
                        ' 生產-SCHE Qty (N_SCHEPROD)
                        xKeepCode = dt_FCT.Rows(i).Item("GR_02").ToString

                        oWaves.GetProductionQtyZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xKeepCode, 0, xQty, xMonScheProd)
                        If CDbl(xQty) > 0 Then
                            sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                            '
                            If Trim(dt_FCT.Rows(i).Item("GR_11").ToString) = "YES" Then
                                xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                                xSumQty = xSumQty + CDbl(xQty) / 10000000
                            End If
                        Else
                            sql &= " " & "0" & ", "
                        End If
                        '
                        ' 生產-ON Qty (N_ONPROD)
                        oWaves.GetProductionQtyZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xKeepCode, 1, xQty, xMonOnProd)
                        If CDbl(xQty) > 0 Then
                            sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                            '
                            If Trim(dt_FCT.Rows(i).Item("GR_11").ToString) = "YES" Then
                                xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                                xSumQty = xSumQty + CDbl(xQty) / 10000000
                            End If
                        Else
                            sql &= " " & "0" & ", "
                        End If
                    End If
                    '
                    ' 在庫-Free Qty (N_FREEINV)
                    oWaves.GetFreeInventoryZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xQty)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xFreeQty = xFreeQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                    '
                    ' 在庫-Keep Qty (N_KEEPINV)
                    oWaves.GetKeepCodeInventoryZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xKeepCode, xQty)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        '
                        If Trim(dt_FCT.Rows(i).Item("GR_11").ToString) = "YES" Then
                            xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                            xSumQty = xSumQty + CDbl(xQty) / 10000000
                        End If
                    Else
                        sql &= " " & "0" & ", "
                    End If
                    '
                    ' 合計=(生產-SCHE Qty + 生產-ON Qty + 在庫-Free Qty +在庫-Keep Qty) (N_TOTAL)
                    sql &= " " & CStr(xSumQty) & ", "
                End If
                ' ---------------------------------------------------------------------------------
                ' FORCAST & PRODUCTION & INVENTORY
                '
                ' ---------------------------------------------------------------------------------
                ' SP00 / OP00 / FS_00 / PS_00 / IS_00
                sql &= " " & CStr(CDbl(xMonScheProd(1)) / 10000000) & ", "
                sql &= " " & CStr(CDbl(xMonOnProd(1)) / 10000000) & ", "
                ' --
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_00")
                ' 
                ' --ERIC&PEGGY 追加 2020/09/23
                xFCTQty = xFCTQty + dt_FCT.Rows(i).Item("FS_00")
                '
                sql &= " " & dt_FCT.Rows(i).Item("FS_00").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' SP01 / OP01 / FS_01 / PS_01 / IS_01
                sql &= " " & CStr(CDbl(xMonScheProd(2)) / 10000000) & ", "
                sql &= " " & CStr(CDbl(xMonOnProd(2)) / 10000000) & ", "

                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_01")
                xFCTQty = xFCTQty + dt_FCT.Rows(i).Item("FS_01")

                sql &= " " & CStr(dt_FCT.Rows(i).Item("FS_01")) & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                    xFCTQty = xFCTQty - xSumQty * -1
                End If
                '
                ' SP02 / OP02 / FS_02 / PS_02 / IS_02
                sql &= " " & CStr(CDbl(xMonScheProd(3)) / 10000000) & ", "
                sql &= " " & CStr(CDbl(xMonOnProd(3)) / 10000000) & ", "

                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_02")
                xFCTQty = xFCTQty + dt_FCT.Rows(i).Item("FS_02")
                '
                sql &= " " & CStr(dt_FCT.Rows(i).Item("FS_02")) & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                    xFCTQty = xFCTQty - xSumQty * -1
                End If
                '
                ' SP03 / OP03 / FS_03 / PS_03 / IS_03
                sql &= " " & CStr(CDbl(xMonScheProd(4)) / 10000000) & ", "
                sql &= " " & CStr(CDbl(xMonOnProd(4)) / 10000000) & ", "

                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_03")
                xFCTQty = xFCTQty + dt_FCT.Rows(i).Item("FS_03")
                '
                sql &= " " & CStr(dt_FCT.Rows(i).Item("FS_03")) & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                    xFCTQty = xFCTQty - xSumQty * -1
                End If
                '
                ' SP04 / OP04 / FS_04 / PS_04 / IS_04
                sql &= " " & CStr(CDbl(xMonScheProd(5)) / 10000000) & ", "
                sql &= " " & CStr(CDbl(xMonOnProd(5)) / 10000000) & ", "

                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_04")
                xFCTQty = xFCTQty + dt_FCT.Rows(i).Item("FS_04")
                '
                sql &= " " & CStr(dt_FCT.Rows(i).Item("FS_04")) & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                    xFCTQty = xFCTQty - xSumQty * -1
                End If
                '
                ' SP05 / OP05 / FS_05 / PS_05 / IS_05
                sql &= " " & CStr(CDbl(xMonScheProd(6)) / 10000000) & ", "
                sql &= " " & CStr(CDbl(xMonOnProd(6)) / 10000000) & ", "

                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_05")
                '
                sql &= " " & dt_FCT.Rows(i).Item("FS_05").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_06 / PS_06 / IS_06
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_06")
                '
                sql &= " " & dt_FCT.Rows(i).Item("FS_06").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_07 / PS_07 / IS_07
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_07")
                '
                sql &= " " & dt_FCT.Rows(i).Item("FS_07").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_08 / PS_08 / IS_08
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_08")
                '
                sql &= " " & dt_FCT.Rows(i).Item("FS_08").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_09 / PS_09 / IS_09
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_09")
                '
                sql &= " " & dt_FCT.Rows(i).Item("FS_09").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_10 / PS_10 / IS_10
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_10")
                '
                sql &= " " & dt_FCT.Rows(i).Item("FS_10").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_11 / PS_11 / IS_11
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_11")
                '
                sql &= " " & dt_FCT.Rows(i).Item("FS_11").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_12 / PS_12 / IS_12
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_12")
                '
                sql &= " " & dt_FCT.Rows(i).Item("FS_12").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_13 / PS_13
                sql &= " " & dt_FCT.Rows(i).Item("FS_13").ToString & ", "                   ' FS_13
                sql &= " " & CStr(xProdQty) & ", "                                          ' PS_13

                ' FREEALLOC
                If xFCTQty <= xKeepQty Then
                    sql &= " " & "0" & ", "
                Else
                    If xFCTQty - xKeepQty >= xFreeQty Then
                        sql &= " " & CStr(xFreeQty) & ", "
                    Else
                        sql &= " " & CStr(xFCTQty - xKeepQty) & ", "
                    End If
                End If
                '
                ' Description
                '  ** Free Qty
                If xFreeQty > 0 Then
                    oWaves.GetFreeByLocation("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xDescr)
                End If
                '  ** Free PROD Qty
                oWaves.GetFreeProductionQty("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xQty)
                If CDbl(xQty) > 0 Then
                    If xDescr <> "" Then
                        xDescr = xDescr + " * Free on Pro:[" + CStr(CDbl(xQty) / 10000000) + "]"
                    Else
                        xDescr = "Free on Pro:[" + CStr(CDbl(xQty) / 10000000) + "]"
                    End If
                End If
                '  ** Keep Qty
                If oWaves.GetMaterialKeepCodeInventory("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, str) = 0 Then
                    If str <> "" Then
                        If xDescr <> "" Then
                            xDescr = xDescr + " * Keep:[" & str & "]"
                        Else
                            xDescr = " * Keep:[" & str & "]"
                        End If
                    End If
                End If

                '  ** 
                If xDescr <> "" Then
                    sql &= " '" & xDescr & "', "
                Else
                    sql &= " '" & "" & "', "
                End If
                '
                sql &= " '" & Trim(dt_FCT.Rows(i).Item("GR_11").ToString) & "', "
                '-----------------------------
                '
                sql &= " '" & "LSPlan" & "', "
                sql &= " '" & NowDateTime & "' "
                sql &= " ) "
                oDataBase.ExecuteNonQuery(sql)
                '
                ' ---------------------------------------------------------------------------------
                ' Update FCT Plan Data (LSNo, LSSubNo)
                sql = "Update ForcastPlan Set "
                sql &= "LSNO = '" & xLSNo + xSeqNoString & "', "
                sql &= "LSSUBNO = " & CStr(xSubNo) & " "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "  And Version = " & xLastVersion & " "
                sql &= "  And Y_Level > 0 "
                For j = 0 To UBound(xFCTGroupField)
                    sql &= " And " & xFCTGroupField(j) & " = '" & dt_FCT.Rows(i).Item(xLSGroupField(j)).ToString & "' "
                Next
                oDataBase.ExecuteNonQuery(sql)
            Next
            ' ***********************************************************************************
            ' 更新下一個可使用No
            ' ***********************************************************************************
            ' 製作SeqNo(不足4位補0)
            xSeqNo = xSeqNo + 1
            xSeqNoString = CStr(xSeqNo)
            If Len(xSeqNoString) < 2 Then
                xSeqNoString = "0000" + CStr(xSeqNo)
            Else
                If Len(xSeqNoString) < 3 Then
                    xSeqNoString = "000" + CStr(xSeqNo)
                Else
                    If Len(xSeqNoString) < 4 Then
                        xSeqNoString = "00" + CStr(xSeqNo)
                    Else
                        If Len(xSeqNoString) < 5 Then
                            xSeqNoString = "0" + CStr(xSeqNo)
                        End If
                    End If
                End If
            End If
            ' 更新下一次可使用PONO
            sql = "Update M_SPControlRecord Set "
            sql = sql + "LSNo = '" & xLSNo + xSeqNoString & "' "
            sql = sql + "Where Code = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
            '
            oDB.AccessLog(pLogID, pBuyer, "SPS_LocalStockPlan", NowDateTime, 0, "999", "LocalStockPlan-End", "PGM=SPLocalStockPlan", "", "", pUserID, "")
            '
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "SPS_LocalStockPlan", NowDateTime, 1, "999", "LocalStockPlan-Error", "PGM=SPLocalStockPlan", "", "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'SPLocalStockPlan-End    '
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([935]-SPActionkPlan)
    '**     SHOPPING Action Plan展開
    '***********************************************************************************************
    'SPActionkPlan-Start
    <WebMethod()> _
    Public Function SPActionkPlan(ByVal pLogID As String, _
                                  ByVal pBuyer As String, _
                                  ByVal pUserID As String, _
                                  ByVal pGRBuyer As String, _
                                  ByVal pFunList As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            '
            oDB.AccessLog(pLogID, pBuyer, "SPS_ActionPlan", NowDateTime, 0, "999", "ActionPlan-Start", "PGM=SPActionkPlan", "", "", pUserID, "")
            '
            Dim Conn As New SqlClient.SqlConnection("Data Source=" & "10.245.0.112" & ";" & _
                                                    "Initial Catalog=" & "EDI" & ";" & _
                                                    "User ID=" & "sa" & ";" & _
                                                    "Password=;")
            Dim cmd As SqlCommand = Conn.CreateCommand
            '
            Conn.Open()
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandTimeout = 180
            cmd.Parameters.Add(New SqlParameter("@Customer", pBuyer))
            cmd.CommandText = "sp_SPActionPlan"
            cmd.ExecuteNonQuery()
            Conn.Close()
            '
            oDB.AccessLog(pLogID, pBuyer, "SPS_ActionPlan", NowDateTime, 0, "999", "ActionPlan-End", "PGM=SPActionkPlan", "", "", pUserID, "")
            '
        Catch ex As Exception
            RtnCode = 1
            '
            oDB.AccessLog(pLogID, pBuyer, "SPS_ActionPlan", NowDateTime, 1, "999", "ActionPlan-Error", "PGM=SPActionkPlan", "", "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'SPActionkPlan-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**[940]-SPKPInterface
    '**     SHOPPING KP I/F展開
    '***********************************************************************************************
    'SPKPInterface-Start
    <WebMethod()> _
    Public Function SPKPInterface(ByVal pLogID As String, ByVal pBuyer As String, _
                                  ByVal pUserID As String, ByVal pGRBuyer As String, _
                                  ByVal pFunList As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i, j As Integer
        Dim xReqDate(5) As Integer
        Dim xMaxQty, xProdQty As Double
        Dim sql, xYYMM As String
        Dim xLSNo, xOrderMonth, xLengthUnit As String
        Dim xSeqNo, xOrderZone As Integer
        '
        Try
            sql = "SELECT * FROM V_SPActionPlan2KP "
            sql &= "Where Customer = '" & pBuyer & "' "
            sql &= "  And Version = " & "99" & " "
            sql &= "  And [Action] in (1,2,3) "
            sql &= "  And N_Qty+N1_Qty+N2_Qty+N3_Qty+N4_Qty > 0 "
            sql &= "Order by SPNO, LSNO, LSSUBNO, [Action] "
            Dim dt_SPActionPlan2KP As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_SPActionPlan2KP.Rows.Count - 1
                ' 取得MaxQty
                ' ------------------------
                xMaxQty = 100000000     ' 1億
                sql = "SELECT * FROM M_Referp "
                sql &= " Where Cat = '" & "114" & "' "
                sql &= "   And DKey = '" & "MAXQTY-" + dt_SPActionPlan2KP.Rows(i).Item("PartType").ToString.Trim + "-" + pBuyer & "' "
                sql &= "Order by DKey, Data "
                Dim dt_Referp As DataTable = oDataBase.GetDataTable(sql)
                If dt_Referp.Rows.Count > 0 Then
                    xMaxQty = CDbl(dt_Referp.Rows(0).Item("Data").ToString.Trim)
                Else
                    dt_Referp.Clear()
                    sql = "SELECT * FROM M_Referp "
                    sql &= " Where Cat = '" & "114" & "' "
                    sql &= "   And DKey = '" & "MAXQTY-" + dt_SPActionPlan2KP.Rows(i).Item("PartType").ToString.Trim & "' "
                    sql &= "Order by DKey, Data "
                    dt_Referp = oDataBase.GetDataTable(sql)
                    If dt_Referp.Rows.Count > 0 Then
                        xMaxQty = CDbl(dt_Referp.Rows(0).Item("Data").ToString.Trim)
                    End If
                End If
                ' 取得N~N+4納期
                ' ------------------------
                For j = 1 To 5
                    xReqDate(j) = j * 30
                Next
                dt_Referp.Clear()
                sql = "SELECT * FROM M_Referp "
                sql &= " Where Cat = '" & "114" & "' "
                sql &= "   And DKey Like '%" & pBuyer & "-REQDATE" & "%' "
                sql &= "Order by DKey, Data "
                dt_Referp = oDataBase.GetDataTable(sql)
                For j = 0 To dt_Referp.Rows.Count - 1
                    xReqDate(j + 1) = CInt(dt_Referp.Rows(j).Item("Data").ToString.Trim)
                Next
                ' 取得BUYER ORDER ZONE
                ' ------------------------
                xOrderZone = 5
                dt_Referp.Clear()
                sql = "SELECT * FROM M_Referp "
                sql &= " Where Cat = '" & "114" & "' "
                sql &= "   And DKey = '" & "ZONE-" + pBuyer & "' "
                sql &= "Order by DKey, Data "
                dt_Referp = oDataBase.GetDataTable(sql)
                If dt_Referp.Rows.Count > 0 Then
                    xOrderZone = CDbl(dt_Referp.Rows(0).Item("Data").ToString.Trim)
                End If
                xLSNo = dt_SPActionPlan2KP.Rows(i).Item("LSNo")

                If dt_SPActionPlan2KP.Rows(i).Item("PartType") = "CH" Then
                    xLengthUnit = "M"
                Else
                    xLengthUnit = "P"
                End If
                ' 
                ' N
                ' ------------------------
                If xOrderZone >= 0 Then
                    ' 準備資料 (OrderMonth, xYYMM, xSeqNo) SPNO=SP20240728140206
                    If CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) < 10 Then
                        xOrderMonth = "0" + CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)))
                        xYYMM = Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 5, 2) + xOrderMonth
                    Else
                        If CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) > 12 Then
                            xOrderMonth = "0" + CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) - 12)
                            xYYMM = CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 5, 2)) + 1) + xOrderMonth
                        Else
                            xOrderMonth = CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)))
                            xYYMM = Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 5, 2) + xOrderMonth
                        End If
                    End If
                    '
                    xSeqNo = 1
                    ' 製作 EDI 資料
                    xProdQty = dt_SPActionPlan2KP.Rows(i).Item("N_Qty")
                    While xProdQty > 0
                        sql = "Insert into B_LS2EDIInterface "
                        sql &= "( "
                        sql &= "BuyerGroup, BULSNo, BULSSubNo, OrderMonth, Seqno, LSType, KeepCode, Item, ItemName1, ItemName2, ItemName3, Color, PartType, "
                        sql &= "Yobi1, Yobi2, Reference, LengthUnit, BuyerCode, ReqDate, ProdQty, CreateUser, CreateTime"
                        sql &= " ) "
                        sql &= "VALUES( "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Customer") & "', "
                        sql &= " '" & xLSNo & "', "
                        sql &= " " & dt_SPActionPlan2KP.Rows(i).Item("LSSubNo") & ", "
                        sql &= " '" & xOrderMonth & "', "
                        sql &= " " & CStr(xSeqNo) & ", "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ActionName") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Keep") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Item") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ItemName1") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ItemName2") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ItemName3") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Color") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("PartType") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("CustomerGr") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("N_YOBI1") & "', "
                        sql &= " '" & "[" & dt_SPActionPlan2KP.Rows(i).Item("SPNO") & "]" & " N BUY" & "', "
                        sql &= " '" & xLengthUnit & "', "
                        sql &= " '" & Mid(dt_SPActionPlan2KP.Rows(i).Item("Buyer"), 1, 6) & "', "
                        sql &= " " & Now.AddDays(xReqDate(1)).ToString("yyyyMMdd") & ", "
                        If xProdQty <= xMaxQty Then
                            sql &= " " & CStr(xProdQty) & ", "
                            xProdQty = 0
                        Else
                            sql &= " " & CStr(xMaxQty) & ", "
                            xProdQty = xProdQty - xMaxQty
                        End If
                        sql &= " '" & "SHOPPING" & "', "
                        sql &= " '" & NowDateTime & "' "
                        sql &= " ) "
                        oDataBase.ExecuteNonQuery(sql)
                        '
                        xSeqNo = xSeqNo + 1
                    End While
                End If
                ' N+1
                ' ------------------------
                If xOrderZone >= 1 Then
                    ' 準備資料 (OrderMonth, xYYMM, xSeqNo)
                    If CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 1 < 10 Then
                        xOrderMonth = "0" + CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 1)
                        xYYMM = Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 5, 2) + xOrderMonth
                    Else
                        If CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 1 > 12 Then
                            xOrderMonth = "0" + CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 1 - 12)
                            xYYMM = CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 5, 2)) + 1) + xOrderMonth
                        Else
                            xOrderMonth = CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 1)
                            xYYMM = Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 5, 2) + xOrderMonth
                        End If
                    End If
                    xSeqNo = 1
                    ' 製作 EDI 資料
                    xProdQty = dt_SPActionPlan2KP.Rows(i).Item("N1_Qty")
                    While xProdQty > 0
                        sql = "Insert into B_LS2EDIInterface "
                        sql &= "( "
                        sql &= "BuyerGroup, BULSNo, BULSSubNo, OrderMonth, Seqno, LSType, KeepCode, Item, ItemName1, ItemName2, ItemName3, Color, PartType, "
                        sql &= "Yobi1, Yobi2, Reference, LengthUnit, BuyerCode, ReqDate, ProdQty, CreateUser, CreateTime"
                        sql &= " ) "
                        sql &= "VALUES( "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Customer") & "', "
                        sql &= " '" & xLSNo & "', "
                        sql &= " " & dt_SPActionPlan2KP.Rows(i).Item("LSSubNo") & ", "
                        sql &= " '" & xOrderMonth & "', "
                        sql &= " " & CStr(xSeqNo) & ", "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ActionName") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Keep") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Item") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ItemName1") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ItemName2") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ItemName3") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Color") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("PartType") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("CustomerGr") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("N1_YOBI1") & "', "
                        sql &= " '" & "[" & dt_SPActionPlan2KP.Rows(i).Item("SPNO") & "]" + " N+1 BUY" & "', "
                        sql &= " '" & xLengthUnit & "', "
                        sql &= " '" & Mid(dt_SPActionPlan2KP.Rows(i).Item("Buyer"), 1, 6) & "', "
                        sql &= " " & Now.AddDays(xReqDate(2)).ToString("yyyyMMdd") & ", "
                        If xProdQty <= xMaxQty Then
                            sql &= " " & CStr(xProdQty) & ", "
                            xProdQty = 0
                        Else
                            sql &= " " & CStr(xMaxQty) & ", "
                            xProdQty = xProdQty - xMaxQty
                        End If
                        sql &= " '" & "SHOPPING" & "', "
                        sql &= " '" & NowDateTime & "' "
                        sql &= " ) "
                        oDataBase.ExecuteNonQuery(sql)

                        xSeqNo = xSeqNo + 1
                    End While
                End If
                ' N+2
                ' ------------------------
                If xOrderZone >= 2 Then
                    ' 準備資料 (OrderMonth, xYYMM, xSeqNo)
                    If CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 2 < 10 Then
                        xOrderMonth = "0" + CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 2)
                        xYYMM = Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 5, 2) + xOrderMonth
                    Else
                        If CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 2 > 12 Then
                            xOrderMonth = "0" + CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 2 - 12)
                            xYYMM = CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 5, 2)) + 1) + xOrderMonth
                        Else
                            xOrderMonth = CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 2)
                            xYYMM = Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 5, 2) + xOrderMonth
                        End If
                    End If
                    xSeqNo = 1
                    ' 製作 EDI 資料
                    xProdQty = dt_SPActionPlan2KP.Rows(i).Item("N2_Qty")
                    While xProdQty > 0
                        sql = "Insert into B_LS2EDIInterface "
                        sql &= "( "
                        sql &= "BuyerGroup, BULSNo, BULSSubNo, OrderMonth, Seqno, LSType, KeepCode, Item, ItemName1, ItemName2, ItemName3, Color, PartType, "
                        sql &= "Yobi1, Yobi2, Reference, LengthUnit, BuyerCode, ReqDate, ProdQty, CreateUser, CreateTime"
                        sql &= " ) "
                        sql &= "VALUES( "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Customer") & "', "
                        sql &= " '" & xLSNo & "', "
                        sql &= " " & dt_SPActionPlan2KP.Rows(i).Item("LSSubNo") & ", "
                        sql &= " '" & xOrderMonth & "', "
                        sql &= " " & CStr(xSeqNo) & ", "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ActionName") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Keep") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Item") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ItemName1") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ItemName2") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ItemName3") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Color") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("PartType") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("CustomerGr") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("N2_YOBI1") & "', "
                        sql &= " '" & "[" & dt_SPActionPlan2KP.Rows(i).Item("SPNO") & "]" + " N+2 BUY" & "', "
                        sql &= " '" & xLengthUnit & "', "
                        sql &= " '" & Mid(dt_SPActionPlan2KP.Rows(i).Item("Buyer"), 1, 6) & "', "
                        sql &= " " & Now.AddDays(xReqDate(3)).ToString("yyyyMMdd") & ", "

                        If xProdQty <= xMaxQty Then
                            sql &= " " & CStr(xProdQty) & ", "
                            xProdQty = 0
                        Else
                            sql &= " " & CStr(xMaxQty) & ", "
                            xProdQty = xProdQty - xMaxQty
                        End If
                        sql &= " '" & "SHOPPING" & "', "
                        sql &= " '" & NowDateTime & "' "
                        sql &= " ) "
                        oDataBase.ExecuteNonQuery(sql)

                        xSeqNo = xSeqNo + 1
                    End While
                End If
                ' N+3
                ' ------------------------
                If xOrderZone >= 3 Then
                    ' 準備資料 (OrderMonth, xYYMM, xSeqNo)
                    If CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 3 < 10 Then
                        xOrderMonth = "0" + CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 3)
                        xYYMM = Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 5, 2) + xOrderMonth
                    Else
                        If CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 3 > 12 Then
                            xOrderMonth = "0" + CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 3 - 12)
                            xYYMM = CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 5, 2)) + 1) + xOrderMonth
                        Else
                            xOrderMonth = CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 3)
                            xYYMM = Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 5, 2) + xOrderMonth
                        End If
                    End If
                    xSeqNo = 1
                    ' 製作 EDI 資料
                    xProdQty = dt_SPActionPlan2KP.Rows(i).Item("N3_Qty")
                    While xProdQty > 0
                        sql = "Insert into B_LS2EDIInterface "
                        sql &= "( "
                        sql &= "BuyerGroup, BULSNo, BULSSubNo, OrderMonth, Seqno, LSType, KeepCode, Item, ItemName1, ItemName2, ItemName3, Color, PartType, "
                        sql &= "Yobi1, Yobi2, Reference, LengthUnit, BuyerCode, ReqDate, ProdQty, CreateUser, CreateTime"
                        sql &= " ) "
                        sql &= "VALUES( "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Customer") & "', "
                        sql &= " '" & xLSNo & "', "
                        sql &= " " & dt_SPActionPlan2KP.Rows(i).Item("LSSubNo") & ", "
                        sql &= " '" & xOrderMonth & "', "
                        sql &= " " & CStr(xSeqNo) & ", "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ActionName") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Keep") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Item") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ItemName1") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ItemName2") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ItemName3") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Color") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("PartType") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("CustomerGr") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("N3_YOBI1") & "', "
                        sql &= " '" & "[" & dt_SPActionPlan2KP.Rows(i).Item("SPNO") & "]" + " N+3 BUY" & "', "
                        sql &= " '" & xLengthUnit & "', "
                        sql &= " '" & Mid(dt_SPActionPlan2KP.Rows(i).Item("Buyer"), 1, 6) & "', "
                        sql &= " " & Now.AddDays(xReqDate(4)).ToString("yyyyMMdd") & ", "

                        If xProdQty <= xMaxQty Then
                            sql &= " " & CStr(xProdQty) & ", "
                            xProdQty = 0
                        Else
                            sql &= " " & CStr(xMaxQty) & ", "
                            xProdQty = xProdQty - xMaxQty
                        End If
                        sql &= " '" & "SHOPPING" & "', "
                        sql &= " '" & NowDateTime & "' "
                        sql &= " ) "
                        oDataBase.ExecuteNonQuery(sql)

                        xSeqNo = xSeqNo + 1
                    End While
                End If
                ' N+4
                ' ------------------------
                If xOrderZone >= 4 Then
                    ' 準備資料 (OrderMonth, xYYMM, xSeqNo)
                    If CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 4 < 10 Then
                        xOrderMonth = "0" + CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 4)
                        xYYMM = Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 5, 2) + xOrderMonth
                    Else
                        If CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 4 > 12 Then
                            xOrderMonth = "0" + CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 4 - 12)
                            xYYMM = CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 5, 2)) + 1) + xOrderMonth
                        Else
                            xOrderMonth = CStr(CInt(Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 7, 2)) + 4)
                            xYYMM = Mid(dt_SPActionPlan2KP.Rows(i).Item("SPNO"), 5, 2) + xOrderMonth
                        End If
                    End If
                    xSeqNo = 1
                    ' 製作 EDI 資料
                    xProdQty = dt_SPActionPlan2KP.Rows(i).Item("N4_Qty")
                    While xProdQty > 0
                        sql = "Insert into B_LS2EDIInterface "
                        sql &= "( "
                        sql &= "BuyerGroup, BULSNo, BULSSubNo, OrderMonth, Seqno, LSType, KeepCode, Item, ItemName1, ItemName2, ItemName3, Color, PartType, "
                        sql &= "Yobi1, Yobi2, Reference, LengthUnit, BuyerCode, ReqDate, ProdQty, CreateUser, CreateTime"
                        sql &= " ) "
                        sql &= "VALUES( "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Customer") & "', "
                        sql &= " '" & xLSNo & "', "
                        sql &= " " & dt_SPActionPlan2KP.Rows(i).Item("LSSubNo") & ", "
                        sql &= " '" & xOrderMonth & "', "
                        sql &= " " & CStr(xSeqNo) & ", "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ActionName") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Keep") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Item") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ItemName1") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ItemName2") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("ItemName3") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("Color") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("PartType") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("CustomerGr") & "', "
                        sql &= " '" & dt_SPActionPlan2KP.Rows(i).Item("N4_YOBI1") & "', "
                        sql &= " '" & "[" & dt_SPActionPlan2KP.Rows(i).Item("SPNO") & "]" + " N+4 BUY" & "', "
                        sql &= " '" & xLengthUnit & "', "
                        sql &= " '" & Mid(dt_SPActionPlan2KP.Rows(i).Item("Buyer"), 1, 6) & "', "
                        sql &= " " & Now.AddDays(xReqDate(5)).ToString("yyyyMMdd") & ", "
                        If xProdQty <= xMaxQty Then
                            sql &= " " & CStr(xProdQty) & ", "
                            xProdQty = 0
                        Else
                            sql &= " " & CStr(xMaxQty) & ", "
                            xProdQty = xProdQty - xMaxQty
                        End If
                        sql &= " '" & "SHOPPING" & "', "
                        sql &= " '" & NowDateTime & "' "
                        sql &= " ) "
                        oDataBase.ExecuteNonQuery(sql)

                        xSeqNo = xSeqNo + 1
                    End While
                End If
                '
            Next
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "SPEDITransfer", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "SPEDITransfer", pUserID, "")
        End Try
        '
        Return RtnCode
        '
    End Function
    'SPKPInterface-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([945]-SPFinalPlan)
    '**     SHOPPING Plan Final
    '***********************************************************************************************
    'SPFinalPlan-Start
    <WebMethod()> _
    Public Function SPFinalPlan(ByVal pLogID As String, _
                                ByVal pBuyer As String, _
                                ByVal pUserID As String, _
                                ByVal pGRBuyer As String, _
                                ByVal pFunList As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            '
            'INSERT WINGS
            Dim i As Integer
            Dim sql As String
            Dim cnW As New OleDbConnection
            Dim cmdW As New OleDbCommand
            Dim ds As New DataSet
            '
            cnW.ConnectionString = ConnectString
            cmdW.Connection = cnW
            cnW.Open()
            '
            '判斷是否可 FINAL
            sql = "SELECT * "
            sql &= "FROM SPActionPlan "
            sql &= "Where Customer = '" & pBuyer & "' "
            sql &= "AND VERSION = 99 "
            sql &= "AND ACTION = 1 "
            sql &= "AND ( "
            sql &= "   ( (N_QTY>0 AND N_YOBI1='') OR (N1_QTY>0 AND N1_YOBI1='') OR (N2_QTY>0 AND N2_YOBI1='') OR (N3_QTY>0 AND N3_YOBI1='') OR (N4_QTY>0 AND N4_YOBI1='') ) "
            sql &= "   OR "
            sql &= "   ( N_YOBI1+'/'+N1_YOBI1+'/'+N2_YOBI1+'/'+N3_YOBI1+'/'+N4_YOBI1+'/' LIKE '%TBA%' ) "
            sql &= "   OR "
            sql &= "   ( N_YOBI1+'/'+N1_YOBI1+'/'+N2_YOBI1+'/'+N3_YOBI1+'/'+N4_YOBI1+'/' LIKE '%KP%' ) "
            sql &= ") "
            Dim dt_ActionPlanErr As DataTable = oDataBase.GetDataTable(sql)
            If dt_ActionPlanErr.Rows.Count <= 0 Then
                '
                ' FINAL=OK
                sql = "SELECT "
                sql &= "CustomerGr, Item, Color, Keep "
                sql &= "FROM SPActionPlan "
                sql &= "Where Customer = '" & pBuyer & "' "
                sql &= "AND VERSION = 99 "
                sql &= "Group by CustomerGr, Item, Color, Keep "
                sql &= "Order by CustomerGr, Item, Color, Keep "
                Dim dt_ActionPlan As DataTable = oDataBase.GetDataTable(sql)
                For i = 0 To dt_ActionPlan.Rows.Count - 1
                    '
                    'ITEM -------------------------------
                    sql = "DELETE FROM TWNDLIB.SPITEM100X "
                    sql &= "Where UIDC10 = '" & pBuyer & "' "
                    cmdW.CommandText = sql
                    cmdW.ExecuteNonQuery()
                    '
                    sql = "INSERT INTO TWNDLIB.SPITEM100X "
                    sql &= "( "
                    sql &= "CSTC10, ITMC10, CLRC10, KEPC10, "
                    sql &= "SCAQ10,ONAQ10,FRQ110,KPQ110, "
                    sql &= "N21I10,N21O10,N22I10,N22O10, "
                    sql &= "N23I10,N23O10,N24I10,N24O10, "
                    sql &= "N11I10,N11O10,N12I10,N12O10, "
                    sql &= "N13I10,N13O10,N14I10,N14O10, "
                    sql &= "N01I10,N01O10,N02I10,N02O10, "
                    sql &= "N03I10,N03O10,N04I10,N04O10, "
                    sql &= "UIDC10,PRGC10,DEVC10, "
                    sql &= "RADU10, RADT10, RUPU10, RUPT10 "
                    sql &= " ) "
                    sql &= "VALUES( "
                    sql &= " '" & dt_ActionPlan.Rows(i).Item("CustomerGr").ToString & "', "
                    sql &= " '" & dt_ActionPlan.Rows(i).Item("Item").ToString & "', "
                    sql &= " '" & dt_ActionPlan.Rows(i).Item("Color").ToString & "', "
                    sql &= " '" & dt_ActionPlan.Rows(i).Item("Keep").ToString & "', "
                    '
                    sql &= " 0, 0, 0, 0, "
                    sql &= " 0, 0, 0, 0, "
                    sql &= " 0, 0, 0, 0, "
                    sql &= " 0, 0, 0, 0, "
                    sql &= " 0, 0, 0, 0, "
                    sql &= " 0, 0, 0, 0, "
                    sql &= " 0, 0, 0, 0, "
                    '
                    If Len(pBuyer) > 10 Then
                        sql &= " '" & Mid(pBuyer, 1, 10) & "', "
                        sql &= " '" & Mid(pBuyer, 11, 5) & "', "
                    Else
                        sql &= " '" & pBuyer & "', "
                        sql &= " '" & "" & "', "
                    End If
                    '
                    sql &= " '', "
                    sql &= " " & Now.ToString("yyyyMMdd") & ", "
                    sql &= " " & Now.ToString("HHmmss") & ", "
                    sql &= " 0, 0 "
                    sql &= " ) "
                    '
                    cmdW.CommandText = sql
                    cmdW.ExecuteNonQuery()
                    '----
                    '
                    'ALL ITEM -------------------------------
                    sql = "SELECT ITMC10 FROM TWNDLIB.SPITEM100 "
                    sql &= "Where ITMC10 = '" & dt_ActionPlan.Rows(i).Item("Item").ToString & "' "
                    sql &= "  And CLRC10 = '" & dt_ActionPlan.Rows(i).Item("Color").ToString & "' "
                    sql &= "  And KEPC10 = '" & dt_ActionPlan.Rows(i).Item("Keep").ToString & "' "
                    Dim DBAdapter1 As New OleDbDataAdapter(sql, cnW)
                    DBAdapter1.Fill(ds, "SPITEM")
                    If ds.Tables("SPITEM").Rows.Count <= 0 Then
                        '
                        sql = "INSERT INTO TWNDLIB.SPITEM100 "
                        sql &= "( "
                        sql &= "CSTC10, ITMC10, CLRC10, KEPC10, "
                        sql &= "SCAQ10,ONAQ10,FRQ110,KPQ110, "
                        sql &= "N21I10,N21O10,N22I10,N22O10, "
                        sql &= "N23I10,N23O10,N24I10,N24O10, "
                        sql &= "N11I10,N11O10,N12I10,N12O10, "
                        sql &= "N13I10,N13O10,N14I10,N14O10, "
                        sql &= "N01I10,N01O10,N02I10,N02O10, "
                        sql &= "N03I10,N03O10,N04I10,N04O10, "
                        sql &= "UIDC10,PRGC10,DEVC10, "
                        sql &= "RADU10, RADT10, RUPU10, RUPT10 "
                        sql &= " ) "
                        sql &= "VALUES( "
                        sql &= " '" & dt_ActionPlan.Rows(i).Item("CustomerGr").ToString & "', "
                        sql &= " '" & dt_ActionPlan.Rows(i).Item("Item").ToString & "', "
                        sql &= " '" & dt_ActionPlan.Rows(i).Item("Color").ToString & "', "
                        sql &= " '" & dt_ActionPlan.Rows(i).Item("Keep").ToString & "', "
                        '
                        sql &= " 0, 0, 0, 0, "
                        sql &= " 0, 0, 0, 0, "
                        sql &= " 0, 0, 0, 0, "
                        sql &= " 0, 0, 0, 0, "
                        sql &= " 0, 0, 0, 0, "
                        sql &= " 0, 0, 0, 0, "
                        sql &= " 0, 0, 0, 0, "
                        '
                        If Len(pBuyer) > 10 Then
                            sql &= " '" & Mid(pBuyer, 1, 10) & "', "
                            sql &= " '" & Mid(pBuyer, 11, 5) & "', "
                        Else
                            sql &= " '" & pBuyer & "', "
                            sql &= " '" & "" & "', "
                        End If
                        '
                        sql &= " '', "
                        sql &= " " & Now.ToString("yyyyMMdd") & ", "
                        sql &= " " & Now.ToString("HHmmss") & ", "
                        sql &= " 0, 0 "
                        sql &= " ) "
                        '
                        cmdW.CommandText = sql
                        cmdW.ExecuteNonQuery()
                        '
                    End If
                    '
                Next
                '
                ' ------------
                'FINAL PROC
                Dim Conn As New SqlClient.SqlConnection("Data Source=" & "10.245.0.112" & ";" & _
                                                        "Initial Catalog=" & "EDI" & ";" & _
                                                        "User ID=" & "sa" & ";" & _
                                                        "Password=;")
                Dim cmd As SqlCommand = Conn.CreateCommand
                '
                Conn.Open()
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandTimeout = 180
                cmd.Parameters.Add(New SqlParameter("@Customer", pBuyer))
                cmd.CommandText = "sp_LFSPFinnalPlan"
                cmd.ExecuteNonQuery()
                Conn.Close()
                '
            Else
                ' FINAL=NG
                RtnCode = 1
            End If
            '
            cnW.Close()
            '
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'SPFinalPlan-End
    '***********************************************************************************************
    '**([950]-SPPendingPlan)
    '**     SHOPPING Plan Pending
    '***********************************************************************************************
    'SPPendingPlan-Start
    <WebMethod()> _
    Public Function SPPendingPlan(ByVal pLogID As String, _
                                  ByVal pBuyer As String, _
                                  ByVal pUserID As String, _
                                  ByVal pGRBuyer As String, _
                                  ByVal pFunList As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            '
            Dim sql As String
            '
            '判斷是否可 FINAL
            sql = "SELECT * "
            sql &= "FROM SPActionPlan "
            sql &= "Where Customer = '" & pBuyer & "' "
            sql &= "AND VERSION = 99 "
            sql &= "AND ACTION = 1 "
            sql &= "AND ( "
            sql &= "      ( (N_QTY>0 AND N_YOBI1='') OR (N1_QTY>0 AND N1_YOBI1='') OR (N2_QTY>0 AND N2_YOBI1='') OR (N3_QTY>0 AND N3_YOBI1='') OR (N4_QTY>0 AND N4_YOBI1='') ) "
            sql &= "      OR "
            sql &= "      ( "
            sql &= "        ( N_YOBI1+N1_YOBI1+N2_YOBI1+N3_YOBI1+N4_YOBI1 <> '' ) "
            sql &= "        AND "
            sql &= "        ( N_YOBI1+'/'+N1_YOBI1+'/'+N2_YOBI1+'/'+N3_YOBI1+'/'+N4_YOBI1+'/' NOT LIKE '%OR%' ) "
            sql &= "        AND "
            sql &= "        ( N_YOBI1+'/'+N1_YOBI1+'/'+N2_YOBI1+'/'+N3_YOBI1+'/'+N4_YOBI1+'/' NOT LIKE '%KP%' ) "
            sql &= "      ) "
            sql &= ") "
            '
            Dim dt_ActionPlanErr As DataTable = oDataBase.GetDataTable(sql)
            If dt_ActionPlanErr.Rows.Count > 0 Then
                ' PENDING=NG
                RtnCode = 1
            Else
                ' PENDING=OK
                Dim Conn As New SqlClient.SqlConnection("Data Source=" & "10.245.0.112" & ";" & _
                                                        "Initial Catalog=" & "EDI" & ";" & _
                                                        "User ID=" & "sa" & ";" & _
                                                        "Password=;")
                Dim cmd As SqlCommand = Conn.CreateCommand
                '
                Conn.Open()
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandTimeout = 180
                cmd.Parameters.Add(New SqlParameter("@Customer", pBuyer))
                cmd.CommandText = "sp_LFSPPendingPlan"
                cmd.ExecuteNonQuery()
                Conn.Close()
                '
            End If
            '
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'SPPendingPlan-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([955]-SPAutoFinalPlan)
    '**     SHOPPING Plan Pending Auto Final
    '***********************************************************************************************
    'SPAutoFinalPlan-Start
    <WebMethod()> _
    Public Function SPAutoFinalPlan(ByVal pLogID As String, _
                                    ByVal pBuyer As String, _
                                    ByVal pUserID As String, _
                                    ByVal pGRBuyer As String, _
                                    ByVal pFunList As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            '
            'INSERT WINGS
            Dim i, j As Integer
            Dim sql As String
            Dim cnW As New OleDbConnection
            Dim cmdW As New OleDbCommand
            Dim ds As New DataSet
            '
            cnW.ConnectionString = ConnectString
            cmdW.Connection = cnW
            cnW.Open()
            '
            '判斷是否可 FINAL
            sql = "SELECT "
            sql &= "Customer, SPNo "
            sql &= "FROM V_PDSPFinalList "
            sql &= "Where FinalFlag = 1 "
            sql &= "order by Customer, SPNo "
            Dim dt_FinalList As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_FinalList.Rows.Count - 1
                '
                ' FINAL=OK
                sql = "SELECT "
                sql = sql & "CustomerGr, Item, Color, Keep "
                sql = sql & "from PD_SPActionPlan "
                sql = sql & "where Customer = '" & dt_FinalList.Rows(i).Item("Customer").ToString & "' "
                sql = sql & "  and SPNo = '" & dt_FinalList.Rows(i).Item("SPNo").ToString & "' "
                sql = sql & "  and VERSION = 99 "
                sql = sql & "Group by CustomerGr, Item, Color, Keep "
                sql = sql & "Order by CustomerGr, Item, Color, Keep "
                Dim dt_ActionPlan As DataTable = oDataBase.GetDataTable(sql)
                For j = 0 To dt_ActionPlan.Rows.Count - 1
                    '
                    'ALL ITEM -------------------------------
                    sql = "SELECT ITMC10 FROM TWNDLIB.SPITEM100 "
                    sql &= "Where ITMC10 = '" & dt_ActionPlan.Rows(j).Item("Item").ToString & "' "
                    sql &= "  And CLRC10 = '" & dt_ActionPlan.Rows(j).Item("Color").ToString & "' "
                    sql &= "  And KEPC10 = '" & dt_ActionPlan.Rows(j).Item("Keep").ToString & "' "
                    Dim DBAdapter1 As New OleDbDataAdapter(sql, cnW)
                    DBAdapter1.Fill(ds, "SPITEM")
                    If ds.Tables("SPITEM").Rows.Count <= 0 Then
                        '
                        sql = "INSERT INTO TWNDLIB.SPITEM100 "
                        sql &= "( "
                        sql &= "CSTC10, ITMC10, CLRC10, KEPC10, "
                        sql &= "SCAQ10,ONAQ10,FRQ110,KPQ110, "
                        sql &= "N21I10,N21O10,N22I10,N22O10, "
                        sql &= "N23I10,N23O10,N24I10,N24O10, "
                        sql &= "N11I10,N11O10,N12I10,N12O10, "
                        sql &= "N13I10,N13O10,N14I10,N14O10, "
                        sql &= "N01I10,N01O10,N02I10,N02O10, "
                        sql &= "N03I10,N03O10,N04I10,N04O10, "
                        sql &= "UIDC10,PRGC10,DEVC10, "
                        sql &= "RADU10, RADT10, RUPU10, RUPT10 "
                        sql &= " ) "
                        sql &= "VALUES( "
                        sql &= " '" & dt_ActionPlan.Rows(j).Item("CustomerGr").ToString & "', "
                        sql &= " '" & dt_ActionPlan.Rows(j).Item("Item").ToString & "', "
                        sql &= " '" & dt_ActionPlan.Rows(j).Item("Color").ToString & "', "
                        sql &= " '" & dt_ActionPlan.Rows(j).Item("Keep").ToString & "', "
                        '
                        sql &= " 0, 0, 0, 0, "
                        sql &= " 0, 0, 0, 0, "
                        sql &= " 0, 0, 0, 0, "
                        sql &= " 0, 0, 0, 0, "
                        sql &= " 0, 0, 0, 0, "
                        sql &= " 0, 0, 0, 0, "
                        sql &= " 0, 0, 0, 0, "
                        '
                        If Len(dt_FinalList.Rows(i).Item("Customer").ToString) > 10 Then
                            sql &= " '" & Mid(dt_FinalList.Rows(i).Item("Customer").ToString, 1, 10) & "', "
                            sql &= " '" & Mid(dt_FinalList.Rows(i).Item("Customer").ToString, 11, 5) & "', "
                        Else
                            sql &= " '" & dt_FinalList.Rows(i).Item("Customer").ToString & "', "
                            sql &= " '" & "" & "', "
                        End If
                        '
                        sql &= " 'AUTO', "
                        sql &= " " & Now.ToString("yyyyMMdd") & ", "
                        sql &= " " & Now.ToString("HHmmss") & ", "
                        sql &= " 0, 0 "
                        sql &= " ) "
                        '
                        cmdW.CommandText = sql
                        cmdW.ExecuteNonQuery()
                        '
                    End If
                    '
                Next
                '
                'FINAL PROC
                Dim Conn As New SqlClient.SqlConnection("Data Source=" & "10.245.0.112" & ";" & _
                                                        "Initial Catalog=" & "EDI" & ";" & _
                                                        "User ID=" & "sa" & ";" & _
                                                        "Password=;")
                Dim cmd As SqlCommand = Conn.CreateCommand
                '
                Conn.Open()
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandTimeout = 600
                cmd.Parameters.Add(New SqlParameter("@Customer", dt_FinalList.Rows(i).Item("Customer").ToString))
                cmd.Parameters.Add(New SqlParameter("@SPNo", dt_FinalList.Rows(i).Item("SPNo").ToString))
                cmd.CommandText = "sp_PDLFSPFinnalPlan"
                cmd.ExecuteNonQuery()
                Conn.Close()
            Next
            '
            cnW.Close()
            '
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'SPAutoFinalPlan-End
    'joy
    '***********************************************************************************************
    '
    '***********************************************************************************************
    '**([100]-MakeForcastNo)
    '**       建置 Forcast No 
    '***********************************************************************************************
    'MakeForcastNo-Start
    <WebMethod()> _
    Public Function MakeForcastNo(ByVal pLogID As String, _
                                  ByVal pBuyer As String, _
                                  ByVal pUserID As String, _
                                  ByVal pGRBuyer As String, _
                                  ByVal pFunList As String, _
                                  ByVal pLastUniqueID As Integer) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            ' ***********************************************************************************
            ' 變數定義及設定初值
            ' ***********************************************************************************
            Dim sql As String
            Dim xFCTNo, xSeqNoString, xGroup As String
            Dim xYFCTNo, xYSeqNoString As String
            Dim xSeqNo, xYSeqNo, i As Integer
            xGroup = "XX"
            ' ***********************************************************************************
            ' 取得客戶FCT FCTNO = GroupCode + "F"  + 年月 + 流水號(4碼)   SA + F  + 136 + 0001
            ' 取得YKK FCT FCTNO = GroupCode + "YF" + 年月 + 流水號(3碼)   SA + AF + 136 + 001
            ' ***********************************************************************************
            ' 取得GroupCode
            sql = "SELECT GroupCode FROM M_FControlRecord "
            sql &= "Where Buyer = '" & pBuyer & "' "
            Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
            If dt_ControlRecord.Rows.Count > 0 Then xGroup = dt_ControlRecord.Rows(0).Item("GroupCode")
            ' 取得年月 
            Select Case Month(Now)
                Case 10
                    xFCTNo = xGroup + "F" + CStr(Year(Now) - 2000) + "A"
                    xYFCTNo = xGroup + "YF" + CStr(Year(Now) - 2000) + "A"
                Case 11
                    xFCTNo = xGroup + "F" + CStr(Year(Now) - 2000) + "B"
                    xYFCTNo = xGroup + "YF" + CStr(Year(Now) - 2000) + "B"
                Case 12
                    xFCTNo = xGroup + "F" + CStr(Year(Now) - 2000) + "C"
                    xYFCTNo = xGroup + "YF" + CStr(Year(Now) - 2000) + "C"
                Case Else
                    xFCTNo = xGroup + "F" + CStr(Year(Now) - 2000) + CStr(Month(Now))
                    xYFCTNo = xGroup + "YF" + CStr(Year(Now) - 2000) + CStr(Month(Now))
            End Select
            ' 流水號(下一可使用No)
            xSeqNo = GetFCTSeqNo(pBuyer, xFCTNo, "C")
            xYSeqNo = GetFCTSeqNo(pBuyer, xFCTNo, "Y")
            '
            ' ***********************************************************************************
            ' 更新 Forcast Plan FCTNo
            ' ***********************************************************************************
            sql = "SELECT Unique_ID, FCTNo FROM ForcastPlan "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And Unique_ID > " & pLastUniqueID & " "
            sql &= "  And Y_Level = 0 "
            sql &= "Order by Unique_ID "
            Dim dt_FCT As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_FCT.Rows.Count - 1
                '
                ' 判斷客戶FCT or YKK FCT
                If dt_FCT.Rows(i).Item("FCTNo").ToString = "" Then
                    ' 客戶FCT
                    ' 製作SeqNo(不足5位補0)
                    xSeqNoString = CStr(xSeqNo)
                    If Len(xSeqNoString) < 2 Then
                        xSeqNoString = "0000" + CStr(xSeqNo)
                    Else
                        If Len(xSeqNoString) < 3 Then
                            xSeqNoString = "000" + CStr(xSeqNo)
                        Else
                            If Len(xSeqNoString) < 4 Then
                                xSeqNoString = "00" + CStr(xSeqNo)
                            Else
                                If Len(xSeqNoString) < 5 Then
                                    xSeqNoString = "0" + CStr(xSeqNo)
                                End If
                            End If
                        End If
                    End If
                    '
                    ' 更新 Forcast Plan
                    sql = "Update ForcastPlan Set "
                    sql = sql + " FCTNo = '" & xFCTNo + xSeqNoString & "' "
                    sql = sql + " Where Unique_ID = '" & dt_FCT.Rows(i).Item("Unique_ID").ToString & "' "
                    oDataBase.ExecuteNonQuery(sql)
                    '
                    xSeqNo = xSeqNo + 1
                Else
                    ' YKK FCT
                    ' 製作SeqNo(不足3位補0)
                    xYSeqNoString = CStr(xYSeqNo)
                    If Len(xYSeqNoString) < 2 Then
                        xYSeqNoString = "00" + CStr(xYSeqNo)
                    Else
                        If Len(xYSeqNoString) < 3 Then
                            xYSeqNoString = "0" + CStr(xYSeqNo)
                        End If
                    End If
                    '
                    ' 更新 Forcast Plan
                    sql = "Update ForcastPlan Set "
                    sql = sql + " FCTNo = '" & xYFCTNo + xYSeqNoString & "' "
                    sql = sql + " Where Unique_ID = '" & dt_FCT.Rows(i).Item("Unique_ID").ToString & "' "
                    oDataBase.ExecuteNonQuery(sql)
                    '
                    xYSeqNo = xYSeqNo + 1
                End If
            Next
            '
            ' ***********************************************************************************
            ' 更新下一個可使用No
            ' ***********************************************************************************
            ' 客戶FCT
            ' 製作SeqNo(不足4位補0)
            xSeqNoString = CStr(xSeqNo)
            If Len(xSeqNoString) < 2 Then
                xSeqNoString = "0000" + CStr(xSeqNo)
            Else
                If Len(xSeqNoString) < 3 Then
                    xSeqNoString = "000" + CStr(xSeqNo)
                Else
                    If Len(xSeqNoString) < 4 Then
                        xSeqNoString = "00" + CStr(xSeqNo)
                    Else
                        If Len(xSeqNoString) < 5 Then
                            xSeqNoString = "0" + CStr(xSeqNo)
                        End If
                    End If
                End If
            End If
            ' 更新下一次可使用PONO
            UpdateNextFCTNo(pBuyer, xFCTNo + xSeqNoString, "C")
            ' ---------------------------------------------------
            ' YKK FCT
            ' 製作SeqNo(不足3位補0)
            xYSeqNoString = CStr(xYSeqNo)
            If Len(xYSeqNoString) < 2 Then
                xYSeqNoString = "00" + CStr(xYSeqNo)
            Else
                If Len(xYSeqNoString) < 3 Then
                    xYSeqNoString = "0" + CStr(xYSeqNo)
                End If
            End If
            ' 更新下一次可使用PONO
            UpdateNextFCTNo(pBuyer, xYFCTNo + xYSeqNoString, "Y")
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "FCTPLAN", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "MakeForcastNo", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'MakeForcastNo-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([110]-ForcastPlan)  (Step-1)
    '**     Forcast Plan展開 
    '***********************************************************************************************
    'ForcastPlan-Start
    <WebMethod()> _
    Public Function ForcastPlan(ByVal pLogID As String, _
                                ByVal pBuyer As String, _
                                ByVal pUserID As String, _
                                ByVal pGRBuyer As String, _
                                ByVal pFunList As String, _
                                ByVal pLastUniqueID As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        '
        Try
            ' ***********************************************************************************
            ' 變數定義及設定初值
            ' ***********************************************************************************
            Dim xClass, xQty As String
            Dim pItem1(10), pItem2(10), pItem3(10) As String
            Dim pItemName1(10), pItemName2(10), pItemName3(10) As String
            Dim pQty1(10), pQty2(10), pQty3(10) As String
            Dim pCount1, pCount2, pCount3 As Integer
            Dim i, j, idx1, idx2, idx3, SubNo As Integer

            oWaves.Timeout = Timeout.Infinite

            ' ***********************************************************************************
            ' Wave's BOM展開
            ' ***********************************************************************************
            sql = "SELECT Y_ItemCode FROM ForcastPlan "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And Unique_ID > " & pLastUniqueID & " "
            sql &= "  And Y_Level = 0 "
            sql &= "Group by Y_ItemCode "
            sql &= "Order by Y_ItemCode "
            Dim dt_FCTItem As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_FCTItem.Rows.Count - 1
                '
                ' 取得Meter換算基準 (取得ItemClass)
                xClass = ""
                oWaves.GetItemClass(dt_FCTItem.Rows(i).Item("Y_ItemCode"), xClass)
                ' 
                sql = "Select * From M_Referp "
                sql = sql & " Where Cat  = '" & "110" & "' "
                sql = sql & "   And DKey = '" & pBuyer + "-" + xClass & "' "
                Dim dt_Referp As DataTable = oDataBase.GetDataTable(sql)
                If dt_Referp.Rows.Count > 0 Then
                    xQty = CStr(CDbl(dt_Referp.Rows(0).Item("Data")) * 100)
                Else
                    xQty = "100"
                End If
                '
                SubNo = 2
                '
                ' -----------------------------------------
                ' BOM 展開
                ' -----------------------------------------
                '
                ' Level 1 展開 (CHAIN)
                oWaves.GetItemStructure("01", dt_FCTItem.Rows(i).Item("Y_ItemCode"), "CH", pItem1, pItemName1, pQty1, pCount1)

                If pCount1 > 0 Then
                    For idx1 = 1 To pCount1
                        '
                        sql = "SELECT * FROM ForcastPlan "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                        sql &= "  And Y_Level = 0 "
                        sql &= "  And Y_ItemCode = " & dt_FCTItem.Rows(i).Item("Y_ItemCode") & " "
                        sql &= "Order by FCTNo, FCTSubNo "
                        Dim dt_FCT1 As DataTable = oDataBase.GetDataTable(sql)
                        For j = 0 To dt_FCT1.Rows.Count - 1
                            WriteFCTPlan(dt_FCT1.Rows(j).Item("Unique_ID"), SubNo, 1, dt_FCTItem.Rows(i).Item("Y_ItemCode"), pItem1(idx1), dt_FCT1.Rows(j).Item("Y_Color").ToString, "CH", xQty)
                        Next
                        SubNo = SubNo + 1
                        '
                        ' Level 2 展開 (CHAIN) 
                        oWaves.GetItemStructure("01", pItem1(idx1), "CH", pItem2, pItemName2, pQty2, pCount2)
                        For idx2 = 1 To pCount2
                            '
                            sql = "SELECT * FROM ForcastPlan "
                            sql &= "Where Buyer = '" & pBuyer & "' "
                            sql &= "  And Unique_ID > " & pLastUniqueID & " "
                            sql &= "  And Y_Level = 0 "
                            sql &= "  And Y_ItemCode = " & dt_FCTItem.Rows(i).Item("Y_ItemCode") & " "
                            sql &= "Order by FCTNo, FCTSubNo "
                            Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                            For j = 0 To dt_FCT2.Rows.Count - 1
                                WriteFCTPlan(dt_FCT2.Rows(j).Item("Unique_ID"), SubNo, 2, dt_FCTItem.Rows(i).Item("Y_ItemCode"), pItem2(idx2), dt_FCT2.Rows(j).Item("Y_Color").ToString, "CH", xQty)
                            Next
                            SubNo = SubNo + 1
                            '
                            ' Level 3 展開 (CHAIN)
                            oWaves.GetItemStructure("01", pItem2(idx2), "CH", pItem3, pItemName3, pQty3, pCount3)
                            For idx3 = 1 To pCount3
                                '
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And Y_ItemCode = " & dt_FCTItem.Rows(i).Item("Y_ItemCode") & " "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT3 As DataTable = oDataBase.GetDataTable(sql)
                                For j = 0 To dt_FCT3.Rows.Count - 1
                                    WriteFCTPlan(dt_FCT3.Rows(j).Item("Unique_ID"), SubNo, 3, dt_FCTItem.Rows(i).Item("Y_ItemCode"), pItem3(idx3), dt_FCT3.Rows(j).Item("Y_Color").ToString, "CH", xQty)
                                Next
                                SubNo = SubNo + 1
                            Next
                        Next
                    Next
                Else
                    ' MF
                    ' GAP CHAIN 對應
                    oWaves.GetItemStructure("01", dt_FCTItem.Rows(i).Item("Y_ItemCode"), "GAP", pItem1, pItemName1, pQty1, pCount1)
                    ' Level 1 展開 (TAPE)
                    oWaves.GetItemStructure("01", pItem1(1), "T", pItem1, pItemName1, pQty1, pCount1)
                    For idx1 = 1 To pCount1
                        '
                        sql = "SELECT * FROM ForcastPlan "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                        sql &= "  And Y_Level = 0 "
                        sql &= "  And Y_ItemCode = " & dt_FCTItem.Rows(i).Item("Y_ItemCode") & " "
                        sql &= "Order by FCTNo, FCTSubNo "
                        Dim dt_FCT1 As DataTable = oDataBase.GetDataTable(sql)
                        For j = 0 To dt_FCT1.Rows.Count - 1
                            WriteFCTPlan(dt_FCT1.Rows(j).Item("Unique_ID"), SubNo, 1, dt_FCTItem.Rows(i).Item("Y_ItemCode"), pItem1(idx1), dt_FCT1.Rows(j).Item("Y_Color").ToString, "T", xQty)
                        Next
                        SubNo = SubNo + 1
                        '
                        ' Level 2 展開 (TAPE) 
                        oWaves.GetItemStructure("01", pItem1(idx1), "T", pItem2, pItemName2, pQty2, pCount2)
                        For idx2 = 1 To pCount2
                            '
                            sql = "SELECT * FROM ForcastPlan "
                            sql &= "Where Buyer = '" & pBuyer & "' "
                            sql &= "  And Unique_ID > " & pLastUniqueID & " "
                            sql &= "  And Y_Level = 0 "
                            sql &= "  And Y_ItemCode = " & dt_FCTItem.Rows(i).Item("Y_ItemCode") & " "
                            sql &= "Order by FCTNo, FCTSubNo "
                            Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                            For j = 0 To dt_FCT2.Rows.Count - 1
                                WriteFCTPlan(dt_FCT2.Rows(j).Item("Unique_ID"), SubNo, 2, dt_FCTItem.Rows(i).Item("Y_ItemCode"), pItem2(idx2), dt_FCT2.Rows(j).Item("Y_Color").ToString, "T", xQty)
                            Next
                            SubNo = SubNo + 1
                            '
                            ' Level 3 展開 (TAPE)
                            oWaves.GetItemStructure("01", pItem2(idx2), "T", pItem3, pItemName3, pQty3, pCount3)
                            For idx3 = 1 To pCount3
                                '
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And Y_ItemCode = " & dt_FCTItem.Rows(i).Item("Y_ItemCode") & " "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT3 As DataTable = oDataBase.GetDataTable(sql)
                                For j = 0 To dt_FCT3.Rows.Count - 1
                                    WriteFCTPlan(dt_FCT3.Rows(j).Item("Unique_ID"), SubNo, 3, dt_FCTItem.Rows(i).Item("Y_ItemCode"), pItem3(idx3), dt_FCT3.Rows(j).Item("Y_Color").ToString, "T", xQty)
                                Next
                                SubNo = SubNo + 1
                            Next
                        Next
                    Next
                End If
                ' -----------------------------------------
                ' BOM SLIDER(PS) 展開
                ' -----------------------------------------
                ' 取得換算基準
                xQty = "100"
                ' 
                ' Level 1 展開
                oWaves.GetItemStructure("01", dt_FCTItem.Rows(i).Item("Y_ItemCode"), "PS", pItem1, pItemName1, pQty1, pCount1)
                For idx1 = 1 To pCount1
                    '
                    sql = "SELECT * FROM ForcastPlan "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "  And Unique_ID > " & pLastUniqueID & " "
                    sql &= "  And Y_Level = 0 "
                    sql &= "  And Y_ItemCode = " & dt_FCTItem.Rows(i).Item("Y_ItemCode") & " "
                    sql &= "Order by FCTNo, FCTSubNo "
                    Dim dt_FCT1 As DataTable = oDataBase.GetDataTable(sql)
                    For j = 0 To dt_FCT1.Rows.Count - 1
                        WriteFCTPlan(dt_FCT1.Rows(j).Item("Unique_ID"), SubNo, 1, dt_FCTItem.Rows(i).Item("Y_ItemCode"), pItem1(idx1), dt_FCT1.Rows(j).Item("Y_Color").ToString, "PS", xQty)
                    Next
                    SubNo = SubNo + 1
                    '
                    ' Level 2 展開 
                    oWaves.GetItemStructure("01", pItem1(idx1), "PS", pItem2, pItemName2, pQty2, pCount2)
                    For idx2 = 1 To pCount2
                        '
                        sql = "SELECT * FROM ForcastPlan "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                        sql &= "  And Y_Level = 0 "
                        sql &= "  And Y_ItemCode = " & dt_FCTItem.Rows(i).Item("Y_ItemCode") & " "
                        sql &= "Order by FCTNo, FCTSubNo "
                        Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                        For j = 0 To dt_FCT2.Rows.Count - 1
                            WriteFCTPlan(dt_FCT2.Rows(j).Item("Unique_ID"), SubNo, 2, dt_FCTItem.Rows(i).Item("Y_ItemCode"), pItem2(idx2), dt_FCT2.Rows(j).Item("Y_Color").ToString, "PS", xQty)
                        Next
                        SubNo = SubNo + 1
                    Next
                Next
            Next
            '
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "FCTPLAN", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "ForcastPlan", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'ForcastPlan-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([112]-NewForcastPlan)   (Step-2)
    '**     New Forcast Plan展開 
    '***********************************************************************************************
    'NewForcastPlan-Start
    <WebMethod()> _
    Public Function NewForcastPlan(ByVal pLogID As String, _
                                   ByVal pBuyer As String, _
                                   ByVal pUserID As String, _
                                   ByVal pGRBuyer As String, _
                                   ByVal pFunList As String, _
                                   ByVal pLastUniqueID As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        '
        Try
            '##
            'MsgBox("IN")   
            '
            ' 變數定義及設定初值
            oWaves.Timeout = Timeout.Infinite
            Dim xItemProduct, xLTType, xPartType, xClass, xQtyMeter, xColor As String
            Dim xSubNo As Integer
            Dim FCTWrite As Boolean = False
            ' 構成展開專用變數
            Dim xItem(5), xItemName(5), xQty(5) As String
            Dim xItem1(5), xItemName1(5), xQty1(5) As String
            Dim xCount, xCount1 As Integer
            ' SearchItem專用變數
            Dim xSearchItem(10) As String
            ' 過濾箱中取得的資料變數
            Dim xRuleNo(5), xAction(5), xObjectType(5), xObjectProduct(5) As String
            Dim xRuleCount As Integer
            '
            ' 讀取-ForcastPlan Data
            sql = "SELECT C_Season As Season, Y_ItemCode AS Item, Y_Color AS Color, C_Color AS CSTColor, C_ShortenLT AS KeepCode FROM ForcastPlan "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And Unique_ID > " & pLastUniqueID & " "
            sql &= "  And Y_Level = 0 "
            sql &= "Group by C_Season, Y_ItemCode, Y_Color, C_Color, C_ShortenLT "
            sql &= "Order by C_Season, Y_ItemCode, Y_Color, C_Color, C_ShortenLT "
            Dim dt_FCTItem As DataTable = oDataBase.GetDataTable(sql)
            For i As Integer = 0 To dt_FCTItem.Rows.Count - 1
                FCTWrite = False
                xSubNo = 2
                ' 決定LT種類(LTType)
                xLTType = "LLT"
                If pBuyer = "FALL-000001" Or pBuyer = "FALL-000016" Then
                    '
                    'MODIFY-START ADIDAS SPEED 240627
                    'If dt_FCTItem.Rows(i).Item("KeepCode") <> "AD-RB" Then
                    'xLTType = "SLT"
                    'End If
                    If dt_FCTItem.Rows(i).Item("KeepCode") <> "TP4ADCLA-1" Then
                        xLTType = "SLT"
                    End If
                    'MODIFY-END
                    '
                End If
                ' 決定CST PULLER COLOR
                Dim xCSTPullerColor As String = ""
                If pBuyer = "FALL-000001" Or pBuyer = "FALL-000016" Then
                    ' ADIDAS / REEBOK
                    xCSTPullerColor = Mid(dt_FCTItem.Rows(i).Item("CSTColor"), InStr(dt_FCTItem.Rows(i).Item("CSTColor"), "/") + 1)
                    xCSTPullerColor = Mid(xCSTPullerColor, InStr(xCSTPullerColor, "/") + 1)
                    xCSTPullerColor = Mid(xCSTPullerColor, 1, 4)
                Else
                    ' Other BUYER
                    xCSTPullerColor = Mid(dt_FCTItem.Rows(i).Item("CSTColor"), InStr(dt_FCTItem.Rows(i).Item("CSTColor"), "/") + 1)
                    xCSTPullerColor = Mid(xCSTPullerColor, InStr(xCSTPullerColor, "/") + 1)
                    xCSTPullerColor = Mid(xCSTPullerColor, 1, InStr(xCSTPullerColor, "/") - 1)
                End If
                '
                ' CHAIN 處理 (不需CHECK LINE-LINE ITEM)
                '--------------------------------------------------------------------------------
                ' 取得Meter換算基準 (取得ItemClass)
                oWaves.GetItemClass(dt_FCTItem.Rows(i).Item("Item"), xClass)
                sql = "Select * From M_Referp "
                sql = sql & " Where Cat  = '" & "110" & "' "
                sql = sql & "   And DKey = '" & pBuyer + "-" + xClass & "' "
                Dim dt_Referp As DataTable = oDataBase.GetDataTable(sql)
                If dt_Referp.Rows.Count > 0 Then
                    xQtyMeter = CStr(CDbl(dt_Referp.Rows(0).Item("Data")) * 100)
                Else
                    xQtyMeter = "100"
                End If
                ' 取得完成品CHAIN(CF/VF) 或 GAP-CHAIN(MF)     (MF=金屬/CF=樹脂/VF=塑鋼)
                xPartType = "CH"                                                            ' 決定材料種類(PartTye)
                oWaves.GetItemProduct(dt_FCTItem.Rows(i).Item("Item"), xItemProduct)        ' 取得製品區分 
                '
                ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)
                If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) = 1 Then
                    If xItemProduct = "MF" Then
                        oWaves.GetChildItemStructure("01", "CH-GAP", dt_FCTItem.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                    Else
                        oWaves.GetChildItemStructure("01", "CH-DYED", dt_FCTItem.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                    End If
                Else
                    xItem(1) = dt_FCTItem.Rows(i).Item("Item")
                    xItemName(1) = ""
                    xQty(1) = CStr(1 * 10000000)
                    xCount = 1
                    For j As Integer = 2 To 5
                        xItem(j) = ""
                        xItemName(j) = ""
                        xQty(j) = "0"
                    Next
                End If
                ' 展開ITEM構成取得所指定ITEM
                If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) <> 3 Then
                    For ItemIndex As Integer = 1 To xCount
                        ' 1. 取得完成品CHAIN兼用色
                        oWaves.GetChangeColor("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCTItem.Rows(i).Item("Color"), xColor)
                        ' 2. 準備備料基準相關SearchItem
                        oWaves.GetSearchItem(xPartType, xItem(ItemIndex), xCSTPullerColor, xSearchItem)
                        ' 3. 尋找適合備料基準-從過濾箱中(M_LSOrderRule)取得適合備料基準
                        GetFilterRuleNo(pBuyer, xLTType, xPartType, xSearchItem, xRuleNo, xAction, xObjectType, xObjectProduct, xRuleCount)
                        ' 4. 依備料基準展開結構
                        For RuleIndex As Integer = 1 To xRuleCount
                            FCTWrite = False
                            ' 第一階結構資料是否OK
                            If InStr(xItemName(ItemIndex), xObjectProduct(RuleIndex)) > 0 Or xObjectProduct(RuleIndex) = "FINISH" Then
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    ' COLOMBIA-客製METER換算-廢除  2016/02/16
                                    'If pBuyer = "FALL-000003" Then
                                    '    If CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) > 0 Then
                                    '        xQtyMeter = CStr(CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) * 2.54 / 100 * 100)
                                    '    End If
                                    'End If
                                    '
                                    WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "CH", xQtyMeter, xRuleNo(RuleIndex))
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                            ' 搜尋下一階結構資料
                            If Not FCTWrite Then
                                '
                                oWaves.GetChildItemStructure("01", xObjectType(RuleIndex) + "-" + xObjectProduct(RuleIndex), xItem(ItemIndex), xItem1, xItemName1, xQty1, xCount1)
                                '
                                For ItemIndex1 As Integer = 1 To xCount1
                                    If InStr(xItemName1(ItemIndex1), xObjectProduct(RuleIndex)) > 0 Then
                                        ' 寫入 ForcastPlan
                                        sql = "SELECT * FROM ForcastPlan "
                                        sql &= "Where Buyer = '" & pBuyer & "' "
                                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                        sql &= "  And Y_Level = 0 "
                                        sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                        sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                        sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                        sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                        sql &= "Order by FCTNo, FCTSubNo "
                                        Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                        For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                            ' COLOMBIA-客製METER換算-廢除  2016/02/16
                                            'If pBuyer = "FALL-000003" Then
                                            '    If CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) > 0 Then
                                            '        xQtyMeter = CStr(CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) * 2.54 / 100 * 100)
                                            '    End If
                                            'End If
                                            '
                                            WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, xItem(ItemIndex), xItem1(ItemIndex1), xColor, "CH", xQtyMeter, xRuleNo(RuleIndex))
                                        Next
                                        xSubNo = xSubNo + 1
                                        FCTWrite = True
                                    End If
                                Next
                            End If
                            ' 搜尋不到任何結構資料(Set Error Rule [9900-990-99])
                            If Not FCTWrite Then
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    ' COLOMBIA-客製METER換算-廢除  2016/02/16
                                    'If pBuyer = "FALL-000003" Then
                                    '    If CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) > 0 Then
                                    '        xQtyMeter = CStr(CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) * 2.54 / 100 * 100)
                                    '    End If
                                    'End If
                                    '
                                    WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "CH", xQtyMeter, xRuleNo(RuleIndex) + "/" + "9900-990-00")
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                        Next
                    Next
                End If

                '##
                'MSGBOX("SLD")
                '
                ' SLIDER 處理  (需CHECK LINE-LINE ITEM)
                '--------------------------------------------------------------------------------
                xQtyMeter = "100"                       ' 取得Meter換算基準 (取得ItemClass)
                xPartType = "SLD"                       ' 決定材料種類(PartTye)
                '
                ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)
                If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) = 1 Then
                    oWaves.GetChildItemStructure("01", "SLD-FINISH", dt_FCTItem.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                Else
                    xItem(1) = dt_FCTItem.Rows(i).Item("Item")
                    xItemName(1) = ""
                    xQty(1) = CStr(1 * 10000000)
                    xCount = 1
                    For j As Integer = 2 To 5
                        xItem(j) = ""
                        xItemName(j) = ""
                        xQty(j) = "0"
                    Next
                End If
                ' 展開ITEM構成取得所指定ITEM
                If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) <> 2 Then
                    For ItemIndex As Integer = 1 To xCount
                        ' 1. 取得完成品SLD兼用色
                        oWaves.GetChangeColor("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCTItem.Rows(i).Item("Color"), xColor)
                        ' 2. 準備備料基準相關SearchItem

                        '##
                        'MsgBox(xPartType & "-" & xItem(ItemIndex))

                        oWaves.GetSearchItem(xPartType, xItem(ItemIndex), xCSTPullerColor, xSearchItem)
                        '##
                        'MsgBox(xSearchItem(1) & "-" & xSearchItem(2) & "-" & xSearchItem(3) & "-" & xSearchItem(4) & "-" & xSearchItem(5) & "-" & xSearchItem(6) & "-" & xSearchItem(7) & "-" & xSearchItem(8) & "-" & xSearchItem(9) & "-" & xSearchItem(10))

                        ' 3. 尋找適合備料基準-從過濾箱中(M_LSOrderRule)取得適合備料基準
                        GetFilterRuleNo(pBuyer, xLTType, xPartType, xSearchItem, xRuleNo, xAction, xObjectType, xObjectProduct, xRuleCount)

                        ' 4. 依備料基準展開結構
                        For RuleIndex As Integer = 1 To xRuleCount
                            '##
                            'MsgBox(xRuleNo(RuleIndex))

                            FCTWrite = False
                            ' 第一階結構資料是否OK
                            If (xObjectProduct(RuleIndex) = "FINISH" Or _
                               (xObjectProduct(RuleIndex) = "E" And InStr(xItemName(ItemIndex), xObjectProduct(RuleIndex)) > 0)) And _
                               oWaves.GetItemLineLine("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex)) = 0 Then         ' 不是LINE-LINE ITEM ((0=不是/1=是)
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "PS", xQtyMeter, xRuleNo(RuleIndex))
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                            ' 搜尋下一階結構資料
                            If Not FCTWrite Then
                                '##
                                'MSGBOX("[" & xItem(ItemIndex) & "][" & xObjectType(RuleIndex) + "-" + xObjectProduct(RuleIndex) & "]")

                                oWaves.GetChildItemStructure("01", xObjectType(RuleIndex) + "-" + xObjectProduct(RuleIndex), xItem(ItemIndex), xItem1, xItemName1, xQty1, xCount1)

                                '##
                                'MSGBOX("[" & xItem1(1) & "][" & xItemName1(1) & "]")

                                '
                                For ItemIndex1 As Integer = 1 To xCount1
                                    If oWaves.GetItemLineLine("01", dt_FCTItem.Rows(i).Item("Item"), xItem1(ItemIndex1)) = 0 Then          ' 不是LINE-LINE ITEM ((0=不是/1=是)
                                        ' 寫入 ForcastPlan
                                        sql = "SELECT * FROM ForcastPlan "
                                        sql &= "Where Buyer = '" & pBuyer & "' "
                                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                        sql &= "  And Y_Level = 0 "
                                        sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                        sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                        sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                        sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                        sql &= "Order by FCTNo, FCTSubNo "
                                        Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                        For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                            WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, xItem(ItemIndex), xItem1(ItemIndex1), xColor, "PS", xQtyMeter, xRuleNo(RuleIndex))
                                        Next
                                        xSubNo = xSubNo + 1
                                        FCTWrite = True
                                    End If
                                Next
                            End If
                            ' 搜尋不到任何結構資料(Set Error Rule [9900-990-99])
                            If Not FCTWrite Then
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "PS", xQtyMeter, xRuleNo(RuleIndex) + "/9900-990-00")
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                        Next
                    Next
                End If
                ' ZIP是否為採購品  (0=不是/1=是)
                If Not FCTWrite Then
                    If oWaves.GetPurchaseItem(dt_FCTItem.Rows(i).Item("Item")) = 1 Then
                        ' 寫入 ForcastPlan
                        sql = "SELECT * FROM ForcastPlan "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                        sql &= "  And Y_Level = 0 "
                        sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                        sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                        sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                        sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                        sql &= "Order by FCTNo, FCTSubNo "
                        Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                        For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                            WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), dt_FCTItem.Rows(i).Item("Item"), dt_FCT2.Rows(j).Item("Y_Color").ToString, "ZIP", 100, "1000-010-00")
                        Next
                        xSubNo = xSubNo + 1
                        FCTWrite = True
                    End If
                End If
            Next
            '
            '##
            'MSGBOX("OUT")
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "FCTPLAN", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "ForcastPlan", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'NewForcastPlan-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([113]-VendorForcastPlan) (Vendor專用)
    '**     Vendor Forcast Plan展開 
    '***********************************************************************************************
    'VendorForcastPlan-Start
    <WebMethod()> _
    Public Function VendorForcastPlan(ByVal pLogID As String, _
                                      ByVal pBuyer As String, _
                                      ByVal pUserID As String, _
                                      ByVal pGRBuyer As String, _
                                      ByVal pFunList As String, _
                                      ByVal pLastUniqueID As Integer) As Integer
        ' 變數定義及設定初值
        Dim RtnCode As Integer = 0
        Dim sql, str, xClass, xGRKey, xFCTNo As String
        Dim Qty, QtyTotal As Double
        '
        Try
            str = ""
            xClass = ""
            '
            ' 讀取-ForcastPlan Data
            sql = "SELECT * FROM ForcastPlan "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And Unique_ID > " & pLastUniqueID & " "
            sql &= "  And Y_Level = 0 "
            sql &= "Order by C_A1, C_B1, C_D1, C_F1, Y_G1, Y_ItemCode, Y_Color "
            Dim dt_FCTData As DataTable = oDataBase.GetDataTable(sql)
            For i As Integer = 0 To dt_FCTData.Rows.Count - 1
                '
                ' 製作 FORECAST
                ' ---------------------------------------------------------
                sql = "Insert into ForcastPlan "
                sql &= "( "
                sql &= "BuyerGroup, "
                sql &= "FCTNo, "
                sql &= "FCTSubNo, "
                sql &= "Version, "
                sql &= "C_Code, "
                sql &= "C_Color, "
                sql &= "C_SpecialRequest, "
                sql &= "C_Season, "
                sql &= "C_ShortenLT, "
                sql &= "C_Style, "
                sql &= "C_A1, "
                sql &= "C_B1, "
                sql &= "C_C1, "
                sql &= "C_D1, "
                sql &= "C_E1, "
                sql &= "C_F1, "
                sql &= "C_G1, "
                sql &= "C_H1, "
                sql &= "C_I1, "
                sql &= "C_J1, "
                sql &= "C_K1, "
                sql &= "C_L1, "
                sql &= "C_M1, "
                sql &= "C_N1, "
                sql &= "C_O1, "

                sql &= "Y_LEVEL, "
                sql &= "Y_ItemCode, "
                sql &= "Y_ItemName1, "
                sql &= "Y_ItemName2, "
                sql &= "Y_ItemName3, "
                sql &= "Y_Color, "
                sql &= "Y_A1, "
                sql &= "Y_B1, "
                sql &= "Y_C1, "
                sql &= "Y_D1, "
                sql &= "Y_E1, "
                sql &= "Y_F1, "
                sql &= "Y_G1, "
                sql &= "Y_H1, "
                sql &= "Y_I1, "
                sql &= "Y_J1, "

                sql &= "N_F, "
                sql &= "N1_F, "
                sql &= "N2_F, "
                sql &= "N3_F, "
                sql &= "N4_F, "
                sql &= "N5_F, "
                sql &= "N6_F, "
                sql &= "N7_F, "
                sql &= "N8_F, "
                sql &= "N9_F, "
                sql &= "N10_F, "
                sql &= "N11_F, "
                sql &= "N12_F, "
                sql &= "Total, "

                sql &= "BUYER, "
                sql &= "ID, "
                sql &= "CreateUser, "
                sql &= "CreateTime "
                sql &= " ) "
                ' ---------------------------------------------------------
                sql &= "VALUES( "
                sql &= " '" & dt_FCTData.Rows(i).Item("BuyerGroup").ToString & "', "                ' BuyerGroup
                sql &= " '" & dt_FCTData.Rows(i).Item("FCTNo").ToString & "', "                     ' FCTNo
                sql &= " " & CStr(dt_FCTData.Rows(i).Item("FCTSubNo") + 1) & ", "                   ' FCTSubNo
                sql &= " " & dt_FCTData.Rows(i).Item("Version").ToString & ", "                     ' Version
                sql &= " '" & dt_FCTData.Rows(i).Item("C_Code").ToString & "', "                    ' C_Code
                sql &= " '" & dt_FCTData.Rows(i).Item("C_Color").ToString & "', "                   ' C_Color
                sql &= " '" & dt_FCTData.Rows(i).Item("C_SpecialRequest").ToString & "', "          ' C_SpecialRequest
                sql &= " '" & dt_FCTData.Rows(i).Item("C_Season").ToString & "', "                  ' C_Season
                sql &= " '" & dt_FCTData.Rows(i).Item("C_ShortenLT").ToString & "', "               ' C_ShortenLT 
                sql &= " '" & dt_FCTData.Rows(i).Item("C_Style").ToString & "', "                   ' C_Style
                sql &= " '" & dt_FCTData.Rows(i).Item("C_A1").ToString & "', "                      ' C_A1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_B1").ToString & "', "                      ' C_B1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_C1").ToString & "', "                      ' C_C1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_D1").ToString & "', "                      ' C_D1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_E1").ToString & "', "                      ' C_E1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_F1").ToString & "', "                      ' C_F1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_G1").ToString & "', "                      ' C_G1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_H1").ToString & "', "                      ' C_H1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_I1").ToString & "', "                      ' C_I1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_J1").ToString & "', "                      ' C_J1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_K1").ToString & "', "                      ' C_K1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_L1").ToString & "', "                      ' C_L1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_M1").ToString & "', "                      ' C_M1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_N1").ToString & "', "                      ' C_N1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_O1").ToString & "', "                      ' C_O1
                ' ------------------------------------------------------------------
                sql &= " " & CStr(dt_FCTData.Rows(i).Item("Y_LEVEL") + 1) & ", "                    ' Y_LEVEL
                sql &= " '" & dt_FCTData.Rows(i).Item("C_I1").ToString & "', "                      ' Y_ItemCode
                oWaves.GetItemName001(dt_FCTData.Rows(i).Item("C_I1").ToString, 1, str)             ' Y_ItemName1
                sql &= " '" & str & "', "
                oWaves.GetItemName001(dt_FCTData.Rows(i).Item("C_I1").ToString, 2, str)             ' Y_ItemName2
                sql &= " '" & str & "', "
                oWaves.GetItemName001(dt_FCTData.Rows(i).Item("C_I1").ToString, 3, str)             ' Y_ItemName3
                sql &= " '" & str & "', "
                sql &= " '" & dt_FCTData.Rows(i).Item("C_K1").ToString & "', "                      ' Y_Color(不考慮兼用色)
                oWaves.GetItemClass(dt_FCTData.Rows(i).Item("C_I1").ToString, xClass)               ' Y_A1 (Item Class)


                If dt_FCTData.Rows(i).Item("Y_I1").ToString <> "" Then
                    sql &= " '" & "ZIP" & "', "
                Else
                    If xClass = "PS" Then
                        sql &= " '" & "PS" & "', "
                    Else
                        sql &= " '" & "CH" & "', "
                    End If
                End If

                sql &= " '" & dt_FCTData.Rows(i).Item("Y_B1").ToString & "', "                      ' Y_B1 (Rule Inf-原LS-RuleNo)
                sql &= " '" & dt_FCTData.Rows(i).Item("Y_C1").ToString & "', "                      ' Y_C1 (Rule Inf-採用LS-RulenNo)
                sql &= " '" & "LS" & "', "                                                          ' Y_D1 (Rule Inf-Action)
                sql &= " '" & dt_FCTData.Rows(i).Item("Y_E1").ToString & " ', "                     ' Y_E1 (SearchItem)
                sql &= " '" & dt_FCTData.Rows(i).Item("Y_F1").ToString & "', "                      ' Y_F1 (Keep)
                sql &= " '" & dt_FCTData.Rows(i).Item("Y_G1").ToString & "', "                      ' Y_G1 (N+1 ~ N+4)
                sql &= " '" & dt_FCTData.Rows(i).Item("Y_H1").ToString & "', "                      ' Y_H1
                sql &= " '" & dt_FCTData.Rows(i).Item("Y_I1").ToString & "', "                      ' Y_I1
                sql &= " '" & dt_FCTData.Rows(i).Item("Y_J1").ToString & "', "                      ' Y_J1
                ' ------------------------------------------------------------------
                Qty = CDbl(dt_FCTData.Rows(i).Item("N_F"))                                          ' N_F
                QtyTotal = Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N1_F"))                                         ' N1_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N2_F"))                                         ' N2_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N3_F"))                                         ' N3_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N4_F"))                                         ' N4_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N5_F"))                                         ' N5_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N6_F"))                                         ' N6_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N7_F"))                                         ' N7_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N8_F"))                                         ' N8_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N9_F"))                                         ' N9_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N10_F"))                                        ' N10_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N11_F"))                                        ' N11_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N12_F"))                                        ' N12_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                sql &= " " & CStr(QtyTotal) & ", "                                                  ' Total

                sql &= " '" & dt_FCTData.Rows(i).Item("BUYER").ToString & "', "                     ' BUYER
                sql &= " " & dt_FCTData.Rows(i).Item("ID").ToString & ", "                          ' ID
                sql &= " '" & "FCTPlan" & "', "
                sql &= " '" & NowDateTime & "' "
                sql &= " ) "
                oDataBase.ExecuteNonQuery(sql)
            Next
            '
            ' 同筆資料整合 17/11/23
            xGRKey = ""
            xFCTNo = ""
            sql = "SELECT * FROM ForcastPlan "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And Y_Level = 0 "
            sql &= "  And Y_G1 <> '" & "" & "' "
            sql &= "Order by C_A1, C_B1, C_D1, Y_G1, FCTNo, FCTSubNo "
            Dim dt_FCTData1 As DataTable = oDataBase.GetDataTable(sql)
            For i As Integer = 0 To dt_FCTData1.Rows.Count - 1
                str = dt_FCTData1.Rows(i).Item("C_A1").ToString + "/" + _
                      dt_FCTData1.Rows(i).Item("C_B1").ToString + "/" + _
                      dt_FCTData1.Rows(i).Item("C_D1").ToString + "/" + _
                      dt_FCTData1.Rows(i).Item("Y_G1").ToString

                If str <> xGRKey Then
                    xGRKey = dt_FCTData1.Rows(i).Item("C_A1").ToString + "/" + _
                             dt_FCTData1.Rows(i).Item("C_B1").ToString + "/" + _
                             dt_FCTData1.Rows(i).Item("C_D1").ToString + "/" + _
                             dt_FCTData1.Rows(i).Item("Y_G1").ToString
                    xFCTNo = dt_FCTData1.Rows(i).Item("FCTNo").ToString
                Else
                    ' Update ForcastPlan
                    sql = "Update ForcastPlan Set "
                    sql &= "Y_J1 = '" & "*" & "' "
                    sql &= "Where Unique_ID = " & dt_FCTData1.Rows(i).Item("Unique_ID") & " "
                    oDataBase.ExecuteNonQuery(sql)
                    ' Update ForcastPlan FCTNo & FCTSubNo
                    sql = "SELECT * FROM ForcastPlan "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "And C_A1 + '/' + C_B1 + '/' + C_D1 + '/' + Y_G1 = '" & str & "' "
                    sql &= "And Y_Level = 1 "
                    sql &= "Order by FCTNo, FCTSubNo "
                    Dim dt_FCTData2 As DataTable = oDataBase.GetDataTable(sql)
                    For j As Integer = 0 To dt_FCTData2.Rows.Count - 1
                        sql = "Update ForcastPlan Set "
                        sql &= "FCTNo = '" & xFCTNo & "', "
                        sql &= "FCTSubNo = " & j + 2 & " "
                        sql &= "Where Unique_ID = " & dt_FCTData2.Rows(j).Item("Unique_ID") & " "
                        oDataBase.ExecuteNonQuery(sql)
                    Next
                End If
            Next
            '
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "FCTPLAN", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "VendorForcastPlan", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'VendorForcastPlan-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([114]-TPForcastPlan) (YKK扣具專用)
    '**     YKK扣具 Forcast Plan展開 
    '***********************************************************************************************
    'TPForcastPlan-Start
    <WebMethod()> _
    Public Function TPForcastPlan(ByVal pLogID As String, _
                                  ByVal pBuyer As String, _
                                  ByVal pUserID As String, _
                                  ByVal pGRBuyer As String, _
                                  ByVal pFunList As String, _
                                  ByVal pLastUniqueID As Integer) As Integer
        ' 變數定義及設定初值
        Dim RtnCode As Integer = 0
        Dim sql, str, xClass As String
        Dim Qty, QtyTotal As Double
        '
        Try
            str = ""
            xClass = ""
            '
            ' 讀取-ForcastPlan Data
            sql = "SELECT * FROM ForcastPlan "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And Unique_ID > " & pLastUniqueID & " "
            sql &= "  And Y_Level = 0 "
            sql &= "Order by FCTNo, FCTSubNo "
            Dim dt_FCTData As DataTable = oDataBase.GetDataTable(sql)
            For i As Integer = 0 To dt_FCTData.Rows.Count - 1
                '
                ' 製作 FORECAST
                ' ---------------------------------------------------------
                sql = "Insert into ForcastPlan "
                sql &= "( "
                sql &= "BuyerGroup, "
                sql &= "FCTNo, "
                sql &= "FCTSubNo, "
                sql &= "Version, "
                sql &= "C_Code, "
                sql &= "C_Color, "
                sql &= "C_SpecialRequest, "
                sql &= "C_Season, "
                sql &= "C_ShortenLT, "
                sql &= "C_Style, "
                sql &= "C_A1, "
                sql &= "C_B1, "
                sql &= "C_C1, "
                sql &= "C_D1, "
                sql &= "C_E1, "
                sql &= "C_F1, "
                sql &= "C_G1, "
                sql &= "C_H1, "
                sql &= "C_I1, "
                sql &= "C_J1, "
                sql &= "C_K1, "
                sql &= "C_L1, "
                sql &= "C_M1, "
                sql &= "C_N1, "
                sql &= "C_O1, "

                sql &= "Y_LEVEL, "
                sql &= "Y_ItemCode, "
                sql &= "Y_ItemName1, "
                sql &= "Y_ItemName2, "
                sql &= "Y_ItemName3, "
                sql &= "Y_Color, "
                sql &= "Y_A1, "
                sql &= "Y_B1, "
                sql &= "Y_C1, "
                sql &= "Y_D1, "
                sql &= "Y_E1, "
                sql &= "Y_F1, "
                sql &= "Y_G1, "
                sql &= "Y_H1, "
                sql &= "Y_I1, "
                sql &= "Y_J1, "

                sql &= "N_F, "
                sql &= "N1_F, "
                sql &= "N2_F, "
                sql &= "N3_F, "
                sql &= "N4_F, "
                sql &= "N5_F, "
                sql &= "N6_F, "
                sql &= "N7_F, "
                sql &= "N8_F, "
                sql &= "N9_F, "
                sql &= "N10_F, "
                sql &= "N11_F, "
                sql &= "N12_F, "
                sql &= "Total, "

                sql &= "BUYER, "
                sql &= "ID, "
                sql &= "CreateUser, "
                sql &= "CreateTime "
                sql &= " ) "
                ' ---------------------------------------------------------
                sql &= "VALUES( "
                sql &= " '" & dt_FCTData.Rows(i).Item("BuyerGroup").ToString & "', "                ' BuyerGroup
                sql &= " '" & dt_FCTData.Rows(i).Item("FCTNo").ToString & "', "                     ' FCTNo
                sql &= " " & CStr(dt_FCTData.Rows(i).Item("FCTSubNo") + 1) & ", "                   ' FCTSubNo
                sql &= " " & dt_FCTData.Rows(i).Item("Version").ToString & ", "                     ' Version
                sql &= " '" & dt_FCTData.Rows(i).Item("C_Code").ToString & "', "                    ' C_Code
                sql &= " '" & dt_FCTData.Rows(i).Item("C_Color").ToString & "', "                   ' C_Color
                sql &= " '" & dt_FCTData.Rows(i).Item("C_SpecialRequest").ToString & "', "          ' C_SpecialRequest
                sql &= " '" & dt_FCTData.Rows(i).Item("C_Season").ToString & "', "                  ' C_Season
                sql &= " '" & dt_FCTData.Rows(i).Item("C_ShortenLT").ToString & "', "               ' C_ShortenLT 
                sql &= " '" & dt_FCTData.Rows(i).Item("C_Style").ToString & "', "                   ' C_Style
                sql &= " '" & dt_FCTData.Rows(i).Item("C_A1").ToString & "', "                      ' C_A1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_B1").ToString & "', "                      ' C_B1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_C1").ToString & "', "                      ' C_C1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_D1").ToString & "', "                      ' C_D1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_E1").ToString & "', "                      ' C_E1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_F1").ToString & "', "                      ' C_F1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_G1").ToString & "', "                      ' C_G1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_H1").ToString & "', "                      ' C_H1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_I1").ToString & "', "                      ' C_I1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_J1").ToString & "', "                      ' C_J1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_K1").ToString & "', "                      ' C_K1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_L1").ToString & "', "                      ' C_L1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_M1").ToString & "', "                      ' C_M1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_N1").ToString & "', "                      ' C_N1
                sql &= " '" & dt_FCTData.Rows(i).Item("C_O1").ToString & "', "                      ' C_O1
                ' ------------------------------------------------------------------
                sql &= " " & CStr(dt_FCTData.Rows(i).Item("Y_LEVEL") + 1) & ", "                    ' Y_LEVEL
                sql &= " '" & dt_FCTData.Rows(i).Item("Y_ItemCode").ToString & "', "                ' Y_ItemCode
                oWaves.GetItemName001(dt_FCTData.Rows(i).Item("Y_ItemCode").ToString, 1, str)       ' Y_ItemName1
                sql &= " '" & str & "', "
                oWaves.GetItemName001(dt_FCTData.Rows(i).Item("Y_ItemCode").ToString, 2, str)       ' Y_ItemName2
                sql &= " '" & str & "', "
                oWaves.GetItemName001(dt_FCTData.Rows(i).Item("Y_ItemCode").ToString, 3, str)       ' Y_ItemName3
                sql &= " '" & str & "', "
                sql &= " '" & dt_FCTData.Rows(i).Item("Y_Color").ToString & "', "                   ' Y_Color(不考慮兼用色)

                oWaves.GetItemClass(dt_FCTData.Rows(i).Item("Y_ItemCode").ToString, xClass)         ' Y_A1 (Item Class)
                If xClass = "" Then
                    sql &= " '" & "TP" & "', "
                Else
                    sql &= " '" & xClass & "', "
                End If

                sql &= " '" & dt_FCTData.Rows(i).Item("Y_B1").ToString & "', "                      ' Y_B1 (Rule Inf-原LS-RuleNo)
                sql &= " '" & dt_FCTData.Rows(i).Item("Y_C1").ToString & "', "                      ' Y_C1 (Rule Inf-採用LS-RulenNo)
                sql &= " '" & "LS" & "', "                                                          ' Y_D1 (Rule Inf-Action)
                sql &= " '" & dt_FCTData.Rows(i).Item("Y_E1").ToString & " ', "                     ' Y_E1 (SearchItem)
                sql &= " '" & dt_FCTData.Rows(i).Item("Y_F1").ToString & "', "                      ' Y_F1 (Keep)
                sql &= " '" & dt_FCTData.Rows(i).Item("Y_G1").ToString & "', "                      ' Y_G1 (N+1 ~ N+4)
                sql &= " '" & dt_FCTData.Rows(i).Item("Y_H1").ToString & "', "                      ' Y_H1
                sql &= " '" & dt_FCTData.Rows(i).Item("Y_I1").ToString & "', "                      ' Y_I1
                sql &= " '" & dt_FCTData.Rows(i).Item("Y_J1").ToString & "', "                      ' Y_J1
                ' ------------------------------------------------------------------
                Qty = CDbl(dt_FCTData.Rows(i).Item("N_F"))                                          ' N_F
                QtyTotal = Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N1_F"))                                         ' N1_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N2_F"))                                         ' N2_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N3_F"))                                         ' N3_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N4_F"))                                         ' N4_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N5_F"))                                         ' N5_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N6_F"))                                         ' N6_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N7_F"))                                         ' N7_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N8_F"))                                         ' N8_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N9_F"))                                         ' N9_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N10_F"))                                        ' N10_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N11_F"))                                        ' N11_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = CDbl(dt_FCTData.Rows(i).Item("N12_F"))                                        ' N12_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                sql &= " " & CStr(QtyTotal) & ", "                                                  ' Total

                sql &= " '" & dt_FCTData.Rows(i).Item("BUYER").ToString & "', "                     ' BUYER
                sql &= " " & dt_FCTData.Rows(i).Item("ID").ToString & ", "                          ' ID
                sql &= " '" & "FCTPlan" & "', "
                sql &= " '" & NowDateTime & "' "
                sql &= " ) "
                oDataBase.ExecuteNonQuery(sql)
            Next
            '
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "FCTPLAN", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "TPForcastPlan", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'TPForcastPlan-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([115]-MatForcastPlan) (材料專用)
    '**     材料 Forcast Plan展開 
    '***********************************************************************************************
    'MatForcastPlan-Start
    <WebMethod()> _
    Public Function MatForcastPlan(ByVal pLogID As String, _
                                   ByVal pBuyer As String, _
                                   ByVal pUserID As String, _
                                   ByVal pGRBuyer As String, _
                                   ByVal pFunList As String, _
                                   ByVal pLastUniqueID As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        '
        Try
            '
            ' 變數定義及設定初值
            oWaves.Timeout = Timeout.Infinite
            Dim xItemProduct, xLTType, xPartType, xClass, xQtyMeter, xColor As String
            Dim xSubNo As Integer
            Dim FCTWrite As Boolean = False
            ' 構成展開專用變數
            Dim xItem(5), xItemName(5), xQty(5) As String
            Dim xItem1(5), xItemName1(5), xQty1(5) As String
            Dim xCount, xCount1 As Integer
            ' SearchItem專用變數
            Dim xSearchItem(10) As String
            ' 過濾箱中取得的資料變數
            Dim xRuleNo(5), xAction(5), xObjectType(5), xObjectProduct(5) As String
            Dim xRuleCount As Integer
            '
            ' 讀取-ForcastPlan Data
            sql = "SELECT C_Season As Season, Y_ItemCode AS Item, Y_Color AS Color, C_Color AS CSTColor, C_ShortenLT AS KeepCode, "
            sql &= "Y_ItemName1 + ' ' + Y_ItemName2 + ' ' + Y_ItemName3 AS ItemName, C_E1 AS Class "
            sql &= "FROM ForcastPlan "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And Unique_ID > " & pLastUniqueID & " "
            'sql &= "  And Y_ItemCode = '" & "3278472" & "' "      'JOY-ADD 180725
            sql &= "  And Y_Level = 0 "
            sql &= "Group by C_Season, Y_ItemCode, Y_Color, C_Color, C_ShortenLT, Y_ItemName1, Y_ItemName2, Y_ItemName3, C_E1 "
            sql &= "Order by C_Season, Y_ItemCode, Y_Color, C_Color, C_ShortenLT, Y_ItemName1, Y_ItemName2, Y_ItemName3, C_E1 "
            Dim dt_FCTItem As DataTable = oDataBase.GetDataTable(sql)
            For i As Integer = 0 To dt_FCTItem.Rows.Count - 1

                FCTWrite = False
                xSubNo = 2
                xLTType = "LLT"                                                         ' 決定LT種類(LTType)
                Dim xCSTPullerColor As String = dt_FCTItem.Rows(i).Item("CSTColor")    ' 決定CST PULLER COLOR

                If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) = 1 Then          ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)
                    ' ZIPPER 處理 
                    '--------------------------------------------------------------------------------
                    '
                    ' CHAIN 處理 (不需CHECK LINE-LINE ITEM)
                    '--------------------------------------------------------------------------------
                    ' 取得Meter換算基準 (取得ItemClass)
                    oWaves.GetItemClass(dt_FCTItem.Rows(i).Item("Item"), xClass)
                    sql = "Select * From M_Referp "
                    sql = sql & " Where Cat  = '" & "110" & "' "
                    sql = sql & "   And DKey = '" & pBuyer + "-" + xClass & "' "
                    Dim dt_Referp As DataTable = oDataBase.GetDataTable(sql)
                    If dt_Referp.Rows.Count > 0 Then
                        xQtyMeter = CStr(CDbl(dt_Referp.Rows(0).Item("Data")) * 100)
                    Else
                        xQtyMeter = "100"
                    End If
                    ' 取得完成品CHAIN(CF/VF) 或 GAP-CHAIN(MF)     (MF=金屬/CF=樹脂/VF=塑鋼)
                    xPartType = "CH"                                                            ' 決定材料種類(PartTye)
                    oWaves.GetItemProduct(dt_FCTItem.Rows(i).Item("Item"), xItemProduct)        ' 取得製品區分 
                    '
                    ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)
                    If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) = 1 Then
                        If xItemProduct = "MF" Then
                            oWaves.GetChildItemStructure("01", "CH-GAP", dt_FCTItem.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                        Else
                            oWaves.GetChildItemStructure("01", "CH-DYED", dt_FCTItem.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                        End If
                    Else
                        xItem(1) = dt_FCTItem.Rows(i).Item("Item")
                        xItemName(1) = ""
                        xQty(1) = CStr(1 * 10000000)
                        xCount = 1
                        For j As Integer = 2 To 5
                            xItem(j) = ""
                            xItemName(j) = ""
                            xQty(j) = "0"
                        Next
                    End If
                    ' 展開ITEM構成取得所指定ITEM
                    If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) <> 3 Then
                        For ItemIndex As Integer = 1 To xCount
                            ' 1. 取得完成品CHAIN兼用色
                            oWaves.GetChangeColor("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCTItem.Rows(i).Item("Color"), xColor)
                            ' 2. 準備備料基準相關SearchItem
                            oWaves.GetSearchItem(xPartType, xItem(ItemIndex), xCSTPullerColor, xSearchItem)
                            ' 3. 尋找適合備料基準-從過濾箱中(M_LSOrderRule)取得適合備料基準
                            GetFilterRuleNo(pBuyer, xLTType, xPartType, xSearchItem, xRuleNo, xAction, xObjectType, xObjectProduct, xRuleCount)
                            ' 4. 依備料基準展開結構
                            For RuleIndex As Integer = 1 To xRuleCount
                                FCTWrite = False
                                ' 第一階結構資料是否OK
                                If InStr(xItemName(ItemIndex), xObjectProduct(RuleIndex)) > 0 Then
                                    ' 寫入 ForcastPlan
                                    sql = "SELECT * FROM ForcastPlan "
                                    sql &= "Where Buyer = '" & pBuyer & "' "
                                    sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                    sql &= "  And Y_Level = 0 "
                                    sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                    sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                    sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                    sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                    sql &= "  And C_E1 = '" & dt_FCTItem.Rows(i).Item("Class") & "' "       'JOY-ADD 180724
                                    sql &= "Order by FCTNo, FCTSubNo "
                                    Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                    For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                        ' COLOMBIA-客製METER換算-廢除  2016/02/16
                                        'If pBuyer = "FALL-000003" Then
                                        '    If CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) > 0 Then
                                        '        xQtyMeter = CStr(CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) * 2.54 / 100 * 100)
                                        '    End If
                                        'End If
                                        '
                                        WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "CH", xQtyMeter, xRuleNo(RuleIndex))
                                    Next
                                    xSubNo = xSubNo + 1
                                    FCTWrite = True
                                End If
                                ' 搜尋下一階結構資料
                                If Not FCTWrite Then
                                    '
                                    oWaves.GetChildItemStructure("01", xObjectType(RuleIndex) + "-" + xObjectProduct(RuleIndex), xItem(ItemIndex), xItem1, xItemName1, xQty1, xCount1)
                                    '
                                    For ItemIndex1 As Integer = 1 To xCount1
                                        If InStr(xItemName1(ItemIndex1), xObjectProduct(RuleIndex)) > 0 Then
                                            ' 寫入 ForcastPlan
                                            sql = "SELECT * FROM ForcastPlan "
                                            sql &= "Where Buyer = '" & pBuyer & "' "
                                            sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                            sql &= "  And Y_Level = 0 "
                                            sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                            sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                            sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                            sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                            sql &= "  And C_E1 = '" & dt_FCTItem.Rows(i).Item("Class") & "' "       'JOY-ADD 180724
                                            sql &= "Order by FCTNo, FCTSubNo "
                                            Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                            For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                                ' COLOMBIA-客製METER換算-廢除  2016/02/16
                                                'If pBuyer = "FALL-000003" Then
                                                '    If CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) > 0 Then
                                                '        xQtyMeter = CStr(CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) * 2.54 / 100 * 100)
                                                '    End If
                                                'End If
                                                '
                                                WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, xItem(ItemIndex), xItem1(ItemIndex1), xColor, "CH", xQtyMeter, xRuleNo(RuleIndex))
                                            Next
                                            xSubNo = xSubNo + 1
                                            FCTWrite = True
                                        End If
                                    Next
                                End If
                                ' 搜尋不到任何結構資料(Set Error Rule [9900-990-99])
                                If Not FCTWrite Then
                                    ' 寫入 ForcastPlan
                                    sql = "SELECT * FROM ForcastPlan "
                                    sql &= "Where Buyer = '" & pBuyer & "' "
                                    sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                    sql &= "  And Y_Level = 0 "
                                    sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                    sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                    sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                    sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                    sql &= "  And C_E1 = '" & dt_FCTItem.Rows(i).Item("Class") & "' "       'JOY-ADD 180724
                                    sql &= "Order by FCTNo, FCTSubNo "
                                    Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                    For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                        ' COLOMBIA-客製METER換算-廢除  2016/02/16
                                        'If pBuyer = "FALL-000003" Then
                                        '    If CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) > 0 Then
                                        '        xQtyMeter = CStr(CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) * 2.54 / 100 * 100)
                                        '    End If
                                        'End If
                                        '
                                        WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "CH", xQtyMeter, xRuleNo(RuleIndex) + "/" + "9900-990-00")
                                    Next
                                    xSubNo = xSubNo + 1
                                    FCTWrite = True
                                End If
                            Next
                        Next
                    End If
                    '
                    ' SLIDER 處理  (需CHECK LINE-LINE ITEM)
                    '--------------------------------------------------------------------------------
                    xQtyMeter = "100"                       ' 取得Meter換算基準 (取得ItemClass)
                    xPartType = "SLD"                       ' 決定材料種類(PartTye)
                    '
                    ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)
                    If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) = 1 Then
                        oWaves.GetChildItemStructure("01", "SLD-FINISH", dt_FCTItem.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                    Else
                        xItem(1) = dt_FCTItem.Rows(i).Item("Item")
                        xItemName(1) = ""
                        xQty(1) = CStr(1 * 10000000)
                        xCount = 1
                        For j As Integer = 2 To 5
                            xItem(j) = ""
                            xItemName(j) = ""
                            xQty(j) = "0"
                        Next
                    End If
                    ' 展開ITEM構成取得所指定ITEM
                    If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) <> 2 Then
                        For ItemIndex As Integer = 1 To xCount
                            ' 1. 取得完成品SLD兼用色
                            oWaves.GetChangeColor("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCTItem.Rows(i).Item("Color"), xColor)
                            ' 2. 準備備料基準相關SearchItem
                            oWaves.GetSearchItem(xPartType, xItem(ItemIndex), xCSTPullerColor, xSearchItem)
                            ' 3. 尋找適合備料基準-從過濾箱中(M_LSOrderRule)取得適合備料基準
                            GetFilterRuleNo(pBuyer, xLTType, xPartType, xSearchItem, xRuleNo, xAction, xObjectType, xObjectProduct, xRuleCount)
                            ' 4. 依備料基準展開結構
                            For RuleIndex As Integer = 1 To xRuleCount
                                FCTWrite = False
                                ' 第一階結構資料是否OK
                                If (xObjectProduct(RuleIndex) = "FINISH" Or _
                                   (xObjectProduct(RuleIndex) = "E" And InStr(xItemName(ItemIndex), xObjectProduct(RuleIndex)) > 0)) And _
                                   oWaves.GetItemLineLine("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex)) = 0 Then         ' 不是LINE-LINE ITEM ((0=不是/1=是)
                                    ' 寫入 ForcastPlan
                                    sql = "SELECT * FROM ForcastPlan "
                                    sql &= "Where Buyer = '" & pBuyer & "' "
                                    sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                    sql &= "  And Y_Level = 0 "
                                    sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                    sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                    sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                    sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                    sql &= "  And C_E1 = '" & dt_FCTItem.Rows(i).Item("Class") & "' "       'JOY-ADD 180724
                                    sql &= "Order by FCTNo, FCTSubNo "
                                    Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                    For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                        WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "PS", xQtyMeter, xRuleNo(RuleIndex))
                                    Next
                                    xSubNo = xSubNo + 1
                                    FCTWrite = True
                                End If
                                ' 搜尋下一階結構資料
                                If Not FCTWrite Then
                                    oWaves.GetChildItemStructure("01", xObjectType(RuleIndex) + "-" + xObjectProduct(RuleIndex), xItem(ItemIndex), xItem1, xItemName1, xQty1, xCount1)
                                    '
                                    For ItemIndex1 As Integer = 1 To xCount1
                                        If oWaves.GetItemLineLine("01", dt_FCTItem.Rows(i).Item("Item"), xItem1(ItemIndex1)) = 0 Then          ' 不是LINE-LINE ITEM ((0=不是/1=是)
                                            ' 寫入 ForcastPlan
                                            sql = "SELECT * FROM ForcastPlan "
                                            sql &= "Where Buyer = '" & pBuyer & "' "
                                            sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                            sql &= "  And Y_Level = 0 "
                                            sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                            sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                            sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                            sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                            sql &= "  And C_E1 = '" & dt_FCTItem.Rows(i).Item("Class") & "' "       'JOY-ADD 180724
                                            sql &= "Order by FCTNo, FCTSubNo "
                                            Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                            For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                                WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, xItem(ItemIndex), xItem1(ItemIndex1), xColor, "PS", xQtyMeter, xRuleNo(RuleIndex))
                                            Next
                                            xSubNo = xSubNo + 1
                                            FCTWrite = True
                                        End If
                                    Next
                                End If
                                ' 搜尋不到任何結構資料(Set Error Rule [9900-990-99])
                                If Not FCTWrite Then
                                    ' 寫入 ForcastPlan
                                    sql = "SELECT * FROM ForcastPlan "
                                    sql &= "Where Buyer = '" & pBuyer & "' "
                                    sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                    sql &= "  And Y_Level = 0 "
                                    sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                    sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                    sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                    sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                    sql &= "  And C_E1 = '" & dt_FCTItem.Rows(i).Item("Class") & "' "       'JOY-ADD 180724
                                    sql &= "Order by FCTNo, FCTSubNo "
                                    Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                    For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                        WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "PS", xQtyMeter, xRuleNo(RuleIndex) + "/9900-990-00")
                                    Next
                                    xSubNo = xSubNo + 1
                                    FCTWrite = True
                                End If
                            Next
                        Next
                    End If
                End If
                '
                If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) = 2 Then          ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)
                    ' CHAIN 處理
                    '--------------------------------------------------------------------------------
                    ' 取得Meter換算基準 (取得ItemClass)
                    sql = "Select * From M_Referp "
                    sql = sql & " Where Cat  = '" & "110" & "' "
                    sql = sql & "   And DKey = '" & pBuyer + "-" + dt_FCTItem.Rows(i).Item("Class") & "' "
                    Dim dt_Referp As DataTable = oDataBase.GetDataTable(sql)
                    If dt_Referp.Rows.Count > 0 Then
                        xQtyMeter = CStr(CDbl(dt_Referp.Rows(0).Item("Data")) * 100)
                    Else
                        xQtyMeter = "100"
                    End If
                    ' 取得完成品CHAIN(CF/VF) 或 GAP-CHAIN(MF)     (MF=金屬/CF=樹脂/VF=塑鋼)
                    xPartType = "CH"                                                            ' 決定材料種類(PartTye)
                    oWaves.GetItemProduct(dt_FCTItem.Rows(i).Item("Item"), xItemProduct)        ' 取得製品區分 
                    '
                    xItem(1) = dt_FCTItem.Rows(i).Item("Item")
                    xItemName(1) = dt_FCTItem.Rows(i).Item("ItemName")
                    xQty(1) = CStr(1 * 10000000)
                    xCount = 1
                    For j As Integer = 2 To 5
                        xItem(j) = ""
                        xItemName(j) = ""
                        xQty(j) = "0"
                    Next
                    '
                    ' 展開ITEM構成取得所指定ITEM
                    For ItemIndex As Integer = 1 To xCount
                        ' 1. 取得完成品CHAIN兼用色
                        oWaves.GetChangeColor("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCTItem.Rows(i).Item("Color"), xColor)
                        ' 2. 準備備料基準相關SearchItem
                        oWaves.GetSearchItem(xPartType, xItem(ItemIndex), xCSTPullerColor, xSearchItem)
                        ' 3. 尋找適合備料基準-從過濾箱中(M_LSOrderRule)取得適合備料基準
                        GetFilterRuleNo(pBuyer, xLTType, xPartType, xSearchItem, xRuleNo, xAction, xObjectType, xObjectProduct, xRuleCount)
                        ' 4. 依備料基準展開結構
                        For RuleIndex As Integer = 1 To xRuleCount
                            FCTWrite = False
                            ' 第一階結構資料是否OK
                            If InStr(xItemName(ItemIndex), xObjectProduct(RuleIndex)) > 0 Then
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "  And C_E1 = '" & dt_FCTItem.Rows(i).Item("Class") & "' "       'JOY-ADD 180724
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "CH", xQtyMeter, xRuleNo(RuleIndex))
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                            ' 搜尋下一階結構資料
                            If Not FCTWrite Then
                                '
                                oWaves.GetChildItemStructure("01", xObjectType(RuleIndex) + "-" + xObjectProduct(RuleIndex), xItem(ItemIndex), xItem1, xItemName1, xQty1, xCount1)
                                '
                                For ItemIndex1 As Integer = 1 To xCount1
                                    If InStr(xItemName1(ItemIndex1), xObjectProduct(RuleIndex)) > 0 Then
                                        ' 寫入 ForcastPlan
                                        sql = "SELECT * FROM ForcastPlan "
                                        sql &= "Where Buyer = '" & pBuyer & "' "
                                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                        sql &= "  And Y_Level = 0 "
                                        sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                        sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                        sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                        sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                        sql &= "  And C_E1 = '" & dt_FCTItem.Rows(i).Item("Class") & "' "       'JOY-ADD 180724
                                        sql &= "Order by FCTNo, FCTSubNo "
                                        Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                        For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                            WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, xItem(ItemIndex), xItem1(ItemIndex1), xColor, "CH", xQtyMeter, xRuleNo(RuleIndex))
                                        Next
                                        xSubNo = xSubNo + 1
                                        FCTWrite = True
                                    End If
                                Next
                            End If
                            ' 搜尋不到任何結構資料(Set Error Rule [9900-990-99])
                            If Not FCTWrite Then
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "  And C_E1 = '" & dt_FCTItem.Rows(i).Item("Class") & "' "       'JOY-ADD 180724
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "CH", xQtyMeter, xRuleNo(RuleIndex) + "/" + "9900-990-00")
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                        Next
                    Next
                End If
                '
                If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) = 3 Then          ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)
                    '
                    ' SLIDER 處理  
                    '--------------------------------------------------------------------------------
                    xQtyMeter = "100"                       ' 取得Meter換算基準 (取得ItemClass)
                    xPartType = "SLD"                       ' 決定材料種類(PartTye)
                    '
                    ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)
                    xItem(1) = dt_FCTItem.Rows(i).Item("Item")
                    xItemName(1) = dt_FCTItem.Rows(i).Item("ItemName")
                    xQty(1) = CStr(1 * 10000000)
                    xCount = 1
                    For j As Integer = 2 To 5
                        xItem(j) = ""
                        xItemName(j) = ""
                        xQty(j) = "0"
                    Next
                    ' 展開ITEM構成取得所指定ITEM
                    For ItemIndex As Integer = 1 To xCount
                        ' 1. 取得完成品SLD兼用色
                        oWaves.GetChangeColor("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCTItem.Rows(i).Item("Color"), xColor)
                        ' 2. 準備備料基準相關SearchItem
                        oWaves.GetSearchItem(xPartType, xItem(ItemIndex), xCSTPullerColor, xSearchItem)
                        ' 3. 尋找適合備料基準-從過濾箱中(M_LSOrderRule)取得適合備料基準
                        GetFilterRuleNo(pBuyer, xLTType, xPartType, xSearchItem, xRuleNo, xAction, xObjectType, xObjectProduct, xRuleCount)
                        ' 4. 依備料基準展開結構
                        For RuleIndex As Integer = 1 To xRuleCount
                            FCTWrite = False
                            ' 第一階結構資料是否OK
                            If (xObjectProduct(RuleIndex) = "FINISH" Or _
                               (xObjectProduct(RuleIndex) = "E" And InStr(xItemName(ItemIndex), xObjectProduct(RuleIndex)) > 0)) And _
                               oWaves.GetItemLineLine("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex)) = 0 Then         ' 不是LINE-LINE ITEM ((0=不是/1=是)
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "  And C_E1 = '" & dt_FCTItem.Rows(i).Item("Class") & "' "       'JOY-ADD 180724
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "PS", xQtyMeter, xRuleNo(RuleIndex))
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                            ' 搜尋下一階結構資料
                            If Not FCTWrite Then
                                oWaves.GetChildItemStructure("01", xObjectType(RuleIndex) + "-" + xObjectProduct(RuleIndex), xItem(ItemIndex), xItem1, xItemName1, xQty1, xCount1)
                                '
                                For ItemIndex1 As Integer = 1 To xCount1
                                    If oWaves.GetItemLineLine("01", dt_FCTItem.Rows(i).Item("Item"), xItem1(ItemIndex1)) = 0 Then          ' 不是LINE-LINE ITEM ((0=不是/1=是)
                                        ' 寫入 ForcastPlan
                                        sql = "SELECT * FROM ForcastPlan "
                                        sql &= "Where Buyer = '" & pBuyer & "' "
                                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                        sql &= "  And Y_Level = 0 "
                                        sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                        sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                        sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                        sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                        sql &= "  And C_E1 = '" & dt_FCTItem.Rows(i).Item("Class") & "' "       'JOY-ADD 180724
                                        sql &= "Order by FCTNo, FCTSubNo "
                                        Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                        For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                            WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, xItem(ItemIndex), xItem1(ItemIndex1), xColor, "PS", xQtyMeter, xRuleNo(RuleIndex))
                                        Next
                                        xSubNo = xSubNo + 1
                                        FCTWrite = True
                                    End If
                                Next
                            End If
                            ' 搜尋不到任何結構資料(Set Error Rule [9900-990-99])
                            If Not FCTWrite Then
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "  And C_E1 = '" & dt_FCTItem.Rows(i).Item("Class") & "' "       'JOY-ADD 180724
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "PS", xQtyMeter, xRuleNo(RuleIndex) + "/9900-990-00")
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                        Next
                    Next
                End If
                ' ZIP是否為採購品  (0=不是/1=是)
                If Not FCTWrite Then
                    If oWaves.GetPurchaseItem(dt_FCTItem.Rows(i).Item("Item")) = 1 Then
                        ' 寫入 ForcastPlan
                        sql = "SELECT * FROM ForcastPlan "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                        sql &= "  And Y_Level = 0 "
                        sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                        sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                        sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                        sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                        sql &= "  And C_E1 = '" & dt_FCTItem.Rows(i).Item("Class") & "' "       'JOY-ADD 180724
                        sql &= "Order by FCTNo, FCTSubNo "
                        Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                        For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                            WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), dt_FCTItem.Rows(i).Item("Item"), dt_FCT2.Rows(j).Item("Y_Color").ToString, "ZIP", 100, "1000-010-00")
                        Next
                        xSubNo = xSubNo + 1
                        FCTWrite = True
                    End If
                End If
            Next
            '
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "FCTPLAN", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "ForcastPlan", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'MatForcastPlan-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([115]-MatForcastPlanUABAG) (材料專用)
    '**     UABAG材料 Forcast Plan展開 
    '***********************************************************************************************
    'MatForcastPlanUABAG-Start
    <WebMethod()> _
    Public Function MatForcastPlanUABAG(ByVal pLogID As String, _
                                   ByVal pBuyer As String, _
                                   ByVal pUserID As String, _
                                   ByVal pGRBuyer As String, _
                                   ByVal pFunList As String, _
                                   ByVal pLastUniqueID As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        '
        Try
            'MsgBox("UABAG-IN")
            '
            ' 變數定義及設定初值
            oWaves.Timeout = Timeout.Infinite
            Dim xItemProduct, xLTType, xPartType, xClass, xQtyMeter, xColor As String
            Dim xSubNo As Integer
            Dim FCTWrite As Boolean = False
            ' 構成展開專用變數
            Dim xItem(5), xItemName(5), xQty(5) As String
            Dim xItem1(5), xItemName1(5), xQty1(5) As String
            Dim xCount, xCount1 As Integer
            ' SearchItem專用變數
            Dim xSearchItem(10) As String
            ' 過濾箱中取得的資料變數
            Dim xRuleNo(5), xAction(5), xObjectType(5), xObjectProduct(5) As String
            Dim xRuleCount As Integer
            '
            ' 讀取-ForcastPlan Data
            sql = "SELECT C_Season As Season, Y_ItemCode AS Item, Y_Color AS Color, C_Color AS CSTColor, C_ShortenLT AS KeepCode, "
            sql &= "Y_ItemName1 + ' ' + Y_ItemName2 + ' ' + Y_ItemName3 AS ItemName, C_B1 AS CUST, C_CODE AS CITEM "
            sql &= "FROM ForcastPlan "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And Unique_ID > " & pLastUniqueID & " "
            sql &= "  And Y_Level = 0 "
            sql &= "Group by C_Season, Y_ItemCode, Y_Color, C_Color, C_ShortenLT, Y_ItemName1, Y_ItemName2, Y_ItemName3, C_B1, C_CODE "
            sql &= "Order by C_Season, Y_ItemCode, Y_Color, C_Color, C_ShortenLT, Y_ItemName1, Y_ItemName2, Y_ItemName3, C_B1, C_CODE "
            Dim dt_FCTItem As DataTable = oDataBase.GetDataTable(sql)
            For i As Integer = 0 To dt_FCTItem.Rows.Count - 1

                FCTWrite = False
                xSubNo = 2
                xLTType = "LLT"                                                         ' 決定LT種類(LTType)
                Dim xCSTPullerColor As String = dt_FCTItem.Rows(i).Item("CSTColor")    ' 決定CST PULLER COLOR

                If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) = 1 Then          ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)

                    'MsgBox("ZIPPER")
                    ' ZIPPER 處理 
                    '--------------------------------------------------------------------------------
                    '
                    ' CHAIN 處理 (不需CHECK LINE-LINE ITEM)
                    '--------------------------------------------------------------------------------
                    ' 取得Meter換算基準 (取得ItemClass)
                    oWaves.GetItemClass(dt_FCTItem.Rows(i).Item("Item"), xClass)
                    sql = "Select * From M_Referp "
                    sql = sql & " Where Cat  = '" & "110" & "' "
                    sql = sql & "   And DKey = '" & pBuyer + "-" + xClass & "' "
                    Dim dt_Referp As DataTable = oDataBase.GetDataTable(sql)
                    If dt_Referp.Rows.Count > 0 Then
                        xQtyMeter = CStr(CDbl(dt_Referp.Rows(0).Item("Data")) * 100)
                    Else
                        xQtyMeter = "100"
                    End If
                    ' 取得完成品CHAIN(CF/VF) 或 GAP-CHAIN(MF)     (MF=金屬/CF=樹脂/VF=塑鋼)
                    xPartType = "CH"                                                            ' 決定材料種類(PartTye)
                    oWaves.GetItemProduct(dt_FCTItem.Rows(i).Item("Item"), xItemProduct)        ' 取得製品區分 
                    '
                    ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)
                    If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) = 1 Then
                        If xItemProduct = "MF" Then
                            oWaves.GetChildItemStructure("01", "CH-GAP", dt_FCTItem.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                        Else
                            oWaves.GetChildItemStructure("01", "CH-DYED", dt_FCTItem.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                        End If
                    Else
                        xItem(1) = dt_FCTItem.Rows(i).Item("Item")
                        xItemName(1) = ""
                        xQty(1) = CStr(1 * 10000000)
                        xCount = 1
                        For j As Integer = 2 To 5
                            xItem(j) = ""
                            xItemName(j) = ""
                            xQty(j) = "0"
                        Next
                    End If
                    ' 展開ITEM構成取得所指定ITEM
                    If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) <> 3 Then
                        For ItemIndex As Integer = 1 To xCount
                            ' 1. 取得完成品CHAIN兼用色
                            oWaves.GetChangeColor("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCTItem.Rows(i).Item("Color"), xColor)
                            ' 2. 準備備料基準相關SearchItem
                            oWaves.GetSearchItem(xPartType, xItem(ItemIndex), xCSTPullerColor, xSearchItem)
                            ' 3. 尋找適合備料基準-從過濾箱中(M_LSOrderRule)取得適合備料基準
                            GetFilterRuleNo(pBuyer, xLTType, xPartType, xSearchItem, xRuleNo, xAction, xObjectType, xObjectProduct, xRuleCount)
                            ' 4. 依備料基準展開結構
                            For RuleIndex As Integer = 1 To xRuleCount
                                FCTWrite = False
                                ' 第一階結構資料是否OK
                                If InStr(xItemName(ItemIndex), xObjectProduct(RuleIndex)) > 0 Then
                                    ' 寫入 ForcastPlan
                                    sql = "SELECT * FROM ForcastPlan "
                                    sql &= "Where Buyer = '" & pBuyer & "' "
                                    sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                    sql &= "  And Y_Level = 0 "
                                    sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                    sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                    sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                    sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                    sql &= "  And C_B1 = '" & dt_FCTItem.Rows(i).Item("CUST") & "' "
                                    sql &= "  And C_CODE = '" & dt_FCTItem.Rows(i).Item("CITEM") & "' "
                                    sql &= "Order by FCTNo, FCTSubNo "
                                    Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                    For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                        ' COLOMBIA-客製METER換算-廢除  2016/02/16
                                        'If pBuyer = "FALL-000003" Then
                                        '    If CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) > 0 Then
                                        '        xQtyMeter = CStr(CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) * 2.54 / 100 * 100)
                                        '    End If
                                        'End If
                                        '
                                        WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "CH", xQtyMeter, xRuleNo(RuleIndex))
                                    Next
                                    xSubNo = xSubNo + 1
                                    FCTWrite = True
                                End If
                                ' 搜尋下一階結構資料
                                If Not FCTWrite Then
                                    '
                                    oWaves.GetChildItemStructure("01", xObjectType(RuleIndex) + "-" + xObjectProduct(RuleIndex), xItem(ItemIndex), xItem1, xItemName1, xQty1, xCount1)
                                    '
                                    For ItemIndex1 As Integer = 1 To xCount1
                                        If InStr(xItemName1(ItemIndex1), xObjectProduct(RuleIndex)) > 0 Then
                                            ' 寫入 ForcastPlan
                                            sql = "SELECT * FROM ForcastPlan "
                                            sql &= "Where Buyer = '" & pBuyer & "' "
                                            sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                            sql &= "  And Y_Level = 0 "
                                            sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                            sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                            sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                            sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                            sql &= "  And C_B1 = '" & dt_FCTItem.Rows(i).Item("CUST") & "' "
                                            sql &= "  And C_CODE = '" & dt_FCTItem.Rows(i).Item("CITEM") & "' "
                                            sql &= "Order by FCTNo, FCTSubNo "
                                            Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                            For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                                ' COLOMBIA-客製METER換算-廢除  2016/02/16
                                                'If pBuyer = "FALL-000003" Then
                                                '    If CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) > 0 Then
                                                '        xQtyMeter = CStr(CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) * 2.54 / 100 * 100)
                                                '    End If
                                                'End If
                                                '
                                                WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, xItem(ItemIndex), xItem1(ItemIndex1), xColor, "CH", xQtyMeter, xRuleNo(RuleIndex))
                                            Next
                                            xSubNo = xSubNo + 1
                                            FCTWrite = True
                                        End If
                                    Next
                                End If
                                ' 搜尋不到任何結構資料(Set Error Rule [9900-990-99])
                                If Not FCTWrite Then
                                    ' 寫入 ForcastPlan
                                    sql = "SELECT * FROM ForcastPlan "
                                    sql &= "Where Buyer = '" & pBuyer & "' "
                                    sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                    sql &= "  And Y_Level = 0 "
                                    sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                    sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                    sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                    sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                    sql &= "  And C_B1 = '" & dt_FCTItem.Rows(i).Item("CUST") & "' "
                                    sql &= "  And C_CODE = '" & dt_FCTItem.Rows(i).Item("CITEM") & "' "
                                    sql &= "Order by FCTNo, FCTSubNo "
                                    Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                    For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                        ' COLOMBIA-客製METER換算-廢除  2016/02/16
                                        'If pBuyer = "FALL-000003" Then
                                        '    If CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) > 0 Then
                                        '        xQtyMeter = CStr(CDbl(dt_FCT2.Rows(j).Item("C_G1").ToString) * 2.54 / 100 * 100)
                                        '    End If
                                        'End If
                                        '
                                        WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "CH", xQtyMeter, xRuleNo(RuleIndex) + "/" + "9900-990-00")
                                    Next
                                    xSubNo = xSubNo + 1
                                    FCTWrite = True
                                End If
                            Next
                        Next
                    End If
                    '
                    ' SLIDER 處理  (需CHECK LINE-LINE ITEM)
                    '--------------------------------------------------------------------------------
                    xQtyMeter = "100"                       ' 取得Meter換算基準 (取得ItemClass)
                    xPartType = "SLD"                       ' 決定材料種類(PartTye)
                    '
                    ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)
                    If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) = 1 Then
                        oWaves.GetChildItemStructure("01", "SLD-FINISH", dt_FCTItem.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                    Else
                        xItem(1) = dt_FCTItem.Rows(i).Item("Item")
                        xItemName(1) = ""
                        xQty(1) = CStr(1 * 10000000)
                        xCount = 1
                        For j As Integer = 2 To 5
                            xItem(j) = ""
                            xItemName(j) = ""
                            xQty(j) = "0"
                        Next
                    End If
                    ' 展開ITEM構成取得所指定ITEM
                    If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) <> 2 Then
                        For ItemIndex As Integer = 1 To xCount
                            ' 1. 取得完成品SLD兼用色
                            oWaves.GetChangeColor("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCTItem.Rows(i).Item("Color"), xColor)
                            ' 2. 準備備料基準相關SearchItem
                            oWaves.GetSearchItem(xPartType, xItem(ItemIndex), xCSTPullerColor, xSearchItem)
                            ' 3. 尋找適合備料基準-從過濾箱中(M_LSOrderRule)取得適合備料基準
                            GetFilterRuleNo(pBuyer, xLTType, xPartType, xSearchItem, xRuleNo, xAction, xObjectType, xObjectProduct, xRuleCount)
                            ' 4. 依備料基準展開結構
                            For RuleIndex As Integer = 1 To xRuleCount
                                FCTWrite = False
                                ' 第一階結構資料是否OK
                                If (xObjectProduct(RuleIndex) = "FINISH" Or _
                                   (xObjectProduct(RuleIndex) = "E" And InStr(xItemName(ItemIndex), xObjectProduct(RuleIndex)) > 0)) And _
                                   oWaves.GetItemLineLine("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex)) = 0 Then         ' 不是LINE-LINE ITEM ((0=不是/1=是)
                                    ' 寫入 ForcastPlan
                                    sql = "SELECT * FROM ForcastPlan "
                                    sql &= "Where Buyer = '" & pBuyer & "' "
                                    sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                    sql &= "  And Y_Level = 0 "
                                    sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                    sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                    sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                    sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                    sql &= "  And C_B1 = '" & dt_FCTItem.Rows(i).Item("CUST") & "' "
                                    sql &= "  And C_CODE = '" & dt_FCTItem.Rows(i).Item("CITEM") & "' "
                                    sql &= "Order by FCTNo, FCTSubNo "
                                    Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                    For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                        WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "PS", xQtyMeter, xRuleNo(RuleIndex))
                                    Next
                                    xSubNo = xSubNo + 1
                                    FCTWrite = True
                                End If
                                ' 搜尋下一階結構資料
                                If Not FCTWrite Then
                                    oWaves.GetChildItemStructure("01", xObjectType(RuleIndex) + "-" + xObjectProduct(RuleIndex), xItem(ItemIndex), xItem1, xItemName1, xQty1, xCount1)
                                    '
                                    For ItemIndex1 As Integer = 1 To xCount1
                                        If oWaves.GetItemLineLine("01", dt_FCTItem.Rows(i).Item("Item"), xItem1(ItemIndex1)) = 0 Then          ' 不是LINE-LINE ITEM ((0=不是/1=是)
                                            ' 寫入 ForcastPlan
                                            sql = "SELECT * FROM ForcastPlan "
                                            sql &= "Where Buyer = '" & pBuyer & "' "
                                            sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                            sql &= "  And Y_Level = 0 "
                                            sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                            sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                            sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                            sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                            sql &= "  And C_B1 = '" & dt_FCTItem.Rows(i).Item("CUST") & "' "
                                            sql &= "  And C_CODE = '" & dt_FCTItem.Rows(i).Item("CITEM") & "' "
                                            sql &= "Order by FCTNo, FCTSubNo "
                                            Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                            For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                                WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, xItem(ItemIndex), xItem1(ItemIndex1), xColor, "PS", xQtyMeter, xRuleNo(RuleIndex))
                                            Next
                                            xSubNo = xSubNo + 1
                                            FCTWrite = True
                                        End If
                                    Next
                                End If
                                ' 搜尋不到任何結構資料(Set Error Rule [9900-990-99])
                                If Not FCTWrite Then
                                    ' 寫入 ForcastPlan
                                    sql = "SELECT * FROM ForcastPlan "
                                    sql &= "Where Buyer = '" & pBuyer & "' "
                                    sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                    sql &= "  And Y_Level = 0 "
                                    sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                    sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                    sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                    sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                    sql &= "  And C_B1 = '" & dt_FCTItem.Rows(i).Item("CUST") & "' "
                                    sql &= "  And C_CODE = '" & dt_FCTItem.Rows(i).Item("CITEM") & "' "
                                    sql &= "Order by FCTNo, FCTSubNo "
                                    Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                    For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                        WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "PS", xQtyMeter, xRuleNo(RuleIndex) + "/9900-990-00")
                                    Next
                                    xSubNo = xSubNo + 1
                                    FCTWrite = True
                                End If
                            Next
                        Next
                    End If
                End If
                '
                If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) = 2 Then          ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)
                    'MsgBox("CHAIN")
                    ' CHAIN 處理
                    oWaves.GetItemClass(dt_FCTItem.Rows(i).Item("Item"), xClass)
                    sql = "Select * From M_Referp "
                    sql = sql & " Where Cat  = '" & "110" & "' "
                    sql = sql & "   And DKey = '" & pBuyer + "-" + xClass & "' "
                    Dim dt_Referp As DataTable = oDataBase.GetDataTable(sql)
                    If dt_Referp.Rows.Count > 0 Then
                        xQtyMeter = CStr(CDbl(dt_Referp.Rows(0).Item("Data")) * 100)
                    Else
                        xQtyMeter = "100"
                    End If
                    ' 取得完成品CHAIN(CF/VF) 或 GAP-CHAIN(MF)     (MF=金屬/CF=樹脂/VF=塑鋼)
                    xPartType = "CH"                                                            ' 決定材料種類(PartTye)
                    oWaves.GetItemProduct(dt_FCTItem.Rows(i).Item("Item"), xItemProduct)        ' 取得製品區分 
                    '
                    xItem(1) = dt_FCTItem.Rows(i).Item("Item")
                    xItemName(1) = dt_FCTItem.Rows(i).Item("ItemName")
                    xQty(1) = CStr(1 * 10000000)
                    xCount = 1
                    For j As Integer = 2 To 5
                        xItem(j) = ""
                        xItemName(j) = ""
                        xQty(j) = "0"
                    Next
                    '
                    ' 展開ITEM構成取得所指定ITEM
                    For ItemIndex As Integer = 1 To xCount
                        ' 1. 取得完成品CHAIN兼用色
                        oWaves.GetChangeColor("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCTItem.Rows(i).Item("Color"), xColor)
                        ' 2. 準備備料基準相關SearchItem
                        oWaves.GetSearchItem(xPartType, xItem(ItemIndex), xCSTPullerColor, xSearchItem)
                        ' 3. 尋找適合備料基準-從過濾箱中(M_LSOrderRule)取得適合備料基準
                        GetFilterRuleNo(pBuyer, xLTType, xPartType, xSearchItem, xRuleNo, xAction, xObjectType, xObjectProduct, xRuleCount)
                        ' 4. 依備料基準展開結構
                        For RuleIndex As Integer = 1 To xRuleCount
                            FCTWrite = False
                            ' 第一階結構資料是否OK
                            If InStr(xItemName(ItemIndex), xObjectProduct(RuleIndex)) > 0 Then
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "  And C_B1 = '" & dt_FCTItem.Rows(i).Item("CUST") & "' "
                                sql &= "  And C_CODE = '" & dt_FCTItem.Rows(i).Item("CITEM") & "' "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "CH", xQtyMeter, xRuleNo(RuleIndex))
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                            ' 搜尋下一階結構資料
                            If Not FCTWrite Then
                                '
                                oWaves.GetChildItemStructure("01", xObjectType(RuleIndex) + "-" + xObjectProduct(RuleIndex), xItem(ItemIndex), xItem1, xItemName1, xQty1, xCount1)
                                '
                                For ItemIndex1 As Integer = 1 To xCount1
                                    If InStr(xItemName1(ItemIndex1), xObjectProduct(RuleIndex)) > 0 Then
                                        ' 寫入 ForcastPlan
                                        sql = "SELECT * FROM ForcastPlan "
                                        sql &= "Where Buyer = '" & pBuyer & "' "
                                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                        sql &= "  And Y_Level = 0 "
                                        sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                        sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                        sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                        sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                        sql &= "  And C_B1 = '" & dt_FCTItem.Rows(i).Item("CUST") & "' "
                                        sql &= "  And C_CODE = '" & dt_FCTItem.Rows(i).Item("CITEM") & "' "
                                        sql &= "Order by FCTNo, FCTSubNo "
                                        Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                        For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                            WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, xItem(ItemIndex), xItem1(ItemIndex1), xColor, "CH", xQtyMeter, xRuleNo(RuleIndex))
                                        Next
                                        xSubNo = xSubNo + 1
                                        FCTWrite = True
                                    End If
                                Next
                            End If
                            ' 搜尋不到任何結構資料(Set Error Rule [9900-990-99])
                            If Not FCTWrite Then
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "  And C_B1 = '" & dt_FCTItem.Rows(i).Item("CUST") & "' "
                                sql &= "  And C_CODE = '" & dt_FCTItem.Rows(i).Item("CITEM") & "' "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "CH", xQtyMeter, xRuleNo(RuleIndex) + "/" + "9900-990-00")
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                        Next
                    Next
                End If
                '
                If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) = 3 Then          ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)
                    'MsgBox("SLD")
                    '
                    ' SLIDER 處理  
                    '--------------------------------------------------------------------------------
                    xQtyMeter = "100"                       ' 取得Meter換算基準 (取得ItemClass)
                    xPartType = "SLD"                       ' 決定材料種類(PartTye)
                    '
                    ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)
                    xItem(1) = dt_FCTItem.Rows(i).Item("Item")
                    xItemName(1) = dt_FCTItem.Rows(i).Item("ItemName")
                    xQty(1) = CStr(1 * 10000000)
                    xCount = 1
                    For j As Integer = 2 To 5
                        xItem(j) = ""
                        xItemName(j) = ""
                        xQty(j) = "0"
                    Next
                    ' 展開ITEM構成取得所指定ITEM
                    For ItemIndex As Integer = 1 To xCount
                        ' 1. 取得完成品SLD兼用色
                        oWaves.GetChangeColor("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCTItem.Rows(i).Item("Color"), xColor)
                        ' 2. 準備備料基準相關SearchItem
                        oWaves.GetSearchItem(xPartType, xItem(ItemIndex), xCSTPullerColor, xSearchItem)
                        ' 3. 尋找適合備料基準-從過濾箱中(M_LSOrderRule)取得適合備料基準
                        GetFilterRuleNo(pBuyer, xLTType, xPartType, xSearchItem, xRuleNo, xAction, xObjectType, xObjectProduct, xRuleCount)
                        ' 4. 依備料基準展開結構
                        For RuleIndex As Integer = 1 To xRuleCount
                            FCTWrite = False
                            ' 第一階結構資料是否OK
                            If (xObjectProduct(RuleIndex) = "FINISH" Or _
                               (xObjectProduct(RuleIndex) = "E" And InStr(xItemName(ItemIndex), xObjectProduct(RuleIndex)) > 0)) And _
                               oWaves.GetItemLineLine("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex)) = 0 Then         ' 不是LINE-LINE ITEM ((0=不是/1=是)
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "  And C_B1 = '" & dt_FCTItem.Rows(i).Item("CUST") & "' "
                                sql &= "  And C_CODE = '" & dt_FCTItem.Rows(i).Item("CITEM") & "' "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "PS", xQtyMeter, xRuleNo(RuleIndex))
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                            ' 搜尋下一階結構資料
                            If Not FCTWrite Then
                                oWaves.GetChildItemStructure("01", xObjectType(RuleIndex) + "-" + xObjectProduct(RuleIndex), xItem(ItemIndex), xItem1, xItemName1, xQty1, xCount1)
                                '
                                For ItemIndex1 As Integer = 1 To xCount1
                                    If oWaves.GetItemLineLine("01", dt_FCTItem.Rows(i).Item("Item"), xItem1(ItemIndex1)) = 0 Then          ' 不是LINE-LINE ITEM ((0=不是/1=是)
                                        ' 寫入 ForcastPlan
                                        sql = "SELECT * FROM ForcastPlan "
                                        sql &= "Where Buyer = '" & pBuyer & "' "
                                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                        sql &= "  And Y_Level = 0 "
                                        sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                        sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                        sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                        sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                        sql &= "  And C_B1 = '" & dt_FCTItem.Rows(i).Item("CUST") & "' "
                                        sql &= "  And C_CODE = '" & dt_FCTItem.Rows(i).Item("CITEM") & "' "
                                        sql &= "Order by FCTNo, FCTSubNo "
                                        Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                        For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                            WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, xItem(ItemIndex), xItem1(ItemIndex1), xColor, "PS", xQtyMeter, xRuleNo(RuleIndex))
                                        Next
                                        xSubNo = xSubNo + 1
                                        FCTWrite = True
                                    End If
                                Next
                            End If
                            ' 搜尋不到任何結構資料(Set Error Rule [9900-990-99])
                            If Not FCTWrite Then
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "  And C_B1 = '" & dt_FCTItem.Rows(i).Item("CUST") & "' "
                                sql &= "  And C_CODE = '" & dt_FCTItem.Rows(i).Item("CITEM") & "' "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "PS", xQtyMeter, xRuleNo(RuleIndex) + "/9900-990-00")
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                        Next
                    Next
                End If
                ' ZIP是否為採購品  (0=不是/1=是)
                If Not FCTWrite Then
                    If oWaves.GetPurchaseItem(dt_FCTItem.Rows(i).Item("Item")) = 1 Then
                        ' 寫入 ForcastPlan
                        sql = "SELECT * FROM ForcastPlan "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                        sql &= "  And Y_Level = 0 "
                        sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                        sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                        sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                        sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                        sql &= "  And C_B1 = '" & dt_FCTItem.Rows(i).Item("CUST") & "' "
                        sql &= "  And C_CODE = '" & dt_FCTItem.Rows(i).Item("CITEM") & "' "
                        sql &= "Order by FCTNo, FCTSubNo "
                        Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                        For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                            WriteNewFCTPlan(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), dt_FCTItem.Rows(i).Item("Item"), dt_FCT2.Rows(j).Item("Y_Color").ToString, "ZIP", 100, "1000-010-00")
                        Next
                        xSubNo = xSubNo + 1
                        FCTWrite = True
                    End If
                End If
            Next
            '
            'MsgBox("UABAG-OUT")
            '
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "FCTPLAN", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "ForcastPlan", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'MatForcastPlanUABAG-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([120]-LocalStockPlan)   (Step-1)
    '**     Local Stock Plan展開 
    '***********************************************************************************************
    'LocalStockPlan-Start
    <WebMethod()> _
    Public Function LocalStockPlan(ByVal pLogID As String, _
                                   ByVal pBuyer As String, _
                                   ByVal pUserID As String, _
                                   ByVal pGRBuyer As String, _
                                   ByVal pFunList As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        '
        Try
            ' ***********************************************************************************
            ' 變數定義及設定初值
            ' ***********************************************************************************
            Dim i, j As Integer
            Dim xLastVersion As Integer = 99    ' 最終版PLAN
            Dim xSystemItem As String = ""
            Dim xGroupItem As String = ""
            Dim xSelectField As String = ""
            Dim xFCTString, xLSString As String
            Dim xFCTGroupField(), xLSGroupField() As String
            Dim xKeepCode As String = ""
            Dim xQty As String = "0"
            Dim xSumQty As Double = 0           ' N月生產+在庫合計
            Dim xProdQty As Double = 0          ' 需生產合計
            Dim xProdQtyUnit As Integer = 1     ' 生產數-數量單位
            Dim xMonScheProd(6), xMonOnProd(6) As String        ' N1 ~ N4 SCHE_PRODQTY, ON_PRODQTY

            oWaves.Timeout = Timeout.Infinite

            ' ---------------------------------------------------------------------------------
            ' 取得 SystemItem 初值
            sql = "SELECT * FROM M_ImportRule "
            sql &= "Where Buyer = '" & pBuyer & "-2' "
            sql &= "  And Sts   = '1' "
            sql &= "  And DataRule Like '" & "SY_" & "%' "
            sql &= "Order by DataRule "
            Dim dt_List As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_List.Rows.Count - 1
                ' SystemItem
                If xSystemItem = "" Then
                    xSystemItem = dt_List.Rows(i).Item("Field").ToString
                Else
                    xSystemItem = xSystemItem & "," & dt_List.Rows(i).Item("Field").ToString
                End If
                ' SelectField
                If xSelectField = "" Then
                    xSelectField = dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("Field").ToString
                Else
                    xSelectField = xSelectField & ", " & dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("Field").ToString
                End If
            Next
            ' 取得FCT & LS Group Field
            xFCTString = xSystemItem
            xLSString = xSystemItem
            ' ---------------------------------------------------------------------------------
            ' 取得 GroupItem 初值
            dt_List.Clear()
            sql = "SELECT * FROM M_ImportRule "
            sql &= "Where Buyer = '" & pBuyer + "-2" & "' "
            sql &= "  And Sts   = '1' "
            sql &= "  And DataRule Like '" & "GR_" & "%' "
            sql &= "Order by DataRule "
            dt_List = oDataBase.GetDataTable(sql)
            For i = 0 To dt_List.Rows.Count - 1
                ' GroupItem
                If xGroupItem = "" Then
                    xGroupItem = dt_List.Rows(i).Item("Field").ToString
                Else
                    xGroupItem = xGroupItem & "," & dt_List.Rows(i).Item("Field").ToString
                End If
                ' SelectField
                If xSelectField = "" Then
                    xSelectField = dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("DataRule").ToString
                Else
                    xSelectField = xSelectField & ", " & dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("DataRule").ToString
                End If
                ' 取得FCT & LS Group Field
                xFCTString = xFCTString & "," & dt_List.Rows(i).Item("Field").ToString
                xLSString = xLSString & "," & dt_List.Rows(i).Item("DataRule").ToString
            Next
            '
            ' 取得FCT & LS Group Field
            xFCTGroupField = xFCTString.Split(",")
            xLSGroupField = xLSString.Split(",")
            '
            ' GroupItem 不足10個時 ==> SelectField 需補至 10 個
            For j = i + 1 To 10
                If j < 10 Then
                    xSelectField = xSelectField & ", " & "'' AS GR_0" & CStr(j)
                Else
                    xSelectField = xSelectField & ", " & "'' AS GR_10"
                End If
            Next
            ' ---------------------------------------------------------------------------------
            ' 取得 SummaryItem 初值
            dt_List.Clear()
            sql = "SELECT * FROM M_ImportRule "
            sql &= "Where Buyer = '" & pBuyer + "-2" & "' "
            sql &= "  And Sts   = '1' "
            sql &= "  And DataRule Like '" & "FS_" & "%' "
            sql &= "Order by DataRule "
            dt_List = oDataBase.GetDataTable(sql)
            For i = 0 To dt_List.Rows.Count - 1
                ' SelectField
                If xSelectField = "" Then
                    xSelectField = "Sum(" & dt_List.Rows(i).Item("Field").ToString & ") AS " & dt_List.Rows(i).Item("DataRule").ToString
                Else
                    xSelectField = xSelectField & ", " & "Sum(" & dt_List.Rows(i).Item("Field").ToString & ") AS " & dt_List.Rows(i).Item("DataRule").ToString
                End If
            Next
            ' ***********************************************************************************
            ' 取得 LSNO = GroupCode + "L" + 年月 + 流水號(4碼)   SA + L + 135 + 0001
            ' 變數定義及設定初值
            ' ***********************************************************************************
            Dim xLSNo, xSeqNoString, xGroup, xItem As String
            Dim xSeqNo, xSubNo As Integer
            xGroup = "XX"
            xItem = ""
            ' 取得GroupCode
            sql = "SELECT GroupCode FROM M_FControlRecord "
            sql &= "Where Buyer = '" & pBuyer & "' "
            Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
            If dt_ControlRecord.Rows.Count > 0 Then xGroup = dt_ControlRecord.Rows(0).Item("GroupCode")
            ' 取得年月 
            Select Case Month(Now)
                Case 10
                    xLSNo = xGroup + "L" + CStr(Year(Now) - 2000) + "A"
                Case 11
                    xLSNo = xGroup + "L" + CStr(Year(Now) - 2000) + "B"
                Case 12
                    xLSNo = xGroup + "L" + CStr(Year(Now) - 2000) + "C"
                Case Else
                    xLSNo = xGroup + "L" + CStr(Year(Now) - 2000) + CStr(Month(Now))
            End Select
            ' 流水號(下一可使用No)及SubNo
            xSeqNo = GetLSSeqNo(pBuyer, xLSNo)
            xSubNo = 1
            '
            ' ***********************************************************************************
            ' Local Stock Plan展開
            ' ***********************************************************************************
            sql = "SELECT "
            sql &= xSelectField
            sql &= " FROM ForcastPlan "
            sql &= " Where Buyer = '" & pBuyer & "' "
            sql &= "   And Version = " & xLastVersion & " "
            sql &= "   And Y_Level > 0 "
            ' test-start 測試使用 
            'sql &= "   And Y_itemcode = '0003203' "
            'sql &= "   And C_SHORTENLT = '0' "
            ' test-end 
            sql &= " Group by " & xSystemItem & ", " & xGroupItem
            sql &= " Order by " & xSystemItem & ", " & xGroupItem
            Dim dt_FCT As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_FCT.Rows.Count - 1
                ' Insert LocalStock Plan Data
                sql = "Insert into LocalStockPlan "
                sql &= "( "
                sql &= "BUYER, "
                sql &= "BUYERGROUP, "
                sql &= "LSNO, "
                sql &= "LSSUBNO, "
                sql &= "VERSION, "

                sql &= "GR_01, "
                sql &= "GR_02, "
                sql &= "GR_03, "
                sql &= "GR_04, "
                sql &= "GR_05, "
                sql &= "GR_06, "
                sql &= "GR_07, "
                sql &= "GR_08, "
                sql &= "GR_09, "
                sql &= "GR_10, "

                sql &= "MINIMUMSTOCK, "
                sql &= "N_SCHEPROD, "
                sql &= "N_ONPROD, "
                sql &= "N_FREEINV, "
                sql &= "N_KEEPINV, "
                sql &= "N_TOTAL, "

                sql &= "FS_01, "
                sql &= "PS_01, "
                sql &= "IS_01, "
                sql &= "FS_02, "
                sql &= "PS_02, "
                sql &= "IS_02, "
                sql &= "FS_03, "
                sql &= "PS_03, "
                sql &= "IS_03, "
                sql &= "FS_04, "
                sql &= "PS_04, "
                sql &= "IS_04, "
                sql &= "FS_05, "
                sql &= "PS_05, "
                sql &= "IS_05, "
                sql &= "FS_06, "
                sql &= "PS_06, "
                sql &= "IS_06, "
                sql &= "FS_07, "
                sql &= "PS_07, "
                sql &= "IS_07, "
                sql &= "FS_08, "
                sql &= "PS_08, "
                sql &= "IS_08, "
                sql &= "FS_09, "
                sql &= "PS_09, "
                sql &= "IS_09, "
                sql &= "FS_10, "
                sql &= "PS_10, "
                sql &= "IS_10, "
                sql &= "FS_11, "
                sql &= "PS_11, "
                sql &= "IS_11, "
                sql &= "FS_12, "
                sql &= "PS_12, "
                sql &= "IS_12, "

                sql &= "FS_13, "
                sql &= "PS_13, "

                sql &= "CreateUser, "
                sql &= "CreateTime "
                sql &= " ) "
                sql &= "VALUES( "
                sql &= " '" & dt_FCT.Rows(i).Item("Buyer").ToString & "', "                 ' Buyer
                sql &= " '" & dt_FCT.Rows(i).Item("BuyerGroup").ToString & "', "            ' BuyerGroup
                ' ---------------------------------------------------------------------------------
                ' 設定 LSNo, SubNo
                ' Item Code是否相同 Counter LSNo, SubNo
                If dt_FCT.Rows(i).Item("GR_03").ToString <> xItem Then
                    If i > 0 Then
                        xSeqNo = xSeqNo + 1
                        xSubNo = 1
                    End If
                    xItem = dt_FCT.Rows(i).Item("GR_03").ToString
                Else
                    If i > 0 Then xSubNo = xSubNo + 1
                End If
                ' 製作SeqNo(不足4位補0)
                xSeqNoString = CStr(xSeqNo)
                If Len(xSeqNoString) < 2 Then
                    xSeqNoString = "000" + CStr(xSeqNo)
                Else
                    If Len(xSeqNoString) < 3 Then
                        xSeqNoString = "00" + CStr(xSeqNo)
                    Else
                        If Len(xSeqNoString) < 4 Then
                            xSeqNoString = "0" + CStr(xSeqNo)
                        End If
                    End If
                End If
                '設定 LSNo, SubNo
                sql &= " '" & xLSNo + xSeqNoString & "', "                                  ' LSNO
                sql &= " " & CStr(xSubNo) & ", "                                            ' LSSUBNO
                ' ---------------------------------------------------------------------------------
                sql &= " " & CStr(xLastVersion) & ", "                                      ' Version

                sql &= " '" & dt_FCT.Rows(i).Item("GR_01").ToString & "', "                 ' GR_01
                sql &= " '" & dt_FCT.Rows(i).Item("GR_02").ToString & "', "                 ' GR_02
                sql &= " '" & dt_FCT.Rows(i).Item("GR_03").ToString & "', "                 ' GR_03
                sql &= " '" & dt_FCT.Rows(i).Item("GR_04").ToString & "', "                 ' GR_04
                sql &= " '" & dt_FCT.Rows(i).Item("GR_05").ToString & "', "                 ' GR_05
                sql &= " '" & dt_FCT.Rows(i).Item("GR_06").ToString & "', "                 ' GR_06
                sql &= " '" & dt_FCT.Rows(i).Item("GR_07").ToString & "', "                 ' GR_07
                sql &= " '" & dt_FCT.Rows(i).Item("GR_08").ToString & "', "                 ' GR_08
                sql &= " '" & dt_FCT.Rows(i).Item("GR_09").ToString & "', "                 ' GR_09
                sql &= " '" & dt_FCT.Rows(i).Item("GR_10").ToString & "', "                 ' GR_10
                ' ---------------------------------------------------------------------------------
                ' 取得 MinimumStock Qty
                oWaves.GetMininumStock("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xQty)
                If CDbl(xQty) > 0 Then
                    sql &= " " & "1" & ", "
                Else
                    sql &= " " & "0" & ", "
                End If
                ' ---------------------------------------------------------------------------------
                ' 取得 生產中Qty (SCHE/ON) , 在庫Qty (Free/Keep)
                ' 取得 KeepCode
                'xKeepCode = ""
                'dt_List.Clear()
                'sql2 = "SELECT Data FROM M_Referp "
                'sql2 &= "Where Cat = '" & "111" & "' "
                'sql2 &= "  And DKey = '" & pGRBuyer + "-" + dt_FCT.Rows(i).Item("GR_05").ToString & "' "
                'sql2 &= "Order by Data "
                'dt_List = oDataBase.GetDataTable(sql2)
                'If dt_List.Rows.Count > 0 Then
                '    xKeepCode = dt_List.Rows(0).Item("Data").ToString
                'End If
                xKeepCode = dt_FCT.Rows(i).Item("GR_02").ToString
                '
                xSumQty = 0
                xProdQty = 0
                '
                ' 生產-SCHE Qty (N_SCHEPROD)
                oWaves.GetProductionQty("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xKeepCode, 0, xQty, xMonScheProd)
                If CDbl(xQty) > 0 Then
                    sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                    xSumQty = xSumQty + CDbl(xQty) / 10000000
                Else
                    sql &= " " & "0" & ", "
                End If
                '
                ' 生產-ON Qty (N_ONPROD)
                oWaves.GetProductionQty("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xKeepCode, 1, xQty, xMonOnProd)
                If CDbl(xQty) > 0 Then
                    sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                    xSumQty = xSumQty + CDbl(xQty) / 10000000
                Else
                    sql &= " " & "0" & ", "
                End If
                '
                ' 在庫-Free Qty (N_FREEINV)
                oWaves.GetFreeInventory("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xQty)
                If CDbl(xQty) > 0 Then
                    sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                    xSumQty = xSumQty + CDbl(xQty) / 10000000
                Else
                    sql &= " " & "0" & ", "
                End If
                '
                ' 在庫-Keep Qty (N_KEEPINV)
                oWaves.GetKeepCodeInventory("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xKeepCode, xQty)
                If CDbl(xQty) > 0 Then
                    sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                    xSumQty = xSumQty + CDbl(xQty) / 10000000
                Else
                    sql &= " " & "0" & ", "
                End If
                '
                ' 合計=(生產-SCHE Qty + 生產-ON Qty + 在庫-Free Qty +在庫-Keep Qty) (N_TOTAL)
                sql &= " " & CStr(xSumQty) & ", "
                ' ---------------------------------------------------------------------------------
                ' FORCAST & PRODUCTION & INVENTORY
                '
                ' FS_01 / PS_01 / IS_01
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_01")
                sql &= " " & dt_FCT.Rows(i).Item("FS_01").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_02 / PS_02 / IS_02
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_02")
                sql &= " " & dt_FCT.Rows(i).Item("FS_02").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_03 / PS_03 / IS_03
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_03")
                sql &= " " & dt_FCT.Rows(i).Item("FS_03").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_04 / PS_04 / IS_04
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_04")
                sql &= " " & dt_FCT.Rows(i).Item("FS_04").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_05 / PS_05 / IS_05
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_05")
                sql &= " " & dt_FCT.Rows(i).Item("FS_05").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_06 / PS_06 / IS_06
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_06")
                sql &= " " & dt_FCT.Rows(i).Item("FS_06").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_07 / PS_07 / IS_07
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_07")
                sql &= " " & dt_FCT.Rows(i).Item("FS_07").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_08 / PS_08 / IS_08
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_08")
                sql &= " " & dt_FCT.Rows(i).Item("FS_08").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_09 / PS_09 / IS_09
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_09")
                sql &= " " & dt_FCT.Rows(i).Item("FS_09").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_10 / PS_10 / IS_10
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_10")
                sql &= " " & dt_FCT.Rows(i).Item("FS_10").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_11 / PS_11 / IS_11
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_11")
                sql &= " " & dt_FCT.Rows(i).Item("FS_11").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_12 / PS_12 / IS_12
                xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_12")
                sql &= " " & dt_FCT.Rows(i).Item("FS_12").ToString & ", "
                If xSumQty >= 0 Then
                    sql &= " " & "0" & ", "
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    sql &= " " & CStr(xSumQty * -1) & ", "
                    sql &= " " & "0" & ", "
                    xProdQty = xProdQty + (xSumQty * -1)
                    xSumQty = 0
                End If
                '
                ' FS_13 / PS_13
                sql &= " " & dt_FCT.Rows(i).Item("FS_13").ToString & ", "                   ' FS_13
                '
                ' 數量單位整合 (目前不整合)
                xProdQtyUnit = 1
                'dt_List.Clear()
                'sql2 = "SELECT Data FROM M_Referp "
                'sql2 &= "Where Cat = '" & "112" & "' "
                'sql2 &= "  And DKey = '" & dt_FCT.Rows(i).Item("GR_04").ToString & "' "
                'sql2 &= "Order by Data "
                'dt_List = oDataBase.GetDataTable(sql2)
                'If dt_List.Rows.Count > 0 Then
                '    xProdQtyUnit = CInt(dt_List.Rows(0).Item("Data"))
                'End If
                ''
                'If xProdQty Mod xProdQtyUnit > 0 Then
                '    xProdQty = (Fix(xProdQty / xProdQtyUnit) + 1) * xProdQtyUnit
                'End If
                sql &= " " & CStr(xProdQty) & ", "                                          ' PS_13
                '
                sql &= " '" & "LSPlan" & "', "
                sql &= " '" & NowDateTime & "' "
                sql &= " ) "
                oDataBase.ExecuteNonQuery(sql)
                '
                ' ---------------------------------------------------------------------------------
                ' Update FCT Plan Data (LSNo, LSSubNo)
                sql = "Update ForcastPlan Set "
                sql &= "LSNO = '" & xLSNo + xSeqNoString & "', "
                sql &= "LSSUBNO = " & CStr(xSubNo) & " "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "  And Version = " & xLastVersion & " "
                sql &= "  And Y_Level > 0 "
                For j = 0 To UBound(xFCTGroupField)
                    sql &= " And " & xFCTGroupField(j) & " = '" & dt_FCT.Rows(i).Item(xLSGroupField(j)).ToString & "' "
                Next
                oDataBase.ExecuteNonQuery(sql)
            Next
            ' ***********************************************************************************
            ' 更新下一個可使用No
            ' ***********************************************************************************
            ' 製作SeqNo(不足4位補0)
            xSeqNo = xSeqNo + 1
            xSeqNoString = CStr(xSeqNo)
            If Len(xSeqNoString) < 2 Then
                xSeqNoString = "000" + CStr(xSeqNo)
            Else
                If Len(xSeqNoString) < 3 Then
                    xSeqNoString = "00" + CStr(xSeqNo)
                Else
                    If Len(xSeqNoString) < 4 Then
                        xSeqNoString = "0" + CStr(xSeqNo)
                    End If
                End If
            End If
            ' 更新下一次可使用PONO
            UpdateNextLSNo(pBuyer, xLSNo + xSeqNoString)
            '
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "LSPLAN", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "LocalStockPlan", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'LocalStockPlan-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([122]-NewLocalStockPlan)    (Step-2)
    '**     New Local Stock Plan展開 
    '***********************************************************************************************
    'NewLocalStockPlan-Start
    <WebMethod()> _
    Public Function NewLocalStockPlan(ByVal pLogID As String, _
                                      ByVal pBuyer As String, _
                                      ByVal pUserID As String, _
                                      ByVal pGRBuyer As String, _
                                      ByVal pFunList As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        '
        'Try
        ' ***********************************************************************************
        ' 變數定義及設定初值
        ' ***********************************************************************************
        Dim i, j As Integer
        Dim xLastVersion As Integer = 99    ' 最終版PLAN
        Dim xSystemItem As String = ""
        Dim xGroupItem As String = ""
        Dim xSelectField As String = ""
        Dim xFCTString, xLSString, xDescr As String
        Dim xFCTGroupField(), xLSGroupField() As String
        Dim xKeepCode As String = ""
        Dim xQty As String = "0"
        Dim xSumQty As Double = 0           ' N月生產+在庫合計
        Dim xKeepQty As Double = 0          ' SCHE-PROD + ON-PROD + KEEP-INV
        Dim xFreeQty As Double = 0          ' FREE-INV
        Dim xFCTQty As Double = 0           ' ORDER ZONE FCT QTY
        Dim xFCTSumQty As Double = 0        ' 合計FCT QTY 
        Dim xProdQty As Double = 0          ' 需生產合計
        Dim xMonScheProd(6), xMonOnProd(6) As String        ' N1 ~ N4 SCHE_PRODQTY, ON_PRODQTY
        Dim xPurchaseItem As Boolean        ' 採購品
        Dim xPriceA As String = "0"         ' A單價
        Dim xPriceB As String = "0"         ' B單價
        Dim xPrice As Double = 0            ' 單價
        Dim xMetterLength As Double = 0     ' Length(公呎單位)

        oWaves.Timeout = Timeout.Infinite
        ' ---------------------------------------------------------------------------------
        ' 取得 SystemItem 初值
        sql = "SELECT * FROM M_ImportRule "
        sql &= "Where Buyer = '" & pBuyer & "-2' "
        sql &= "  And Sts   = '1' "
        sql &= "  And DataRule Like '" & "SY_" & "%' "
        sql &= "Order by DataRule "
        Dim dt_List As DataTable = oDataBase.GetDataTable(sql)
        For i = 0 To dt_List.Rows.Count - 1
            ' SystemItem
            If xSystemItem = "" Then
                xSystemItem = dt_List.Rows(i).Item("Field").ToString
            Else
                xSystemItem = xSystemItem & "," & dt_List.Rows(i).Item("Field").ToString
            End If
            ' SelectField
            If xSelectField = "" Then
                xSelectField = dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("Field").ToString
            Else
                xSelectField = xSelectField & ", " & dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("Field").ToString
            End If
        Next
        ' 取得FCT & LS Group Field
        xFCTString = xSystemItem
        xLSString = xSystemItem
        ' ---------------------------------------------------------------------------------
        ' 取得 GroupItem 初值
        dt_List.Clear()
        sql = "SELECT * FROM M_ImportRule "
        sql &= "Where Buyer = '" & pBuyer + "-2" & "' "
        sql &= "  And Sts   = '1' "
        sql &= "  And DataRule Like '" & "GR_" & "%' "
        sql &= "Order by DataRule "
        dt_List = oDataBase.GetDataTable(sql)
        For i = 0 To dt_List.Rows.Count - 1
            ' GroupItem
            If xGroupItem = "" Then
                xGroupItem = dt_List.Rows(i).Item("Field").ToString
            Else
                xGroupItem = xGroupItem & "," & dt_List.Rows(i).Item("Field").ToString
            End If
            ' SelectField
            If xSelectField = "" Then
                xSelectField = dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("DataRule").ToString
            Else
                xSelectField = xSelectField & ", " & dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("DataRule").ToString
            End If
            ' 取得FCT & LS Group Field
            xFCTString = xFCTString & "," & dt_List.Rows(i).Item("Field").ToString
            xLSString = xLSString & "," & dt_List.Rows(i).Item("DataRule").ToString
        Next
        '
        ' 取得FCT & LS Group Field
        xFCTGroupField = xFCTString.Split(",")
        xLSGroupField = xLSString.Split(",")
        '
        ' GroupItem 不足10個時 ==> SelectField 需補至 10 個
        For j = i + 1 To 10
            If j < 10 Then
                xSelectField = xSelectField & ", " & "'' AS GR_0" & CStr(j)
            Else
                xSelectField = xSelectField & ", " & "'' AS GR_10"
            End If
        Next
        ' ---------------------------------------------------------------------------------
        ' 取得 SummaryItem 初值
        dt_List.Clear()
        sql = "SELECT * FROM M_ImportRule "
        sql &= "Where Buyer = '" & pBuyer + "-2" & "' "
        sql &= "  And Sts   = '1' "
        sql &= "  And DataRule Like '" & "FS_" & "%' "
        sql &= "Order by DataRule "
        dt_List = oDataBase.GetDataTable(sql)
        For i = 0 To dt_List.Rows.Count - 1
            ' SelectField
            If xSelectField = "" Then
                xSelectField = "Sum(" & dt_List.Rows(i).Item("Field").ToString & ") AS " & dt_List.Rows(i).Item("DataRule").ToString
            Else
                xSelectField = xSelectField & ", " & "Sum(" & dt_List.Rows(i).Item("Field").ToString & ") AS " & dt_List.Rows(i).Item("DataRule").ToString
            End If
        Next
        ' ***********************************************************************************
        ' 取得 LSNO = GroupCode + "L" + 年月 + 流水號(4碼)   SA + L + 135 + 0001
        ' 變數定義及設定初值
        ' ***********************************************************************************
        Dim xLSNo, xSeqNoString, xGroup, xItem As String
        Dim xSeqNo, xSubNo As Integer
        xGroup = "XX"
        xItem = ""
        ' 取得GroupCode
        sql = "SELECT GroupCode FROM M_FControlRecord "
        sql &= "Where Buyer = '" & pBuyer & "' "
        Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
        If dt_ControlRecord.Rows.Count > 0 Then xGroup = dt_ControlRecord.Rows(0).Item("GroupCode")
        ' 取得年月 
        Select Case Month(Now)
            Case 10
                xLSNo = xGroup + "L" + CStr(Year(Now) - 2000) + "A"
            Case 11
                xLSNo = xGroup + "L" + CStr(Year(Now) - 2000) + "B"
            Case 12
                xLSNo = xGroup + "L" + CStr(Year(Now) - 2000) + "C"
            Case Else
                xLSNo = xGroup + "L" + CStr(Year(Now) - 2000) + CStr(Month(Now))
        End Select
        ' 流水號(下一可使用No)及SubNo
        xSeqNo = GetLSSeqNo(pBuyer, xLSNo)
        xSubNo = 1
        '
        ' ***********************************************************************************
        ' Local Stock Plan展開
        ' ***********************************************************************************
        sql = "SELECT "
        sql &= xSelectField
        sql &= " FROM ForcastPlan "
        sql &= " Where Buyer = '" & pBuyer & "' "
        sql &= "   And Version = " & xLastVersion & " "
        sql &= "   And Y_Level > 0 "
        'VENDOR-FC 17/11/24 ADD
        sql &= "   And Y_J1 <> '" & "*" & "' "
        sql &= " Group by " & xSystemItem & ", " & xGroupItem
        sql &= " Order by " & xSystemItem & ", " & xGroupItem
        Dim dt_FCT As DataTable = oDataBase.GetDataTable(sql)
        For i = 0 To dt_FCT.Rows.Count - 1
            ' Insert LocalStock Plan Data


            sql = "Insert into LocalStockPlan "
            sql &= "( "
            sql &= "BUYER, "
            sql &= "BUYERGROUP, "
            sql &= "LSNO, "
            sql &= "LSSUBNO, "
            sql &= "VERSION, "

            sql &= "GR_01, "
            sql &= "GR_02, "
            sql &= "GR_03, "
            sql &= "GR_04, "
            sql &= "GR_05, "
            sql &= "GR_06, "
            sql &= "GR_07, "
            sql &= "GR_08, "
            sql &= "GR_09, "
            sql &= "GR_10, "

            sql &= "MINIMUMSTOCK, "
            sql &= "N_SCHEPROD, "
            sql &= "N_ONPROD, "
            sql &= "N_FREEINV, "
            sql &= "N_KEEPINV, "
            sql &= "N_TOTAL, "

            sql &= "SP_00, "        'ADD-START 13/11/11
            sql &= "OP_00, "        'ADD-START 13/11/11
            sql &= "FS_00, "        'ADD-START 13/12/4
            sql &= "PS_00, "        'ADD-START 13/12/4
            sql &= "IS_00, "        'ADD-START 13/12/4
            sql &= "SP_01, "        'ADD-START 13/11/11
            sql &= "OP_01, "        'ADD-START 13/11/11
            sql &= "FS_01, "
            sql &= "PS_01, "
            sql &= "IS_01, "
            sql &= "SP_02, "        'ADD-START 13/11/11
            sql &= "OP_02, "        'ADD-START 13/11/11
            sql &= "FS_02, "
            sql &= "PS_02, "
            sql &= "IS_02, "
            sql &= "SP_03, "        'ADD-START 13/11/11
            sql &= "OP_03, "        'ADD-START 13/11/11
            sql &= "FS_03, "
            sql &= "PS_03, "
            sql &= "IS_03, "
            sql &= "SP_04, "        'ADD-START 13/11/11
            sql &= "OP_04, "        'ADD-START 13/11/11
            sql &= "FS_04, "
            sql &= "PS_04, "
            sql &= "IS_04, "
            sql &= "SP_05, "        'ADD-START 13/11/11
            sql &= "OP_05, "        'ADD-START 13/11/11
            sql &= "FS_05, "
            sql &= "PS_05, "
            sql &= "IS_05, "
            sql &= "FS_06, "
            sql &= "PS_06, "
            sql &= "IS_06, "
            sql &= "FS_07, "
            sql &= "PS_07, "
            sql &= "IS_07, "
            sql &= "FS_08, "
            sql &= "PS_08, "
            sql &= "IS_08, "
            sql &= "FS_09, "
            sql &= "PS_09, "
            sql &= "IS_09, "
            sql &= "FS_10, "
            sql &= "PS_10, "
            sql &= "IS_10, "
            sql &= "FS_11, "
            sql &= "PS_11, "
            sql &= "IS_11, "
            sql &= "FS_12, "
            sql &= "PS_12, "
            sql &= "IS_12, "

            sql &= "FS_13, "
            sql &= "PS_13, "
            sql &= "FREEALLOC, "
            sql &= "DESCRIPTION, "

            sql &= "CreateUser, "
            sql &= "CreateTime "
            sql &= " ) "
            sql &= "VALUES( "
            sql &= " '" & dt_FCT.Rows(i).Item("Buyer").ToString & "', "                 ' Buyer
            sql &= " '" & dt_FCT.Rows(i).Item("BuyerGroup").ToString & "', "            ' BuyerGroup
            ' ---------------------------------------------------------------------------------
            ' 設定 LSNo, SubNo
            ' Item Code是否相同 Counter LSNo, SubNo
            If dt_FCT.Rows(i).Item("GR_03").ToString <> xItem Then
                If i > 0 Then
                    xSeqNo = xSeqNo + 1
                    xSubNo = 1
                End If
                xItem = dt_FCT.Rows(i).Item("GR_03").ToString
            Else
                If i > 0 Then xSubNo = xSubNo + 1
            End If

            ' 製作SeqNo(不足4位補0)
            xSeqNoString = CStr(xSeqNo)
            If Len(xSeqNoString) < 2 Then
                xSeqNoString = "000" + CStr(xSeqNo)
            Else
                If Len(xSeqNoString) < 3 Then
                    xSeqNoString = "00" + CStr(xSeqNo)
                Else
                    If Len(xSeqNoString) < 4 Then
                        xSeqNoString = "0" + CStr(xSeqNo)
                    End If
                End If
            End If

            '設定 LSNo, SubNo
            sql &= " '" & xLSNo + xSeqNoString & "', "                                  ' LSNO
            sql &= " " & CStr(xSubNo) & ", "                                            ' LSSUBNO
            ' ---------------------------------------------------------------------------------
            sql &= " " & CStr(xLastVersion) & ", "                                      ' Version

            sql &= " '" & dt_FCT.Rows(i).Item("GR_01").ToString & "', "                 ' GR_01
            sql &= " '" & dt_FCT.Rows(i).Item("GR_02").ToString & "', "                 ' GR_02
            sql &= " '" & dt_FCT.Rows(i).Item("GR_03").ToString & "', "                 ' GR_03
            sql &= " '" & dt_FCT.Rows(i).Item("GR_04").ToString & "', "                 ' GR_04
            sql &= " '" & dt_FCT.Rows(i).Item("GR_05").ToString & "', "                 ' GR_05
            sql &= " '" & dt_FCT.Rows(i).Item("GR_06").ToString & "', "                 ' GR_06
            sql &= " '" & dt_FCT.Rows(i).Item("GR_07").ToString & "', "                 ' GR_07
            sql &= " '" & dt_FCT.Rows(i).Item("GR_08").ToString & "', "                 ' GR_08
            sql &= " '" & dt_FCT.Rows(i).Item("GR_09").ToString & "', "                 ' GR_09
            sql &= " '" & dt_FCT.Rows(i).Item("GR_10").ToString & "', "                 ' GR_10
            ' ---------------------------------------------------------------------------------

            If dt_FCT.Rows(i).Item("GR_08").ToString <> "ZIP" Then
                '** PS / CH
                ' ---------------------------------------------------------------------------------
                ' 取得 MinimumStock Qty
                oWaves.GetMininumStock("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xQty)
                If CDbl(xQty) > 0 Then
                    sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                Else
                    sql &= " " & "0" & ", "
                End If
                ' ---------------------------------------------------------------------------------
                ' 歸零作業
                xSumQty = 0
                xProdQty = 0
                xKeepQty = 0            ' SCHE-PROD + ON-PROD + KEEP-INV
                xFreeQty = 0            ' FREE-INV
                xFCTQty = 0             ' ORDER ZONE FCT QTY
                xDescr = ""             ' Free Qty備註說明
                For j = 1 To 6          ' N1~N4 PROD-QTY
                    xMonScheProd(j) = 0
                    xMonOnProd(j) = 0
                Next
                ' 判斷是否採購品(0 = 不是 / 1 = 是)
                xPurchaseItem = False
                If oWaves.GetPurchaseItem(dt_FCT.Rows(i).Item("GR_03").ToString) = 1 Then
                    xPurchaseItem = True
                End If
                ' ---------------------------------------------------------------------------------
                If xPurchaseItem Then
                    ' 採購品
                    ' 不使用 (N_SCHEPROD)
                    sql &= " " & "0" & ", "
                    '
                    ' 採購-ON Qty (N_ONPROD)
                    xKeepCode = dt_FCT.Rows(i).Item("GR_02").ToString

                    oWaves.GetPurchaseQty("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xKeepCode, xQty, xMonOnProd)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                Else
                    ' 製造品
                    ' 生產-SCHE Qty (N_SCHEPROD)
                    xKeepCode = dt_FCT.Rows(i).Item("GR_02").ToString

                    oWaves.GetProductionQty("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xKeepCode, 0, xQty, xMonScheProd)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                    '
                    ' 生產-ON Qty (N_ONPROD)
                    oWaves.GetProductionQty("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xKeepCode, 1, xQty, xMonOnProd)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                End If
                '
                ' 在庫-Free Qty (N_FREEINV)
                oWaves.GetFreeInventory("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xQty)
                If CDbl(xQty) > 0 Then
                    sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                    xFreeQty = xFreeQty + CDbl(xQty) / 10000000
                    xSumQty = xSumQty + CDbl(xQty) / 10000000
                Else
                    sql &= " " & "0" & ", "
                End If
                '
                ' 在庫-Keep Qty (N_KEEPINV)
                oWaves.GetKeepCodeInventory("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xKeepCode, xQty)
                If CDbl(xQty) > 0 Then
                    sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                    xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                    xSumQty = xSumQty + CDbl(xQty) / 10000000
                Else
                    sql &= " " & "0" & ", "
                End If
                '
                ' 合計=(生產-SCHE Qty + 生產-ON Qty + 在庫-Free Qty +在庫-Keep Qty) (N_TOTAL)
                sql &= " " & CStr(xSumQty) & ", "
            Else
                '** VNDOR-FC ADD 17/11/24
                '** ZIP 
                ' ---------------------------------------------------------------------------------
                ' 取得 MinimumStock Qty
                oWaves.GetMininumStockZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xQty)
                If CDbl(xQty) > 0 Then
                    sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                Else
                    sql &= " " & "0" & ", "
                End If
                ' ---------------------------------------------------------------------------------
                ' 歸零作業
                xSumQty = 0
                xProdQty = 0
                xKeepQty = 0            ' SCHE-PROD + ON-PROD + KEEP-INV
                xFreeQty = 0            ' FREE-INV
                xFCTQty = 0             ' ORDER ZONE FCT QTY
                xDescr = ""             ' Free Qty備註說明
                For j = 1 To 6          ' N1~N4 PROD-QTY
                    xMonScheProd(j) = 0
                    xMonOnProd(j) = 0
                Next
                ' 判斷是否採購品(0 = 不是 / 1 = 是)
                xPurchaseItem = False

                If oWaves.GetPurchaseItem(dt_FCT.Rows(i).Item("GR_03").ToString) = 1 Then
                    xPurchaseItem = True
                End If
                ' ---------------------------------------------------------------------------------
                If xPurchaseItem Then
                    ' 採購品
                    ' 不使用 (N_SCHEPROD)
                    sql &= " " & "0" & ", "
                    '
                    ' 採購-ON Qty (N_ONPROD)
                    xKeepCode = dt_FCT.Rows(i).Item("GR_02").ToString
                    oWaves.GetPurchaseQtyZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xKeepCode, xQty, xMonOnProd)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                Else
                    ' 製造品
                    ' 生產-SCHE Qty (N_SCHEPROD)
                    xKeepCode = dt_FCT.Rows(i).Item("GR_02").ToString

                    oWaves.GetProductionQtyZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xKeepCode, 0, xQty, xMonScheProd)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                    '
                    ' 生產-ON Qty (N_ONPROD)
                    oWaves.GetProductionQtyZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xKeepCode, 1, xQty, xMonOnProd)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                End If
                '
                ' 在庫-Free Qty (N_FREEINV)
                oWaves.GetFreeInventoryZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xQty)
                If CDbl(xQty) > 0 Then
                    sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                    xFreeQty = xFreeQty + CDbl(xQty) / 10000000
                    xSumQty = xSumQty + CDbl(xQty) / 10000000
                Else
                    sql &= " " & "0" & ", "
                End If
                '
                ' 在庫-Keep Qty (N_KEEPINV)
                oWaves.GetKeepCodeInventoryZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xKeepCode, xQty)
                If CDbl(xQty) > 0 Then
                    sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                    xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                    xSumQty = xSumQty + CDbl(xQty) / 10000000
                Else
                    sql &= " " & "0" & ", "
                End If
                '
                ' 合計=(生產-SCHE Qty + 生產-ON Qty + 在庫-Free Qty +在庫-Keep Qty) (N_TOTAL)
                sql &= " " & CStr(xSumQty) & ", "
            End If
            ' ---------------------------------------------------------------------------------
            ' FORCAST & PRODUCTION & INVENTORY
            '
            ' ---------------------------------------------------------------------------------
            ' SP00 / OP00 / FS_00 / PS_00 / IS_00
            ' ** MOD-START 17/12/13 VENDOR FC
            If InStr(pBuyer, "F-VENDOR") > 0 Or InStr(pBuyer, "FALL-VENDOR") > 0 Then
                ' SP00:單價 / OP00:金額
                xPriceA = "0"         ' A單價
                xPriceB = "0"         ' B單價
                oWaves.GetCostPrice(dt_FCT.Rows(i).Item("GR_03").ToString, xPriceA, xPriceB)
                If CDbl(xPriceA) > 0 Or CDbl(xPriceB) > 0 Then
                    If dt_FCT.Rows(i).Item("GR_08").ToString = "ZIP" Then
                        If dt_FCT.Rows(i).Item("GR_10").ToString = "C" Then
                            xMetterLength = CDbl(dt_FCT.Rows(i).Item("GR_09").ToString) / 100
                        Else
                            xMetterLength = CDbl(dt_FCT.Rows(i).Item("GR_09").ToString) * 2.54 / 100
                        End If
                        xPrice = xMetterLength * CDbl(xPriceA) + CDbl(xPriceB)
                    Else
                        xPrice = CDbl(xPriceA) + CDbl(xPriceB)
                    End If
                    sql &= " " & CStr(xPrice) & ", "
                    sql &= " " & CStr(xKeepQty * xPrice) & ", "
                Else
                    sql &= " " & "0" & ", "
                    sql &= " " & "0" & ", "
                End If
            Else
                sql &= " " & CStr(CDbl(xMonScheProd(1)) / 10000000) & ", "
                sql &= " " & CStr(CDbl(xMonOnProd(1)) / 10000000) & ", "
            End If
            ' --
            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_00")
            ' 
            ' --ERIC&PEGGY 追加 2020/09/23
            xFCTQty = xFCTQty + dt_FCT.Rows(i).Item("FS_00")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_00").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' SP01 / OP01 / FS_01 / PS_01 / IS_01
            sql &= " " & CStr(CDbl(xMonScheProd(2)) / 10000000) & ", "
            sql &= " " & CStr(CDbl(xMonOnProd(2)) / 10000000) & ", "

            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_01")
            xFCTQty = xFCTQty + dt_FCT.Rows(i).Item("FS_01")

            sql &= " " & CStr(dt_FCT.Rows(i).Item("FS_01")) & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
                xFCTQty = xFCTQty - xSumQty * -1
            End If
            '
            ' SP02 / OP02 / FS_02 / PS_02 / IS_02
            sql &= " " & CStr(CDbl(xMonScheProd(3)) / 10000000) & ", "
            sql &= " " & CStr(CDbl(xMonOnProd(3)) / 10000000) & ", "

            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_02")
            xFCTQty = xFCTQty + dt_FCT.Rows(i).Item("FS_02")
            '
            sql &= " " & CStr(dt_FCT.Rows(i).Item("FS_02")) & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
                xFCTQty = xFCTQty - xSumQty * -1
            End If
            '
            ' SP03 / OP03 / FS_03 / PS_03 / IS_03
            sql &= " " & CStr(CDbl(xMonScheProd(4)) / 10000000) & ", "
            sql &= " " & CStr(CDbl(xMonOnProd(4)) / 10000000) & ", "

            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_03")
            xFCTQty = xFCTQty + dt_FCT.Rows(i).Item("FS_03")
            '
            sql &= " " & CStr(dt_FCT.Rows(i).Item("FS_03")) & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
                xFCTQty = xFCTQty - xSumQty * -1
            End If
            '
            ' SP04 / OP04 / FS_04 / PS_04 / IS_04
            sql &= " " & CStr(CDbl(xMonScheProd(5)) / 10000000) & ", "
            sql &= " " & CStr(CDbl(xMonOnProd(5)) / 10000000) & ", "

            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_04")
            xFCTQty = xFCTQty + dt_FCT.Rows(i).Item("FS_04")
            '
            sql &= " " & CStr(dt_FCT.Rows(i).Item("FS_04")) & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
                xFCTQty = xFCTQty - xSumQty * -1
            End If
            '
            ' SP05 / OP05 / FS_05 / PS_05 / IS_05
            sql &= " " & CStr(CDbl(xMonScheProd(6)) / 10000000) & ", "
            sql &= " " & CStr(CDbl(xMonOnProd(6)) / 10000000) & ", "

            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_05")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_05").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' FS_06 / PS_06 / IS_06
            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_06")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_06").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' FS_07 / PS_07 / IS_07
            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_07")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_07").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' FS_08 / PS_08 / IS_08
            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_08")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_08").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' FS_09 / PS_09 / IS_09
            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_09")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_09").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' FS_10 / PS_10 / IS_10
            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_10")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_10").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' FS_11 / PS_11 / IS_11
            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_11")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_11").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' FS_12 / PS_12 / IS_12
            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_12")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_12").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' FS_13 / PS_13
            sql &= " " & dt_FCT.Rows(i).Item("FS_13").ToString & ", "                   ' FS_13
            sql &= " " & CStr(xProdQty) & ", "                                          ' PS_13

            ' FREEALLOC
            If xFCTQty <= xKeepQty Then
                sql &= " " & "0" & ", "
            Else
                If xFCTQty - xKeepQty >= xFreeQty Then
                    sql &= " " & CStr(xFreeQty) & ", "
                Else
                    sql &= " " & CStr(xFCTQty - xKeepQty) & ", "
                End If
            End If
            '
            ' Description
            '  ** Free Qty
            If xFreeQty > 0 Then
                oWaves.GetFreeByLocation("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xDescr)
            End If
            '  ** Free PROD Qty
            oWaves.GetFreeProductionQty("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xQty)
            If CDbl(xQty) > 0 Then
                If xDescr <> "" Then
                    xDescr = xDescr + " * Free on Pro:[" + CStr(CDbl(xQty) / 10000000) + "]"
                Else
                    xDescr = "Free on Pro:[" + CStr(CDbl(xQty) / 10000000) + "]"
                End If
            End If
            'VENDIR FC-17112
            If dt_FCT.Rows(i).Item("GR_09").ToString <> "" And dt_FCT.Rows(i).Item("GR_10").ToString <> "" Then
                If xDescr <> "" Then
                    xDescr = "Length & Unit:[" + dt_FCT.Rows(i).Item("GR_09").ToString + "/" + dt_FCT.Rows(i).Item("GR_10").ToString + "]" + xDescr
                Else
                    xDescr = "Length & Unit:[" + dt_FCT.Rows(i).Item("GR_09").ToString + "/" + dt_FCT.Rows(i).Item("GR_10").ToString + "]"
                End If
            End If
            '  ** 
            If xDescr <> "" Then
                sql &= " '" & xDescr & "', "
            Else
                sql &= " '" & "" & "', "
            End If
            '-----------------------------
            '
            sql &= " '" & "LSPlan" & "', "
            sql &= " '" & NowDateTime & "' "
            sql &= " ) "
            oDataBase.ExecuteNonQuery(sql)
            '
            ' ---------------------------------------------------------------------------------
            ' Update FCT Plan Data (LSNo, LSSubNo)
            sql = "Update ForcastPlan Set "
            sql &= "LSNO = '" & xLSNo + xSeqNoString & "', "
            sql &= "LSSUBNO = " & CStr(xSubNo) & " "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And Version = " & xLastVersion & " "
            sql &= "  And Y_Level > 0 "
            For j = 0 To UBound(xFCTGroupField)
                sql &= " And " & xFCTGroupField(j) & " = '" & dt_FCT.Rows(i).Item(xLSGroupField(j)).ToString & "' "
            Next
            oDataBase.ExecuteNonQuery(sql)
        Next
        ' ***********************************************************************************
        ' 更新下一個可使用No
        ' ***********************************************************************************
        ' 製作SeqNo(不足4位補0)
        xSeqNo = xSeqNo + 1
        xSeqNoString = CStr(xSeqNo)
        If Len(xSeqNoString) < 2 Then
            xSeqNoString = "000" + CStr(xSeqNo)
        Else
            If Len(xSeqNoString) < 3 Then
                xSeqNoString = "00" + CStr(xSeqNo)
            Else
                If Len(xSeqNoString) < 4 Then
                    xSeqNoString = "0" + CStr(xSeqNo)
                End If
            End If
        End If
        ' 更新下一次可使用PONO
        UpdateNextLSNo(pBuyer, xLSNo + xSeqNoString)
        '
        ' ***********************************************************************************
        ' 更新VENDOR FC (VENDOR FC 限用)
        ' ***********************************************************************************
        If InStr(pBuyer, "F-VENDOR") > 0 Then
            oWavesEDI.UpdateVendorFC(pBuyer, pUserID)
        End If
        'Catch ex As Exception
        'RtnCode = 9
        'oDB.AccessLog(pLogID, pBuyer, "LSPLAN", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "LocalStockPlan", pUserID, "")
        'End Try
        '
        Return RtnCode
    End Function
    'NewLocalStockPlan-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([125]-LocalStockPlanProdInf)
    '**     Local Stock Plan-Prod Inf. 
    '***********************************************************************************************
    'LocalStockPlanProdInf-Start
    <WebMethod()> _
    Public Function LocalStockPlanProdInf(ByVal pLogID As String, _
                                          ByVal pBuyer As String, _
                                          ByVal pUserID As String, _
                                          ByVal pGRBuyer As String, _
                                          ByVal pFunList As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            ' ***********************************************************************************
            ' 變數定義及設定初值
            ' ***********************************************************************************
            Dim i, j, idx, xCount As Integer
            Dim Sql As String
            Dim xProdInfList(50) As String      ' Prod Inf List 
            Dim xProdInf(7) As String           ' Prod-Inf 
            Dim xPurchaseItem As Boolean        ' 採購品
            Dim NowMonth As Integer = Year(Now) * 100 + Month(Now) * 1

            Dim SqlString As String = "Insert into LOCALSTOCKPlan_ProdInf " + _
                                      "(BUYER, LSNO, LSSUBNO, VERSION, PRODMONTH, PRODTYPE, PRNO, CANUSINGDATE, QTY, ORDERNO, REFERENCE, REQDATE, CreateUser, CreateTime) "
            '
            ' ***********************************************************************************
            ' Local Stock Plan-Prod Inf.展開
            ' ***********************************************************************************
            Sql = "SELECT * FROM LocalStockPlan "
            Sql &= " Where Buyer = '" & pBuyer & "' "
            Sql &= "   And Version = " & "99" & " "
            Sql &= "   And (N_ScheProd > 0 Or N_OnProd > 0) "

            'Sql &= "   And GR_03 = '1458867' "

            Sql &= " Order by LSNo, LSSubNo "
            Dim dt_LSPlan As DataTable = oDataBase.GetDataTable(Sql)
            For i = 0 To dt_LSPlan.Rows.Count - 1
                ' 判斷是否採購品(0 = 不是 / 1 = 是)
                xPurchaseItem = False
                If oWaves.GetPurchaseItem(dt_LSPlan.Rows(i).Item("GR_03").ToString) = 1 Then
                    xPurchaseItem = True
                End If
                If xPurchaseItem Then
                    ' Get Purchase Inf.   DEPO, ITEM, COLOR, KEEPCODE, (R) - PRODINF, (R) - COUNT
                    oWaves.GetPurchaseInf("01", _
                                          dt_LSPlan.Rows(i).Item("GR_03").ToString, _
                                          dt_LSPlan.Rows(i).Item("GR_07").ToString, _
                                          dt_LSPlan.Rows(i).Item("GR_02").ToString, _
                                          xProdInfList, _
                                          xCount)
                Else
                    ' Get Production Inf.   DEPO, ITEM, COLOR, KEEPCODE, (R) - PRODINF, (R) - COUNT
                    oWaves.GetProductionInf("01", _
                                            dt_LSPlan.Rows(i).Item("GR_03").ToString, _
                                            dt_LSPlan.Rows(i).Item("GR_07").ToString, _
                                            dt_LSPlan.Rows(i).Item("GR_02").ToString, _
                                            xProdInfList, _
                                            xCount)
                End If

                For j = 0 To xCount - 1
                    Sql = SqlString
                    Sql &= "VALUES( "
                    Sql &= " '" & dt_LSPlan.Rows(i).Item("Buyer").ToString & "', "                  ' Buyer
                    Sql &= " '" & dt_LSPlan.Rows(i).Item("LSNO").ToString & "', "                   ' LSNO
                    Sql &= " " & dt_LSPlan.Rows(i).Item("LSSUBNO").ToString & ", "                  ' LSSUBNO
                    Sql &= " " & dt_LSPlan.Rows(i).Item("VERSION").ToString & ", "                  ' VERSION
                    '
                    xProdInf = xProdInfList(j).Split("^")
                    '
                    ' PRODMONTH
                    If Mid(CStr(NowMonth), 1, 4) = Mid(xProdInf(2), 1, 4) Then
                        idx = CInt(Mid(xProdInf(2), 1, 6)) - NowMonth
                    Else
                        If Mid(CStr(NowMonth), 1, 4) < Mid(xProdInf(2), 1, 4) Then
                            idx = (CInt(Mid(xProdInf(2), 1, 4)) - 1) * 100 + CInt(Mid(xProdInf(2), 5, 2)) + 12 - NowMonth
                        Else
                            idx = CInt(xProdInf(2)) - (CInt(Mid(CStr(NowMonth), 1, 4)) - 1) * 100 + CInt(Mid(CStr(NowMonth), 5, 2)) + 12
                        End If
                    End If
                    Select Case idx
                        Case Is <= 0
                            Sql &= " '" & "N月前(含N)" & "', "
                        Case 1
                            Sql &= " '" & "N+1月" & "', "
                        Case 2
                            Sql &= " '" & "N+2月" & "', "
                        Case 3
                            Sql &= " '" & "N+3月" & "', "
                        Case 4
                            Sql &= " '" & "N+4月" & "', "
                        Case Else
                            Sql &= " '" & "N+5月後(含N+5)" & "', "
                    End Select
                    '
                    ' SCHE/ON PROD, PRNO, CANUSINGDATE, QTY, ORDERNO, REFERENCE, REQDATE
                    If xProdInf(0) = "SCHE" Then
                        Sql &= " '" & "SCHE-PROD" & "', "   ' PRODTYPE
                    Else
                        If xPurchaseItem Then
                            Sql &= " '" & "ON-PURC" & "', "
                        Else
                            Sql &= " '" & "ON-PROD" & "', "
                        End If
                    End If
                    Sql &= " '" & xProdInf(1) & "', "   ' PRNO / PONO
                    Sql &= " '" & xProdInf(2) & "', "   ' CANUSINGDATE
                    Sql &= " " & CStr(CDbl(xProdInf(3)) / 10000000) & ", "      ' QTY
                    Sql &= " '" & xProdInf(4) & "', "   ' ORDERNO
                    Sql &= " '" & xProdInf(5) & "', "   ' REFERENCE
                    Sql &= " '" & xProdInf(6) & "', "   ' REQDATE

                    Sql &= " '" & "LS-PRODINF" & "', "
                    Sql &= " '" & NowDateTime & "' "
                    Sql &= " ) "
                    oDataBase.ExecuteNonQuery(Sql)
                Next
            Next
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "LS-PRODINF", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "LS-PRODINF", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'LocalStockPlanProdInf-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([130]-BuyerLocalStockPlan)
    '**     Buyer Local Stock Plan展開 
    '***********************************************************************************************
    'BuyerLocalStockPlan-Start
    <WebMethod()> _
    Public Function BuyerLocalStockPlan(ByVal pLogID As String, _
                                        ByVal pBuyer As String, _
                                        ByVal pUserID As String, _
                                        ByVal pGRBuyer As String, _
                                        ByVal pFunList As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        '
        Try
            ' ***********************************************************************************
            ' 變數定義及設定初值
            ' ***********************************************************************************
            Dim i, j As Integer
            Dim xLastVersion As Integer = 99    ' 最終版PLAN
            Dim xSystemItem As String = ""
            Dim xGroupItem As String = ""
            Dim xSelectField As String = ""
            Dim xFCTString, xLSString, xDescr As String
            Dim xFCTGroupField(), xLSGroupField() As String
            Dim xKeepCode As String = ""
            Dim xQty As String = "0"
            Dim xSumQty As Double = 0           ' N月生產+在庫合計
            Dim xKeepQty As Double = 0          ' SCHE-PROD + ON-PROD + KEEP-INV
            Dim xFreeQty As Double = 0          ' FREE-INV
            Dim xFCTQty As Double = 0           ' ORDER ZONE FCT QTY
            Dim xMonthProdQty As Double = 0     ' ORDER ZONE 生產數量
            Dim xProdQty As Double = 0          ' 需生產合計
            Dim xProdQtyUnit As Integer = 1     ' 生產數-數量單位
            Dim xMonScheProd(6), xMonOnProd(6) As String        ' N1 ~ N4 SCHE_PRODQTY, ON_PRODQTY
            Dim xPurchaseItem As Boolean        ' 採購品

            oWaves.Timeout = Timeout.Infinite

            ' ---------------------------------------------------------------------------------
            ' 取得 SystemItem 初值
            sql = "SELECT * FROM M_ImportRule "
            sql &= "Where Buyer = '" & pBuyer & "-5' "
            sql &= "  And Sts   = '1' "
            sql &= "  And DataRule Like '" & "SY_" & "%' "
            sql &= "Order by DataRule "
            Dim dt_List As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_List.Rows.Count - 1
                ' SystemItem
                If xSystemItem = "" Then
                    xSystemItem = dt_List.Rows(i).Item("Field").ToString
                Else
                    xSystemItem = xSystemItem & "," & dt_List.Rows(i).Item("Field").ToString
                End If
                ' SelectField
                If xSelectField = "" Then
                    xSelectField = dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("Field").ToString
                Else
                    xSelectField = xSelectField & ", " & dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("Field").ToString
                End If
            Next
            ' 取得FCT & LS Group Field
            xFCTString = xSystemItem
            xLSString = xSystemItem
            ' ---------------------------------------------------------------------------------
            ' 取得 GroupItem 初值
            dt_List.Clear()
            sql = "SELECT * FROM M_ImportRule "
            sql &= "Where Buyer = '" & pBuyer + "-5" & "' "
            sql &= "  And Sts   = '1' "
            sql &= "  And DataRule Like '" & "GR_" & "%' "
            sql &= "Order by DataRule "
            dt_List = oDataBase.GetDataTable(sql)
            For i = 0 To dt_List.Rows.Count - 1
                ' GroupItem
                If xGroupItem = "" Then
                    xGroupItem = dt_List.Rows(i).Item("Field").ToString
                Else
                    xGroupItem = xGroupItem & ", " & dt_List.Rows(i).Item("Field").ToString
                End If
                ' SelectField
                If xSelectField = "" Then
                    xSelectField = dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("DataRule").ToString
                Else
                    xSelectField = xSelectField & ", " & dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("DataRule").ToString
                End If
                ' 取得FCT & LS Group Field
                xFCTString = xFCTString & "," & dt_List.Rows(i).Item("Field").ToString
                xLSString = xLSString & "," & dt_List.Rows(i).Item("DataRule").ToString
            Next
            '
            ' 取得FCT & LS Group Field
            xFCTGroupField = xFCTString.Split(",")
            xLSGroupField = xLSString.Split(",")
            '
            ' GroupItem 不足10個時 ==> SelectField 需補至 10 個
            For j = i + 1 To 10
                If j < 10 Then
                    xSelectField = xSelectField & ", " & "'' AS GR_0" & CStr(j)
                Else
                    xSelectField = xSelectField & ", " & "'' AS GR_10"
                End If
            Next
            ' ---------------------------------------------------------------------------------
            ' 取得 SummaryItem 初值
            dt_List.Clear()
            sql = "SELECT * FROM M_ImportRule "
            sql &= "Where Buyer = '" & pBuyer + "-5" & "' "
            sql &= "  And Sts   = '1' "
            sql &= "  And ( DataRule Like '" & "FS_" & "%' Or DataRule Like '" & "PS_" & "%' ) "
            sql &= "Order by DataRule "
            dt_List = oDataBase.GetDataTable(sql)
            For i = 0 To dt_List.Rows.Count - 1
                ' SelectField
                If xSelectField = "" Then
                    xSelectField = "Sum(" & dt_List.Rows(i).Item("Field").ToString & ") AS " & dt_List.Rows(i).Item("DataRule").ToString
                Else
                    xSelectField = xSelectField & ", " & "Sum(" & dt_List.Rows(i).Item("Field").ToString & ") AS " & dt_List.Rows(i).Item("DataRule").ToString
                End If
            Next
            ' ***********************************************************************************
            ' 取得 BULSNO = BuyerGroupCode + "L" + 年月 + 流水號(4碼)   AD + L + 135 + 0001
            ' 變數定義及設定初值
            ' ***********************************************************************************
            Dim xBULSNo, xSeqNoString, xBUGroup, xItem As String
            Dim xSeqNo, xSubNo As Integer
            xBUGroup = "XX"
            xItem = ""
            ' 取得BuyerGroupCode
            sql = "SELECT BUGroupCode FROM M_FControlRecord "
            sql &= "Where Buyer = '" & pBuyer & "' "
            Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
            If dt_ControlRecord.Rows.Count > 0 Then xBUGroup = dt_ControlRecord.Rows(0).Item("BUGroupCode")
            ' 取得年月 
            Select Case Month(Now)
                Case 10
                    xBULSNo = xBUGroup + "L" + CStr(Year(Now) - 2000) + "A"
                Case 11
                    xBULSNo = xBUGroup + "L" + CStr(Year(Now) - 2000) + "B"
                Case 12
                    xBULSNo = xBUGroup + "L" + CStr(Year(Now) - 2000) + "C"
                Case Else
                    xBULSNo = xBUGroup + "L" + CStr(Year(Now) - 2000) + CStr(Month(Now))
            End Select
            ' 流水號(下一可使用No)及SubNo
            xSeqNo = GetBuyerLSSeqNo(pBuyer, xBULSNo)
            xSubNo = 1
            ' ***********************************************************************************
            ' Buyer Local Stock Plan展開
            ' ***********************************************************************************
            sql = "SELECT "
            sql &= xSelectField
            sql &= " FROM LocalStockPlan "
            sql &= " Where BuyerGroup = '" & pGRBuyer & "' "
            sql &= "   And Version = " & xLastVersion & " "
            sql &= " Group by " & xSystemItem & ", " & xGroupItem
            sql &= " Order by " & xSystemItem & ", " & xGroupItem
            Dim dt_LSPlan As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_LSPlan.Rows.Count - 1
                '
                sql = "Insert into BuyerLocalStockPlan "
                sql &= "( "
                sql &= "BUYER, "
                sql &= "BUYERGROUP, "
                sql &= "BULSNO, "
                sql &= "BULSSUBNO, "
                sql &= "VERSION, "

                sql &= "GR_01, "
                sql &= "GR_02, "
                sql &= "GR_03, "
                sql &= "GR_04, "
                sql &= "GR_05, "
                sql &= "GR_06, "
                sql &= "GR_07, "
                sql &= "GR_08, "
                sql &= "GR_09, "
                sql &= "GR_10, "

                sql &= "MINIMUMSTOCK, "
                sql &= "N_SCHEPROD, "
                sql &= "N_ONPROD, "
                sql &= "N_FREEINV, "
                sql &= "N_KEEPINV, "
                sql &= "N_TOTAL, "

                sql &= "SP_00, "        'ADD-START 13/11/11
                sql &= "OP_00, "        'ADD-START 13/11/11
                sql &= "FS_00, "        'ADD-START 13/12/4
                sql &= "PS_00, "        'ADD-START 13/12/4
                sql &= "IS_00, "        'ADD-START 13/12/4
                sql &= "SP_01, "        'ADD-START 13/11/11
                sql &= "OP_01, "        'ADD-START 13/11/11
                sql &= "FS_01, "
                sql &= "PS_01, "
                sql &= "IS_01, "
                sql &= "SP_02, "        'ADD-START 13/11/11
                sql &= "OP_02, "        'ADD-START 13/11/11
                sql &= "FS_02, "
                sql &= "PS_02, "
                sql &= "IS_02, "
                sql &= "SP_03, "        'ADD-START 13/11/11
                sql &= "OP_03, "        'ADD-START 13/11/11
                sql &= "FS_03, "
                sql &= "PS_03, "
                sql &= "IS_03, "
                sql &= "SP_04, "        'ADD-START 13/11/11
                sql &= "OP_04, "        'ADD-START 13/11/11
                sql &= "FS_04, "
                sql &= "PS_04, "
                sql &= "IS_04, "
                sql &= "SP_05, "        'ADD-START 13/11/11
                sql &= "OP_05, "        'ADD-START 13/11/11
                sql &= "FS_05, "
                sql &= "PS_05, "
                sql &= "IS_05, "
                sql &= "FS_06, "
                sql &= "PS_06, "
                sql &= "IS_06, "
                sql &= "FS_07, "
                sql &= "PS_07, "
                sql &= "IS_07, "
                sql &= "FS_08, "
                sql &= "PS_08, "
                sql &= "IS_08, "
                sql &= "FS_09, "
                sql &= "PS_09, "
                sql &= "IS_09, "
                sql &= "FS_10, "
                sql &= "PS_10, "
                sql &= "IS_10, "
                sql &= "FS_11, "
                sql &= "PS_11, "
                sql &= "IS_11, "
                sql &= "FS_12, "
                sql &= "PS_12, "
                sql &= "IS_12, "

                sql &= "FS_13, "
                sql &= "PS_13, "
                sql &= "FREEALLOC, "
                sql &= "DESCRIPTION, "

                sql &= "CreateUser, "
                sql &= "CreateTime "
                sql &= " ) "
                sql &= "VALUES( "
                sql &= " '" & "BULS" & "', "                                                    ' Buyer
                sql &= " '" & dt_LSPlan.Rows(i).Item("BuyerGroup").ToString & "', "             ' BuyerGroup
                ' ---------------------------------------------------------------------------------
                ' 設定 BULSNo, SubNo
                ' Item Code是否相同 Counter LSNo, SubNo
                If dt_LSPlan.Rows(i).Item("GR_03").ToString <> xItem Then
                    If i > 0 Then
                        xSeqNo = xSeqNo + 1
                        xSubNo = 1
                    End If
                    xItem = dt_LSPlan.Rows(i).Item("GR_03").ToString
                Else
                    If i > 0 Then xSubNo = xSubNo + 1
                End If
                ' 製作SeqNo(不足4位補0)
                xSeqNoString = CStr(xSeqNo)
                If Len(xSeqNoString) < 2 Then
                    xSeqNoString = "000" + CStr(xSeqNo)
                Else
                    If Len(xSeqNoString) < 3 Then
                        xSeqNoString = "00" + CStr(xSeqNo)
                    Else
                        If Len(xSeqNoString) < 4 Then
                            xSeqNoString = "0" + CStr(xSeqNo)
                        End If
                    End If
                End If
                '設定 BULSNo, SubNo
                sql &= " '" & xBULSNo + xSeqNoString & "', "                                    ' LSNO
                sql &= " " & CStr(xSubNo) & ", "                                                ' LSSUBNO
                ' ---------------------------------------------------------------------------------
                sql &= " " & CStr(xLastVersion) & ", "                                          ' Version

                sql &= " '" & dt_LSPlan.Rows(i).Item("GR_01").ToString & "', "                 ' GR_01
                sql &= " '" & dt_LSPlan.Rows(i).Item("GR_02").ToString & "', "                 ' GR_02
                sql &= " '" & dt_LSPlan.Rows(i).Item("GR_03").ToString & "', "                 ' GR_03
                sql &= " '" & dt_LSPlan.Rows(i).Item("GR_04").ToString & "', "                 ' GR_04
                sql &= " '" & dt_LSPlan.Rows(i).Item("GR_05").ToString & "', "                 ' GR_05
                sql &= " '" & dt_LSPlan.Rows(i).Item("GR_06").ToString & "', "                 ' GR_06
                sql &= " '" & dt_LSPlan.Rows(i).Item("GR_07").ToString & "', "                 ' GR_07
                sql &= " '" & dt_LSPlan.Rows(i).Item("GR_08").ToString & "', "                 ' GR_08
                sql &= " '" & dt_LSPlan.Rows(i).Item("GR_09").ToString & "', "                 ' GR_09
                sql &= " '" & dt_LSPlan.Rows(i).Item("GR_10").ToString & "', "                 ' GR_10

                If dt_LSPlan.Rows(i).Item("GR_08").ToString <> "ZIP" Then
                    '** PS / CH
                    ' ---------------------------------------------------------------------------------
                    ' 取得 MinimumStock Qty
                    oWaves.GetMininumStock("01", dt_LSPlan.Rows(i).Item("GR_03").ToString, dt_LSPlan.Rows(i).Item("GR_07").ToString, xQty)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                    Else
                        sql &= " " & "0" & ", "
                    End If
                    ' ---------------------------------------------------------------------------------
                    ' 歸零作業
                    xSumQty = 0
                    xMonthProdQty = 0       ' MONTH PROD-QTY
                    xProdQty = 0            ' PROD-QTY
                    xKeepQty = 0            ' SCHE-PROD + ON-PROD + KEEP-INV
                    xFreeQty = 0            ' FREE-INV
                    xFCTQty = 0             ' ORDER ZONE FCT QTY
                    xDescr = ""             ' Free Qty備註說明
                    For j = 1 To 6          ' N1~N4 PROD-QTY
                        xMonScheProd(j) = 0
                        xMonOnProd(j) = 0
                    Next

                    ' 取得PROD-QTY Unit
                    GetProdQtyUnit(pBuyer, dt_LSPlan.Rows(i).Item("GR_08").ToString, xProdQtyUnit)

                    ' 判斷是否採購品(0 = 不是 / 1 = 是)
                    xPurchaseItem = False
                    If oWaves.GetPurchaseItem(dt_LSPlan.Rows(i).Item("GR_03").ToString) = 1 Then
                        xPurchaseItem = True
                    End If
                    ' ---------------------------------------------------------------------------------
                    If xPurchaseItem Then
                        ' 採購品
                        ' 不使用 (N_SCHEPROD)
                        sql &= " " & "0" & ", "
                        '
                        ' 採購-ON Qty (N_ONPROD)
                        xKeepCode = dt_LSPlan.Rows(i).Item("GR_02").ToString

                        oWaves.GetPurchaseQty("01", dt_LSPlan.Rows(i).Item("GR_03").ToString, dt_LSPlan.Rows(i).Item("GR_07").ToString, xKeepCode, xQty, xMonOnProd)
                        If CDbl(xQty) > 0 Then
                            sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                            xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                            xSumQty = xSumQty + CDbl(xQty) / 10000000
                        Else
                            sql &= " " & "0" & ", "
                        End If
                    Else
                        ' 製造品
                        ' 生產-SCHE Qty (N_SCHEPROD)
                        xKeepCode = dt_LSPlan.Rows(i).Item("GR_02").ToString

                        oWaves.GetProductionQty("01", dt_LSPlan.Rows(i).Item("GR_03").ToString, dt_LSPlan.Rows(i).Item("GR_07").ToString, xKeepCode, 0, xQty, xMonScheProd)
                        If CDbl(xQty) > 0 Then
                            sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                            xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                            xSumQty = xSumQty + CDbl(xQty) / 10000000
                        Else
                            sql &= " " & "0" & ", "
                        End If
                        '
                        ' 生產-ON Qty (N_ONPROD)
                        oWaves.GetProductionQty("01", dt_LSPlan.Rows(i).Item("GR_03").ToString, dt_LSPlan.Rows(i).Item("GR_07").ToString, xKeepCode, 1, xQty, xMonOnProd)
                        If CDbl(xQty) > 0 Then
                            sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                            xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                            xSumQty = xSumQty + CDbl(xQty) / 10000000
                        Else
                            sql &= " " & "0" & ", "
                        End If
                    End If
                    '
                    ' 在庫-Free Qty (N_FREEINV)
                    oWaves.GetFreeInventory("01", dt_LSPlan.Rows(i).Item("GR_03").ToString, dt_LSPlan.Rows(i).Item("GR_07").ToString, xQty)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xFreeQty = xFreeQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                    '
                    ' 在庫-Keep Qty (N_KEEPINV)
                    oWaves.GetKeepCodeInventory("01", dt_LSPlan.Rows(i).Item("GR_03").ToString, dt_LSPlan.Rows(i).Item("GR_07").ToString, xKeepCode, xQty)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                    '
                    ' 合計=(生產-SCHE Qty + 生產-ON Qty + 在庫-Free Qty +在庫-Keep Qty) (N_TOTAL)
                    sql &= " " & CStr(xSumQty) & ", "
                Else
                    '** ZIP
                    ' ---------------------------------------------------------------------------------
                    ' 取得 MinimumStock Qty
                    oWaves.GetMininumStockZIP("01", dt_LSPlan.Rows(i).Item("GR_03").ToString, dt_LSPlan.Rows(i).Item("GR_07").ToString, dt_LSPlan.Rows(i).Item("GR_09").ToString, dt_LSPlan.Rows(i).Item("GR_10").ToString, xQty)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                    Else
                        sql &= " " & "0" & ", "
                    End If
                    ' ---------------------------------------------------------------------------------
                    ' 歸零作業
                    xSumQty = 0
                    xMonthProdQty = 0       ' MONTH PROD-QTY
                    xProdQty = 0            ' PROD-QTY
                    xKeepQty = 0            ' SCHE-PROD + ON-PROD + KEEP-INV
                    xFreeQty = 0            ' FREE-INV
                    xFCTQty = 0             ' ORDER ZONE FCT QTY
                    xDescr = ""             ' Free Qty備註說明
                    For j = 1 To 6          ' N1~N4 PROD-QTY
                        xMonScheProd(j) = 0
                        xMonOnProd(j) = 0
                    Next

                    ' 取得PROD-QTY Unit
                    GetProdQtyUnit(pBuyer, dt_LSPlan.Rows(i).Item("GR_08").ToString, xProdQtyUnit)

                    ' 判斷是否採購品(0 = 不是 / 1 = 是)
                    xPurchaseItem = False
                    If oWaves.GetPurchaseItem(dt_LSPlan.Rows(i).Item("GR_03").ToString) = 1 Then
                        xPurchaseItem = True
                    End If
                    ' ---------------------------------------------------------------------------------
                    If xPurchaseItem Then
                        ' 採購品
                        ' 不使用 (N_SCHEPROD)
                        sql &= " " & "0" & ", "
                        '
                        ' 採購-ON Qty (N_ONPROD)
                        xKeepCode = dt_LSPlan.Rows(i).Item("GR_02").ToString

                        oWaves.GetPurchaseQtyZIP("01", dt_LSPlan.Rows(i).Item("GR_03").ToString, dt_LSPlan.Rows(i).Item("GR_07").ToString, dt_LSPlan.Rows(i).Item("GR_09").ToString, dt_LSPlan.Rows(i).Item("GR_10").ToString, xKeepCode, xQty, xMonOnProd)
                        If CDbl(xQty) > 0 Then
                            sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                            xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                            xSumQty = xSumQty + CDbl(xQty) / 10000000
                        Else
                            sql &= " " & "0" & ", "
                        End If
                    Else
                        ' 製造品
                        ' 生產-SCHE Qty (N_SCHEPROD)
                        xKeepCode = dt_LSPlan.Rows(i).Item("GR_02").ToString

                        oWaves.GetProductionQtyZIP("01", dt_LSPlan.Rows(i).Item("GR_03").ToString, dt_LSPlan.Rows(i).Item("GR_07").ToString, xKeepCode, dt_LSPlan.Rows(i).Item("GR_09").ToString, dt_LSPlan.Rows(i).Item("GR_10").ToString, 0, xQty, xMonScheProd)
                        If CDbl(xQty) > 0 Then
                            sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                            xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                            xSumQty = xSumQty + CDbl(xQty) / 10000000
                        Else
                            sql &= " " & "0" & ", "
                        End If
                        '
                        ' 生產-ON Qty (N_ONPROD)
                        oWaves.GetProductionQtyZIP("01", dt_LSPlan.Rows(i).Item("GR_03").ToString, dt_LSPlan.Rows(i).Item("GR_07").ToString, dt_LSPlan.Rows(i).Item("GR_09").ToString, dt_LSPlan.Rows(i).Item("GR_10").ToString, xKeepCode, 1, xQty, xMonOnProd)
                        If CDbl(xQty) > 0 Then
                            sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                            xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                            xSumQty = xSumQty + CDbl(xQty) / 10000000
                        Else
                            sql &= " " & "0" & ", "
                        End If
                    End If
                    '
                    ' 在庫-Free Qty (N_FREEINV)
                    oWaves.GetFreeInventoryZIP("01", dt_LSPlan.Rows(i).Item("GR_03").ToString, dt_LSPlan.Rows(i).Item("GR_07").ToString, dt_LSPlan.Rows(i).Item("GR_09").ToString, dt_LSPlan.Rows(i).Item("GR_10").ToString, xQty)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xFreeQty = xFreeQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                    '
                    ' 在庫-Keep Qty (N_KEEPINV)
                    oWaves.GetKeepCodeInventoryZIP("01", dt_LSPlan.Rows(i).Item("GR_03").ToString, dt_LSPlan.Rows(i).Item("GR_07").ToString, dt_LSPlan.Rows(i).Item("GR_09").ToString, dt_LSPlan.Rows(i).Item("GR_10").ToString, xKeepCode, xQty)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                    '
                    ' 合計=(生產-SCHE Qty + 生產-ON Qty + 在庫-Free Qty +在庫-Keep Qty) (N_TOTAL)
                    sql &= " " & CStr(xSumQty) & ", "
                End If
                ' ---------------------------------------------------------------------------------
                ' FORCAST & PRODUCTION & INVENTORY
                ' SP00 / OP00 / FS_00 / PS_00 / IS_00
                sql &= " " & CStr(CDbl(xMonScheProd(1)) / 10000000) & ", "
                sql &= " " & CStr(CDbl(xMonOnProd(1)) / 10000000) & ", "

                sql &= " " & dt_LSPlan.Rows(i).Item("FS_00").ToString & ", "
                sql &= " " & dt_LSPlan.Rows(i).Item("PS_00").ToString & ", "
                xSumQty = xSumQty + dt_LSPlan.Rows(i).Item("PS_00") - dt_LSPlan.Rows(i).Item("FS_00")
                sql &= " " & CStr(xSumQty) & ", "

                xProdQty = xProdQty + dt_LSPlan.Rows(i).Item("PS_00")                   ' PROD-QTY合計
                '
                ' SP01 / OP01 / FS_01 / PS_01 / IS_01
                sql &= " " & CStr(CDbl(xMonScheProd(2)) / 10000000) & ", "
                sql &= " " & CStr(CDbl(xMonOnProd(2)) / 10000000) & ", "

                sql &= " " & dt_LSPlan.Rows(i).Item("FS_01").ToString & ", "            ' FS_01
                xMonthProdQty = dt_LSPlan.Rows(i).Item("PS_01")                         ' PS_01
                If xMonthProdQty Mod xProdQtyUnit > 0 Then
                    xMonthProdQty = (Fix(xMonthProdQty / xProdQtyUnit) + 1) * xProdQtyUnit
                End If
                sql &= " " & CStr(xMonthProdQty) & ", "
                xSumQty = xSumQty + xMonthProdQty - dt_LSPlan.Rows(i).Item("FS_01")     ' IS_01
                sql &= " " & CStr(xSumQty) & ", "

                xFCTQty = xFCTQty + dt_LSPlan.Rows(i).Item("FS_01") - xMonthProdQty     ' FREEALLOC計算
                xProdQty = xProdQty + xMonthProdQty                                     ' PROD-QTY合計
                ' ----------------------
                ' SP02 / OP02 / FS_02 / PS_02 / IS_02
                sql &= " " & CStr(CDbl(xMonScheProd(3)) / 10000000) & ", "
                sql &= " " & CStr(CDbl(xMonOnProd(3)) / 10000000) & ", "

                sql &= " " & dt_LSPlan.Rows(i).Item("FS_02").ToString & ", "            ' FS_02
                xMonthProdQty = dt_LSPlan.Rows(i).Item("PS_02")                         ' PS_02
                If xMonthProdQty Mod xProdQtyUnit > 0 Then
                    xMonthProdQty = (Fix(xMonthProdQty / xProdQtyUnit) + 1) * xProdQtyUnit
                End If
                sql &= " " & CStr(xMonthProdQty) & ", "
                xSumQty = xSumQty + xMonthProdQty - dt_LSPlan.Rows(i).Item("FS_02")     ' IS_02
                sql &= " " & CStr(xSumQty) & ", "

                xFCTQty = xFCTQty + dt_LSPlan.Rows(i).Item("FS_02") - xMonthProdQty     ' FREEALLOC計算
                xProdQty = xProdQty + xMonthProdQty                                     ' PROD-QTY合計
                ' ----------------------
                ' SP03 / OP03 / FS_03 / PS_03 / IS_03
                sql &= " " & CStr(CDbl(xMonScheProd(4)) / 10000000) & ", "
                sql &= " " & CStr(CDbl(xMonOnProd(4)) / 10000000) & ", "

                sql &= " " & dt_LSPlan.Rows(i).Item("FS_03").ToString & ", "            ' FS_03
                xMonthProdQty = dt_LSPlan.Rows(i).Item("PS_03")                         ' PS_03
                If xMonthProdQty Mod xProdQtyUnit > 0 Then
                    xMonthProdQty = (Fix(xMonthProdQty / xProdQtyUnit) + 1) * xProdQtyUnit
                End If
                sql &= " " & CStr(xMonthProdQty) & ", "
                xSumQty = xSumQty + xMonthProdQty - dt_LSPlan.Rows(i).Item("FS_03")     ' IS_03
                sql &= " " & CStr(xSumQty) & ", "

                xFCTQty = xFCTQty + dt_LSPlan.Rows(i).Item("FS_03") - xMonthProdQty     ' FREEALLOC計算
                xProdQty = xProdQty + xMonthProdQty                                     ' PROD-QTY合計
                ' ----------------------
                ' SP04 / OP04 / FS_04 / PS_04 / IS_04
                sql &= " " & CStr(CDbl(xMonScheProd(5)) / 10000000) & ", "
                sql &= " " & CStr(CDbl(xMonOnProd(5)) / 10000000) & ", "

                sql &= " " & dt_LSPlan.Rows(i).Item("FS_04").ToString & ", "            ' FS_04
                xMonthProdQty = dt_LSPlan.Rows(i).Item("PS_04")                         ' PS_04
                If xMonthProdQty Mod xProdQtyUnit > 0 Then
                    xMonthProdQty = (Fix(xMonthProdQty / xProdQtyUnit) + 1) * xProdQtyUnit
                End If
                sql &= " " & CStr(xMonthProdQty) & ", "
                xSumQty = xSumQty + xMonthProdQty - dt_LSPlan.Rows(i).Item("FS_04")     ' IS_04
                sql &= " " & CStr(xSumQty) & ", "

                xFCTQty = xFCTQty + dt_LSPlan.Rows(i).Item("FS_04") - xMonthProdQty     ' FREEALLOC計算
                xProdQty = xProdQty + xMonthProdQty                                     ' PROD-QTY合計
                ' ----------------------
                ' SP05 / OP05 / FS_05 / PS_05 / IS_05
                sql &= " " & CStr(CDbl(xMonScheProd(6)) / 10000000) & ", "
                sql &= " " & CStr(CDbl(xMonOnProd(6)) / 10000000) & ", "

                sql &= " " & dt_LSPlan.Rows(i).Item("FS_05").ToString & ", "
                sql &= " " & dt_LSPlan.Rows(i).Item("PS_05").ToString & ", "
                xSumQty = xSumQty + dt_LSPlan.Rows(i).Item("PS_05") - dt_LSPlan.Rows(i).Item("FS_05")
                sql &= " " & CStr(xSumQty) & ", "

                xProdQty = xProdQty + dt_LSPlan.Rows(i).Item("PS_05")                   ' PROD-QTY合計
                ' ----------------------
                ' FS_06 / PS_06 / IS_06
                sql &= " " & dt_LSPlan.Rows(i).Item("FS_06").ToString & ", "
                sql &= " " & dt_LSPlan.Rows(i).Item("PS_06").ToString & ", "
                xSumQty = xSumQty + dt_LSPlan.Rows(i).Item("PS_06") - dt_LSPlan.Rows(i).Item("FS_06")
                sql &= " " & CStr(xSumQty) & ", "

                xProdQty = xProdQty + dt_LSPlan.Rows(i).Item("PS_06")                   ' PROD-QTY合計
                ' ----------------------
                ' FS_07 / PS_07 / IS_07
                sql &= " " & dt_LSPlan.Rows(i).Item("FS_07").ToString & ", "
                sql &= " " & dt_LSPlan.Rows(i).Item("PS_07").ToString & ", "
                xSumQty = xSumQty + dt_LSPlan.Rows(i).Item("PS_07") - dt_LSPlan.Rows(i).Item("FS_07")
                sql &= " " & CStr(xSumQty) & ", "

                xProdQty = xProdQty + dt_LSPlan.Rows(i).Item("PS_07")                   ' PROD-QTY合計
                ' ----------------------
                ' FS_08 / PS_08 / IS_08
                sql &= " " & dt_LSPlan.Rows(i).Item("FS_08").ToString & ", "
                sql &= " " & dt_LSPlan.Rows(i).Item("PS_08").ToString & ", "
                xSumQty = xSumQty + dt_LSPlan.Rows(i).Item("PS_08") - dt_LSPlan.Rows(i).Item("FS_08")
                sql &= " " & CStr(xSumQty) & ", "

                xProdQty = xProdQty + dt_LSPlan.Rows(i).Item("PS_08")                   ' PROD-QTY合計
                ' ----------------------
                ' FS_09 / PS_09 / IS_09
                sql &= " " & dt_LSPlan.Rows(i).Item("FS_09").ToString & ", "
                sql &= " " & dt_LSPlan.Rows(i).Item("PS_09").ToString & ", "
                xSumQty = xSumQty + dt_LSPlan.Rows(i).Item("PS_09") - dt_LSPlan.Rows(i).Item("FS_09")
                sql &= " " & CStr(xSumQty) & ", "

                xProdQty = xProdQty + dt_LSPlan.Rows(i).Item("PS_09")                   ' PROD-QTY合計
                ' ----------------------
                ' FS_10 / PS_10 / IS_10
                sql &= " " & dt_LSPlan.Rows(i).Item("FS_10").ToString & ", "
                sql &= " " & dt_LSPlan.Rows(i).Item("PS_10").ToString & ", "
                xSumQty = xSumQty + dt_LSPlan.Rows(i).Item("PS_10") - dt_LSPlan.Rows(i).Item("FS_10")
                sql &= " " & CStr(xSumQty) & ", "

                xProdQty = xProdQty + dt_LSPlan.Rows(i).Item("PS_10")                   ' PROD-QTY合計
                ' ----------------------
                ' FS_11 / PS_11 / IS_11
                sql &= " " & dt_LSPlan.Rows(i).Item("FS_11").ToString & ", "
                sql &= " " & dt_LSPlan.Rows(i).Item("PS_11").ToString & ", "
                xSumQty = xSumQty + dt_LSPlan.Rows(i).Item("PS_11") - dt_LSPlan.Rows(i).Item("FS_11")
                sql &= " " & CStr(xSumQty) & ", "

                xProdQty = xProdQty + dt_LSPlan.Rows(i).Item("PS_11")                   ' PROD-QTY合計
                ' ----------------------
                ' FS_12 / PS_12 / IS_12
                sql &= " " & dt_LSPlan.Rows(i).Item("FS_12").ToString & ", "
                sql &= " " & dt_LSPlan.Rows(i).Item("PS_12").ToString & ", "
                xSumQty = xSumQty + dt_LSPlan.Rows(i).Item("PS_12") - dt_LSPlan.Rows(i).Item("FS_12")
                sql &= " " & CStr(xSumQty) & ", "

                xProdQty = xProdQty + dt_LSPlan.Rows(i).Item("PS_12")                   ' PROD-QTY合計
                ' ----------------------
                ' FS_13 / PS_13
                sql &= " " & dt_LSPlan.Rows(i).Item("FS_13").ToString & ", "            ' FS_13
                sql &= " " & CStr(xProdQty) & ", "                                      ' PS_13

                If xFCTQty <= xKeepQty Then                                             ' FREEALLOC
                    sql &= " " & "0" & ", "
                Else
                    If xFCTQty - xKeepQty >= xFreeQty Then
                        sql &= " " & CStr(xFreeQty) & ", "
                    Else
                        sql &= " " & CStr(xFCTQty - xKeepQty) & ", "
                    End If
                End If
                '
                ' Description
                '  ** Free Qty
                If xFreeQty > 0 Then
                    oWaves.GetFreeByLocation("01", dt_LSPlan.Rows(i).Item("GR_03").ToString, dt_LSPlan.Rows(i).Item("GR_07").ToString, xDescr)
                End If
                '  ** Free PROD Qty
                oWaves.GetFreeProductionQty("01", dt_LSPlan.Rows(i).Item("GR_03").ToString, dt_LSPlan.Rows(i).Item("GR_07").ToString, xQty)
                If CDbl(xQty) > 0 Then
                    If xDescr <> "" Then
                        xDescr = xDescr + " * Free on Pro:[" + CStr(CDbl(xQty) / 10000000) + "]"
                    Else
                        xDescr = "Free on Pro:[" + CStr(CDbl(xQty) / 10000000) + "]"
                    End If
                End If
                'VENDIR FC-17112
                If dt_LSPlan.Rows(i).Item("GR_09").ToString <> "" And dt_LSPlan.Rows(i).Item("GR_10").ToString <> "" Then
                    If xDescr <> "" Then
                        xDescr = "Length & Unit:[" + dt_LSPlan.Rows(i).Item("GR_09").ToString + "/" + dt_LSPlan.Rows(i).Item("GR_10").ToString + "]" + xDescr
                    Else
                        xDescr = "Length & Unit:[" + dt_LSPlan.Rows(i).Item("GR_09").ToString + "/" + dt_LSPlan.Rows(i).Item("GR_10").ToString + "]"
                    End If
                End If
                '  ** 
                If xDescr <> "" Then
                    sql &= " '" & xDescr & "', "
                Else
                    sql &= " '" & "" & "', "
                End If
                ' ---------------------------------------------------------------------------------
                sql &= " '" & "BULSPlan" & "', "
                sql &= " '" & NowDateTime & "' "
                sql &= " ) "
                oDataBase.ExecuteNonQuery(sql)
                '
                ' ---------------------------------------------------------------------------------
                ' Update LS Plan Data (BULSNo, BULSSubNo)
                sql = "Update LocalStockPlan Set "
                sql &= "BULSNO = '" & xBULSNo + xSeqNoString & "', "
                sql &= "BULSSUBNO = " & CStr(xSubNo) & " "
                sql &= " Where BuyerGroup = '" & pGRBuyer & "' "
                sql &= "   And Version = " & xLastVersion & " "
                For j = 0 To UBound(xFCTGroupField)
                    sql &= " And " & xFCTGroupField(j) & " = '" & dt_LSPlan.Rows(i).Item(xLSGroupField(j)).ToString & "' "
                Next
                oDataBase.ExecuteNonQuery(sql)
                '
            Next
            ' ***********************************************************************************
            ' 更新下一個可使用No
            ' ***********************************************************************************
            ' 製作SeqNo(不足4位補0)
            xSeqNo = xSeqNo + 1
            xSeqNoString = CStr(xSeqNo)
            If Len(xSeqNoString) < 2 Then
                xSeqNoString = "000" + CStr(xSeqNo)
            Else
                If Len(xSeqNoString) < 3 Then
                    xSeqNoString = "00" + CStr(xSeqNo)
                Else
                    If Len(xSeqNoString) < 4 Then
                        xSeqNoString = "0" + CStr(xSeqNo)
                    End If
                End If
            End If
            ' 更新下一次可使用PONO
            UpdateNextBuyerLSNo(pBuyer, xBULSNo + xSeqNoString)
            '
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "BULSPLAN", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "BuyerLocalStockPlan", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'BuyerLocalStockPlan-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([140]-LastFinalLocalStockPlan)
    '**     Buyer Local Stock Plan最終確定 
    '***********************************************************************************************
    'LastFinalLocalStockPlan-Start
    <WebMethod()> _
    Public Function LastFinalLocalStockPlan(ByVal pLogID As String, _
                                            ByVal pBuyer As String, _
                                            ByVal pUserID As String, _
                                            ByVal pGRBuyer As String, _
                                            ByVal pFunList As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        '
        Try
            ' -----------------------------------------------------------------------------------
            ' Forcast Plan ---> LF Forcast Plan
            ' -----------------------------------------------------------------------------------
            '
            ' Insert LastFinal ForcastPlan
            sql = "Insert into LF_ForcastPlan "
            sql &= "Select *, '" & pLogID & "', '' "
            sql &= "From ForcastPlan "
            sql &= "Where BuyerGroup = '" & pGRBuyer & "' "
            'sql &= "  And Version = " & "99" & " "
            sql &= "Order by FCTNo, FCTSubNo "
            oDataBase.ExecuteNonQuery(sql)
            '
            ' Delete ForcastPlan
            sql = "Delete From ForcastPlan "
            sql &= "Where BuyerGroup = '" & pGRBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
            ' -----------------------------------------------------------------------------------
            ' LocalStock Plan ---> LF LocalStock Plan
            ' -----------------------------------------------------------------------------------
            '
            ' Insert LastFinal LocalStockPlan
            sql = "Insert into LF_LocalStockPlan "
            sql &= "Select *, '" & pLogID & "', '' "
            sql &= "From LocalStockPlan "
            sql &= "Where BuyerGroup = '" & pGRBuyer & "' "
            'sql &= "  And Version = " & "99" & " "
            sql &= "Order by LSNo, LSSubNo "
            oDataBase.ExecuteNonQuery(sql)
            '
            ' Delete LocalStockPlan
            sql = "Delete From LocalStockPlan "
            sql &= "Where BuyerGroup = '" & pGRBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
            ' -----------------------------------------------------------------------------------
            ' BuyerLocalStock Plan ---> LF Buyer LocalStock Plan
            ' -----------------------------------------------------------------------------------
            '
            ' Insert LastFinal BuyerLocalStockPlan
            sql = "Insert into LF_BuyerLocalStockPlan "
            sql &= "Select *, '" & pLogID & "', '' "
            sql &= "From BuyerLocalStockPlan "
            sql &= "Where BuyerGroup = '" & pGRBuyer & "' "
            'sql &= "  And Version = " & "99" & " "
            sql &= "Order by BULSNo, BULSSubNo "
            oDataBase.ExecuteNonQuery(sql)
            '
            ' Delete BuyerLocalStockPlan
            sql = "Delete From BuyerLocalStockPlan "
            sql &= "Where BuyerGroup = '" & pGRBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
            ' -----------------------------------------------------------------------------------
            ' InputSheet ---> LF_InputSheet
            ' -----------------------------------------------------------------------------------
            '
            sql = "Insert into LF_InputSheet "
            sql &= "Select *, '" & pLogID & "', '' "
            sql &= "From E_InputSheet "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "Order by Unique_ID "
            oDataBase.ExecuteNonQuery(sql)
            ' -----------------------------------------------------------------------------------
            ' Update M_Referp Cat=200 DKey=AGENT-FCT2ANA
            ' 觸發Batch程式執行 (FCT分析使用)
            ' -----------------------------------------------------------------------------------
            '
            ' 更新 M_Referp
            sql = "Update M_Referp Set "
            sql = sql + " Data = '" & "Y" & "/" & pLogID & "' "
            sql = sql + " Where Cat = '200' "
            sql = sql + "   And DKey = 'AGENT-FCT2ANA' "
            oDataBase.ExecuteNonQuery(sql)
            '
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "LFLSPLAN", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "LastFinalLocalStockPlan", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'BuyerLocalStockPlan-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([150]-BYFCTDataCheck)
    '**     Buyer FCT Data Check
    '***********************************************************************************************
    'BYFCTDataCheck-Start
    <WebMethod()> _
    Public Function BYFCTDataCheck(ByVal pLogID As String, _
                                   ByVal pBuyer As String, _
                                   ByVal pUserID As String, _
                                   ByVal pGRBuyer As String, _
                                   ByVal pFunList As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i, j As Integer
        Dim sql, xBuyerName, str As String
        Dim StsCode As String = ""
        '
        Try
            ' -----------------------------------------------------------------------------------
            ' 設定 BuyerName
            ' -----------------------------------------------------------------------------------
            xBuyerName = ""
            If InStr(pBuyer, "000001") > 0 Then xBuyerName = "ADIDAS"
            If InStr(pBuyer, "000016") > 0 Then xBuyerName = "REEBOK"
            '
            ' -----------------------------------------------------------------------------------
            ' 10    BRAND判斷(A1)→非ADIDAS             
            ' -----------------------------------------------------------------------------------
            StsCode = "10"
            If oDB.GetBYFCTDataCheckFun(pBuyer, 1) = "1" Then
                sql = "Update E_InputSheetBY Set "
                sql &= " EZ1 = '10', "
                sql &= " ModifyUser = '" & pUserID & "', "
                sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EY1   = '" & "*" & "' "
                sql &= "   And A1   <> '" & xBuyerName & "' "
                oDataBase.ExecuteNonQuery(sql)
                ' BY FCT 資料同步化
                UpdateBYFCTStatus(pBuyer, pUserID, "10")
            End If
            '
            ' -----------------------------------------------------------------------------------
            ' 20    T2 SUPPLIER NAME判斷(B1)→非YKK        
            ' -----------------------------------------------------------------------------------
            StsCode = "20"
            If oDB.GetBYFCTDataCheckFun(pBuyer, 2) = "1" Then
                sql = "Update E_InputSheetBY Set "
                sql &= " EZ1 = '20', "
                sql &= " ModifyUser = '" & pUserID & "', "
                sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EY1   = '" & "*" & "' "
                sql &= "   And Not B1 LIKE '%" & "YKK" & "%' "
                oDataBase.ExecuteNonQuery(sql)
                ' BY FCT 資料同步化
                UpdateBYFCTStatus(pBuyer, pUserID, "20")
            End If
            '
            ' -----------------------------------------------------------------------------------
            ' FACTORY NAME判斷(L1)
            ' 40    FACTORY(NAME(L1) = M_NativeVendor(FactoryName)不存在                 →新成衣廠
            ' 41    FACTORY NAME(L1) = M_NativeVendor(FactoryName)存在但CustName=空白    →姐妹社
            ' -----------------------------------------------------------------------------------
            ' 40→新成衣廠
            StsCode = "40"
            If oDB.GetBYFCTDataCheckFun(pBuyer, 4) = "1" Then
                sql = "Update E_InputSheetBY Set "
                sql &= " EZ1 = '40', "
                sql &= " ModifyUser = '" & pUserID & "', "
                sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EY1   = '" & "*" & "' "
                sql &= "   And L1 Not In (Select Factoryname From M_NativeVendor Where M_NativeVendor.Buyer=E_InputSheetBY.Buyer And Active=1 Group By Factoryname) "
                oDataBase.ExecuteNonQuery(sql)
                ' BY FCT 資料同步化
                UpdateBYFCTStatus(pBuyer, pUserID, "40")
            End If
            ' 41→姐妹社
            StsCode = "41"
            If oDB.GetBYFCTDataCheckFun(pBuyer, 5) = "1" Then
                sql = "Update E_InputSheetBY Set "
                sql &= " EZ1 = '41', "
                sql &= " ModifyUser = '" & pUserID & "', "
                sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EY1   = '" & "*" & "' "
                sql &= "   And (Select CustName From M_NativeVendor Where M_NativeVendor.Buyer=E_InputSheetBY.Buyer And Active=1 And Factoryname=L1 Group By CustName) = '" & "" & "' "
                oDataBase.ExecuteNonQuery(sql)
                ' 增加-YKK GROUP管理單位
                sql = "Update E_InputSheetBY Set "
                sql &= " N1 = '[' + (Select YKKManageLocation From M_NativeVendor Where M_NativeVendor.Buyer=E_InputSheetBY.Buyer And Active=1 And Factoryname=L1 Group By YKKManageLocation) + ']-' + N1 "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EZ1 = '41' "
                sql &= "   And N1 NOT LIKE (Select '%'+YKKManageLocation+'%' From M_NativeVendor Where M_NativeVendor.Buyer=E_InputSheetBY.Buyer And Active=1 And Factoryname=L1 Group By YKKManageLocation) "
                oDataBase.ExecuteNonQuery(sql)
                ' BY FCT 資料同步化
                UpdateBYFCTStatus(pBuyer, pUserID, "41")
            End If
            '
            ' -----------------------------------------------------------------------------------
            ' 50    CCS判斷 WORKING_NO-(D1) = A 開始 / S結尾 / CCS PLM ITEM
            ' -----------------------------------------------------------------------------------
            StsCode = "50"
            If oDB.GetBYFCTDataCheckFun(pBuyer, 7) = "1" Then
                sql = "Update E_InputSheetBY Set "
                sql &= " EZ1 = '50', "
                sql &= " ModifyUser = '" & pUserID & "', "
                sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EY1   = '" & "*" & "' "
                sql &= "   And (   D1 LIKE '" & "A" & "%' "
                sql &= "        Or D1 LIKE '%" & "S" & "' "
                sql &= "        Or C1 In (Select PLM From M_ItemDescriptionConvert Where M_ItemDescriptionConvert.Buyer=E_InputSheetBY.Buyer And DataType='2' Group By PLM) "
                sql &= "       ) "
                oDataBase.ExecuteNonQuery(sql)
                ' BY FCT 資料同步化
                UpdateBYFCTStatus(pBuyer, pUserID, "50")
            End If
            '
            ' -----------------------------------------------------------------------------------
            ' 42/52   TWN特殊部品判斷(對象=41:姐妹社, 50:CCS)  42:姐妹社有TWN特殊部品 / 52:CCS有TWN特殊部品
            ' -----------------------------------------------------------------------------------
            ' 判斷TWN特殊部品 
            StsCode = "42/52-準備"
            sql = "SELECT Part, Condition FROM M_SpecialMaterial "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And Active = 1 "
            sql &= "Order by Part "
            Dim dt_SpecialMaterial As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_SpecialMaterial.Rows.Count - 1
                '
                str = Replace(dt_SpecialMaterial.Rows(i).Item("Condition").ToString, "@ItemDescription", "AG1+AH1+AI1")
                '
                sql = "Update E_InputSheetBY Set "
                sql &= " EX1 = '" & dt_SpecialMaterial.Rows(i).Item("Part").ToString & "' "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EY1   = '" & "*" & "' "
                sql &= "   And (EZ1 = '41' or EZ1 = '50') "
                sql &= "   And " & str & " "
                oDataBase.ExecuteNonQuery(sql)
            Next
            '  42:姐妹社有TWN特殊部品
            StsCode = "42"
            If oDB.GetBYFCTDataCheckFun(pBuyer, 6) = "1" Then
                sql = "Update E_InputSheetBY Set "
                sql &= " EZ1 = '42', "
                sql &= " ModifyUser = '" & pUserID & "', "
                sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EY1   = '" & "*" & "' "
                sql &= "   And EZ1   = '" & "41" & "' "
                sql &= "   And EX1  <> '" & "" & "' "
                oDataBase.ExecuteNonQuery(sql)
                ' BY FCT 資料同步化
                UpdateBYFCTStatus(pBuyer, pUserID, "42")
            End If
            '  52:CCS有TWN特殊部品
            StsCode = "52"
            If oDB.GetBYFCTDataCheckFun(pBuyer, 8) = "1" Then
                sql = "Update E_InputSheetBY Set "
                sql &= " EZ1 = '52', "
                sql &= " ModifyUser = '" & pUserID & "', "
                sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EY1   = '" & "*" & "' "
                sql &= "   And EZ1   = '" & "50" & "' "
                sql &= "   And EX1  <> '" & "" & "' "
                oDataBase.ExecuteNonQuery(sql)
                ' BY FCT 資料同步化
                UpdateBYFCTStatus(pBuyer, pUserID, "52")
            End If
            '
            ' -----------------------------------------------------------------------------------
            ' 非42/52的41/50資料    清除需處理FLAG(EY1)    
            ' -----------------------------------------------------------------------------------
            ' 非 42/52 的 41/50 BY FCT 資料同步化
            sql = "Update E_InputSheetBY Set "
            sql &= " EY1 = '', "
            sql &= " ModifyUser = '" & pUserID & "', "
            sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
            sql &= " Where Buyer = '" & pBuyer & "' "
            sql &= "   And EY1   = '" & "*" & "' "
            sql &= "   And (EZ1 = '41' or EZ1 = '50') "
            oDataBase.ExecuteNonQuery(sql)
            '
            ' -----------------------------------------------------------------------------------
            ' 30    Part With Leve2判斷(J1)→非SOLID及MULTI(含WRONG)  (空白=solid X=multi wrong=與客戶確認)
            ' -----------------------------------------------------------------------------------
            StsCode = "30"
            If oDB.GetBYFCTDataCheckFun(pBuyer, 3) = "1" Then
                sql = "Update E_InputSheetBY Set "
                sql &= " EZ1 = '30', "
                sql &= " ModifyUser = '" & pUserID & "', "
                sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EY1   = '" & "*" & "' "
                sql &= "   And J1   <> '" & "" & "' "
                sql &= "   And J1 NOT LIKE '%" & "X" & "%' "
                oDataBase.ExecuteNonQuery(sql)
                ' BY FCT 資料同步化
                UpdateBYFCTStatus(pBuyer, pUserID, "30")
            End If
            '
            ' -----------------------------------------------------------------------------------
            ' 80    Part With Leve2(J1)=空白(solid) FCT QTY檢查
            ' 81    Part With Leve2(J1)=X(MULTI) FCT QTY檢查
            ' MKT-SUZY 2013/9/23 變更   移至Check 60 Status前 
            ' -----------------------------------------------------------------------------------
            ' 80:solid FCT QTY檢查
            StsCode = "80"
            If oDB.GetBYFCTDataCheckFun(pBuyer, 13) = "1" Then
                sql = "Update E_InputSheetBY Set "
                sql &= " EZ1 = '80', "
                sql &= " ModifyUser = '" & pUserID & "', "
                sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EY1   = '" & "*" & "' "
                sql &= "   And J1    = '" & "" & "' "
                sql &= "   And CONVERT(int,AF1) <= 0 "
                oDataBase.ExecuteNonQuery(sql)
                ' BY FCT 資料同步化
                UpdateBYFCTStatus(pBuyer, pUserID, "80")
            End If
            ' 81:MULTI FCT QTY檢查
            StsCode = "81"
            If oDB.GetBYFCTDataCheckFun(pBuyer, 14) = "1" Then
                sql = "SELECT C1, D1, E1, G1, L1, "
                sql &= "SUM(CONVERT(int,AF1)) AS QTY "
                sql &= "FROM E_InputSheetBY "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EY1   = '" & "*" & "' "
                sql &= "   And J1 LIKE '%" & "X" & "%' "
                sql &= "Group by C1, D1, E1, G1, L1 "
                sql &= "Order by C1, D1, E1, G1, L1 "
                Dim dt_InputSheet1 As DataTable = oDataBase.GetDataTable(sql)
                For i = 0 To dt_InputSheet1.Rows.Count - 1
                    '
                    If dt_InputSheet1.Rows(i).Item("QTY") <= 0 Then
                        '
                        sql = "Update E_InputSheetBY Set "
                        sql &= " EY1 = '', "
                        sql &= " EZ1 = '81', "
                        sql &= " ModifyUser = '" & pUserID & "', "
                        sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                        sql &= " Where Buyer = '" & pBuyer & "' "
                        sql &= "   And C1    = '" & dt_InputSheet1.Rows(i).Item("C1").ToString & "' "
                        sql &= "   And D1    = '" & dt_InputSheet1.Rows(i).Item("D1").ToString & "' "
                        sql &= "   And E1    = '" & dt_InputSheet1.Rows(i).Item("E1").ToString & "' "
                        sql &= "   And G1    = '" & dt_InputSheet1.Rows(i).Item("G1").ToString & "' "
                        sql &= "   And L1    = '" & dt_InputSheet1.Rows(i).Item("L1").ToString & "' "
                        oDataBase.ExecuteNonQuery(sql)
                    End If
                Next
            End If
            '
            ' -----------------------------------------------------------------------------------
            ' 60/61     PLM(C1)/COLOR判斷(K1)   ELEMENT(F1)   
            '           Part With Leve2(J1)=(空白=solid X=multi wrong=與客戶確認)
            ' -----------------------------------------------------------------------------------
            ' 60:solid-PLM/COLOR判斷 
            StsCode = "60"
            If oDB.GetBYFCTDataCheckFun(pBuyer, 9) = "1" Then
                sql = "Update E_InputSheetBY Set "
                sql &= " EZ1 = '60', "
                sql &= " ModifyUser = '" & pUserID & "', "
                sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EY1   = '" & "*" & "' "
                sql &= "   And J1    = '" & "" & "' "
                sql &= "   And (C1 = '" & "" & "' or K1 = '" & "" & "') "
                oDataBase.ExecuteNonQuery(sql)
                ' BY FCT 資料同步化
                UpdateBYFCTStatus(pBuyer, pUserID, "60")
            End If
            ' 61:multi-PLM/COLOR判斷 
            StsCode = "61"
            If oDB.GetBYFCTDataCheckFun(pBuyer, 10) = "1" Then
                sql = "Update E_InputSheetBY Set "
                sql &= " EZ1 = '61', "
                sql &= " ModifyUser = '" & pUserID & "', "
                sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EY1   = '" & "*" & "' "
                sql &= "   And J1 LIKE '%" & "X" & "%' "
                sql &= "   And ( "
                sql &= "         C1 = '" & "" & "' or "
                sql &= "         (K1 = '" & "" & "' And F1 Not Like '%" & "ZIPPER" & "%') "
                sql &= "       ) "
                oDataBase.ExecuteNonQuery(sql)
                ' BY FCT 資料同步化
                UpdateBYFCTStatus(pBuyer, pUserID, "61")
            End If
            '
            ' -----------------------------------------------------------------------------------
            ' 70    拉鍊結構判斷(ELEMENT)  Part With Leve2(J1)=X(multi) / ELEMENT(F1)<>拉鍊結構(系統)
            '       拉鍊結構(系統)=系統參數 CAT=150 
            ' -----------------------------------------------------------------------------------
            StsCode = "70"
            If oDB.GetBYFCTDataCheckFun(pBuyer, 11) = "1" Then
                sql = "Update E_InputSheetBY Set "
                sql &= " EZ1 = '70', "
                sql &= " ModifyUser = '" & pUserID & "', "
                sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EY1   = '" & "*" & "' "
                sql &= "   And J1 LIKE '%" & "X" & "%' "
                sql &= "   And F1 Not In (Select Data From M_Referp Where Cat='150' And DKey='" & pBuyer & "' Group By Data) "
                oDataBase.ExecuteNonQuery(sql)
                ' BY FCT 資料同步化
                UpdateBYFCTStatus(pBuyer, pUserID, "70")
            End If
            '
            ' -----------------------------------------------------------------------------------
            ' 71    ITEM & COLOR 轉換3要素(TAPE, TEETH, PULLER)檢查     
            '       Part With Leve2(J1)=X(multi)    PLM(C1), ELEMENT(F1), COLOR(K1)
            ' -----------------------------------------------------------------------------------
            StsCode = "71"
            If oDB.GetBYFCTDataCheckFun(pBuyer, 12) = "1" Then
                Dim TapeColor, TeethColor, PullerColor As Integer
                sql = "SELECT C1, D1, E1, G1, L1 FROM E_InputSheetBY "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EY1   = '" & "*" & "' "
                sql &= "   And J1 LIKE '%" & "X" & "%' "
                sql &= "Group by C1, D1, E1, G1, L1 "
                sql &= "Order by C1, D1, E1, G1, L1 "
                Dim dt_InputSheet As DataTable = oDataBase.GetDataTable(sql)
                For i = 0 To dt_InputSheet.Rows.Count - 1
                    TapeColor = 0
                    TeethColor = 0
                    PullerColor = 0
                    '
                    sql = "SELECT F1 FROM E_InputSheetBY "
                    sql &= " Where Buyer = '" & pBuyer & "' "
                    sql &= "   And C1    = '" & dt_InputSheet.Rows(i).Item("C1").ToString & "' "
                    sql &= "   And D1    = '" & dt_InputSheet.Rows(i).Item("D1").ToString & "' "
                    sql &= "   And E1    = '" & dt_InputSheet.Rows(i).Item("E1").ToString & "' "
                    sql &= "   And G1    = '" & dt_InputSheet.Rows(i).Item("G1").ToString & "' "
                    sql &= "   And L1    = '" & dt_InputSheet.Rows(i).Item("L1").ToString & "' "
                    sql &= "Order by F1 "
                    Dim dt_ELEMENT As DataTable = oDataBase.GetDataTable(sql)
                    For j = 0 To dt_ELEMENT.Rows.Count - 1
                        If InStr(dt_ELEMENT.Rows(j).Item("F1").ToString.ToUpper, "TAPE") > 0 Then TapeColor = TapeColor + 1
                        If InStr(dt_ELEMENT.Rows(j).Item("F1").ToString.ToUpper, "TEETH") > 0 Then TeethColor = TeethColor + 1
                        If InStr(dt_ELEMENT.Rows(j).Item("F1").ToString.ToUpper, "PULLER") > 0 Then PullerColor = PullerColor + 1
                    Next
                    '
                    If TapeColor = 0 Or TeethColor = 0 Or PullerColor = 0 Then
                        sql = "Update E_InputSheetBY Set "
                        sql &= " EY1 = '', "
                        sql &= " EZ1 = '71', "
                        sql &= " ModifyUser = '" & pUserID & "', "
                        sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                        sql &= " Where Buyer = '" & pBuyer & "' "
                        sql &= "   And C1    = '" & dt_InputSheet.Rows(i).Item("C1").ToString & "' "
                        sql &= "   And D1    = '" & dt_InputSheet.Rows(i).Item("D1").ToString & "' "
                        sql &= "   And E1    = '" & dt_InputSheet.Rows(i).Item("E1").ToString & "' "
                        sql &= "   And G1    = '" & dt_InputSheet.Rows(i).Item("G1").ToString & "' "
                        sql &= "   And L1    = '" & dt_InputSheet.Rows(i).Item("L1").ToString & "' "
                        oDataBase.ExecuteNonQuery(sql)
                    End If
                Next
            End If
            ' 01:國內進口半成品 
            ' AG1 --> [材料類別/表面處理/進口SLIDER/進口CHAIN]-ITEMDESCR
            StsCode = "01"
            If oDB.GetBYFCTDataCheckFun(pBuyer, 15) = "1" Then
                sql = "Update E_InputSheetBY Set "
                sql &= " EZ1 = '01', "
                sql &= " ModifyUser = '" & pUserID & "', "
                sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EY1   = '" & "*" & "' "
                sql &= "   And AG1 Not LIKE '%" & "/TW/TW" & "%' "
                oDataBase.ExecuteNonQuery(sql)
                ' BY FCT 資料同步化
                UpdateBYFCTStatus(pBuyer, pUserID, "01")
            End If            '
            ' -----------------------------------------------------------------------------------
            ' 清除EZ1='00'的需處理FLAG(EY1)     EZ1='00' AND EY1='*' --> EY1=空白
            ' -----------------------------------------------------------------------------------
            sql = "Update E_InputSheetBY Set "
            sql &= " EY1 = '', "
            sql &= " ModifyUser = '" & pUserID & "', "
            sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
            sql &= " Where Buyer = '" & pBuyer & "' "
            sql &= "   And EY1   = '" & "*" & "' "
            sql &= "   And EZ1   = '" & "00" & "' "
            oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "BYFCTDataCheck", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "Status=[" + StsCode + "]", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'BYFCTDataCheck-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([160]-NewBYFCTImport)
    '**     New Buyer FCT Data Import
    '***********************************************************************************************
    'NewBYFCTImport-Start
    <WebMethod()> _
    Public Function NewBYFCTImport(ByVal pLogID As String, _
                                   ByVal pBuyer As String, _
                                   ByVal pUserID As String, _
                                   ByVal pGRBuyer As String, _
                                   ByVal pFunList As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i As Integer
        Dim sql As String
        '
        Try
            ' -----------------------------------------------------------------------------------
            ' 清除前回 BY FCT 資料
            ' -----------------------------------------------------------------------------------
            sql = "Delete From E_InputSheetBY "
            sql &= "Where Buyer = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
            '
            ' -----------------------------------------------------------------------------------
            ' 將WF中BY FCT Qty合計後轉至正式FILE             
            ' -----------------------------------------------------------------------------------
            sql = "SELECT "
            sql &= "Buyer, "
            sql &= "A1, B1, C1, D1, E1, F1, G1, H1, I1, J1, K1, L1, M1, N1, O1, P1, Q1, R1, "
            sql &= "AG1, AH1, AI1, "
            sql &= "EY1, EZ1, "
            sql &= "sum(convert(int,S1)) As N_F, "
            sql &= "sum(convert(int,T1)) As N1_F, "
            sql &= "sum(convert(int,U1)) As N2_F, "
            sql &= "sum(convert(int,V1)) As N3_F, "
            sql &= "sum(convert(int,W1)) As N4_F, "
            sql &= "sum(convert(int,X1)) As N5_F, "
            sql &= "sum(convert(int,Y1)) As N6_F, "
            sql &= "sum(convert(int,Z1)) As N7_F, "
            sql &= "sum(convert(int,AA1)) As N8_F, "
            sql &= "sum(convert(int,AB1)) As N9_F, "
            sql &= "sum(convert(int,AC1)) As N10_F, "
            sql &= "sum(convert(int,AD1)) As N11_F, "
            sql &= "sum(convert(int,AE1)) As N12_F, "
            sql &= "sum(convert(int,S1)) + sum(convert(int,T1)) + sum(convert(int,U1)) + sum(convert(int,V1)) + sum(convert(int,W1)) + "
            sql &= "sum(convert(int,X1)) + sum(convert(int,Y1)) + sum(convert(int,Z1)) + sum(convert(int,AA1)) +  "
            sql &= "sum(convert(int,AB1)) + sum(convert(int,AC1)) + sum(convert(int,AD1)) + sum(convert(int,AE1)) As TOTAL "
            sql &= "FROM W_InputSheetBY "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "Group by Buyer, A1, B1, C1, D1, E1, F1, G1, H1, I1, J1, K1, L1, M1, N1, O1, P1, Q1, R1, AG1, AH1, AI1, EY1, EZ1 "
            sql &= "Order by Buyer, A1, B1, C1, D1, E1, F1, G1, H1, I1, J1, K1, L1, M1, N1, O1, P1, Q1, R1, AG1, AH1, AI1, EY1, EZ1 "
            Dim dt_InputSheet As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_InputSheet.Rows.Count - 1
                '
                sql = "Insert into E_InputSheetBY ( "
                sql &= "Buyer, "
                sql &= "A1, B1, C1, D1, E1, F1, G1, H1, I1, J1, K1, L1, M1, N1, O1, P1, Q1, R1, "
                sql &= "S1, T1, U1, V1, W1, X1, Y1, Z1, AA1, AB1, AC1, AD1, AE1, AF1, "
                sql &= "AG1, AH1, AI1, "
                sql &= "EY1, EZ1, "
                sql &= "CreateUser, CreateTime "
                sql &= ") "
                sql &= "VALUES( "
                sql &= "'" & dt_InputSheet.Rows(i).Item("Buyer").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("A1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("B1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("C1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("D1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("E1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("F1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("G1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("H1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("I1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("J1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("K1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("L1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("M1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("N1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("O1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("P1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("Q1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("R1").ToString & "', "
                '
                sql &= "'" & dt_InputSheet.Rows(i).Item("N_F").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("N1_F").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("N2_F").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("N3_F").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("N4_F").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("N5_F").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("N6_F").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("N7_F").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("N8_F").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("N9_F").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("N10_F").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("N11_F").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("N12_F").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("TOTAL").ToString & "', "
                '
                sql &= "'" & dt_InputSheet.Rows(i).Item("AG1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("AH1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("AI1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("EY1").ToString & "', "
                sql &= "'" & dt_InputSheet.Rows(i).Item("EZ1").ToString & "', "
                '
                sql &= "'" & pUserID & "', "
                sql &= "'" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= ") "
                oDataBase.ExecuteNonQuery(sql)
            Next
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "NewBYFCTImport", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "NewBYFCTImport", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'NewBYFCTImport-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([170]-ConvertToBYFCT)
    '**     轉換成FAS使用FCT DATA
    '***********************************************************************************************
    'ConvertToBYFCT-Start
    <WebMethod()> _
    Public Function ConvertToBYFCT(ByVal pLogID As String, _
                                   ByVal pBuyer As String, _
                                   ByVal pUserID As String, _
                                   ByVal pGRBuyer As String, _
                                   ByVal pFunList As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i, j, k As Integer
        Dim xSum(14) As Integer
        Dim sql, xBuyerCode, xColorString, xColorCode As String
        '
        'Try
        ' -----------------------------------------------------------------------------------
        ' 設定 BuyerName, Color Array
        ' -----------------------------------------------------------------------------------
        xBuyerCode = ""
        If InStr(pBuyer, "000001") > 0 Then xBuyerCode = "000001"
        If InStr(pBuyer, "000016") > 0 Then xBuyerCode = "000016"
        ' -----------------------------------------------------------------------------------
        ' 檢查是否有待處理資料
        ' -----------------------------------------------------------------------------------
        sql = "SELECT Top 1 Unique_ID FROM E_InputSheetBY "
        sql &= " Where Buyer = '" & pBuyer & "' "
        sql &= "   And EY1   = '" & "*" & "' "
        sql &= "Order by Unique_ID "
        Dim dt_CheckInputSheet As DataTable = oDataBase.GetDataTable(sql)
        If dt_CheckInputSheet.Rows.Count > 0 Then
            RtnCode = 1         ' 有待處理資料
        End If
        ' -----------------------------------------------------------------------------------
        ' 清除前回 FAS使用FCT DATA
        ' -----------------------------------------------------------------------------------
        If RtnCode = 0 Then
            sql = "Delete From E_BYFCTSheet "
            sql &= "Where Buyer = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
        End If
        ' -----------------------------------------------------------------------------------
        ' SOLID TYPE DATA   轉換成FAS使用FCT DATA 
        ' -----------------------------------------------------------------------------------
        If RtnCode = 0 Then
            '
            sql = "SELECT C1, D1, E1, G1, L1 FROM E_InputSheetBY "
            sql &= " Where Buyer = '" & pBuyer & "' "
            sql &= "   And EZ1   < '" & "10" & "' "
            sql &= "   And J1    = '" & "" & "' "
            sql &= "Group by C1, D1, E1, G1, L1 "
            sql &= "Order by C1, D1, E1, G1, L1 "
            Dim dt_InputSheet As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_InputSheet.Rows.Count - 1
                '
                sql = "SELECT * FROM E_InputSheetBY "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EZ1   < '" & "10" & "' "
                sql &= "   And J1    = '" & "" & "' "
                sql &= "   And C1    = '" & dt_InputSheet.Rows(i).Item("C1").ToString & "' "
                sql &= "   And D1    = '" & dt_InputSheet.Rows(i).Item("D1").ToString & "' "
                sql &= "   And E1    = '" & dt_InputSheet.Rows(i).Item("E1").ToString & "' "
                sql &= "   And G1    = '" & dt_InputSheet.Rows(i).Item("G1").ToString & "' "
                sql &= "   And L1    = '" & dt_InputSheet.Rows(i).Item("L1").ToString & "' "
                sql &= "Order by Unique_ID "
                Dim dt_SourceData As DataTable = oDataBase.GetDataTable(sql)
                For j = 0 To dt_SourceData.Rows.Count - 1
                    '
                    sql = "Insert into E_BYFCTSheet ( "
                    sql &= "Buyer, "
                    sql &= "BRAND,	SUPPLIER, PLM, STYLE, ARTICLE, PART, COLOR, CUSTOMER, CUSTOMERCODE, GARMENT_LT, GRANGE, SEASON, "
                    sql &= "N_F, N1_F, N2_F, N3_F, N4_F, N5_F, N6_F, N7_F, N8_F, N9_F, N10_F, N11_F, N12_F, TOTAL, "
                    sql &= "CreateUser, CreateTime, ModifyUser "
                    sql &= ") "
                    sql &= "VALUES( "
                    sql &= "'" & dt_SourceData.Rows(j).Item("Buyer").ToString & "', "   ' BUYER

                    sql &= "'" & dt_SourceData.Rows(j).Item("A1").ToString & "', "      ' BRAND    
                    sql &= "'" & dt_SourceData.Rows(j).Item("B1").ToString & "', "      ' SUPPLIER
                    sql &= "'" & dt_SourceData.Rows(j).Item("C1").ToString & "', "      ' PLM
                    sql &= "'" & dt_SourceData.Rows(j).Item("D1").ToString & "', "      ' STYLE
                    sql &= "'" & dt_SourceData.Rows(j).Item("E1").ToString & "', "      ' ARTICLE
                    sql &= "'" & dt_SourceData.Rows(j).Item("G1").ToString & "', "      ' PART
                    ' 設定COLOR
                    xColorString = "////////////"                        ' 1/2/3/4/5/6/7/8/9/10/11/12
                    xColorCode = Mid(dt_SourceData.Rows(j).Item("K1").ToString.Trim, Len(dt_SourceData.Rows(j).Item("K1").ToString.Trim) - 3, 4)
                    sql &= "'" & GetColor(1, xBuyerCode, xColorString, dt_SourceData.Rows(j).Item("F1").ToString, xColorCode) & "', "       ' COLOR
                    ' 設定CUSTOMER
                    sql &= "'" & GetNativeVendorCode(dt_SourceData.Rows(j).Item("Buyer").ToString, dt_SourceData.Rows(j).Item("L1").ToString, "NAME") & "', "     ' CUSTOMER
                    sql &= "'" & GetNativeVendorCode(dt_SourceData.Rows(j).Item("Buyer").ToString, dt_SourceData.Rows(j).Item("L1").ToString, "CODE") & "', "     ' CUSTOMERCODE
                    ' GARMENT_LT
                    If dt_SourceData.Rows(j).Item("P1").ToString = "" Then
                        sql &= "'" & dt_SourceData.Rows(j).Item("O1").ToString & "', "  ' GARMENT_LT
                    Else
                        sql &= "'" & dt_SourceData.Rows(j).Item("P1").ToString & "', "  ' GARMENT_LT(REVERSE)
                    End If
                    sql &= "'" & dt_SourceData.Rows(j).Item("Q1").ToString & "', "      ' GARMENT_RANGE
                    sql &= "'" & dt_SourceData.Rows(j).Item("R1").ToString & "', "      ' SEASON
                    '
                    sql &= "" & dt_SourceData.Rows(j).Item("S1").ToString & ", "      ' N_F
                    sql &= "" & dt_SourceData.Rows(j).Item("T1").ToString & ", "      ' N1_F
                    sql &= "" & dt_SourceData.Rows(j).Item("U1").ToString & ", "      ' N2_F
                    sql &= "" & dt_SourceData.Rows(j).Item("V1").ToString & ", "      ' N3_F
                    sql &= "" & dt_SourceData.Rows(j).Item("W1").ToString & ", "      ' N4_F
                    sql &= "" & dt_SourceData.Rows(j).Item("X1").ToString & ", "      ' N5_F
                    sql &= "" & dt_SourceData.Rows(j).Item("Y1").ToString & ", "      ' N6_F
                    sql &= "" & dt_SourceData.Rows(j).Item("Z1").ToString & ", "      ' N7_F
                    sql &= "" & dt_SourceData.Rows(j).Item("AA1").ToString & ", "     ' N8_F
                    sql &= "" & dt_SourceData.Rows(j).Item("AB1").ToString & ", "     ' N9_F
                    sql &= "" & dt_SourceData.Rows(j).Item("AC1").ToString & ", "     ' N10_F
                    sql &= "" & dt_SourceData.Rows(j).Item("AD1").ToString & ", "     ' N11_F
                    sql &= "" & dt_SourceData.Rows(j).Item("AE1").ToString & ", "     ' N12_F
                    sql &= "" & dt_SourceData.Rows(j).Item("AF1").ToString & ", "     ' TOTAL
                    '
                    sql &= "'" & pUserID & "', "
                    sql &= "'" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "', "
                    sql &= "'" & "SOLID" & "' "
                    sql &= ") "
                    oDataBase.ExecuteNonQuery(sql)
                Next
            Next
        End If
        ' -----------------------------------------------------------------------------------
        ' MULTI TYPE DATA   轉換成FAS使用FCT DATA 
        ' -----------------------------------------------------------------------------------
        If RtnCode = 0 Then
            '
            sql = "SELECT C1, D1, E1, G1, L1 FROM E_InputSheetBY "
            sql &= " Where Buyer = '" & pBuyer & "' "
            sql &= "   And EZ1   < '" & "10" & "' "
            sql &= "   And J1 LIKE '%" & "X" & "%' "
            ' Test-Add-Start
            'sql &= "   And C1 = '" & "61030096" & "' "
            'sql &= "   And D1 = '" & "SMSURUS14BMW15" & "' "
            'sql &= "   And E1 = '" & "D87169" & "' "
            'sql &= "   And G1 = '" & "220" & "' "
            'sql &= "   And L1 = '" & "T.M.T." & "' "
            ' Test-Add-End
            sql &= "Group by C1, D1, E1, G1, L1 "
            sql &= "Order by C1, D1, E1, G1, L1 "
            Dim dt_InputSheet As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_InputSheet.Rows.Count - 1
                '2014/5/22 Debug-Start 無SUM造成有 Forecast Qty=0
                For k = 1 To 14
                    xSum(k) = 0
                Next
                'Debug-End
                xColorString = "////////////"                        ' 1/2/3/4/5/6/7/8/9/10/11/12
                '
                sql = "SELECT * FROM E_InputSheetBY "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And EZ1   < '" & "10" & "' "
                sql &= "   And J1 LIKE '%" & "X" & "%' "
                sql &= "   And C1    = '" & dt_InputSheet.Rows(i).Item("C1").ToString & "' "
                sql &= "   And D1    = '" & dt_InputSheet.Rows(i).Item("D1").ToString & "' "
                sql &= "   And E1    = '" & dt_InputSheet.Rows(i).Item("E1").ToString & "' "
                sql &= "   And G1    = '" & dt_InputSheet.Rows(i).Item("G1").ToString & "' "
                sql &= "   And L1    = '" & dt_InputSheet.Rows(i).Item("L1").ToString & "' "
                sql &= "Order by Unique_ID "
                Dim dt_SourceData As DataTable = oDataBase.GetDataTable(sql)
                For j = 0 To dt_SourceData.Rows.Count - 1
                    '
                    '2014/5/22 Debug-Start 無SUM造成有 Forecast Qty=0
                    xSum(1) = xSum(1) + dt_SourceData.Rows(j).Item("S1")        ' N_F
                    xSum(2) = xSum(2) + dt_SourceData.Rows(j).Item("T1")        ' N1_F
                    xSum(3) = xSum(3) + dt_SourceData.Rows(j).Item("U1")        ' N2_F
                    xSum(4) = xSum(4) + dt_SourceData.Rows(j).Item("V1")        ' N3_F
                    xSum(5) = xSum(5) + dt_SourceData.Rows(j).Item("W1")        ' N4_F
                    xSum(6) = xSum(6) + dt_SourceData.Rows(j).Item("X1")        ' N5_F
                    xSum(7) = xSum(7) + dt_SourceData.Rows(j).Item("Y1")        ' N6_F
                    xSum(8) = xSum(8) + dt_SourceData.Rows(j).Item("Z1")        ' N7_F
                    xSum(9) = xSum(9) + dt_SourceData.Rows(j).Item("AA1")       ' N8_F
                    xSum(10) = xSum(10) + dt_SourceData.Rows(j).Item("AB1")     ' N9_F
                    xSum(11) = xSum(11) + dt_SourceData.Rows(j).Item("AC1")     ' N10_F
                    xSum(12) = xSum(12) + dt_SourceData.Rows(j).Item("AD1")     ' N11_F
                    xSum(13) = xSum(13) + dt_SourceData.Rows(j).Item("AE1")     ' N12_F
                    xSum(14) = xSum(14) + dt_SourceData.Rows(j).Item("AF1")     ' TOTAL
                    'Debug-End
                    If j = dt_SourceData.Rows.Count - 1 Then
                        '
                        If dt_SourceData.Rows(j).Item("K1").ToString.Trim <> "" Then
                            xColorCode = Mid(dt_SourceData.Rows(j).Item("K1").ToString.Trim, Len(dt_SourceData.Rows(j).Item("K1").ToString.Trim) - 3, 4)
                            xColorString = GetColor(2, xBuyerCode, xColorString, dt_SourceData.Rows(j).Item("F1").ToString, xColorCode)
                        End If
                        '
                        sql = "Insert into E_BYFCTSheet ( "
                        sql &= "Buyer, "
                        sql &= "BRAND,	SUPPLIER, PLM, STYLE, ARTICLE, PART, COLOR, CUSTOMER, CUSTOMERCODE, GARMENT_LT, GRANGE, SEASON, "
                        sql &= "N_F, N1_F, N2_F, N3_F, N4_F, N5_F, N6_F, N7_F, N8_F, N9_F, N10_F, N11_F, N12_F, TOTAL, "
                        sql &= "CreateUser, CreateTime, ModifyUser "
                        sql &= ") "
                        sql &= "VALUES( "
                        sql &= "'" & dt_SourceData.Rows(j).Item("Buyer").ToString & "', "   ' BUYER
                        sql &= "'" & dt_SourceData.Rows(j).Item("A1").ToString & "', "      ' BRAND    
                        sql &= "'" & dt_SourceData.Rows(j).Item("B1").ToString & "', "      ' SUPPLIER
                        sql &= "'" & dt_SourceData.Rows(j).Item("C1").ToString & "', "      ' PLM
                        sql &= "'" & dt_SourceData.Rows(j).Item("D1").ToString & "', "      ' STYLE
                        sql &= "'" & dt_SourceData.Rows(j).Item("E1").ToString & "', "      ' ARTICLE
                        sql &= "'" & dt_SourceData.Rows(j).Item("G1").ToString & "', "      ' PART
                        ' 設定COLOR
                        sql &= "'" & xColorString & "', "                                   ' COLOR
                        ' 設定CUSTOMER
                        sql &= "'" & GetNativeVendorCode(dt_SourceData.Rows(j).Item("Buyer").ToString, dt_SourceData.Rows(j).Item("L1").ToString, "NAME") & "', "     ' CUSTOMER
                        sql &= "'" & GetNativeVendorCode(dt_SourceData.Rows(j).Item("Buyer").ToString, dt_SourceData.Rows(j).Item("L1").ToString, "CODE") & "', "     ' CUSTOMERCODE
                        ' GARMENT_LT
                        If dt_SourceData.Rows(j).Item("P1").ToString = "" Then
                            sql &= "'" & dt_SourceData.Rows(j).Item("O1").ToString & "', "  ' GARMENT_LT
                        Else
                            sql &= "'" & dt_SourceData.Rows(j).Item("P1").ToString & "', "  ' GARMENT_LT(REVERSE)
                        End If
                        sql &= "'" & dt_SourceData.Rows(j).Item("Q1").ToString & "', "      ' GARMENT_RANGE
                        sql &= "'" & dt_SourceData.Rows(j).Item("R1").ToString & "', "      ' SEASON
                        '
                        sql &= "" & CStr(xSum(1)) & ", "      ' N_F
                        sql &= "" & CStr(xSum(2)) & ", "      ' N1_F
                        sql &= "" & CStr(xSum(3)) & ", "      ' N2_F
                        sql &= "" & CStr(xSum(4)) & ", "      ' N3_F
                        sql &= "" & CStr(xSum(5)) & ", "      ' N4_F
                        sql &= "" & CStr(xSum(6)) & ", "      ' N5_F
                        sql &= "" & CStr(xSum(7)) & ", "      ' N6_F
                        sql &= "" & CStr(xSum(8)) & ", "      ' N7_F
                        sql &= "" & CStr(xSum(9)) & ", "     ' N8_F
                        sql &= "" & CStr(xSum(10)) & ", "     ' N9_F
                        sql &= "" & CStr(xSum(11)) & ", "     ' N10_F
                        sql &= "" & CStr(xSum(12)) & ", "     ' N11_F
                        sql &= "" & CStr(xSum(13)) & ", "     ' N12_F
                        sql &= "" & CStr(xSum(14)) & ", "     ' TOTAL
                        '
                        sql &= "'" & pUserID & "', "
                        sql &= "'" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "', "
                        sql &= "'" & "MULTI" & "' "
                        sql &= ") "
                        oDataBase.ExecuteNonQuery(sql)
                    Else
                        If dt_SourceData.Rows(j).Item("K1").ToString.Trim <> "" Then
                            xColorCode = Mid(dt_SourceData.Rows(j).Item("K1").ToString.Trim, Len(dt_SourceData.Rows(j).Item("K1").ToString.Trim) - 3, 4)
                            xColorString = GetColor(2, xBuyerCode, xColorString, dt_SourceData.Rows(j).Item("F1").ToString, xColorCode)       ' COLOR
                        End If
                    End If
                Next
                '
            Next
        End If
        'Catch ex As Exception
        'RtnCode = 9
        'oDB.AccessLog(pLogID, pBuyer, "ConvertToBYFCT", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "ConvertToBYFCT", pUserID, "")
        'End Try
        '
        Return RtnCode
    End Function
    'ConvertToBYFCT-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([180]-EDITransfer   						                             
    '**     轉換EDI-DATA
    '***********************************************************************************************
    'EDITransfer-Start
    <WebMethod()> _
    Public Function EDITransfer(ByVal pLogID As String, ByVal pBuyer As String, _
                                ByVal pUserID As String, ByVal pGRBuyer As String, _
                                ByVal pFunList As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i, j As Integer
        Dim xReqDate(4) As Integer
        Dim xMaxQty, xProdQty As Double
        Dim sql, xYYMM As String
        Dim xLSNo, xOrderMonth, xLengthUnit As String
        Dim xSeqNo, xOrderZone As Integer
        '
        Try
            sql = "SELECT * FROM LF_BuyerLocalStockPlan "
            sql &= "Where LogID = '" & pLogID & "' "
            sql &= "  And BuyerGroup = '" & pGRBuyer & "' "
            sql &= "  And Version = " & "99" & " "
            sql &= "  And GR_01 = '" & "LS" & "' "
            sql &= "Order by GR_02, GR_03, GR_07, GR_08 "
            Dim dt_LocalStockPlan As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_LocalStockPlan.Rows.Count - 1
                ' 取得MaxQty
                ' ------------------------
                xMaxQty = 100000000     ' 1億
                sql = "SELECT * FROM M_Referp "
                sql &= " Where Cat = '" & "114" & "' "
                sql &= "   And DKey = '" & "MAXQTY-" + dt_LocalStockPlan.Rows(i).Item("GR_08").ToString.Trim + "-" + pBuyer & "' "
                sql &= "Order by DKey, Data "
                Dim dt_Referp As DataTable = oDataBase.GetDataTable(sql)
                If dt_Referp.Rows.Count > 0 Then
                    xMaxQty = CDbl(dt_Referp.Rows(0).Item("Data").ToString.Trim)
                Else
                    dt_Referp.Clear()
                    sql = "SELECT * FROM M_Referp "
                    sql &= " Where Cat = '" & "114" & "' "
                    sql &= "   And DKey = '" & "MAXQTY-" + dt_LocalStockPlan.Rows(i).Item("GR_08").ToString.Trim & "' "
                    sql &= "Order by DKey, Data "
                    dt_Referp = oDataBase.GetDataTable(sql)
                    If dt_Referp.Rows.Count > 0 Then
                        xMaxQty = CDbl(dt_Referp.Rows(0).Item("Data").ToString.Trim)
                    End If
                End If
                ' 取得N+1~N+4納期
                ' ------------------------
                For j = 1 To 4
                    xReqDate(j) = j * 30
                Next
                dt_Referp.Clear()
                sql = "SELECT * FROM M_Referp "
                sql &= " Where Cat = '" & "114" & "' "
                sql &= "   And DKey Like '%" & pBuyer & "-REQDATE" & "%' "
                sql &= "Order by DKey, Data "
                dt_Referp = oDataBase.GetDataTable(sql)
                For j = 0 To dt_Referp.Rows.Count - 1
                    xReqDate(j + 1) = CInt(dt_Referp.Rows(j).Item("Data").ToString.Trim)
                Next
                ' 取得BUYER ORDER ZONE
                ' ------------------------
                xOrderZone = 4
                dt_Referp.Clear()
                sql = "SELECT * FROM M_Referp "
                sql &= " Where Cat = '" & "114" & "' "
                sql &= "   And DKey = '" & "ZONE-" + pBuyer & "' "
                sql &= "Order by DKey, Data "
                dt_Referp = oDataBase.GetDataTable(sql)
                If dt_Referp.Rows.Count > 0 Then
                    xOrderZone = CDbl(dt_Referp.Rows(0).Item("Data").ToString.Trim)
                End If
                'MsgBox("[" + pBuyer & "-REQDATE" + "]" + "[" + "ZONE-" + pBuyer + "]" + "[" + CStr(xOrderZone) + "]")
                ' 共同資料準備 (BULSNo, xLengthUnit)
                ' ------------------------
                'If dt_LocalStockPlan.Rows(i).Item("BULSSubNo") < 10 Then
                '    xLSNo = dt_LocalStockPlan.Rows(i).Item("BULSNo") + "0" + dt_LocalStockPlan.Rows(i).Item("BULSSubNo").ToString
                'Else
                '    xLSNo = dt_LocalStockPlan.Rows(i).Item("BULSNo") + dt_LocalStockPlan.Rows(i).Item("BULSSubNo").ToString
                'End If
                xLSNo = dt_LocalStockPlan.Rows(i).Item("BULSNo")

                If dt_LocalStockPlan.Rows(i).Item("GR_08") = "CH" Then
                    xLengthUnit = "M"
                Else
                    xLengthUnit = "P"
                End If
                ' 
                ' N (只限VENDOR)
                ' ------------------------
                If pGRBuyer = "VENDORB" Then
                    If xOrderZone >= 0 Then
                        ' 準備資料 (OrderMonth, xYYMM, xSeqNo)
                        If CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) < 10 Then
                            xOrderMonth = "0" + CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)))
                            xYYMM = Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2) + xOrderMonth
                        Else
                            If CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) > 12 Then
                                xOrderMonth = "0" + CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) - 12)
                                xYYMM = CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2)) + 1) + xOrderMonth
                            Else
                                xOrderMonth = CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)))
                                xYYMM = Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2) + xOrderMonth
                            End If
                        End If
                        xSeqNo = 1
                        ' 製作 EDI 資料
                        xProdQty = dt_LocalStockPlan.Rows(i).Item("PS_00")
                        While xProdQty > 0
                            sql = "Insert into B_LS2EDIInterface "
                            sql &= "( "
                            sql &= "BuyerGroup, BULSNo, BULSSubNo, OrderMonth, Seqno, LSType, KeepCode, Item, ItemName1, ItemName2, ItemName3, Color, PartType, "
                            sql &= "Yobi1, Yobi2, Reference, LengthUnit, BuyerCode, ReqDate, ProdQty, CreateUser, CreateTime"
                            sql &= " ) "
                            sql &= "VALUES( "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("BuyerGroup") & "', "
                            sql &= " '" & xLSNo & "', "
                            sql &= " " & dt_LocalStockPlan.Rows(i).Item("BULSSubNo") & ", "
                            sql &= " '" & xOrderMonth & "', "
                            sql &= " " & CStr(xSeqNo) & ", "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_01") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_02") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_03") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_04") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_05") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_06") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_07") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_08") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_09") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_10") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("LogID") + " N " + xYYMM + " BUY" & "', "
                            sql &= " '" & xLengthUnit & "', "
                            sql &= " '" & Mid(dt_LocalStockPlan.Rows(i).Item("BuyerGroup"), 1, 6) & "', "
                            sql &= " " & Now.AddDays(xReqDate(1)).ToString("yyyyMMdd") & ", "
                            If xProdQty <= xMaxQty Then
                                sql &= " " & CStr(xProdQty) & ", "
                                xProdQty = 0
                            Else
                                sql &= " " & CStr(xMaxQty) & ", "
                                xProdQty = xProdQty - xMaxQty
                            End If
                            sql &= " '" & "EDI" & "', "
                            sql &= " '" & NowDateTime & "' "
                            sql &= " ) "
                            oDataBase.ExecuteNonQuery(sql)

                            xSeqNo = xSeqNo + 1
                        End While
                    End If
                End If
                ' N+1
                ' ------------------------
                If xOrderZone >= 1 Then
                    ' 準備資料 (OrderMonth, xYYMM, xSeqNo)
                    If CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 1 < 10 Then
                        xOrderMonth = "0" + CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 1)
                        xYYMM = Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2) + xOrderMonth
                    Else
                        If CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 1 > 12 Then
                            xOrderMonth = "0" + CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 1 - 12)
                            xYYMM = CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2)) + 1) + xOrderMonth
                        Else
                            xOrderMonth = CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 1)
                            xYYMM = Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2) + xOrderMonth
                        End If
                    End If
                    xSeqNo = 1
                    ' 製作 EDI 資料
                    xProdQty = dt_LocalStockPlan.Rows(i).Item("PS_01")
                    While xProdQty > 0
                        sql = "Insert into B_LS2EDIInterface "
                        sql &= "( "
                        sql &= "BuyerGroup, BULSNo, BULSSubNo, OrderMonth, Seqno, LSType, KeepCode, Item, ItemName1, ItemName2, ItemName3, Color, PartType, "
                        sql &= "Yobi1, Yobi2, Reference, LengthUnit, BuyerCode, ReqDate, ProdQty, CreateUser, CreateTime"
                        sql &= " ) "
                        sql &= "VALUES( "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("BuyerGroup") & "', "
                        sql &= " '" & xLSNo & "', "
                        sql &= " " & dt_LocalStockPlan.Rows(i).Item("BULSSubNo") & ", "
                        sql &= " '" & xOrderMonth & "', "
                        sql &= " " & CStr(xSeqNo) & ", "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_01") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_02") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_03") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_04") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_05") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_06") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_07") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_08") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_09") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_10") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("LogID") + " N+1 " + xYYMM + " BUY" & "', "
                        sql &= " '" & xLengthUnit & "', "
                        sql &= " '" & Mid(dt_LocalStockPlan.Rows(i).Item("BuyerGroup"), 1, 6) & "', "
                        sql &= " " & Now.AddDays(xReqDate(1)).ToString("yyyyMMdd") & ", "
                        If xProdQty <= xMaxQty Then
                            sql &= " " & CStr(xProdQty) & ", "
                            xProdQty = 0
                        Else
                            sql &= " " & CStr(xMaxQty) & ", "
                            xProdQty = xProdQty - xMaxQty
                        End If
                        sql &= " '" & "EDI" & "', "
                        sql &= " '" & NowDateTime & "' "
                        sql &= " ) "
                        oDataBase.ExecuteNonQuery(sql)

                        xSeqNo = xSeqNo + 1
                    End While
                End If
                ' N+2
                ' ------------------------
                If xOrderZone >= 2 Then
                    ' 準備資料 (OrderMonth, xYYMM, xSeqNo)
                    If CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 2 < 10 Then
                        xOrderMonth = "0" + CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 2)
                        xYYMM = Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2) + xOrderMonth
                    Else
                        If CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 2 > 12 Then
                            xOrderMonth = "0" + CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 2 - 12)
                            xYYMM = CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2)) + 1) + xOrderMonth
                        Else
                            xOrderMonth = CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 2)
                            xYYMM = Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2) + xOrderMonth
                        End If
                    End If
                    xSeqNo = 1
                    ' 製作 EDI 資料
                    xProdQty = dt_LocalStockPlan.Rows(i).Item("PS_02")
                    While xProdQty > 0
                        sql = "Insert into B_LS2EDIInterface "
                        sql &= "( "
                        sql &= "BuyerGroup, BULSNo, BULSSubNo, OrderMonth, Seqno, LSType, KeepCode, Item, ItemName1, ItemName2, ItemName3, Color, PartType, "
                        sql &= "Yobi1, Yobi2, Reference, LengthUnit, BuyerCode, ReqDate, ProdQty, CreateUser, CreateTime"
                        sql &= " ) "
                        sql &= "VALUES( "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("BuyerGroup") & "', "
                        sql &= " '" & xLSNo & "', "
                        sql &= " " & dt_LocalStockPlan.Rows(i).Item("BULSSubNo") & ", "
                        sql &= " '" & xOrderMonth & "', "
                        sql &= " " & CStr(xSeqNo) & ", "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_01") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_02") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_03") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_04") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_05") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_06") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_07") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_08") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_09") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_10") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("LogID") + " N+2 " + xYYMM + " BUY" & "', "
                        sql &= " '" & xLengthUnit & "', "
                        sql &= " '" & Mid(dt_LocalStockPlan.Rows(i).Item("BuyerGroup"), 1, 6) & "', "
                        sql &= " " & Now.AddDays(xReqDate(2)).ToString("yyyyMMdd") & ", "

                        If xProdQty <= xMaxQty Then
                            sql &= " " & CStr(xProdQty) & ", "
                            xProdQty = 0
                        Else
                            sql &= " " & CStr(xMaxQty) & ", "
                            xProdQty = xProdQty - xMaxQty
                        End If
                        sql &= " '" & "EDI" & "', "
                        sql &= " '" & NowDateTime & "' "
                        sql &= " ) "
                        oDataBase.ExecuteNonQuery(sql)

                        xSeqNo = xSeqNo + 1
                    End While
                End If
                ' N+3
                ' ------------------------
                If xOrderZone >= 3 Then
                    ' 準備資料 (OrderMonth, xYYMM, xSeqNo)
                    If CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 3 < 10 Then
                        xOrderMonth = "0" + CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 3)
                        xYYMM = Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2) + xOrderMonth
                    Else
                        If CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 3 > 12 Then
                            xOrderMonth = "0" + CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 3 - 12)
                            xYYMM = CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2)) + 1) + xOrderMonth
                        Else
                            xOrderMonth = CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 3)
                            xYYMM = Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2) + xOrderMonth
                        End If
                    End If
                    xSeqNo = 1
                    ' 製作 EDI 資料
                    xProdQty = dt_LocalStockPlan.Rows(i).Item("PS_03")
                    While xProdQty > 0
                        sql = "Insert into B_LS2EDIInterface "
                        sql &= "( "
                        sql &= "BuyerGroup, BULSNo, BULSSubNo, OrderMonth, Seqno, LSType, KeepCode, Item, ItemName1, ItemName2, ItemName3, Color, PartType, "
                        sql &= "Yobi1, Yobi2, Reference, LengthUnit, BuyerCode, ReqDate, ProdQty, CreateUser, CreateTime"
                        sql &= " ) "
                        sql &= "VALUES( "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("BuyerGroup") & "', "
                        sql &= " '" & xLSNo & "', "
                        sql &= " " & dt_LocalStockPlan.Rows(i).Item("BULSSubNo") & ", "
                        sql &= " '" & xOrderMonth & "', "
                        sql &= " " & CStr(xSeqNo) & ", "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_01") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_02") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_03") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_04") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_05") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_06") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_07") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_08") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_09") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_10") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("LogID") + " N+3 " + xYYMM + " BUY" & "', "
                        sql &= " '" & xLengthUnit & "', "
                        sql &= " '" & Mid(dt_LocalStockPlan.Rows(i).Item("BuyerGroup"), 1, 6) & "', "
                        sql &= " " & Now.AddDays(xReqDate(3)).ToString("yyyyMMdd") & ", "

                        If xProdQty <= xMaxQty Then
                            sql &= " " & CStr(xProdQty) & ", "
                            xProdQty = 0
                        Else
                            sql &= " " & CStr(xMaxQty) & ", "
                            xProdQty = xProdQty - xMaxQty
                        End If
                        sql &= " '" & "EDI" & "', "
                        sql &= " '" & NowDateTime & "' "
                        sql &= " ) "
                        oDataBase.ExecuteNonQuery(sql)

                        xSeqNo = xSeqNo + 1
                    End While
                End If
                ' N+4
                ' ------------------------
                If xOrderZone >= 4 Then
                    ' 準備資料 (OrderMonth, xYYMM, xSeqNo)
                    If CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 4 < 10 Then
                        xOrderMonth = "0" + CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 4)
                        xYYMM = Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2) + xOrderMonth
                    Else
                        If CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 4 > 12 Then
                            xOrderMonth = "0" + CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 4 - 12)
                            xYYMM = CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2)) + 1) + xOrderMonth
                        Else
                            xOrderMonth = CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 4)
                            xYYMM = Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2) + xOrderMonth
                        End If
                    End If
                    xSeqNo = 1
                    ' 製作 EDI 資料
                    xProdQty = dt_LocalStockPlan.Rows(i).Item("PS_04")
                    While xProdQty > 0
                        sql = "Insert into B_LS2EDIInterface "
                        sql &= "( "
                        sql &= "BuyerGroup, BULSNo, BULSSubNo, OrderMonth, Seqno, LSType, KeepCode, Item, ItemName1, ItemName2, ItemName3, Color, PartType, "
                        sql &= "Yobi1, Yobi2, Reference, LengthUnit, BuyerCode, ReqDate, ProdQty, CreateUser, CreateTime"
                        sql &= " ) "
                        sql &= "VALUES( "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("BuyerGroup") & "', "
                        sql &= " '" & xLSNo & "', "
                        sql &= " " & dt_LocalStockPlan.Rows(i).Item("BULSSubNo") & ", "
                        sql &= " '" & xOrderMonth & "', "
                        sql &= " " & CStr(xSeqNo) & ", "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_01") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_02") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_03") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_04") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_05") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_06") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_07") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_08") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_09") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_10") & "', "
                        sql &= " '" & dt_LocalStockPlan.Rows(i).Item("LogID") + " N+4 " + xYYMM + " BUY" & "', "
                        sql &= " '" & xLengthUnit & "', "
                        sql &= " '" & Mid(dt_LocalStockPlan.Rows(i).Item("BuyerGroup"), 1, 6) & "', "
                        sql &= " " & Now.AddDays(xReqDate(4)).ToString("yyyyMMdd") & ", "
                        If xProdQty <= xMaxQty Then
                            sql &= " " & CStr(xProdQty) & ", "
                            xProdQty = 0
                        Else
                            sql &= " " & CStr(xMaxQty) & ", "
                            xProdQty = xProdQty - xMaxQty
                        End If
                        sql &= " '" & "EDI" & "', "
                        sql &= " '" & NowDateTime & "' "
                        sql &= " ) "
                        oDataBase.ExecuteNonQuery(sql)

                        xSeqNo = xSeqNo + 1
                    End While
                End If
                ' 
                ' N+5 (只限VENDOR)
                ' ------------------------
                If pGRBuyer = "VENDORB" Then
                    If xOrderZone >= 5 Then
                        ' 準備資料 (OrderMonth, xYYMM, xSeqNo)
                        If CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 5 < 10 Then
                            xOrderMonth = "0" + CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 5)
                            xYYMM = Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2) + xOrderMonth
                        Else
                            If CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 5 > 12 Then
                                xOrderMonth = "0" + CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 5 - 12)
                                xYYMM = CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2)) + 1) + xOrderMonth
                            Else
                                xOrderMonth = CStr(CInt(Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 5, 2)) + 5)
                                xYYMM = Mid(dt_LocalStockPlan.Rows(i).Item("LogID"), 3, 2) + xOrderMonth
                            End If
                        End If
                        xSeqNo = 1
                        ' 製作 EDI 資料
                        xProdQty = dt_LocalStockPlan.Rows(i).Item("PS_05")
                        While xProdQty > 0
                            sql = "Insert into B_LS2EDIInterface "
                            sql &= "( "
                            sql &= "BuyerGroup, BULSNo, BULSSubNo, OrderMonth, Seqno, LSType, KeepCode, Item, ItemName1, ItemName2, ItemName3, Color, PartType, "
                            sql &= "Yobi1, Yobi2, Reference, LengthUnit, BuyerCode, ReqDate, ProdQty, CreateUser, CreateTime"
                            sql &= " ) "
                            sql &= "VALUES( "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("BuyerGroup") & "', "
                            sql &= " '" & xLSNo & "', "
                            sql &= " " & dt_LocalStockPlan.Rows(i).Item("BULSSubNo") & ", "
                            sql &= " '" & xOrderMonth & "', "
                            sql &= " " & CStr(xSeqNo) & ", "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_01") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_02") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_03") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_04") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_05") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_06") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_07") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_08") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_09") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("GR_10") & "', "
                            sql &= " '" & dt_LocalStockPlan.Rows(i).Item("LogID") + " N+5 " + xYYMM + " BUY" & "', "
                            sql &= " '" & xLengthUnit & "', "
                            sql &= " '" & Mid(dt_LocalStockPlan.Rows(i).Item("BuyerGroup"), 1, 6) & "', "
                            sql &= " " & Now.AddDays(xReqDate(4)).ToString("yyyyMMdd") & ", "
                            If xProdQty <= xMaxQty Then
                                sql &= " " & CStr(xProdQty) & ", "
                                xProdQty = 0
                            Else
                                sql &= " " & CStr(xMaxQty) & ", "
                                xProdQty = xProdQty - xMaxQty
                            End If
                            sql &= " '" & "EDI" & "', "
                            sql &= " '" & NowDateTime & "' "
                            sql &= " ) "
                            oDataBase.ExecuteNonQuery(sql)

                            xSeqNo = xSeqNo + 1
                        End While
                    End If
                End If
            Next
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "EDITransfer", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "EDITransfer", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'EDITransfer-End
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
                    sql = "Update M_FControlRecord Set "
                    sql = sql + " Active = '1', "
                    sql = sql + " Customer = '2', "
                    sql = sql + " DataPrepare = '1', "
                    sql = sql + " DataConvert = '0', "
                    sql = sql + " FCTPlan = '0', "
                    sql = sql + " LSPlan = '0', "
                    sql = sql + " Report = '0', "
                    sql = sql + " IPLSPLan = '0', "
                    sql = sql + " BULSPlan = '0', "
                    sql = sql + " BUReport = '0', "
                    sql = sql + " LFLSPlan = '0', "
                    sql = sql + " EPEDI = '0', "
                    '
                    sql = sql + " BYFMTChange = '0', "
                    sql = sql + " BYImport = '0', "
                    sql = sql + " BYDataCheck = '0', "
                    sql = sql + " BYReport = '0', "
                    sql = sql + " BYConvert = '0', "
                    sql = sql + " BYFASReport = '0', "
                    '
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where Buyer = '" & pBuyer & "' "
                    oDataBase.ExecuteNonQuery(sql)
                Case "StartBY"
                    sql = "Update M_FControlRecord Set "
                    sql = sql + " BYFMTChange = '1', "
                    sql = sql + " BYImport = '1', "
                    sql = sql + " BYDataCheck = '0', "
                    sql = sql + " BYReport = '0', "
                    sql = sql + " BYConvert = '0', "
                    sql = sql + " BYFASReport = '0', "
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where Buyer = '" & pBuyer & "' "
                    oDataBase.ExecuteNonQuery(sql)
                Case "End"
                    sql = "Update M_FControlRecord Set "
                    sql = sql + " Active = '0', "
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where Buyer = '" & pBuyer & "' "
                    oDataBase.ExecuteNonQuery(sql)
                Case "Reset"
                    sql = "Update M_FControlRecord Set "
                    sql = sql + " Active = '0', "
                    sql = sql + " Customer = '0', "
                    sql = sql + " DataPrepare = '0', "
                    sql = sql + " DataConvert = '0', "
                    sql = sql + " FCTPlan = '0', "
                    sql = sql + " LSPlan = '0', "
                    sql = sql + " Report = '0', "
                    sql = sql + " IPLSPLan = '0', "
                    sql = sql + " BULSPlan = '0', "
                    sql = sql + " BUReport = '0', "
                    sql = sql + " LFLSPlan = '0', "
                    sql = sql + " EPEDI = '0', "
                    '
                    sql = sql + " BYFMTChange = '0', "
                    sql = sql + " BYImport = '0', "
                    sql = sql + " BYDataCheck = '0', "
                    sql = sql + " BYReport = '0', "
                    sql = sql + " BYConvert = '0', "
                    sql = sql + " BYFASReport = '0', "
                    '
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where Buyer = '" & pBuyer & "' "
                    oDataBase.ExecuteNonQuery(sql)
                Case "ResetBY"
                    sql = "Update M_FControlRecord Set "
                    sql = sql + " BYFMTChange = '1', "
                    sql = sql + " BYImport = '1', "
                    sql = sql + " BYDataCheck = '0', "
                    sql = sql + " BYReport = '0', "
                    sql = sql + " BYConvert = '0', "
                    sql = sql + " BYFASReport = '0', "
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where Buyer = '" & pBuyer & "' "
                    oDataBase.ExecuteNonQuery(sql)
                Case "Action2"
                    sql = "Update M_FControlRecord Set "
                    sql = sql + pAction2 + " = '" & CStr(pStatus2) & "', "
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where Buyer = '" & pBuyer & "' "
                    oDataBase.ExecuteNonQuery(sql)
                Case Else
                    sql = "Update M_FControlRecord Set "
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
            sql = "Select * From M_FControlRecord "
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
    '**([540]-DeleteFCTData)
    '**     刪除 FCT Plan 資料
    '***********************************************************************************************
    'DeleteFCTData-Start
    <WebMethod()> _
    Public Function DeleteFCTData(ByVal pBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Delete From ForcastPlan "
            sql = sql + "Where Buyer = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'DeleteFCTData-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([541]-DeleteFCTLevelData)
    '**     刪除 FCT Plan 資料(對象Level<>0)
    '***********************************************************************************************
    'DeleteFCTLevelData-Start
    <WebMethod()> _
    Public Function DeleteFCTLevelData(ByVal pBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Delete From ForcastPlan "
            sql = sql + "Where Buyer = '" & pBuyer & "' "
            sql = sql + "  And Y_Level <> " & "0" & " "
            oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'DeleteFCTLevelData-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([550]-DeleteLSData)
    '**     刪除 LS Plan 資料
    '***********************************************************************************************
    'DeleteLSData-Start
    <WebMethod()> _
    Public Function DeleteLSData(ByVal pBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Delete From LocalStockPlan "
            sql = sql + "Where Buyer = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
            '
            ' Delete LocalStockPlan Prod-Inf Data
            If DeleteLSProdInfData(pBuyer) <> 0 Then
                RtnCode = 1
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'DeleteLSData-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([551]-DeleteBuyerLSData)
    '**     刪除 Buyer LS Plan 資料
    '***********************************************************************************************
    'DeleteBuyerLSData-Start
    <WebMethod()> _
    Public Function DeleteBuyerLSData(ByVal pBuyer As String, ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Delete From BuyerLocalStockPlan "
            sql = sql + "Where Buyer = '" & pBuyer & "' "
            sql = sql + "  And BuyerGroup = '" & pGRBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'DeleteBuyerLSData-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([552]-DeleteLSProdInfData)
    '**     刪除 LS Plan Prod Inf 資料
    '***********************************************************************************************
    'DeleteLSProdInfData-Start
    <WebMethod()> _
    Public Function DeleteLSProdInfData(ByVal pBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Delete From LOCALSTOCKPlan_ProdInf "
            sql = sql + "Where Buyer = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'DeleteLSProdInfData-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([555]-DeleteLS2EDIInterface	    						           
    '**     刪除 EDI Interface 資料
    '***********************************************************************************************
    'DeleteLS2EDIInterface-Start
    <WebMethod()> _
    Public Function DeleteLS2EDIInterface(ByVal pBuyer As String, ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Delete From B_LS2EDIInterface "
            sql = sql + "Where BuyerGroup = '" & pGRBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'DeleteLS2EDIInterface-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([560]-GetLSSeqNo)
    '**     取得LS SeqNo
    '***********************************************************************************************
    'GetLSSeqNo-Start
    <WebMethod()> _
    Public Function GetLSSeqNo(ByVal pBuyer As String, ByVal pLSNo As String) As Integer
        Dim RtnCode As Integer = 1
        '
        Dim sql As String
        Try
            sql = "Select LSNo From M_FControlRecord "
            sql = sql + "Where Buyer = '" & pBuyer & "' "
            sql = sql + "  And LSNo Like '" & pLSNo & "%' "
            Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
            If dt_ControlRecord.Rows.Count > 0 Then
                RtnCode = CInt(Mid(dt_ControlRecord.Rows(0).Item("LSNo"), 7, 4))
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'GetLSSeqNo-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([561]-GetBuyerLSSeqNo)
    '**     取得Buyer LS SeqNo
    '***********************************************************************************************
    'GetBuyerLSSeqNo-Start
    <WebMethod()> _
    Public Function GetBuyerLSSeqNo(ByVal pBuyer As String, ByVal pBULSNo As String) As Integer
        Dim RtnCode As Integer = 1
        '
        Dim sql As String
        Try
            sql = "Select BULSNo From M_FControlRecord "
            sql = sql + "Where Buyer = '" & pBuyer & "' "
            sql = sql + "  And BULSNo Like '" & pBULSNo & "%' "
            Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
            If dt_ControlRecord.Rows.Count > 0 Then
                RtnCode = CInt(Mid(dt_ControlRecord.Rows(0).Item("BULSNo"), 7, 4))
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'GetBuyerLSSeqNo-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([570]-GetFCTSeqNo)
    '**     取得FCT SeqNo
    '***********************************************************************************************
    'GetFCTSeqNo-Start
    <WebMethod()> _
    Public Function GetFCTSeqNo(ByVal pBuyer As String, ByVal pFCTNo As String, ByVal pFCTType As String) As Integer
        Dim RtnCode As Integer = 1
        '
        Dim sql As String
        Try
            If pFCTType = "C" Then
                sql = "Select FCTNo From M_FControlRecord "
                sql = sql + "Where Buyer = '" & pBuyer & "' "
                sql = sql + "  And FCTNo Like '" & pFCTNo & "%' "
                Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
                If dt_ControlRecord.Rows.Count > 0 Then
                    RtnCode = CInt(Mid(dt_ControlRecord.Rows(0).Item("FCTNo"), 7, 5))
                End If
            Else
                sql = "Select YFCTNo From M_FControlRecord "
                sql = sql + "Where Buyer = '" & pBuyer & "' "
                sql = sql + "  And YFCTNo Like '" & pFCTNo & "%' "
                Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
                If dt_ControlRecord.Rows.Count > 0 Then
                    RtnCode = CInt(Mid(dt_ControlRecord.Rows(0).Item("YFCTNo"), 8, 3))
                End If
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'GetFCTSeqNo-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([575]-GetLastPlanVersion)
    '**     取得最後Final Plan Version
    '***********************************************************************************************
    'GetLastPlanVersion-Start
    <WebMethod()> _
    Public Function GetLastPlanVersion(ByVal pGRBuyer As String) As String
        Dim RtnCode As String = ""
        '
        Dim sql As String
        Try
            sql = "SELECT Top 1 LogID FROM LF_BuyerLocalStockPlan "
            sql &= "Where BuyerGroup = '" & pGRBuyer & "' "
            sql &= "  And Version = " & "99" & " "
            sql &= "Order by LogID Desc "
            Dim dt_PlanVersion As DataTable = oDataBase.GetDataTable(sql)
            If dt_PlanVersion.Rows.Count > 0 Then
                RtnCode = dt_PlanVersion.Rows(0).Item("LogID")
            End If
        Catch ex As Exception
            RtnCode = ""
        End Try
        '
        Return RtnCode
    End Function
    'GetLastPlanVersion-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([580]-UpdateNextFCTNo)
    '**     更新下一次可使用FCTNo
    '***********************************************************************************************
    'UpdateNextFCTNo-Start
    <WebMethod()> _
    Public Function UpdateNextFCTNo(ByVal pBuyer As String, ByVal pFCTNo As String, ByVal pFCTType As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            If pFCTType = "C" Then
                sql = "Update M_FControlRecord Set "
                sql = sql + "FCTNo = '" & pFCTNo & "' "
                sql = sql + "Where Buyer = '" & pBuyer & "' "
                oDataBase.ExecuteNonQuery(sql)
            Else
                sql = "Update M_FControlRecord Set "
                sql = sql + "YFCTNo = '" & pFCTNo & "' "
                sql = sql + "Where Buyer = '" & pBuyer & "' "
                oDataBase.ExecuteNonQuery(sql)
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'UpdateNextFCTNo-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([590]-UpdateNextLSNo)
    '**     更新下一次可使用LSNo
    '***********************************************************************************************
    'UpdateNextLSNo-Start
    <WebMethod()> _
    Public Function UpdateNextLSNo(ByVal pBuyer As String, ByVal pLSNo As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Update M_FControlRecord Set "
            sql = sql + "LSNo = '" & pLSNo & "' "
            sql = sql + "Where Buyer = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([591]-UpdateNextBuyerLSNo)
    '**     更新下一次可使用BuyerLSNo
    '***********************************************************************************************
    'UpdateNextBuyerLSNo-Start
    <WebMethod()> _
    Public Function UpdateNextBuyerLSNo(ByVal pBuyer As String, ByVal pBULSNo As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Update M_FControlRecord Set "
            sql = sql + "BULSNo = '" & pBULSNo & "' "
            sql = sql + "Where Buyer = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'UpdateNextBuyerLSNo-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([600]-WriteFCTPlan)
    '**     Write Forcast Plan
    '***********************************************************************************************
    'WriteFCTPlan-Start
    <WebMethod()> _
    Public Function WriteFCTPlan(ByVal pID As Integer, _
                                 ByVal pSubNo As Integer, _
                                 ByVal pLevel As Integer, _
                                 ByVal pFatherItem As String, _
                                 ByVal pItem As String, _
                                 ByVal pColor As String, _
                                 ByVal pClass As String, _
                                 ByVal pQty As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql, str As String
        Dim Qty, QtyTotal As Double
        '
        Try
            oWaves.Timeout = Timeout.Infinite

            sql = "SELECT * FROM ForcastPlan "
            sql &= "Where Unique_ID = '" & CStr(pID) & "' "
            Dim dt_FCTData As DataTable = oDataBase.GetDataTable(sql)
            If dt_FCTData.Rows.Count > 0 Then
                '
                sql = "Insert into ForcastPlan "
                sql &= "( "
                sql &= "BuyerGroup, "
                sql &= "FCTNo, "
                sql &= "FCTSubNo, "
                sql &= "Version, "
                sql &= "C_Code, "
                sql &= "C_Color, "
                sql &= "C_SpecialRequest, "
                sql &= "C_Season, "
                sql &= "C_ShortenLT, "
                sql &= "C_Style, "
                sql &= "C_A1, "
                sql &= "C_B1, "
                sql &= "C_C1, "
                sql &= "C_D1, "
                sql &= "C_E1, "
                sql &= "C_F1, "
                sql &= "C_G1, "
                sql &= "C_H1, "
                sql &= "C_I1, "
                sql &= "C_J1, "
                sql &= "C_K1, "
                sql &= "C_L1, "
                sql &= "C_M1, "
                sql &= "C_N1, "
                sql &= "C_O1, "

                sql &= "Y_LEVEL, "
                sql &= "Y_ItemCode, "
                sql &= "Y_ItemName1, "
                sql &= "Y_ItemName2, "
                sql &= "Y_ItemName3, "
                sql &= "Y_Color, "
                sql &= "Y_A1, "
                sql &= "Y_B1, "
                sql &= "Y_C1, "
                sql &= "Y_D1, "
                sql &= "Y_E1, "
                sql &= "Y_F1, "
                sql &= "Y_G1, "
                sql &= "Y_H1, "
                sql &= "Y_I1, "
                sql &= "Y_J1, "

                sql &= "N_F, "
                sql &= "N1_F, "
                sql &= "N2_F, "
                sql &= "N3_F, "
                sql &= "N4_F, "
                sql &= "N5_F, "
                sql &= "N6_F, "
                sql &= "N7_F, "
                sql &= "N8_F, "
                sql &= "N9_F, "
                sql &= "N10_F, "
                sql &= "N11_F, "
                sql &= "N12_F, "
                sql &= "Total, "

                sql &= "BUYER, "
                sql &= "ID, "
                sql &= "CreateUser, "
                sql &= "CreateTime "
                sql &= " ) "
                sql &= "VALUES( "
                sql &= " '" & dt_FCTData.Rows(0).Item("BuyerGroup").ToString & "', "                ' BuyerGroup
                sql &= " '" & dt_FCTData.Rows(0).Item("FCTNo").ToString & "', "                     ' FCTNo
                sql &= " " & CStr(pSubNo) & ", "                                                    ' FCTSubNo
                sql &= " " & dt_FCTData.Rows(0).Item("Version").ToString & ", "                     ' Version
                sql &= " '" & dt_FCTData.Rows(0).Item("C_Code").ToString & "', "                    ' C_Code
                sql &= " '" & dt_FCTData.Rows(0).Item("C_Color").ToString & "', "                   ' C_Color
                sql &= " '" & dt_FCTData.Rows(0).Item("C_SpecialRequest").ToString & "', "          ' C_SpecialRequest
                sql &= " '" & dt_FCTData.Rows(0).Item("C_Season").ToString & "', "                  ' C_Season
                sql &= " '" & dt_FCTData.Rows(0).Item("C_ShortenLT").ToString & "', "               ' C_ShortenLT
                sql &= " '" & dt_FCTData.Rows(0).Item("C_Style").ToString & "', "                   ' C_Style
                sql &= " '" & dt_FCTData.Rows(0).Item("C_A1").ToString & "', "                      ' C_A1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_B1").ToString & "', "                      ' C_B1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_C1").ToString & "', "                      ' C_C1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_D1").ToString & "', "                      ' C_D1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_E1").ToString & "', "                      ' C_E1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_F1").ToString & "', "                      ' C_F1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_G1").ToString & "', "                      ' C_G1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_H1").ToString & "', "                      ' C_H1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_I1").ToString & "', "                      ' C_I1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_J1").ToString & "', "                      ' C_J1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_K1").ToString & "', "                      ' C_K1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_L1").ToString & "', "                      ' C_L1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_M1").ToString & "', "                      ' C_M1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_N1").ToString & "', "                      ' C_N1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_O1").ToString & "', "                      ' C_O1

                sql &= " " & CStr(pLevel) & ", "                                                    ' Y_LEVEL
                sql &= " '" & pItem & "', "                                                         ' Y_ItemCode

                oWaves.GetItemName001(pItem, 1, str)                                                ' Y_ItemName1
                sql &= " '" & str & "', "
                oWaves.GetItemName001(pItem, 2, str)                                                ' Y_ItemName2
                sql &= " '" & str & "', "
                oWaves.GetItemName001(pItem, 3, str)                                                ' Y_ItemName3
                sql &= " '" & str & "', "

                oWaves.GetChangeColor("01", pFatherItem, pItem, pColor, str)                        ' Y_Color(考慮兼用色)
                sql &= " '" & str & "', "

                sql &= " '" & pClass & "', "                                                        ' Y_A1 (固定使用 Item Class)

                sql &= " '" & dt_FCTData.Rows(0).Item("Y_B1").ToString & "', "                      ' Y_B1
                sql &= " '" & dt_FCTData.Rows(0).Item("Y_C1").ToString & "', "                      ' Y_C1
                sql &= " '" & dt_FCTData.Rows(0).Item("Y_D1").ToString & "', "                      ' Y_D1
                sql &= " '" & dt_FCTData.Rows(0).Item("Y_E1").ToString & "', "                      ' Y_E1
                sql &= " '" & dt_FCTData.Rows(0).Item("Y_F1").ToString & "', "                      ' Y_F1
                sql &= " '" & dt_FCTData.Rows(0).Item("Y_G1").ToString & "', "                      ' Y_G1
                sql &= " '" & dt_FCTData.Rows(0).Item("Y_H1").ToString & "', "                      ' Y_H1
                sql &= " '" & dt_FCTData.Rows(0).Item("Y_I1").ToString & "', "                      ' Y_I1
                sql &= " '" & dt_FCTData.Rows(0).Item("Y_J1").ToString & "', "                      ' Y_J1

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N_F"))                             ' N_F
                QtyTotal = Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N1_F"))                            ' N1_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N2_F"))                            ' N2_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N3_F"))                            ' N3_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N4_F"))                            ' N4_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N5_F"))                            ' N5_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N6_F"))                            ' N6_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N7_F"))                            ' N7_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N8_F"))                            ' N8_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N9_F"))                            ' N9_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N10_F"))                           ' N10_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N11_F"))                           ' N11_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N12_F"))                           ' N12_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                sql &= " " & CStr(QtyTotal) & ", "                                                  ' Total

                sql &= " '" & dt_FCTData.Rows(0).Item("BUYER").ToString & "', "                     ' BUYER
                sql &= " " & dt_FCTData.Rows(0).Item("ID").ToString & ", "                          ' ID
                sql &= " '" & "FCTPlan" & "', "
                sql &= " '" & NowDateTime & "' "
                sql &= " ) "
                oDataBase.ExecuteNonQuery(sql)
                '
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'WriteFCTPlan-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([600]-WriteNewFCTPlan)
    '**     Write Forcast Plan
    '***********************************************************************************************
    'WriteFCTPlan-Start
    <WebMethod()> _
    Public Function WriteNewFCTPlan(ByVal pBuyer As String, _
                                   ByVal pID As Integer, _
                                   ByVal pSubNo As Integer, _
                                   ByVal pLevel As Integer, _
                                   ByVal pFatherItem As String, _
                                   ByVal pItem As String, _
                                   ByVal pColor As String, _
                                   ByVal pClass As String, _
                                   ByVal pQty As String, _
                                   ByVal pRuleNo As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql, xsql, str As String
        Dim Qty, QtyTotal As Double
        '
        Try
            oWaves.Timeout = Timeout.Infinite

            sql = "SELECT * FROM ForcastPlan "
            sql &= "Where Unique_ID = '" & CStr(pID) & "' "
            Dim dt_FCTData As DataTable = oDataBase.GetDataTable(sql)
            If dt_FCTData.Rows.Count > 0 Then
                '
                sql = "Insert into ForcastPlan "
                sql &= "( "
                sql &= "BuyerGroup, "
                sql &= "FCTNo, "
                sql &= "FCTSubNo, "
                sql &= "Version, "
                sql &= "C_Code, "
                sql &= "C_Color, "
                sql &= "C_SpecialRequest, "
                sql &= "C_Season, "
                sql &= "C_ShortenLT, "
                sql &= "C_Style, "
                sql &= "C_A1, "
                sql &= "C_B1, "
                sql &= "C_C1, "
                sql &= "C_D1, "
                sql &= "C_E1, "
                sql &= "C_F1, "
                sql &= "C_G1, "
                sql &= "C_H1, "
                sql &= "C_I1, "
                sql &= "C_J1, "
                sql &= "C_K1, "
                sql &= "C_L1, "
                sql &= "C_M1, "
                sql &= "C_N1, "
                sql &= "C_O1, "

                sql &= "Y_LEVEL, "
                sql &= "Y_ItemCode, "
                sql &= "Y_ItemName1, "
                sql &= "Y_ItemName2, "
                sql &= "Y_ItemName3, "
                sql &= "Y_Color, "
                sql &= "Y_A1, "
                sql &= "Y_B1, "
                sql &= "Y_C1, "
                sql &= "Y_D1, "
                sql &= "Y_E1, "
                sql &= "Y_F1, "
                sql &= "Y_G1, "
                sql &= "Y_H1, "
                sql &= "Y_I1, "
                sql &= "Y_J1, "

                sql &= "N_F, "
                sql &= "N1_F, "
                sql &= "N2_F, "
                sql &= "N3_F, "
                sql &= "N4_F, "
                sql &= "N5_F, "
                sql &= "N6_F, "
                sql &= "N7_F, "
                sql &= "N8_F, "
                sql &= "N9_F, "
                sql &= "N10_F, "
                sql &= "N11_F, "
                sql &= "N12_F, "
                sql &= "Total, "

                sql &= "BUYER, "
                sql &= "ID, "
                sql &= "CreateUser, "
                sql &= "CreateTime "
                sql &= " ) "
                '
                ' 取得 備料MASTER-INF
                ' ------------------------------------------------------
                Dim xLTType, xPartType, xRuleCode, xRuleNo, xRuleSubno, xKeep As String
                Dim xRatio1, xRatio2, xRatio3, xRatio4 As Integer
                ' 決定LT種類(LTType)
                xLTType = "LLT"
                If pBuyer = "FALL-000001" Or pBuyer = "FALL-000016" Then
                    '
                    'MODIFY-START ADIDAS SPEED 240627
                    'If dt_FCTData.Rows(0).Item("C_ShortenLT").ToString <> "AD-RB" Then
                    'xLTType = "SLT"
                    'End If
                    If dt_FCTData.Rows(0).Item("C_ShortenLT").ToString <> "TP4ADCLA-1" Then
                        xLTType = "SLT"
                    End If
                    'MODIFY-END
                    '
                End If
                ' 決定Part Type
                xPartType = "CH"
                If pClass = "PS" Then xPartType = "SLD" ' PARTTYPE      
                If pClass = "ZIP" Then xPartType = "ZIP" ' PARTTYPE      
                ' 初值
                xRatio1 = 100
                xRatio2 = 100
                xRatio3 = 100
                xRatio4 = 100
                xKeep = ""
                '
                str = pRuleNo   ' RULE-CODE, RULE-N0, RULESUBNO
                If InStr(str, "/") > 0 Then str = Mid(str, InStr(str, "/") + 1)
                xRuleCode = Mid(str, 1, InStr(str, "-") - 1)
                str = Mid(str, InStr(str, "-") + 1)
                xRuleNo = Mid(str, 1, InStr(str, "-") - 1)
                str = Mid(str, InStr(str, "-") + 1)
                xRuleSubno = str
                '
                xsql = "SELECT Action, Ratio1, Ratio2, Ratio3, Ratio4, Keep, "
                xsql &= "str(Ratio1,3,0) + '/' + str(Ratio2,3,0) + '/' + str(Ratio3,3,0) + '/' + str(Ratio4,3,0) + '/' As Ratio, "
                xsql &= "SearchItem1 + '/' + SearchItem2 + '/' + SearchItem3 + '/' +SearchItem4 + '/' + SearchItem5 + '/' + "
                xsql &= "SearchItem6 + '/' + SearchItem7 + '/' + SearchItem8 + '/' +SearchItem9 + '/' + SearchItem10 + '/' As SearchItem "
                xsql &= "FROM M_LSOrderRule "
                xsql &= "Where Active = '1' "
                xsql &= "  And Buyer = '" & pBuyer & "' "
                xsql &= "  And LTType = '" & xLTType & "' "
                xsql &= "  And PartType = '" & xPartType & "' "
                xsql &= "  And RuleCode = '" & xRuleCode & "' "
                xsql &= "  And RuleNo = '" & xRuleNo & "' "
                xsql &= "  And RuleSubno = '" & xRuleSubno & "' "
                xsql &= "Order by RuleCode, RuleNo, RuleSubno "
                Dim dt_LSOrderRule As DataTable = oDataBase.GetDataTable(xsql)
                If dt_LSOrderRule.Rows.Count > 0 Then
                    xRatio1 = dt_LSOrderRule.Rows(0).Item("Ratio1")
                    xRatio2 = dt_LSOrderRule.Rows(0).Item("Ratio2")
                    xRatio3 = dt_LSOrderRule.Rows(0).Item("Ratio3")
                    xRatio4 = dt_LSOrderRule.Rows(0).Item("Ratio4")
                    xKeep = dt_LSOrderRule.Rows(0).Item("Keep").ToString
                End If
                '
                ' 製作 FORECAST
                ' ---------------------------------------------------------
                sql &= "VALUES( "
                sql &= " '" & dt_FCTData.Rows(0).Item("BuyerGroup").ToString & "', "                ' BuyerGroup
                sql &= " '" & dt_FCTData.Rows(0).Item("FCTNo").ToString & "', "                     ' FCTNo
                sql &= " " & CStr(pSubNo) & ", "                                                    ' FCTSubNo
                sql &= " " & dt_FCTData.Rows(0).Item("Version").ToString & ", "                     ' Version
                sql &= " '" & dt_FCTData.Rows(0).Item("C_Code").ToString & "', "                    ' C_Code
                sql &= " '" & dt_FCTData.Rows(0).Item("C_Color").ToString & "', "                   ' C_Color
                sql &= " '" & dt_FCTData.Rows(0).Item("C_SpecialRequest").ToString & "', "          ' C_SpecialRequest
                sql &= " '" & dt_FCTData.Rows(0).Item("C_Season").ToString & "', "                  ' C_Season

                If xKeep = "Y" Then                                                                 ' C_ShortenLT    
                    sql &= " '" & dt_FCTData.Rows(0).Item("C_ShortenLT").ToString & "', "
                Else
                    sql &= " '" & "" & "', "
                End If
                sql &= " '" & dt_FCTData.Rows(0).Item("C_Style").ToString & "', "                   ' C_Style
                sql &= " '" & dt_FCTData.Rows(0).Item("C_A1").ToString & "', "                      ' C_A1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_B1").ToString & "', "                      ' C_B1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_C1").ToString & "', "                      ' C_C1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_D1").ToString & "', "                      ' C_D1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_E1").ToString & "', "                      ' C_E1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_F1").ToString & "', "                      ' C_F1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_G1").ToString & "', "                      ' C_G1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_H1").ToString & "', "                      ' C_H1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_I1").ToString & "', "                      ' C_I1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_J1").ToString & "', "                      ' C_J1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_K1").ToString & "', "                      ' C_K1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_L1").ToString & "', "                      ' C_L1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_M1").ToString & "', "                      ' C_M1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_N1").ToString & "', "                      ' C_N1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_O1").ToString & "', "                      ' C_O1
                ' ------------------------------------------------------------------
                sql &= " " & CStr(pLevel) & ", "                                                    ' Y_LEVEL
                sql &= " '" & pItem & "', "                                                         ' Y_ItemCode
                oWaves.GetItemName001(pItem, 1, str)                                                ' Y_ItemName1
                sql &= " '" & str & "', "
                oWaves.GetItemName001(pItem, 2, str)                                                ' Y_ItemName2
                sql &= " '" & str & "', "
                oWaves.GetItemName001(pItem, 3, str)                                                ' Y_ItemName3
                sql &= " '" & str & "', "

                If xPartType = "ZIP" Then                                                           ' Y_Color(考慮兼用色)(ZIP時不考慮)
                    str = pColor
                Else
                    oWaves.GetChangeColor("01", pFatherItem, pItem, pColor, str)
                End If
                sql &= " '" & str & "', "

                sql &= " '" & pClass & "', "                                                        ' Y_A1 (Item Class)
                If InStr(pRuleNo, "/") > 0 Then                                                     ' 異常
                    sql &= " '" & Mid(pRuleNo, 1, InStr(pRuleNo, "/") - 1) & "', "                  ' Y_B1 (Rule Inf-原LS-RuleNo)
                    sql &= " '" & Mid(pRuleNo, InStr(pRuleNo, "/") + 1) & "', "                     ' Y_C1 (Rule Inf-採用LS-RulenNo)
                Else                                                                                ' 正常
                    sql &= " '" & pRuleNo & "', "
                    sql &= " '" & pRuleNo & "', "
                End If

                If dt_LSOrderRule.Rows.Count > 0 Then
                    sql &= " '" & dt_LSOrderRule.Rows(0).Item("Action").ToString & "', "            ' Y_D1 (Rule Inf-Action)
                    sql &= " '" & dt_LSOrderRule.Rows(0).Item("SearchItem").ToString & "', "        ' Y_E1 (SearchItem)
                    sql &= " '" & dt_LSOrderRule.Rows(0).Item("Keep").ToString & "', "              ' Y_F1 (Keep)
                    sql &= " '" & dt_LSOrderRule.Rows(0).Item("Ratio").ToString & "', "             ' Y_G1 (N+1 ~ N+4)
                Else
                    sql &= " '" & dt_FCTData.Rows(0).Item("Y_D1").ToString & "', "                  ' Y_D1
                    sql &= " '" & dt_FCTData.Rows(0).Item("Y_E1").ToString & "', "                  ' Y_E1
                    sql &= " '" & dt_FCTData.Rows(0).Item("Y_F1").ToString & "', "                  ' Y_F1
                    sql &= " '" & dt_FCTData.Rows(0).Item("Y_G1").ToString & "', "                  ' Y_G1
                End If
                sql &= " '" & dt_FCTData.Rows(0).Item("Y_H1").ToString & "', "                      ' Y_H1
                sql &= " '" & dt_FCTData.Rows(0).Item("Y_I1").ToString & "', "                      ' Y_I1
                sql &= " '" & dt_FCTData.Rows(0).Item("Y_J1").ToString & "', "                      ' Y_J1

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N_F"))                             ' N_F
                QtyTotal = Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N1_F"))                            ' N1_F
                If pClass = "PS" Then
                    Qty = Fix(Qty * (xRatio1 / 100) + 0.9)
                Else
                    Qty = Qty * (xRatio1 / 100)
                End If
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N2_F"))                            ' N2_F
                If pClass = "PS" Then
                    Qty = Fix(Qty * (xRatio2 / 100) + 0.9)
                Else
                    Qty = Qty * (xRatio2 / 100)
                End If
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N3_F"))                            ' N3_F
                If pClass = "PS" Then
                    Qty = Fix(Qty * (xRatio3 / 100) + 0.9)
                Else
                    Qty = Qty * (xRatio3 / 100)
                End If
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N4_F"))                            ' N4_F
                If pClass = "PS" Then
                    Qty = Fix(Qty * (xRatio4 / 100) + 0.9)
                Else
                    Qty = Qty * (xRatio4 / 100)
                End If
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N5_F"))                            ' N5_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N6_F"))                            ' N6_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N7_F"))                            ' N7_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N8_F"))                            ' N8_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N9_F"))                            ' N9_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N10_F"))                           ' N10_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N11_F"))                           ' N11_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N12_F"))                           ' N12_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                sql &= " " & CStr(QtyTotal) & ", "                                                  ' Total

                sql &= " '" & dt_FCTData.Rows(0).Item("BUYER").ToString & "', "                     ' BUYER
                sql &= " " & dt_FCTData.Rows(0).Item("ID").ToString & ", "                          ' ID
                sql &= " '" & "FCTPlan" & "', "
                sql &= " '" & NowDateTime & "' "
                sql &= " ) "
                oDataBase.ExecuteNonQuery(sql)
                '
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'WriteFCTPlan-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([610]-CanRunLFLocalStockPlan)
    '**     判斷是否可執行 LS Plan 最終確定
    '***********************************************************************************************
    'CanRunLFLocalStockPlan-Start
    <WebMethod()> _
    Public Function CanRunLFLocalStockPlan(ByVal pBuyer As String, ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql, xLFBULSNo, xBULSNo As String
        Try
            xLFBULSNo = ""
            xBULSNo = ""
            ' 取得 LF-BuyerLocalStockPlan
            sql = "SELECT Top 1 BULSNo FROM LF_BuyerLocalStockPlan "
            sql &= "Where Buyer = '" & "BULS" & "' "
            sql &= "  And BuyerGroup = '" & pGRBuyer & "' "
            sql &= "Order by BULSNo Desc "
            Dim dt_LFBULSPlan As DataTable = oDataBase.GetDataTable(sql)
            If dt_LFBULSPlan.Rows.Count > 0 Then
                xLFBULSNo = dt_LFBULSPlan.Rows(0).Item("BULSNo")
            End If
            ' 取得 BuyerLocalStockPlan
            sql = "SELECT Top 1 BULSNo FROM BuyerLocalStockPlan "
            sql &= "Where Buyer = '" & "BULS" & "' "
            sql &= "  And BuyerGroup = '" & pGRBuyer & "' "
            sql &= "Order by BULSNo Desc "
            Dim dt_BULSPlan As DataTable = oDataBase.GetDataTable(sql)
            If dt_BULSPlan.Rows.Count > 0 Then
                xBULSNo = dt_BULSPlan.Rows(0).Item("BULSNo")
            End If
            ' 
            If xBULSNo <= xLFBULSNo Then RtnCode = 1
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'CanRunLFLocalStockPlan-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([620]-GetLastUniqueID)
    '**     取得最後1筆資料
    '***********************************************************************************************
    'GetLastUniqueID-Start
    <WebMethod()> _
    Public Function GetLastUniqueID(ByVal pBuyer As String, ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            ' 取得 Forcast Plan 最後1筆資料
            sql = "SELECT Top 1 Unique_ID FROM ForcastPlan "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "Order by Unique_ID Desc "
            Dim dt_FCTPlan As DataTable = oDataBase.GetDataTable(sql)
            If dt_FCTPlan.Rows.Count > 0 Then
                RtnCode = dt_FCTPlan.Rows(0).Item("Unique_ID")
            End If
        Catch ex As Exception
            RtnCode = 0
        End Try
        '
        Return RtnCode
    End Function
    'GetLastUniqueID-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([630]-ResetFCTNo)
    '**     重新設定下一次可使用FCTNo
    '***********************************************************************************************
    'ResetFCTNo-Start
    <WebMethod()> _
    Public Function ResetFCTNo(ByVal pBuyer As String, ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Dim xSeqno As Integer
        Dim xSeqNoString, xGroup, LF_LastNo, CTL_LastNo As String
        Try
            ' -- FCT NO ---------------------------------------------------------------
            ' 取得 ControlRecord Last FCTNo
            CTL_LastNo = ""
            xGroup = ""
            sql = "SELECT GroupCode, FCTNo FROM M_FControlRecord "
            sql &= "Where Buyer = '" & pBuyer & "' "
            Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
            If dt_ControlRecord.Rows.Count > 0 Then
                CTL_LastNo = dt_ControlRecord.Rows(0).Item("FCTNo")
                xGroup = dt_ControlRecord.Rows(0).Item("GroupCode")
            End If
            ' 取得 LastFinal Last FCTNo
            LF_LastNo = xGroup + "F" + "13100001"
            sql = "SELECT Top 1 FCTNo FROM LF_ForcastPlan "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And FCTNo Like '" & xGroup + "F" & "%' "
            sql &= "Order by FCTNo Desc "
            Dim dt_LFPlan As DataTable = oDataBase.GetDataTable(sql)
            If dt_LFPlan.Rows.Count > 0 Then
                xSeqno = CInt(Mid(dt_LFPlan.Rows(0).Item("FCTNo"), 7, 5)) + 1
                xSeqNoString = CStr(xSeqno)
                If Len(xSeqNoString) < 2 Then
                    xSeqNoString = "0000" + CStr(xSeqno)
                Else
                    If Len(xSeqNoString) < 3 Then
                        xSeqNoString = "000" + CStr(xSeqno)
                    Else
                        If Len(xSeqNoString) < 4 Then
                            xSeqNoString = "00" + CStr(xSeqno)
                        Else
                            If Len(xSeqNoString) < 5 Then
                                xSeqNoString = "0" + CStr(xSeqno)
                            End If
                        End If
                    End If
                End If
                LF_LastNo = Mid(dt_LFPlan.Rows(0).Item("FCTNo"), 1, 6) + xSeqNoString
            End If
            ' 重新設定下一次可使用FCTNo
            If CTL_LastNo <> LF_LastNo Then
                UpdateNextFCTNo(pBuyer, LF_LastNo, "C")
            End If
            '
            ' -- YKK FCT NO ---------------------------------------------------------------
            ' 取得 ControlRecord Last FCTNo
            CTL_LastNo = ""
            xGroup = ""
            sql = "SELECT GroupCode, YFCTNo FROM M_FControlRecord "
            sql &= "Where Buyer = '" & pBuyer & "' "
            Dim dt_ControlRecord1 As DataTable = oDataBase.GetDataTable(sql)
            If dt_ControlRecord1.Rows.Count > 0 Then
                CTL_LastNo = dt_ControlRecord1.Rows(0).Item("YFCTNo")
                xGroup = dt_ControlRecord1.Rows(0).Item("GroupCode")
            End If
            ' 取得 LastFinal Last FCTNo
            LF_LastNo = xGroup + "YF" + "131001"
            sql = "SELECT Top 1 FCTNo FROM LF_ForcastPlan "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And FCTNo Like '" & xGroup + "YF" & "%' "
            sql &= "Order by FCTNo Desc "
            Dim dt_LFPlan1 As DataTable = oDataBase.GetDataTable(sql)
            If dt_LFPlan1.Rows.Count > 0 Then
                xSeqno = CInt(Mid(dt_LFPlan1.Rows(0).Item("FCTNo"), 8, 3)) + 1
                xSeqNoString = CStr(xSeqno)
                If Len(xSeqNoString) < 2 Then
                    xSeqNoString = "00" + CStr(xSeqno)
                Else
                    If Len(xSeqNoString) < 3 Then
                        xSeqNoString = "0" + CStr(xSeqno)
                    End If
                End If
                LF_LastNo = Mid(dt_LFPlan1.Rows(0).Item("FCTNo"), 1, 7) + xSeqNoString
            End If
            ' 重新設定下一次可使用FCTNo
            If CTL_LastNo <> LF_LastNo Then
                UpdateNextFCTNo(pBuyer, LF_LastNo, "Y")
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'ResetFCTNo-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([631]-ResetPlanFCTNo)
    '**     重新設定 FCT Plan FCTNo
    '***********************************************************************************************
    'ResetPlanFCTNo-Start
    <WebMethod()> _
    Public Function ResetPlanFCTNo(ByVal pBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Update ForcastPlan Set "
            sql = sql + "FCTNo = '" & "" & "', "
            sql = sql + "FCTSubNo = " & "1" & ", "
            sql = sql + "LSNo = '" & "" & "', "
            sql = sql + "LSSubNo = " & "0" & " "
            sql = sql + "Where Buyer = '" & pBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'ResetPlanFCTNo-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([640]-ResetLSNo)
    '**     重新設定下一次可使用LSNo
    '***********************************************************************************************
    'ResetLSNo-Start
    <WebMethod()> _
    Public Function ResetLSNo(ByVal pBuyer As String, ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Dim xSeqno As Integer
        Dim xSeqNoString, xGroup, LF_LastNo, CTL_LastNo As String
        Try
            ' 取得 ControlRecord Last LSNo
            CTL_LastNo = ""
            xGroup = ""
            sql = "SELECT GroupCode, LSNo FROM M_FControlRecord "
            sql &= "Where Buyer = '" & pBuyer & "' "
            Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
            If dt_ControlRecord.Rows.Count > 0 Then
                CTL_LastNo = dt_ControlRecord.Rows(0).Item("LSNo")
                xGroup = dt_ControlRecord.Rows(0).Item("GroupCode")
            End If
            ' 取得 LastFinal Last LSNo
            LF_LastNo = xGroup + "L" + "1310001"
            sql = "SELECT Top 1 LSNo FROM LF_LocalStockPlan "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And LSNo Like '" & xGroup + "L" & "%' "
            sql &= "Order by LSNo Desc "
            Dim dt_LFPlan As DataTable = oDataBase.GetDataTable(sql)
            If dt_LFPlan.Rows.Count > 0 Then
                xSeqno = CInt(Mid(dt_LFPlan.Rows(0).Item("LSNo"), 7, 4)) + 1
                xSeqNoString = CStr(xSeqno)
                If Len(xSeqNoString) < 2 Then
                    xSeqNoString = "000" + CStr(xSeqno)
                Else
                    If Len(xSeqNoString) < 3 Then
                        xSeqNoString = "00" + CStr(xSeqno)
                    Else
                        If Len(xSeqNoString) < 4 Then
                            xSeqNoString = "0" + CStr(xSeqno)
                        End If
                    End If
                End If
                LF_LastNo = Mid(dt_LFPlan.Rows(0).Item("LSNo"), 1, 6) + xSeqNoString
            End If
            ' 重新設定下一次可使用LSNo
            If CTL_LastNo <> LF_LastNo Then
                UpdateNextLSNo(pBuyer, LF_LastNo)
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'ResetLSNo-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([650]-ResetBuyerLSNo)
    '**     重新設定下一次可使用BuyerLSNo
    '***********************************************************************************************
    'ResetBuyerLSNo-Start
    <WebMethod()> _
    Public Function ResetBuyerLSNo(ByVal pBuyer As String, ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Dim xSeqno As Integer
        Dim xSeqNoString, xBUGroup, LF_LastNo, CTL_LastNo As String
        Try
            ' 取得 ControlRecord Last BuyerLSNo
            CTL_LastNo = ""
            xBUGroup = ""
            sql = "SELECT BUGroupCode, BULSNo FROM M_FControlRecord "
            sql &= "Where Buyer = '" & pBuyer & "' "
            Dim dt_ControlRecord As DataTable = oDataBase.GetDataTable(sql)
            If dt_ControlRecord.Rows.Count > 0 Then
                CTL_LastNo = dt_ControlRecord.Rows(0).Item("BULSNo")
                xBUGroup = dt_ControlRecord.Rows(0).Item("BUGroupCode")
            End If
            ' 取得 LastFinal Last LSNo
            LF_LastNo = xBUGroup + "L" + "1310001"
            sql = "SELECT Top 1 BULSNo FROM LF_BuyerLocalStockPlan "
            sql &= "Where BuyerGroup = '" & pGRBuyer & "' "
            sql &= "  And BULSNo Like '" & xBUGroup + "L" & "%' "
            sql &= "Order by BULSNo Desc "
            Dim dt_LFPlan As DataTable = oDataBase.GetDataTable(sql)
            If dt_LFPlan.Rows.Count > 0 Then
                xSeqno = CInt(Mid(dt_LFPlan.Rows(0).Item("BULSNo"), 7, 4)) + 1
                xSeqNoString = CStr(xSeqno)
                If Len(xSeqNoString) < 2 Then
                    xSeqNoString = "000" + CStr(xSeqno)
                Else
                    If Len(xSeqNoString) < 3 Then
                        xSeqNoString = "00" + CStr(xSeqno)
                    Else
                        If Len(xSeqNoString) < 4 Then
                            xSeqNoString = "0" + CStr(xSeqno)
                        End If
                    End If
                End If
                LF_LastNo = Mid(dt_LFPlan.Rows(0).Item("BULSNo"), 1, 6) + xSeqNoString
            End If
            ' 重新設定下一次可使用LSNo
            If CTL_LastNo <> LF_LastNo Then
                UpdateNextBuyerLSNo(pBuyer, LF_LastNo)
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'ResetBuyerLSNo-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([660]-UpdateBYFCTStatus)
    '**     BY FCT 資料同步化(ADIDAS REF + WORKING + ARTICLE)
    '***********************************************************************************************
    'UpdateBYFCTStatus-Start
    <WebMethod()> _
    Public Function UpdateBYFCTStatus(ByVal pBuyer As String, ByVal pUserID As String, ByVal pStatus As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim i As Integer
        Dim sql As String
        Try
            '
            ' -----------------------------------------------------------------------------------
            ' 資料同步化    更新同一 ADIDAS REF + WORKING + ARTICLE + PART + FACTORY NAME 中資料狀態
            ' -----------------------------------------------------------------------------------
            sql = "SELECT C1, D1, E1, G1, L1 FROM E_InputSheetBY "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And EY1   = '" & "*" & "' "
            sql &= "  And EZ1   = '" & pStatus & "' "
            sql &= "Group by C1, D1, E1, G1, L1 "
            sql &= "Order by C1, D1, E1, G1, L1 "
            Dim dt_InputSheet As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_InputSheet.Rows.Count - 1
                '
                sql = "Update E_InputSheetBY Set "
                If pStatus <> "41" And pStatus <> "50" Then
                    sql &= " EY1 = '', "
                End If
                sql &= " EZ1 = '" & pStatus & "', "
                sql &= " ModifyUser = '" & pUserID & "', "
                sql &= " ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " Where Buyer = '" & pBuyer & "' "
                sql &= "   And C1    = '" & dt_InputSheet.Rows(i).Item("C1").ToString & "' "
                sql &= "   And D1    = '" & dt_InputSheet.Rows(i).Item("D1").ToString & "' "
                sql &= "   And E1    = '" & dt_InputSheet.Rows(i).Item("E1").ToString & "' "
                sql &= "   And G1    = '" & dt_InputSheet.Rows(i).Item("G1").ToString & "' "
                sql &= "   And L1    = '" & dt_InputSheet.Rows(i).Item("L1").ToString & "' "
                oDataBase.ExecuteNonQuery(sql)
            Next
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'UpdateBYFCTStatus-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([670]-GetNativeVendorCode)
    '**     取得Native Vendor Code
    '***********************************************************************************************
    'GetNativeVendorCode-Start
    <WebMethod()> _
    Public Function GetNativeVendorCode(ByVal pBuyer As String, ByVal pFactory As String, ByVal pType As String) As String
        Dim RtnCode As String = ""
        '
        Dim sql As String
        Try
            sql = "SELECT CustName, CustCode FROM M_NativeVendor "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And Active = 1 "
            sql &= "  And Factoryname = '" & pFactory & "' "
            Dim dt_Vendor As DataTable = oDataBase.GetDataTable(sql)
            If dt_Vendor.Rows.Count > 0 Then
                If pType = "NAME" Then
                    RtnCode = dt_Vendor.Rows(0).Item("CustName")
                Else
                    RtnCode = dt_Vendor.Rows(0).Item("CustCode")
                End If
            End If
        Catch ex As Exception
            RtnCode = ""
        End Try
        '
        Return RtnCode
    End Function
    'GetNativeVendorCode-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([680]-GetColor)
    '**     取得Color Master String
    '***********************************************************************************************
    'GetColor-Start
    <WebMethod()> _
    Public Function GetColor(ByVal pType As String, ByVal pBuyerCode As String, ByVal pColorString As String, _
                             ByVal pColor As String, ByVal pColorCode As String) As String
        Dim RtnCode As String = ""
        '
        Dim i As Integer
        Dim sql, str As String
        Dim xColorString As Object
        Try
            '
            ' 分解 COLOR STRING
            xColorString = pColorString.Split("/")
            '
            ' TYPE=1    SOLID   /   TYPE=2  MULTI
            If pType = 1 Then
                ' --------------------------------------------------------------------------
                ' SOLID PROC.
                ' --------------------------------------------------------------------------
                xColorString(0) = pColorCode
                xColorString(2) = pColorCode
                xColorString(4) = pColorCode
            Else
                ' --------------------------------------------------------------------------
                ' MULTI PROC.
                ' --------------------------------------------------------------------------
                ' 取得此結構(ELEMENT)之INDEX
                sql = "Select Data From M_Referp "
                sql = sql & " Where Cat  = '" & "150" & "' "
                sql = sql & "   And DKey = '" & pBuyerCode + "-" + pColor & "' "
                sql = sql & "   And DKey <> '" & pBuyerCode + "-ZIPPER" & "' "
                Dim dt_Referp As DataTable = oDataBase.GetDataTable(sql)
                If dt_Referp.Rows.Count > 0 Then
                    '
                    ' 依INDEX設定 COLOR STRING
                    str = dt_Referp.Rows(0).Item("Data")
                    While InStr(str, ",") > 0
                        '
                        i = CInt(Mid(str, 1, InStr(str, ",") - 1)) - 1
                        xColorString(i) = pColorCode
                        str = Mid(str, InStr(str, ",") + 1)
                    End While
                    i = CInt(str) - 1
                    xColorString(i) = pColorCode
                End If
            End If
            '
            ' 各COLOR連結成String後送回主程式
            For i = 0 To 11
                If i = 0 Then
                    RtnCode = xColorString(i)
                Else
                    RtnCode = RtnCode + "/" + xColorString(i)
                End If
            Next
            RtnCode = RtnCode + "/"
            '
        Catch ex As Exception
            RtnCode = "ERROR"
        End Try
        '
        Return RtnCode
    End Function
    'GetColor-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([690]-CheckItemNoDisplay)
    '**     檢測Item NODISPLAY
    '***********************************************************************************************
    'CheckItemNoDisplay-Start
    <WebMethod()> _
    Public Function CheckItemNoDisplay(ByVal pCode As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            RtnCode = oWaves.CheckItemCode(pCode)
        Catch ex As Exception
            RtnCode = 9
        End Try
        '
        Return RtnCode
    End Function
    'CheckItemNoDisplay-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([695]-CheckColorNotFound)
    '**     檢測Color Code  Not Found
    '***********************************************************************************************
    'CheckColorCode-Start
    <WebMethod()> _
    Public Function CheckColorNotFound(ByVal pCode As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            RtnCode = oWaves.CheckColorCode(pCode)
        Catch ex As Exception
            RtnCode = 9
        End Try
        '
        Return RtnCode
    End Function
    'CheckItemNoDisplay-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([700]-GetProdQtyUnit)
    '**     取得LS中PROD的數量單位
    '***********************************************************************************************
    'GetProdQtyUnit-Start
    <WebMethod()> _
    Public Function GetProdQtyUnit(ByVal pBuyer As String, ByVal pPartType As String, ByRef pQtyUnit As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        pQtyUnit = 1
        Try
            sql = "SELECT Data FROM M_Referp "
            sql &= "Where Cat = '" & "112" & "' "
            sql &= "  And DKey = '" & pBuyer & "-" & pPartType & "' "
            sql &= "Order by Data "
            Dim dt_List As DataTable = oDataBase.GetDataTable(sql)
            If dt_List.Rows.Count > 0 Then
                pQtyUnit = CInt(dt_List.Rows(0).Item("Data"))
            End If
        Catch ex As Exception
            RtnCode = 9
        End Try
        '
        Return RtnCode
    End Function
    'GetProdQtyUnit-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([705]-GetZipper)
    '**     判斷是否為ZIPPER (0:無, 1:ZIPPER, 2:CH, 3:SLD)
    '***********************************************************************************************
    'GetZipper-Start
    <WebMethod()> _
    Public Function GetZipper(ByVal pBuyer As String, ByVal pItem As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Try
            sql = "SELECT Top 1 YName1+YName2+YName3 As ItemName FROM M_ItemConvert "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And YCode = '" & pItem & "' "
            sql &= "Order by ItemName "
            Dim dt_ItemConvert As DataTable = oDataBase.GetDataTable(sql)
            If dt_ItemConvert.Rows.Count > 0 Then
                If InStr(dt_ItemConvert.Rows(0).Item("ItemName"), "1/1/") > 0 Then
                    RtnCode = 1
                End If
                If InStr(dt_ItemConvert.Rows(0).Item("ItemName"), "1/2/") > 0 Then
                    RtnCode = 2
                End If
                If InStr(dt_ItemConvert.Rows(0).Item("ItemName"), "1/3/") > 0 Then
                    RtnCode = 3
                End If
            End If
            '
            If RtnCode = 0 Then
                Dim xCode As Integer = 0
                Dim xItemStatistics As String
                xCode = oWaves.GetItemStatistics(pItem, xItemStatistics)
                If xCode = 0 Then
                    Dim xItemStat As Object = Split(xItemStatistics, "/")
                    If xItemStat(0) = "1" And xItemStat(3) = "1" Then
                        RtnCode = 1
                    End If
                    If xItemStat(0) = "1" And xItemStat(3) = "2" Then
                        RtnCode = 2
                    End If
                    If xItemStat(0) = "1" And xItemStat(3) = "3" Then
                        RtnCode = 3
                    End If
                End If
            End If
        Catch ex As Exception
            RtnCode = 9
        End Try
        '
        Return RtnCode
    End Function
    'GetZipper-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([710]-GetFilterRuleNo)
    '**     尋找適合備料基準    從過濾箱中(M_LSOrderRule)取得適合備料基準
    '***********************************************************************************************
    'GetFilterRuleNo-Start
    <WebMethod()> _
    Public Function GetFilterRuleNo(ByVal pBuyer As String, ByVal pLTType As String, ByVal pPartType As String, ByVal pSearchItem() As String, _
                                    ByRef pRuleNo() As String, ByRef pAction() As String, ByRef pObjectType() As String, ByRef pObjectProduct() As String, ByRef pCount As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim xRuleIndex As Integer = 1
        Dim sql, xField As String
        Dim xRuleCode, xRuleNo As String
        Dim xKeyData(10) As String
        Dim FoundRule As Boolean
        '
        Try
            For i As Integer = 1 To 5
                pRuleNo(i) = ""
                pAction(i) = ""
                pObjectType(i) = ""
                pObjectProduct(i) = ""
            Next
            xRuleIndex = 1
            pCount = 0
            '
            ' 過濾箱-過濾處理
            sql = "SELECT * FROM M_LSOrderRule "
            sql &= "Where Active = '1' "
            sql &= "  And Buyer = '" & pBuyer & "' "
            sql &= "  And LTType = '" & pLTType & "' "
            sql &= "  And PartType = '" & pPartType & "' "
            sql &= "Order by RuleCode, RuleNo, RuleSubno "
            Dim dt_LSOrderRule As DataTable = oDataBase.GetDataTable(sql)
            '
            If dt_LSOrderRule.Rows.Count > 0 Then
                ' 變數初值設定
                xRuleCode = dt_LSOrderRule.Rows(0).Item("RuleCode").ToString
                xRuleNo = dt_LSOrderRule.Rows(0).Item("RuleNo").ToString
            End If
            For i As Integer = 0 To dt_LSOrderRule.Rows.Count - 1
                ' 已搜尋到Rule-No 並 RuleCode+RuleNo 不合
                If xRuleIndex - 1 > 0 Then
                    If xRuleCode <> dt_LSOrderRule.Rows(i).Item("RuleCode").ToString Or _
                       xRuleNo <> dt_LSOrderRule.Rows(i).Item("RuleNo").ToString Then
                        Exit For
                    End If
                End If
                ' 分解SearchIndex--建置SearchItem的 [KeyData]
                If dt_LSOrderRule.Rows(i).Item("SearchIndex") = "----------" Then
                    xRuleCode = dt_LSOrderRule.Rows(0).Item("RuleCode").ToString
                    xRuleNo = dt_LSOrderRule.Rows(0).Item("RuleNo").ToString
                    pRuleNo(xRuleIndex) = xRuleCode + "-" + xRuleNo + "-" + dt_LSOrderRule.Rows(i).Item("RuleSubno").ToString
                    pAction(xRuleIndex) = dt_LSOrderRule.Rows(i).Item("Action").ToString
                    pObjectType(xRuleIndex) = dt_LSOrderRule.Rows(i).Item("ObjectType").ToString
                    ' 引手對應
                    'If Trim(dt_LSOrderRule.Rows(i).Item("ObjectProduct").ToString) = "Z" Then
                    '    pObjectProduct(xRuleIndex) = Chr(32) & dt_LSOrderRule.Rows(i).Item("ObjectProduct").ToString & Chr(32)
                    'Else
                    'End If
                    pObjectProduct(xRuleIndex) = dt_LSOrderRule.Rows(i).Item("ObjectProduct").ToString
                    xRuleIndex = xRuleIndex + 1
                    Exit For
                Else
                    For j As Integer = 1 To 10
                        xKeyData(j) = ""
                        If Mid(dt_LSOrderRule.Rows(i).Item("SearchIndex"), j, 1) = "1" Then
                            xField = "SearchItem" + CStr(j)
                            xKeyData(j) = dt_LSOrderRule.Rows(i).Item(xField).ToString
                        End If
                    Next
                End If
                ' 比較-----[%]+[?]處理-----[%x??x]處理(特別要求時處理需REVIEW)
                '       |               |
                '       |               +--[x??x%]處理(特別要求時處理需REVIEW)
                '       |               |
                '       |               +--[%x??x%]處理(不對應)
                '       |
                '       +--[%]處理---------[%xxxx]處理
                '       |               |
                '       |               +--[xxxx%]處理
                '       |               |
                '       |               +--[%xxx%]處理
                '       |
                '       +--[?]處理
                '       |
                '       +--[xxxx]處理
                FoundRule = True
                For j As Integer = 1 To 10
                    If xKeyData(j) <> "" Then
                        If InStr(xKeyData(j), "%") > 0 And InStr(xKeyData(j), "?") > 0 Then
                            ' [%]+[?]處理
                            If Mid(xKeyData(j), 1, 1) = "%" And Mid(xKeyData(j), Len(xKeyData(j)), 1) <> "%" Then
                                ' [%x??x]處理
                                Dim xItemLength As Integer = Len(pSearchItem(j))
                                Dim xKeyDataLength As Integer = Len(xKeyData(j)) - 1
                                Dim xItemDataStr As String = Mid(pSearchItem(j), xItemLength - xKeyDataLength + 1)
                                Dim xKeyDataStr As String = Mid(xKeyData(j), 2)
                                '
                                For k As Integer = 1 To xKeyDataLength
                                    If Mid(xKeyDataStr, k, 1) <> "?" Then
                                        If Mid(xItemDataStr, k, 1) <> Mid(xKeyDataStr, k, 1) Then
                                            FoundRule = False
                                            Exit For
                                        End If
                                    End If
                                Next
                            Else
                                If Mid(xKeyData(j), 1, 1) <> "%" And Mid(xKeyData(j), Len(xKeyData(j)), 1) = "%" Then
                                    ' [x??x%]處理
                                    Dim xKeyDataLength As Integer = Len(xKeyData(j)) - 1
                                    Dim xItemDataStr As String = Mid(pSearchItem(j), 1, xKeyDataLength)
                                    Dim xKeyDataStr As String = Mid(xKeyData(j), 1, xKeyDataLength)
                                    '
                                    For k As Integer = 1 To xKeyDataLength
                                        If Mid(xKeyDataStr, k, 1) <> "?" Then
                                            If Mid(xItemDataStr, k, 1) <> Mid(xKeyDataStr, k, 1) Then
                                                FoundRule = False
                                                Exit For
                                            End If
                                        End If
                                    Next
                                Else
                                    ' [%x??x%] or 其他處理
                                    FoundRule = False
                                    Exit For
                                End If
                            End If
                        Else
                            If InStr(xKeyData(j), "%") > 0 Then
                                ' [%]處理
                                If Mid(xKeyData(j), 1, 1) = "%" And Mid(xKeyData(j), Len(xKeyData(j)), 1) <> "%" Then
                                    ' [%xxxx]處理
                                    Dim xItemLength As Integer = Len(pSearchItem(j))
                                    Dim xKeyDataLength As Integer = Len(xKeyData(j)) - 1

                                    If InStr(pSearchItem(j), "/") > 0 Then
                                        ' 特別要求時
                                        Dim OK As Boolean = False
                                        Dim xStr1, xStr2 As String
                                        xStr1 = pSearchItem(j)
                                        While InStr(xStr1, "/") > 0
                                            xStr2 = Mid(xStr1, 1, InStr(xStr1, "/") - 1)
                                            xItemLength = Len(xStr2)
                                            If Mid(xStr2, xItemLength - xKeyDataLength + 1) <> Mid(xKeyData(j), 2) Then
                                                xStr1 = Mid(xStr1, InStr(xStr1, "/") + 1)
                                            Else
                                                OK = True
                                                Exit While
                                            End If
                                        End While
                                        If OK = False Then
                                            FoundRule = False
                                            Exit For
                                        End If
                                    Else
                                        If Mid(pSearchItem(j), xItemLength - xKeyDataLength + 1) <> Mid(xKeyData(j), 2) Then
                                            FoundRule = False
                                            Exit For
                                        End If
                                    End If
                                Else
                                    If Mid(xKeyData(j), 1, 1) <> "%" And Mid(xKeyData(j), Len(xKeyData(j)), 1) = "%" Then
                                        ' [xxxx%]處理
                                        Dim xKeyDataLength As Integer = Len(xKeyData(j)) - 1

                                        If InStr(pSearchItem(j), "/") > 0 Then
                                            ' 特別要求時
                                            Dim OK As Boolean = False
                                            Dim xStr1, xStr2 As String
                                            xStr1 = pSearchItem(j)
                                            While InStr(xStr1, "/") > 0
                                                xStr2 = Mid(xStr1, 1, InStr(xStr1, "/") - 1)
                                                If Mid(xStr2, 1, xKeyDataLength) <> Mid(xKeyData(j), 1, xKeyDataLength) Then
                                                    xStr1 = Mid(xStr1, InStr(xStr1, "/") + 1)
                                                Else
                                                    OK = True
                                                    Exit While
                                                End If
                                            End While
                                            If OK = False Then
                                                FoundRule = False
                                                Exit For
                                            End If
                                        Else
                                            If Mid(pSearchItem(j), 1, xKeyDataLength) <> Mid(xKeyData(j), 1, xKeyDataLength) Then
                                                FoundRule = False
                                                Exit For
                                            End If
                                        End If
                                    Else
                                        ' [%xxxx%]處理
                                        Dim Str As String = Mid(xKeyData(j), 2)
                                        Str = Mid(Str, 1, Len(Str) - 1)

                                        If InStr(pSearchItem(j), "/") > 0 Then
                                            ' 特別要求時
                                            Dim OK As Boolean = False
                                            Dim xStr1, xStr2 As String
                                            xStr1 = pSearchItem(j)
                                            While InStr(xStr1, "/") > 0
                                                xStr2 = Mid(xStr1, 1, InStr(xStr1, "/") - 1)
                                                If InStr(xStr2, Str) <= 0 Then
                                                    xStr1 = Mid(xStr1, InStr(xStr1, "/") + 1)
                                                Else
                                                    OK = True
                                                    Exit While
                                                End If
                                            End While
                                            If OK = False Then
                                                FoundRule = False
                                                Exit For
                                            End If
                                        Else
                                            If InStr(pSearchItem(j), Str) <= 0 Then
                                                FoundRule = False
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            Else

                                If InStr(xKeyData(j), "?") > 0 Then
                                    ' [?]處理  P??L / ??PL / PL??
                                    If InStr(pSearchItem(j), "/") > 0 Then
                                        ' 特別要求時
                                        Dim OK, xErr As Boolean
                                        Dim xStr1, xStr2 As String
                                        OK = False
                                        xStr1 = pSearchItem(j)
                                        While InStr(xStr1, "/") > 0
                                            xStr2 = Mid(xStr1, 1, InStr(xStr1, "/") - 1)
                                            xErr = False
                                            For k As Integer = 1 To Len(xKeyData(j))
                                                If Mid(xKeyData(j), k, 1) <> "?" Then
                                                    If Mid(xStr2, k, 1) <> Mid(xKeyData(j), k, 1) Then
                                                        xStr1 = Mid(xStr1, InStr(xStr1, "/") + 1)
                                                        xErr = True
                                                        Exit For
                                                    End If
                                                End If
                                            Next
                                            If xErr = False Then
                                                OK = True
                                                Exit While
                                            End If
                                        End While
                                        If OK = False Then
                                            FoundRule = False
                                            Exit For
                                        End If
                                    Else
                                        For k As Integer = 1 To Len(xKeyData(j))
                                            If Mid(xKeyData(j), k, 1) <> "?" Then
                                                If Mid(pSearchItem(j), k, 1) <> Mid(xKeyData(j), k, 1) Then
                                                    FoundRule = False
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                    End If
                                Else
                                    ' [xxxx]處理
                                    If InStr(pSearchItem(j), "/") > 0 Then
                                        ' 特別要求時
                                        Dim OK As Boolean
                                        Dim xStr1, xStr2 As String
                                        OK = False
                                        xStr1 = pSearchItem(j)
                                        While InStr(xStr1, "/") > 0
                                            xStr2 = Mid(xStr1, 1, InStr(xStr1, "/") - 1)
                                            If xStr2 <> xKeyData(j) Then
                                                xStr1 = Mid(xStr1, InStr(xStr1, "/") + 1)
                                            Else
                                                OK = True
                                                Exit While
                                            End If
                                        End While
                                        If OK = False Then
                                            FoundRule = False
                                            Exit For
                                        End If
                                    Else
                                        If pSearchItem(j) <> xKeyData(j) Then
                                            FoundRule = False
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If          ' xKeyData(j) = ""
                Next
                '
                If FoundRule Then
                    ' 儲存搜尋到 Rule-No
                    xRuleCode = dt_LSOrderRule.Rows(i).Item("RuleCode").ToString
                    xRuleNo = dt_LSOrderRule.Rows(i).Item("RuleNo").ToString
                    pRuleNo(xRuleIndex) = xRuleCode + "-" + xRuleNo + "-" + dt_LSOrderRule.Rows(i).Item("RuleSubno").ToString
                    pAction(xRuleIndex) = dt_LSOrderRule.Rows(i).Item("Action").ToString
                    pObjectType(xRuleIndex) = dt_LSOrderRule.Rows(i).Item("ObjectType").ToString
                    ' 引手對應
                    'If Trim(dt_LSOrderRule.Rows(i).Item("ObjectProduct").ToString) = "Z" Then
                    '    pObjectProduct(xRuleIndex) = Chr(32) & dt_LSOrderRule.Rows(i).Item("ObjectProduct").ToString & Chr(32)
                    'Else
                    'End If
                    '
                    pObjectProduct(xRuleIndex) = dt_LSOrderRule.Rows(i).Item("ObjectProduct").ToString
                    xRuleIndex = xRuleIndex + 1
                End If
            Next
            '
            pCount = xRuleIndex - 1
        Catch ex As Exception
            RtnCode = 9
        End Try
        '
        Return RtnCode
    End Function
    'GetFilterRuleNo-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([800]-MaterialExpansion )
    '**       材料分析BOM展開 ADIDAS, REEBOK
    '***********************************************************************************************
    'MaterialExpansion-Start
    <WebMethod()> _
    Public Sub MaterialExpansion()
        ' ***********************************************************************************
        ' 變數定義及設定初值
        ' ***********************************************************************************
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim i As Integer = 0
        Dim Run As String = ""
        Dim RunBuyer As String = ""
        Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")     ' 現在日時    
        Dim xSeasonYY As String = ""
        Dim xBuyerList As String = ""
        '
        Try
            sql = "Select * From M_Referp "
            sql = sql & " Where Cat  = '" & "200" & "' "
            sql = sql & "   And DKey = '" & "AGENT-MATERIAL" & "' "
            Dim dt_Run As DataTable = oDataBase.GetDataTable(sql)
            If dt_Run.Rows.Count > 0 Then
                Run = Mid(dt_Run.Rows(0).Item("Data"), 1, 1)
            End If

            sql = "Select * From M_Referp "
            sql = sql & " Where Cat  = '" & "200" & "' "
            sql = sql & "   And DKey = '" & "MATERIAL-BUYER" & "' "
            sql = sql & "   And Data Like '%" & "Y" & "%' "
            Dim dt_RunBuyer As DataTable = oDataBase.GetDataTable(sql)
            If dt_RunBuyer.Rows.Count > 0 Then
                RunBuyer = "Y"
            End If
            ' ***********************************************************************************
            ' 材料展開
            ' ***********************************************************************************
            If Run = "Y" And RunBuyer = "Y" Then
                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                ' 處理FLAG = N
                '
                sql = "Update M_Referp Set "
                sql = sql + " Data = '" & "N" & "/" & Now.ToString("yyyyMMddHHmmss") & "' "
                sql = sql & " Where Cat  = '" & "200" & "' "
                sql = sql & "   And DKey = '" & "AGENT-MATERIAL" & "' "
                oDataBase.ExecuteNonQuery(sql)
                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                ' 取得對象BUYER
                '
                xBuyerList = ""
                sql = "Select Top 1 Data From M_Referp "
                sql = sql & " Where Cat  = '" & "200" & "' "
                sql = sql & "   And DKey = '" & "MATERIAL-BUYER" & "' "
                sql = sql & "   And Data Like '%" & "Y" & "%' "
                sql = sql & "Order by Data "
                Dim dt_Buyer As DataTable = oDataBase.GetDataTable(sql)
                For i = 0 To dt_Buyer.Rows.Count - 1
                    xBuyerList = Mid(dt_Buyer.Rows(i).Item("Data"), 1, 6)
                Next
                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                ' 取得 VDP計算開始年度
                '
                xSeasonYY = "99"
                sql = "Select Data From M_Referp "
                sql = sql & " Where Cat  = '" & "200" & "' "
                sql = sql & "   And DKey = '" & "AGENT-ACT-VDPSTART" & "' "
                sql = sql & "Order by Data "
                Dim dt_Referp As DataTable = oDataBase.GetDataTable(sql)
                If dt_Referp.Rows.Count > 0 Then
                    xSeasonYY = dt_Referp.Rows(0).Item("Data")
                End If
                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                ' 資料清除
                sql = "Delete From A_MaterialsAnalysis_SLD "
                sql &= "Where Buyer = '" & xBuyerList & "' "
                sql &= "  And Substring(Season, 3,2) >= '" & xSeasonYY & "' "
                oDataBase.ExecuteNonQuery(sql)
                '
                sql = "Delete From A_MaterialsAnalysis_CH "
                sql &= "Where Buyer = '" & xBuyerList & "' "
                sql &= "  And Substring(Season, 3,2) >= '" & xSeasonYY & "' "
                oDataBase.ExecuteNonQuery(sql)
                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                ' 構成資料展開
                '
                oWaves.Timeout = Timeout.Infinite
                '
                Dim xRubberColorList() As String
                Dim xPartType, xItemProduct, xClass, xQtyMeter, xColor, xName, xSales, xPurchase, xProduction, xFinish, xUpRubberColor, xLoRubberColor As String
                Dim xItem(5), xItemName(5), xQty(5) As String
                Dim xCount, idx As Integer
                '
                ' 清除 ZIP-處理FLAG(Yobi1)
                sql = "Update A_MaterialsAnalysis_ZIP Set "
                sql &= "Yobi1 = '" & "" & "', "
                sql &= "ModifyUser = '" & "" & "', "
                sql &= "ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= "Where Buyer = '" & xBuyerList & "' "
                sql &= "  And Substring(Season, 3,2) >= '" & xSeasonYY & "' "
                oDataBase.ExecuteNonQuery(sql)
                ' ----------------------------------------------------------------
                ' Trace Log
                ' ----------------------------------------------------------------
                sql = "Delete From W_TraceLog "
                sql &= "Where Pgm = '" & "FAS-Material" & "' "
                oDataBase.ExecuteNonQuery(sql)
                '
                sql = "Insert into W_TraceLog "
                sql &= "( Pgm, Data, CreateTime ) "
                sql &= "VALUES( "
                sql &= " '" & "FAS-Material" & "', "
                sql &= " '" & xBuyerList & "--- Start ---" & "', "
                sql &= " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " ) "
                oDataBase.ExecuteNonQuery(sql)
                ' ----------------------------------------------------------------
                '
                sql = "Select Item From A_MaterialsAnalysis_ZIP "
                sql &= "Where Buyer = '" & xBuyerList & "' "
                sql &= "  And Substring(Season, 3,2) >= '" & xSeasonYY & "' "
                sql &= "Group by Item "
                sql &= "Order by Item "
                Dim dt_ZIPITEM As DataTable = oDataBase.GetDataTable(sql)
                For i = 0 To dt_ZIPITEM.Rows.Count - 1
                    ' ----------------------------------------------------------------
                    ' CHAIN 展開
                    ' ----------------------------------------------------------------
                    '
                    ' 取得完成品CHAIN(CF/VF) 或 GAP-CHAIN(MF)     (MF=金屬/CF=樹脂/VF=塑鋼)
                    xPartType = "CH"                                                            ' 決定材料種類(PartTye)
                    oWaves.GetItemProducta(dt_ZIPITEM.Rows(i).Item("Item"), xItemProduct)        ' 取得製品區分 
                    If xItemProduct = "MF" Then
                        oWaves.GetChildItemStructurea("01", "CH-GAP", dt_ZIPITEM.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                    Else
                        oWaves.GetChildItemStructurea("01", "CH-DYED", dt_ZIPITEM.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                    End If
                    ' 展開ITEM構成取得所指定ITEM
                    For ItemIndex As Integer = 1 To xCount
                        '
                        sql = "Select Color From A_MaterialsAnalysis_ZIP "
                        sql &= "Where Buyer = '" & xBuyerList & "' "
                        sql &= "  And Substring(Season, 3,2) >= '" & xSeasonYY & "' "
                        sql &= "  And Item = '" & dt_ZIPITEM.Rows(i).Item("Item") & "' "
                        sql &= "Group by Color "
                        sql &= "Order by Color "
                        Dim dt_ZIPCOLOR As DataTable = oDataBase.GetDataTable(sql)
                        For j As Integer = 0 To dt_ZIPCOLOR.Rows.Count - 1
                            ' 取得完成品CHAIN兼用色
                            oWaves.GetChangeColora("01", dt_ZIPITEM.Rows(i).Item("Item"), xItem(ItemIndex), dt_ZIPCOLOR.Rows(j).Item("Color"), xColor)
                            ' 產生 CHAIN材料構成檔
                            sql = "Select * From A_MaterialsAnalysis_ZIP "
                            sql &= "Where Buyer = '" & xBuyerList & "' "
                            sql &= "  And Substring(Season, 3,2) >= '" & xSeasonYY & "' "
                            sql &= "  And Item = '" & dt_ZIPITEM.Rows(i).Item("Item") & "' "
                            sql &= "  And Color = '" & dt_ZIPCOLOR.Rows(j).Item("Color") & "' "
                            sql &= "Order by Buyer, Season, Month, Version, CustCode, CustItem, CustColor, Item, Color "
                            Dim dt_ZIPDATA As DataTable = oDataBase.GetDataTable(sql)
                            For k As Integer = 0 To dt_ZIPDATA.Rows.Count - 1
                                '
                                ' 取得Meter換算基準 (取得ItemClass)
                                oWaves.GetItemClassa(dt_ZIPDATA.Rows(k).Item("Item"), xClass)
                                sql = "Select * From M_Referp "
                                sql = sql & " Where Cat  = '" & "110" & "' "
                                sql = sql & "   And DKey = 'FALL-" & dt_ZIPDATA.Rows(k).Item("Buyer") + "-" + xClass & "' "
                                Dim dt_Meter As DataTable = oDataBase.GetDataTable(sql)
                                If dt_Meter.Rows.Count > 0 Then
                                    xQtyMeter = CStr(CDbl(dt_Meter.Rows(0).Item("Data")) * 100)
                                Else
                                    xQtyMeter = "100"
                                End If
                                '
                                sql = "Insert into A_MaterialsAnalysis_CH "
                                sql &= "( "
                                sql &= "DataType, Buyer, Season, CustCode, Month, "
                                sql &= "Version, CustItem, CustColor, ParentItem, ParentColor, SeqNo, "
                                sql &= "Item, Color, ItemName1, ItemName2, ItemName3, "
                                sql &= "Import, Produce, Yobi1, Yobi2, Yobi3, "
                                sql &= "FCTQty, ACTQty, "
                                sql &= "CreateUser, CreateTime, ModifyUser, ModifyTime "
                                sql &= " ) "
                                sql &= "VALUES( "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("DataType").ToString & "', "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("Buyer").ToString & "', "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("Season").ToString & "', "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("CustCode").ToString & "', "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("Month").ToString & "', "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("Version").ToString & "', "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("CustItem").ToString & "', "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("CustColor").ToString & "', "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("Item").ToString & "', "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("Color").ToString & "', "
                                sql &= " '" & "0" + CStr(ItemIndex) & "', "

                                sql &= " '" & xItem(ItemIndex) & "', "
                                sql &= " '" & xColor & "', "
                                oWaves.GetItemName001a(xItem(ItemIndex), 1, xName)
                                sql &= " '" & xName & "', "
                                oWaves.GetItemName001a(xItem(ItemIndex), 2, xName)
                                sql &= " '" & xName & "', "
                                oWaves.GetItemName001a(xItem(ItemIndex), 3, xName)
                                sql &= " '" & xName & "', "

                                oWaves.GetItemProdType(xItem(ItemIndex), xSales, xPurchase, xProduction)
                                sql &= " '" & xPurchase & "', "
                                sql &= " '" & xProduction & "', "
                                sql &= " '" & "" & "', "
                                sql &= " '" & "" & "', "
                                sql &= " '" & "" & "', "

                                sql &= " " & CStr(Fix(CInt(xQtyMeter) / 100 * CDbl(dt_ZIPDATA.Rows(k).Item("FCTQty")) + 0.99)) & ", "
                                sql &= " " & CStr(Fix(CInt(xQtyMeter) / 100 * CDbl(dt_ZIPDATA.Rows(k).Item("ACTQty")) + 0.99)) & ", "
                                sql &= " '" & "FCTACT-BOM" & "', "
                                sql &= " '" & NowDateTime & "', "
                                sql &= " '" & "" & "', "
                                sql &= " Null "
                                sql &= " ) "
                                oDataBase.ExecuteNonQuery(sql)
                            Next
                        Next
                    Next    ' CH
                    ' ----------------------------------------------------------------
                    ' Trace Log
                    ' ----------------------------------------------------------------
                    sql = "Insert into W_TraceLog "
                    sql &= "( Pgm, Data, CreateTime ) "
                    sql &= "VALUES( "
                    sql &= " '" & "FAS-Material" & "', "
                    sql &= " '" & xBuyerList & "CH=" & dt_ZIPITEM.Rows(i).Item("Item") & "', "
                    sql &= " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                    sql &= " ) "
                    oDataBase.ExecuteNonQuery(sql)
                    ' ----------------------------------------------------------------
                    ' SLIDER 展開
                    ' ----------------------------------------------------------------
                    '
                    xQtyMeter = "100"                       ' 取得Meter換算基準 (取得ItemClass)
                    xPartType = "SLD"                       ' 決定材料種類(PartTye)
                    oWaves.GetChildItemStructurea("01", "SLD-FINISH", dt_ZIPITEM.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                    ' 展開ITEM構成取得所指定ITEM
                    For ItemIndex As Integer = 1 To xCount
                        '
                        sql = "Select Color From A_MaterialsAnalysis_ZIP "
                        sql &= "Where Buyer = '" & xBuyerList & "' "
                        sql &= "  And Substring(Season, 3,2) >= '" & xSeasonYY & "' "
                        sql &= "  And Item = '" & dt_ZIPITEM.Rows(i).Item("Item") & "' "
                        sql &= "Group by Color "
                        sql &= "Order by Color "
                        Dim dt_ZIPCOLOR As DataTable = oDataBase.GetDataTable(sql)
                        For j As Integer = 0 To dt_ZIPCOLOR.Rows.Count - 1
                            ' 取得完成品SLIDER兼用色
                            oWaves.GetChangeColora("01", dt_ZIPITEM.Rows(i).Item("Item"), xItem(ItemIndex), dt_ZIPCOLOR.Rows(j).Item("Color"), xColor)
                            ' 產生 SLIDER材料構成檔
                            sql = "Select * From A_MaterialsAnalysis_ZIP "
                            sql &= "Where Buyer = '" & xBuyerList & "' "
                            sql &= "  And Substring(Season, 3,2) >= '" & xSeasonYY & "' "
                            sql &= "  And Item = '" & dt_ZIPITEM.Rows(i).Item("Item") & "' "
                            sql &= "  And Color = '" & dt_ZIPCOLOR.Rows(j).Item("Color") & "' "
                            sql &= "Order by Buyer, Season, Month, Version, CustCode, CustItem, CustColor, Item, Color "
                            Dim dt_ZIPDATA As DataTable = oDataBase.GetDataTable(sql)
                            For k As Integer = 0 To dt_ZIPDATA.Rows.Count - 1
                                '
                                sql = "Insert into A_MaterialsAnalysis_SLD "
                                sql &= "( "
                                sql &= "DataType, Buyer, Season, CustCode, Month, "
                                sql &= "Version, CustItem, CustColor, RubberColor, ParentItem, ParentColor, SeqNo, "
                                sql &= "Item, Color, ItemName1, ItemName2, ItemName3, "
                                sql &= "Import, Produce, Finish, Yobi1, Yobi2, Yobi3, "
                                sql &= "FCTQty, ACTQty, "
                                sql &= "CreateUser, CreateTime, ModifyUser, ModifyTime "
                                sql &= " ) "
                                sql &= "VALUES( "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("DataType").ToString & "', "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("Buyer").ToString & "', "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("Season").ToString & "', "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("CustCode").ToString & "', "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("Month").ToString & "', "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("Version").ToString & "', "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("CustItem").ToString & "', "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("CustColor").ToString & "', "
                                ' 設定 RubberColor
                                ' 當 ItemIndex=1 時為上拉頭 採用 CUSTCOLOR 第 1, 2 COLOR
                                ' 當 ItemIndex=2 時為下拉頭 採用 CUSTCOLOR 第 3, 4 COLOR
                                If dt_ZIPDATA.Rows(k).Item("CustColor").ToString <> "" Then
                                    idx = 0
                                    xUpRubberColor = ""
                                    xLoRubberColor = ""
                                    xRubberColorList = Split(dt_ZIPDATA.Rows(k).Item("CustColor").ToString + "-", "-")
                                    For Each xStr As String In xRubberColorList
                                        Select Case idx
                                            Case 0
                                                xUpRubberColor = xStr
                                            Case 1
                                                If Len(xStr) > 1 Then xUpRubberColor = xUpRubberColor + "-" + xStr
                                            Case 2
                                                xLoRubberColor = xStr
                                            Case 3
                                                If Len(xStr) > 1 Then xLoRubberColor = xLoRubberColor + "-" + xStr
                                            Case Else
                                        End Select
                                        idx = idx + 1
                                    Next
                                    If ItemIndex = 1 Then
                                        sql &= " '" & xUpRubberColor & "', "
                                    Else
                                        sql &= " '" & xLoRubberColor & "', "
                                    End If
                                Else
                                    sql &= " '" & "" & "', "
                                End If
                                '
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("Item").ToString & "', "
                                sql &= " '" & dt_ZIPDATA.Rows(k).Item("Color").ToString & "', "
                                sql &= " '" & "0" + CStr(ItemIndex) & "', "

                                sql &= " '" & xItem(ItemIndex) & "', "
                                sql &= " '" & xColor & "', "
                                oWaves.GetItemName001a(xItem(ItemIndex), 1, xName)
                                sql &= " '" & xName & "', "
                                oWaves.GetItemName001a(xItem(ItemIndex), 2, xName)
                                sql &= " '" & xName & "', "
                                oWaves.GetItemName001a(xItem(ItemIndex), 3, xName)
                                sql &= " '" & xName & "', "

                                oWaves.GetItemProdType(xItem(ItemIndex), xSales, xPurchase, xProduction)
                                sql &= " '" & xPurchase & "', "
                                sql &= " '" & xProduction & "', "
                                oWaves.GetItemSliderFinish(xItem(ItemIndex), xFinish)
                                sql &= " '" & xFinish & "', "
                                sql &= " '" & "" & "', "
                                sql &= " '" & "" & "', "
                                sql &= " '" & "" & "', "

                                sql &= " " & CStr(Fix(CInt(xQtyMeter) / 100 * CDbl(dt_ZIPDATA.Rows(k).Item("FCTQty")) + 0.99)) & ", "
                                sql &= " " & CStr(Fix(CInt(xQtyMeter) / 100 * CDbl(dt_ZIPDATA.Rows(k).Item("ACTQty")) + 0.99)) & ", "
                                sql &= " '" & "FCTACT-BOM" & "', "
                                sql &= " '" & NowDateTime & "', "
                                sql &= " '" & "" & "', "
                                sql &= " Null "
                                sql &= " ) "
                                oDataBase.ExecuteNonQuery(sql)
                            Next
                        Next
                    Next    ' SLD
                    ' ----------------------------------------------------------------
                    ' Trace Log
                    ' ----------------------------------------------------------------
                    sql = "Insert into W_TraceLog "
                    sql &= "( Pgm, Data, CreateTime ) "
                    sql &= "VALUES( "
                    sql &= " '" & "FAS-Material" & "', "
                    sql &= " '" & xBuyerList & "SLD=" & dt_ZIPITEM.Rows(i).Item("Item") & "', "
                    sql &= " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                    sql &= " ) "
                    oDataBase.ExecuteNonQuery(sql)
                    '
                    ' 更新 ZIP-處理FLAG(Yobi1)
                    sql = "Update A_MaterialsAnalysis_ZIP Set "
                    sql &= "Yobi1 = '" & "*" & "', "
                    sql &= "ModifyUser = '" & "BOM" & "', "
                    sql &= "ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                    sql &= "Where Buyer = '" & xBuyerList & "' "
                    sql &= "  And Substring(Season, 3,2) >= '" & xSeasonYY & "' "
                    sql &= "  And Item = '" & dt_ZIPITEM.Rows(i).Item("Item") & "' "
                    oDataBase.ExecuteNonQuery(sql)
                Next
                '
                sql = "Insert into W_TraceLog "
                sql &= "( Pgm, Data, CreateTime ) "
                sql &= "VALUES( "
                sql &= " '" & "FAS-Material" & "', "
                sql &= " '" & xBuyerList & "--- End ---" & "', "
                sql &= " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " ) "
                oDataBase.ExecuteNonQuery(sql)
                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                ' 更新系統設定
                '
                sql = "Update M_Referp Set "
                sql = sql + " Data = '" & Run & "/" & Now.ToString("yyyyMMddHHmmss") & "' "
                sql = sql + " Where Cat = '200' "
                sql = sql + "   And DKey = 'AGENT-MATERIAL' "
                oDataBase.ExecuteNonQuery(sql)
                '
                sql = "Update M_Referp Set "
                sql = sql + " Data = '" & xBuyerList & "/" & "N" & "' "
                sql = sql + " Where Cat = '200' "
                sql = sql + "   And DKey = 'MATERIAL-BUYER' "
                sql = sql + "   And Data Like '" & xBuyerList & "%' "
                oDataBase.ExecuteNonQuery(sql)
            End If
        Catch ex As Exception
            ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            ' Error Log
            sql = "Insert into W_TraceLog "
            sql &= "( Pgm, Data, CreateTime ) "
            Sql &= "VALUES( "
            sql &= " '" & "FAS-Material" & "', "
            sql &= " '" & xBuyerList & "--- Error ---" & "', "
            Sql &= " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
            Sql &= " ) "
            oDataBase.ExecuteNonQuery(Sql)
            ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            ' 異常後重新可處理  處理FLAG = Y
            sql = "Update M_Referp Set "
            sql = sql + " Data = '" & "Y" & "/" & Now.ToString("yyyyMMddHHmmss") & "error" & "' "
            sql = sql & " Where Cat  = '" & "200" & "' "
            sql = sql & "   And DKey = '" & "AGENT-MATERIAL" & "' "
            oDataBase.ExecuteNonQuery(sql)
        End Try
    End Sub
    'MaterialExpansion-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([850]-CanRunVendorFCFinal)
    '**     判斷是否可執行Vendor FC最終確定
    '***********************************************************************************************
    'CanRunVendorFCFinal-Start
    <WebMethod()> _
    Public Function CanRunVendorFCFinal(ByVal pBuyer As String, ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql, xLFBULSNo, xBULSNo As String
        Try
            xLFBULSNo = ""
            xBULSNo = ""
            ' 取得 LF-BuyerLocalStockPlan
            sql = "SELECT Top 1 BULSNo FROM LF_BuyerLocalStockPlan "
            sql &= "Where Buyer = '" & "BULS" & "' "
            sql &= "  And BuyerGroup = '" & pGRBuyer & "' "
            sql &= "Order by BULSNo Desc "
            Dim dt_LFBULSPlan As DataTable = oDataBase.GetDataTable(sql)
            If dt_LFBULSPlan.Rows.Count > 0 Then
                xLFBULSNo = dt_LFBULSPlan.Rows(0).Item("BULSNo")
            End If
            ' 取得 BuyerLocalStockPlan
            sql = "SELECT Top 1 BULSNo FROM BuyerLocalStockPlan "
            sql &= "Where Buyer = '" & "BULS" & "' "
            sql &= "  And BuyerGroup = '" & pGRBuyer & "' "
            sql &= "Order by BULSNo Desc "
            Dim dt_BULSPlan As DataTable = oDataBase.GetDataTable(sql)
            If dt_BULSPlan.Rows.Count > 0 Then
                xBULSNo = dt_BULSPlan.Rows(0).Item("BULSNo")
            End If
            ' 
            If xBULSNo <= xLFBULSNo Then RtnCode = 1
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'CanRunVendorFCFinal-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([855]-VendorFCFinal)
    '**     Vendor FC 最終確定
    '***********************************************************************************************
    'VendorFCFinal-Start
    <WebMethod()> _
    Public Function VendorFCFinal(ByVal pLogID As String, _
                                            ByVal pBuyer As String, _
                                            ByVal pUserID As String, _
                                            ByVal pGRBuyer As String, _
                                            ByVal pFunList As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        '
        Try
            ' -----------------------------------------------------------------------------------
            ' Forcast Plan ---> LF Forcast Plan
            ' -----------------------------------------------------------------------------------
            '
            ' Insert LastFinal ForcastPlan
            sql = "Insert into LF_ForcastPlan "
            sql &= "Select *, '" & pLogID & "', '' "
            sql &= "From ForcastPlan "
            sql &= "Where BuyerGroup = '" & pGRBuyer & "' "
            'sql &= "  And Version = " & "99" & " "
            sql &= "Order by FCTNo, FCTSubNo "
            oDataBase.ExecuteNonQuery(sql)
            '
            ' Delete ForcastPlan
            sql = "Delete From ForcastPlan "
            sql &= "Where BuyerGroup = '" & pGRBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
            ' -----------------------------------------------------------------------------------
            ' LocalStock Plan ---> LF LocalStock Plan
            ' -----------------------------------------------------------------------------------
            '
            ' Insert LastFinal LocalStockPlan
            sql = "Insert into LF_LocalStockPlan "
            sql &= "Select *, '" & pLogID & "', '' "
            sql &= "From LocalStockPlan "
            sql &= "Where BuyerGroup = '" & pGRBuyer & "' "
            'sql &= "  And Version = " & "99" & " "
            sql &= "Order by LSNo, LSSubNo "
            oDataBase.ExecuteNonQuery(sql)
            '
            ' Delete LocalStockPlan
            sql = "Delete From LocalStockPlan "
            sql &= "Where BuyerGroup = '" & pGRBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
            ' -----------------------------------------------------------------------------------
            ' BuyerLocalStock Plan ---> LF Buyer LocalStock Plan
            ' -----------------------------------------------------------------------------------
            '
            ' Insert LastFinal BuyerLocalStockPlan
            sql = "Insert into LF_BuyerLocalStockPlan "
            sql &= "Select *, '" & pLogID & "', '' "
            sql &= "From BuyerLocalStockPlan "
            sql &= "Where BuyerGroup = '" & pGRBuyer & "' "
            'sql &= "  And Version = " & "99" & " "
            sql &= "Order by BULSNo, BULSSubNo "
            oDataBase.ExecuteNonQuery(sql)
            '
            ' Delete BuyerLocalStockPlan
            sql = "Delete From BuyerLocalStockPlan "
            sql &= "Where BuyerGroup = '" & pGRBuyer & "' "
            oDataBase.ExecuteNonQuery(sql)
            ' -----------------------------------------------------------------------------------
            ' Update M_Referp Cat=200 DKey=AGENT-FCT2ANA
            ' 觸發Batch程式執行 (FCT分析使用)
            ' -----------------------------------------------------------------------------------
            '
            ' 更新 M_Referp
            sql = "Update M_Referp Set "
            sql = sql + " Data = '" & "Y" & "/" & pLogID & "' "
            sql = sql + " Where Cat = '200' "
            sql = sql + "   And DKey = 'AGENT-FCT2ANA' "
            oDataBase.ExecuteNonQuery(sql)
            '
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "LFLSPLAN", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "LastFinalLocalStockPlan", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'VendorFCFinal-End
    '***********************************************************************************************
    '**([900]-KPIExpansion)
    '**     eKPI Delay Reason 判定
    '***********************************************************************************************
    'KPIExpansion-Start
    <WebMethod()> _
    Public Function KPIExpansion(ByVal pLogID As String, _
                                            ByVal pBuyer As String, _
                                            ByVal pUserID As String, _
                                            ByVal pGRBuyer As String, _
                                            ByVal pFunList As String) As Integer
        Dim xOFCTQty, xQty As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        '
        Try
            ' -----------------------------------------------------------------------------------
            ' 調整 KPI DATA
            ' 
            'MsgBox("Exec sp_KPIDataAdjustment-BEF")
            sql = "Exec sp_KPIDataAdjustment "
            oDataBase.ExecuteNonQuery(sql)
            'MsgBox("Exec sp_KPIDataAdjustment-AFT")
            ' -----------------------------------------------------------------------------------
            ' 判定延遲理由
            ' 
            'MsgBox("Exec sp_KPIDelayReason-BEF")
            sql = "Exec sp_KPIDelayReason "
            oDataBase.ExecuteNonQuery(sql)
            'MsgBox("Exec sp_KPIDelayReason-AFT")
            ' -----------------------------------------------------------------------------------
            ' Overflow FCT展開
            ' 
            'MsgBox("Overflow FCT展開-BEF")
            sql = "SELECT DataKey FROM I_KPIDataAdjustment "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And Version = 99 "
            sql &= "  And A_DelayReason Like '%Production Capacity Problem%' "
            sql &= "Group by DataKey "
            sql &= "Order by DataKey "
            Dim dt_DataKey As DataTable = oDataBase.GetDataTable(sql)
            For i As Integer = 0 To dt_DataKey.Rows.Count - 1
                ' 
                ' 有Capacity Problem --> 取得 Overflow FCT Qty
                'MsgBox("GET Overflow FCT QTY-BEF")
                xOFCTQty = 0
                xQty = 0
                sql = "SELECT Qty FROM V_KPIFCTOverflow "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataKey = '" & dt_DataKey.Rows(i).Item("DataKey") & "' "
                sql &= "Order by DataKey "
                Dim dt_OFCT As DataTable = oDataBase.GetDataTable(sql)
                If dt_OFCT.Rows.Count > 0 Then
                    xOFCTQty = dt_OFCT.Rows(0).Item("Qty")
                End If
                'MsgBox("GET Overflow FCT QTY-AFT")
                ' 
                ' Overflow FCT展開
                If xOFCTQty > 0 Then
                    ' 
                    ' 有Capacity Problem & 有Overflow FCT Qty
                    ' 
                    ' 單筆-更新 Delay Reason
                    'MsgBox("Overflow FCT QTY 分配-BEF")
                    xQty = xOFCTQty
                    sql = "SELECT Unique_ID, ISNULL(CAST(E1 AS INT),0) As ACTQty, ISNULL(CAST(O1 AS INT),0) As SHIPQty FROM I_KPIDataAdjustment "
                    sql &= "Where DataKey = '" & dt_DataKey.Rows(i).Item("DataKey") & "' "
                    sql &= "  And A_DelayReason Like '%Production Capacity Problem%' "
                    sql &= "  And Version = 99 "
                    sql &= "Order by CAST(E1 AS INT) Desc "
                    Dim dt_SalesData1 As DataTable = oDataBase.GetDataTable(sql)
                    For j As Integer = 0 To dt_SalesData1.Rows.Count - 1
                        If xQty > 0 Then
                            'MsgBox(CStr(xQty) & "-" & CStr(dt_SalesData1.Rows(j).Item("ACTQty")))
                            If xQty >= dt_SalesData1.Rows(j).Item("ACTQty") Then
                                'MsgBox("足夠分配-BEF")
                                ' 
                                sql = "Update I_KPIDataAdjustment Set "
                                sql &= "A_DelayReason = '" & "T1 - Unforecast Q" & "' + CHAR(39) + '" & "ty (Pull Forward)" & "', "
                                sql &= "A_DRRemark = A_DRRemark + '" & "/" & "[Overflow FCT]" & "', "
                                sql &= "A_OFRemark = '" & CStr(xQty) & "-" & CStr(dt_SalesData1.Rows(j).Item("ACTQty")) & "', "
                                sql &= "ModifyUser = '" & "KPI-ADJUST" & "', "
                                sql &= "ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                                sql &= "Where Unique_ID = " & CStr(dt_SalesData1.Rows(j).Item("Unique_ID")) & " "
                                oDataBase.ExecuteNonQuery(sql)
                                xQty = xQty - dt_SalesData1.Rows(j).Item("ACTQty")
                                'MsgBox("足夠分配-AFT")
                            Else
                                ' 
                                '  追加-[Overflow FCT]調整後-剩餘QTY
                                'MsgBox("不足分配-ADD-BEF")
                                sql = "Insert into I_KPIDataAdjustment "
                                sql &= "Select " & _
                                       "Version, DataKey, A_OrderType, A_IdealDate, A_IdealDate15, A_FirstETD, A_ETD, " & _
                                       "A_LastETD, A_PackingDate, A_ShipDate, A_ShipQty, A_DelayQty, " & _
                                       "A_PackingQty, A_NotPackingQty, A_Confirm_Pack, A_Ideal_Pack, A_First_Ideal, " & _
                                       "A_Ship_Pack, A_WavesCode, A_FCTCode, A_MDP, A_PKMDP, A_OCP, A_SDP, A_DOL15, A_POCFR_PER, A_LT_PER, A_DelayReason, " & _
                                       "A_DRRemark + '" & "/[Overflow FCT(不足-調整)]" & "', " & _
                                       "'" & "0" & "-" & CStr(dt_SalesData1.Rows(j).Item("ACTQty") - xQty) & "', " & _
                                       "UID, Buyer, A1, B1, C1, D1, " & _
                                       "'" & CStr(dt_SalesData1.Rows(j).Item("ACTQty") - xQty) & "', " & _
                                       "F1, G1, H1, I1, J1, K1, L1, M1, N1, " & _
                                       "'" & CStr(dt_SalesData1.Rows(j).Item("SHIPQty") - xQty) & "', " & _
                                       "P1, Q1, R1, S1, " & _
                                       "T1, U1, V1, W1, X1, Y1, Z1, AA1, AB1, AC1, AD1, AE1, AF1, " & _
                                       "AG1, AH1, AI1, AJ1, AK1, AL1, AM1, AN1, AO1, AP1, AQ1, " & _
                                       "AR1, AS1, AT1, AU1, AV1, AW1, AX1, AY1, AZ1, BA1, BB1, " & _
                                       "BC1, BD1, BE1, BF1, BG1, BH1, BI1, BJ1, BK1, BL1, BM1, " & _
                                       "BN1, BO1, BP1, BQ1, BR1, BS1, BT1, BU1, BV1, BW1, " & _
                                       "CreateUser, CreateTime, " & _
                                       "'" & "KPI-ADJUST" & "', " & _
                                       "'" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                                sql &= "From I_KPIDataAdjustment "
                                sql &= "Where Unique_ID = " & CStr(dt_SalesData1.Rows(j).Item("Unique_ID")) & " "
                                oDataBase.ExecuteNonQuery(sql)
                                'MsgBox("不足分配-ADD-AFT")
                                ' 
                                '  變更-[Overflow FCT]調整後-FCTQTY
                                'MsgBox("不足分配-ADD+CHG-BEF")
                                sql = "Update I_KPIDataAdjustment Set "
                                sql &= "A_DelayReason = '" & "T1 - Unforecast Q" & "' + CHAR(39) + '" & "ty (Pull Forward)" & "', "
                                sql &= "A_DRRemark = A_DRRemark + '" & "/" & "[Overflow FCT(不足-調整)]" & "', "
                                sql &= "A_OFRemark = '" & CStr(xQty) & "-" & CStr(dt_SalesData1.Rows(j).Item("ACTQty")) & "', "
                                sql &= "E1 = '" & CStr(xQty) & "', "
                                sql &= "O1 = '" & CStr(xQty) & "', "
                                sql &= "ModifyUser = '" & "KPI-ADJUST" & "', "
                                sql &= "ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                                sql &= "Where Unique_ID = " & CStr(dt_SalesData1.Rows(j).Item("Unique_ID")) & " "
                                oDataBase.ExecuteNonQuery(sql)
                                '
                                xQty = xQty - dt_SalesData1.Rows(j).Item("ACTQty")
                                'MsgBox("不足分配-ADD+CHG-AFT")
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    'MsgBox("Overflow FCT QTY 分配-AFT")
                    ' 
                    ' 總數量比較 - 取消
                    'sql = "SELECT ISNULL(SUM(CAST(E1 AS INT)),0) As ACTQty FROM I_KPIDataAdjustment "
                    'sql &= "Where DataKey = '" & dt_DataKey.Rows(i).Item("DataKey") & "' "
                    'Dim dt_SalesData As DataTable = oDataBase.GetDataTable(sql)
                    'If dt_SalesData.Rows(0).Item("ACTQty") <= xOFCTQty Then
                    '    ' 
                    '    ' 全資料-更新 Delay Reason
                    '    sql = "Update I_KPIDataAdjustment Set "
                    '    sql &= "A_DelayReason = '" & "T1 - Unforecast Q'ty (Pull Forward)" & "', "
                    '    sql &= "A_DRRemark = A_DRRemark + '" & "/" & "[Overflow FCT]" & "', "
                    '    sql &= "A_OFCTQty = " & CStr(xOFCTQty) & " "
                    '    sql &= "Where DataKey = '" & dt_DataKey.Rows(i).Item("DataKey") & "' "
                    '    oDataBase.ExecuteNonQuery(sql)
                    'Else
                    'End If
                End If
            Next
            ' -----------------------------------------------------------------------------------
            ' 判定延遲理由
            ' 
            'MsgBox("Exec sp_KPIDataRefresh-BEF")
            sql = "Exec sp_KPIDataRefresh "
            oDataBase.ExecuteNonQuery(sql)
            'MsgBox("Exec sp_KPIDataRefresh-AFT")

            'MsgBox("Overflow FCT展開-AFT")
            '
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "eKPI", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "KPIExpansion", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'KPIExpansion-End
    '-----------------------------------------------------------------------------------------------
    'ISOS-JOY
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([400]-ISOS2FAS)
    '**     Forcast Plan & LS Plan 展開 
    '***********************************************************************************************
    'ISOS2FAS-Start
    <WebMethod()> _
    Public Function ISOS2FAS(ByVal pLogID As String, _
                               ByVal pBuyer As String, _
                               ByVal pUserID As String, _
                               ByVal pGRBuyer As String, _
                               ByVal pFunList As String, _
                               ByVal pLastUniqueID As Integer) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            '##
            '
            Dim xCode As Integer = 0
            Dim Sql As String
            '
            Sql = "SELECT * FROM ForcastPlan_ISOS "
            Sql &= "Where Buyer = '" & pBuyer & "' "
            Sql &= "  And ModifyUser = '" & pUserID & "' "
            Sql &= "  And Y_Level = 0 "
            Dim dt_FCT As DataTable = oDataBase.GetDataTable(Sql)
            If dt_FCT.Rows.Count > 0 Then
                xCode = ForcastPlanISOS(pLogID, pBuyer, pUserID, pGRBuyer, pFunList, 0)
                If xCode = 0 Then
                    LocalStockPlanISOS(pLogID, pBuyer, pUserID, pGRBuyer, pFunList)
                End If
            End If
            '
            '##
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "ISOS", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "ISOS2FAS", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'ForcastPlanISOS-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([410]-ForcastPlanISOS)
    '**     ISOS專用 Forcast Plan展開 
    '***********************************************************************************************
    'ForcastPlanISOS-Start
    <WebMethod()> _
    Public Function ForcastPlanISOS(ByVal pLogID As String, _
                                   ByVal pBuyer As String, _
                                   ByVal pUserID As String, _
                                   ByVal pGRBuyer As String, _
                                   ByVal pFunList As String, _
                                   ByVal pLastUniqueID As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        '
        Try
            '##
            'MSGBOX("IN")
            '
            ' 變數定義及設定初值
            oWaves.Timeout = Timeout.Infinite
            Dim xItemProduct, xLTType, xPartType, xClass, xQtyMeter, xColor As String
            Dim xSubNo As Integer
            Dim FCTWrite As Boolean = False
            ' 構成展開專用變數
            Dim xItem(5), xItemName(5), xQty(5) As String
            Dim xItem1(5), xItemName1(5), xQty1(5) As String
            Dim xCount, xCount1 As Integer
            ' SearchItem專用變數
            Dim xSearchItem(10) As String
            ' 過濾箱中取得的資料變數
            Dim xRuleNo(5), xAction(5), xObjectType(5), xObjectProduct(5) As String
            Dim xRuleCount As Integer
            '
            ' 讀取-ForcastPlan Data
            sql = "SELECT C_Season As Season, Y_ItemCode AS Item, Y_Color AS Color, C_Color AS CSTColor, C_ShortenLT AS KeepCode FROM ForcastPlan_ISOS "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And ModifyUser = '" & pUserID & "' "
            sql &= "  And Unique_ID > " & pLastUniqueID & " "
            sql &= "  And Y_Level = 0 "
            sql &= "Group by C_Season, Y_ItemCode, Y_Color, C_Color, C_ShortenLT "
            sql &= "Order by C_Season, Y_ItemCode, Y_Color, C_Color, C_ShortenLT "
            Dim dt_FCTItem As DataTable = oDataBase.GetDataTable(sql)
            For i As Integer = 0 To dt_FCTItem.Rows.Count - 1
                FCTWrite = False
                xSubNo = 2
                ' 決定LT種類(LTType)
                xLTType = "LLT"
                If pBuyer = "FALL-000001" Or pBuyer = "FALL-000016" Then
                    If dt_FCTItem.Rows(i).Item("KeepCode") <> "AD-RB" Then
                        xLTType = "SLT"
                    End If
                End If
                ' 決定CST PULLER COLOR
                Dim xCSTPullerColor As String = dt_FCTItem.Rows(i).Item("CSTColor")
                '
                ' CHAIN 處理 (不需CHECK LINE-LINE ITEM)
                '--------------------------------------------------------------------------------
                ' 取得Meter換算基準 (取得ItemClass)
                oWaves.GetItemClass(dt_FCTItem.Rows(i).Item("Item"), xClass)
                sql = "Select * From M_Referp "
                sql = sql & " Where Cat  = '" & "110" & "' "
                sql = sql & "   And DKey = '" & pBuyer + "-" + xClass & "' "
                Dim dt_Referp As DataTable = oDataBase.GetDataTable(sql)
                If dt_Referp.Rows.Count > 0 Then
                    xQtyMeter = CStr(CDbl(dt_Referp.Rows(0).Item("Data")) * 100)
                Else
                    xQtyMeter = "100"
                End If
                ' 取得完成品CHAIN(CF/VF) 或 GAP-CHAIN(MF)     (MF=金屬/CF=樹脂/VF=塑鋼)
                xPartType = "CH"                                                            ' 決定材料種類(PartTye)
                oWaves.GetItemProduct(dt_FCTItem.Rows(i).Item("Item"), xItemProduct)        ' 取得製品區分 
                '
                ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)
                If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) = 1 Then
                    If xItemProduct = "MF" Then
                        oWaves.GetChildItemStructure("01", "CH-GAP", dt_FCTItem.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                    Else
                        oWaves.GetChildItemStructure("01", "CH-DYED", dt_FCTItem.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                    End If
                Else
                    xItem(1) = dt_FCTItem.Rows(i).Item("Item")
                    xItemName(1) = ""
                    xQty(1) = CStr(1 * 10000000)
                    xCount = 1
                    For j As Integer = 2 To 5
                        xItem(j) = ""
                        xItemName(j) = ""
                        xQty(j) = "0"
                    Next
                End If
                ' 展開ITEM構成取得所指定ITEM
                If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) <> 3 Then
                    For ItemIndex As Integer = 1 To xCount
                        ' 1. 取得完成品CHAIN兼用色
                        oWaves.GetChangeColor("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCTItem.Rows(i).Item("Color"), xColor)
                        ' 2. 準備備料基準相關SearchItem
                        oWaves.GetSearchItem(xPartType, xItem(ItemIndex), xCSTPullerColor, xSearchItem)
                        ' 3. 尋找適合備料基準-從過濾箱中(M_LSOrderRule)取得適合備料基準
                        GetFilterRuleNo(pBuyer, xLTType, xPartType, xSearchItem, xRuleNo, xAction, xObjectType, xObjectProduct, xRuleCount)
                        ' 4. 依備料基準展開結構
                        For RuleIndex As Integer = 1 To xRuleCount
                            FCTWrite = False
                            ' 第一階結構資料是否OK
                            If InStr(xItemName(ItemIndex), xObjectProduct(RuleIndex)) > 0 Then
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan_ISOS "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    '
                                    WriteNewFCTPlanISOS(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "CH", xQtyMeter, pUserID, xRuleNo(RuleIndex), pLogID)
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                            ' 搜尋下一階結構資料
                            If Not FCTWrite Then
                                '
                                oWaves.GetChildItemStructure("01", xObjectType(RuleIndex) + "-" + xObjectProduct(RuleIndex), xItem(ItemIndex), xItem1, xItemName1, xQty1, xCount1)
                                '
                                For ItemIndex1 As Integer = 1 To xCount1
                                    If InStr(xItemName1(ItemIndex1), xObjectProduct(RuleIndex)) > 0 Then
                                        ' 寫入 ForcastPlan
                                        sql = "SELECT * FROM ForcastPlan_ISOS "
                                        sql &= "Where Buyer = '" & pBuyer & "' "
                                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                        sql &= "  And Y_Level = 0 "
                                        sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                        sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                        sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                        sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                        sql &= "Order by FCTNo, FCTSubNo "
                                        Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                        For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                            '
                                            WriteNewFCTPlanISOS(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, xItem(ItemIndex), xItem1(ItemIndex1), xColor, "CH", xQtyMeter, pUserID, xRuleNo(RuleIndex), pLogID)
                                        Next
                                        xSubNo = xSubNo + 1
                                        FCTWrite = True
                                    End If
                                Next
                            End If
                            ' 搜尋不到任何結構資料(Set Error Rule [9900-990-99])
                            If Not FCTWrite Then
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan_ISOS "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    '
                                    WriteNewFCTPlanISOS(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "CH", xQtyMeter, pUserID, xRuleNo(RuleIndex) + "/" + "9900-990-00", pLogID)
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                        Next
                    Next
                End If
                '##
                'MSGBOX("SLD")
                '
                ' SLIDER 處理  (需CHECK LINE-LINE ITEM)
                '--------------------------------------------------------------------------------
                xQtyMeter = "100"                       ' 取得Meter換算基準 (取得ItemClass)
                xPartType = "SLD"                       ' 決定材料種類(PartTye)
                '
                ' 判斷是否為ZIPPER 0=無, 1=ZIPPER, 2=CH, 3=SLD)
                If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) = 1 Then
                    oWaves.GetChildItemStructure("01", "SLD-FINISH", dt_FCTItem.Rows(i).Item("Item"), xItem, xItemName, xQty, xCount)
                Else
                    xItem(1) = dt_FCTItem.Rows(i).Item("Item")
                    xItemName(1) = ""
                    xQty(1) = CStr(1 * 10000000)
                    xCount = 1
                    For j As Integer = 2 To 5
                        xItem(j) = ""
                        xItemName(j) = ""
                        xQty(j) = "0"
                    Next
                End If
                ' 展開ITEM構成取得所指定ITEM
                If GetZipper(pBuyer, dt_FCTItem.Rows(i).Item("Item")) <> 2 Then
                    For ItemIndex As Integer = 1 To xCount
                        ' 1. 取得完成品SLD兼用色
                        oWaves.GetChangeColor("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCTItem.Rows(i).Item("Color"), xColor)
                        ' 2. 準備備料基準相關SearchItem

                        '##
                        'MsgBox(xPartType & "-" & xItem(ItemIndex))

                        oWaves.GetSearchItem(xPartType, xItem(ItemIndex), xCSTPullerColor, xSearchItem)
                        '##
                        'MsgBox(xSearchItem(1) & "-" & xSearchItem(2) & "-" & xSearchItem(3) & "-" & xSearchItem(4) & "-" & xSearchItem(5) & "-" & xSearchItem(6) & "-" & xSearchItem(7) & "-" & xSearchItem(8) & "-" & xSearchItem(9) & "-" & xSearchItem(10))

                        ' 3. 尋找適合備料基準-從過濾箱中(M_LSOrderRule)取得適合備料基準
                        GetFilterRuleNo(pBuyer, xLTType, xPartType, xSearchItem, xRuleNo, xAction, xObjectType, xObjectProduct, xRuleCount)

                        ' 4. 依備料基準展開結構
                        For RuleIndex As Integer = 1 To xRuleCount
                            '##
                            'MsgBox(xRuleNo(RuleIndex))

                            FCTWrite = False
                            ' 第一階結構資料是否OK
                            If (xObjectProduct(RuleIndex) = "FINISH" Or _
                               (xObjectProduct(RuleIndex) = "E" And InStr(xItemName(ItemIndex), xObjectProduct(RuleIndex)) > 0)) And _
                               oWaves.GetItemLineLine("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex)) = 0 Then         ' 不是LINE-LINE ITEM ((0=不是/1=是)
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan_ISOS "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    WriteNewFCTPlanISOS(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "PS", xQtyMeter, pUserID, xRuleNo(RuleIndex), pLogID)
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                            ' 搜尋下一階結構資料
                            If Not FCTWrite Then
                                '##
                                'MSGBOX("[" & xItem(ItemIndex) & "][" & xObjectType(RuleIndex) + "-" + xObjectProduct(RuleIndex) & "]")

                                oWaves.GetChildItemStructure("01", xObjectType(RuleIndex) + "-" + xObjectProduct(RuleIndex), xItem(ItemIndex), xItem1, xItemName1, xQty1, xCount1)

                                '##
                                'MSGBOX("[" & xItem1(1) & "][" & xItemName1(1) & "]")

                                '
                                For ItemIndex1 As Integer = 1 To xCount1
                                    If oWaves.GetItemLineLine("01", dt_FCTItem.Rows(i).Item("Item"), xItem1(ItemIndex1)) = 0 Then          ' 不是LINE-LINE ITEM ((0=不是/1=是)
                                        ' 寫入 ForcastPlan
                                        sql = "SELECT * FROM ForcastPlan_ISOS "
                                        sql &= "Where Buyer = '" & pBuyer & "' "
                                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                        sql &= "  And Y_Level = 0 "
                                        sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                        sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                        sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                        sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                        sql &= "Order by FCTNo, FCTSubNo "
                                        Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                        For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                            WriteNewFCTPlanISOS(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, xItem(ItemIndex), xItem1(ItemIndex1), xColor, "PS", xQtyMeter, pUserID, xRuleNo(RuleIndex), pLogID)
                                        Next
                                        xSubNo = xSubNo + 1
                                        FCTWrite = True
                                    End If
                                Next
                            End If
                            ' 搜尋不到任何結構資料(Set Error Rule [9900-990-99])
                            If Not FCTWrite Then
                                ' 寫入 ForcastPlan
                                sql = "SELECT * FROM ForcastPlan_ISOS "
                                sql &= "Where Buyer = '" & pBuyer & "' "
                                sql &= "  And Unique_ID > " & pLastUniqueID & " "
                                sql &= "  And Y_Level = 0 "
                                sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                                sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                                sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                                sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                                sql &= "Order by FCTNo, FCTSubNo "
                                Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                                For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                                    WriteNewFCTPlanISOS(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCT2.Rows(j).Item("Y_Color").ToString, "PS", xQtyMeter, pUserID, xRuleNo(RuleIndex) + "/9900-990-00", pLogID)
                                Next
                                xSubNo = xSubNo + 1
                                FCTWrite = True
                            End If
                        Next
                    Next
                End If
                ' ZIP是否為採購品  (0=不是/1=是)
                If Not FCTWrite Then
                    If oWaves.GetPurchaseItem(dt_FCTItem.Rows(i).Item("Item")) = 1 Then
                        ' 寫入 ForcastPlan
                        sql = "SELECT * FROM ForcastPlan_ISOS "
                        sql &= "Where Buyer = '" & pBuyer & "' "
                        sql &= "  And Unique_ID > " & pLastUniqueID & " "
                        sql &= "  And Y_Level = 0 "
                        sql &= "  And C_Season    = '" & dt_FCTItem.Rows(i).Item("Season") & "' "
                        sql &= "  And Y_ItemCode  = '" & dt_FCTItem.Rows(i).Item("Item") & "' "
                        sql &= "  And C_Color     = '" & dt_FCTItem.Rows(i).Item("CSTColor") & "' "
                        sql &= "  And C_ShortenLT = '" & dt_FCTItem.Rows(i).Item("KeepCode") & "' "
                        sql &= "Order by FCTNo, FCTSubNo "
                        Dim dt_FCT2 As DataTable = oDataBase.GetDataTable(sql)
                        For j As Integer = 0 To dt_FCT2.Rows.Count - 1
                            WriteNewFCTPlanISOS(pBuyer, dt_FCT2.Rows(j).Item("Unique_ID"), xSubNo, 1, dt_FCTItem.Rows(i).Item("Item"), dt_FCTItem.Rows(i).Item("Item"), dt_FCT2.Rows(j).Item("Y_Color").ToString, "ZIP", 100, pUserID, "1000-010-00", pLogID)
                        Next
                        xSubNo = xSubNo + 1
                        FCTWrite = True
                    End If
                End If
            Next
            '
            '##
            'MSGBOX("OUT")
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "FCTPLAN", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "ForcastPlanISOS", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'ForcastPlanISOS-End
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([410]-WriteNewFCTPlanISOS)
    '**     Write Forcast Plan ISOS
    '***********************************************************************************************
    'WriteFCTPlanISOS-Start
    <WebMethod()> _
    Public Function WriteNewFCTPlanISOS(ByVal pBuyer As String, _
                                   ByVal pID As Integer, _
                                   ByVal pSubNo As Integer, _
                                   ByVal pLevel As Integer, _
                                   ByVal pFatherItem As String, _
                                   ByVal pItem As String, _
                                   ByVal pColor As String, _
                                   ByVal pClass As String, _
                                   ByVal pQty As String, _
                                   ByVal pUserID As String, _
                                   ByVal pRuleNo As String, _
                                   ByVal pLogID As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql, xsql, str As String
        Dim Qty, QtyTotal As Double
        '
        Try
            oWaves.Timeout = Timeout.Infinite

            sql = "SELECT * FROM ForcastPlan_ISOS "
            sql &= "Where Unique_ID = '" & CStr(pID) & "' "
            Dim dt_FCTData As DataTable = oDataBase.GetDataTable(sql)
            If dt_FCTData.Rows.Count > 0 Then
                '
                sql = "Insert into ForcastPlan_ISOS "
                sql &= "( "
                sql &= "BuyerGroup, "
                sql &= "FCTNo, "
                sql &= "FCTSubNo, "
                sql &= "Version, "
                sql &= "C_Code, "
                sql &= "C_Color, "
                sql &= "C_SpecialRequest, "
                sql &= "C_Season, "
                sql &= "C_ShortenLT, "
                sql &= "C_Style, "
                sql &= "C_A1, "
                sql &= "C_B1, "
                sql &= "C_C1, "
                sql &= "C_D1, "
                sql &= "C_E1, "
                sql &= "C_F1, "
                sql &= "C_G1, "
                sql &= "C_H1, "
                sql &= "C_I1, "
                sql &= "C_J1, "
                sql &= "C_K1, "
                sql &= "C_L1, "
                sql &= "C_M1, "
                sql &= "C_N1, "
                sql &= "C_O1, "

                sql &= "Y_LEVEL, "
                sql &= "Y_ItemCode, "
                sql &= "Y_ItemName1, "
                sql &= "Y_ItemName2, "
                sql &= "Y_ItemName3, "
                sql &= "Y_Color, "
                sql &= "Y_A1, "
                sql &= "Y_B1, "
                sql &= "Y_C1, "
                sql &= "Y_D1, "
                sql &= "Y_E1, "
                sql &= "Y_F1, "
                sql &= "Y_G1, "
                sql &= "Y_H1, "
                sql &= "Y_I1, "
                sql &= "Y_J1, "

                sql &= "N_F, "
                sql &= "N1_F, "
                sql &= "N2_F, "
                sql &= "N3_F, "
                sql &= "N4_F, "
                sql &= "N5_F, "
                sql &= "N6_F, "
                sql &= "N7_F, "
                sql &= "N8_F, "
                sql &= "N9_F, "
                sql &= "N10_F, "
                sql &= "N11_F, "
                sql &= "N12_F, "
                sql &= "Total, "

                sql &= "BUYER, "
                sql &= "ID, "
                sql &= "CreateUser, "
                sql &= "ModifyUser, "
                sql &= "CreateTime "
                sql &= " ) "
                '
                ' 取得 備料MASTER-INF
                ' ------------------------------------------------------
                Dim xLTType, xPartType, xRuleCode, xRuleNo, xRuleSubno, xKeep As String
                Dim xRatio1, xRatio2, xRatio3, xRatio4 As Integer
                ' 決定LT種類(LTType)
                xLTType = "LLT"
                If pBuyer = "FALL-000001" Or pBuyer = "FALL-000016" Then
                    If dt_FCTData.Rows(0).Item("C_ShortenLT").ToString <> "AD-RB" Then
                        xLTType = "SLT"
                    End If
                End If
                ' 決定Part Type
                xPartType = "CH"
                If pClass = "PS" Then xPartType = "SLD" ' PARTTYPE      
                If pClass = "ZIP" Then xPartType = "ZIP" ' PARTTYPE      
                ' 初值
                xRatio1 = 100
                xRatio2 = 100
                xRatio3 = 100
                xRatio4 = 100
                xKeep = ""
                '
                str = pRuleNo   ' RULE-CODE, RULE-N0, RULESUBNO
                If InStr(str, "/") > 0 Then str = Mid(str, InStr(str, "/") + 1)
                xRuleCode = Mid(str, 1, InStr(str, "-") - 1)
                str = Mid(str, InStr(str, "-") + 1)
                xRuleNo = Mid(str, 1, InStr(str, "-") - 1)
                str = Mid(str, InStr(str, "-") + 1)
                xRuleSubno = str
                '
                xsql = "SELECT Action, Ratio1, Ratio2, Ratio3, Ratio4, Keep, "
                xsql &= "str(Ratio1,3,0) + '/' + str(Ratio2,3,0) + '/' + str(Ratio3,3,0) + '/' + str(Ratio4,3,0) + '/' As Ratio, "
                xsql &= "SearchItem1 + '/' + SearchItem2 + '/' + SearchItem3 + '/' +SearchItem4 + '/' + SearchItem5 + '/' + "
                xsql &= "SearchItem6 + '/' + SearchItem7 + '/' + SearchItem8 + '/' +SearchItem9 + '/' + SearchItem10 + '/' As SearchItem "
                xsql &= "FROM M_LSOrderRule "
                xsql &= "Where Active = '1' "
                xsql &= "  And Buyer = '" & pBuyer & "' "
                xsql &= "  And LTType = '" & xLTType & "' "
                xsql &= "  And PartType = '" & xPartType & "' "
                xsql &= "  And RuleCode = '" & xRuleCode & "' "
                xsql &= "  And RuleNo = '" & xRuleNo & "' "
                xsql &= "  And RuleSubno = '" & xRuleSubno & "' "
                xsql &= "Order by RuleCode, RuleNo, RuleSubno "
                Dim dt_LSOrderRule As DataTable = oDataBase.GetDataTable(xsql)
                If dt_LSOrderRule.Rows.Count > 0 Then
                    xRatio1 = dt_LSOrderRule.Rows(0).Item("Ratio1")
                    xRatio2 = dt_LSOrderRule.Rows(0).Item("Ratio2")
                    xRatio3 = dt_LSOrderRule.Rows(0).Item("Ratio3")
                    xRatio4 = dt_LSOrderRule.Rows(0).Item("Ratio4")
                    xKeep = dt_LSOrderRule.Rows(0).Item("Keep").ToString
                End If
                '
                ' 製作 FORECAST
                ' ---------------------------------------------------------
                sql &= "VALUES( "
                sql &= " '" & dt_FCTData.Rows(0).Item("BuyerGroup").ToString & "', "                ' BuyerGroup
                sql &= " '" & dt_FCTData.Rows(0).Item("FCTNo").ToString & "', "                     ' FCTNo
                sql &= " " & CStr(pSubNo) & ", "                                                    ' FCTSubNo
                sql &= " " & dt_FCTData.Rows(0).Item("Version").ToString & ", "                     ' Version
                sql &= " '" & dt_FCTData.Rows(0).Item("C_Code").ToString & "', "                    ' C_Code
                sql &= " '" & dt_FCTData.Rows(0).Item("C_Color").ToString & "', "                   ' C_Color
                sql &= " '" & dt_FCTData.Rows(0).Item("C_SpecialRequest").ToString & "', "          ' C_SpecialRequest
                sql &= " '" & dt_FCTData.Rows(0).Item("C_Season").ToString & "', "                  ' C_Season

                If xKeep = "Y" Then                                                                 ' C_ShortenLT    
                    sql &= " '" & dt_FCTData.Rows(0).Item("C_ShortenLT").ToString & "', "
                Else
                    sql &= " '" & "" & "', "
                End If
                sql &= " '" & dt_FCTData.Rows(0).Item("C_Style").ToString & "', "                   ' C_Style
                sql &= " '" & dt_FCTData.Rows(0).Item("C_A1").ToString & "', "                      ' C_A1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_B1").ToString & "', "                      ' C_B1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_C1").ToString & "', "                      ' C_C1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_D1").ToString & "', "                      ' C_D1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_E1").ToString & "', "                      ' C_E1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_F1").ToString & "', "                      ' C_F1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_G1").ToString & "', "                      ' C_G1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_H1").ToString & "', "                      ' C_H1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_I1").ToString & "', "                      ' C_I1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_J1").ToString & "', "                      ' C_J1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_K1").ToString & "', "                      ' C_K1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_L1").ToString & "', "                      ' C_L1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_M1").ToString & "', "                      ' C_M1
                sql &= " '" & dt_FCTData.Rows(0).Item("C_N1").ToString & "', "                      ' C_N1
                sql &= " '" & pLogID & "', "                      ' C_O1
                ' ------------------------------------------------------------------
                sql &= " " & CStr(pLevel) & ", "                                                    ' Y_LEVEL
                sql &= " '" & pItem & "', "                                                         ' Y_ItemCode
                oWaves.GetItemName001(pItem, 1, str)                                                ' Y_ItemName1
                sql &= " '" & str & "', "
                oWaves.GetItemName001(pItem, 2, str)                                                ' Y_ItemName2
                sql &= " '" & str & "', "
                oWaves.GetItemName001(pItem, 3, str)                                                ' Y_ItemName3
                sql &= " '" & str & "', "

                If xPartType = "ZIP" Then                                                           ' Y_Color(考慮兼用色)(ZIP時不考慮)
                    str = pColor
                Else
                    oWaves.GetChangeColor("01", pFatherItem, pItem, pColor, str)
                End If
                sql &= " '" & str & "', "

                sql &= " '" & pClass & "', "                                                        ' Y_A1 (Item Class)
                If InStr(pRuleNo, "/") > 0 Then                                                     ' 異常
                    sql &= " '" & Mid(pRuleNo, 1, InStr(pRuleNo, "/") - 1) & "', "                  ' Y_B1 (Rule Inf-原LS-RuleNo)
                    sql &= " '" & Mid(pRuleNo, InStr(pRuleNo, "/") + 1) & "', "                     ' Y_C1 (Rule Inf-採用LS-RulenNo)
                Else                                                                                ' 正常
                    sql &= " '" & pRuleNo & "', "
                    sql &= " '" & pRuleNo & "', "
                End If

                If dt_LSOrderRule.Rows.Count > 0 Then
                    sql &= " '" & dt_LSOrderRule.Rows(0).Item("Action").ToString & "', "            ' Y_D1 (Rule Inf-Action)
                    sql &= " '" & dt_LSOrderRule.Rows(0).Item("SearchItem").ToString & "', "        ' Y_E1 (SearchItem)
                    sql &= " '" & dt_LSOrderRule.Rows(0).Item("Keep").ToString & "', "              ' Y_F1 (Keep)
                    sql &= " '" & dt_LSOrderRule.Rows(0).Item("Ratio").ToString & "', "             ' Y_G1 (N+1 ~ N+4)
                Else
                    sql &= " '" & dt_FCTData.Rows(0).Item("Y_D1").ToString & "', "                  ' Y_D1
                    sql &= " '" & dt_FCTData.Rows(0).Item("Y_E1").ToString & "', "                  ' Y_E1
                    sql &= " '" & dt_FCTData.Rows(0).Item("Y_F1").ToString & "', "                  ' Y_F1
                    sql &= " '" & dt_FCTData.Rows(0).Item("Y_G1").ToString & "', "                  ' Y_G1
                End If
                sql &= " '" & dt_FCTData.Rows(0).Item("Y_H1").ToString & "', "                      ' Y_H1
                sql &= " '" & dt_FCTData.Rows(0).Item("Y_I1").ToString & "', "                      ' Y_I1
                sql &= " '" & dt_FCTData.Rows(0).Item("Y_J1").ToString & "', "                      ' Y_J1

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N_F"))                             ' N_F
                QtyTotal = Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N1_F"))                            ' N1_F
                If pClass = "PS" Then
                    Qty = Fix(Qty * (xRatio1 / 100) + 0.9)
                Else
                    Qty = Qty * (xRatio1 / 100)
                End If
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N2_F"))                            ' N2_F
                If pClass = "PS" Then
                    Qty = Fix(Qty * (xRatio2 / 100) + 0.9)
                Else
                    Qty = Qty * (xRatio2 / 100)
                End If
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N3_F"))                            ' N3_F
                If pClass = "PS" Then
                    Qty = Fix(Qty * (xRatio3 / 100) + 0.9)
                Else
                    Qty = Qty * (xRatio3 / 100)
                End If
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N4_F"))                            ' N4_F
                If pClass = "PS" Then
                    Qty = Fix(Qty * (xRatio4 / 100) + 0.9)
                Else
                    Qty = Qty * (xRatio4 / 100)
                End If
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N5_F"))                            ' N5_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N6_F"))                            ' N6_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N7_F"))                            ' N7_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N8_F"))                            ' N8_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N9_F"))                            ' N9_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N10_F"))                           ' N10_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N11_F"))                           ' N11_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                Qty = pQty / 100 * CDbl(dt_FCTData.Rows(0).Item("N12_F"))                           ' N12_F
                QtyTotal = QtyTotal + Qty
                sql &= " " & CStr(Qty) & ", "

                sql &= " " & CStr(QtyTotal) & ", "                                                  ' Total

                sql &= " '" & dt_FCTData.Rows(0).Item("BUYER").ToString & "', "                     ' BUYER
                sql &= " " & dt_FCTData.Rows(0).Item("ID").ToString & ", "                          ' ID
                sql &= " '" & "FCTPlan" & "', "
                sql &= " '" & pUserID & "', "
                sql &= " '" & NowDateTime & "' "
                sql &= " ) "
                oDataBase.ExecuteNonQuery(sql)
                '
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    'WriteFCTPlanISOS-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([420]-LocalStockPlanISOS)    (Step-2)
    '**     ISOS Local Stock Plan展開 
    '***********************************************************************************************
    'LocalStockPlanISOS-Start
    <WebMethod()> _
    Public Function LocalStockPlanISOS(ByVal pLogID As String, _
                                      ByVal pBuyer As String, _
                                      ByVal pUserID As String, _
                                      ByVal pGRBuyer As String, _
                                      ByVal pFunList As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        '
        'Try
        ' ***********************************************************************************
        ' 變數定義及設定初值
        ' ***********************************************************************************
        Dim i, j As Integer
        Dim xLastVersion As Integer = 99    ' 最終版PLAN
        Dim xSystemItem As String = ""
        Dim xGroupItem As String = ""
        Dim xSelectField As String = ""
        Dim xFCTString, xLSString, xDescr As String
        Dim xFCTGroupField(), xLSGroupField() As String
        Dim xKeepCode As String = ""
        Dim xQty As String = "0"
        Dim xSumQty As Double = 0           ' N月生產+在庫合計
        Dim xKeepQty As Double = 0          ' SCHE-PROD + ON-PROD + KEEP-INV
        Dim xFreeQty As Double = 0          ' FREE-INV
        Dim xFCTQty As Double = 0           ' ORDER ZONE FCT QTY
        Dim xFCTSumQty As Double = 0        ' 合計FCT QTY 
        Dim xProdQty As Double = 0          ' 需生產合計
        Dim xMonScheProd(6), xMonOnProd(6) As String        ' N1 ~ N4 SCHE_PRODQTY, ON_PRODQTY
        Dim xPurchaseItem As Boolean        ' 採購品
        Dim xPriceA As String = "0"         ' A單價
        Dim xPriceB As String = "0"         ' B單價
        Dim xPrice As Double = 0            ' 單價
        Dim xMetterLength As Double = 0     ' Length(公呎單位)

        oWaves.Timeout = Timeout.Infinite
        ' ---------------------------------------------------------------------------------
        ' 取得 SystemItem 初值
        sql = "SELECT * FROM M_ImportRule "
        sql &= "Where Buyer = '" & pBuyer & "-2' "
        sql &= "  And Sts   = '1' "
        sql &= "  And DataRule Like '" & "SY_" & "%' "
        sql &= "Order by DataRule "
        Dim dt_List As DataTable = oDataBase.GetDataTable(sql)
        For i = 0 To dt_List.Rows.Count - 1
            ' SystemItem
            If xSystemItem = "" Then
                xSystemItem = dt_List.Rows(i).Item("Field").ToString
            Else
                xSystemItem = xSystemItem & "," & dt_List.Rows(i).Item("Field").ToString
            End If
            ' SelectField
            If xSelectField = "" Then
                xSelectField = dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("Field").ToString
            Else
                xSelectField = xSelectField & ", " & dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("Field").ToString
            End If
        Next
        ' 取得FCT & LS Group Field
        xFCTString = xSystemItem
        xLSString = xSystemItem
        ' ---------------------------------------------------------------------------------
        ' 取得 GroupItem 初值
        dt_List.Clear()
        sql = "SELECT * FROM M_ImportRule "
        sql &= "Where Buyer = '" & pBuyer + "-2" & "' "
        sql &= "  And Sts   = '1' "
        sql &= "  And DataRule Like '" & "GR_" & "%' "
        sql &= "Order by DataRule "
        dt_List = oDataBase.GetDataTable(sql)
        For i = 0 To dt_List.Rows.Count - 1
            ' GroupItem
            If xGroupItem = "" Then
                xGroupItem = dt_List.Rows(i).Item("Field").ToString
            Else
                xGroupItem = xGroupItem & "," & dt_List.Rows(i).Item("Field").ToString
            End If
            ' SelectField
            If xSelectField = "" Then
                xSelectField = dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("DataRule").ToString
            Else
                xSelectField = xSelectField & ", " & dt_List.Rows(i).Item("Field").ToString & " AS " & dt_List.Rows(i).Item("DataRule").ToString
            End If
            ' 取得FCT & LS Group Field
            xFCTString = xFCTString & "," & dt_List.Rows(i).Item("Field").ToString
            xLSString = xLSString & "," & dt_List.Rows(i).Item("DataRule").ToString
        Next
        '
        ' 取得FCT & LS Group Field
        xFCTGroupField = xFCTString.Split(",")
        xLSGroupField = xLSString.Split(",")
        '
        ' GroupItem 不足10個時 ==> SelectField 需補至 10 個
        For j = i + 1 To 10
            If j < 10 Then
                xSelectField = xSelectField & ", " & "'' AS GR_0" & CStr(j)
            Else
                xSelectField = xSelectField & ", " & "'' AS GR_10"
            End If
        Next
        ' ---------------------------------------------------------------------------------
        ' 取得 SummaryItem 初值
        dt_List.Clear()
        sql = "SELECT * FROM M_ImportRule "
        sql &= "Where Buyer = '" & pBuyer + "-2" & "' "
        sql &= "  And Sts   = '1' "
        sql &= "  And DataRule Like '" & "FS_" & "%' "
        sql &= "Order by DataRule "
        dt_List = oDataBase.GetDataTable(sql)
        For i = 0 To dt_List.Rows.Count - 1
            ' SelectField
            If xSelectField = "" Then
                xSelectField = "Sum(" & dt_List.Rows(i).Item("Field").ToString & ") AS " & dt_List.Rows(i).Item("DataRule").ToString
            Else
                xSelectField = xSelectField & ", " & "Sum(" & dt_List.Rows(i).Item("Field").ToString & ") AS " & dt_List.Rows(i).Item("DataRule").ToString
            End If
        Next
        ' ***********************************************************************************
        ' 取得 LSNO = GroupCode + "L" + 年月 + 流水號(4碼)   SA + L + 135 + 0001
        ' 變數定義及設定初值
        ' ***********************************************************************************
        Dim xLSNo, xSeqNoString, xGroup, xItem As String
        Dim xSeqNo, xSubNo As Integer
        xGroup = "ZZ"
        xItem = ""
        ' 取得年月 
        Select Case Month(Now)
            Case 10
                xLSNo = xGroup + "L" + CStr(Year(Now) - 2000) + "A"
            Case 11
                xLSNo = xGroup + "L" + CStr(Year(Now) - 2000) + "B"
            Case 12
                xLSNo = xGroup + "L" + CStr(Year(Now) - 2000) + "C"
            Case Else
                xLSNo = xGroup + "L" + CStr(Year(Now) - 2000) + CStr(Month(Now))
        End Select
        ' 流水號(下一可使用No)及SubNo
        xSeqNo = 1
        xSubNo = 1
        '
        ' ***********************************************************************************
        ' Local Stock Plan展開
        ' ***********************************************************************************
        sql = "SELECT "
        sql &= xSelectField
        sql &= " FROM ForcastPlan_ISOS "
        sql &= " Where Buyer = '" & pBuyer & "' "
        sql &= "   And ModifyUser = '" & pUserID & "' "
        sql &= "   And Version = " & xLastVersion & " "
        sql &= "   And Y_Level > 0 "
        'VENDOR-FC 17/11/24 ADD
        sql &= "   And Y_J1 <> '" & "*" & "' "
        sql &= " Group by " & xSystemItem & ", " & xGroupItem
        sql &= " Order by " & xSystemItem & ", " & xGroupItem
        Dim dt_FCT As DataTable = oDataBase.GetDataTable(sql)
        For i = 0 To dt_FCT.Rows.Count - 1
            ' Insert LocalStock Plan Data


            sql = "Insert into LocalStockPlan_ISOS "
            sql &= "( "
            sql &= "BUYER, "
            sql &= "BUYERGROUP, "
            sql &= "LSNO, "
            sql &= "LSSUBNO, "
            sql &= "VERSION, "

            sql &= "GR_01, "
            sql &= "GR_02, "
            sql &= "GR_03, "
            sql &= "GR_04, "
            sql &= "GR_05, "
            sql &= "GR_06, "
            sql &= "GR_07, "
            sql &= "GR_08, "
            sql &= "GR_09, "
            sql &= "GR_10, "

            sql &= "MINIMUMSTOCK, "
            sql &= "N_SCHEPROD, "
            sql &= "N_ONPROD, "
            sql &= "N_FREEINV, "
            sql &= "N_KEEPINV, "
            sql &= "N_TOTAL, "

            sql &= "SP_00, "        'ADD-START 13/11/11
            sql &= "OP_00, "        'ADD-START 13/11/11
            sql &= "FS_00, "        'ADD-START 13/12/4
            sql &= "PS_00, "        'ADD-START 13/12/4
            sql &= "IS_00, "        'ADD-START 13/12/4
            sql &= "SP_01, "        'ADD-START 13/11/11
            sql &= "OP_01, "        'ADD-START 13/11/11
            sql &= "FS_01, "
            sql &= "PS_01, "
            sql &= "IS_01, "
            sql &= "SP_02, "        'ADD-START 13/11/11
            sql &= "OP_02, "        'ADD-START 13/11/11
            sql &= "FS_02, "
            sql &= "PS_02, "
            sql &= "IS_02, "
            sql &= "SP_03, "        'ADD-START 13/11/11
            sql &= "OP_03, "        'ADD-START 13/11/11
            sql &= "FS_03, "
            sql &= "PS_03, "
            sql &= "IS_03, "
            sql &= "SP_04, "        'ADD-START 13/11/11
            sql &= "OP_04, "        'ADD-START 13/11/11
            sql &= "FS_04, "
            sql &= "PS_04, "
            sql &= "IS_04, "
            sql &= "SP_05, "        'ADD-START 13/11/11
            sql &= "OP_05, "        'ADD-START 13/11/11
            sql &= "FS_05, "
            sql &= "PS_05, "
            sql &= "IS_05, "
            sql &= "FS_06, "
            sql &= "PS_06, "
            sql &= "IS_06, "
            sql &= "FS_07, "
            sql &= "PS_07, "
            sql &= "IS_07, "
            sql &= "FS_08, "
            sql &= "PS_08, "
            sql &= "IS_08, "
            sql &= "FS_09, "
            sql &= "PS_09, "
            sql &= "IS_09, "
            sql &= "FS_10, "
            sql &= "PS_10, "
            sql &= "IS_10, "
            sql &= "FS_11, "
            sql &= "PS_11, "
            sql &= "IS_11, "
            sql &= "FS_12, "
            sql &= "PS_12, "
            sql &= "IS_12, "

            sql &= "FS_13, "
            sql &= "PS_13, "
            sql &= "FREEALLOC, "
            sql &= "DESCRIPTION, "

            sql &= "CreateUser, "
            sql &= "ModifyUser, "
            sql &= "CreateTime "
            sql &= " ) "
            sql &= "VALUES( "
            sql &= " '" & dt_FCT.Rows(i).Item("Buyer").ToString & "', "                 ' Buyer
            sql &= " '" & dt_FCT.Rows(i).Item("BuyerGroup").ToString & "', "            ' BuyerGroup
            ' ---------------------------------------------------------------------------------
            ' 設定 LSNo, SubNo
            ' Item Code是否相同 Counter LSNo, SubNo
            If dt_FCT.Rows(i).Item("GR_03").ToString <> xItem Then
                If i > 0 Then
                    xSeqNo = xSeqNo + 1
                    xSubNo = 1
                End If
                xItem = dt_FCT.Rows(i).Item("GR_03").ToString
            Else
                If i > 0 Then xSubNo = xSubNo + 1
            End If

            ' 製作SeqNo(不足4位補0)
            xSeqNoString = CStr(xSeqNo)
            If Len(xSeqNoString) < 2 Then
                xSeqNoString = "000" + CStr(xSeqNo)
            Else
                If Len(xSeqNoString) < 3 Then
                    xSeqNoString = "00" + CStr(xSeqNo)
                Else
                    If Len(xSeqNoString) < 4 Then
                        xSeqNoString = "0" + CStr(xSeqNo)
                    End If
                End If
            End If

            '設定 LSNo, SubNo
            sql &= " '" & xLSNo + xSeqNoString & "', "                                  ' LSNO
            sql &= " " & CStr(xSubNo) & ", "                                            ' LSSUBNO
            ' ---------------------------------------------------------------------------------
            sql &= " " & CStr(xLastVersion) & ", "                                      ' Version

            sql &= " '" & dt_FCT.Rows(i).Item("GR_01").ToString & "', "                 ' GR_01
            sql &= " '" & dt_FCT.Rows(i).Item("GR_02").ToString & "', "                 ' GR_02
            sql &= " '" & dt_FCT.Rows(i).Item("GR_03").ToString & "', "                 ' GR_03
            sql &= " '" & dt_FCT.Rows(i).Item("GR_04").ToString & "', "                 ' GR_04
            sql &= " '" & dt_FCT.Rows(i).Item("GR_05").ToString & "', "                 ' GR_05
            sql &= " '" & dt_FCT.Rows(i).Item("GR_06").ToString & "', "                 ' GR_06
            sql &= " '" & dt_FCT.Rows(i).Item("GR_07").ToString & "', "                 ' GR_07
            sql &= " '" & dt_FCT.Rows(i).Item("GR_08").ToString & "', "                 ' GR_08
            sql &= " '" & dt_FCT.Rows(i).Item("GR_09").ToString & "', "                 ' GR_09
            sql &= " '" & pLogID & "', "                 ' GR_10
            ' ---------------------------------------------------------------------------------

            If dt_FCT.Rows(i).Item("GR_08").ToString <> "ZIP" Then
                '** PS / CH
                ' ---------------------------------------------------------------------------------
                ' 取得 MinimumStock Qty
                oWaves.GetMininumStock("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xQty)
                If CDbl(xQty) > 0 Then
                    sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                Else
                    sql &= " " & "0" & ", "
                End If
                ' ---------------------------------------------------------------------------------
                ' 歸零作業
                xSumQty = 0
                xProdQty = 0
                xKeepQty = 0            ' SCHE-PROD + ON-PROD + KEEP-INV
                xFreeQty = 0            ' FREE-INV
                xFCTQty = 0             ' ORDER ZONE FCT QTY
                xDescr = ""             ' Free Qty備註說明
                For j = 1 To 6          ' N1~N4 PROD-QTY
                    xMonScheProd(j) = 0
                    xMonOnProd(j) = 0
                Next
                ' 判斷是否採購品(0 = 不是 / 1 = 是)
                xPurchaseItem = False
                If oWaves.GetPurchaseItem(dt_FCT.Rows(i).Item("GR_03").ToString) = 1 Then
                    xPurchaseItem = True
                End If
                ' ---------------------------------------------------------------------------------
                If xPurchaseItem Then
                    ' 採購品
                    ' 不使用 (N_SCHEPROD)
                    sql &= " " & "0" & ", "
                    '
                    ' 採購-ON Qty (N_ONPROD)
                    xKeepCode = dt_FCT.Rows(i).Item("GR_02").ToString

                    oWaves.GetPurchaseQty("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xKeepCode, xQty, xMonOnProd)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                Else
                    ' 製造品
                    ' 生產-SCHE Qty (N_SCHEPROD)
                    xKeepCode = dt_FCT.Rows(i).Item("GR_02").ToString

                    oWaves.GetProductionQty("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xKeepCode, 0, xQty, xMonScheProd)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                    '
                    ' 生產-ON Qty (N_ONPROD)
                    oWaves.GetProductionQty("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xKeepCode, 1, xQty, xMonOnProd)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                End If
                '
                ' 在庫-Free Qty (N_FREEINV)
                oWaves.GetFreeInventory("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xQty)
                If CDbl(xQty) > 0 Then
                    sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                    xFreeQty = xFreeQty + CDbl(xQty) / 10000000
                    xSumQty = xSumQty + CDbl(xQty) / 10000000
                Else
                    sql &= " " & "0" & ", "
                End If
                '
                ' 在庫-Keep Qty (N_KEEPINV)
                oWaves.GetKeepCodeInventory("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xKeepCode, xQty)
                If CDbl(xQty) > 0 Then
                    sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                    xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                    xSumQty = xSumQty + CDbl(xQty) / 10000000
                Else
                    sql &= " " & "0" & ", "
                End If
                '
                ' 合計=(生產-SCHE Qty + 生產-ON Qty + 在庫-Free Qty +在庫-Keep Qty) (N_TOTAL)
                sql &= " " & CStr(xSumQty) & ", "
            Else
                '** VNDOR-FC ADD 17/11/24
                '** ZIP 
                ' ---------------------------------------------------------------------------------
                ' 取得 MinimumStock Qty
                oWaves.GetMininumStockZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xQty)
                If CDbl(xQty) > 0 Then
                    sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                Else
                    sql &= " " & "0" & ", "
                End If
                ' ---------------------------------------------------------------------------------
                ' 歸零作業
                xSumQty = 0
                xProdQty = 0
                xKeepQty = 0            ' SCHE-PROD + ON-PROD + KEEP-INV
                xFreeQty = 0            ' FREE-INV
                xFCTQty = 0             ' ORDER ZONE FCT QTY
                xDescr = ""             ' Free Qty備註說明
                For j = 1 To 6          ' N1~N4 PROD-QTY
                    xMonScheProd(j) = 0
                    xMonOnProd(j) = 0
                Next
                ' 判斷是否採購品(0 = 不是 / 1 = 是)
                xPurchaseItem = False

                If oWaves.GetPurchaseItem(dt_FCT.Rows(i).Item("GR_03").ToString) = 1 Then
                    xPurchaseItem = True
                End If
                ' ---------------------------------------------------------------------------------
                If xPurchaseItem Then
                    ' 採購品
                    ' 不使用 (N_SCHEPROD)
                    sql &= " " & "0" & ", "
                    '
                    ' 採購-ON Qty (N_ONPROD)
                    xKeepCode = dt_FCT.Rows(i).Item("GR_02").ToString
                    oWaves.GetPurchaseQtyZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xKeepCode, xQty, xMonOnProd)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                Else
                    ' 製造品
                    ' 生產-SCHE Qty (N_SCHEPROD)
                    xKeepCode = dt_FCT.Rows(i).Item("GR_02").ToString

                    oWaves.GetProductionQtyZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xKeepCode, 0, xQty, xMonScheProd)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                    '
                    ' 生產-ON Qty (N_ONPROD)
                    oWaves.GetProductionQtyZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xKeepCode, 1, xQty, xMonOnProd)
                    If CDbl(xQty) > 0 Then
                        sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                        xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                        xSumQty = xSumQty + CDbl(xQty) / 10000000
                    Else
                        sql &= " " & "0" & ", "
                    End If
                End If
                '
                ' 在庫-Free Qty (N_FREEINV)
                oWaves.GetFreeInventoryZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xQty)
                If CDbl(xQty) > 0 Then
                    sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                    xFreeQty = xFreeQty + CDbl(xQty) / 10000000
                    xSumQty = xSumQty + CDbl(xQty) / 10000000
                Else
                    sql &= " " & "0" & ", "
                End If
                '
                ' 在庫-Keep Qty (N_KEEPINV)
                oWaves.GetKeepCodeInventoryZIP("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, dt_FCT.Rows(i).Item("GR_09").ToString, dt_FCT.Rows(i).Item("GR_10").ToString, xKeepCode, xQty)
                If CDbl(xQty) > 0 Then
                    sql &= " " & CStr(CDbl(xQty) / 10000000) & ", "
                    xKeepQty = xKeepQty + CDbl(xQty) / 10000000
                    xSumQty = xSumQty + CDbl(xQty) / 10000000
                Else
                    sql &= " " & "0" & ", "
                End If
                '
                ' 合計=(生產-SCHE Qty + 生產-ON Qty + 在庫-Free Qty +在庫-Keep Qty) (N_TOTAL)
                sql &= " " & CStr(xSumQty) & ", "
            End If
            ' ---------------------------------------------------------------------------------
            ' FORCAST & PRODUCTION & INVENTORY
            '
            ' ---------------------------------------------------------------------------------
            ' SP00 / OP00 / FS_00 / PS_00 / IS_00
            ' ** MOD-START 17/12/13 VENDOR FC
            If InStr(pBuyer, "F-VENDOR") > 0 Or InStr(pBuyer, "FALL-VENDOR") > 0 Then
                ' SP00:單價 / OP00:金額
                xPriceA = "0"         ' A單價
                xPriceB = "0"         ' B單價
                oWaves.GetCostPrice(dt_FCT.Rows(i).Item("GR_03").ToString, xPriceA, xPriceB)
                If CDbl(xPriceA) > 0 Or CDbl(xPriceB) > 0 Then
                    If dt_FCT.Rows(i).Item("GR_08").ToString = "ZIP" Then
                        If dt_FCT.Rows(i).Item("GR_10").ToString = "C" Then
                            xMetterLength = CDbl(dt_FCT.Rows(i).Item("GR_09").ToString) / 100
                        Else
                            xMetterLength = CDbl(dt_FCT.Rows(i).Item("GR_09").ToString) * 2.54 / 100
                        End If
                        xPrice = xMetterLength * CDbl(xPriceA) + CDbl(xPriceB)
                    Else
                        xPrice = CDbl(xPriceA) + CDbl(xPriceB)
                    End If
                    sql &= " " & CStr(xPrice) & ", "
                    sql &= " " & CStr(xKeepQty * xPrice) & ", "
                Else
                    sql &= " " & "0" & ", "
                    sql &= " " & "0" & ", "
                End If
            Else
                sql &= " " & CStr(CDbl(xMonScheProd(1)) / 10000000) & ", "
                sql &= " " & CStr(CDbl(xMonOnProd(1)) / 10000000) & ", "
            End If
            ' --
            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_00")
            ' 
            ' --ERIC&PEGGY 追加 2020/09/23
            xFCTQty = xFCTQty + dt_FCT.Rows(i).Item("FS_00")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_00").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' SP01 / OP01 / FS_01 / PS_01 / IS_01
            sql &= " " & CStr(CDbl(xMonScheProd(2)) / 10000000) & ", "
            sql &= " " & CStr(CDbl(xMonOnProd(2)) / 10000000) & ", "

            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_01")
            xFCTQty = xFCTQty + dt_FCT.Rows(i).Item("FS_01")

            sql &= " " & CStr(dt_FCT.Rows(i).Item("FS_01")) & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
                xFCTQty = xFCTQty - xSumQty * -1
            End If
            '
            ' SP02 / OP02 / FS_02 / PS_02 / IS_02
            sql &= " " & CStr(CDbl(xMonScheProd(3)) / 10000000) & ", "
            sql &= " " & CStr(CDbl(xMonOnProd(3)) / 10000000) & ", "

            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_02")
            xFCTQty = xFCTQty + dt_FCT.Rows(i).Item("FS_02")
            '
            sql &= " " & CStr(dt_FCT.Rows(i).Item("FS_02")) & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
                xFCTQty = xFCTQty - xSumQty * -1
            End If
            '
            ' SP03 / OP03 / FS_03 / PS_03 / IS_03
            sql &= " " & CStr(CDbl(xMonScheProd(4)) / 10000000) & ", "
            sql &= " " & CStr(CDbl(xMonOnProd(4)) / 10000000) & ", "

            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_03")
            xFCTQty = xFCTQty + dt_FCT.Rows(i).Item("FS_03")
            '
            sql &= " " & CStr(dt_FCT.Rows(i).Item("FS_03")) & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
                xFCTQty = xFCTQty - xSumQty * -1
            End If
            '
            ' SP04 / OP04 / FS_04 / PS_04 / IS_04
            sql &= " " & CStr(CDbl(xMonScheProd(5)) / 10000000) & ", "
            sql &= " " & CStr(CDbl(xMonOnProd(5)) / 10000000) & ", "

            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_04")
            xFCTQty = xFCTQty + dt_FCT.Rows(i).Item("FS_04")
            '
            sql &= " " & CStr(dt_FCT.Rows(i).Item("FS_04")) & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
                xFCTQty = xFCTQty - xSumQty * -1
            End If
            '
            ' SP05 / OP05 / FS_05 / PS_05 / IS_05
            sql &= " " & CStr(CDbl(xMonScheProd(6)) / 10000000) & ", "
            sql &= " " & CStr(CDbl(xMonOnProd(6)) / 10000000) & ", "

            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_05")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_05").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' FS_06 / PS_06 / IS_06
            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_06")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_06").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' FS_07 / PS_07 / IS_07
            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_07")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_07").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' FS_08 / PS_08 / IS_08
            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_08")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_08").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' FS_09 / PS_09 / IS_09
            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_09")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_09").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' FS_10 / PS_10 / IS_10
            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_10")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_10").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' FS_11 / PS_11 / IS_11
            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_11")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_11").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' FS_12 / PS_12 / IS_12
            xSumQty = xSumQty - dt_FCT.Rows(i).Item("FS_12")
            '
            sql &= " " & dt_FCT.Rows(i).Item("FS_12").ToString & ", "
            If xSumQty >= 0 Then
                sql &= " " & "0" & ", "
                sql &= " " & CStr(xSumQty) & ", "
            Else
                sql &= " " & CStr(xSumQty * -1) & ", "
                sql &= " " & "0" & ", "
                xProdQty = xProdQty + (xSumQty * -1)
                xSumQty = 0
            End If
            '
            ' FS_13 / PS_13
            sql &= " " & dt_FCT.Rows(i).Item("FS_13").ToString & ", "                   ' FS_13
            sql &= " " & CStr(xProdQty) & ", "                                          ' PS_13

            ' FREEALLOC
            If xFCTQty <= xKeepQty Then
                sql &= " " & "0" & ", "
            Else
                If xFCTQty - xKeepQty >= xFreeQty Then
                    sql &= " " & CStr(xFreeQty) & ", "
                Else
                    sql &= " " & CStr(xFCTQty - xKeepQty) & ", "
                End If
            End If
            '
            ' Description
            '  ** Free Qty
            If xFreeQty > 0 Then
                oWaves.GetFreeByLocation("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xDescr)
            End If
            '  ** Free PROD Qty
            oWaves.GetFreeProductionQty("01", dt_FCT.Rows(i).Item("GR_03").ToString, dt_FCT.Rows(i).Item("GR_07").ToString, xQty)
            If CDbl(xQty) > 0 Then
                If xDescr <> "" Then
                    xDescr = xDescr + " * Free on Pro:[" + CStr(CDbl(xQty) / 10000000) + "]"
                Else
                    xDescr = "Free on Pro:[" + CStr(CDbl(xQty) / 10000000) + "]"
                End If
            End If
            'VENDIR FC-17112
            If dt_FCT.Rows(i).Item("GR_09").ToString <> "" And dt_FCT.Rows(i).Item("GR_10").ToString <> "" Then
                If xDescr <> "" Then
                    xDescr = "Length & Unit:[" + dt_FCT.Rows(i).Item("GR_09").ToString + "/" + dt_FCT.Rows(i).Item("GR_10").ToString + "]" + xDescr
                Else
                    xDescr = "Length & Unit:[" + dt_FCT.Rows(i).Item("GR_09").ToString + "/" + dt_FCT.Rows(i).Item("GR_10").ToString + "]"
                End If
            End If
            '  ** 
            If xDescr <> "" Then
                sql &= " '" & xDescr & "', "
            Else
                sql &= " '" & "" & "', "
            End If
            '-----------------------------
            '
            sql &= " '" & "LSPlan" & "', "
            sql &= " '" & pUserID & "', "
            sql &= " '" & NowDateTime & "' "
            sql &= " ) "
            oDataBase.ExecuteNonQuery(sql)
            '
            ' ---------------------------------------------------------------------------------
            ' Update FCT Plan Data (LSNo, LSSubNo)
            sql = "Update ForcastPlan_ISOS Set "
            sql &= "LSNO = '" & xLSNo + xSeqNoString & "', "
            sql &= "LSSUBNO = " & CStr(xSubNo) & " "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And ModifyUser = '" & pUserID & "' "
            sql &= "  And Version = " & xLastVersion & " "
            sql &= "  And Y_Level > 0 "
            For j = 0 To UBound(xFCTGroupField)
                sql &= " And " & xFCTGroupField(j) & " = '" & dt_FCT.Rows(i).Item(xLSGroupField(j)).ToString & "' "
            Next
            oDataBase.ExecuteNonQuery(sql)
        Next

        '' ***********************************************************************************
        '' 更新下一個可使用No
        '' ***********************************************************************************
        '' 製作SeqNo(不足4位補0)
        'xSeqNo = xSeqNo + 1
        'xSeqNoString = CStr(xSeqNo)
        'If Len(xSeqNoString) < 2 Then
        '    xSeqNoString = "000" + CStr(xSeqNo)
        'Else
        '    If Len(xSeqNoString) < 3 Then
        '        xSeqNoString = "00" + CStr(xSeqNo)
        '    Else
        '        If Len(xSeqNoString) < 4 Then
        '            xSeqNoString = "0" + CStr(xSeqNo)
        '        End If
        '    End If
        'End If
        '' 更新下一次可使用PONO
        'UpdateNextLSNo(pBuyer, xLSNo + xSeqNoString)
        ''
        '' ***********************************************************************************
        '' 更新VENDOR FC (VENDOR FC 限用)
        '' ***********************************************************************************
        'If InStr(pBuyer, "F-VENDOR") > 0 Then
        '    oWavesEDI.UpdateVendorFC(pBuyer, pUserID)
        'End If
        ''Catch ex As Exception
        ''RtnCode = 9
        ''oDB.AccessLog(pLogID, pBuyer, "LSPLAN", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "LocalStockPlan", pUserID, "")
        ''End Try
        '
        Return RtnCode
    End Function
    'NewLocalStockPlan-End


    '***********************************************************************************************


End Class