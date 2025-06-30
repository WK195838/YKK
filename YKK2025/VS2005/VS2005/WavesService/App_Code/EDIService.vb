Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'** Database
'***************************************************************************************************
'Database-Start
Imports System.Data             'SQL
'Database-End

<WebService(Namespace:="http://tempuri.org/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class EDIService
     Inherits System.Web.Services.WebService

    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '** 全域變數
    '***********************************************************************************************
    '全域變數-Start
    Dim oDB As New ForProject
    Dim oDataBase As Utility.DataBase = oDB.GetDataBaseObj()
    Dim oWaves As New CommonService
    Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")     ' 現在日時               
    '-----------------------------------------------------------------------------------------------
    '[000]-Service				    				                                |  
    '	+									                                        |
    '	+----[100]-EDICheckKeepCode   						                        | EDI Check Keep Code
    '	+									                                        |
    '	+----[110]-EDICheckColorCode   						                        | EDI Check Color Code
    '	+									                                        |
    '	+----[120]-EDICheckItemCode   						                        | EDI Check Item Code
    '	+									                                        |
    '	+----[130]-EDICheckDuplicateData 				                            | EDI Check Duplicate Data
    '	+									                                        |
    '	+----[140]-EDIEDI2Waves          				                            | EDI Data to Waves System
    '	+									                                        |
    '	+----[700]-UpdateVendorFC(V-FAS)		                                    | 更新VENDOR FC (0=存在/1=不存在) 
    '	+									                                        |
    '	+----									                                    |
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([100]-EDICheckKeepCode)
    '**     EDI Check Keep Code 
    '***********************************************************************************************
    'EDICheckKeepCode-Start
    <WebMethod()> _
    Public Function EDICheckKeepCode(ByVal pLogID As String, _
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
            ' S5K00-檢測
            Row = 1
            sql = "SELECT CORN5K, GRPC5K, AKPC5K FROM S5K00 "
            sql &= "Where Buyer   = '" & pBuyer & "' "
            ' 
            ' Delete-Start 2014/6/23   鼎暉-TUMI
            ' LS Order時追加條件
            'If Mid(pBuyer, 1, 3) = "FCT" Then
            'sql &= "  And AKPC5K <> '" & "" & "' "
            'End If
            ' Delete-End
            ' 
            sql &= "Order by Unique_ID "
            Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_S5K00.Rows.Count - 1
                ' ActionLog-Key欄
                eCORN5K = dt_S5K00.Rows(i).Item("CORN5K").ToString
                eGRPC5K = dt_S5K00.Rows(i).Item("GRPC5K")

                ' [Keep Code檢測]
                If Not err Then
                    ' 
                    ' Delete-Start 2014/6/23   鼎暉-TUMI
                    'If dt_S5K00.Rows(i).Item("AKPC5K") <> "" Then
                    ' Delete-End
                    '
                    RtnCode = oWaves.CheckKeepCode(dt_S5K00.Rows(i).Item("AKPC5K"))
                    If RtnCode <> 0 Then
                        err = True
                        msg = "Check-KeepCode:[KeepCode (AKPC5K=" + dt_S5K00.Rows(i).Item("AKPC5K") + ") !]"
                    End If
                    ' 
                    ' Delete-Start 2014/6/23   鼎暉-TUMI
                    'End If
                    ' Delete-End
                    '
                End If
                ' [檢測異常處理]
                If err Then
                    oDB.AccessLog(pLogID, pBuyer, "CheckKeepCode", NowDateTime, 1, oDB.GetAccessLogLine(pLogID, pBuyer, eCORN5K, eGRPC5K), msg, "=", CStr(RtnCode), "", pUserID, "")
                    Exit For
                Else
                    Row = Row + 1
                End If
            Next
            '
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "CheckKeepCode", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'EDICheckKeepCode-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([110]-EDICheckColorCode)
    '**     EDI Check Color Code 
    '***********************************************************************************************
    'EDICheckColorCode-Start
    <WebMethod()> _
    Public Function EDICheckColorCode(ByVal pLogID As String, _
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
            ' S5K00-檢測
            Row = 1
            sql = "SELECT CORN5K, GRPC5K, CLRC5K FROM S5K00 "
            sql &= "Where Buyer = '" & pBuyer & "' "
            ' 
            ' Delete-Start 2014/6/23   鼎暉-TUMI
            ' LS Order時追加條件
            'If Mid(pBuyer, 1, 3) = "FCT" Then
            sql &= "  And CLRC5K <> '" & "" & "' "
            'End If
            ' Delete-End
            '
            sql &= "Order by Unique_ID "
            Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_S5K00.Rows.Count - 1
                ' ActionLog-Key欄
                eCORN5K = dt_S5K00.Rows(i).Item("CORN5K").ToString
                eGRPC5K = dt_S5K00.Rows(i).Item("GRPC5K")

                ' [Color Code檢測-1]
                If Not err Then
                    If Len(dt_S5K00.Rows(i).Item("CLRC5K")) > 5 Then
                        err = True
                        msg = "Check-ColorCode:[Color Code Length > 5 !][" + CStr(i) + "]"
                        RtnCode = 1
                    End If
                End If
                ' [Color Code檢測-2]
                If Not err Then
                    RtnCode = oWaves.CheckColorCode(dt_S5K00.Rows(i).Item("CLRC5K"))
                    If RtnCode <> 0 Then
                        err = True
                        msg = "Check-ColorCode[Color Not Found (Color=" + dt_S5K00.Rows(i).Item("CLRC5K") + ") !][" + CStr(i) + "]"
                    End If
                End If
                ' [檢測異常處理]
                If err Then
                    oDB.AccessLog(pLogID, pBuyer, "CheckColorCode", NowDateTime, 1, oDB.GetAccessLogLine(pLogID, pBuyer, eCORN5K, eGRPC5K), msg, "=", CStr(RtnCode), "", pUserID, "")
                    Exit For
                Else
                    Row = Row + 1
                End If
            Next
            '
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "CheckColorCode", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'EDICheckColorCode-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([120]-EDICheckItemCode)
    '**     EDI Check Item Code 
    '***********************************************************************************************
    'EDICheckItemCode-Start
    <WebMethod()> _
    Public Function EDICheckItemCode(ByVal pLogID As String, _
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
            ' S5K00-檢測
            Row = 1
            sql = "SELECT CORN5K, GRPC5K, ITMC5K, RCMC5K FROM S5K00 "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "Order by Unique_ID "
            Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_S5K00.Rows.Count - 1
                ' ActionLog-Key欄
                eCORN5K = dt_S5K00.Rows(i).Item("CORN5K").ToString
                eGRPC5K = dt_S5K00.Rows(i).Item("GRPC5K")
                '
                ' PFAS -[Item Code檢測]-START 230826
                ' [PFAS Item Code檢測]
                If Not err Then
                    RtnCode = oWaves.CheckPFASItem(dt_S5K00.Rows(i).Item("ITMC5K"), Mid(dt_S5K00.Rows(i).Item("RCMC5K"), 1, 6))
                    If RtnCode <> 0 Then
                        err = True
                        msg = "Check-ItemCode[Item Not Found or PFAS Error (Item/Buyer=" + dt_S5K00.Rows(i).Item("ITMC5K") + "/" + Mid(dt_S5K00.Rows(i).Item("RCMC5K"), 1, 6) + ") !]"
                    End If
                End If
                '
                '' [Item Code檢測]
                'If Not err Then
                '    RtnCode = oWaves.CheckItemCode(dt_S5K00.Rows(i).Item("ITMC5K"))
                '    If RtnCode <> 0 Then
                '        err = True
                '        msg = "Check-ItemCode[Item Not Found (Item=" + dt_S5K00.Rows(i).Item("ITMC5K") + ") !]"
                '    End If
                'End If
                ' PFAS -[Item Code檢測]-END
                ' [檢測異常處理]
                If err Then
                    oDB.AccessLog(pLogID, pBuyer, "CheckItemCode", NowDateTime, 1, oDB.GetAccessLogLine(pLogID, pBuyer, eCORN5K, eGRPC5K), msg, "=", CStr(RtnCode), "", pUserID, "")
                    Exit For
                Else
                    Row = Row + 1
                End If
            Next
            '
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "CheckItemCode", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'EDICheckItemCode-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([130]-EDICheckDuplicateData)
    '**     EDI Check Duplicate EDI-Data
    '***********************************************************************************************
    'EDICheckDuplicateData-Start
    <WebMethod()> _
    Public Function EDICheckDuplicateData(ByVal pLogID As String, _
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
            Row = 1
            '** Modify-Start by Joy 2012/10/8
            '
            sql = "SELECT CORN5K, GRPC5K FROM S5K00 "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "Order by Unique_ID "
            Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
            '
            For i = 0 To dt_S5K00.Rows.Count - 1
                ' ActionLog-Key欄
                eCORN5K = dt_S5K00.Rows(i).Item("CORN5K").ToString
                eGRPC5K = dt_S5K00.Rows(i).Item("GRPC5K")

                ' [資料重複檢測]
                If Not err Then
                    RtnCode = oWaves.CheckDuplicateData("*", dt_S5K00.Rows(i).Item("CORN5K"))
                    If RtnCode <> 0 Then
                        err = True
                        msg = "Check-DuplicateData[Data Duplication (CORN5K=" + dt_S5K00.Rows(i).Item("CORN5K") + ") !]"
                    End If
                End If
                ' [檢測異常處理]
                If err Then
                    oDB.AccessLog(pLogID, pBuyer, "CheckDuplicateData", NowDateTime, 1, oDB.GetAccessLogLine(pLogID, pBuyer, eCORN5K, eGRPC5K), msg, "=", CStr(RtnCode), "", pUserID, "")
                    Exit For
                Else
                    Row = Row + 1
                End If
            Next
            '
            'If oDB.GetFunctionCode(pGRBuyer, 2) = "P" Then        '製作採購號碼 --> GRBuyer(2)=P 
            '    If oDB.GetFunctionCode(pGRBuyer, 5) = "R" Then    'PUCN5K --> GRBuyer(5)=R (PUCN5K+ITMC5K)
            '        sql = "SELECT PUCN5K FROM S5K00 "
            '        sql &= "Where Buyer = '" & pBuyer & "' "
            '        sql &= "Order by Unique_ID "
            '        Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
            '        '
            '        For i = 0 To dt_S5K00.Rows.Count - 1
            '            ' [資料重複檢測]
            '            If Not err Then
            '                RtnCode = oWaves.CheckDuplicateData("R", dt_S5K00.Rows(i).Item("PUCN5K"))
            '                If RtnCode <> 0 Then
            '                    err = True
            '                    msg = "Check-DuplicateData[Data Duplication (PUCN5K=" + dt_S5K00.Rows(i).Item("PUCN5K") + ") !]"
            '                End If
            '            End If
            '            ' [檢測異常處理]
            '            If err Then
            '                oDB.AccessLog(pLogID, pBuyer, "CheckDuplicateData", NowDateTime, 1, CStr(Row), msg, "=", CStr(RtnCode), "", pUserID, "")
            '                Exit For
            '            Else
            '                Row = Row + 1
            '            End If
            '        Next
            '    Else
            '        sql = "SELECT ORFN5K FROM S5K00 "
            '        sql &= "Where Buyer = '" & pBuyer & "' "
            '        sql &= "Order by Unique_ID "
            '        Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
            '        For i = 0 To dt_S5K00.Rows.Count - 1
            '            ' [資料重複檢測]
            '            If Not err Then
            '                RtnCode = oWaves.CheckDuplicateData("P", dt_S5K00.Rows(i).Item("ORFN5K"))
            '                If RtnCode <> 0 Then
            '                    err = True
            '                    msg = "Check-DuplicateData[Data Duplication (ORFN5K=" + dt_S5K00.Rows(i).Item("ORFN5K") + ") !]"
            '                End If
            '            End If
            '            ' [檢測異常處理]
            '            If err Then
            '                oDB.AccessLog(pLogID, pBuyer, "CheckDuplicateData", NowDateTime, 1, CStr(Row), msg, "=", CStr(RtnCode), "", pUserID, "")
            '                Exit For
            '            Else
            '                Row = Row + 1
            '            End If
            '        Next
            '    End If
            'Else
            '    sql = "SELECT PODN5K FROM S5K00 "
            '    sql &= "Where Buyer = '" & pBuyer & "' "
            '    sql &= "Order by Unique_ID "
            '    Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
            '    For i = 0 To dt_S5K00.Rows.Count - 1
            '        ' [資料重複檢測]
            '        If Not err Then
            '            RtnCode = oWaves.CheckDuplicateData("X", dt_S5K00.Rows(i).Item("PODN5K"))
            '            If RtnCode <> 0 Then
            '                err = True
            '                msg = "Check-DuplicateData[Data Duplication (PODN5K=" + dt_S5K00.Rows(i).Item("PODN5K") + ") !]"
            '            End If
            '        End If
            '        ' [檢測異常處理]
            '        If err Then
            '            oDB.AccessLog(pLogID, pBuyer, "CheckDuplicateData", NowDateTime, 1, CStr(Row), msg, "=", CStr(RtnCode), "", pUserID, "")
            '            Exit For
            '        Else
            '            Row = Row + 1
            '        End If
            '    Next
            'End If
            '
            '** Modify-End by Joy 2012/10/8
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "CheckDuplicateData", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'EDICheckDuplicateData-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([140]-EDIEDI2Waves)
    '**     EDI Data to Waves System
    '***********************************************************************************************
    'EDIEDI2Waves-Start
    <WebMethod()> _
    Public Function EDIEDI2Waves(ByVal pLogID As String, _
                                 ByVal pBuyer As String, _
                                 ByVal pUserID As String, _
                                 ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim msg As String = ""
        '
        Try
            RtnCode = oWaves.EDI2Waves(pBuyer, pGRBuyer)
            If RtnCode <> 0 Then
                msg = "EDI2Waves-[EDI To WavesData Error !]"
                oDB.AccessLog(pLogID, pBuyer, "EDI2Waves", NowDateTime, 1, "", msg, "=", CStr(RtnCode), "", pUserID, "")
            End If
            '
        Catch ex As Exception
            RtnCode = 9
            oDB.AccessLog(pLogID, pBuyer, "EDI2Waves", NowDateTime, 1, "Failed", "Code", "=", CStr(RtnCode), "", pUserID, "")
        End Try
        '
        Return RtnCode
    End Function
    'EDIEDI2Waves-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([700]-UpdateVendorFC) 更新VENDOR FC DATA (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'UpdateVendorFC-Start
    <WebMethod()> _
    Public Function UpdateVendorFC(ByVal pBuyer As String, ByVal pUserID As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim i As Integer
        '
        Try
            sql = "SELECT GR_02, GR_03, GR_07, GR_09, GR_10, "
            sql &= "SUM(N_ScheProd) AS ScheProd, SUM(N_OnProd) AS OnProd, "
            sql &= "SUM(N_FreeInv) AS FreeInv, SUM(N_KeepInv) AS KeepInv, SUM(SP_00) AS UnitPrice, "
            sql &= "SUM(N_Total)-SUM(N_FreeInv) AS Total, SUM(OP_00) AS AMT "
            sql &= "FROM LocalStockPlan "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "  And Version = 99 "
            sql &= "Group by GR_02, GR_03, GR_07, GR_09, GR_10 "
            sql &= "Order by GR_02, GR_03, GR_07, GR_09, GR_10 "
            Dim dt_LSPlan As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_LSPlan.Rows.Count - 1
                RtnCode = oWaves.UpdateFCData(dt_LSPlan.Rows(i).Item("GR_02").ToString, dt_LSPlan.Rows(i).Item("GR_03").ToString, dt_LSPlan.Rows(i).Item("GR_07").ToString, _
                                              dt_LSPlan.Rows(i).Item("GR_09").ToString, dt_LSPlan.Rows(i).Item("GR_10").ToString, Now.ToString("yyyy/MM"), pUserID, _
                                              dt_LSPlan.Rows(i).Item("ScheProd").ToString, dt_LSPlan.Rows(i).Item("OnProd").ToString, _
                                              dt_LSPlan.Rows(i).Item("FreeInv").ToString, dt_LSPlan.Rows(i).Item("KeepInv").ToString, dt_LSPlan.Rows(i).Item("UnitPrice").ToString, _
                                              dt_LSPlan.Rows(i).Item("Total").ToString, dt_LSPlan.Rows(i).Item("AMT").ToString)
            Next
            '
        Catch ex As Exception
            RtnCode = 9
        End Try
        '
        Return RtnCode
    End Function
    'UpdateVendorFC-End


End Class
