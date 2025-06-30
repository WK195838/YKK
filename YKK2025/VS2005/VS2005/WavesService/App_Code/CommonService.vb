Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'** Database
'***************************************************************************************************
'Database-Start
Imports System.Data             'SQL
Imports System.Data.OleDb
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
    Dim oCommon As New Utility.Common
    Dim ConnectString As String = oCommon.GetAppSetting("WAVESOLEDB")
    Dim ConnectWFS As String = oCommon.GetAppSetting("WFSDB")

    Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")     ' 現在日時    
    '全域變數-End
    '-----------------------------------------------------------------------------------------------
    '[000]-Service				    				                                |  
    '	+									                                        |
    '	+----[100]-GetPartType        						                        | 取得Part Type ((G=環保 / IW=進口防水))
    '	+									                                        |
    '	+----[105]-GetSpecialChain     						                        | 取得特殊CHAIN (FAS專用)
    '	+									                                        |
    '	+----[110]-CheckKeepCode									                | 檢測Keep Code (0=存在/1=不存在)
    '	+									                                        |
    '	+----[120]-CheckColorCode   							                    | 檢測Color Code (0=存在/1=不存在)
    '	+									                                        |
    '	+----[130]-CheckItemCode   							                        | 檢測Item Code (0=存在/1=不存在)
    '	+									                                        |
    '	+----[140]-CheckDuplicateData  						                        | 檢測重覆EDI Data (0=不重覆/1=重覆)
    '	+									                                        |
    '	+----[150]-EDI2Waves             				                            | EDI To Waves System
    '	+									                                        |
    '	+----[160]-GetCustomerList         				                            | Customer資料集 (0=存在/1=不存在)
    '	+									                                        |
    '	+----[170]-GetStandardPrice       				                            | 標準單價 (0=存在/1=不存在)
    '	+									                                        |
    '	+----[171]-GetStandardPrice1(EDI)  				                            | 標準單價 (0=存在/1=不存在)
    '	+									                                        |
    '	+----[180]-GetItemName      				                                | 取得Item Name (0=存在/1=不存在)
    '	+									                                        |
    '	+----[185]-GetItemName001      				                                | 取得Item Name1/2/3 (0=存在/1=不存在)
    '	+									                                        |
    '	+----[186]-GetItemName001a(FAS-統計分析)不考慮NODISPLAY                     | 取得Item Name1/2/3 (0=存在/1=不存在)
    '	+									                                        |
    '	+----[190]-GetItemStatistics   				                                | 取得Item統計區分 (0=存在/1=不存在)
    '	+									                                        |
    '	+----[200]-GetItemSize        				                                | 取得Item Size (0=存在/1=不存在)
    '	+									                                        |
    '	+----[210]-GetItemChain       				                                | 取得Item Chain (0=存在/1=不存在)
    '	+									                                        |
    '	+----[215]-GetItemClass       				                                | 取得Item Class (0=存在/1=不存在)
    '	+									                                        |
    '	+----[215a]-GetItemClassa(FAS-統計分析)不考慮NODISPLAY                      | 取得Item Class (0=存在/1=不存在)
    '	+									                                        |
    '	+----[216]-GetItemTape       				                                | 取得Item Tape (0=存在/1=不存在)
    '	+									                                        |
    '	+----[217]-GetItemSpecialFeature			                                | 取得Item Special Feature (0=存在/1=不存在)
    '	+									                                        |
    '	+----[218]-GetItemProdType(FAS-統計分析)不考慮NODISPLAY                     | 取得Item 銷售/採購/內製 (0=存在/1=不存在)
    '	+									                                        |
    '	+----[219]-GetItemSliderFinish(FAS-統計分析)不考慮NODISPLAY                 | 取得Item Slider Finish Code (0=存在/1=不存在)
    '	+					
    '	+----[220]-GetPriceDiscount    				                                | 取得Price Discount (0=存在/1=不存在)
    '	+									                                        |
    '	+----[221]-GetPriceDiscount1(EDI)			                                | 取得Price Discount (0=存在/1=不存在)
    '	+									                                        |
    '	+----[222]-GetItemQtyUnit    				                                | 取得Item Qty Unit
    '	+									                                        |
    '	+----[230]-GetCustomerBuyerPrice       			                            | Customer Buyer Price List (0=存在/1=不存在)
    '	+									                                        |
    '	+----[240]-GetCustomerPrice       			                                | Customer Price List (0=存在/1=不存在)
    '	+									                                        |
    '	+----[250]-GetBuyerPrice       			                                    | Buyer Price List (0=存在/1=不存在)
    '	+									                                        |
    '	+----[260]-GetOrderPrice    			                                    | Order Price List (0=存在/1=不存在)
    '	+									                                        |
    '	+----[270]-GetOrderPriceDetail    			                                | Order Price Detail (0=存在/1=不存在)
    '	+									                                        |
    '	+----[280]-GetOrderProgress    			                                    | Order Progress (0=存在/1=不存在)
    '	+									                                        |
    '	+----[290]-GetDescriptionMaster			                                    | Description Master (0=存在/1=不存在)
    '	+									                                        |
    '	+----[300]-GetPackingListInf			                                    | Packing List (0=OK/1=NG)
    '	+									                                        |
    '	+----[310]-GetChildItemStructure(FAS)		                                | ItemStructure (0=存在/1=不存在)
    '	+									                                        |
    '	+----[311]-GetChildItemStructurea(FAS-統計分析)不考慮NODISPLAY              | ItemStructure (0=存在/1=不存在)
    '	+									                                        |
    '	+----[320]-GetKeepCodeInventory(FAS)		                                | KeepCodeInventory (0=存在/1=不存在)
    '	+									                                        |
    '	+----[321]-GetKeepCodeInventoryZIP(FAS)		                                | KeepCodeInventory (0=存在/1=不存在)
    '	+									                                        |
    '	+----[322]-GetMaterialKeepCodeInventory(SPS)	                            | Material KeepCodeInventory (0=存在/1=不存在)
    '	+									                                        |
    '	+----[330]-GetFreeInventory(FAS)		                                    | FreeInventory (0=存在/1=不存在)
    '	+									                                        |
    '	+----[331]-GetFreeInventoryZIP(FAS)	                                        | FreeInventory (0=存在/1=不存在)
    '	+									                                        |
    '	+----[333]-GetFreeByLocation(FAS)		                                    | Location別FreeInventory (0=存在/1=不存在)
    '	+									                                        |
    '	+----[335]-GetMininumStock(FAS)		                                        | Mininum Stock (0=存在/1=不存在)
    '	+									                                        |
    '	+----[336]-GetMininumStockZIP(FAS)		                                    | Mininum Stock (0=存在/1=不存在)
    '	+									                                        |
    '	+----[340]-GetProductionQty(FAS)		                                    | ProductionQty (0=存在/1=不存在)
    '	+									                                        |
    '	+----[3401]-GetProductionQtyZIP(FAS)		                                | ProductionQty (0=存在/1=不存在)
    '	+									                                        |
    '	+----[341]-GetFreeProductionQty(FAS)		                                | Free ProductionQty (0=存在/1=不存在)
    '	+									                                        |
    '	+----[343]-GetPurchaseQty(FAS)		                                        | PurchaseQty (0=存在/1=不存在)
    '	+									                                        |
    '	+----[344]-GetPurchaseQtyZIP(FAS)                                           | PurchaseQty (0=存在/1=不存在)
    '	+									                                        |
    '	+----[345]-GetProductionInf(FAS)		                                    | Production Inf. (0=存在/1=不存在)
    '	+									                                        |
    '	+----[347]-GetPurchaseInf(FAS)		                                        | Purchase Inf. (0=存在/1=不存在)
    '	+									                                        |
    '	+----[350]-GetChangeColor(FAS)		                                        | ChangeColor (0=存在/1=不存在)
    '	+									                                        |
    '	+----[351]-GetChangeColora(FAS-統計分析)不考慮NODISPLAY                     | ChangeColor (0=存在/1=不存在)
    '	+									                                        |
    '	+----[360]-GetOutsideProduction(FAS)		                                | 判斷是否內製/外注 (0=內製/1=外注)
    '	+									                                        |
    '	+----[370]-GetPlatingRubberSlider(FAS)		                                | 判斷是否電鍍橡膠拉頭 (0=不是/1=是)
    '	+									                                        |
    '	+----[371]-GetPlatingSlider(FAS)		                                    | 判斷是否電鍍或烤漆拉頭 (0=烤漆/1=電鍍)
    '	+									                                        |
    '	+----[380]-GetItemLineLine(FAS)		                                        | 判斷是否LINE-LINE ITEM (0=不是/1=是)
    '	+									                                        |
    '	+----[390]-GetPurchaseItem(FAS)		                                        | 判斷是否採購品 (0=不是/1=是)
    '	+									                                        |
    '	+----[395]-GetItemSlider(FAS)		                                        | 判斷Slider (DAxxxx, DSxxxxx)
    '	+									                                        |
    '	+----[400]-GetItemProduct(FAS)		                                        | 判斷製品區分 (MF=金屬/CF=樹脂/VF=塑鋼)
    '	+									                                        |
    '	+----[401]-GetItemProducta(FAS)(FAS-統計分析)不考慮NODISPLAY                | 判斷製品區分 (MF=金屬/CF=樹脂/VF=塑鋼)
    '	+									                                        |
    '	+----[410]-GetSearchItem(FAS)		                                        | 取得過濾箱搜尋Key 
    '	+									                                        |
    '	+----[420]-GetSalesManCode		                                            | 取得業務員代碼 (0=存在/1=不存在) 
    '	+									                                        |
    '	+----[430]-GetCostPrice(V-FAS)		                                        | 取得成本單價 (0=存在/1=不存在) 
    '	+									                                        |
    '	+----[700]-UpdateFCData(V-FAS)		                                        | 更新VENDOR FC DATA (0=存在/1=不存在) 
    '	+									                                        |
    '	+----									                                    |
    '	+									                                        |
    '	+----[800]-NewColor2Wings(DTMW)	                                            | 新增DTMW新色至WINGS FA300WK  (0=OK/1=NG) 
    '	+									                                        |
    '	+----[810]-OldColor2Wings(DTMW)	                                            | 新增DTMW舊色至WINGS FA200WK  (0=OK/1=NG) 
    '	+									                                        |
    '	+----[820]-NewColorType2Wings(DTMW)	                                        | 新增DTMW新色至WINGS TP200    (0=OK/1=NG) 
    '	+									                                        |
    '	+----									                                    |
    '	+									                                        |
    '	+----[850]-StockInNew2Wings(棧板在庫-STOCK IN NEW)	                        | STOCK DATA至WINGS TWS830A(0=OK/1=NG) 
    '	+									                                        |
    '	+----[851]-StockOutNew2Wings(棧板在庫-STOCK OUT NEW)	                    | STOCK DATA至WINGS TWS830AO(0=OK/1=NG) 
    '	+									                                        |
    '	+----									                                    |
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([851]-StockOutNew2Wings(棧板在庫-STOCK OUT NEW) / STOCK DATA至WINGS TWS830AO(0=OK/1=NG)
    '**使用 棧板在庫
    '**    
    '***********************************************************************************************
    'StockOutNew2Wings-Start
    <WebMethod()> _
    Public Function StockOutNew2Wings(ByVal pNo As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i As Integer
        Dim sql As String
        '
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim cnW As New OleDbConnection
        Dim cmdW As New OleDbCommand
        Dim ds1 As New DataSet
        '
        Try
            'MsgBox("StockOutNew2Wings-IN")
            cn.ConnectionString = ConnectString
            cmd.Connection = cn
            cn.Open()
            '
            cnW.ConnectionString = ConnectWFS
            cmdW.Connection = cnW
            cnW.Open()
            '
            sql = "SELECT "
            sql &= "a.Itemcode as ITMCXA, a.color as CLRCXA, 0 as WQTYXA, 0 as TRNQXA, '' as QUNCXA, '' as BINCXA, "
            sql &= "b.stockno as WSHCXA, 'C' AS REM2XA, 'WEBUPLOAD' AS UIDCXA, "
            sql &= "CONVERT(CHAR(10),GETDATE(),112) AS RADUXA, "
            sql &= "Replace(Convert(varchar(8),Getdate(),108),':','') AS RADTXA "
            '
            sql &= "From F_StockOutSheet a, F_StockOutSheetdt b "
            sql &= "Where a.No = b.No "
            sql &= "And a.No = '" & pNo & "' "
            sql &= "And b.WingSTS = " & "0" & " "
            sql &= "Order by a.No "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cnW)
            DBAdapter1.Fill(ds1, "wStockOut")
            For i = 0 To ds1.Tables("wStockOut").Rows.Count - 1
                '    
                sql = "INSERT INTO WAVEALIB.TWS830AO "
                sql &= "( "
                sql &= "ITMCXX,LNGVXX,LUNCXX,CLRCXX,WQTYXX, "           'A/S/A/A/S/
                sql &= "OQTYXX,RQTYXX,TRNQXX,QUNCXX,WSHCXX, "           'S/S/S/A/A/
                sql &= "REM1XX,REM2XX,SLCCXX,EMPCXX,UIDCXX, "           'A/A/A/A/A/
                sql &= "DEVCXX,RADUXX,RADTXX,RUPUXX,RUPTXX, "           'A/S/S/S/S/
                sql &= "EMP2XX,RCD1XX,RCT1XX,RCD2XX,RCT2XX,BINCXX "     'A/S/S/S/S/A
                sql &= " ) "
                sql &= "VALUES( "
                sql &= " '" & ds1.Tables("wStockOut").Rows(i).Item("ITMCXA").ToString & "', "
                sql &= "  " & "0" & ", "
                sql &= " '" & "" & "', "
                sql &= " '" & ds1.Tables("wStockOut").Rows(i).Item("CLRCXA").ToString & "', "
                sql &= "  " & ds1.Tables("wStockOut").Rows(i).Item("WQTYXA").ToString & ", "
                '
                sql &= "  " & "0" & ", "
                sql &= "  " & "0" & ", "
                sql &= "  " & ds1.Tables("wStockOut").Rows(i).Item("TRNQXA").ToString & ", "
                sql &= " '" & ds1.Tables("wStockOut").Rows(i).Item("QUNCXA").ToString & "', "
                sql &= " '" & ds1.Tables("wStockOut").Rows(i).Item("WSHCXA").ToString & "', "
                '
                sql &= " '" & "" & "', "
                sql &= " '" & ds1.Tables("wStockOut").Rows(i).Item("REM2XA").ToString & "', "
                sql &= " '" & "" & "', "
                sql &= " '" & "" & "', "
                sql &= " '" & ds1.Tables("wStockOut").Rows(i).Item("UIDCXA").ToString & "', "
                '
                sql &= " '" & "" & "', "
                sql &= "  " & ds1.Tables("wStockOut").Rows(i).Item("RADUXA").ToString & ", "
                sql &= "  " & ds1.Tables("wStockOut").Rows(i).Item("RADTXA").ToString & ", "
                sql &= "  " & "0" & ", "
                sql &= "  " & "0" & ", "
                '
                sql &= " '" & "" & "', "
                sql &= "  " & "0" & ", "
                sql &= "  " & "0" & ", "
                sql &= "  " & "0" & ", "
                sql &= "  " & "0" & ", "
                '
                sql &= " '" & ds1.Tables("wStockOut").Rows(i).Item("BINCXA").ToString & "' "
                sql &= " ) "
                cmd.CommandText = sql
                cmd.ExecuteNonQuery()
                '
                sql = "UPDATE F_StockOutSheetDT SET WingSTS=1 "
                sql &= "WHERE "
                sql &= "    NO = '" & pNo & "' "
                sql &= "AND WingSTS = " & "0" & " "
                sql &= "AND STOCKNO = '" & ds1.Tables("wStockOut").Rows(i).Item("WSHCXA").ToString & "' "
                cmdW.CommandText = sql
                cmdW.ExecuteNonQuery()
            Next
            '
            cn.Close()
            cnW.Close()
            '
            'MsgBox("StockOutNew2Wings-OUT")
        Catch ex As Exception
            RtnCode = 1
            cn.Close()
            cnW.Close()
        End Try
        '
        Return RtnCode
    End Function
    '***********************************************************************************************
    '**([850]-StockInNew2Wings(棧板在庫-STOCK IN NEW) / STOCK DATA至WINGS TWS830A(0=OK/1=NG) 
    '**使用 棧板在庫
    '**    
    '***********************************************************************************************
    'StockInNew2Wings-Start
    <WebMethod()> _
    Public Function StockInNew2Wings(ByVal pStockNo As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i As Integer
        Dim sql As String
        '
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim cnW As New OleDbConnection
        Dim cmdW As New OleDbCommand
        Dim ds1 As New DataSet
        '
        Try
            'MsgBox("StockInNew2Wings-IN")
            cn.ConnectionString = ConnectString
            cmd.Connection = cn
            cn.Open()
            '
            cnW.ConnectionString = ConnectWFS
            cmdW.Connection = cnW
            cnW.Open()
            '
            sql = "SELECT "
            sql &= "*, CONVERT(CHAR(10),GETDATE(),112) AS RADUXA, Replace(Convert(varchar(8),Getdate(),108),':','') AS RADTXA "
            sql &= "FROM F_StockInSheetDT "
            sql &= "Where StockNo = '" & pStockNo & "' "
            sql &= "And WingSTS = " & "0" & " "
            sql &= "Order by CreateTime "
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cnW)
            DBAdapter1.Fill(ds1, "wStockIn")
            For i = 0 To ds1.Tables("wStockIn").Rows.Count - 1
                '    
                'sql = "INSERT INTO WAVEALIB.TWS830A "
                '
                sql = "INSERT INTO SONGOLIB.TWS830A "
                sql &= "( "
                sql &= "ITMCXA,	LNGVXA,	LUNCXA,	CLRCXA,	WQTYXA, "       'A/S/A/A/S/
                sql &= "OQTYXA,	RQTYXA,	TRNQXA,	QUNCXA,	WSHCXA, "       'S/S/S/A/A/
                sql &= "REM1XA,	REM2XA,	SLCCXA,	EMPCXA,	UIDCXA, "       'A/A/A/A/A/
                sql &= "DEVCXA,	RADUXA,	RADTXA,	RUPUXA,	RUPTXA, "       'A/S/S/S/S/
                sql &= "EMP2XA,	RCD1XA,	RCT1XA,	RCD2XA,	RCT2XA, "       'A/S/S/S/S/
                sql &= "BINCXA "                                        'A
                sql &= " ) "
                sql &= "VALUES( "
                sql &= " '" & ds1.Tables("wStockIn").Rows(i).Item("ITEMCODE").ToString & "', "
                sql &= "  " & "0" & ", "
                sql &= " '" & "" & "', "
                sql &= " '" & ds1.Tables("wStockIn").Rows(i).Item("COLOR").ToString & "', "
                sql &= "  " & ds1.Tables("wStockIn").Rows(i).Item("QTY").ToString & ", "
                '
                sql &= "  " & "0" & ", "
                sql &= "  " & "0" & ", "
                sql &= "  " & ds1.Tables("wStockIn").Rows(i).Item("QTY").ToString & ", "
                sql &= " '" & ds1.Tables("wStockIn").Rows(i).Item("UNIT").ToString & "', "
                sql &= " '" & ds1.Tables("wStockIn").Rows(i).Item("STOCKNO").ToString & "', "
                '
                sql &= " '" & "C" & "', "
                sql &= " '" & "" & "', "
                sql &= " '" & "" & "', "
                sql &= " '" & "" & "', "
                sql &= " '" & "WEBUPLOAD" & "', "
                '
                sql &= " '" & "" & "', "
                sql &= "  " & ds1.Tables("wStockIn").Rows(i).Item("RADUXA").ToString & ", "
                sql &= "  " & ds1.Tables("wStockIn").Rows(i).Item("RADTXA").ToString & ", "
                sql &= "  " & "0" & ", "
                sql &= "  " & "0" & ", "
                '
                sql &= " '" & "" & "', "
                sql &= "  " & "0" & ", "
                sql &= "  " & "0" & ", "
                sql &= "  " & "0" & ", "
                sql &= "  " & "0" & ", "
                '
                sql &= " '" & ds1.Tables("wStockIn").Rows(i).Item("BOXNO").ToString & "' "
                sql &= " ) "
                cmd.CommandText = sql
                cmd.ExecuteNonQuery()
                '
                sql = "UPDATE F_StockInSheetDT SET WingSTS=1 "
                sql &= "WHERE "
                sql &= "    NO = '" & ds1.Tables("wStockIn").Rows(i).Item("NO").ToString & "' "
                sql &= "AND WingSTS = " & "0" & " "
                sql &= "AND STOCKNO = '" & pStockNo & "' "
                sql &= "AND ITEMCODE = '" & ds1.Tables("wStockIn").Rows(i).Item("ITEMCODE").ToString & "' "
                sql &= "AND COLOR = '" & ds1.Tables("wStockIn").Rows(i).Item("COLOR").ToString & "' "
                cmdW.CommandText = sql
                cmdW.ExecuteNonQuery()
            Next
            '
            cn.Close()
            cnW.Close()
            '
            'MsgBox("StockInNew2Wings-OUT")
        Catch ex As Exception
            RtnCode = 1
            cn.Close()
            cnW.Close()
        End Try
        '
        Return RtnCode
    End Function
    'StockInNew2Wings-End
    '***********************************************************************************************
    '**([800]-NewColor2Wings 新增DTMW新色至WINGS FA300WK  (0=OK/1=NG) 
    '**使用 1.DTMW
    '**    
    '***********************************************************************************************
    'NewColor2Wings-Start
    <WebMethod()> _
    Public Function NewColor2Wings(ByVal pNo As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i As Integer
        Dim sql As String
        '
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim cnW As New OleDbConnection
        Dim cmdW As New OleDbCommand
        Dim ds1 As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            cmd.Connection = cn
            cn.Open()
            '
            cnW.ConnectionString = ConnectWFS
            cmdW.Connection = cnW
            cnW.Open()
            '
            sql = "SELECT * FROM R_NCFA300WK "
            sql &= "Where No = '" & pNo & "' "
            sql &= "And STS = '" & "0" & "' "
            sql &= "Order by RADUW3, RADTW3 "
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cnW)
            DBAdapter1.Fill(ds1, "wFA300")
            For i = 0 To ds1.Tables("wFA300").Rows.Count - 1
                sql = "INSERT INTO WAVEALIB.FA300WK "
                sql &= "( "
                sql &= "PDPCW3, CTBNW3, PACCW3, ST1CW3, ST6CW3, "       'A/A/A/A/A/
                sql &= "SCSVW3, CCLCW3, UIDCW3, PRGCW3, DEVCW3, "       'S/A/A/A/A/
                sql &= "RADUW3, RADTW3, RUPUW3, RUPTW3, RCMCW3, "       'S(A)/S(A)/S/S/A/
                sql &= "PRBFW3, CHKFW3, WKDTW3 "                        'A/A/S/
                sql &= " ) "
                sql &= "VALUES( "
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("PDPCW3").ToString & "', "
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("CTBNW3").ToString & "', "
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("PACCW3").ToString & "', "
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("ST1CW3").ToString & "', "
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("ST6CW3").ToString & "', "
                '
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("SCSVW3").ToString & "', "
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("CCLCW3").ToString & "', "
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("UIDCW3").ToString & "', "
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("PRGCW3").ToString & "', "
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("DEVCW3").ToString & "', "
                '
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("RADUW3").ToString & "', "
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("RADTW3").ToString & "', "
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("RUPUW3").ToString & "', "
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("RUPTW3").ToString & "', "
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("RCMCW3").ToString & "', "
                '
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("PRBFW3").ToString & "', "
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("CHKFW3").ToString & "', "
                sql &= " '" & ds1.Tables("wFA300").Rows(i).Item("WKDTW3").ToString & "'  "
                sql &= " ) "
                cmd.CommandText = sql
                cmd.ExecuteNonQuery()
                '
                sql = "UPDATE R_NCFA300WK SET STS='1' "
                sql &= "WHERE "
                sql &= "    NO = '" & pNo & "' "
                sql &= "AND STS = '" & "0" & "' "
                sql &= "AND PDPCW3 = '" & ds1.Tables("wFA300").Rows(i).Item("PDPCW3").ToString & "' "
                sql &= "AND CTBNW3 = '" & ds1.Tables("wFA300").Rows(i).Item("CTBNW3").ToString & "' "
                sql &= "AND PACCW3 = '" & ds1.Tables("wFA300").Rows(i).Item("PACCW3").ToString & "' "
                sql &= "AND ST1CW3 = '" & ds1.Tables("wFA300").Rows(i).Item("ST1CW3").ToString & "' "
                sql &= "AND ST6CW3 = '" & ds1.Tables("wFA300").Rows(i).Item("ST6CW3").ToString & "' "
                sql &= "AND SCSVW3 = '" & ds1.Tables("wFA300").Rows(i).Item("SCSVW3").ToString & "' "
                sql &= "AND CCLCW3 = '" & ds1.Tables("wFA300").Rows(i).Item("CCLCW3").ToString & "' "
                cmdW.CommandText = sql
                cmdW.ExecuteNonQuery()
            Next
            '
            cn.Close()
            cnW.Close()
        Catch ex As Exception
            RtnCode = 1
            cn.Close()
            cnW.Close()
        End Try
        '
        Return RtnCode
    End Function
    'NewColor2Wings-End
    '***********************************************************************************************
    '**([810]-OldColor2Wings 新增DTMW新色至WINGS FA200WK  (0=OK/1=NG) 
    '**使用 1.DTMW
    '**    
    '***********************************************************************************************
    'OldColor2Wings-Start
    <WebMethod()> _
    Public Function OldColor2Wings(ByVal pNo As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i As Integer
        Dim sql As String
        '
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim cnW As New OleDbConnection
        Dim cmdW As New OleDbCommand
        Dim ds1 As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            cmd.Connection = cn
            cn.Open()
            '
            cnW.ConnectionString = ConnectWFS
            cmdW.Connection = cnW
            cnW.Open()
            '
            sql = "SELECT * FROM R_NCFA200WK "
            sql &= "Where No = '" & pNo & "' "
            sql &= "And STS = '" & "0" & "' "
            sql &= "Order by RADUA2, RADTA2 "
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cnW)
            DBAdapter1.Fill(ds1, "wFA200")
            For i = 0 To ds1.Tables("wFA200").Rows.Count - 1
                sql = "INSERT INTO WAVEALIB.FA200WX "
                sql &= "( "
                sql &= "PDPCW2,PCCCW2,CLRCW2,SIZCW2,FMLCW2, "       'A/A/A/A/A/
                sql &= "ITPCW2,PCACW2,ST1CW2,ST2CW2,ST3CW2, "       'A/A/A/A/A/
                sql &= "ST4CW2,ST5CW2,ST6CW2,ST7CW2,UIDCW2, "       'A/A/A/A/A/
                sql &= "PRGCW2,DEVCW2,RADUW2,RADTW2,RUPUW2, "       'A/A/S/S/S/
                sql &= "RUPTW2,RCMCW2 "                             'S/A/
                sql &= " ) "
                sql &= "VALUES( "
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("PDPCA2").ToString & "', "
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("PCCCA2").ToString & "', "
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("CLRCA2").ToString & "', "
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("SIZCA2").ToString & "', "
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("FMLCA2").ToString & "', "
                '
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("ITPCA2").ToString & "', "
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("PCACA2").ToString & "', "
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("ST1CA2").ToString & "', "
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("ST2CA2").ToString & "', "
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("ST3CA2").ToString & "', "
                '
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("ST4CA2").ToString & "', "
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("ST5CA2").ToString & "', "
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("ST6CA2").ToString & "', "
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("ST7CA2").ToString & "', "
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("UIDCA2").ToString & "', "
                '
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("PRGCA2").ToString & "', "
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("DEVCA2").ToString & "', "
                sql &= " " & ds1.Tables("wFA200").Rows(i).Item("RADUA2").ToString & ", "
                sql &= " " & ds1.Tables("wFA200").Rows(i).Item("RADTA2").ToString & ", "
                sql &= " " & "0" & ", "
                '
                sql &= " " & "0" & ", "
                sql &= " '" & ds1.Tables("wFA200").Rows(i).Item("RCMCA2").ToString & "' "
                sql &= " ) "
                cmd.CommandText = sql
                cmd.ExecuteNonQuery()
                '
                'joy
                sql = "UPDATE R_NCFA200WK SET STS='1' "
                sql &= "WHERE "
                sql &= "    NO = '" & pNo & "' "
                sql &= "AND STS = '" & "0" & "' "
                sql &= "AND PDPCA2 = '" & ds1.Tables("wFA200").Rows(i).Item("PDPCA2").ToString & "' "
                sql &= "AND PCCCA2 = '" & ds1.Tables("wFA200").Rows(i).Item("PCCCA2").ToString & "' "
                sql &= "AND CLRCA2 = '" & ds1.Tables("wFA200").Rows(i).Item("CLRCA2").ToString & "' "
                sql &= "AND SIZCA2 = '" & ds1.Tables("wFA200").Rows(i).Item("SIZCA2").ToString & "' "
                sql &= "AND FMLCA2 = '" & ds1.Tables("wFA200").Rows(i).Item("FMLCA2").ToString & "' "
                sql &= "AND ITPCA2 = '" & ds1.Tables("wFA200").Rows(i).Item("ITPCA2").ToString & "' "
                cmdW.CommandText = sql
                cmdW.ExecuteNonQuery()
            Next
            '
            cn.Close()
            cnW.Close()
        Catch ex As Exception
            RtnCode = 1
            cn.Close()
            cnW.Close()
        End Try
        '
        Return RtnCode
    End Function
    'NewColor2Wings-End
    '***********************************************************************************************
    '**([820]-NewColorType2Wings 新增DTMW新色至WINGS TP200    (0=OK/1=NG)
    '**使用 1.DTMW
    '**    
    '***********************************************************************************************
    'NewColorType2Wings-Start
    <WebMethod()> _
    Public Function NewColorType2Wings(ByVal pNo As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i As Integer
        Dim sql As String
        '
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim cnW As New OleDbConnection
        Dim cmdW As New OleDbCommand
        Dim ds1 As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            cmd.Connection = cn
            cn.Open()
            '
            cnW.ConnectionString = ConnectWFS
            cmdW.Connection = cnW
            cnW.Open()
            '
            sql = "SELECT * FROM R_NCTP200WK "
            sql &= "Where No = '" & pNo & "' "
            sql &= "And STS = '" & "0" & "' "
            sql &= "Order by RADUP2, RADTP2 "
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cnW)
            DBAdapter1.Fill(ds1, "wTP200")
            For i = 0 To ds1.Tables("wTP200").Rows.Count - 1
                sql = "INSERT INTO WAVEDLIB.TP200 "
                sql &= "( "
                sql &= "CLRCP2,LHTCP2,CNLVP2,CNAVP2,CNBVP2, "       'A/A/S/S/S/
                sql &= "YB1NP2,YB2NP2,YB3NP2,YB4NP2,YB5NP2, "       'S/S/S/S/S/
                sql &= "YB1CP2,YB2CP2,YB3CP2,YB4CP2,YB5CP2, "       'A/A/A/A/A/
                sql &= "YB1FP2,YB2FP2,YB3FP2,YB4FP2,YB5FP2, "       'A/A/A/A/A/
                sql &= "YB1IP2,YB2IP2,YB3IP2,YB4IP2,YB5IP2, "       'A/A/A/A/A/
                sql &= "UIDCP2,PRGCP2,DEVCP2, "                     'A/A/A/
                sql &= "RADUP2,RADTP2,RUPUP2,RUPTP2, "              'S/S/S/S/
                sql &= "RCMCP2 "                                    'A/
                sql &= " ) "
                sql &= "VALUES( "
                sql &= " '" & ds1.Tables("wTP200").Rows(i).Item("CLRCP2").ToString & "', "
                sql &= " '" & "" & "', "
                sql &= " " & "0" & ", "
                sql &= " " & "0" & ", "
                sql &= " " & "0" & ", "
                '
                sql &= " " & "0" & ", "
                sql &= " " & "0" & ", "
                sql &= " " & "0" & ", "
                sql &= " " & "0" & ", "
                sql &= " " & "0" & ", "
                '
                sql &= " '" & ds1.Tables("wTP200").Rows(i).Item("YB1CP2").ToString & "', "
                sql &= " '" & "" & "', "
                sql &= " '" & "" & "', "
                sql &= " '" & "" & "', "
                sql &= " '" & "" & "', "
                '
                sql &= " '" & "" & "', "
                sql &= " '" & "" & "', "
                sql &= " '" & "" & "', "
                sql &= " '" & "" & "', "
                sql &= " '" & "" & "', "
                '
                sql &= " '" & "" & "', "
                sql &= " '" & "" & "', "
                sql &= " '" & "" & "', "
                sql &= " '" & "" & "', "
                sql &= " '" & "" & "', "
                '
                sql &= " '" & ds1.Tables("wTP200").Rows(i).Item("UIDCP2").ToString & "', "
                sql &= " '" & ds1.Tables("wTP200").Rows(i).Item("PRGCP2").ToString & "', "
                sql &= " '" & ds1.Tables("wTP200").Rows(i).Item("DEVCP2").ToString & "', "
                '
                sql &= " " & ds1.Tables("wTP200").Rows(i).Item("RADUP2").ToString & ", "
                sql &= " " & ds1.Tables("wTP200").Rows(i).Item("RADTP2").ToString & ", "
                sql &= " " & ds1.Tables("wTP200").Rows(i).Item("RUPUP2").ToString & ", "
                sql &= " " & ds1.Tables("wTP200").Rows(i).Item("RUPTP2").ToString & ", "
                '
                sql &= " '" & ds1.Tables("wTP200").Rows(i).Item("RCMCP2").ToString & "' "
                '
                sql &= " ) "
                cmd.CommandText = sql
                cmd.ExecuteNonQuery()
                '
                'joy
                sql = "UPDATE R_NCTP200WK SET STS='1' "
                sql &= "WHERE "
                sql &= "    NO = '" & pNo & "' "
                sql &= "AND STS = '" & "0" & "' "
                sql &= "AND CLRCP2 = '" & ds1.Tables("wTP200").Rows(i).Item("CLRCP2").ToString & "' "
                sql &= "AND YB1CP2 =  " & ds1.Tables("wTP200").Rows(i).Item("YB1CP2").ToString & " "
                cmdW.CommandText = sql
                cmdW.ExecuteNonQuery()
            Next
            '
            cn.Close()
            cnW.Close()
        Catch ex As Exception
            RtnCode = 1
            cn.Close()
            cnW.Close()
        End Try
        '
        Return RtnCode
    End Function
    'NewColor2Wings-End
    '***********************************************************************************************
    '**([100]-GetPartType)  取得Part Type (G=環保 / IW=進口防水)
    '**使用 1.EDI System
    '**    
    '***********************************************************************************************
    'GetPartType-Start
    <WebMethod()> _
    Public Function GetPartType(ByVal pCode As String) As String
        Dim RtnString As String = ""
        Dim sql As String

        Dim cn As New OleDbConnection
        Dim ds, ds1, ds2, ds3 As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            '
            'ITEM NAME：[GREEN-F] --> [G]
            sql = "SELECT ITMCA0 FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pCode & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            sql &= "  And IT1IA0||IT2IA0||IT3IA0 LIKE '" & "%GREEN-F%" & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "FA000")
            If ds.Tables("FA000").Rows.Count > 0 Then RtnString = "G"
            '
            'ITEM NAME：[CNT／CFT／CNDT1／CNDT2／CFDT1／CFDT2／FL-MAT／VT8／VT10] --> [IW]
            sql = "SELECT ITMCA0 FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pCode & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            sql &= "  And ( "
            sql &= "    IT1IA0||IT2IA0||IT3IA0 LIKE '" & "%CNT%" & "' "
            sql &= " or IT1IA0||IT2IA0||IT3IA0 LIKE '" & "%CFT%" & "' "
            sql &= " or IT1IA0||IT2IA0||IT3IA0 LIKE '" & "%CNDT1%" & "' "
            sql &= " or IT1IA0||IT2IA0||IT3IA0 LIKE '" & "%CNDT2%" & "' "
            sql &= " or IT1IA0||IT2IA0||IT3IA0 LIKE '" & "%CFDT1%" & "' "
            sql &= " or IT1IA0||IT2IA0||IT3IA0 LIKE '" & "%CFDT2%" & "' "
            sql &= " or IT1IA0||IT2IA0||IT3IA0 LIKE '" & "%CIDT1%" & "' "
            sql &= " or IT1IA0||IT2IA0||IT3IA0 LIKE '" & "%CIDT2%" & "' "
            sql &= " or IT1IA0||IT2IA0||IT3IA0 LIKE '" & "%FL-MAT%" & "' "
            sql &= " or IT1IA0||IT2IA0||IT3IA0 LIKE '" & "%VT8%" & "' "
            sql &= " or IT1IA0||IT2IA0||IT3IA0 LIKE '" & "%VT10%" & "' "
            sql &= "      ) "
            '
            Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
            DBAdapter2.Fill(ds1, "FA000-1")
            If ds1.Tables("FA000-1").Rows.Count > 0 Then
                If RtnString = "" Then
                    RtnString = "IW"
                Else
                    RtnString = RtnString + "/IW"
                End If
            End If
            'If RtnString = "" Then
            'End If
            '
            'ITEM NAME：[RFA／RFB／TY] --> [TSCC]
            sql = "SELECT ITMCA0 FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pCode & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            sql &= "  And ( "
            sql &= "    IT1IA0||IT2IA0||IT3IA0 LIKE '" & "%RFA%" & "' "
            sql &= " or IT1IA0||IT2IA0||IT3IA0 LIKE '" & "%RFB%" & "' "
            sql &= " or IT1IA0||IT2IA0||IT3IA0 LIKE '" & "%TY%" & "' "
            sql &= "      ) "
            '
            Dim DBAdapter3 As New OleDbDataAdapter(sql, cn)
            DBAdapter3.Fill(ds2, "FA000-2")
            If ds2.Tables("FA000-2").Rows.Count > 0 Then
                If RtnString = "" Then
                    RtnString = "TSCC"
                Else
                    RtnString = RtnString + "/TSCC"
                End If
            End If
            'If RtnString = "" Then
            'End If
            '
            'ITEM NAME：[I-AD/ILAD] --> [Q501]
            sql = "SELECT ITMCA0 FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pCode & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            sql &= "  And ( "
            sql &= "    IT1IA0||IT2IA0||IT3IA0 LIKE '" & "%I-AD%" & "' "
            sql &= " or IT1IA0||IT2IA0||IT3IA0 LIKE '" & "%ILAD%" & "' "
            sql &= "      ) "
            '
            Dim DBAdapter4 As New OleDbDataAdapter(sql, cn)
            DBAdapter4.Fill(ds3, "FA000-3")
            If ds3.Tables("FA000-3").Rows.Count > 0 Then
                If RtnString = "" Then
                    RtnString = "Q501"
                Else
                    RtnString = RtnString + "/Q501"
                End If
            End If
            '
            cn.Close()
        Catch ex As Exception
            cn.Close()
        End Try
        '
        Return RtnString
    End Function
    'GetGreeItem-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([105]-GetSpecialChain )  取得特殊CHAIN (FAS專用)
    '**使用 1.FAS System
    '**    
    '***********************************************************************************************
    'GetSpecialChain-Start
    <WebMethod()> _
    Public Function GetSpecialChain(ByVal pCode As String, ByVal pChain As String) As String
        Dim RtnString As String = ""
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds, ds1 As New DataSet
        '
        Try
            ' 取得 ITEM NAME
            Dim xItemName As String = ""
            cn.ConnectionString = ConnectString
            '
            sql = "SELECT IT1IA0||IT2IA0||IT3IA0 AS ITEMNAME FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pCode & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "FA000")
            If ds.Tables("FA000").Rows.Count > 0 Then
                xItemName = ds.Tables(0).Rows(0).Item("ITEMNAME")
            End If
            '
            ' 判斷是否特殊CHAIN
            ' T8/CFT8/CNT8/;T9/CFT9/CNT9/;T10/CFT10/CNT10/;DT1/CFDT1/CNDT1/;DT2/CFDT2/CNDT2/;BP/BP12/BP14/BP15/BP16/BP18/;PBR/PBR12/PBR14/PBR16/;
            If xItemName <> "" Then
                Dim i As Integer
                Dim str As String = pChain
                Dim xChainString As String = ""
                Dim xChain As Object
                '
                ' 解析比較基準
                While InStr(str, ";") > 0
                    xChainString = Mid(str, 1, InStr(str, ";") - 1)
                    xChain = Split(xChainString, "/")
                    '
                    For i = 1 To UBound(xChain) - 1

                        If InStr(xItemName, xChain(i)) > 0 Then
                            '
                            If RtnString = "" Then
                                RtnString = xChain(0)
                            Else
                                RtnString = RtnString + "/" + xChain(0)
                            End If
                        End If
                    Next
                    '
                    str = Mid(str, InStr(str, ";") + 1)
                End While
            End If
            '
            cn.Close()
        Catch ex As Exception
            cn.Close()
        End Try
        '
        Return RtnString
    End Function
    'GetSpecialChain-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([110]-CheckKeepCode)  檢測KeepCode (0=存在/1=不存在)
    '**使用 1.EDI System
    '**    
    '***********************************************************************************************
    'CheckKeepCode-Start
    <WebMethod()> _
    Public Function CheckKeepCode(ByVal pCode As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            sql = "SELECT DGRC09 FROM WAVEDLIB.C0900 "
            sql &= "Where DGRC09 = '" & "KEPC" & "' "
            sql &= "  And DDTC09 = '" & pCode & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "C0900")
            If ds.Tables("C0900").Rows.Count <= 0 Then RtnCode = 1
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'CheckKeepCode-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([120]-CheckColorCode)  檢測Color Code (0=存在/1=不存在)
    '**使用 1.EDI System
    '**    
    '***********************************************************************************************
    'CheckColorCode-Start
    <WebMethod()> _
    Public Function CheckColorCode(ByVal pCode As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            sql = "SELECT CLRCA2 FROM WAVEDLIB.FA200 "
            sql &= "Where CLRCA2 = '" & pCode & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "FA200")
            If ds.Tables("FA200").Rows.Count <= 0 Then RtnCode = 1
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'CheckColorCode-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([130]-CheckPFASItem) 檢測Item Code (0=存在/1=不存在)
    '**使用 1.EDI System
    '**    
    '***********************************************************************************************
    ' PFAS -[Item Code檢測]-START 230826
    'CheckPFASItem-Start
    <WebMethod()> _
    Public Function CheckPFASItem(ByVal pCode As String, ByVal pBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            sql = "SELECT ITMCA0, ST1CA0, IF8CA0 FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pCode & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "FA000")
            If ds.Tables("FA000").Rows.Count > 0 Then
                'ZIPPER
                If ds.Tables("FA000").Rows(0).Item("ST1CA0") = "1" Then
                    If ds.Tables("FA000").Rows(0).Item("IF8CA0") <> "W" Then
                        'BUYER MASTER extention
                        ds.Clear()
                        sql = "SELECT BYRC35 FROM WAVEDLIB.S35E0 "
                        sql &= "Where BYRC35 = '" & pBuyer & "' "
                        sql &= "  And PFCK35 = '" & "1" & "' "
                        '
                        Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
                        DBAdapter2.Fill(ds, "S35E0")
                        If ds.Tables("S35E0").Rows.Count > 0 Then RtnCode = 1
                        'If ds.Tables("S35E0").Rows.Count > 0 Then RtnCode = 0
                        '
                    End If
                End If
                '
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'CheckPFASItem-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([130]-CheckItemCode) 檢測Item Code (0=存在/1=不存在)
    '**使用 1.EDI System
    '**    
    '***********************************************************************************************
    ' PFAS -[Item Code檢測]-START 230826
    'CheckPFASItem-Start
    <WebMethod()> _
    Public Function CheckItemCode(ByVal pCode As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            sql = "SELECT ITMCA0 FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pCode & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "FA000")
            If ds.Tables("FA000").Rows.Count <= 0 Then RtnCode = 1
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'CheckItemCode-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([140]-CheckDuplicateData) 檢測重覆EDI Data (0=不重覆/1=重覆)
    '**使用 1.EDI System
    '**    
    '***********************************************************************************************
    'CheckDuplicateData-Start
    <WebMethod()> _
    Public Function CheckDuplicateData(ByVal pType As String, ByVal pCode As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            '** Modify-Start by Joy 2012/10/8
            '
            sql = "SELECT CORN5K FROM WAVEDLIB.S5K00 "
            sql &= "Where CORN5K = '" & pCode & "' "
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "WS5K00")
            If ds.Tables("WS5K00").Rows.Count > 0 Then RtnCode = 1
            '
            'sql = "SELECT PODN5K FROM WAVEDLIB.S5K00 "
            'If pType = "R" Then
            '    sql &= "Where PUCN5K = '" & pCode & "' "
            'Else
            '    If pType = "P" Then
            '        sql &= "Where ORFN5K = '" & pCode & "' "
            '    Else
            '        sql &= "Where PODN5K = '" & pCode & "' "
            '    End If
            'End If
            ''
            'Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            'DBAdapter1.Fill(ds, "WS5K00")
            'If ds.Tables("WS5K00").Rows.Count > 0 Then RtnCode = 1
            '** Modify-End by Joy 2012/10/8
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'CheckDuplicateData-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([150]-EDI2Waves)
    '**     EDI To Waves System
    '***********************************************************************************************
    'EDI2Waves-Start
    <WebMethod()> _
    Public Function EDI2Waves(ByVal pBuyer As String, ByVal pGRBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim i As Integer
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim ds1, ds2, ds3, ds4 As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            cmd.Connection = cn
            cn.Open()
            '          
            ' JOY 2017/6/20
            sql = "SELECT PODN5K, POSN5K FROM S5K00 "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "Order by PODN5K, POSN5K "
            Dim dt_S5KXX As DataTable = oDataBase.GetDataTable(sql)
            For i = 0 To dt_S5KXX.Rows.Count - 1
                sql = "SELECT PODN5K, POSN5K FROM WAVEDLIB.S5K00 "
                sql &= "Where PODN5K = '" & dt_S5KXX.Rows(i).Item("PODN5K").ToString & "' "
                sql &= "And   POSN5K = '" & dt_S5KXX.Rows(i).Item("POSN5K").ToString & "' "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds1, "S5KXX")
                If ds1.Tables("S5KXX").Rows.Count > 0 Then
                    RtnCode = 8
                    Exit For
                End If
            Next
            '  
            If RtnCode = 0 Then
                ' [S5K00]
                sql = "SELECT * FROM S5K00 "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "Order by PODN5K, POSN5K "
                Dim dt_S5K00 As DataTable = oDataBase.GetDataTable(sql)
                For i = 0 To dt_S5K00.Rows.Count - 1
                    ' JOY 2017/6/20
                    sql = "SELECT PODN5K, POSN5K FROM WAVEDLIB.S5K00 "
                    sql &= "Where PODN5K = '" & dt_S5K00.Rows(i).Item("PODN5K").ToString & "' "
                    sql &= "And   POSN5K = '" & dt_S5K00.Rows(i).Item("POSN5K").ToString & "' "
                    Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
                    DBAdapter2.Fill(ds2, "WS5K00")
                    If ds2.Tables("WS5K00").Rows.Count <= 0 Then
                        sql = "INSERT INTO WAVEDLIB.S5K00 "
                        sql &= "( "
                        sql &= "DPDC5K, OTYC5K, VNDC5K, POSN5K, PODN5K, "
                        sql &= "GRPC5K, STLN5K, PFMC5K, PUCN5K, CORN5K, "
                        sql &= "ORFN5K, SBKN5K, PODU5K, POEU5K, PRDU5K, "
                        sql &= "PRDC5K, RDRV5K, CDLF5K, MCNC5K, MARC5K, "
                        sql &= "TRNC5K, TTRC5K, TCRC5K, IDPR5K, PMTC5K, "
                        sql &= "PTYC5K, PD1V5K, PD2V5K, PD3V5K, US1V5K, "
                        sql &= "US2V5K, US3V5K, PDOC5K, PLTC5K, RDPC5K, PDDF5K, "
                        sql &= "DK1C5K, DK2C5K, RACC5K, RVNC5K, ITMC5K, "
                        sql &= "LUNC5K, LNGV5K, CLRC5K, POOC5K, AKPC5K, "
                        sql &= "ABKN5K, BSBN5K, PODQ5K, PQUC5K, PPVC5K, "
                        sql &= "PRTC5K, PUNP5K, TTYC5K, SUNP5K, DTCX5K, "
                        sql &= "NCMF5K, SMPF5K, USGC5K, RFCC5K, PRBF5K, "
                        sql &= "ERRC5K, UIDC5K, PRGC5K, DEVC5K, RADU5K, "
                        sql &= "RADT5K, RUPU5K, RUPT5K, RCMC5K "
                        sql &= " ) "
                        sql &= "VALUES( "
                        sql &= " '" & dt_S5K00.Rows(i).Item("DPDC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("OTYC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("VNDC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("POSN5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PODN5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("GRPC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("STLN5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PFMC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PUCN5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("CORN5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("ORFN5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("SBKN5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PODU5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("POEU5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PRDU5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PRDC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("RDRV5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("CDLF5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("MCNC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("MARC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("TRNC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("TTRC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("TCRC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("IDPR5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PMTC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PTYC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PD1V5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PD2V5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PD3V5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("US1V5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("US2V5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("US3V5K").ToString & "', "
                        'MODIFY-START  2017/9/12
                        'PDOC中止, 改用PLTC
                        'sql &= " '" & dt_S5K00.Rows(i).Item("PDOC5K").ToString & "', "
                        sql &= " '" & "" & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PDOC5K").ToString & "', "
                        'MODIFY-END
                        sql &= " '" & dt_S5K00.Rows(i).Item("RDPC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PDDF5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("DK1C5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("DK2C5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("RACC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("RVNC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("ITMC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("LUNC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("LNGV5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("CLRC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("POOC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("AKPC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("ABKN5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("BSBN5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PODQ5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PQUC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PPVC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PRTC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PUNP5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("TTYC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("SUNP5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("DTCX5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("NCMF5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("SMPF5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("USGC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("RFCC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PRBF5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("ERRC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("UIDC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("PRGC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("DEVC5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("RADU5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("RADT5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("RUPU5K").ToString & "', "
                        sql &= " '" & dt_S5K00.Rows(i).Item("RUPT5K").ToString & "', "
                        ' MODIFY JOY BY 2017/7/24
                        'sql &= " '" & dt_S5K00.Rows(i).Item("RCMC5K").ToString & "'  "
                        sql &= " '" & Mid(dt_S5K00.Rows(i).Item("RCMC5K").ToString, 1, 6) & "'  "
                        sql &= " ) "
                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    End If
                Next
                '
                ' [S5L00]
                sql = "SELECT * FROM S5L00 "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "Order by PODN5L "
                Dim dt_S5L00 As DataTable = oDataBase.GetDataTable(sql)
                For i = 0 To dt_S5L00.Rows.Count - 1
                    ' JOY 2017/6/20
                    sql = "SELECT PODN5L FROM WAVEDLIB.S5L00 "
                    sql &= "Where PODN5L = '" & dt_S5L00.Rows(i).Item("PODN5L").ToString & "' "
                    Dim DBAdapter3 As New OleDbDataAdapter(sql, cn)
                    DBAdapter3.Fill(ds3, "WS5L00")
                    If ds3.Tables("WS5L00").Rows.Count <= 0 Then
                        sql = "INSERT INTO WAVEDLIB.S5L00 "
                        sql &= "( "
                        sql &= "DPDC5L, PODN5L, PO1X5L, PO2X5L, PO3X5L, "
                        sql &= "PO4X5L, FN1I5L, FN2I5L, DL1A5L, DL2A5L, "
                        sql &= "DL3A5L, FD1I5L, FD2I5L, FD3I5L, FM1X5L, "
                        sql &= "CM1C5L, FM2X5L, CM2C5L, FM3X5L, CM3C5L, "
                        sql &= "FM4X5L, CM4C5L, FM5X5L, CM5C5L, FM6X5L, "
                        sql &= "CM6C5L, FM7X5L, CM7C5L, FM8X5L, CM8C5L, "
                        sql &= "FM9X5L, CM9C5L, FMAX5L, CMAC5L, PCSN5L, "
                        sql &= "RFCC5L, PRBF5L, UIDC5L, PRGC5L, DEVC5L, "
                        sql &= "RADU5L, RADT5L, RUPU5L, RUPT5L, RCMC5L  "
                        sql &= " ) "
                        sql &= "VALUES( "
                        sql &= " '" & dt_S5L00.Rows(i).Item("DPDC5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("PODN5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("PO1X5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("PO2X5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("PO3X5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("PO4X5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("FN1I5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("FN2I5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("DL1A5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("DL2A5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("DL3A5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("FD1I5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("FD2I5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("FD3I5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("FM1X5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("CM1C5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("FM2X5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("CM2C5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("FM3X5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("CM3C5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("FM4X5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("CM4C5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("FM5X5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("CM5C5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("FM6X5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("CM6C5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("FM7X5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("CM7C5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("FM8X5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("CM8C5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("FM9X5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("CM9C5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("FMAX5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("CMAC5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("PCSN5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("RFCC5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("PRBF5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("UIDC5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("PRGC5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("DEVC5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("RADU5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("RADT5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("RUPU5L").ToString & "', "
                        sql &= " '" & dt_S5L00.Rows(i).Item("RUPT5L").ToString & "', "
                        ' MODIFY JOY BY 2017/7/24
                        'sql &= " '" & dt_S5L00.Rows(i).Item("RCMC5L").ToString & "'  "
                        sql &= " '" & Mid(dt_S5L00.Rows(i).Item("RCMC5L").ToString, 1, 6) & "'  "
                        sql &= " ) "
                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    End If
                Next
                '
                ' [SC760W1]
                If oDB.GetFunctionCode(pGRBuyer, 1) = "N" Then        'VDP --> GRBuyer(1)=N (NIKE)
                    sql = "SELECT * FROM SC760W1 "
                    sql &= "Where Buyer = '" & pBuyer & "' "
                    sql &= "Order by ORDNW1, OSBNW1 "
                    Dim dt_SC760W1 As DataTable = oDataBase.GetDataTable(sql)
                    For i = 0 To dt_SC760W1.Rows.Count - 1
                        ' JOY 2017/6/20
                        sql = "SELECT ORDNW1, OSBNW1 FROM WAVEDLIB.SC760W1 "
                        sql &= "Where ORDNW1 = '" & dt_SC760W1.Rows(i).Item("ORDNW1").ToString & "' "
                        sql &= "And   OSBNW1 = '" & dt_SC760W1.Rows(i).Item("OSBNW1").ToString & "' "
                        Dim DBAdapter4 As New OleDbDataAdapter(sql, cn)
                        DBAdapter4.Fill(ds4, "WSC760W")
                        If ds4.Tables("WSC760W").Rows.Count <= 0 Then
                            sql = "INSERT INTO WAVEDLIB.SC760W1 "
                            sql &= "( "
                            sql &= "ORDNW1, OSBNW1, SNCDW1, SNYRW1, BYMHW1, BYMSW1, "
                            sql &= "NFCDW1, NSNOW1, NINOW1, NCCDW1, NCN1W1, "
                            sql &= "NCN2W1, NCN3W1, NCN4W1, REFNW1, CMCDW1, "
                            sql &= "CMMTW1, NOICW1, UIDCW1, PRGCW1, DEVCW1, "
                            sql &= "RADUW1, RADTW1, RUPUW1, RUPTW1, RCMCW1, "
                            sql &= "GRPCW1, "
                            '
                            sql &= "CSTCW1, SMTCW1, ACRNW1, ACR2W1, ACR3W1, "
                            sql &= "ADCDW1, DVSCW1, YBF1W1, YBF2W1, YBF3W1, "
                            sql &= "YBF4W1, YBF5W1, YBD1W1, YBD2W1, YBD3W1 "
                            '
                            sql &= " ) "
                            sql &= "VALUES( "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("ORDNW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("OSBNW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("SNCDW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("SNYRW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("BYMHW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("BYMSW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("NFCDW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("NSNOW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("NINOW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("NCCDW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("NCN1W1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("NCN2W1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("NCN3W1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("NCN4W1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("REFNW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("CMCDW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("CMMTW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("NOICW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("UIDCW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("PRGCW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("DEVCW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("RADUW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("RADTW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("RUPUW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("RUPTW1").ToString & "', "
                            ' MODIFY JOY BY 2017/7/24
                            'sql &= " '" & dt_SC760W1.Rows(i).Item("RCMCW1").ToString & "', "
                            sql &= " '" & Mid(dt_SC760W1.Rows(i).Item("RCMCW1").ToString, 1, 6) & "', "

                            sql &= " '" & dt_SC760W1.Rows(i).Item("GRPCW1").ToString & "', "
                            '
                            sql &= " '" & dt_SC760W1.Rows(i).Item("CSTCW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("SMTCW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("ACRNW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("ACR2W1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("ACR3W1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("ADCDW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("DVSCW1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("YBF1W1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("YBF2W1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("YBF3W1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("YBF4W1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("YBF5W1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("YBD1W1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("YBD2W1").ToString & "', "
                            sql &= " '" & dt_SC760W1.Rows(i).Item("YBD3W1").ToString & "'  "
                            '
                            sql &= " ) "
                            cmd.CommandText = sql
                            cmd.ExecuteNonQuery()
                        End If
                    Next
                End If
                '
                ' JOY 2019/2/21
                ' [PEE50]
                'sql = "SELECT DPDC5K, PODN5K, RFCC5K, VNDC5K, RCMC5K, PUCN5K, CORN5K, ORFN5K  FROM S5K00 "
                'sql &= "Where Buyer = '" & pBuyer & "' "
                'sql &= "  And RFCC5K <> '" & "000034" & "' "
                'sql &= "  And RFCC5K <> '" & "" & "' "
                'sql &= "Group by DPDC5K, PODN5K, RFCC5K, VNDC5K, RCMC5K, PUCN5K, CORN5K, ORFN5K "
                'sql &= "Order by DPDC5K, PODN5K, RFCC5K, VNDC5K, RCMC5K, PUCN5K, CORN5K, ORFN5K "
                'Dim dt_S5KPEE As DataTable = oDataBase.GetDataTable(sql)
                'For i = 0 To dt_S5KPEE.Rows.Count - 1
                '    sql = "SELECT PODNEE, RFCCEE FROM WAVEDLIB.PEE50 "
                '    sql &= "Where PODNEE = '" & dt_S5KPEE.Rows(i).Item("PODN5K").ToString & "' "
                '    sql &= "And   RFCCEE = '" & dt_S5KPEE.Rows(i).Item("RFCC5K").ToString & "' "
                '    Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
                '    DBAdapter2.Fill(ds2, "PEE50")
                '    If ds2.Tables("PEE50").Rows.Count <= 0 Then
                '        sql = "INSERT INTO WAVEDLIB.PEE50 "
                '        sql &= "( "
                '        sql &= "DPDCEE, PODNEE, RFCCEE, VNDCEE, BYRCEE, "
                '        sql &= "MCNCEE, MARCEE, CORNEE, CCNNEE, ORFNEE, "
                '        sql &= "ICSCEE, PO1XEE, PO2XEE, PO3XEE, PO4XEE, "
                '        sql &= "FN1IEE, FN2IEE, DL1AEE, DL2AEE, DL3AEE, "
                '        sql &= "FD1IEE, FD2IEE, FD3IEE, PRBFEE, UIDCEE, "
                '        sql &= "PRGCEE, DEVCEE, RADUEE, RADTEE, RUPUEE, "
                '        sql &= "RUPTEE, RCMCEE "
                '        sql &= " ) "
                '        sql &= "VALUES( "
                '        sql &= " '" & dt_S5KPEE.Rows(i).Item("DPDC5K").ToString & "', "
                '        sql &= " '" & dt_S5KPEE.Rows(i).Item("PODN5K").ToString & "', "
                '        sql &= " '" & dt_S5KPEE.Rows(i).Item("RFCC5K").ToString & "', "
                '        sql &= " '" & dt_S5KPEE.Rows(i).Item("VNDC5K").ToString & "', "
                '        sql &= " '" & Mid(dt_S5KPEE.Rows(i).Item("RCMC5K").ToString, 1, 6) & "', "
                '        '
                '        sql &= " '" & "" & "', "
                '        sql &= " '" & "" & "', "
                '        sql &= " '" & dt_S5KPEE.Rows(i).Item("PUCN5K").ToString & "', "
                '        sql &= " '" & dt_S5KPEE.Rows(i).Item("CORN5K").ToString & "', "
                '        sql &= " '" & dt_S5KPEE.Rows(i).Item("ORFN5K").ToString & "', "
                '        '
                '        sql &= " '" & "" & "', "
                '        sql &= " '" & "" & "', "
                '        sql &= " '" & "" & "', "
                '        sql &= " '" & "" & "', "
                '        sql &= " '" & "" & "', "
                '        '
                '        sql &= " '" & "" & "', "
                '        sql &= " '" & "" & "', "
                '        sql &= " '" & "" & "', "
                '        sql &= " '" & "" & "', "
                '        sql &= " '" & "" & "', "
                '        '
                '        sql &= " '" & "" & "', "
                '        sql &= " '" & "" & "', "
                '        sql &= " '" & "" & "', "
                '        sql &= " '" & "" & "', "
                '        sql &= " '" & "JOY" & "', "
                '        '
                '        sql &= " '" & "EDI" & "', "
                '        sql &= " '" & "JOY" & "', "
                '        sql &= " " & Now.ToString("yyyyMMdd") & ", "
                '        sql &= " " & Now.ToString("HHmmss") & ", "
                '        sql &= " " & Now.ToString("yyyyMMdd") & ", "
                '        '
                '        sql &= " " & Now.ToString("HHmmss") & ", "
                '        sql &= " '" & dt_S5KPEE.Rows(i).Item("RFCC5K").ToString & "'  "
                '        '
                '        sql &= " ) "
                '        cmd.CommandText = sql
                '        cmd.ExecuteNonQuery()
                '    End If
                'Next
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'EDI2Waves-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([160]-GetCustomerList) Customer資料集  (0=存在/1=不存在)
    '**使用 1.Sales Price
    '**    
    '***********************************************************************************************
    'GetCustomerList-Start
    <WebMethod()> _
    Public Function GetCustomerList(ByVal pCode As String, ByRef pDS As DataSet) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT A.CSTC36 AS CODE, B.FL1I39 AS NAME FROM WAVEDLIB.S3600 A, WAVEDLIB.S3900 B "
            sql &= "WHERE A.CSTC36 = B.CLNC39 "
            sql &= "  AND A.CSTC36 LIKE '%" & pCode & "%' "
            sql &= "ORDER BY A.CSTC36 "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count <= 0 Then RtnCode = 1
            pDS = ds
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetCustomerList-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([170]-GetStandardPrice) 標準單價  (0=存在/1=不存在)
    '**使用 1.Sales Price
    '**    
    '***********************************************************************************************
    'GetStandardPrice-Start
    <WebMethod()> _
    Public Function GetStandardPrice(ByVal pType As String, _
                                     ByVal pVersion As String, _
                                     ByVal pCurrency As String, _
                                     ByVal pCode As String, _
                                     ByRef pPriceA As Single, _
                                     ByRef pPriceB As Single, _
                                     ByRef pPriceC As Single, _
                                     ByRef pPriceD As Single) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        pPriceA = 0
        pPriceB = 0
        pPriceC = 0
        pPriceD = 0
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT "
            sql &= "PA1P41+PA2P41 AS PriceA, "
            sql &= "PB1P41+PB2P41+PB3P41+PB4P41+PB5P41+PB6P41 AS PriceB, "
            sql &= "PC1P41+PC2P41+PC3P41 AS PriceC, "
            sql &= "PCFR41 AS PriceD "
            sql &= "FROM WAVEDLIB.S4100 "
            sql &= "WHERE PTPC41 = '" & pType & "' "
            sql &= "  AND PLVC41 = '" & pVersion & "' "
            sql &= "  AND PCRC41 = '" & pCurrency & "' "
            sql &= "  AND ITMC41 = '" & pCode & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                pPriceA = ds.Tables(0).Rows(0).Item("PriceA")
                pPriceB = ds.Tables(0).Rows(0).Item("PriceB")
                pPriceC = ds.Tables(0).Rows(0).Item("PriceC")
                pPriceD = ds.Tables(0).Rows(0).Item("PriceD")
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetStandardPrice-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([171]-GetStandardPrice1) 標準單價  (0=存在/1=不存在)
    '**使用 1.EDI System
    '**    
    '***********************************************************************************************
    'GetStandardPrice1-Start
    <WebMethod()> _
    Public Function GetStandardPrice1(ByVal pType As String, _
                                      ByVal pVersion As String, _
                                      ByVal pCurrency As String, _
                                      ByVal pCode As String, _
                                      ByRef pLastTime As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT * FROM WAVEDLIB.S4100 "
            sql &= "WHERE PTPC41 = '" & pType & "' "
            sql &= "  AND PLVC41 = '" & pVersion & "' "
            sql &= "  AND PCRC41 = '" & pCurrency & "' "
            sql &= "  AND ITMC41 = '" & pCode & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).Item("RUPU41") = 0 Then
                    pLastTime = ds.Tables(0).Rows(0).Item("RADU41").ToString + ds.Tables(0).Rows(0).Item("RADT41").ToString
                Else
                    pLastTime = ds.Tables(0).Rows(0).Item("RUPU41").ToString + ds.Tables(0).Rows(0).Item("RUPT41").ToString
                End If
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetStandardPrice1-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([180]-GetItemName) 取得Item Name  (0=存在/1=不存在)
    '**使用 1.Sales Price
    '**    
    '***********************************************************************************************
    'GetItemName-Start
    <WebMethod()> _
    Public Function GetItemName(ByVal pItemCode As String, _
                                ByRef pItemName As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        pItemName = ""
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT RTRIM(IT1IA0) || ' ' || RTRIM(IT2IA0) || ' ' || RTRIM(IT3IA0) AS ITEMNAME FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pItemCode & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                pItemName = ds.Tables(0).Rows(0).Item("ITEMNAME")
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemName-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([185]-GetItemName001) 取得Item Name1 / Item Name2 / Item Name3  (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetItemName001-Start
    <WebMethod()> _
    Public Function GetItemName001(ByVal pItemCode As String, _
                                   ByVal pIndex As Integer, _
                                   ByRef pItemName As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        pItemName = ""
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT "
            sql &= "RTRIM(IT1IA0) AS ITEMNAME1, RTRIM(IT2IA0) AS ITEMNAME2, RTRIM(IT3IA0) AS ITEMNAME3 "
            sql &= "FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pItemCode & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                If pIndex = 1 Then
                    pItemName = ds.Tables(0).Rows(0).Item("ITEMNAME1")
                Else
                    If pIndex = 2 Then
                        pItemName = ds.Tables(0).Rows(0).Item("ITEMNAME2")
                    Else
                        pItemName = ds.Tables(0).Rows(0).Item("ITEMNAME3")
                    End If
                End If
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemName001-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([186]-GetItemName001a) 取得Item Name1 / Item Name2 / Item Name3  (0=存在/1=不存在)
    '**(FAS-統計分析)不考慮NODISPLAY
    '**    
    '***********************************************************************************************
    'GetItemName001a-Start
    <WebMethod()> _
    Public Function GetItemName001a(ByVal pItemCode As String, _
                                    ByVal pIndex As Integer, _
                                    ByRef pItemName As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        pItemName = ""
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT "
            sql &= "RTRIM(IT1IA0) AS ITEMNAME1, RTRIM(IT2IA0) AS ITEMNAME2, RTRIM(IT3IA0) AS ITEMNAME3 "
            sql &= "FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pItemCode & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                If pIndex = 1 Then
                    pItemName = ds.Tables(0).Rows(0).Item("ITEMNAME1")
                Else
                    If pIndex = 2 Then
                        pItemName = ds.Tables(0).Rows(0).Item("ITEMNAME2")
                    Else
                        pItemName = ds.Tables(0).Rows(0).Item("ITEMNAME3")
                    End If
                End If
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemName001a-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([190]-GetItemStatistics) 取得Item統計區分  (0=存在/1=不存在)
    '**使用 1.Sales Price
    '**    
    '***********************************************************************************************
    'GetItemStatistics-Start
    <WebMethod()> _
    Public Function GetItemStatistics(ByVal pItemCode As String, _
                                      ByRef pItemStatistics As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        pItemStatistics = ""
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT "
            sql &= "ST1CA0 || '/' || "
            sql &= "ST2CA0 || '/' || "
            sql &= "ST3CA0 || '/' || "
            sql &= "ST4CA0 || '/' || "
            sql &= "ST5CA0 || '/' || "
            sql &= "ST6CA0 || '/' || "
            sql &= "ST7CA0 || '/' AS ITEMSTAT "
            sql &= "FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pItemCode & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                pItemStatistics = ds.Tables(0).Rows(0).Item("ITEMSTAT")
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemStatistics-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([200]-GetItemSize) 取得Item Size  (0=存在/1=不存在)
    '**使用 1.Sales Price
    '**    
    '***********************************************************************************************
    'GetItemSize-Start
    <WebMethod()> _
    Public Function GetItemSize(ByVal pItemCode As String, _
                                ByRef pItemSize As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        pItemSize = ""
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT SIZCA0 AS SIZE "
            sql &= "FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pItemCode & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                pItemSize = ds.Tables(0).Rows(0).Item("SIZE")
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemSize-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([210]-GetItemChain) 取得Item Chain  (0=存在/1=不存在)
    '**使用 1.Sales Price
    '**    
    '***********************************************************************************************
    'GetItemChain-Start
    <WebMethod()> _
    Public Function GetItemChain(ByVal pItemCode As String, _
                                 ByRef pItemChain As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        pItemChain = ""
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT CHNCA0 AS CHAIN "
            sql &= "FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pItemCode & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                pItemChain = ds.Tables(0).Rows(0).Item("CHAIN")
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemChain-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([215]-GetItemClass) 取得Item Class  (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetItemClass-Start
    <WebMethod()> _
    Public Function GetItemClass(ByVal pItemCode As String, _
                                 ByRef pItemClass As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        pItemClass = ""
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT CLSCA0 AS CLASS "
            sql &= "FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pItemCode & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                pItemClass = ds.Tables(0).Rows(0).Item("CLASS")
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemClass-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([215a]-GetItemClass) 取得Item Class  (0=存在/1=不存在)
    '**(FAS-統計分析)不考慮NODISPLAY
    '**    
    '***********************************************************************************************
    'GetItemClassa-Start
    <WebMethod()> _
    Public Function GetItemClassa(ByVal pItemCode As String, _
                                  ByRef pItemClass As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        pItemClass = ""
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT CLSCA0 AS CLASS "
            sql &= "FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pItemCode & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                pItemClass = ds.Tables(0).Rows(0).Item("CLASS")
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemClassa-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([216]-GetItemTape       取得Item Tape (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetItemTape-Start
    <WebMethod()> _
    Public Function GetItemTape(ByVal pItemCode As String, _
                                ByRef pItemTape As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        pItemTape = ""
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT TAPCA0 AS TAPE "
            sql &= "FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pItemCode & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                pItemTape = ds.Tables(0).Rows(0).Item("TAPE")
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemTape-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([217]-GetItemSpecialFeature     取得Item Special Feature (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetItemSpecialFeature-Start
    <WebMethod()> _
    Public Function GetItemSpecialFeature(ByVal pItemCode As String, _
                                          ByRef pItemSpecialFeature As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        pItemSpecialFeature = ""
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT TRIM(SF1CA0) || '/' || TRIM(SF2CA0) || '/' || TRIM(SF3CA0) || '/' || "
            sql &= "      TRIM(SF4CA0) || '/' || TRIM(SF5CA0) || '/' || TRIM(SF6CA0) || '/' AS SF "
            sql &= "FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pItemCode & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                pItemSpecialFeature = ds.Tables(0).Rows(0).Item("SF").ToString
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemSpecialFeature-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([218]-GetItemProdType       取得Item 銷售/採購/內製 (0=存在/1=不存在)
    '**(FAS-統計分析)不考慮NODISPLAY
    '**    
    '***********************************************************************************************
    'GetItemProdType-Start
    <WebMethod()> _
    Public Function GetItemProdType(ByVal pItemCode As String, _
                                    ByRef pSales As String, _
                                    ByRef pPurchase As String, _
                                    ByRef pProduction As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        pSales = ""
        pPurchase = ""
        pProduction = ""
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT SAVFA0 AS SALES, PUVFA0 AS PURCHASE, PALFA0 AS PRODUCTION "
            sql &= "FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pItemCode & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                pSales = ds.Tables(0).Rows(0).Item("SALES").ToString
                If pSales = "" Then pSales = "0"
                pPurchase = ds.Tables(0).Rows(0).Item("PURCHASE").ToString
                If pPurchase = "" Then pPurchase = "0"
                pProduction = ds.Tables(0).Rows(0).Item("PRODUCTION").ToString
                If pProduction = "" Then pProduction = "0"
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemProdType-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([219]-GetItemSliderFinish         取得Item Slider Finish Code (0=存在/1=不存在)
    '**(FAS-統計分析)不考慮NODISPLAY
    '**    
    '***********************************************************************************************
    'GetItemSliderFinish-Start
    <WebMethod()> _
    Public Function GetItemSliderFinish(ByVal pItemCode As String, _
                                        ByRef pFinish As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        pFinish = ""
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT SFNCA0 AS FINISH "
            sql &= "FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pItemCode & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                pFinish = ds.Tables(0).Rows(0).Item("FINISH").ToString
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemSliderFinish-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([220]-GetPriceDiscount) 取得Price Discount (0=存在/1=不存在)
    '**使用 1.Sales Price
    '**    
    '***********************************************************************************************
    'GetPriceDiscount-Start
    <WebMethod()> _
    Public Function GetPriceDiscount(ByVal pVersion As String, _
                                     ByVal pCustomer As String, _
                                     ByVal pItemStat As String, _
                                     ByVal pItemSize As String, _
                                     ByVal pItemChain As String, _
                                     ByRef pDiscount As Single) As Integer
        Dim RtnCode As Integer = 0
        Dim xItemStat As Object = Split(pItemStat, "/")
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        pDiscount = 0
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT DPRR31 AS DISCOUNT "
            sql &= "FROM WAVEDLIB.S3100 "
            sql &= "Where PLVC31 =  '" & pVersion & "' "
            sql &= "  And CSTC31 =  '" & pCustomer & "' "
            sql &= "  And ST1C31 =  '" & xItemStat(0) & "' "
            sql &= "  And ST2C31 =  '" & xItemStat(1) & "' "
            sql &= "  And ST3C31 =  '" & xItemStat(2) & "' "
            sql &= "  And ST4C31 =  '" & xItemStat(3) & "' "
            sql &= "  And ST5C31 =  '" & xItemStat(4) & "' "
            sql &= "  And SIZC31 =  '" & pItemSize & "' "
            sql &= "  And CHNC31 =  '" & pItemChain & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                pDiscount = ds.Tables(0).Rows(0).Item("DISCOUNT")
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetPriceDiscount-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([221]-GetPriceDiscount1) 取得Price Discount (0=存在/1=不存在)
    '**使用 1.EDI System
    '**    
    '***********************************************************************************************
    'GetPriceDiscount1-Start
    <WebMethod()> _
    Public Function GetPriceDiscount1(ByVal pVersion As String, _
                                      ByVal pCustomer As String, _
                                      ByVal pItemStat As String, _
                                      ByVal pItemSize As String, _
                                      ByVal pItemChain As String, _
                                      ByRef pLastTime As String) As Integer
        Dim RtnCode As Integer = 0
        Dim xItemStat As Object = Split(pItemStat, "/")
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT * AS DISCOUNT "
            sql &= "FROM WAVEDLIB.S3100 "
            sql &= "Where PLVC31 =  '" & pVersion & "' "
            sql &= "  And CSTC31 =  '" & pCustomer & "' "
            sql &= "  And ST1C31 =  '" & xItemStat(0) & "' "
            sql &= "  And ST2C31 =  '" & xItemStat(1) & "' "
            sql &= "  And ST3C31 =  '" & xItemStat(2) & "' "
            sql &= "  And ST4C31 =  '" & xItemStat(3) & "' "
            sql &= "  And ST5C31 =  '" & xItemStat(4) & "' "
            sql &= "  And SIZC31 =  '" & pItemSize & "' "
            sql &= "  And CHNC31 =  '" & pItemChain & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).Item("RUPU31") = 0 Then
                    pLastTime = ds.Tables(0).Rows(0).Item("RADU31").ToString + ds.Tables(0).Rows(0).Item("RADT31").ToString
                Else
                    pLastTime = ds.Tables(0).Rows(0).Item("RUPU31").ToString + ds.Tables(0).Rows(0).Item("RUPT31").ToString
                End If
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetPriceDiscount1-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([222]-GetItemQtyUnit      取得Item Qty Unit
    '**
    '**    
    '***********************************************************************************************
    'GetItemQtyUnit-Start
    <WebMethod()> _
    Public Function GetItemQtyUnit(ByVal pItemCode As String) As String
        Dim RtnCode As String = ""
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT QUNCA0 AS QTYUNIT "
            sql &= "FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 =  '" & pItemCode & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                RtnCode = ds.Tables(0).Rows(0).Item("QTYUNIT")
            End If
            cn.Close()
        Catch ex As Exception
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemQtyUnit-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([230]-GetCustomerBuyerPrice) Customer Buyer Price List (0=存在/1=不存在)
    '**使用 1.EDI System
    '**    
    '***********************************************************************************************
    'GetCustomerBuyerPrice-Start
    <WebMethod()> _
    Public Function GetCustomerBuyerPrice(ByVal pCustomer As String, _
                                          ByVal pBuyer As String, _
                                          ByVal pVersion As String, _
                                          ByVal pCurrency As String, _
                                          ByVal pTradeTerms As String, _
                                          ByVal pCode As String, _
                                          ByVal pColor As String, _
                                          ByRef pLastTime As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT * FROM WAVEDLIB.S3M00 "
            sql &= "WHERE CSTC3M = '" & pCustomer & "' "
            sql &= "  AND BYRC3M = '" & pBuyer & "' "
            sql &= "  AND PLVC3M = '" & pVersion & "' "
            sql &= "  AND CRRC3M = '" & pCurrency & "' "
            sql &= "  AND TTRC3M = '" & pTradeTerms & "' "
            sql &= "  AND ITMC3M = '" & pCode & "' "
            sql &= "  AND ( CLRC3M = '" & pColor & "' "
            sql &= "     OR CLRC3M = '" & "" & "' ) "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).Item("RUPU3M") = 0 Then
                    pLastTime = ds.Tables(0).Rows(0).Item("RADU3M").ToString + ds.Tables(0).Rows(0).Item("RADT3M").ToString
                Else
                    pLastTime = ds.Tables(0).Rows(0).Item("RUPU3M").ToString + ds.Tables(0).Rows(0).Item("RUPT3M").ToString
                End If
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetCustomerBuyerPrice-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([240]-GetCustomerPrice) Customer Price List (0=存在/1=不存在)
    '**使用 1.EDI System
    '**    
    '***********************************************************************************************
    'GetCustomerPrice-Start
    <WebMethod()> _
    Public Function GetCustomerPrice(ByVal pCustomer As String, _
                                     ByVal pVersion As String, _
                                     ByVal pCurrency As String, _
                                     ByVal pTradeTerms As String, _
                                     ByVal pCode As String, _
                                     ByVal pColor As String, _
                                     ByRef pLastTime As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT * FROM WAVEDLIB.S3000 "
            sql &= "WHERE CSTC30 = '" & pCustomer & "' "
            sql &= "  AND PLVC30 = '" & pVersion & "' "
            sql &= "  AND CRRC30 = '" & pCurrency & "' "
            sql &= "  AND TTRC30 = '" & pTradeTerms & "' "
            sql &= "  AND ITMC30 = '" & pCode & "' "
            sql &= "  AND ( CLRC30 = '" & pColor & "' "
            sql &= "     OR CLRC30 = '" & "" & "' ) "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).Item("RUPU30") = 0 Then
                    pLastTime = ds.Tables(0).Rows(0).Item("RADU30").ToString + ds.Tables(0).Rows(0).Item("RADT30").ToString
                Else
                    pLastTime = ds.Tables(0).Rows(0).Item("RUPU30").ToString + ds.Tables(0).Rows(0).Item("RUPT30").ToString
                End If
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetCustomerPrice-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([250]-GetBuyerPrice) Buyer Price List (0=存在/1=不存在)
    '**使用 1.EDI System
    '**    
    '***********************************************************************************************
    'GetBuyerPrice-Start
    <WebMethod()> _
    Public Function GetBuyerPrice(ByVal pBuyer As String, _
                                  ByVal pVersion As String, _
                                  ByVal pCurrency As String, _
                                  ByVal pTradeTerms As String, _
                                  ByVal pCode As String, _
                                  ByVal pColor As String, _
                                  ByRef pLastTime As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT * FROM WAVEDLIB.S3L00 "
            sql &= "WHERE BYRC3L = '" & pBuyer & "' "
            sql &= "  AND PLVC3L = '" & pVersion & "' "
            sql &= "  AND CRRC3L = '" & pCurrency & "' "
            sql &= "  AND TTRC3L = '" & pTradeTerms & "' "
            sql &= "  AND ITMC3L = '" & pCode & "' "
            sql &= "  AND ( CLRC3L = '" & pColor & "' "
            sql &= "     OR CLRC3L = '" & "" & "' ) "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).Item("RUPU3L") = 0 Then
                    pLastTime = ds.Tables(0).Rows(0).Item("RADU3L").ToString + ds.Tables(0).Rows(0).Item("RADT3L").ToString
                Else
                    pLastTime = ds.Tables(0).Rows(0).Item("RUPU3L").ToString + ds.Tables(0).Rows(0).Item("RUPT3L").ToString
                End If
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetBuyerPrice-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([260]-GetOrderPrice) Order Price List (0=存在/1=不存在)
    '**使用 1.EDI System
    '**    
    '***********************************************************************************************
    'GetOrderPrice-Start
    <WebMethod()> _
    Public Function GetOrderPrice(ByVal pCustomer As String, _
                                  ByVal pBuyer As String, _
                                  ByVal pPO As String, _
                                  ByVal pSeqno As Integer, _
                                  ByRef pOrderNo As String, _
                                  ByRef pItem As String, _
                                  ByRef pColor As String, _
                                  ByRef pSalesPrice As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            pOrderNo = ""
            pItem = ""
            pColor = ""
            pSalesPrice = "0"
            '
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT A.ORDN5C AS ORDN, B.ITMC5E AS ITEM, B.CLRC5E AS COLOR, B.SUNP5E AS PRICE "
            sql &= "FROM WAVEDLIB.S5C00 A, WAVEDLIB.S5E00 B "
            sql &= "WHERE A.ORDN5C = B.ORDN5E "
            sql &= "  AND A.CSTC5C = '" & pCustomer & "' "
            sql &= "  AND A.BYRC5C = '" & pBuyer & "' "
            sql &= "  AND A.CCNN5C = '" & pPO & "' "
            sql &= "  AND B.COSN5E = " & CStr(pSeqno) & " "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                pOrderNo = ds.Tables(0).Rows(0).Item("ORDN").ToString
                pItem = ds.Tables(0).Rows(0).Item("ITEM").ToString
                pColor = ds.Tables(0).Rows(0).Item("COLOR").ToString
                pSalesPrice = CStr(ds.Tables(0).Rows(0).Item("PRICE") * 1000000)
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetOrderPrice-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([270]-GetOrderPriceDetail) Order Price Detail (0=存在/1=不存在)
    '**使用 1.EDI System
    '**    
    '***********************************************************************************************
    'GetOrderPriceDetail-Start
    <WebMethod()> _
    Public Function GetOrderPriceDetail(ByVal pOrderNo As String, _
                                        ByVal pSubNo As Integer, _
                                        ByRef pPriceVersion As String, _
                                        ByRef pOrderQty As String, _
                                        ByRef pListPrice As String, _
                                        ByRef pSalesPrice As String, _
                                        ByRef pSalesAmount As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            pPriceVersion = ""
            pOrderQty = "0"
            pListPrice = "0"
            pSalesPrice = "0"
            pSalesAmount = "0"
            '
            cn.ConnectionString = ConnectString
            cn.Open()
            '
            sql = "SELECT A.PLVC5C AS PLVC5C, B.OCAQ5E AS OCAQ5E, B.LUPP5E AS LUPP5E, B.SUNP5E AS SUNP5E, B.TCAK5E AS TCAK5E "
            sql &= "FROM WAVEDLIB.S5C00 A, WAVEDLIB.S5E00 B "
            sql &= "WHERE A.ORDN5C = B.ORDN5E "
            sql &= "  AND B.ORDN5E = '" & pOrderNo & "' "
            sql &= "  AND B.OSBN5E = " & CStr(pSubNo) & " "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                pPriceVersion = ds.Tables(0).Rows(0).Item("PLVC5C").ToString
                pOrderQty = CStr(ds.Tables(0).Rows(0).Item("OCAQ5E") * 10000000)
                pListPrice = CStr(ds.Tables(0).Rows(0).Item("LUPP5E") * 1000000)
                pSalesPrice = CStr(ds.Tables(0).Rows(0).Item("SUNP5E") * 1000000)
                pSalesAmount = CStr(ds.Tables(0).Rows(0).Item("TCAK5E") * 100)
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetOrderPriceDetail-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([280]-GetOrderProgress) Order Progress (0=存在/1=不存在)
    '**使用 1.EDI System
    '**    
    '***********************************************************************************************
    'GetOrderProgress-Start
    <WebMethod()> _
    Public Function GetOrderProgress(ByVal pOrderNo As String, _
                                     ByVal pSubNo As Integer, _
                                     ByRef pStatus As String, _
                                     ByRef pPackQty As String, _
                                     ByRef pDeliveryQty As String, _
                                     ByRef pPlanDate As String, _
                                     ByRef pInvoiceQty As String, _
                                     ByRef pCompeltedDate As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds, ds1 As New DataSet
        '
        Try
            pStatus = ""
            pPackQty = "0"
            pDeliveryQty = "0"
            pPlanDate = ""
            pInvoiceQty = "0"
            pCompeltedDate = ""
            '
            cn.ConnectionString = ConnectString
            cn.Open()
            ' S5Q00
            sql = "SELECT ODCU5Q "
            sql &= "FROM WAVEDLIB.S5Q00 "
            sql &= "WHERE ORDN5Q = '" & pOrderNo & "' "
            '
            Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
            DBAdapter2.Fill(ds1)
            If ds1.Tables(0).Rows.Count > 0 Then
                pCompeltedDate = ds1.Tables(0).Rows(0).Item("ODCU5Q").ToString
            Else
                RtnCode = 1
            End If
            '
            If RtnCode = 0 Then
                sql = "SELECT PSTC5F, PAKQ5F, DCNQ5F, NDCU5F, IVCQ5F "
                sql &= "FROM WAVEDLIB.S5F00 "
                sql &= "WHERE ORDN5F = '" & pOrderNo & "' "
                sql &= "  AND OSBN5F = " & CStr(pSubNo) & " "
                '
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds)
                If ds.Tables(0).Rows.Count > 0 Then
                    pStatus = ds.Tables(0).Rows(0).Item("PSTC5F").ToString
                    pPackQty = CStr(ds.Tables(0).Rows(0).Item("PAKQ5F") * 10000000)
                    pDeliveryQty = CStr(ds.Tables(0).Rows(0).Item("DCNQ5F") * 10000000)
                    pPlanDate = ds.Tables(0).Rows(0).Item("NDCU5F").ToString
                    pInvoiceQty = CStr(ds.Tables(0).Rows(0).Item("IVCQ5F") * 10000000)
                Else
                    RtnCode = 1
                End If
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetOrderProgress-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([290]-GetDescriptionMaster)  Description Master (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetDescriptionMaster-Start
    <WebMethod()> _
    Public Function GetDescriptionMaster(ByVal pKey As String, ByVal pCode As String, ByRef pData As String) As Integer

        Dim RtnCode As Integer = 0
        Dim sql As String
        'DB連結設定
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            cn.Open()

            sql = "SELECT '('||RTRIM(CN1I09)||'):'||CN3I09 AS CN3I09 FROM WAVEDLIB.C0900 "
            sql &= "Where DGRC09 = '" & pKey & "' "
            sql &= "  And DDTC09 = '" & pCode & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "C0900")
            If ds.Tables("C0900").Rows.Count > 0 Then
                pData = ds.Tables(0).Rows(0).Item("CN3I09").ToString
            Else
                RtnCode = 1
            End If
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetDescriptionMaster-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([300]-GetPackingListInf)     Packing List (0=OK/1=NG)
    '**
    '**    
    '***********************************************************************************************
    'GetPackingListInf-Start
    <WebMethod()> _
    Public Function GetPackingListInf(ByVal pLogID As String, ByVal pBuyer As String, ByVal pOrder As String, ByVal pUserID As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i As Integer

        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            Dim xRegisterTime As String = NowDateTime
            Dim xItemName As String = ""
            Dim xOrderSubNo As Integer = 0

            cn.ConnectionString = ConnectString
            sql = "SELECT "
            sql &= "CCNN6P, COSN6P, ORDN6P, OSBN6P, ITMC6P, WCFN6P, WCTN6P, "
            sql &= "LNGV6P, CLRC6P, PCKQ6P, IPCV6P, OTNW6N, "
            sql &= "OPKW6N, OTGW6N, DNIF6P, OPIC6N "
            sql &= "FROM WAVEDLIB.S6P00 A, WAVEDLIB.S6N00 B "
            sql &= "WHERE A.PKIN6P = B.PKIN6N "
            sql &= "  AND A.WCFN6P = B.WCFN6N "
            sql &= "  AND A.WCTN6P = B.WCTN6N "
            sql &= "  AND A.WCFN6P = B.WCFN6N "
            sql &= "  AND A.ORDN6P = '" & pOrder & "' "
            sql &= "ORDER BY A.OSBN6P "
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "PACKLIST")
            For i = 0 To ds.Tables("PACKLIST").Rows.Count - 1

                If ds.Tables(0).Rows(i).Item("OSBN6P") <> xOrderSubNo Then
                    sql = "SELECT RegisterTime From B_PackingInstruction "
                    sql &= "Where OrderNo = '" & ds.Tables(0).Rows(i).Item("ORDN6P").ToString & "' "
                    sql &= "  And OrderSubNo = " & ds.Tables(0).Rows(i).Item("OSBN6P").ToString & " "
                    Dim dt_PackingInstruction As DataTable = oDataBase.GetDataTable(sql)
                    If dt_PackingInstruction.Rows.Count > 0 Then
                        xRegisterTime = CDate(dt_PackingInstruction.Rows(0).Item("RegisterTime")).ToString("yyyy/MM/dd HH:mm:ss")

                        sql = "Delete From B_PackingInstruction "
                        sql &= "Where OrderNo = '" & ds.Tables(0).Rows(i).Item("ORDN6P").ToString & "' "
                        sql &= "  And OrderSubNo = " & ds.Tables(0).Rows(i).Item("OSBN6P").ToString & " "
                        oDataBase.ExecuteNonQuery(sql)
                    End If

                    xOrderSubNo = ds.Tables(0).Rows(i).Item("OSBN6P")
                End If

                sql = "Insert Into B_PackingInstruction "
                sql &= "( "
                sql &= "CustomerBuyer, PO, Seqno, RegisterTime, OrderNo, "
                sql &= "OrderSubNo, Item, ItemName1, ItemName2, ItemName3, "
                sql &= "CaseNoStart, CaseNoEnd, Length, Color, PackQty, "
                sql &= "Count, ItemNet, OuterNet, Gross, "
                sql &= "CreateTime, Delivery, OItem, OItemName1, OItemName2, OItemName3 "
                sql &= " ) "
                sql &= "VALUES( "
                sql &= " '" & pBuyer & "', "
                sql &= " '" & ds.Tables(0).Rows(i).Item("CCNN6P").ToString & "', "
                sql &= " " & ds.Tables(0).Rows(i).Item("COSN6P").ToString & ", "
                sql &= " '" & xRegisterTime & "', "
                sql &= " '" & ds.Tables(0).Rows(i).Item("ORDN6P").ToString & "', "

                sql &= " " & ds.Tables(0).Rows(i).Item("OSBN6P").ToString & ", "
                sql &= " '" & ds.Tables(0).Rows(i).Item("ITMC6P").ToString & "', "
                If GetItemName001(ds.Tables(0).Rows(i).Item("ITMC6P").ToString, 1, xItemName) = 0 Then
                    sql &= " '" & xItemName & "', "
                Else
                    sql &= " '" & "" & "', "
                End If
                If GetItemName001(ds.Tables(0).Rows(i).Item("ITMC6P").ToString, 2, xItemName) = 0 Then
                    sql &= " '" & xItemName & "', "
                Else
                    sql &= " '" & "" & "', "
                End If
                If GetItemName001(ds.Tables(0).Rows(i).Item("ITMC6P").ToString, 3, xItemName) = 0 Then
                    sql &= " '" & xItemName & "', "
                Else
                    sql &= " '" & "" & "', "
                End If
                sql &= " " & ds.Tables(0).Rows(i).Item("WCFN6P").ToString & ", "
                sql &= " " & ds.Tables(0).Rows(i).Item("WCTN6P").ToString & ", "
                sql &= " " & ds.Tables(0).Rows(i).Item("LNGV6P").ToString & ", "
                sql &= " '" & ds.Tables(0).Rows(i).Item("CLRC6P").ToString & "', "
                sql &= " " & ds.Tables(0).Rows(i).Item("PCKQ6P").ToString & ", "
                sql &= " " & ds.Tables(0).Rows(i).Item("IPCV6P").ToString & ", "
                sql &= " " & ds.Tables(0).Rows(i).Item("OTNW6N").ToString & ", "

                sql &= " " & ds.Tables(0).Rows(i).Item("OPKW6N").ToString & ", "
                sql &= " " & ds.Tables(0).Rows(i).Item("OTGW6N").ToString & ", "

                sql &= " '" & NowDateTime & "', "

                sql &= " '" & ds.Tables(0).Rows(i).Item("DNIF6P").ToString & "', "


                sql &= " '" & ds.Tables(0).Rows(i).Item("OPIC6N").ToString & "', "
                If GetItemName001(ds.Tables(0).Rows(i).Item("OPIC6N").ToString, 1, xItemName) = 0 Then
                    sql &= " '" & xItemName & "', "
                Else
                    sql &= " '" & "" & "', "
                End If
                If GetItemName001(ds.Tables(0).Rows(i).Item("OPIC6N").ToString, 2, xItemName) = 0 Then
                    sql &= " '" & xItemName & "', "
                Else
                    sql &= " '" & "" & "', "
                End If
                If GetItemName001(ds.Tables(0).Rows(i).Item("OPIC6N").ToString, 3, xItemName) = 0 Then
                    sql &= " '" & xItemName & "' "
                Else
                    sql &= " '" & "" & "' "
                End If
                sql &= " ) "
                oDataBase.ExecuteNonQuery(sql)
            Next

            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetPackingListInf-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([310]-GetItemStructure)(FAS)     ItemStructure (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetItemStructure-Start
    <WebMethod()> _
    Public Function GetItemStructure(ByVal pDepo As String, ByVal pPItem As String, ByVal pClass As String, _
                                     ByRef pItem() As String, ByRef pItemName() As String, ByRef pQty() As String, ByRef pCount As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim i, j As Integer
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim ds1 As New DataSet
        '
        Try
            For j = 1 To 10
                pItem(j) = ""
                pItemName(j) = ""
                pQty(j) = "0"
            Next
            pCount = 0
            '
            cn.ConnectionString = ConnectString
            sql = "SELECT CITCB0, UUSQB0 FROM WAVEDLIB.FB000 "
            sql &= "Where PIPCB0 = '" & pPItem & "' "
            sql &= "  And DPTCB0 = '" & pDepo & "' "
            sql &= "Order By SCSVB0 "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "FB000")
            If ds.Tables("FB000").Rows.Count > 0 Then
                j = 1
                '
                For i = 0 To ds.Tables("FB000").Rows.Count - 1
                    ds1.Clear()
                    '
                    If pClass <> "GAP" Then
                        sql = "SELECT RTRIM(IT1IA0) AS ITEMNAME, CLSCA0 FROM WAVEDLIB.FA000 "
                        sql &= "Where ITMCA0 = '" & ds.Tables(0).Rows(i).Item("CITCB0") & "' "
                        sql &= "  And NDPCA0 <> '" & "1" & "' "
                        If pClass <> "" Then
                            sql &= "  And CLSCA0 = '" & UCase(pClass) & "' "
                        End If
                    Else
                        sql = "SELECT RTRIM(IT1IA0) AS ITEMNAME, CLSCA0 FROM WAVEDLIB.FA000 "
                        sql &= "Where ITMCA0 = '" & ds.Tables(0).Rows(i).Item("CITCB0") & "' "
                        sql &= "  And IT1IA0 || IT2IA0 LIKE '%" & "GAP" & "%' "
                        sql &= "  And NDPCA0 <> '" & "1" & "' "
                    End If
                    '
                    Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
                    DBAdapter2.Fill(ds1, "C0900")
                    If ds1.Tables("FA000").Rows.Count > 0 Then
                        pItemName(j) = ds1.Tables(0).Rows(0).Item("ITEMNAME").ToString
                        pItem(j) = ds.Tables(0).Rows(i).Item("CITCB0").ToString
                        pQty(j) = CStr(ds.Tables(0).Rows(i).Item("UUSQB0") * 10000000)
                        j = j + 1
                    End If
                Next
                pCount = j - 1
            Else
                RtnCode = 1
            End If

            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemStructure-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([310]-GetChildItemStructure)(FAS)     ItemStructure (0=存在/1=不存在)
    '**
    '** +----10.決定CLASS, ST, ITEMNAME   
    '**   |    
    '**   +--20.第一階展開 
    '**   |  |    
    '**   |  +--21.符合CLASS(自已+上階材料) or ITEMNAME 的所有ITEM
    '**   |  |    
    '**   |  +--22.檢查每個ITEM是否符合 ST and ITEMNAME 
    '**   |
    '**   +--30.第二階展開 
    '**   |  |    
    '**   |  +--31.符合CLASS(自已+上階材料) or ITEMNAME 的所有ITEM
    '**   |  |     符合自已CLASS=找到 or 符合ITEMNAME=找到
    '**   |  |     上述找不到,有符合上階CLASS時以此階材料繼續往下階尋找
    '**   |  |    
    '**   |  +--32.檢查每個ITEM是否符合 ST and ITEMNAME 
    '**   |  |     符合ST and ITEMNAME=找到
    '**   |  |     上述找不到,有符合上階CLASS時以此階材料繼續往下階尋找
    '**   |  |    
    '**   |  |    
    '**   |  |    
    '**    
    '***********************************************************************************************
    'GetChildItemStructure-Start
    <WebMethod()> _
    Public Function GetChildItemStructure(ByVal pDepo As String, ByVal pObjectProduct As String, ByVal pPItem As String, _
                                          ByRef pItem() As String, ByRef pItemName() As String, ByRef pQty() As String, ByRef pCount As Integer) As Integer
        Dim RtnCode As Integer = 0

        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim ds1 As New DataSet

        Dim xItem(5), xItemName1(5), xItemName(5), xQty(5), xSTC(5) As String
        Dim sql, xClass, wItem, xObjectProduct As String
        Dim i, idx, xItemCount, LoopMax As Integer
        Dim xFound As Boolean = False
        Dim xYES As Boolean = False
        '
        Try
            '##
            'MsgBox("IN")
            ' 變數及回傳變數清初值
            For i = 1 To 5
                pItem(i) = ""
                pItemName(i) = ""
                pQty(i) = "0"
                xItem(i) = ""
                xItemName1(i) = ""
                xItemName(i) = ""
                xQty(i) = "0"
                xSTC(i) = ""
            Next
            pCount = 0
            '------------------------------------------------------------
            '** +----10.決定CLASS, ST, ITEMNAME   
            '**   |    
            '------------------------------------------------------------
            xClass = Mid(pObjectProduct, 1, InStr(pObjectProduct, "-") - 1)
            xObjectProduct = Mid(pObjectProduct, InStr(pObjectProduct, "-") + 1)
            '
            Select Case xClass
                Case "BOTTOMSTOP"
                    xClass = "PB"
                    xSTC(4) = "3"
                    xSTC(5) = "3"
                Case "CH"
                    If xObjectProduct <> "GAP" Then
                        xClass = "CH"
                        xSTC(4) = "2"
                    Else
                        xClass = ""
                        xSTC(4) = "1"
                    End If
                Case "SLD"
                    xClass = "PS"
                    xSTC(4) = "3"
                    xSTC(5) = "1"
                Case "SLIDERPART"
                    xClass = "SP"
                    xSTC(4) = "4"
                    If xObjectProduct <> "ZPULLER" Then
                        xSTC(5) = "1"
                    Else
                        xSTC(5) = "3"
                    End If
                Case "TAPE"
                    xClass = "T"
                    xSTC(1) = "5"
                    xSTC(2) = "2"
                    xSTC(3) = "1"
                Case "TOPSTOP"
                    xClass = "PT"
                    xSTC(4) = "3"
                    xSTC(5) = "2"
                Case "ZIP"
                    xClass = ""
                    xSTC(4) = "1"
                    '
                Case Else
                    xClass = "ZZ"
                    xSTC(1) = "Z"
            End Select
            '
            ' 決定KEY-ITEM_NAME
            '
            Select Case xObjectProduct
                Case "FINISH"
                    xObjectProduct = "PARTS"
                Case "SEMI"
                    xObjectProduct = "PARTS"
                Case "ZSEMI"
                    xObjectProduct = "Z PARTS"
                Case "ZCSEMI"
                    xObjectProduct = "ZC PARTS"
                Case "ZPULLER"
                    xObjectProduct = " Z"
                    '
                Case Else
            End Select
            '
            cn.ConnectionString = ConnectString
            '------------------------------------------------------------
            '**   +--20.第一階展開 
            '**   |  |    
            '**   |  +--21.符合CLASS(自已+上階材料) or ITEMNAME 的所有ITEM
            '**   |  |    
            '**   |  +--22.檢查每個ITEM是否符合 ST and ITEMNAME 
            '**   |
            '------------------------------------------------------------
            '
            '**   |  |    
            '**   |  +--21.符合CLASS(自已+上階材料) or ITEMNAME 的所有ITEM
            '##
            'MsgBox("PITEM=" & pPItem)
            xItemCount = 1
            sql = "SELECT CITCB0, UUSQB0 FROM WAVEDLIB.FB000 "
            sql &= "Where PIPCB0 = '" & pPItem & "' "
            sql &= "  And DPTCB0 = '" & pDepo & "' "
            sql &= "Order By SCSVB0 "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "FB000")
            For i = 0 To ds.Tables("FB000").Rows.Count - 1
                ds1.Clear()
                '
                sql = "SELECT RTRIM(IT1IA0) AS ITEMNAME1, RTRIM(IT1IA0) || RTRIM(IT2IA0) || RTRIM(IT3IA0) AS ITEMNAME "
                sql &= "FROM WAVEDLIB.FA000 "
                sql &= "Where ITMCA0 = '" & ds.Tables(0).Rows(i).Item("CITCB0") & "' "
                sql &= "  And NDPCA0 <> '" & "1" & "' "
                ' ---
                ' KEY 
                ' CLASS / ITEM_NAME
                Select Case xClass
                    Case "SP"
                        sql &= "  And (CLSCA0 = 'PS' OR CLSCA0 = 'SP') "
                    Case "T"
                        sql &= "  And (CLSCA0 = 'CH' OR CLSCA0 = 'T') "
                    Case ""
                        'ITEMNAME
                        sql &= "  And IT1IA0 || IT2IA0 || IT3IA0 LIKE '%" & xObjectProduct & "%' "
                    Case Else
                        sql &= "  And CLSCA0 = '" & xClass & "' "
                End Select
                '
                Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
                DBAdapter2.Fill(ds1, "FA000")
                If ds1.Tables("FA000").Rows.Count > 0 Then
                    xItem(xItemCount) = ds.Tables(0).Rows(i).Item("CITCB0").ToString
                    xItemName1(xItemCount) = ds1.Tables(0).Rows(0).Item("ITEMNAME1").ToString
                    xItemName(xItemCount) = ds1.Tables(0).Rows(0).Item("ITEMNAME").ToString
                    xQty(xItemCount) = CStr(ds.Tables(0).Rows(i).Item("UUSQB0") * 10000000)
                    xItemCount = xItemCount + 1
                End If
            Next
            '
            '**   |  |    
            '**   |  +--22.檢查每個ITEM是否符合 ST and ITEMNAME 
            '##
            'MsgBox("2")
            '
            idx = 1
            For ItemIndex As Integer = 1 To xItemCount - 1
                '##
                'MsgBox("[" & xClass & "]")
                'MsgBox("[" & xObjectProduct & "]")
                'MsgBox(xItem(ItemIndex))

                ds1.Clear()
                '
                sql = "SELECT RTRIM(IT1IA0) AS ITEMNAME1, RTRIM(IT1IA0) || RTRIM(IT2IA0) || RTRIM(IT3IA0) AS ITEMNAME, "
                sql &= "ST1CA0, ST2CA0, ST3CA0, ST4CA0, ST5CA0, ST6CA0, ST7CA0 "
                sql &= "FROM WAVEDLIB.FA000 "
                sql &= "Where ITMCA0 = '" & xItem(ItemIndex) & "' "
                sql &= "  And NDPCA0 <> '" & "1" & "' "
                ' ---
                ' KEY 
                ' CLASS
                Select Case xClass
                    Case "SP"
                        sql &= "  And (CLSCA0 = 'PS' OR CLSCA0 = 'SP') "
                    Case "T"
                        sql &= "  And (CLSCA0 = 'CH' OR CLSCA0 = 'T') "
                    Case ""
                    Case Else
                        sql &= "  And CLSCA0 = '" & xClass & "' "
                End Select
                ' ---
                ' 取得 ITEM
                Dim DBAdapter3 As New OleDbDataAdapter(sql, cn)
                DBAdapter3.Fill(ds1, "FA000")
                If ds1.Tables("FA000").Rows.Count > 0 Then
                    ' ---
                    ' KEY ITEM_NAME
                    If InStr(ds1.Tables(0).Rows(0).Item("ITEMNAME"), xObjectProduct) > 0 Then
                        xYES = True
                        ' ---
                        ' 統計區分
                        If xSTC(1) <> "" Then
                            If ds1.Tables(0).Rows(0).Item("ST1CA0") <> xSTC(1) Then xYES = False
                        End If
                        If xSTC(2) <> "" Then
                            If ds1.Tables(0).Rows(0).Item("ST2CA0") <> xSTC(2) Then xYES = False
                        End If
                        If xSTC(3) <> "" Then
                            If ds1.Tables(0).Rows(0).Item("ST3CA0") <> xSTC(3) Then xYES = False
                        End If
                        If xSTC(4) <> "" Then
                            If ds1.Tables(0).Rows(0).Item("ST4CA0") <> xSTC(4) Then xYES = False
                        End If
                        If xSTC(5) <> "" Then
                            If ds1.Tables(0).Rows(0).Item("ST5CA0") <> xSTC(5) Then xYES = False
                        End If
                        '
                        If xYES = True Then
                            pItem(idx) = xItem(ItemIndex)
                            pItemName(idx) = xItemName(ItemIndex)
                            pQty(idx) = xQty(ItemIndex)
                            idx = idx + 1
                        End If
                    End If
                End If
            Next
            '
            '------------------------------------------------------------
            '**   +--30.第二階展開 
            '**   |  |    
            '**   |  +--31.符合CLASS(自已+上階材料) or ITEMNAME 的所有ITEM
            '**   |  |     符合自已CLASS=找到 or 符合ITEMNAME=找到
            '**   |  |     上述找不到,有符合上階CLASS時以此階材料繼續往下階尋找
            '**   |  |    
            '**   |  +--32.檢查每個ITEM是否符合 ST and ITEMNAME 
            '**   |  |     符合ST and ITEMNAME=找到
            '**   |  |     上述找不到,有符合上階CLASS時以此階材料繼續往下階尋找
            '------------------------------------------------------------
            If xYES = False Then
                ' 變數及回傳變數清初值
                For i = 1 To 5
                    pItem(i) = ""
                    pItemName(i) = ""
                    pQty(i) = "0"
                    xItem(i) = ""
                    xItemName1(i) = ""
                    xItemName(i) = ""
                    xQty(i) = "0"
                    xSTC(i) = ""
                Next
                pCount = 0
                '
                '**   |  |    
                '**   |  +--31.符合CLASS(自已+上階材料) or ITEMNAME 的所有ITEM
                '**   |  |     符合自已CLASS=找到 or 符合ITEMNAME=找到
                '**   |  |     上述找不到,有符合上階CLASS時以此階材料繼續往下階尋找
                '##
                'MsgBox("PITEM=" & pPItem)
                '
                xItemCount = 1
                LoopMax = 1
                xFound = False
                wItem = pPItem
                '
                While LoopMax < 99 And Not xFound
                    '##
                    'MsgBox("wITEM=" & wItem)
                    ds.Clear()
                    '
                    sql = "SELECT CITCB0, UUSQB0 FROM WAVEDLIB.FB000 "
                    sql &= "Where PIPCB0 = '" & wItem & "' "
                    sql &= "  And DPTCB0 = '" & pDepo & "' "
                    sql &= "Order By SCSVB0 "
                    '
                    Dim DBAdapter4 As New OleDbDataAdapter(sql, cn)
                    DBAdapter4.Fill(ds, "FB000")
                    For i = 0 To ds.Tables("FB000").Rows.Count - 1
                        ds1.Clear()
                        '
                        sql = "SELECT RTRIM(IT1IA0) AS ITEMNAME1, RTRIM(IT1IA0) || RTRIM(IT2IA0) || RTRIM(IT3IA0) AS ITEMNAME, CLSCA0 "
                        sql &= "FROM WAVEDLIB.FA000 "
                        sql &= "Where ITMCA0 = '" & ds.Tables(0).Rows(i).Item("CITCB0") & "' "
                        sql &= "  And NDPCA0 <> '" & "1" & "' "
                        ' ---
                        ' KEY 
                        ' CLASS / ITEM_NAME
                        Select Case xClass
                            Case "SP"
                                sql &= "  And (CLSCA0 = 'PS' OR CLSCA0 = 'SP') "
                            Case "T"
                                sql &= "  And (CLSCA0 = 'CH' OR CLSCA0 = 'T') "
                            Case ""
                                'ITEMNAME
                                sql &= "  And IT1IA0 || IT2IA0 || IT3IA0 LIKE '%" & xObjectProduct & "%' "
                            Case Else
                                sql &= "  And CLSCA0 = '" & xClass & "' "
                        End Select
                        '
                        Dim DBAdapter5 As New OleDbDataAdapter(sql, cn)
                        DBAdapter5.Fill(ds1, "FA000")
                        If ds1.Tables("FA000").Rows.Count > 0 Then
                            If xClass <> "" Then
                                If ds1.Tables(0).Rows(0).Item("CLSCA0").ToString = xClass Then
                                    xItem(xItemCount) = ds.Tables(0).Rows(i).Item("CITCB0").ToString
                                    xItemName1(xItemCount) = ds1.Tables(0).Rows(0).Item("ITEMNAME1").ToString
                                    xItemName(xItemCount) = ds1.Tables(0).Rows(0).Item("ITEMNAME").ToString
                                    xQty(xItemCount) = CStr(ds.Tables(0).Rows(i).Item("UUSQB0") * 10000000)
                                    xItemCount = xItemCount + 1
                                    xFound = True
                                End If
                            Else
                                xItem(xItemCount) = ds.Tables(0).Rows(i).Item("CITCB0").ToString
                                xItemName1(xItemCount) = ds1.Tables(0).Rows(0).Item("ITEMNAME1").ToString
                                xItemName(xItemCount) = ds1.Tables(0).Rows(0).Item("ITEMNAME").ToString
                                xQty(xItemCount) = CStr(ds.Tables(0).Rows(i).Item("UUSQB0") * 10000000)
                                xItemCount = xItemCount + 1
                                xFound = True
                            End If
                            '
                            If xFound = False Then
                                wItem = ds.Tables(0).Rows(i).Item("CITCB0").ToString
                            End If
                        End If
                    Next
                    '
                    LoopMax = LoopMax + 1
                End While
                '
                '**   |  |    
                '**   |  +--32.檢查每個ITEM是否符合 ST and ITEMNAME 
                '**   |  |     符合ST and ITEMNAME=找到
                '**   |  |     上述找不到,有符合上階CLASS時以此階材料繼續往下階尋找
                'For ItemIndex As Integer = 1 To xItemCount - 1
                'MsgBox(CStr(ItemIndex) & "=" & xItem(ItemIndex))
                'Next
                '
                idx = 1
                For ItemIndex As Integer = 1 To xItemCount - 1
                    '##
                    'MsgBox("[" & xClass & "]")
                    'MsgBox("[" & xObjectProduct & "]")
                    'MsgBox(CStr(ItemIndex) & "=" & xItem(ItemIndex))
                    '
                    LoopMax = 1
                    wItem = xItem(ItemIndex)
                    '
                    While LoopMax < 99
                        ds.Clear()
                        '
                        sql = "SELECT CITCB0, UUSQB0 FROM WAVEDLIB.FB000 "
                        sql &= "Where PIPCB0 = '" & wItem & "' "
                        sql &= "  And DPTCB0 = '" & pDepo & "' "
                        sql &= "Order By SCSVB0 "
                        '
                        Dim DBAdapter6 As New OleDbDataAdapter(sql, cn)
                        DBAdapter6.Fill(ds, "FB000")
                        If ds.Tables("FB000").Rows.Count > 0 Then
                            '
                            For i = 0 To ds.Tables("FB000").Rows.Count - 1
                                '##
                                'MsgBox(CStr(idx) & ds.Tables(0).Rows(i).Item("CITCB0"))
                                ds1.Clear()
                                '
                                sql = "SELECT RTRIM(IT1IA0) AS ITEMNAME1, RTRIM(IT1IA0) || RTRIM(IT2IA0) || RTRIM(IT3IA0) AS ITEMNAME, "
                                sql &= "ST1CA0, ST2CA0, ST3CA0, ST4CA0, ST5CA0, ST6CA0, ST7CA0 "
                                sql &= "FROM WAVEDLIB.FA000 "
                                sql &= "Where ITMCA0 = '" & ds.Tables(0).Rows(i).Item("CITCB0") & "' "
                                sql &= "  And NDPCA0 <> '" & "1" & "' "
                                ' ---
                                ' KEY 
                                ' CLASS / ITEM_NAME
                                Select Case xClass
                                    Case "SP"
                                        sql &= "  And (CLSCA0 = 'PS' OR CLSCA0 = 'SP') "
                                    Case "T"
                                        sql &= "  And (CLSCA0 = 'CH' OR CLSCA0 = 'T') "
                                    Case ""
                                        '無條件
                                    Case Else
                                        sql &= "  And CLSCA0 = '" & xClass & "' "
                                End Select
                                ' ---
                                ' 取得 ITEM
                                Dim DBAdapter7 As New OleDbDataAdapter(sql, cn)
                                DBAdapter7.Fill(ds1, "FA000")
                                If ds1.Tables("FA000").Rows.Count > 0 Then
                                    ' ---
                                    ' KEY ITEM_NAME
                                    If InStr(ds1.Tables(0).Rows(0).Item("ITEMNAME"), xObjectProduct) > 0 Then
                                        xYES = True
                                        ' ---
                                        ' 統計區分
                                        If xSTC(1) <> "" Then
                                            If ds1.Tables(0).Rows(0).Item("ST1CA0") <> xSTC(1) Then xYES = False
                                        End If
                                        If xSTC(2) <> "" Then
                                            If ds1.Tables(0).Rows(0).Item("ST2CA0") <> xSTC(2) Then xYES = False
                                        End If
                                        If xSTC(3) <> "" Then
                                            If ds1.Tables(0).Rows(0).Item("ST3CA0") <> xSTC(3) Then xYES = False
                                        End If
                                        If xSTC(4) <> "" Then
                                            If ds1.Tables(0).Rows(0).Item("ST4CA0") <> xSTC(4) Then xYES = False
                                        End If
                                        If xSTC(5) <> "" Then
                                            If ds1.Tables(0).Rows(0).Item("ST5CA0") <> xSTC(5) Then xYES = False
                                        End If
                                        '
                                        If xYES = True Then
                                            pItem(idx) = ds.Tables(0).Rows(i).Item("CITCB0").ToString
                                            pItemName(idx) = ds1.Tables(0).Rows(0).Item("ITEMNAME").ToString
                                            pQty(idx) = CStr(ds.Tables(0).Rows(i).Item("UUSQB0") * 10000000)
                                            idx = idx + 1
                                            Exit While
                                        End If
                                    End If
                                    '
                                    wItem = ds.Tables(0).Rows(i).Item("CITCB0").ToString
                                    '
                                End If
                            Next
                        End If
                        '
                        LoopMax = LoopMax + 1
                    End While
                Next
            End If
            '
            pCount = idx - 1
            '
            cn.Close()
            '##
            'MsgBox("out")
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetChildItemStructure-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([311]-GetChildItemStructurea)(FAS-統計分析)不考慮NODISPLAY     ItemStructure (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetChildItemStructurea-Start
    <WebMethod()> _
    Public Function GetChildItemStructurea(ByVal pDepo As String, ByVal pObjectProduct As String, ByVal pPItem As String, _
                                           ByRef pItem() As String, ByRef pItemName() As String, ByRef pQty() As String, ByRef pCount As Integer) As Integer

        Dim RtnCode As Integer = 0

        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim ds1 As New DataSet

        Dim xItem(5), xItemName1(5), xItemName(5), xQty(5) As String
        Dim sql, xClass, wItem, xObjectProduct As String
        Dim i, idx, xItemCount, LoopMax As Integer
        Dim xFound As Boolean = False
        '
        Try
            '##
            'MsgBox("IN")

            ' 變數及回傳變數清初值
            For i = 1 To 5
                pItem(i) = ""
                pItemName(i) = ""
                pQty(i) = "0"
                xItem(i) = ""
                xItemName1(i) = ""
                xItemName(i) = ""
                xQty(i) = "0"
            Next
            pCount = 0
            ' 決定-ITEM CLASS / 產品
            xClass = Mid(pObjectProduct, 1, InStr(pObjectProduct, "-") - 1)
            xObjectProduct = Mid(pObjectProduct, InStr(pObjectProduct, "-") + 1)
            '
            If xClass = "TAPE" Then
                xClass = "T"
            Else
                If xClass = "SLD" Then
                    xClass = "PS"
                Else
                    If (xClass = "CH" And xObjectProduct = "GAP") Then
                        xClass = ""
                    Else
                        If xClass = "SLIDERPART" Or xClass = "PULLER" Then
                            xClass = "SP"
                        Else
                            If xClass = "TOPSTOP" Then
                                xClass = "PT"
                            Else
                                If xClass = "BOTTOMSTOP" Then
                                    xClass = "PB"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            '
            If xObjectProduct = "FINISH" Or xObjectProduct = "SEMI" Then
                xObjectProduct = "PARTS"
            Else
                If xObjectProduct = "ZSEMI" Then
                    xObjectProduct = "Z PARTS"
                Else
                    If xObjectProduct = "ZCSEMI" Then
                        xObjectProduct = "ZC PARTS"
                    End If
                End If
            End If
            '
            cn.ConnectionString = ConnectString
            '
            ' 第 1 層   GAP CHAIN-->TAPE DYED
            '##
            'MsgBox("1")
            'MsgBox("[" & xClass & "]")
            'MsgBox("[" & xObjectProduct & "]")

            'xItemCount = 1
            'sql = "SELECT CITCB0, UUSQB0 FROM WAVEDLIB.FB000 "
            'sql &= "Where PIPCB0 = '" & pPItem & "' "
            'sql &= "  And DPTCB0 = '" & pDepo & "' "
            'sql &= "Order By SCSVB0 "
            ''
            'Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            'DBAdapter1.Fill(ds, "FB000")
            'For i = 0 To ds.Tables("FB000").Rows.Count - 1
            '    ds1.Clear()
            '    '
            '    '##
            '    'MsgBox("[" & xClass & "]")
            '    '
            '    sql = "SELECT RTRIM(IT1IA0) AS ITEMNAME1, RTRIM(IT1IA0) || RTRIM(IT2IA0) || RTRIM(IT3IA0) AS ITEMNAME "
            '    sql &= "FROM WAVEDLIB.FA000 "
            '    sql &= "Where ITMCA0 = '" & ds.Tables(0).Rows(i).Item("CITCB0") & "' "
            '    sql &= "  And NDPCA0 <> '" & "1" & "' "
            '    '
            '    'If xClass <> "" Then
            '    '    sql &= "  And CLSCA0 = '" & xClass & "' "
            '    'Else
            '    '    sql &= "  And IT1IA0 || IT2IA0 || IT3IA0 LIKE '%" & xObjectProduct & "%' "
            '    'End If
            '    '
            '    Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
            '    DBAdapter2.Fill(ds1, "FA000")
            '    If ds1.Tables("FA000").Rows.Count > 0 Then
            '        xItem(xItemCount) = ds.Tables(0).Rows(i).Item("CITCB0").ToString
            '        xItemName1(xItemCount) = ds1.Tables(0).Rows(0).Item("ITEMNAME1").ToString
            '        xItemName(xItemCount) = ds1.Tables(0).Rows(0).Item("ITEMNAME").ToString
            '        xQty(xItemCount) = CStr(ds.Tables(0).Rows(i).Item("UUSQB0") * 10000000)
            '        xItemCount = xItemCount + 1
            '    End If
            'Next

            xItemCount = 1
            xItem(xItemCount) = pPItem
            xItemName1(xItemCount) = ""
            xItemName(xItemCount) = ""
            xQty(xItemCount) = CStr(1 * 10000000)
            xItemCount = xItemCount + 1
            '
            ' 第 2 層   TAPE DYED-->TAPE SET-->TAPE NAT
            '
            '##
            'MsgBox("2")
            idx = 1
            For ItemIndex As Integer = 1 To xItemCount - 1
                If InStr(xItemName(ItemIndex), xObjectProduct) > 0 Then
                    '##
                    'MsgBox("3-1")
                    'MsgBox(xItem(ItemIndex))

                    pItem(idx) = xItem(ItemIndex)
                    pItemName(idx) = xItemName(ItemIndex)
                    pQty(idx) = xQty(ItemIndex)
                    idx = idx + 1
                Else
                    '##
                    'MsgBox("3-2")
                    'MsgBox("[" & xClass & "]")
                    'MsgBox("[" & xObjectProduct & "]")

                    LoopMax = 1
                    xFound = False
                    wItem = xItem(ItemIndex)
                    While LoopMax < 11 And Not xFound
                        ds.Clear()
                        '
                        sql = "SELECT CITCB0, UUSQB0 FROM WAVEDLIB.FB000 "
                        sql &= "Where PIPCB0 = '" & wItem & "' "
                        sql &= "  And DPTCB0 = '" & pDepo & "' "
                        sql &= "Order By SCSVB0 "
                        '
                        Dim DBAdapter3 As New OleDbDataAdapter(sql, cn)
                        DBAdapter3.Fill(ds, "FB000")
                        If ds.Tables("FB000").Rows.Count > 0 Then
                            '
                            For i = 0 To ds.Tables("FB000").Rows.Count - 1
                                wItem = ds.Tables(0).Rows(0).Item("CITCB0")
                                '##
                                'MsgBox(wItem)
                                '
                                ds1.Clear()
                                '
                                sql = "SELECT RTRIM(IT1IA0) AS ITEMNAME1, RTRIM(IT1IA0) || RTRIM(IT2IA0) || RTRIM(IT3IA0) AS ITEMNAME, ST5CA0 "
                                sql &= "FROM WAVEDLIB.FA000 "
                                sql &= "Where ITMCA0 = '" & ds.Tables(0).Rows(i).Item("CITCB0") & "' "
                                'sql &= "  And NDPCA0 <> '" & "1" & "' "
                                '
                                'MODIFY-START 
                                sql &= "  And CLSCA0 = '" & xClass & "' "
                                '
                                If xObjectProduct = "PULLER" Then
                                    '##
                                    'MsgBox("A")
                                    sql &= "  And ST5CA0 = '" & "3" & "' "
                                End If
                                '
                                'If xClass <> "" And Mid(pObjectProduct, 1, InStr(pObjectProduct, "-") - 1) <> "PULLER" Then
                                '    sql &= "  And CLSCA0 = '" & xClass & "' "
                                'Else
                                '    sql &= "  And IT1IA0 || IT2IA0 || IT3IA0 LIKE '%" & xObjectProduct & "%' "
                                'End If
                                '
                                'MODIFY-END 
                                '
                                Dim DBAdapter4 As New OleDbDataAdapter(sql, cn)
                                DBAdapter4.Fill(ds1, "FA000")
                                If ds1.Tables("FA000").Rows.Count > 0 Then
                                    '
                                    If xObjectProduct = "PULLER" Then
                                        '##
                                        'MsgBox("A")
                                        pItem(idx) = ds.Tables(0).Rows(i).Item("CITCB0").ToString
                                        pItemName(idx) = ds1.Tables(0).Rows(0).Item("ITEMNAME").ToString
                                        pQty(idx) = CStr(ds.Tables(0).Rows(i).Item("UUSQB0") * 10000000)
                                        idx = idx + 1
                                        xFound = True
                                    Else
                                        If InStr(ds1.Tables(0).Rows(0).Item("ITEMNAME"), xObjectProduct) > 0 Then
                                            '##
                                            'MsgBox(ds.Tables(0).Rows(i).Item("CITCB0").ToString)

                                            pItem(idx) = ds.Tables(0).Rows(i).Item("CITCB0").ToString
                                            pItemName(idx) = ds1.Tables(0).Rows(0).Item("ITEMNAME").ToString
                                            pQty(idx) = CStr(ds.Tables(0).Rows(i).Item("UUSQB0") * 10000000)
                                            idx = idx + 1
                                            xFound = True
                                        End If
                                    End If
                                End If
                            Next
                        End If
                        '
                        LoopMax = LoopMax + 1
                    End While
                End If
            Next
            pCount = idx - 1
            '
            cn.Close()
            '##
            'MsgBox("out")
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
        '
        '=====================================================
        'Dim RtnCode As Integer = 0

        'Dim cn As New OleDbConnection
        'Dim ds As New DataSet
        'Dim ds1 As New DataSet

        'Dim xItem(5), xItemName1(5), xItemName(5), xQty(5) As String
        'Dim sql, xClass, wItem, xObjectProduct As String
        'Dim i, idx, xItemCount, LoopMax As Integer
        'Dim xFound As Boolean = False
        ''
        'Try
        '    ' 變數及回傳變數清初值
        '    For i = 1 To 5
        '        pItem(i) = ""
        '        pItemName(i) = ""
        '        pQty(i) = "0"
        '        xItem(i) = ""
        '        xItemName1(i) = ""
        '        xItemName(i) = ""
        '        xQty(i) = "0"
        '    Next
        '    pCount = 0
        '    ' 決定-ITEM CLASS / 產品
        '    xClass = Mid(pObjectProduct, 1, InStr(pObjectProduct, "-") - 1)
        '    xObjectProduct = Mid(pObjectProduct, InStr(pObjectProduct, "-") + 1)

        '    If xClass = "TAPE" Then
        '        xClass = "T"
        '    Else
        '        If xClass = "SLD" Or xClass = "PULLER" Then
        '            xClass = "PS"
        '        Else
        '            If (xClass = "CH" And xObjectProduct = "GAP") Then
        '                xClass = ""
        '            End If
        '        End If
        '    End If
        '    If xObjectProduct = "FINISH" Or xObjectProduct = "SEMI" Then
        '        xObjectProduct = "PARTS"
        '    End If
        '    '
        '    cn.ConnectionString = ConnectString
        '    '
        '    ' 第 1 層   GAP CHAIN-->TAPE DYED
        '    xItemCount = 1
        '    sql = "SELECT CITCB0, UUSQB0 FROM WAVEDLIB.FB000 "
        '    sql &= "Where PIPCB0 = '" & pPItem & "' "
        '    sql &= "  And DPTCB0 = '" & pDepo & "' "
        '    sql &= "Order By SCSVB0 "
        '    '
        '    Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
        '    DBAdapter1.Fill(ds, "FB000")
        '    For i = 0 To ds.Tables("FB000").Rows.Count - 1
        '        ds1.Clear()
        '        '
        '        sql = "SELECT RTRIM(IT1IA0) AS ITEMNAME1, RTRIM(IT1IA0) || RTRIM(IT2IA0) || RTRIM(IT3IA0) AS ITEMNAME "
        '        sql &= "FROM WAVEDLIB.FA000 "
        '        sql &= "Where ITMCA0 = '" & ds.Tables(0).Rows(i).Item("CITCB0") & "' "
        '        If xClass <> "" Then
        '            sql &= "  And CLSCA0 = '" & xClass & "' "
        '        Else
        '            sql &= "  And IT1IA0 || IT2IA0 || IT3IA0 LIKE '%" & xObjectProduct & "%' "
        '        End If
        '        '
        '        Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
        '        DBAdapter2.Fill(ds1, "FA000")
        '        If ds1.Tables("FA000").Rows.Count > 0 Then
        '            xItem(xItemCount) = ds.Tables(0).Rows(i).Item("CITCB0").ToString
        '            xItemName1(xItemCount) = ds1.Tables(0).Rows(0).Item("ITEMNAME1").ToString
        '            xItemName(xItemCount) = ds1.Tables(0).Rows(0).Item("ITEMNAME").ToString
        '            xQty(xItemCount) = CStr(ds.Tables(0).Rows(i).Item("UUSQB0") * 10000000)
        '            xItemCount = xItemCount + 1
        '        End If
        '    Next
        '    '
        '    ' 第 2 層   TAPE DYED-->TAPE SET-->TAPE NAT
        '    '
        '    idx = 1
        '    For ItemIndex As Integer = 1 To xItemCount - 1
        '        If InStr(xItemName(ItemIndex), xObjectProduct) > 0 Then
        '            pItem(idx) = xItem(ItemIndex)
        '            pItemName(idx) = xItemName(ItemIndex)
        '            pQty(idx) = xQty(ItemIndex)
        '            idx = idx + 1
        '        Else
        '            LoopMax = 1
        '            xFound = False
        '            wItem = xItem(ItemIndex)
        '            While LoopMax < 11 And Not xFound
        '                ds.Clear()
        '                '
        '                sql = "SELECT CITCB0, UUSQB0 FROM WAVEDLIB.FB000 "
        '                sql &= "Where PIPCB0 = '" & wItem & "' "
        '                sql &= "  And DPTCB0 = '" & pDepo & "' "
        '                sql &= "Order By SCSVB0 "
        '                '
        '                Dim DBAdapter3 As New OleDbDataAdapter(sql, cn)
        '                DBAdapter3.Fill(ds, "FB000")
        '                If ds.Tables("FB000").Rows.Count > 0 Then
        '                    '
        '                    For i = 0 To ds.Tables("FB000").Rows.Count - 1
        '                        wItem = ds.Tables(0).Rows(0).Item("CITCB0")
        '                        '
        '                        ds1.Clear()
        '                        '
        '                        sql = "SELECT RTRIM(IT1IA0) AS ITEMNAME1, RTRIM(IT1IA0) || RTRIM(IT2IA0) || RTRIM(IT3IA0) AS ITEMNAME "
        '                        sql &= "FROM WAVEDLIB.FA000 "
        '                        sql &= "Where ITMCA0 = '" & ds.Tables(0).Rows(i).Item("CITCB0") & "' "
        '                        If xClass <> "" And Mid(pObjectProduct, 1, InStr(pObjectProduct, "-") - 1) <> "PULLER" Then
        '                            sql &= "  And CLSCA0 = '" & xClass & "' "
        '                        Else
        '                            sql &= "  And IT1IA0 || IT2IA0 || IT3IA0 LIKE '%" & xObjectProduct & "%' "
        '                        End If
        '                        '
        '                        Dim DBAdapter4 As New OleDbDataAdapter(sql, cn)
        '                        DBAdapter4.Fill(ds1, "FA000")
        '                        If ds1.Tables("FA000").Rows.Count > 0 Then
        '                            '
        '                            If InStr(ds1.Tables(0).Rows(0).Item("ITEMNAME"), xObjectProduct) > 0 Then
        '                                pItem(idx) = ds.Tables(0).Rows(i).Item("CITCB0").ToString
        '                                pItemName(idx) = ds1.Tables(0).Rows(0).Item("ITEMNAME").ToString
        '                                pQty(idx) = CStr(ds.Tables(0).Rows(i).Item("UUSQB0") * 10000000)
        '                                idx = idx + 1
        '                                xFound = True
        '                            End If
        '                        End If
        '                    Next
        '                End If
        '                '
        '                LoopMax = LoopMax + 1
        '            End While
        '        End If
        '    Next
        '    pCount = idx - 1
        '    '
        '    cn.Close()
        'Catch ex As Exception
        '    RtnCode = 9
        '    cn.Close()
        'End Try
        ''
        'Return RtnCode
    End Function
    'GetChildItemStructurea-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([320]-GetKeepCodeInventory(FAS)     KeepCodeInventory (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetKeepCodeInventory-Start
    <WebMethod()> _
    Public Function GetKeepCodeInventory(ByVal pDepo As String, ByVal pItem As String, ByVal pColor As String, ByVal pKeepCode As String, _
                                         ByRef pQty As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            pQty = "0"
            '
            cn.ConnectionString = ConnectString
            sql = "SELECT SUM(OKSQDB) AS QTY FROM WAVEDLIB.TDB00 "
            sql &= "Where DPTCDB = '" & pDepo & "' "
            sql &= "  And ITMCDB = '" & pItem & "' "
            If pColor <> "" Then
                sql &= "  And CLRCDB = '" & pColor & "' "
            End If
            If pKeepCode <> "" Then
                sql &= "  And KEPCDB = '" & UCase(pKeepCode) & "' "
            End If
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "TDB00")
            If ds.Tables("TDB00").Rows.Count > 0 Then
                pQty = CStr(ds.Tables(0).Rows(0).Item("QTY") * 10000000)
            Else
                RtnCode = 1
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetKeepCodeInventory-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([321]-GetKeepCodeInventoryZIP(FAS)     KeepCodeInventory (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetKeepCodeInventoryZIP-Start
    <WebMethod()> _
    Public Function GetKeepCodeInventoryZIP(ByVal pDepo As String, ByVal pItem As String, ByVal pColor As String, ByVal pLength As String, ByVal pLengthUnit As String, ByVal pKeepCode As String, _
                                         ByRef pQty As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            pQty = "0"
            '
            cn.ConnectionString = ConnectString
            sql = "SELECT SUM(OKSQDB) AS QTY FROM WAVEDLIB.TDB00 "
            sql &= "Where DPTCDB = '" & pDepo & "' "
            sql &= "  And ITMCDB = '" & pItem & "' "
            If pColor <> "" Then
                sql &= "  And CLRCDB = '" & pColor & "' "
            End If
            If pKeepCode <> "" Then
                sql &= "  And KEPCDB = '" & UCase(pKeepCode) & "' "
            End If
            sql &= "  And LNGVDB = " & pLength & " "
            sql &= "  And LUNCDB = '" & pLengthUnit & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "TDB00")
            If ds.Tables("TDB00").Rows.Count > 0 Then
                pQty = CStr(ds.Tables(0).Rows(0).Item("QTY") * 10000000)
            Else
                RtnCode = 1
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetKeepCodeInventoryZIP-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([322]-GetMaterialKeepCodeInventory(SPS)      Material KeepCodeInventory (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetMaterialKeepCodeInventory-Start
    <WebMethod()> _
    Public Function GetMaterialKeepCodeInventory(ByVal pDepo As String, ByVal pItem As String, ByVal pColor As String, _
                                                 ByRef pDescr As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i As Integer
        Dim xQty As Double
        Dim sql, xDescr As String
        Dim cn As New OleDbConnection
        Dim ds, ds1 As New DataSet
        '
        Try
            pDescr = ""
            xDescr = ""
            xQty = 0
            '
            cn.ConnectionString = ConnectString
            sql = "SELECT KEPCDB, SLCCDB, SUM(OKSQDB) AS QTY "
            sql &= "FROM WAVEDLIB.TDB00 "
            sql &= "Where DPTCDB = '" & pDepo & "' "
            sql &= "  And ITMCDB = '" & pItem & "' "
            If pColor <> "" Then
                sql &= "  And CLRCDB = '" & pColor & "' "
            End If
            sql &= "GROUP BY KEPCDB, SLCCDB "
            sql &= "ORDER BY SUM(OKSQDB) DESC, KEPCDB, SLCCDB "
            '
            Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
            DBAdapter2.Fill(ds, "TDB00")
            For i = 0 To ds.Tables("TDB00").Rows.Count - 1
                xQty = ds.Tables("TDB00").Rows(i).Item("QTY")
                '
                If xQty > 0 Then
                    If xDescr = "" Then
                        xDescr = ds.Tables("TDB00").Rows(i).Item("KEPCDB") & "/" & ds.Tables("TDB00").Rows(i).Item("SLCCDB") & "=" & CStr(xQty)
                    Else
                        xDescr = xDescr & "," & ds.Tables("TDB00").Rows(i).Item("KEPCDB") & "/" & ds.Tables("TDB00").Rows(i).Item("SLCCDB") & "=" & CStr(xQty)
                    End If
                End If
                '
            Next
            '
            If xDescr <> "" Then pDescr = xDescr
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetMaterialKeepCodeInventory-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([330]-GetFreeInventory(FAS)     FreeInventory (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetFreeInventory-Start
    <WebMethod()> _
    Public Function GetFreeInventory(ByVal pDepo As String, ByVal pItem As String, ByVal pColor As String, _
                                     ByRef pQty As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            pQty = "0"
            '
            cn.ConnectionString = ConnectString
            sql = "SELECT SUM(SFRQDC) AS QTY FROM WAVEDLIB.TDC00 "
            sql &= "Where DPTCDC = '" & pDepo & "' "
            sql &= "  And ITMCDC = '" & pItem & "' "
            If pColor <> "" Then
                sql &= "  And CLRCDC = '" & pColor & "' "
            End If
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "TDC00")
            If ds.Tables("TDC00").Rows.Count > 0 Then
                pQty = CStr(ds.Tables(0).Rows(0).Item("QTY") * 10000000)
            Else
                RtnCode = 1
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetFreeInventory-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([331]-GetFreeInventoryZIP(FAS)     FreeInventory (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetFreeInventoryZIP-Start
    <WebMethod()> _
    Public Function GetFreeInventoryZIP(ByVal pDepo As String, ByVal pItem As String, ByVal pColor As String, ByVal pLength As String, ByVal pLengthUnit As String, _
                                     ByRef pQty As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            pQty = "0"
            '
            cn.ConnectionString = ConnectString
            sql = "SELECT SUM(SFRQDC) AS QTY FROM WAVEDLIB.TDC00 "
            sql &= "Where DPTCDC = '" & pDepo & "' "
            sql &= "  And ITMCDC = '" & pItem & "' "
            If pColor <> "" Then
                sql &= "  And CLRCDC = '" & pColor & "' "
            End If
            sql &= "  And LNGVDC = " & pLength & " "
            sql &= "  And LUNCDC = '" & pLengthUnit & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "TDC00")
            If ds.Tables("TDC00").Rows.Count > 0 Then
                pQty = CStr(ds.Tables(0).Rows(0).Item("QTY") * 10000000)
            Else
                RtnCode = 1
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetFreeInventoryZIP-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([333]-GetFreeByLocation(FAS)     Location別FreeInventory (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetFreeByLocation-Start
    <WebMethod()> _
    Public Function GetFreeByLocation(ByVal pDepo As String, ByVal pItem As String, ByVal pColor As String, _
                                      ByRef pDescr As String) As Integer
        Dim RtnCode As Integer = 0
        Dim xUseScheQty, xLocFreeQty As Double
        Dim i As Integer
        Dim sql, xDescr As String
        Dim cn As New OleDbConnection
        Dim ds, ds1 As New DataSet
        '
        Try
            pDescr = ""
            xDescr = ""
            xUseScheQty = 0
            '
            ' 取得 Use Schedule Qty
            cn.ConnectionString = ConnectString
            sql = "SELECT SUM(ASCQDC) AS ACTQTY, SUM(SFRQDC) AS FREEQTY, SUM(USTQDC) AS INUSEQTY, SUM(USSQDC) AS USESCHEQTY, SUM(OKSQDC) AS KEEPQTY "
            sql &= "FROM WAVEDLIB.TDC00 "
            sql &= "Where DPTCDC = '" & pDepo & "' "
            sql &= "  And ITMCDC = '" & pItem & "' "
            If pColor <> "" Then
                sql &= "  And CLRCDC = '" & pColor & "' "
            End If
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "TDC00")
            If ds.Tables("TDC00").Rows.Count > 0 Then
                xDescr = "Free Qty:["
                xUseScheQty = ds.Tables(0).Rows(0).Item("USESCHEQTY")
            End If

            If xUseScheQty > 0 Then
                xDescr = xDescr + "Use Schedule=" + CStr(xUseScheQty) + "/"
            End If

            ds.Clear()
            '
            ' 取得 Location別-Free Qty
            cn.ConnectionString = ConnectString
            sql = "SELECT TRIM(SLCCDE) AS SLCCDE, ASCQDE, USTQDE, SFRQDE "
            sql &= "FROM WAVEDLIB.TDE00 "
            sql &= "Where DPTCDE = '" & pDepo & "' "
            sql &= "  And ITMCDE = '" & pItem & "' "
            If pColor <> "" Then
                sql &= "  And CLRCDE = '" & pColor & "' "
            End If
            sql &= "  And SFRQDE > 0 "
            '
            Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
            DBAdapter2.Fill(ds, "TDE00")
            For i = 0 To ds.Tables("TDE00").Rows.Count - 1
                ds1.Clear()
                xLocFreeQty = 0
                xLocFreeQty = ds.Tables("TDE00").Rows(i).Item("SFRQDE")
                xDescr = xDescr + ds.Tables("TDE00").Rows(i).Item("SLCCDE") + "=" + CStr(xLocFreeQty) + "/"
                '
                'MODIFY-START BY 20240716 (目的不清DELETE)
                '
                ' 取得 Location別-On Keep Qty
                ' Location別-Free Qty - Location別-On Keep Qty
                'sql = "SELECT SLCCDB, SUM(OKSQDB) AS KEEPQTY FROM WAVEDLIB.TDB00 "
                'sql &= "Where DPTCDB = '" & pDepo & "' "
                'sql &= "  And ITMCDB = '" & pItem & "' "
                'If pColor <> "" Then
                '    sql &= "  And CLRCDB = '" & pColor & "' "
                'End If
                'sql &= "  And KEPCDB <> '' "
                'sql &= "  And SLCCDB = '" & ds.Tables("TDE00").Rows(i).Item("SLCCDE") & "' "
                'sql &= "Group By SLCCDB "
                ''
                'Dim DBAdapter3 As New OleDbDataAdapter(sql, cn)
                'DBAdapter3.Fill(ds1, "TDB00")

                'If ds1.Tables("TDB00").Rows.Count > 0 Then
                '    xLocFreeQty = ds.Tables("TDE00").Rows(i).Item("SFRQDE") - ds1.Tables("TDB00").Rows(0).Item("KEEPQTY")
                'Else
                '    xLocFreeQty = ds.Tables("TDE00").Rows(i).Item("SFRQDE")
                'End If

                'If xLocFreeQty > 0 Then
                '    xDescr = xDescr + ds.Tables("TDE00").Rows(i).Item("SLCCDE") + "=" + CStr(xLocFreeQty) + "/"
                'End If
                'MODIFY-END
                '
            Next
            '
            If xDescr <> "" Then pDescr = xDescr + "]"
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetFreeByLocation-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([335]-GetMininumStock(FAS)     Mininum Stock (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetMininumStock-Start
    <WebMethod()> _
    Public Function GetMininumStock(ByVal pDepo As String, ByVal pItem As String, ByVal pColor As String, _
                                    ByRef pQty As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            pQty = "0"
            '
            cn.ConnectionString = ConnectString
            sql = "SELECT SUM(MSTQDC) AS QTY FROM WAVEDLIB.TDC00 "
            sql &= "Where DPTCDC = '" & pDepo & "' "
            sql &= "  And ITMCDC = '" & pItem & "' "
            If pColor <> "" Then
                sql &= "  And CLRCDC = '" & pColor & "' "
            End If
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "TDC00")
            If ds.Tables("TDC00").Rows.Count > 0 Then
                pQty = CStr(ds.Tables(0).Rows(0).Item("QTY") * 10000000)
            Else
                RtnCode = 1
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetMininumStock-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([336]-GetMininumStockZIP(FAS)     Mininum Stock (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetMininumStockZIP-Start
    <WebMethod()> _
    Public Function GetMininumStockZIP(ByVal pDepo As String, ByVal pItem As String, ByVal pColor As String, ByVal pLength As String, ByVal pLengthUnit As String, _
                                    ByRef pQty As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            pQty = "0"
            '
            cn.ConnectionString = ConnectString
            sql = "SELECT SUM(MSTQDC) AS QTY FROM WAVEDLIB.TDC00 "
            sql &= "Where DPTCDC = '" & pDepo & "' "
            sql &= "  And ITMCDC = '" & pItem & "' "
            If pColor <> "" Then
                sql &= "  And CLRCDC = '" & pColor & "' "
            End If
            sql &= "  And LNGVDC = " & pLength & " "
            sql &= "  And LUNCDC = '" & pLengthUnit & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "TDC00")
            If ds.Tables("TDC00").Rows.Count > 0 Then
                pQty = CStr(ds.Tables(0).Rows(0).Item("QTY") * 10000000)
            Else
                RtnCode = 1
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetMininumStockZIP-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([340]-GetProductionQty(FAS)     ProductionQty (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetProductionQty-Start
    <WebMethod()> _
    Public Function GetProductionQty(ByVal pDepo As String, ByVal pItem As String, ByVal pColor As String, _
                                     ByVal pKeepCode As String, ByVal pOnProd As Integer, _
                                     ByRef pQty As String, ByRef pProdQty() As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i, idx, NowMonth As Integer
        Dim Qty As Single
        Dim xProdQty(6) As Single
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim ds1 As New DataSet
        '
        Try
            Qty = "0"
            For i = 1 To 6                      ' N1~N4 PROD-QTY
                xProdQty(i) = 0
                pProdQty(i) = "0"
            Next
            idx = 1
            NowMonth = Year(Now) * 100 + Month(Now) * 1
            '
            cn.ConnectionString = ConnectString
            sql = "SELECT PSCN9F, RLTN9F, ALMQ9F, KEPC9F, RLON9F, UAVU9F FROM WAVEDLIB.F9F00 "
            sql &= "Where DPTC9F = '" & pDepo & "' "
            sql &= "  And ITMC9F = '" & pItem & "' "
            If pColor <> "" Then
                sql &= "  And CLRC9F = '" & pColor & "' "
            End If
            If pOnProd = 0 Then
                sql &= "  And PSHN9F = '" & "" & "' "       ' SCHE PROD
            Else
                sql &= "  And PSHN9F <> '" & "" & "' "      ' ON PROD
            End If
            sql &= "  And PCPU9F = '" & "0" & "' "          ' COMPLETE DATE (PRODUCTION) 
            sql &= "  And PSCN9F <> RLTN9F "                ' PRODUCTION SHEDULE NO. <> RELATION NO.
            sql &= "  And ALMQ9F > 0 "                      ' ALLOCATE QUANTITY
            'sql &= "Order By PSCN9F, RLTN9F "              ' 速度慢-Delete
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "F9F00")
            If ds.Tables("F9F00").Rows.Count > 0 Then
                For i = 0 To ds.Tables("F9F00").Rows.Count - 1

                    ' --2013/11/12 變更後 ----------------------------------------
                    ' RELATION NO or ORIGIN RELATION NO 是 OR~
                    If InStr(ds.Tables(0).Rows(i).Item("RLTN9F"), "OR") > 0 Or InStr(ds.Tables(0).Rows(i).Item("RLON9F"), "OR") > 0 Then

                        If Trim(ds.Tables(0).Rows(i).Item("KEPC9F")) = UCase(pKeepCode) Then
                            Qty = Qty + ds.Tables(0).Rows(i).Item("ALMQ9F")
                            '
                            If Mid(CStr(NowMonth), 1, 4) = Mid(ds.Tables(0).Rows(i).Item("UAVU9F").ToString, 1, 4) Then
                                idx = CInt(Mid(ds.Tables(0).Rows(i).Item("UAVU9F").ToString, 1, 6)) - NowMonth
                            Else
                                If Mid(CStr(NowMonth), 1, 4) < Mid(ds.Tables(0).Rows(i).Item("UAVU9F").ToString, 1, 4) Then
                                    idx = (CInt(Mid(ds.Tables(0).Rows(i).Item("UAVU9F").ToString, 1, 4)) - 1) * 100 + _
                                          CInt(Mid(ds.Tables(0).Rows(i).Item("UAVU9F").ToString, 5, 2)) + 12 - NowMonth
                                Else
                                    idx = ds.Tables(0).Rows(i).Item("UAVU9F") - (CInt(Mid(CStr(NowMonth), 1, 4)) - 1) * 100 + CInt(Mid(CStr(NowMonth), 5, 2)) + 12
                                End If
                            End If
                            If idx <= 0 Then
                                xProdQty(1) = xProdQty(1) + ds.Tables(0).Rows(i).Item("ALMQ9F")
                            Else
                                If idx <= 4 Then
                                    xProdQty(idx + 1) = xProdQty(idx + 1) + ds.Tables(0).Rows(i).Item("ALMQ9F")
                                Else
                                    xProdQty(6) = xProdQty(6) + ds.Tables(0).Rows(i).Item("ALMQ9F")
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            pQty = CStr(Qty * 10000000)
            For i = 1 To 6
                pProdQty(i) = CStr(xProdQty(i) * 10000000)
            Next
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetProductionQty-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([3401]-GetProductionQtyZIP(FAS)     ProductionQty (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetProductionQtyZIP-Start
    <WebMethod()> _
    Public Function GetProductionQtyZIP(ByVal pDepo As String, ByVal pItem As String, ByVal pColor As String, ByVal pLength As String, ByVal pLengthUnit As String, _
                                     ByVal pKeepCode As String, ByVal pOnProd As Integer, _
                                     ByRef pQty As String, ByRef pProdQty() As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i, idx, NowMonth As Integer
        Dim Qty As Single
        Dim xProdQty(6) As Single
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim ds1 As New DataSet
        '
        Try
            Qty = "0"
            For i = 1 To 6                      ' N1~N4 PROD-QTY
                xProdQty(i) = 0
                pProdQty(i) = "0"
            Next
            idx = 1
            NowMonth = Year(Now) * 100 + Month(Now) * 1
            '
            cn.ConnectionString = ConnectString
            sql = "SELECT PSCN9F, RLTN9F, ALMQ9F, KEPC9F, RLON9F, UAVU9F FROM WAVEDLIB.F9F00 "
            sql &= "Where DPTC9F = '" & pDepo & "' "
            sql &= "  And ITMC9F = '" & pItem & "' "
            If pColor <> "" Then
                sql &= "  And CLRC9F = '" & pColor & "' "
            End If
            If pOnProd = 0 Then
                sql &= "  And PSHN9F = '" & "" & "' "       ' SCHE PROD
            Else
                sql &= "  And PSHN9F <> '" & "" & "' "      ' ON PROD
            End If
            sql &= "  And PCPU9F = '" & "0" & "' "          ' COMPLETE DATE (PRODUCTION) 
            sql &= "  And PSCN9F <> RLTN9F "                ' PRODUCTION SHEDULE NO. <> RELATION NO.
            sql &= "  And ALMQ9F > 0 "                      ' ALLOCATE QUANTITY
            sql &= "  And LNGV9F = " & pLength & " "
            sql &= "  And LUNC9F = '" & pLengthUnit & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "F9F00")
            If ds.Tables("F9F00").Rows.Count > 0 Then
                For i = 0 To ds.Tables("F9F00").Rows.Count - 1

                    ' --2013/11/12 變更後 ----------------------------------------
                    ' RELATION NO or ORIGIN RELATION NO 是 OR~
                    If InStr(ds.Tables(0).Rows(i).Item("RLTN9F"), "OR") > 0 Or InStr(ds.Tables(0).Rows(i).Item("RLON9F"), "OR") > 0 Then

                        If Trim(ds.Tables(0).Rows(i).Item("KEPC9F")) = UCase(pKeepCode) Then
                            Qty = Qty + ds.Tables(0).Rows(i).Item("ALMQ9F")
                            '
                            If Mid(CStr(NowMonth), 1, 4) = Mid(ds.Tables(0).Rows(i).Item("UAVU9F").ToString, 1, 4) Then
                                idx = CInt(Mid(ds.Tables(0).Rows(i).Item("UAVU9F").ToString, 1, 6)) - NowMonth
                            Else
                                If Mid(CStr(NowMonth), 1, 4) < Mid(ds.Tables(0).Rows(i).Item("UAVU9F").ToString, 1, 4) Then
                                    idx = (CInt(Mid(ds.Tables(0).Rows(i).Item("UAVU9F").ToString, 1, 4)) - 1) * 100 + _
                                          CInt(Mid(ds.Tables(0).Rows(i).Item("UAVU9F").ToString, 5, 2)) + 12 - NowMonth
                                Else
                                    idx = ds.Tables(0).Rows(i).Item("UAVU9F") - (CInt(Mid(CStr(NowMonth), 1, 4)) - 1) * 100 + CInt(Mid(CStr(NowMonth), 5, 2)) + 12
                                End If
                            End If
                            If idx <= 0 Then
                                xProdQty(1) = xProdQty(1) + ds.Tables(0).Rows(i).Item("ALMQ9F")
                            Else
                                If idx <= 4 Then
                                    xProdQty(idx + 1) = xProdQty(idx + 1) + ds.Tables(0).Rows(i).Item("ALMQ9F")
                                Else
                                    xProdQty(6) = xProdQty(6) + ds.Tables(0).Rows(i).Item("ALMQ9F")
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            pQty = CStr(Qty * 10000000)
            For i = 1 To 6
                pProdQty(i) = CStr(xProdQty(i) * 10000000)
            Next
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetProductionQtyZIP-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([341]-GetFreeProductionQty(FAS)	    Free ProductionQty (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetFreeProductionQty-Start
    <WebMethod()> _
    Public Function GetFreeProductionQty(ByVal pDepo As String, ByVal pItem As String, ByVal pColor As String, _
                                         ByRef pQty As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        Try
            pQty = "0"
            '
            cn.ConnectionString = ConnectString
            sql = "SELECT SUM(RFRQDC) AS QTY FROM WAVEDLIB.TDC00 "
            sql &= "Where DPTCDC = '" & pDepo & "' "
            sql &= "  And ITMCDC = '" & pItem & "' "
            If pColor <> "" Then
                sql &= "  And CLRCDC = '" & pColor & "' "
            End If
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "TDC00")
            If ds.Tables("TDC00").Rows.Count > 0 Then
                pQty = CStr(ds.Tables(0).Rows(0).Item("QTY") * 10000000)
            Else
                RtnCode = 1
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetFreeProductionQty-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([343]-GetPurchaseQty(FAS)     PurchaseQty (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetProductionQty-Start
    <WebMethod()> _
    Public Function GetPurchaseQty(ByVal pDepo As String, ByVal pItem As String, ByVal pColor As String, ByVal pKeepCode As String, _
                                   ByRef pQty As String, ByRef pProdQty() As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i, idx, NowMonth As Integer
        Dim Qty As Single
        Dim xProdQty(6) As Single
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim ds1 As New DataSet
        '
        Try
            Qty = "0"
            For i = 1 To 6                      ' N1~N4 PROD-QTY
                xProdQty(i) = 0
                pProdQty(i) = "0"
            Next
            idx = 1
            NowMonth = Year(Now) * 100 + Month(Now) * 1
            '
            cn.ConnectionString = ConnectString
            sql = "SELECT PDTCEL, ITMCEL, CLRCEL, COMFEL, KEPCEL, RLTNEL, PRIQEL, UAVUEL FROM WAVEDLIB.PEL00 "
            sql &= "Where PDTCEL = '" & pDepo & "' "
            sql &= "  And ITMCEL = '" & pItem & "' "
            If pColor <> "" Then
                sql &= "  And CLRCEL = '" & pColor & "' "
            End If
            sql &= "  And COMFEL = '" & "" & "' "          ' COMPLETE FLAG=未完成
            sql &= "  And PRIQEL > 0 "                     ' REQUEST QTY (PURCHASE)
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "PEL00")
            If ds.Tables("PEL00").Rows.Count > 0 Then
                For i = 0 To ds.Tables("PEL00").Rows.Count - 1
                    ' RELATION NO 是 OR~
                    If InStr(ds.Tables(0).Rows(i).Item("RLTNEL"), "OR") > 0 Then

                        ' KEEPCODE需相同
                        If Trim(ds.Tables(0).Rows(i).Item("KEPCEL")) = UCase(pKeepCode) Then
                            Qty = Qty + ds.Tables(0).Rows(i).Item("PRIQEL")
                            '
                            If Mid(CStr(NowMonth), 1, 4) = Mid(ds.Tables(0).Rows(i).Item("UAVUEL").ToString, 1, 4) Then
                                idx = CInt(Mid(ds.Tables(0).Rows(i).Item("UAVUEL").ToString, 1, 6)) - NowMonth
                            Else
                                If Mid(CStr(NowMonth), 1, 4) < Mid(ds.Tables(0).Rows(i).Item("UAVUEL").ToString, 1, 4) Then
                                    idx = (CInt(Mid(ds.Tables(0).Rows(i).Item("UAVUEL").ToString, 1, 4)) - 1) * 100 + _
                                          CInt(Mid(ds.Tables(0).Rows(i).Item("UAVUEL").ToString, 5, 2)) + 12 - NowMonth
                                Else
                                    idx = ds.Tables(0).Rows(i).Item("UAVUEL") - (CInt(Mid(CStr(NowMonth), 1, 4)) - 1) * 100 + CInt(Mid(CStr(NowMonth), 5, 2)) + 12
                                End If
                            End If
                            If idx <= 0 Then
                                xProdQty(1) = xProdQty(1) + ds.Tables(0).Rows(i).Item("PRIQEL")
                            Else
                                If idx <= 4 Then
                                    xProdQty(idx + 1) = xProdQty(idx + 1) + ds.Tables(0).Rows(i).Item("PRIQEL")
                                Else
                                    xProdQty(6) = xProdQty(6) + ds.Tables(0).Rows(i).Item("PRIQEL")
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            pQty = CStr(Qty * 10000000)
            For i = 1 To 6
                pProdQty(i) = CStr(xProdQty(i) * 10000000)
            Next
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetPurchaseQty-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([344]-GetPurchaseQtyZIP(FAS)     PurchaseQty (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetProductionQtyZIP-Start
    <WebMethod()> _
    Public Function GetPurchaseQtyZIP(ByVal pDepo As String, ByVal pItem As String, ByVal pColor As String, ByVal pLength As String, ByVal pLengthUnit As String, ByVal pKeepCode As String, _
                                   ByRef pQty As String, ByRef pProdQty() As String) As Integer
        Dim RtnCode As Integer = 0
        Dim i, idx, NowMonth As Integer
        Dim Qty As Single
        Dim xProdQty(6) As Single
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim ds1 As New DataSet
        '
        Try
            Qty = "0"
            For i = 1 To 6                      ' N1~N4 PROD-QTY
                xProdQty(i) = 0
                pProdQty(i) = "0"
            Next
            idx = 1
            NowMonth = Year(Now) * 100 + Month(Now) * 1
            '
            cn.ConnectionString = ConnectString
            sql = "SELECT PDTCEL, ITMCEL, CLRCEL, COMFEL, KEPCEL, RLTNEL, PRIQEL, UAVUEL FROM WAVEDLIB.PEL00 "
            sql &= "Where PDTCEL = '" & pDepo & "' "
            sql &= "  And ITMCEL = '" & pItem & "' "
            If pColor <> "" Then
                sql &= "  And CLRCEL = '" & pColor & "' "
            End If
            sql &= "  And COMFEL = '" & "" & "' "          ' COMPLETE FLAG=未完成
            sql &= "  And PRIQEL > 0 "                     ' REQUEST QTY (PURCHASE)
            sql &= "  And LNGVEL = " & pLength & " "
            sql &= "  And LUNCEL = '" & pLengthUnit & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "PEL00")
            If ds.Tables("PEL00").Rows.Count > 0 Then
                For i = 0 To ds.Tables("PEL00").Rows.Count - 1
                    ' RELATION NO 是 OR~
                    If InStr(ds.Tables(0).Rows(i).Item("RLTNEL"), "OR") > 0 Then

                        ' KEEPCODE需相同
                        If Trim(ds.Tables(0).Rows(i).Item("KEPCEL")) = UCase(pKeepCode) Then
                            Qty = Qty + ds.Tables(0).Rows(i).Item("PRIQEL")
                            '
                            If Mid(CStr(NowMonth), 1, 4) = Mid(ds.Tables(0).Rows(i).Item("UAVUEL").ToString, 1, 4) Then
                                idx = CInt(Mid(ds.Tables(0).Rows(i).Item("UAVUEL").ToString, 1, 6)) - NowMonth
                            Else
                                If Mid(CStr(NowMonth), 1, 4) < Mid(ds.Tables(0).Rows(i).Item("UAVUEL").ToString, 1, 4) Then
                                    idx = (CInt(Mid(ds.Tables(0).Rows(i).Item("UAVUEL").ToString, 1, 4)) - 1) * 100 + _
                                          CInt(Mid(ds.Tables(0).Rows(i).Item("UAVUEL").ToString, 5, 2)) + 12 - NowMonth
                                Else
                                    idx = ds.Tables(0).Rows(i).Item("UAVUEL") - (CInt(Mid(CStr(NowMonth), 1, 4)) - 1) * 100 + CInt(Mid(CStr(NowMonth), 5, 2)) + 12
                                End If
                            End If
                            If idx <= 0 Then
                                xProdQty(1) = xProdQty(1) + ds.Tables(0).Rows(i).Item("PRIQEL")
                            Else
                                If idx <= 4 Then
                                    xProdQty(idx + 1) = xProdQty(idx + 1) + ds.Tables(0).Rows(i).Item("PRIQEL")
                                Else
                                    xProdQty(6) = xProdQty(6) + ds.Tables(0).Rows(i).Item("PRIQEL")
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            pQty = CStr(Qty * 10000000)
            For i = 1 To 6
                pProdQty(i) = CStr(xProdQty(i) * 10000000)
            Next
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetPurchaseQtyZIP-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([345]-GetProductionInf(FAS)     Production Inf. (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetProductionInf-Start
    <WebMethod()> _
    Public Function GetProductionInf(ByVal pDepo As String, ByVal pItem As String, ByVal pColor As String, ByVal pKeepCode As String, _
                                     ByRef pProdInf() As String, ByRef pCount As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim i As Integer
        Dim sql, ProdStr As String
        Dim cn As New OleDbConnection
        Dim ds, ds1 As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            sql = "SELECT PSCN9F, RLTN9F, ALMQ9F, KEPC9F, RLON9F, UAVU9F, PSHN9F FROM WAVEDLIB.F9F00 "
            sql &= "Where DPTC9F = '" & pDepo & "' "
            sql &= "  And ITMC9F = '" & pItem & "' "
            If pColor <> "" Then
                sql &= "  And CLRC9F = '" & pColor & "' "
            End If
            sql &= "  And (RLTN9F LIKE 'OR%' Or RLON9F LIKE 'OR%') "    ' 是否是訂單的日程
            sql &= "  And KEPC9F = '" & UCase(pKeepCode) & "' "         ' KEEPCODE
            sql &= "  And PCPU9F = '" & "0" & "' "          ' COMPLETE DATE (PRODUCTION) 
            sql &= "  And PSCN9F <> RLTN9F "                ' PRODUCTION SHEDULE NO. <> RELATION NO.
            sql &= "  And ALMQ9F > 0 "                      ' ALLOCATE QUANTITY
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "F9F00")
            If ds.Tables("F9F00").Rows.Count > 0 Then
                ' 設定初值
                pCount = 0
                For i = 0 To 49
                    pProdInf(i) = ""
                Next
                ' 取得 PROD-INF (SCHE/ON PROD, PRNO, CANUSINGDATE, QTY, ORDERNO, REFERENCE, REQDATE)
                ' ------------------------------------------------------------------------------
                For i = 0 To ds.Tables("F9F00").Rows.Count - 1
                    ProdStr = ""
                    ' SCHE/ON PROD + PRNO
                    If ds.Tables(0).Rows(i).Item("PSHN9F").ToString.Trim = "" Then
                        ProdStr = "SCHE" + "^" + "^"
                    Else
                        ProdStr = "ON" + "^" + ds.Tables(0).Rows(i).Item("PSHN9F").ToString.Trim + "^"
                    End If
                    ' CANUSINGDATE
                    ProdStr = ProdStr + ds.Tables(0).Rows(i).Item("UAVU9F").ToString + "^"
                    ' QTY
                    ProdStr = ProdStr + CStr(ds.Tables(0).Rows(i).Item("ALMQ9F") * 10000000) + "^"
                    '
                    ' 取得 OR-INF
                    ds1.Clear()
                    sql = "SELECT ORDN5C, RDLU5C, CORN5C, ORFN5C FROM WAVEDLIB.S5C00 "
                    If InStr(ds.Tables(0).Rows(i).Item("RLTN9F"), "OR") > 0 Then
                        sql &= "Where ORDN5C = '" & ds.Tables(0).Rows(i).Item("RLTN9F") & "' "
                    Else
                        sql &= "Where ORDN5C = '" & ds.Tables(0).Rows(i).Item("RLON9F") & "' "
                    End If
                    '
                    Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
                    DBAdapter2.Fill(ds1, "S5C00")
                    If ds1.Tables("S5C00").Rows.Count > 0 Then
                        ' ORDERNO
                        ProdStr = ProdStr + ds1.Tables(0).Rows(0).Item("ORDN5C") + "^"
                        ' REFERENCE
                        ProdStr = ProdStr + "[" + ds1.Tables(0).Rows(0).Item("ORFN5C").ToString.Trim + "]/[" + ds1.Tables(0).Rows(0).Item("CORN5C").ToString.Trim + "]" + "^"
                        ' REQDATE
                        ProdStr = ProdStr + ds1.Tables(0).Rows(0).Item("RDLU5C").ToString + "^"
                    Else
                        ProdStr = ProdStr + "^^^"
                    End If
                    '
                    pProdInf(i) = ProdStr
                    pCount = pCount + 1
                Next
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetProductionInf-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([347]-GetPurchaseInf(FAS)     Purchase Inf. (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetPurchaseInf-Start
    <WebMethod()> _
    Public Function GetPurchaseInf(ByVal pDepo As String, ByVal pItem As String, ByVal pColor As String, ByVal pKeepCode As String, _
                                   ByRef pProdInf() As String, ByRef pCount As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim i As Integer
        Dim sql, ProdStr As String
        Dim cn As New OleDbConnection
        Dim ds, ds1 As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            '
            sql = "SELECT PDTCEL, ITMCEL, CLRCEL, COMFEL, KEPCEL, RLTNEL, PRIQEL, UAVUEL, PODNEL FROM WAVEDLIB.PEL00 "
            sql &= "Where PDTCEL = '" & pDepo & "' "
            sql &= "  And ITMCEL = '" & pItem & "' "
            If pColor <> "" Then
                sql &= "  And CLRCEL = '" & pColor & "' "
            End If
            sql &= "  And RLTNEL LIKE 'OR%' "                           ' 是否是訂單的日程
            sql &= "  And KEPCEL = '" & UCase(pKeepCode) & "' "         ' KEEPCODE
            sql &= "  And COMFEL = '' "                                 ' COMPLETE FLAG=未完成
            sql &= "  And PRIQEL > 0 "                                  ' REQUEST QTY (PURCHASE)
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "PEL00")
            If ds.Tables("PEL00").Rows.Count > 0 Then
                ' 設定初值
                pCount = 0
                For i = 0 To 49
                    pProdInf(i) = ""
                Next
                ' 取得 PURCHASE-INF (ON PURCHASE, PONO, CANUSINGDATE, QTY, ORDERNO, REFERENCE, REQDATE)
                ' ------------------------------------------------------------------------------
                For i = 0 To ds.Tables("PEL00").Rows.Count - 1
                    ProdStr = ""
                    ' ON PURCHASE  + PONO
                    ProdStr = "ON" + "^" + ds.Tables(0).Rows(i).Item("PODNEL").ToString.Trim + "^"
                    ' CANUSINGDATE
                    ProdStr = ProdStr + ds.Tables(0).Rows(i).Item("UAVUEL").ToString + "^"
                    ' QTY
                    ProdStr = ProdStr + CStr(ds.Tables(0).Rows(i).Item("PRIQEL") * 10000000) + "^"
                    '
                    ' 取得 OR-INF
                    ds1.Clear()
                    sql = "SELECT ORDN5C, RDLU5C, CORN5C, ORFN5C FROM WAVEDLIB.S5C00 "
                    sql &= "Where ORDN5C = '" & ds.Tables(0).Rows(i).Item("RLTNEL") & "' "
                    '
                    Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
                    DBAdapter2.Fill(ds1, "S5C00")
                    If ds1.Tables("S5C00").Rows.Count > 0 Then
                        ' ORDERNO
                        ProdStr = ProdStr + ds1.Tables(0).Rows(0).Item("ORDN5C") + "^"
                        ' REFERENCE
                        ProdStr = ProdStr + "[" + ds1.Tables(0).Rows(0).Item("ORFN5C").ToString.Trim + "]/[" + ds1.Tables(0).Rows(0).Item("CORN5C").ToString.Trim + "]" + "^"
                        ' REQDATE
                        ProdStr = ProdStr + ds1.Tables(0).Rows(0).Item("RDLU5C").ToString + "^"
                    Else
                        ProdStr = ProdStr + "^^^"
                    End If
                    '
                    pProdInf(i) = ProdStr
                    pCount = pCount + 1
                Next
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetPurchaseInf-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([350]-GetChangeColor(FAS)     ChangeColor (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetChangeColor-Start
    <WebMethod()> _
    Public Function GetChangeColor(ByVal pDepo As String, ByVal pFatherItem As String, ByVal pItem As String, ByVal pOriColor As String, _
                                   ByRef pColor As String) As Integer
        Dim RtnCode As Integer = 0
        Dim RunChangeColor As Boolean
        Dim sql, xColor, xClass, xChainCode As String
        Dim cn As New OleDbConnection
        Dim ds1 As New DataSet
        Dim ds2 As New DataSet
        Dim ds3 As New DataSet
        '
        Try
            ' '' 1. 取得完成品CHAIN兼用色
            ''oWaves.GetChangeColor("01", dt_FCTItem.Rows(i).Item("Item"), xItem(ItemIndex), dt_FCTItem.Rows(i).Item("Color"), xColor)
            cn.ConnectionString = ConnectString
            RunChangeColor = False
            xClass = ""
            xChainCode = ""
            '
            ' 取得父ITEM ChainCode
            sql = "SELECT RTRIM(CLSCA0) AS CLASS, RTRIM(CHNCA0) AS CHAINCODE FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 = '" & pFatherItem & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds1, "FA000-3")
            If ds1.Tables("FA000-3").Rows.Count > 0 Then
                xChainCode = ds1.Tables("FA000-3").Rows(0).Item("CHAINCODE")
            End If
            '
            ' 是否  上色ITEM 及 ITEM Class (SLD/CH ST7C=3) 
            ds1.Clear()
            sql = "SELECT RTRIM(ST7CA0) AS ST7, RTRIM(CLSCA0) AS CLASS FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 = '" & pItem & "' "
            sql &= "  And ST7CA0 = '" & "3" & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            '
            Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
            DBAdapter2.Fill(ds1, "FA000-1")
            If ds1.Tables("FA000-1").Rows.Count > 0 Then
                pColor = pOriColor      ' 上色ITEM
                xClass = ds1.Tables("FA000-1").Rows(0).Item("CLASS")
            Else
                pColor = ""             ' 未上色ITEM
            End If
            '
            ' 上色ITEM 限SLIDER轉兼用色 / 表面處理=E~ (不含ET)
            If pColor <> "" Then
                '
                ' ITEM構成-是否有指定色
                xColor = ""
                sql = "SELECT RTRIM(CLRCB0) AS COLOR FROM WAVEDLIB.FB000 "
                sql &= "Where DPTCB0 = '" & pDepo & "' "
                sql &= "  And PIPCB0 = '" & pFatherItem & "' "
                sql &= "  And CITCB0 = '" & pItem & "' "
                '
                sql &= "  And CLRCB0 <> '" & "" & "' "
                '
                Dim DBAdapter5 As New OleDbDataAdapter(sql, cn)
                DBAdapter5.Fill(ds2, "FB000")
                If ds2.Tables("FB000").Rows.Count > 0 Then
                    pColor = ds2.Tables("FB000").Rows(0).Item("COLOR")
                Else
                    '
                    ' 特殊色指定處理
                    ds1.Clear()
                    sql = "SELECT RTRIM(CCLCB9) AS COLOR FROM WAVEDLIB.FB900 "
                    sql &= "Where PDPCB9 = '" & pDepo & "' "
                    sql &= "  And CHNCB9 = '" & xChainCode & "' "
                    sql &= "  And ST1CB9 = '" & "1" & "' "
                    If xClass = "PS" Then
                        sql &= "  And ST6CB9 = '" & "P" & "' "
                    Else
                        sql &= "  And ST6CB9 = '" & "M" & "' "
                    End If
                    '
                    Dim DBAdapter3 As New OleDbDataAdapter(sql, cn)
                    DBAdapter3.Fill(ds1, "FB900-1")
                    If ds1.Tables("FB900-1").Rows.Count > 0 Then
                        pColor = ds1.Tables("FB900-1").Rows(0).Item("COLOR")
                    Else
                        '
                        ' 限SLIDER轉兼用色 / 表面處理=E~ (不含ET)
                        '20231012 MODIFY JOY
                        ds1.Clear()
                        sql = "SELECT RTRIM(ST7CA0) AS ST7 FROM WAVEDLIB.FA000 "
                        sql &= "Where ITMCA0 = '" & pItem & "' "
                        sql &= "  And CLSCA0 = '" & "PS" & "' "
                        sql &= "  And ( "
                        sql &= "        (SFNCA0 LIKE '" & "E" & "%' And SFNCA0 <> '" & "ET" & "') "
                        sql &= "        OR "
                        sql &= "        (SFNCA0 LIKE '" & "J" & "%' And SFNCA0 <> '" & "J" & "' And SFNCA0 <> '" & "JAJ" & "') "
                        sql &= "  ) "
                        sql &= "  And NDPCA0 <> '" & "1" & "' "
                        '
                        Dim DBAdapter4 As New OleDbDataAdapter(sql, cn)
                        DBAdapter4.Fill(ds1, "FA000-2")
                        If ds1.Tables("FA000-2").Rows.Count > 0 Then
                            RunChangeColor = True
                        Else
                            ' 20240522 開放 CH轉兼用色
                            ' --START
                            ds1.Clear()
                            sql = "SELECT RTRIM(ST7CA0) AS ST7 FROM WAVEDLIB.FA000 "
                            sql &= "Where ITMCA0 = '" & pItem & "' "
                            sql &= "  And CLSCA0 = '" & "CH" & "' "
                            sql &= "  And NDPCA0 <> '" & "1" & "' "
                            '
                            Dim DBAdapter6 As New OleDbDataAdapter(sql, cn)
                            DBAdapter6.Fill(ds1, "FA000-3")
                            If ds1.Tables("FA000-3").Rows.Count > 0 Then
                                RunChangeColor = True
                            End If
                            ' --END
                        End If
                    End If
                End If
            End If
            '
            ' 轉兼用色處理
            If RunChangeColor Then
                '
                ' ITEM構成-是否有指定色
                xColor = ""
                'sql = "SELECT RTRIM(CLRCB0) AS COLOR FROM WAVEDLIB.FB000 "
                'sql &= "Where DPTCB0 = '" & pDepo & "' "
                'sql &= "  And PIPCB0 = '" & pFatherItem & "' "
                'sql &= "  And CITCB0 = '" & pItem & "' "
                ''
                'Dim DBAdapter5 As New OleDbDataAdapter(sql, cn)
                'DBAdapter5.Fill(ds2, "FB000")
                'If ds2.Tables("FB000").Rows.Count > 0 Then
                '    xColor = ds2.Tables("FB000").Rows(0).Item("COLOR")
                'End If
                '
                ' 取得 成品COLOR TABLE
                If xColor = "" Then
                    ds2.Clear()
                    sql = "SELECT RTRIM(CTBNA0) AS COLORTBL FROM WAVEDLIB.FA000 "
                    sql &= "Where ITMCA0 = '" & pFatherItem & "' "
                    sql &= "  And NDPCA0 <> '" & "1" & "' "
                    '
                    Dim DBAdapter6 As New OleDbDataAdapter(sql, cn)
                    DBAdapter6.Fill(ds2, "FA000")
                    If ds2.Tables("FA000").Rows.Count > 0 Then
                        '
                        ' COLOR構成-尋找兼用色
                        sql = "SELECT RTRIM(CCLCA3) AS COLOR FROM WAVEDLIB.FA300 "
                        sql &= "Where PDPCA3 = '" & pDepo & "' "
                        sql &= "  And CTBNA3 = '" & ds2.Tables("FA000").Rows(0).Item("COLORTBL") & "' "
                        sql &= "  And PACCA3 = '" & pOriColor & "' "
                        sql &= "  And ST1CA3 = '" & "1" & "' "
                        ' 20240522 開放 CH轉兼用色
                        ' --START
                        'sql &= "  And ST6CA3 = '" & "P" & "' "
                        If xClass = "PS" Then
                            sql &= "  And ST6CA3 = '" & "P" & "' "
                        Else
                            sql &= "  And ST6CA3 = '" & "M" & "' "
                        End If
                        ' --END
                        '
                        Dim DBAdapter7 As New OleDbDataAdapter(sql, cn)
                        DBAdapter7.Fill(ds3, "FA300")
                        If ds3.Tables("FA300").Rows.Count > 0 Then
                            pColor = ds3.Tables("FA300").Rows(0).Item("COLOR")
                        End If
                    End If
                Else
                    pColor = xColor
                End If
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetChangeColor-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([351]-GetChangeColora(FAS-統計分析)不考慮NODISPLAY     ChangeColor (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetChangeColora-Start
    <WebMethod()> _
    Public Function GetChangeColora(ByVal pDepo As String, ByVal pFatherItem As String, ByVal pItem As String, ByVal pOriColor As String, _
                                    ByRef pColor As String) As Integer
        Dim RtnCode As Integer = 0
        Dim RunChangeColor As Boolean
        Dim sql, xColor, xClass, xChainCode As String
        Dim cn As New OleDbConnection
        Dim ds1 As New DataSet
        Dim ds2 As New DataSet
        Dim ds3 As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            RunChangeColor = False
            xClass = ""
            xChainCode = ""
            '
            ' 取得父ITEM ChainCode
            sql = "SELECT RTRIM(CLSCA0) AS CLASS, RTRIM(CHNCA0) AS CHAINCODE FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 = '" & pFatherItem & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds1, "FA000-3")
            If ds1.Tables("FA000-3").Rows.Count > 0 Then
                xChainCode = ds1.Tables("FA000-3").Rows(0).Item("CHAINCODE")
            End If
            '
            ' 是否  上色ITEM 及 ITEM Class (SLD/CH ST7C=3) 
            ds1.Clear()
            sql = "SELECT RTRIM(ST7CA0) AS ST7, RTRIM(CLSCA0) AS CLASS FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 = '" & pItem & "' "
            sql &= "  And ST7CA0 = '" & "3" & "' "
            '
            Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
            DBAdapter2.Fill(ds1, "FA000-1")
            If ds1.Tables("FA000-1").Rows.Count > 0 Then
                pColor = pOriColor      ' 上色ITEM
                xClass = ds1.Tables("FA000-1").Rows(0).Item("CLASS")
            Else
                pColor = ""             ' 未上色ITEM
            End If
            '
            ' 上色ITEM  限SLIDER轉兼用色 / 表面處理=E~ (不含ET)
            If pColor <> "" Then
                '
                ' 特殊色指定處理
                ds1.Clear()
                sql = "SELECT RTRIM(CCLCB9) AS COLOR FROM WAVEDLIB.FB900 "
                sql &= "Where PDPCB9 = '" & pDepo & "' "
                sql &= "  And CHNCB9 = '" & xChainCode & "' "
                sql &= "  And ST1CB9 = '" & "1" & "' "
                If xClass = "PS" Then
                    sql &= "  And ST6CB9 = '" & "P" & "' "
                Else
                    sql &= "  And ST6CB9 = '" & "M" & "' "
                End If
                '
                Dim DBAdapter3 As New OleDbDataAdapter(sql, cn)
                DBAdapter3.Fill(ds1, "FB900-1")
                If ds1.Tables("FB900-1").Rows.Count > 0 Then
                    pColor = ds1.Tables("FB900-1").Rows(0).Item("COLOR")
                Else
                    '
                    ' 限SLIDER轉兼用色 / 表面處理=E~ (不含ET)
                    ds1.Clear()
                    sql = "SELECT RTRIM(ST7CA0) AS ST7 FROM WAVEDLIB.FA000 "
                    sql &= "Where ITMCA0 = '" & pItem & "' "
                    sql &= "  And CLSCA0 = '" & "PS" & "' "
                    sql &= "  And SFNCA0 LIKE '" & "E" & "%' "
                    sql &= "  And SFNCA0 <> '" & "ET" & "' "
                    '
                    Dim DBAdapter4 As New OleDbDataAdapter(sql, cn)
                    DBAdapter4.Fill(ds1, "FA000-2")
                    If ds1.Tables("FA000-2").Rows.Count > 0 Then
                        RunChangeColor = True
                    End If
                End If
            End If
            '
            ' 轉兼用色處理
            If RunChangeColor Then
                '
                ' ITEM構成-是否有指定色
                xColor = ""
                sql = "SELECT RTRIM(CLRCB0) AS COLOR FROM WAVEDLIB.FB000 "
                sql &= "Where DPTCB0 = '" & pDepo & "' "
                sql &= "  And PIPCB0 = '" & pFatherItem & "' "
                sql &= "  And CITCB0 = '" & pItem & "' "
                '
                Dim DBAdapter5 As New OleDbDataAdapter(sql, cn)
                DBAdapter5.Fill(ds2, "FB000")
                If ds2.Tables("FB000").Rows.Count > 0 Then
                    xColor = ds2.Tables("FB000").Rows(0).Item("COLOR")
                End If
                '
                ' 取得 成品COLOR TABLE
                If xColor = "" Then
                    ds2.Clear()
                    sql = "SELECT RTRIM(CTBNA0) AS COLORTBL FROM WAVEDLIB.FA000 "
                    sql &= "Where ITMCA0 = '" & pFatherItem & "' "
                    '
                    Dim DBAdapter6 As New OleDbDataAdapter(sql, cn)
                    DBAdapter6.Fill(ds2, "FA000")
                    If ds2.Tables("FA000").Rows.Count > 0 Then
                        '
                        ' COLOR構成-尋找兼用色
                        sql = "SELECT RTRIM(CCLCA3) AS COLOR FROM WAVEDLIB.FA300 "
                        sql &= "Where PDPCA3 = '" & pDepo & "' "
                        sql &= "  And CTBNA3 = '" & ds2.Tables("FA000").Rows(0).Item("COLORTBL") & "' "
                        sql &= "  And PACCA3 = '" & pOriColor & "' "
                        sql &= "  And ST1CA3 = '" & "1" & "' "
                        sql &= "  And ST6CA3 = '" & "P" & "' "
                        '
                        Dim DBAdapter7 As New OleDbDataAdapter(sql, cn)
                        DBAdapter7.Fill(ds3, "FA300")
                        If ds3.Tables("FA300").Rows.Count > 0 Then
                            pColor = ds3.Tables("FA300").Rows(0).Item("COLOR")
                        End If
                    End If
                Else
                    pColor = xColor
                End If
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetChangeColora-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([360]-GetOutsideProduction(FAS)     判斷是否內製/外注 (0=內製/1=外注)
    '**
    '**    
    '***********************************************************************************************
    'GetOutsideProduction-Start
    <WebMethod()> _
    Public Function GetOutsideProduction(ByVal pDepo As String, ByVal pItem As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds1, ds2 As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            '
            ' 取得當ITEM PROD-LINE
            ' WINGS Fa100 改成FZ030
            'sql = "SELECT LN1CA1, LN2CA1 FROM WAVEDLIB.FA100 "
            sql = "SELECT LN1CA1, LN2CA1 FROM WAVEDLIB.FZ030 "
            sql &= "Where PDPCA1 = '" & pDepo & "' "
            sql &= "  And ITMCA1 = '" & pItem & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds1, "FA100")
            If ds1.Tables("FA100").Rows.Count > 0 Then
                If Mid(ds1.Tables("FA100").Rows(0).Item("LN1CA1").ToString, 1, 2) = "59" Then
                    RtnCode = 1         ' PROD=LINE=59~~~ = 外注
                Else
                    '
                    ' 非59~~!檢查是外注 PROD-LINE
                    sql = "SELECT LN1CCA, LN2CCA FROM WAVEDLIB.FCA00 "
                    sql &= "Where LN1CCA = '" & ds1.Tables("FA100").Rows(0).Item("LN1CA1") & "' "
                    sql &= "  And LN2CCA = '" & ds1.Tables("FA100").Rows(0).Item("LN2CA1") & "' "
                    sql &= "Group By LN1CCA, LN2CCA "
                    '
                    Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
                    DBAdapter2.Fill(ds2, "FCA00")
                    If ds2.Tables("FCA00").Rows.Count > 0 Then
                        RtnCode = 1         ' 外注
                    End If
                End If
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetOutsideProduction-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([370]-GetPlatingRubberSlider(FAS)       判斷是否電鍍橡膠拉頭 (0=不是/1=是)
    '**
    '**    
    '***********************************************************************************************
    'GetPlatingRubberSlider-Start
    <WebMethod()> _
    Public Function GetPlatingRubberSlider(ByVal pItem As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds1 As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            '
            ' 取得當ITEM PROD-LINE
            sql = "SELECT UT5CA0 FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 = '" & pItem & "' "
            'sql &= "  And ST7CA0 = '" & "4" & "' "
            sql &= "  And ( UT5CA0 = '" & "R" & "' Or UT5CA0 = '" & "G" & "' ) "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds1, "FA000")
            If ds1.Tables("FA000").Rows.Count > 0 Then
                RtnCode = 1
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetPlatingRubberSlider-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([371]-GetPlatingSlider(FAS)         判斷是否電鍍或烤漆拉頭 (0=烤漆/1=電鍍)
    '**
    '**    
    '***********************************************************************************************
    'GetPlatingSlider-Start
    <WebMethod()> _
    Public Function GetPlatingSlider(ByVal pItem As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds1 As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            '
            ' 取得當ITEM PROD-LINE
            sql = "SELECT ST7CA0 FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 = '" & pItem & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds1, "FA000")
            If ds1.Tables("FA000").Rows.Count > 0 Then
                If ds1.Tables("FA000").Rows(0).Item("ST7CA0").ToString = "4" Then
                    RtnCode = 1
                End If
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetPlatingSlider-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([380]-GetItemLineLine(FAS)          判斷是否LINE-LINE ITEM (0=不是/1=是)
    '**
    '**    
    '***********************************************************************************************
    'GetItemLineLine-Start
    <WebMethod()> _
    Public Function GetItemLineLine(ByVal pDepo As String, ByVal pFatherItem As String, ByVal pItem As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql, xPItemLine, xCItemLine As String
        Dim cn As New OleDbConnection
        Dim ds1 As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            xPItemLine = "999-99"
            xCItemLine = "999-99"
            '
            ' 取得 父ITEM PROD-LINE
            ' WINGS Fa100 改成 FZ030
            'sql = "SELECT LN1CA1 || LN2CA1 AS LINE FROM WAVEDLIB.FA100 "
            sql = "SELECT LN1CA1 || LN2CA1 AS LINE FROM WAVEDLIB.FZ030 "
            sql &= "Where PDPCA1 = '" & pDepo & "' "
            sql &= "  And ITMCA1 = '" & pFatherItem & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds1, "FA100-1")
            If ds1.Tables("FA100-1").Rows.Count > 0 Then
                xPItemLine = ds1.Tables("FA100-1").Rows(0).Item("LINE").ToString
            End If
            '
            ' 取得 子ITEM PROD-LINE
            ' WINGS Fa100 改成 FZ030
            ds1.Clear()
            'sql = "SELECT LN1CA1 || LN2CA1 AS LINE FROM WAVEDLIB.FA100 "
            sql = "SELECT LN1CA1 || LN2CA1 AS LINE FROM WAVEDLIB.FZ030 "
            sql &= "Where PDPCA1 = '" & pDepo & "' "
            sql &= "  And ITMCA1 = '" & pItem & "' "
            '
            Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
            DBAdapter2.Fill(ds1, "FA100-2")
            If ds1.Tables("FA100-2").Rows.Count > 0 Then
                xCItemLine = ds1.Tables("FA100-2").Rows(0).Item("LINE").ToString
            End If
            '
            If xPItemLine <> "" And xCItemLine <> "" Then
                '
                ' 檢測是否 ITEM LINE-LINE
                ds1.Clear()
                sql = "SELECT CP1CA7 FROM WAVEDLIB.FA700 "
                sql &= "Where PDPCA7 = '" & pDepo & "' "
                sql &= "  And LN1CA7 = '" & Mid(xPItemLine, 1, 3) & "' "
                sql &= "  And LN2CA7 = '" & Mid(xPItemLine, 4, 2) & "' "
                sql &= "  And CP1CA7 || CP2CA7 || CP3CA7 || CP4CA7 || CP5CA7 || CP6CA7 || CP7CA7 || CP8CA7 || CP9CA7 || CPACA7 Like '%" & xCItemLine & "%' "
                '
                Dim DBAdapter3 As New OleDbDataAdapter(sql, cn)
                DBAdapter3.Fill(ds1, "FA700")
                If ds1.Tables("FA700").Rows.Count > 0 Then
                    RtnCode = 1
                End If
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemLineLine-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([390]-GetPurchaseItem(FAS)              判斷是否採購品 (0=不是/1=是)
    '**
    '**    
    '***********************************************************************************************
    'GetPurchaseItem-Start
    <WebMethod()> _
    Public Function GetPurchaseItem(ByVal pItem As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds1 As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            '
            ' 是否採購品
            sql = "SELECT ITMCA0 FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 = '" & pItem & "' "
            sql &= "  And PUVFA0 = '" & "1" & "' "
            sql &= "  And PALFA0 <> '" & "1" & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds1, "FA000")
            If ds1.Tables("FA000").Rows.Count > 0 Then
                RtnCode = 1
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetPurchaseItem-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([395]-GetItemSlider(FAS)              判斷Slider (DAxxxx, DSxxxxx)
    '**
    '**    
    '***********************************************************************************************
    'GetItemSlider-Start
    <WebMethod()> _
   Public Function GetItemSlider(ByVal pItem As String, _
                                 ByRef pItemSlider As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds1 As New DataSet
        '
        Try
            pItemSlider = ""
            cn.ConnectionString = ConnectString
            '
            ' 判斷製品區分 2020/11/10 MODIFY
            'sql = "SELECT SLDCA0 FROM WAVEDLIB.FA000 "
            sql = "SELECT TRIM(SLDCA0) || '-' || TRIM(SFNCA0) AS SLDCA0 FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 = '" & pItem & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds1, "FA000")
            If ds1.Tables("FA000").Rows.Count > 0 Then
                pItemSlider = ds1.Tables("FA000").Rows(0).Item("SLDCA0").ToString
            Else
                RtnCode = 1
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemSlider-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([400]-GetItemProduct(FAS)           判斷製品區分 (MF=金屬/CF=樹脂/VF=塑鋼)
    '**
    '**    
    '***********************************************************************************************
    'GetItemProduct-Start
    <WebMethod()> _
    Public Function GetItemProduct(ByVal pItem As String, _
                                   ByRef pItemProduct As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds1 As New DataSet
        '
        Try
            pItemProduct = ""
            cn.ConnectionString = ConnectString
            '
            ' 判斷製品區分
            sql = "SELECT ST1CA0, ST2CA0, ST4CA0 FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 = '" & pItem & "' "
            sql &= "  And NDPCA0 <> '" & "1" & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds1, "FA000")
            If ds1.Tables("FA000").Rows.Count > 0 Then
                If ds1.Tables("FA000").Rows(0).Item("ST1CA0").ToString = "1" Then
                    '
                    Select Case ds1.Tables("FA000").Rows(0).Item("ST2CA0").ToString
                        Case "1"
                            pItemProduct = "MF"      ' 金屬-F
                        Case "2"
                            pItemProduct = "CF"      ' 樹脂-F
                        Case "3"
                            pItemProduct = "VF"      ' 塑鋼-F
                        Case Else
                            RtnCode = 1
                    End Select
                    '
                End If
            Else
                RtnCode = 1
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemProduct-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([401]-GetItemProduct(FAS)           判斷製品區分 (MF=金屬/CF=樹脂/VF=塑鋼)
    '**(FAS-統計分析)不考慮NODISPLAY
    '**    
    '***********************************************************************************************
    'GetItemProducta-Start
    <WebMethod()> _
    Public Function GetItemProducta(ByVal pItem As String, _
                                    ByRef pItemProduct As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds1 As New DataSet
        '
        Try
            pItemProduct = ""
            cn.ConnectionString = ConnectString
            '
            ' 判斷製品區分
            sql = "SELECT ST1CA0, ST2CA0, ST4CA0 FROM WAVEDLIB.FA000 "
            sql &= "Where ITMCA0 = '" & pItem & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds1, "FA000")
            If ds1.Tables("FA000").Rows.Count > 0 Then
                If ds1.Tables("FA000").Rows(0).Item("ST1CA0").ToString = "1" Then
                    '
                    Select Case ds1.Tables("FA000").Rows(0).Item("ST2CA0").ToString
                        Case "1"
                            pItemProduct = "MF"      ' 金屬-F
                        Case "2"
                            pItemProduct = "CF"      ' 樹脂-F
                        Case "3"
                            pItemProduct = "VF"      ' 塑鋼-F
                        Case Else
                            RtnCode = 1
                    End Select
                    '
                End If
            Else
                RtnCode = 1
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetItemProducta-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([410]-GetSearchItem(FAS)        取得過濾箱搜尋Key 
    '**
    '**    
    '***********************************************************************************************
    'GetSearchItem-Start
    <WebMethod()> _
    Public Function GetSearchItem(ByVal pPartType As String, ByVal pItem As String, ByVal pColor As String, _
                                  ByRef pSearchItem() As String) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            For i As Integer = 1 To 10
                pSearchItem(i) = ""
            Next
            '
            '過濾箱搜尋Key 
            '           1       2       3           4           5           6               7               8               9       10
            ' CHAIN     購入品  SIZE    CHAINTYPE   CHAINCODE   TAPECODE    SPECIALFEATURE  內製(I)/外注(O) 電鍍(E)/烤漆(P) 橡膠    COLOR
            ' SLD       購入品  SIZE    SLIDER      CHAINCODE   TAPECODE    SPECIALFEATURE  內製(I)/外注(O) 電鍍(E)/烤漆(P) 橡膠    COLOR
            ' 
            If GetPurchaseItem(pItem) = 1 Then                              ' 取得是否採購品        (0=不是/1=是) 
                pSearchItem(1) = "P"
            End If
            GetItemSize(pItem, pSearchItem(2))                              ' 取得ItemSize          (0=存在/1=不存在)
            If pPartType = "SLD" Then
                GetItemSlider(pItem, pSearchItem(3))                        ' SLD:取得Slider        (DAxxx, DSxxxx, ....)
            Else
                GetItemProduct(pItem, pSearchItem(3))                       ' CHAIN:取得製品區分    (MF=金屬/CF=樹脂/VF=塑鋼)
            End If
            GetItemChain(pItem, pSearchItem(4))                             ' 取得Item Chain        (0=存在/1=不存在)
            GetItemTape(pItem, pSearchItem(5))                              ' 取得Item TAPE         (0=存在/1=不存在)
            GetItemSpecialFeature(pItem, pSearchItem(6))                    ' 取得Item Special Feature      (0=存在/1=不存在)
            If GetOutsideProduction("01", pItem) = 0 Then
                pSearchItem(7) = "I"
            Else
                pSearchItem(7) = "O"
            End If
            If GetPlatingSlider(pItem) = 1 Then                             ' 判斷是否電鍍或烤漆拉頭 (0=烤漆/1=電鍍)
                pSearchItem(8) = "P"
            Else
                pSearchItem(8) = "E"
            End If
            If GetPlatingRubberSlider(pItem) <> 0 Then pSearchItem(9) = "R" ' 取得電鍍橡膠拉頭      (0=不是/1=是)
            pSearchItem(10) = pColor                                        ' 取得COLOR             由CST_PullerColor
            '
            'For i As Integer = 1 To 10
            'pSearchItem(i) = Replace(pSearchItem(i), " ", "")
            'Next
            '
        Catch ex As Exception
            RtnCode = 9
        End Try
        '
        Return RtnCode
    End Function
    'GetSearchItem-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([420]-GetSalesManCode) 取得業務員代碼   (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetSalesManCode-Start
    <WebMethod()> _
    Public Function GetSalesManCode(ByVal pCode As String, ByRef pSalesManCode As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds1 As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            '
            sql = "SELECT SPRC36 FROM WAVEDLIB.S3600 "
            sql &= "WHERE CSTC36 = '" & pCode & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds1, "S3600")
            If ds1.Tables("S3600").Rows.Count > 0 Then
                pSalesManCode = ds1.Tables("S3600").Rows(0).Item("SPRC36").ToString.Trim
            Else
                RtnCode = 1
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetSalesManCode-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([430]-GetCostPrice) 取得成本單價 (0=存在/1=不存在)
    '**
    '**    
    '***********************************************************************************************
    'GetCostPrice-Start
    <WebMethod()> _
    Public Function GetCostPrice(ByVal pCode As String, ByRef pPriceA As String, ByRef pPriceB As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds1 As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            '
            sql = "SELECT CFAP81, CFBP81 FROM WAVEDLIB.S8100 "
            sql &= "WHERE CATC81 = 'E' "
            sql &= "AND   ITMC81 = '" & pCode & "' "
            sql &= "AND   CAVC81 <= " & Now.ToString("yyyyMM") & " "
            sql &= "ORDER BY CAVC81 DESC "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds1, "S8100")
            If ds1.Tables("S8100").Rows.Count > 0 Then
                pPriceA = CStr(ds1.Tables("S8100").Rows(0).Item("CFAP81") * 1000000)
                pPriceB = CStr(ds1.Tables("S8100").Rows(0).Item("CFBP81") * 1000000)
            Else
                RtnCode = 1
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetCostPrice-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([700]-UpdateFCData(V-FAS)) 新VENDOR FC DATA (0=存在/1=不存在) 
    '**
    '**    
    '***********************************************************************************************
    'UpdateFCData-Start
    <WebMethod()> _
    Public Function UpdateFCData(ByVal pKey1 As String, ByVal pKey2 As String, ByVal pKey3 As String, _
                                 ByVal pKey4 As String, ByVal pKey5 As String, ByVal pKey6 As String, ByVal pKey7 As String, _
                                 ByVal pValue1 As String, ByVal pValue2 As String, ByVal pValue3 As String, _
                                 ByVal pValue4 As String, ByVal pValue5 As String, ByVal pValue6 As String, ByVal pValue7 As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cnString As String = "Provider=sqloledb;Data Source=10.245.1.20;Initial Catalog=WorkFlow;User Id=sa;"
        '
        Try
            cn.ConnectionString = cnString
            cn.Open()
            '
            sql = "Update F_FCDATA Set "
            sql &= "Q_ScheProd = " & pValue1 & ", "
            sql &= "Q_OnProd = " & pValue2 & ", "
            sql &= "Q_FreeInv = " & pValue3 & ", "
            sql &= "Q_KeepInv = " & pValue4 & ", "
            sql &= "UnitPrice = " & CStr(CDbl(pValue5) / 1000000) & ", "
            sql &= "INVQTY = " & pValue6 & ", "
            sql &= "INVAMT = " & CStr(CDbl(pValue7) / 1000000) & ", "
            sql &= "ModifyUser = '" & "FM-LSPLAN" & "', "
            sql &= "ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
            '
            sql &= "WHERE KEEPCODE = '" & pKey1 & "' "
            sql &= "AND   ITEMCODE1 = '" & pKey2 & "' "
            sql &= "AND   COLOR1 = '" & pKey3 & "' "
            sql &= "AND   LENGTH1 = '" & pKey4 & "' "
            sql &= "AND   LENGTHU1 = '" & pKey5 & "' "
            sql &= "AND   BUYMONTH >= '" & pKey6 & "' "
            sql &= "AND   CREATEUSER = '" & pKey7 & "' "
            Dim cmd As OleDbCommand = New OleDbCommand(sql, cn)
            cmd.ExecuteNonQuery()
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'UpdateFCData-End
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([710]-GetPFASBuyer) 取得PFAS Buyer (0=NIT PFAS/1=PFAS)
    '**
    '**    
    '***********************************************************************************************
    'GetPFASBuyer-Start
    <WebMethod()> _
    Public Function GetPFASBuyer(ByVal pBuyer As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim ds1 As New DataSet
        '
        Try
            cn.ConnectionString = ConnectString
            '
            sql = "SELECT BYRC35, PFCK35 FROM WAVEDLIB.S35E0 "
            sql &= "WHERE BYRC35 = '" & pBuyer & "' "
            sql &= "  AND PFCK35 = '1' "
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds1, "S35E0")
            If ds1.Tables("S35E0").Rows.Count > 0 Then
                RtnCode = 1
            End If
            '
            cn.Close()
        Catch ex As Exception
            RtnCode = 9
            cn.Close()
        End Try
        '
        Return RtnCode
    End Function
    'GetPFASBuyer-End
End Class
