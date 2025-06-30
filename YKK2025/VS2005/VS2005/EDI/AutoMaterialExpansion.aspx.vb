Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Threading

Partial Class AutoMaterialExpansion
    Inherits System.Web.UI.Page
    'iexplore.exe "http://10.245.0.153/EDI/AutoMaterialExpansion.aspx"
    '
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim fpObj As New ForProject             ' 操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim uFASMapping As New EDI2011.FMappingService
    Dim uFASCommon As New EDI2011.FCommonService
    Dim uWFSCommon As New WFS.CommonService
    Dim oWaves As New Waves.CommonService

    Dim NowDateTime As String               ' 現在日時
    Dim NowTime As String           '現在日期時間
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                      ' 設定參數
        If Not IsPostBack Then

            DSProgressBar.Text = Now.ToString("yyyy/MM/dd HH:mm:ss") & "    pFun=ADIDAS 處理中.........."
            WebService()
            DEProgressBar.Text = Now.ToString("yyyy/MM/dd HH:mm:ss") & "    pFun=ADIDAS 處理完成........"

            '限IE11
            Response.Write("<script>window.open('', '_self', ''); window.close();</script>")

        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定參數
    '**
    '*****************************************************************
    Sub SetParameter()
        '-----------------------------------------------------------------
        '-- 系統參數
        '-----------------------------------------------------------------
        Server.ScriptTimeout = 900                                      '設定逾時時間
        Response.Cookies("PGM").Value = "AutoMaterialExpansion.aspx"    '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")
        NowTime = Now.ToString("HH:mm:ss")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(WebService)
    '**     
    '**
    '*****************************************************************
    Sub WebService()
        Dim EndTime As String = ""
        Dim Run As String = ""
        Dim sql As String
        Dim i As Integer

        sql = "Select * From M_Referp "
        sql = sql & " Where Cat  = '" & "200" & "' "
        sql = sql & "   And DKey = '" & "AGENT-MATERIAL" & "' "
        Dim dt_Run As DataTable = uDataBase.GetDataTable(sql)
        If dt_Run.Rows.Count > 0 Then
            Run = Mid(dt_Run.Rows(0).Item("Data"), 1, 1)
        End If

        If Run = "Y" Then
            sql = "Select * From M_Referp "
            sql = sql & " Where Cat  = '" & "200" & "' "
            sql = sql & "   And DKey = '" & "AGENT-MATERIAL-END" & "' "
            Dim dt_EndTime As DataTable = uDataBase.GetDataTable(sql)
            If dt_EndTime.Rows.Count > 0 Then
                EndTime = dt_EndTime.Rows(0).Item("Data")
            End If

            If NowTime > EndTime Then
                sql = "Select Data From M_Referp "
                sql = sql & " Where Cat  = '" & "200" & "' "
                sql = sql & "   And DKey = '" & "MATERIAL-BUYER" & "' "
                sql = sql & "Order by Data "
                Dim dt_Buyer As DataTable = uDataBase.GetDataTable(sql)
                For i = 0 To dt_Buyer.Rows.Count - 1

                    sql = "Update M_Referp Set "
                    sql = sql + " Data = '" & Mid(dt_Buyer.Rows(i).Item("Data"), 1, 6) & "/" & "Y" & "' "
                    sql = sql + " Where Cat = '200' "
                    sql = sql + "   And DKey = 'MATERIAL-BUYER' "
                    sql = sql + "   And Data = '" & dt_Buyer.Rows(i).Item("Data") & "' "
                    uDataBase.ExecuteNonQuery(sql)
                Next
            End If

            uFASCommon.MaterialExpansion()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([800]-MaterialExpansion )
    '**       材料分析BOM展開 ADIDAS, REEBOK
    '***********************************************************************************************
    'MaterialExpansion-Start
    Sub xxMaterialExpansion()
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
            Dim dt_Run As DataTable = uDataBase.GetDataTable(sql)
            If dt_Run.Rows.Count > 0 Then
                Run = Mid(dt_Run.Rows(0).Item("Data"), 1, 1)
            End If

            sql = "Select * From M_Referp "
            sql = sql & " Where Cat  = '" & "200" & "' "
            sql = sql & "   And DKey = '" & "MATERIAL-BUYER" & "' "
            sql = sql & "   And Data Like '%" & "Y" & "%' "
            Dim dt_RunBuyer As DataTable = uDataBase.GetDataTable(sql)
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
                uDataBase.ExecuteNonQuery(sql)
                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                ' 取得對象BUYER
                '
                xBuyerList = ""
                sql = "Select Top 1 Data From M_Referp "
                sql = sql & " Where Cat  = '" & "200" & "' "
                sql = sql & "   And DKey = '" & "MATERIAL-BUYER" & "' "
                sql = sql & "   And Data Like '%" & "Y" & "%' "
                sql = sql & "Order by Data "
                Dim dt_Buyer As DataTable = uDataBase.GetDataTable(sql)
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
                Dim dt_Referp As DataTable = uDataBase.GetDataTable(sql)
                If dt_Referp.Rows.Count > 0 Then
                    xSeasonYY = dt_Referp.Rows(0).Item("Data")
                End If
                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                ' 資料清除
                sql = "Delete From A_MaterialsAnalysis_SLD "
                sql &= "Where Buyer = '" & xBuyerList & "' "
                sql &= "  And Substring(Season, 3,2) >= '" & xSeasonYY & "' "
                uDataBase.ExecuteNonQuery(sql)
                '
                sql = "Delete From A_MaterialsAnalysis_CH "
                sql &= "Where Buyer = '" & xBuyerList & "' "
                sql &= "  And Substring(Season, 3,2) >= '" & xSeasonYY & "' "
                uDataBase.ExecuteNonQuery(sql)
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
                uDataBase.ExecuteNonQuery(sql)
                ' ----------------------------------------------------------------
                ' Trace Log
                ' ----------------------------------------------------------------
                sql = "Delete From W_TraceLog "
                sql &= "Where Pgm = '" & "FAS-Material" & "' "
                uDataBase.ExecuteNonQuery(sql)
                '
                sql = "Insert into W_TraceLog "
                sql &= "( Pgm, Data, CreateTime ) "
                sql &= "VALUES( "
                sql &= " '" & "FAS-Material" & "', "
                sql &= " '" & xBuyerList & "--- Start ---" & "', "
                sql &= " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " ) "
                uDataBase.ExecuteNonQuery(sql)
                ' ----------------------------------------------------------------
                '
                sql = "Select Item From A_MaterialsAnalysis_ZIP "
                sql &= "Where Buyer = '" & xBuyerList & "' "
                sql &= "  And Substring(Season, 3,2) >= '" & xSeasonYY & "' "
                sql &= "Group by Item "
                sql &= "Order by Item "
                Dim dt_ZIPITEM As DataTable = uDataBase.GetDataTable(sql)
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
                        Dim dt_ZIPCOLOR As DataTable = uDataBase.GetDataTable(sql)
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
                            Dim dt_ZIPDATA As DataTable = uDataBase.GetDataTable(sql)
                            For k As Integer = 0 To dt_ZIPDATA.Rows.Count - 1
                                '
                                ' 取得Meter換算基準 (取得ItemClass)
                                oWaves.GetItemClassa(dt_ZIPDATA.Rows(k).Item("Item"), xClass)
                                sql = "Select * From M_Referp "
                                sql = sql & " Where Cat  = '" & "110" & "' "
                                sql = sql & "   And DKey = 'FALL-" & dt_ZIPDATA.Rows(k).Item("Buyer") + "-" + xClass & "' "
                                Dim dt_Meter As DataTable = uDataBase.GetDataTable(sql)
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
                                uDataBase.ExecuteNonQuery(sql)
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
                    uDataBase.ExecuteNonQuery(sql)
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
                        Dim dt_ZIPCOLOR As DataTable = uDataBase.GetDataTable(sql)
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
                            Dim dt_ZIPDATA As DataTable = uDataBase.GetDataTable(sql)
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
                                uDataBase.ExecuteNonQuery(sql)
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
                    uDataBase.ExecuteNonQuery(sql)
                    '
                    ' 更新 ZIP-處理FLAG(Yobi1)
                    sql = "Update A_MaterialsAnalysis_ZIP Set "
                    sql &= "Yobi1 = '" & "*" & "', "
                    sql &= "ModifyUser = '" & "BOM" & "', "
                    sql &= "ModifyTime = '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                    sql &= "Where Buyer = '" & xBuyerList & "' "
                    sql &= "  And Substring(Season, 3,2) >= '" & xSeasonYY & "' "
                    sql &= "  And Item = '" & dt_ZIPITEM.Rows(i).Item("Item") & "' "
                    uDataBase.ExecuteNonQuery(sql)
                Next
                '
                sql = "Insert into W_TraceLog "
                sql &= "( Pgm, Data, CreateTime ) "
                sql &= "VALUES( "
                sql &= " '" & "FAS-Material" & "', "
                sql &= " '" & xBuyerList & "--- End ---" & "', "
                sql &= " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "' "
                sql &= " ) "
                uDataBase.ExecuteNonQuery(sql)
                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                ' 更新系統設定
                '
                sql = "Update M_Referp Set "
                sql = sql + " Data = '" & Run & "/" & Now.ToString("yyyyMMddHHmmss") & "' "
                sql = sql + " Where Cat = '200' "
                sql = sql + "   And DKey = 'AGENT-MATERIAL' "
                uDataBase.ExecuteNonQuery(sql)
                '
                sql = "Update M_Referp Set "
                sql = sql + " Data = '" & xBuyerList & "/" & "N" & "' "
                sql = sql + " Where Cat = '200' "
                sql = sql + "   And DKey = 'MATERIAL-BUYER' "
                sql = sql + "   And Data Like '" & xBuyerList & "%' "
                uDataBase.ExecuteNonQuery(sql)
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
            uDataBase.ExecuteNonQuery(sql)
            ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            ' 異常後重新可處理  處理FLAG = Y
            sql = "Update M_Referp Set "
            sql = sql + " Data = '" & "Y" & "/" & Now.ToString("yyyyMMddHHmmss") & "error" & "' "
            sql = sql & " Where Cat  = '" & "200" & "' "
            sql = sql & "   And DKey = '" & "AGENT-MATERIAL" & "' "
            uDataBase.ExecuteNonQuery(sql)
        End Try
    End Sub
    'MaterialExpansion-End


End Class
