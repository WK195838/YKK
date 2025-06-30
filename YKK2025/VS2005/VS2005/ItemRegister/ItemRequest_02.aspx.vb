Imports System.Data
Imports System.Drawing

Partial Class ItemRequest_02
    Inherits System.Web.UI.Page
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim uID As Integer               'Data Unique_ID
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
            ShowItemRequestData()   '顯示資料
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        Response.Cookies("PGM").Value = "ItemRequest_02.aspx"
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日時
        uID = Request.QueryString("pID")                    'Data Unique_ID
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetReadOnly)
    '**     設定畫面屬性
    '**
    '*****************************************************************
    Sub SetReadOnly(ByVal pType As Boolean)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetButtonFunction)
    '**     功能按鈕顯示
    '**
    '*****************************************************************
    Sub SetButtonFunction()
        '搜尋參考Item
        BFindItem.Attributes("onclick") = "window.open('FindItemPage_01.aspx','FindItemPage','status=0,toolbar=0,width=620,height=650,resizable=yes,scrollbars=yes');"
        '儲存
        BSave.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BSave.Text + "]作業嗎？\n\r請確認...." + "');if(!ok){return false;}"
        '刪除
        BDelete.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BDelete.Text + "]作業嗎？\n\r請確認...." + "');if(!ok){return false;}"
        '關閉
        BClose.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BClose.Text + "]視窗嗎？\n\r請確認...." + "');if(!ok){return false;} else {window.close();}"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowItemRequestData)
    '**     顯示資料
    '**
    '*****************************************************************
    Sub ShowItemRequestData()
        Dim SQL As String = ""
        '
        SQL = "Select * "
        SQL = SQL & "From W_ItemRequest "
        SQL = SQL & "Where Unique_ID = '" & CStr(uID) & "' "
        Dim dt_ItemRequest As DataTable = uDataBase.GetDataTable(SQL)
        If dt_ItemRequest.Rows.Count > 0 Then
            DID.Text = dt_ItemRequest.Rows(0).Item("Unique_ID").ToString.Trim   ' Unique_ID
            DSts.SelectedIndex = dt_ItemRequest.Rows(0).Item("Sts")             ' Status
            '
            DRCode.Text = dt_ItemRequest.Rows(0).Item("RITMC16").ToString.Trim  ' R Code
            DRItemName1.Text = dt_ItemRequest.Rows(0).Item("RIT1I16").ToString.Trim           ' R Item Name-1
            DRItemName2.Text = dt_ItemRequest.Rows(0).Item("RIT2I16").ToString.Trim           ' R Item Name-2
            DRItemName3.Text = dt_ItemRequest.Rows(0).Item("RIT3I16").ToString.Trim           ' R Item Name-3
            DRSize.Text = dt_ItemRequest.Rows(0).Item("RSIZC16").ToString.Trim                ' R Size
            DRChain.Text = dt_ItemRequest.Rows(0).Item("RCHNC16").ToString.Trim               ' R Chain
            DRClass.Text = dt_ItemRequest.Rows(0).Item("RCLSC16").ToString.Trim               ' R Class
            DRTape.Text = dt_ItemRequest.Rows(0).Item("RTAPC16").ToString.Trim                ' R Tape
            DRSlider1.Text = dt_ItemRequest.Rows(0).Item("RSLDC16").ToString.Trim             ' R Slider1
            DRFinish1.Text = dt_ItemRequest.Rows(0).Item("RSFNC16").ToString.Trim             ' R Finish1
            DRSlider2.Text = dt_ItemRequest.Rows(0).Item("RSL2C16").ToString.Trim             ' R Slider2
            DRFinish2.Text = dt_ItemRequest.Rows(0).Item("RSE2C16").ToString.Trim             ' R Finish2
            DRSRequest1.Text = dt_ItemRequest.Rows(0).Item("RSF1C16").ToString.Trim           ' R 特殊要求1
            DRSRequest2.Text = dt_ItemRequest.Rows(0).Item("RSF2C16").ToString.Trim           ' R 特殊要求2
            DRSRequest3.Text = dt_ItemRequest.Rows(0).Item("RSF3C16").ToString.Trim           ' R 特殊要求3
            DRSRequest4.Text = dt_ItemRequest.Rows(0).Item("RSF4C16").ToString.Trim           ' R 特殊要求4
            DRSRequest5.Text = dt_ItemRequest.Rows(0).Item("RSF5C16").ToString.Trim           ' R 特殊要求5
            DRSRequest6.Text = dt_ItemRequest.Rows(0).Item("RSF6C16").ToString.Trim           ' R 特殊要求6
            DRFamily.Text = dt_ItemRequest.Rows(0).Item("RFMLC16").ToString.Trim              ' R Family Code
            DRST1.Text = dt_ItemRequest.Rows(0).Item("RST1C16").ToString.Trim                 ' R 統計區分1
            DRST2.Text = dt_ItemRequest.Rows(0).Item("RST2C16").ToString.Trim                 ' R 統計區分2
            DRST3.Text = dt_ItemRequest.Rows(0).Item("RST3C16").ToString.Trim                 ' R 統計區分3
            DRST4.Text = dt_ItemRequest.Rows(0).Item("RST4C16").ToString.Trim                 ' R 統計區分4
            DRST5.Text = dt_ItemRequest.Rows(0).Item("RST5C16").ToString.Trim                 ' R 統計區分5
            DRST6.Text = dt_ItemRequest.Rows(0).Item("RST6C16").ToString.Trim                 ' R 統計區分6
            DRST7.Text = dt_ItemRequest.Rows(0).Item("RST7C16").ToString.Trim                 ' R 統計區分7
            '---------------------------------------------------------
            DCode.Text = dt_ItemRequest.Rows(0).Item("ITMC16").ToString.Trim                  '  Code
            DItemName1.Text = dt_ItemRequest.Rows(0).Item("IT1I16").ToString.Trim             '  Item Name-1
            DItemName2.Text = dt_ItemRequest.Rows(0).Item("IT2I16").ToString.Trim             '  Item Name-2
            DItemName3.Text = dt_ItemRequest.Rows(0).Item("IT3I16").ToString.Trim             '  Item Name-3
            DSize.Text = dt_ItemRequest.Rows(0).Item("SIZC16").ToString.Trim                  '  Size
            DChain.Text = dt_ItemRequest.Rows(0).Item("CHNC16").ToString.Trim                 '  Chain
            DClass.Text = dt_ItemRequest.Rows(0).Item("CLSC16").ToString.Trim                 '  Class
            DTape.Text = dt_ItemRequest.Rows(0).Item("TAPC16").ToString.Trim                  '  Tape
            DSlider1.Text = dt_ItemRequest.Rows(0).Item("SLDC16").ToString.Trim               '  Slider1
            DFinish1.Text = dt_ItemRequest.Rows(0).Item("SFNC16").ToString.Trim               '  Finish1
            DSlider2.Text = dt_ItemRequest.Rows(0).Item("SL2C16").ToString.Trim               '  Slider2
            DFinish2.Text = dt_ItemRequest.Rows(0).Item("SE2C16").ToString.Trim               '  Finish2
            DSRequest1.Text = dt_ItemRequest.Rows(0).Item("SF1C16").ToString.Trim             '  特殊要求1
            DSRequest2.Text = dt_ItemRequest.Rows(0).Item("SF2C16").ToString.Trim             '  特殊要求2
            DSRequest3.Text = dt_ItemRequest.Rows(0).Item("SF3C16").ToString.Trim             '  特殊要求3
            DSRequest4.Text = dt_ItemRequest.Rows(0).Item("SF4C16").ToString.Trim             '  特殊要求4
            DSRequest5.Text = dt_ItemRequest.Rows(0).Item("SF5C16").ToString.Trim             '  特殊要求5
            DSRequest6.Text = dt_ItemRequest.Rows(0).Item("SF6C16").ToString.Trim             '  特殊要求6
            DFamily.Text = dt_ItemRequest.Rows(0).Item("FMLC16").ToString.Trim                '  Family Code
            DST1.Text = dt_ItemRequest.Rows(0).Item("ST1C16").ToString.Trim                   '  統計區分1
            DST2.Text = dt_ItemRequest.Rows(0).Item("ST2C16").ToString.Trim                   '  統計區分2
            DST3.Text = dt_ItemRequest.Rows(0).Item("ST3C16").ToString.Trim                   '  統計區分3
            DST4.Text = dt_ItemRequest.Rows(0).Item("ST4C16").ToString.Trim                   '  統計區分4
            DST5.Text = dt_ItemRequest.Rows(0).Item("ST5C16").ToString.Trim                   '  統計區分5
            DST6.Text = dt_ItemRequest.Rows(0).Item("ST6C16").ToString.Trim                   '  統計區分6
            DST7.Text = dt_ItemRequest.Rows(0).Item("ST7C16").ToString.Trim                   '  統計區分7
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Click Save Button)
    '**     儲存資料
    '**
    '*****************************************************************
    Protected Sub BSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSave.Click
        Dim sql As String = ""
        '
        sql = "SELECT * From MST_C0100 "
        sql = sql + "WHERE ITMC01 = '" & DRCode.Text & "' "
        Dim dtITEM As DataTable = uDataBase.GetDataTable(sql)
        If dtITEM.Rows.Count > 0 Then
            sql = "Update W_ItemRequest Set "
            '
            sql &= " Sts = '" & CStr(DSts.SelectedIndex) & "', "
            '
            sql &= " SEMI16 = '" & "" & "',"                'EMPLOYEE NAME (SECOND LANG  
            sql &= " SSNN16 = '" & "" & "',"                'SESSION NO.
            'sql &= " IT1I16 = '" & "" & "',"               'ITEM NAME 1
            'sql &= " IT2I16 = '" & "" & "',"               'ITEM NAME 2
            'sql &= " IT3I16 = '" & "" & "',"               'ITEM NAME 3
            '
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
            sql &= " TAPC16 = '" & DTape.Text & "',"                  'TAPE CODE
            sql &= " SIMF16 = '" & dtITEM.Rows(0)("SIMF01").ToString.Trim & "',"                                     'SET ITEM FLAG
            sql &= " SLDC16 = '" & DSlider1.Text & "',"                  'SLIDER CODE
            sql &= " SFNC16 = '" & DFinish1.Text & "',"                  'SLIDER FINISH CODE
            '   
            sql &= " SL2C16 = '" & DSlider2.Text & "',"                  'SLIDER CODE 2
            sql &= " SE2C16 = '" & DFinish2.Text & "',"                  'SLIDER FINISH CODE 2
            sql &= " SF1C16 = '" & DSRequest1.Text & "', "
            sql &= " SF2C16 = '" & DSRequest2.Text & "', "
            sql &= " SF3C16 = '" & DSRequest3.Text & "', "
            '   
            sql &= " SF4C16 = '" & DSRequest4.Text & "', "
            sql &= " SF5C16 = '" & DSRequest5.Text & "', "
            sql &= " SF6C16 = '" & DSRequest6.Text & "', "
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
            sql &= " ICQC16 = '" & dtITEM.Rows(0)("ICQC01").ToString.Trim & "',"                  'ITEM CODE REQUEST CODE
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
            sql &= " Where Unique_ID  =  '" & CStr(uID) & "'"
            '
            uDataBase.ExecuteNonQuery(sql)
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Click Delete Button)
    '**     刪除資料
    '**
    '*****************************************************************
    Protected Sub BDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BDelete.Click
        Dim sql As String = ""
        '
        sql = "Delete From W_ItemRequest "
        sql &= " Where Unique_ID  =  '" & CStr(uID) & "'"
        uDataBase.ExecuteNonQuery(sql)
        '
        Response.Write("<script> window.close(); </script>")
    End Sub
End Class
