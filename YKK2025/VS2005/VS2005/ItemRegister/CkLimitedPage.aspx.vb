Imports System.Data
Imports System.Drawing

Partial Class CkLimitedPage
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim wUserID As String = ""      '申請者

    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim sUniqueID As String
    Dim rowItem As DataRow
    Dim dtResult As DataTable

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900 ' 設定逾時時間
        SetParameter()             ' 設定共用參數
        getItemData()              ' 從W_LimitedItemRegister取得限定錯誤的Item資料
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日時
        Response.Cookies("PGM").Value = "CkLimitedPage.aspx"

    End Sub

    '從W_LimitedItemRegister取得限定錯誤的Item資料
    Sub getItemData()
        sUniqueID = Request.QueryString("id")
        Dim sql As String
        sql = "SELECT ItemName1, ItemName2, ItemName3, Size, Chain, Class, Tape, Slider1, Finish1, Slider2, " & _
                "Finish2, SRequest1, SRequest2, SRequest3, SRequest4, SRequest5, SRequest6, Family, " & _
                "ST1, ST2, ST3, ST4, ST5, ST6, ST7, A001, A206, A211, A999, K206, K211, " & _
                "Customer, CustomerCode, Buyer, BuyerCode, Remark, ForUse " & _
                "FROM W_LimitedItemRegister WHERE Unique_ID='" & sUniqueID & "'"
        DLimitedItem.Text = "Unique_ID:" & sUniqueID
        Dim dtItem As DataTable
        dtItem = uDataBase.GetDataTable(sql)
        If dtItem.Rows.Count > 0 Then
            rowItem = dtItem.Rows(0)
            DCustomer.Text = rowItem.Item("Customer")
            DCustomerCode.Text = rowItem.Item("CustomerCode")
            DBuyer.Text = rowItem.Item("Buyer")
            DBuyerCode.Text = rowItem.Item("BuyerCode")
            DForUse.Text = rowItem.Item("ForUse")
            DRemark.Text = rowItem.Item("Remark")
            DA001.Checked = rowItem.Item("A001") = "1"
            DA206.Checked = rowItem.Item("A206") = "1"
            DA211.Checked = rowItem.Item("A211") = "1"
            DA999.Checked = rowItem.Item("A999") = "1"
            DK206.Checked = rowItem.Item("K206") = "1"
            DK211.Checked = rowItem.Item("K211") = "1"
            '執行限定檢查
            If UCase(Request.QueryString("msg")) = "NEW" Then
                LimitedItemError("ALL")         '所有限定檢查
            Else
                If DCustomerCode.Text = "" And DBuyerCode.Text = "" Then
                    LimitedItemError("ITEM")    'Item限定
                Else
                    LimitedItemError("")        'Buyer+Customer限定
                End If
            End If
        Else
            uJavaScript.PopMsg(Me, "找不到Item限定檢查資料! UniqueID:" & sUniqueID)
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(LimitedItemError)
    '**     檢查BUYER或ITEM限定 pCat='BUYER' OR 'ITEM'
    '**
    '*****************************************************************
    Public Function LimitedItemError(ByVal pCat As String) As Boolean
        Dim i, j As Integer
        Dim x, y As Integer
        Dim sql, msg, xSpc, sqlISIP As String
        Dim str, str1, str2 As String
        Dim uError As Boolean = False
        Dim uRun As Boolean = False
        Dim iNew As String()
        Dim ckTime As Integer

        If rowItem Is Nothing Then
            Return False
        End If
        '
        xSpc = GetFullName(rowItem.Item("SRequest1").ToString()) & "!" & _
                GetFullName(rowItem.Item("SRequest2").ToString()) & "!" & _
                GetFullName(rowItem.Item("SRequest3").ToString()) & "!" & _
                GetFullName(rowItem.Item("SRequest4").ToString()) & "!" & _
                GetFullName(rowItem.Item("SRequest5").ToString()) & "!" & _
                GetFullName(rowItem.Item("SRequest6").ToString()) & "!"
        '
        str = rowItem.Item("Size").ToString() & "!" & _
                rowItem.Item("Chain").ToString() & "!" & _
                rowItem.Item("Class").ToString() & "!" & _
                rowItem.Item("Tape").ToString() & "!" & _
                rowItem.Item("Slider1").ToString() & "!" & _
                rowItem.Item("Finish1").ToString() & "!" & _
                rowItem.Item("Slider2").ToString() & "!" & _
                rowItem.Item("Finish2").ToString() & "!" & _
                rowItem.Item("SRequest1").ToString() & "!" & _
                rowItem.Item("Family").ToString() & "!" & _
                rowItem.Item("ST1").ToString() & "!" & _
                rowItem.Item("ST2").ToString() & "!" & _
                rowItem.Item("ST3").ToString() & "!" & _
                rowItem.Item("ST4").ToString() & "!" & _
                rowItem.Item("ST5").ToString() & "!" & _
                rowItem.Item("ST6").ToString() & "!" & _
                rowItem.Item("ST7").ToString() & "!"
        '
        'TEST

        'xSpc = "BOX!ND-B!P-PACK!XXX!!!"
        'xSpc = "GREEN-F!KENSIN!N-ANTI!P-TB!SLS-N412!SLS-TP926"
        'xSpc = "GREEN-F!KENSIN!N-ANTI!P-TB!REVERSE!SLS156" 
        'str = "05!RC!PS!!DF42TM54-B!VK!!!BOX!RC!1!2!2!3!1!P!4!"
        'str = "03!CF!C!PB12!DABHH054AJ!EFJ!!!GREEN-F!C!1!2!2!1!1!!!"
        'str = "03!CF!C!PB12!DABHH054P!EFJ!!!GREEN-F!C!1!2!2!1!1!!!"
        DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "xSpc:" & xSpc & Chr(13) & "str:" & str
        '
        iNew = str.ToString.Split("!")
        '
        ' SLIDER
        '去除字尾SK及-B關鍵字
        DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "原來       Slider1:" & iNew(4) & "/Slider2:" & iNew(6)
        iNew(4) = fpObj.ReplaceSliderString(iNew(4))
        iNew(6) = fpObj.ReplaceSliderString(iNew(6))
        DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "替換SK-B後 Slider1:" & iNew(4) & "/Slider2:" & iNew(6)
        '
        If Not uError Then
            sql = "SELECT "
            sql = sql & "[Size], [Size]+'/'+[SizeA_A]+'/'+[SizeA_C]+'/'+[SizeB_A]+'/'+[SizeB_C]+'/'+[SizeC_A]+'/'+[SizeC_C]+'/', "            '0
            sql = sql & "[Chain], [Chain]+'/'+[ChainA_A]+'/'+[ChainA_C]+'/'+[ChainB_A]+'/'+[ChainB_C]+'/'+[ChainC_A]+'/'+[ChainC_C]+'/', "     '2
            sql = sql & "[Class], [Class]+'/'+[ClassA_A]+'/'+[ClassA_C]+'/'+[ClassB_A]+'/'+[ClassB_C]+'/'+[ClassC_A]+'/'+[ClassC_C]+'/', "     '4
            sql = sql & "[Tape], [Tape]+'/'+[TapeA_A]+'/'+[TapeA_C]+'/'+[TapeB_A]+'/'+[TapeB_C]+'/'+[TapeC_A]+'/'+[TapeC_C]+'/', "            '6
            sql = sql & "[Slider1], [Slider1]+'/'+[SLIDER1A_A]+'/'+[SLIDER1A_C]+'/'+[SLIDER1B_A]+'/'+[SLIDER1B_C]+'/'+[SLIDER1C_A]+'/'+[SLIDER1C_C]+'/', "       '8
            sql = sql & "[Finish1], [Finish1]+'/'+[Finish1A_A]+'/'+[Finish1A_C]+'/'+[Finish1B_A]+'/'+[Finish1B_C]+'/'+[Finish1C_A]+'/'+[Finish1C_C]+'/', "       '10
            sql = sql & "[Slider2], [Slider2]+'/'+[Slider2A_A]+'/'+[Slider2A_C]+'/'+[Slider2B_A]+'/'+[Slider2B_C]+'/'+[Slider2C_A]+'/'+[Slider2C_C]+'/', "       '12
            sql = sql & "[Finish2], [Finish2]+'/'+[Finish2A_A]+'/'+[Finish2A_C]+'/'+[Finish2B_A]+'/'+[Finish2B_C]+'/'+[Finish2C_A]+'/'+[Finish2C_C]+'/', "       '14
            sql = sql & "[SRequestList], [SRequestList]+'/'+[SRequestListA_A]+'/'+[SRequestListA_C]+'/'+[SRequestListB_A]+'/'+[SRequestListB_C]+'/'+[SRequestListC_A]+'/'+[SRequestListC_C]+'/', "        '16
            sql = sql & "[Family], [Family]+'/'+[FamilyA_A]+'/'+[FamilyA_C]+'/'+[FamilyB_A]+'/'+[FamilyB_C]+'/'+[FamilyC_A]+'/'+[FamilyC_C]+'/', "      '18
            sql = sql & "[ST1], [ST1]+'/'+[ST1A_A]+'/'+[ST1A_C]+'/'+[ST1B_A]+'/'+[ST1B_C]+'/'+[ST1C_A]+'/'+[ST1C_C]+'/', "       '20
            sql = sql & "[ST2], [ST2]+'/'+[ST2A_A]+'/'+[ST2A_C]+'/'+[ST2B_A]+'/'+[ST2B_C]+'/'+[ST2C_A]+'/'+[ST2C_C]+'/', "       '22
            sql = sql & "[ST3], [ST3]+'/'+[ST3A_A]+'/'+[ST3A_C]+'/'+[ST3B_A]+'/'+[ST3B_C]+'/'+[ST3C_A]+'/'+[ST3C_C]+'/', "       '24
            sql = sql & "[ST4], [ST4]+'/'+[ST4A_A]+'/'+[ST4A_C]+'/'+[ST4B_A]+'/'+[ST4B_C]+'/'+[ST4C_A]+'/'+[ST4C_C]+'/', "       '26
            sql = sql & "[ST5], [ST5]+'/'+[ST5A_A]+'/'+[ST5A_C]+'/'+[ST5B_A]+'/'+[ST5B_C]+'/'+[ST5C_A]+'/'+[ST5C_C]+'/', "       '28
            sql = sql & "[ST6], [ST6]+'/'+[ST6A_A]+'/'+[ST6A_C]+'/'+[ST6B_A]+'/'+[ST6B_C]+'/'+[ST6C_A]+'/'+[ST6C_C]+'/', "       '30
            sql = sql & "[ST7], [ST7]+'/'+[ST7A_A]+'/'+[ST7A_C]+'/'+[ST7B_A]+'/'+[ST7B_C]+'/'+[ST7C_A]+'/'+[ST7C_C]+'/', "       '32
            '
            sql = sql & "[Active],[Cat],[Rno],[SubNo], "                '34
            sql = sql & "[Action], [ACTION_A], [ACTION_C], [msg], Unique_ID "      '38
            '
            sql = sql & "From W_BUYERLIMITED "
            sql = sql & "where Active = 1 "
            '
            sqlISIP = sql & "and Cat = '4.ITEM' and SUBSTRING([Rno],1,1)='M' "        '只抓ISIP轉入的ITEM限定資料
            '
            'PERFORMACE-START
            '[ST1] ~ [ST7]
            'sql = sql & "AND ( [ST1] = '*' OR [ST1] = '" & iNew(10) & "' ) "
            'sql = sql & "AND ( [ST2] = '*' OR [ST2] = '" & iNew(11) & "' ) "
            'sql = sql & "AND ( [ST3] = '*' OR [ST3] = '" & iNew(12) & "' ) "
            'sql = sql & "AND ( [ST4] = '*' OR [ST4] = '" & iNew(13) & "' ) "
            'sql = sql & "AND ( [ST5] = '*' OR [ST5] = '" & iNew(14) & "' ) "
            'sql = sql & "AND ( [ST6] = '*' OR [ST6] = '" & iNew(15) & "' ) "
            'sql = sql & "AND ( [ST7] = '*' OR [ST7] = '" & iNew(16) & "' ) "
            '
            If pCat = "ALL" Then
                sql = sql & "and Cat != '9.EDX' and SUBSTRING([Rno],1,1)!='M' "         'Item+Customer+Buyer+Warning限定檢查(除了9.EDX以外)
            ElseIf pCat = "ITEM" Then
                sql = sql & "and Cat = '4.ITEM' and SUBSTRING([Rno],1,1)!='M' "         'Item限定檢查
            Else
                sql = sql & "and CAT NOT IN ('4.ITEM','5.WARNING','9.EDX') "            'Customer+Buyer限定檢查
                sqlISIP = ""
            End If
            '
            'PERFORMACE-END
            '
            sql = sql & "order by Active,Cat,Rno,SubNo "
            'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & sql
            '
            ckTime = 0
            Dim dtRule As DataTable = uDataBase.GetDataTable(sql)
            '
            '取得ISIP轉入ITEM限定的規則
            If sqlISIP <> "" Then
                'Item限定中針對ISIP轉入特定SLIDER(PULLER+COLOR)筆數多的資料只取得Slider相符的資料
                sqlISIP = sqlISIP & " and ('" & iNew(4) & "' like Slider1 "
                If iNew(6) <> "" Then
                    sqlISIP = sqlISIP & " or '" & iNew(6) & "' like Slider2 "
                End If
                sqlISIP = sqlISIP & ") order by Active,Cat,Rno,SubNo "
                Dim dtRuleISIP As DataTable = uDataBase.GetDataTable(sqlISIP)
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & Chr(13) & "限定筆數:" & dtRule.Rows.Count & "/ ISIP限定筆數:" & dtRuleISIP.Rows.Count
                dtRule.Merge(dtRuleISIP)    '合併限定規則資料
            End If
            '
            dtResult = dtRule.Clone()
            If dtRule.Rows.Count > 0 Then

                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & Chr(13) & "限定檢測-START"

                For i = 0 To dtRule.Rows.Count - 1
                    uError = False
                    ckTime = ckTime + 1
                    'MsgBox(pCat & "--" & CStr(dtRule.Rows.Count) & dtRule.Rows(i)(37))
                    ' -----------------------------------------------------------------
                    ' 是否限定ITEM
                    ' -----------------------------------------------------------------
                    uRun = False
                    j = 0
                    If (dtRule.Rows(i)(j) = "*" Or LimitedItem(iNew(j), dtRule.Rows(i)(j + 1)) = True) And _
                       (dtRule.Rows(i)(j + 2) = "*" Or LimitedItem(iNew(j + 1), dtRule.Rows(i)(j + 3)) = True) And _
                       (dtRule.Rows(i)(j + 4) = "*" Or LimitedItem(iNew(j + 2), dtRule.Rows(i)(j + 5)) = True) And _
                       (dtRule.Rows(i)(j + 6) = "*" Or LimitedItem(iNew(j + 3), dtRule.Rows(i)(j + 7)) = True) And _
                       (dtRule.Rows(i)(j + 8) = "*" Or LimitedItem(iNew(j + 4), dtRule.Rows(i)(j + 9)) = True) And _
                       (dtRule.Rows(i)(j + 10) = "*" Or LimitedItem(iNew(j + 5), dtRule.Rows(i)(j + 11)) = True) And _
                       (dtRule.Rows(i)(j + 12) = "*" Or LimitedItem(iNew(j + 6), dtRule.Rows(i)(j + 13)) = True) And _
                       (dtRule.Rows(i)(j + 14) = "*" Or LimitedItem(iNew(j + 7), dtRule.Rows(i)(j + 15)) = True) And _
                       (dtRule.Rows(i)(j + 16) = "*" Or LimitedItem(xSpc, dtRule.Rows(i)(j + 17)) = True) And _
                       (dtRule.Rows(i)(j + 18) = "*" Or LimitedItem(iNew(j + 9), dtRule.Rows(i)(j + 19)) = True) And _
                       (dtRule.Rows(i)(j + 20) = "*" Or LimitedItem(iNew(j + 10), dtRule.Rows(i)(j + 21)) = True) And _
                       (dtRule.Rows(i)(j + 22) = "*" Or LimitedItem(iNew(j + 11), dtRule.Rows(i)(j + 23)) = True) And _
                       (dtRule.Rows(i)(j + 24) = "*" Or LimitedItem(iNew(j + 12), dtRule.Rows(i)(j + 25)) = True) And _
                       (dtRule.Rows(i)(j + 26) = "*" Or LimitedItem(iNew(j + 13), dtRule.Rows(i)(j + 27)) = True) And _
                       (dtRule.Rows(i)(j + 28) = "*" Or LimitedItem(iNew(j + 14), dtRule.Rows(i)(j + 29)) = True) And _
                       (dtRule.Rows(i)(j + 30) = "*" Or LimitedItem(iNew(j + 15), dtRule.Rows(i)(j + 31)) = True) And _
                       (dtRule.Rows(i)(j + 32) = "*" Or LimitedItem(iNew(j + 16), dtRule.Rows(i)(j + 33)) = True) Then
                        uRun = True
                    End If
                    '
                    ' -----------------------------------------------------------------
                    ' 檢查 限定BUYER, CUSTOMER, ITEM
                    ' -----------------------------------------------------------------
                    If uRun Then

                        Select Case dtRule.Rows(i)(35)
                            Case "1.BUYER"
                                If DBuyerCode.Text <> "" Then
                                    str = DBuyerCode.Text
                                    Select Case dtRule.Rows(i)(39)
                                        Case "!%"
                                            If InStr(dtRule.Rows(i)(40), str) > 0 Then uError = True
                                        Case "%"
                                            If InStr(dtRule.Rows(i)(40), str) <= 0 Then uError = True
                                        Case "!="
                                            If dtRule.Rows(i)(40) = str Then uError = True
                                        Case "="
                                            If dtRule.Rows(i)(40) <> str Then uError = True
                                        Case Else
                                    End Select
                                End If
                            Case "2.CUSTOMER"
                                If DCustomerCode.Text <> "" Then
                                    str = DCustomerCode.Text
                                    Select Case dtRule.Rows(i)(39)
                                        Case "!%"
                                            If InStr(dtRule.Rows(i)(40), str) > 0 Then uError = True
                                        Case "%"
                                            If InStr(dtRule.Rows(i)(40), str) <= 0 Then uError = True
                                        Case "!="
                                            If dtRule.Rows(i)(40) = str Then uError = True
                                        Case "="
                                            If dtRule.Rows(i)(40) <> str Then uError = True
                                        Case Else
                                    End Select
                                End If
                            Case "3.BUY-CUST"
                                'TW5759-G2600/000151-ALL/ALL-000036/
                                '!% 任一有含ERROR
                                '%  全部不含ERROR
                                '!= 任一等於 ERROR
                                '=  全部不等於 ERRPR
                                If DBuyerCode.Text <> "" And DCustomerCode.Text <> "" Then
                                    str = DBuyerCode.Text & "-" & DCustomerCode.Text
                                    str1 = DBuyerCode.Text & "-" & "ALL"
                                    str2 = "ALL" & "-" & DCustomerCode.Text

                                    Select Case dtRule.Rows(i)(39)
                                        Case "!%"
                                            If InStr(dtRule.Rows(i)(40), str) > 0 Or InStr(dtRule.Rows(i)(40), str1) > 0 Or InStr(dtRule.Rows(i)(40), str2) > 0 Then uError = True
                                        Case "%"
                                            If InStr(dtRule.Rows(i)(40), str) <= 0 And InStr(dtRule.Rows(i)(40), str1) <= 0 And InStr(dtRule.Rows(i)(40), str2) <= 0 Then uError = True
                                        Case "!="
                                            If dtRule.Rows(i)(40) = str Or dtRule.Rows(i)(40) = str1 Or dtRule.Rows(i)(40) = str2 Then uError = True
                                        Case "="
                                            If dtRule.Rows(i)(40) <> str And dtRule.Rows(i)(40) <> str1 And dtRule.Rows(i)(40) <> str2 Then uError = True
                                        Case Else
                                    End Select
                                End If
                            Case "5.WARNING"
                                If dtRule.Rows(i)(39) <> "" And dtRule.Rows(i)(40) <> "" And InStr(dtRule.Rows(i)(40).ToString(), "[") > 0 _
                                    And InStr(dtRule.Rows(i)(40).ToString(), "]") > 0 Then
                                    Dim strCmp() = dtRule.Rows(i)(40).ToString().Trim().Substring(1, Len(dtRule.Rows(i)(40).ToString().Trim()) - 2).Split("|")
                                    Dim strFD1 As String
                                    Dim strFD2 As String
                                    '取得比對欄位的值
                                    strFD1 = fpObj.GetCompareVal(str, strCmp(0))
                                    strFD2 = fpObj.GetCompareVal(str, strCmp(1))
                                    Select Case dtRule.Rows(i)(39)
                                        Case "!="
                                            If strFD1 <> strFD2 Then uError = True
                                        Case "="
                                            If strFD1 = strFD2 Then uError = True
                                        Case Else
                                    End Select
                                Else
                                    uError = True
                                End If
                            Case "4.ITEM"
                                uError = True
                            Case Else
                        End Select
                        '
                        DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & dtRule.Rows(i)(35) & "-" & dtRule.Rows(i)(36) & "-" & dtRule.Rows(i)(37) & "]"
                        If dtRule.Rows(i)(35) <> "4.ITEM" And dtRule.Rows(i)(35) <> "5.WARNING" Then
                            DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & dtRule.Rows(i)(40) & "-" & str & "]"
                        End If
                        '
                        If uError = True Then
                            '
                            DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "限定中" & "]"
                            '
                            ' --------------------------------------------------------------------------------------
                            ' 例外檢測
                            ' --------------------------------------------------------------------------------------
                            '
                            sql = "SELECT "
                            sql = sql & "[Size], [Size]+'/'+[SizeA_A]+'/'+[SizeA_C]+'/'+[SizeB_A]+'/'+[SizeB_C]+'/'+[SizeC_A]+'/'+[SizeC_C]+'/', "            '0
                            sql = sql & "[Chain], [Chain]+'/'+[ChainA_A]+'/'+[ChainA_C]+'/'+[ChainB_A]+'/'+[ChainB_C]+'/'+[ChainC_A]+'/'+[ChainC_C]+'/', "     '2
                            sql = sql & "[Class], [Class]+'/'+[ClassA_A]+'/'+[ClassA_C]+'/'+[ClassB_A]+'/'+[ClassB_C]+'/'+[ClassC_A]+'/'+[ClassC_C]+'/', "     '4
                            sql = sql & "[Tape], [Tape]+'/'+[TapeA_A]+'/'+[TapeA_C]+'/'+[TapeB_A]+'/'+[TapeB_C]+'/'+[TapeC_A]+'/'+[TapeC_C]+'/', "            '6
                            sql = sql & "[Slider1], [Slider1]+'/'+[SLIDER1A_A]+'/'+[SLIDER1A_C]+'/'+[SLIDER1B_A]+'/'+[SLIDER1B_C]+'/'+[SLIDER1C_A]+'/'+[SLIDER1C_C]+'/', "       '8
                            sql = sql & "[Finish1], [Finish1]+'/'+[Finish1A_A]+'/'+[Finish1A_C]+'/'+[Finish1B_A]+'/'+[Finish1B_C]+'/'+[Finish1C_A]+'/'+[Finish1C_C]+'/', "       '10
                            sql = sql & "[Slider2], [Slider2]+'/'+[Slider2A_A]+'/'+[Slider2A_C]+'/'+[Slider2B_A]+'/'+[Slider2B_C]+'/'+[Slider2C_A]+'/'+[Slider2C_C]+'/', "       '12
                            sql = sql & "[Finish2], [Finish2]+'/'+[Finish2A_A]+'/'+[Finish2A_C]+'/'+[Finish2B_A]+'/'+[Finish2B_C]+'/'+[Finish2C_A]+'/'+[Finish2C_C]+'/', "       '14
                            sql = sql & "[SRequestList], [SRequestList]+'/'+[SRequestListA_A]+'/'+[SRequestListA_C]+'/'+[SRequestListB_A]+'/'+[SRequestListB_C]+'/'+[SRequestListC_A]+'/'+[SRequestListC_C]+'/', "        '16
                            sql = sql & "[Family], [Family]+'/'+[FamilyA_A]+'/'+[FamilyA_C]+'/'+[FamilyB_A]+'/'+[FamilyB_C]+'/'+[FamilyC_A]+'/'+[FamilyC_C]+'/', "      '18
                            sql = sql & "[ST1], [ST1]+'/'+[ST1A_A]+'/'+[ST1A_C]+'/'+[ST1B_A]+'/'+[ST1B_C]+'/'+[ST1C_A]+'/'+[ST1C_C]+'/', "       '20
                            sql = sql & "[ST2], [ST2]+'/'+[ST2A_A]+'/'+[ST2A_C]+'/'+[ST2B_A]+'/'+[ST2B_C]+'/'+[ST2C_A]+'/'+[ST2C_C]+'/', "       '22
                            sql = sql & "[ST3], [ST3]+'/'+[ST3A_A]+'/'+[ST3A_C]+'/'+[ST3B_A]+'/'+[ST3B_C]+'/'+[ST3C_A]+'/'+[ST3C_C]+'/', "       '24
                            sql = sql & "[ST4], [ST4]+'/'+[ST4A_A]+'/'+[ST4A_C]+'/'+[ST4B_A]+'/'+[ST4B_C]+'/'+[ST4C_A]+'/'+[ST4C_C]+'/', "       '26
                            sql = sql & "[ST5], [ST5]+'/'+[ST5A_A]+'/'+[ST5A_C]+'/'+[ST5B_A]+'/'+[ST5B_C]+'/'+[ST5C_A]+'/'+[ST5C_C]+'/', "       '28
                            sql = sql & "[ST6], [ST6]+'/'+[ST6A_A]+'/'+[ST6A_C]+'/'+[ST6B_A]+'/'+[ST6B_C]+'/'+[ST6C_A]+'/'+[ST6C_C]+'/', "       '30
                            sql = sql & "[ST7], [ST7]+'/'+[ST7A_A]+'/'+[ST7A_C]+'/'+[ST7B_A]+'/'+[ST7B_C]+'/'+[ST7C_A]+'/'+[ST7C_C]+'/', "       '32
                            sql = sql & "[OTHER], [OTHER]+'/'+[OTHERA_A]+'/'+[OTHERA_C]+'/'+[OTHERB_A]+'/'+[OTHERB_C]+'/'+[OTHERC_A]+'/'+[OTHERC_C]+'/' "       '34
                            '
                            sql = sql & "From W_EXCEPTLIMITED "
                            sql = sql & "where Active = 1 "
                            sql = sql & "and Cat = '" & dtRule.Rows(i)(35) & "' "
                            sql = sql & "and Rno = '" & dtRule.Rows(i)(36) & "' "
                            sql = sql & "and SubNo = '" & dtRule.Rows(i)(37) & "' "
                            '
                            '如果SubNo是SLS(PULLER是HH054)，只取得Slider相符的資料
                            If InStr(dtRule.Rows(i)(37), "SLS-1") > 0 Then
                                sql = sql & " and '" & iNew(4) & "' like Slider1 "
                            ElseIf InStr(dtRule.Rows(i)(37), "SLS-2") > 0 Then
                                sql = sql & " and '" & iNew(6) & "' like Slider2 "
                            End If
                            '
                            Dim xSpcExcep As String
                            xSpcExcep = xSpc
                            'PULLER是HH054'，要把特殊要求中的顏色做關鍵字Replace，再跟限定規則比對
                            If InStr(dtRule.Rows(i)(37), "SLS") > 0 Then
                                xSpcExcep = fpObj.ReplaceColorString(xSpc)
                                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "Replace顏色後spec:" & xSpcExcep
                            End If
                            '
                            sql = sql & "order by Active,Cat,Rno,SubNo "
                            'DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "sql:" & sql
                            '
                            Dim dtExceptRule As DataTable = uDataBase.GetDataTable(sql)
                            '
                            If dtExceptRule.Rows.Count > 0 Then
                                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & ">>例外檢測-START"
                                '
                                For x = 0 To dtExceptRule.Rows.Count - 1
                                    ' -----------------------------------------------------------------
                                    ' 是否限定ITEM
                                    ' -----------------------------------------------------------------
                                    y = 0
                                    If (dtExceptRule.Rows(x)(y) = "*" Or LimitedItem(iNew(y), dtExceptRule.Rows(x)(y + 1)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 2) = "*" Or LimitedItem(iNew(y + 1), dtExceptRule.Rows(x)(y + 3)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 4) = "*" Or LimitedItem(iNew(y + 2), dtExceptRule.Rows(x)(y + 5)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 6) = "*" Or LimitedItem(iNew(y + 3), dtExceptRule.Rows(x)(y + 7)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 8) = "*" Or LimitedItem(iNew(y + 4), dtExceptRule.Rows(x)(y + 9)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 10) = "*" Or LimitedItem(iNew(y + 5), dtExceptRule.Rows(x)(y + 11)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 12) = "*" Or LimitedItem(iNew(y + 6), dtExceptRule.Rows(x)(y + 13)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 14) = "*" Or LimitedItem(iNew(y + 7), dtExceptRule.Rows(x)(y + 15)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 16) = "*" Or LimitedItem(xSpcExcep, dtExceptRule.Rows(x)(y + 17)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 18) = "*" Or LimitedItem(iNew(y + 9), dtExceptRule.Rows(x)(y + 19)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 20) = "*" Or LimitedItem(iNew(y + 10), dtExceptRule.Rows(x)(y + 21)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 22) = "*" Or LimitedItem(iNew(y + 11), dtExceptRule.Rows(x)(y + 23)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 24) = "*" Or LimitedItem(iNew(y + 12), dtExceptRule.Rows(x)(y + 25)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 26) = "*" Or LimitedItem(iNew(y + 13), dtExceptRule.Rows(x)(y + 27)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 28) = "*" Or LimitedItem(iNew(y + 14), dtExceptRule.Rows(x)(y + 29)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 30) = "*" Or LimitedItem(iNew(y + 15), dtExceptRule.Rows(x)(y + 31)) = True) And _
                                       (dtExceptRule.Rows(x)(y + 32) = "*" Or LimitedItem(iNew(y + 16), dtExceptRule.Rows(x)(y + 33)) = True) Then
                                        '
                                        'MsgBox("[" & dtExceptRule.Rows(x)(y + 34) & "]")
                                        If dtExceptRule.Rows(x)(y + 34) <> "*" Then
                                            'PRICE A001/A206/A211/A999/K206/K211
                                            If DA001.Checked = True Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), DA001.Text) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                            If DA206.Checked = True Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), DA206.Text) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                            If DA211.Checked = True Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), DA211.Text) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                            If DA999.Checked = True Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), DA999.Text) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                            If DK206.Checked = True Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), DK206.Text) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                            If DK211.Checked = True Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), DK211.Text) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                            'BUYER/CUSTOMER/BUYER-CUSTOMER
                                            str = ""
                                            If DBuyerCode.Text <> "" And DCustomerCode.Text <> "" Then
                                                str = DBuyerCode.Text & "-" & DCustomerCode.Text
                                            Else
                                                If DBuyerCode.Text <> "" Then
                                                    str = DBuyerCode.Text
                                                Else
                                                    If DCustomerCode.Text <> "" Then
                                                        str = DCustomerCode.Text
                                                    End If
                                                End If
                                            End If
                                            If str <> "" Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), str) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                            '用途DForUse
                                            If DForUse.Text <> "" Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), DForUse.Text) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                            '備註DRemark
                                            If DRemark.Text <> "" Then
                                                If InStr(dtExceptRule.Rows(x)(y + 34), DRemark.Text) > 0 Then
                                                    DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                                    uError = False
                                                End If
                                            End If
                                        Else
                                            DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[例外中]"
                                            uError = False
                                        End If
                                    End If
                                    '只要符合一筆例外就跳出白名單迴圈
                                    If uError = False Then
                                        Exit For
                                    End If
                                Next
                                '
                                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & ">>例外檢測-END"
                            End If
                        Else
                            DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "[" & "無限定" & "]"
                        End If
                        '
                        If uError = True Then
                            msg = dtRule.Rows(i)(41)
                            '把錯誤資料存入dtResult中
                            DLimitedItem.Text = DLimitedItem.Text & Chr(13) & dtRule.Rows(i)(42) & ": " & msg & Chr(13)
                            dtResult.ImportRow(dtRule.Rows(i))
                            'AppendLimitedItem()         '限定ERROR --> ADD DB
                            '
                            'Exit For 不跳出，做完全部檢查
                        End If
                    End If
                Next
                DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "限定檢測-END"
            End If
        End If
        '
        'If uError Then uJavaScript.PopMsg(Me, msg)

        DLimitedItem.Text = DLimitedItem.Text & Chr(13) & "Check: " & ckTime.ToString() & "次" & Chr(13) & "Error: " & dtResult.Rows.Count.ToString() & "筆"
        DataList()
        '
        Return uError
    End Function
    '------------------------------------------

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ItemError)
    '**     檢查限定是否限定ITEM
    '**
    '*****************************************************************
    Public Function LimitedItem(ByVal pApply As String, ByVal pDataStr As String) As Boolean
        Dim uLimitedItem As Boolean = True
        Dim p0, p1A, p1C, p2A, p2C, p3A, p3C, sql As String
        Dim xDataStr As String()
        Dim i As Integer
        '
        xDataStr = pDataStr.Split("/")
        For i = 0 To xDataStr.Length - 1
            Select Case i
                Case 0
                    p0 = xDataStr(i)
                Case 1
                    p1A = xDataStr(i)
                Case 2
                    p1C = xDataStr(i)
                Case 3
                    p2A = xDataStr(i)
                Case 4
                    p2C = xDataStr(i)
                Case 5
                    p3A = xDataStr(i)
                Case 6
                    p3C = xDataStr(i)
                Case Else
            End Select
        Next
        '
        If p0 <> "*" Then
            'CONDITION-1
            If uLimitedItem = True Then
                Select Case p1A
                    Case "!%"
                        sql = "select "
                        sql = sql & "ISNULL( (SELECT TOP 1 1 FROM M_Referp WHERE '" & pApply & "' LIKE '" & ReplaceString(p1C) & "'), 0) AS WK "
                        Dim dtWK As DataTable = uDataBase.GetDataTable(sql)
                        If dtWK.Rows(0)("WK") = 1 Then uLimitedItem = False
                    Case "%"
                        sql = "select "
                        sql = sql & "ISNULL( (SELECT TOP 1 1 FROM M_Referp WHERE '" & pApply & "' LIKE '" & ReplaceString(p1C) & "'), 0) AS WK "
                        Dim dtWK As DataTable = uDataBase.GetDataTable(sql)
                        If dtWK.Rows(0)("WK") = 0 Then uLimitedItem = False
                    Case "!="
                        If pApply = ReplaceString(p1C) Then uLimitedItem = False
                    Case "="
                        If pApply <> ReplaceString(p1C) Then uLimitedItem = False
                    Case Else
                End Select
            End If
            'CONDITION-2
            If uLimitedItem = True Then
                Select Case p2A
                    Case "!%"
                        sql = "select "
                        sql = sql & "ISNULL( (SELECT TOP 1 1 FROM M_Referp WHERE '" & pApply & "' LIKE '" & ReplaceString(p2C) & "'), 0) AS WK "
                        Dim dtWK As DataTable = uDataBase.GetDataTable(sql)
                        If dtWK.Rows(0)("WK") = 1 Then uLimitedItem = False
                    Case "%"
                        sql = "select "
                        sql = sql & "ISNULL( (SELECT TOP 1 1 FROM M_Referp WHERE '" & pApply & "' LIKE '" & ReplaceString(p2C) & "'), 0) AS WK "
                        Dim dtWK As DataTable = uDataBase.GetDataTable(sql)
                        If dtWK.Rows(0)("WK") = 0 Then uLimitedItem = False
                    Case "!="
                        If pApply = ReplaceString(p2C) Then uLimitedItem = False
                    Case "="
                        If pApply <> ReplaceString(p2C) Then uLimitedItem = False
                    Case Else
                End Select
            End If
            'CONDITION-3
            If uLimitedItem = True Then
                Select Case p3A
                    Case "!%"
                        sql = "select "
                        sql = sql & "ISNULL( (SELECT TOP 1 1 FROM M_Referp WHERE '" & pApply & "' LIKE '" & ReplaceString(p3C) & "'), 0) AS WK "
                        Dim dtWK As DataTable = uDataBase.GetDataTable(sql)
                        If dtWK.Rows(0)("WK") = 1 Then uLimitedItem = False
                    Case "%"
                        sql = "select "
                        sql = sql & "ISNULL( (SELECT TOP 1 1 FROM M_Referp WHERE '" & pApply & "' LIKE '" & ReplaceString(p3C) & "'), 0) AS WK "
                        Dim dtWK As DataTable = uDataBase.GetDataTable(sql)
                        If dtWK.Rows(0)("WK") = 0 Then uLimitedItem = False
                    Case "!="
                        If pApply = ReplaceString(p3C) Then uLimitedItem = False
                    Case "="
                        If pApply <> ReplaceString(p3C) Then uLimitedItem = False
                    Case Else
                End Select
            End If
        End If
        '
        Return uLimitedItem
    End Function

    Sub DataList()
        If dtResult.Rows.Count <= 0 Then
            uJavaScript.PopMsg(Me, "限定檢查完成，沒有限定錯誤 ! ")
        Else
            '取得User設定的規則表的資料顯示在畫面上
            Dim i As Integer
            Dim sql As String
            sql = ""
            For i = 0 To dtResult.Rows.Count - 1
                sql = sql & "(Cat = '" & dtResult.Rows(i).Item("Cat").ToString() & "' and "
                sql = sql & "Rno = '" & dtResult.Rows(i).Item("Rno").ToString() & "' and "
                sql = sql & "SubNo = '" & dtResult.Rows(i).Item("SubNo").ToString() & "') or "
            Next
            sql = "SELECT Rno, SubNo, MSG " & _
                    "FROM M_BUYERLIMITED WHERE Active=1 AND " & sql.Remove(sql.Length - 4, 3) & _
                    "GROUP BY Rno, SubNo, MSG ORDER BY Rno, SubNo, MSG"
            Dim dtShow As DataTable = uDataBase.GetDataTable(sql)
            GridView1.DataSource = dtShow
            GridView1.DataBind()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ReplaceFamily)
    '**     
    '**
    '*****************************************************************
    Public Function GetFullName(ByVal pShortName As String) As String
        Dim Sql As String
        Dim xFullName As String
        '
        xFullName = ""
        Sql = "Select FullName From M_LimitedReviationList "
        Sql &= "Where Active = 1 "
        Sql &= "And ShortName = '" & pShortName & "' "
        Dim dt_ReviationList As DataTable = uDataBase.GetDataTable(Sql)
        If dt_ReviationList.Rows.Count > 0 Then
            xFullName = dt_ReviationList.Rows(0).Item("FullName")
        Else
            xFullName = ""
        End If
        '
        If xFullName = "" Then xFullName = pShortName
        '
        Return xFullName
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ReplaceString)
    '**     置換限定ITEM關鍵字
    '**
    '*****************************************************************
    Public Function ReplaceString(ByVal pData As String) As String
        Dim i As Integer
        Dim str As String = pData
        Dim xFrom As String = "[SP]/"
        Dim xTo As String = "/"
        Dim xFromWord As String()
        Dim xToWord As String()
        '
        xFromWord = xFrom.Split("/")
        xToWord = xTo.Split("/")
        For i = 0 To xFromWord.Length - 1
            str = Replace(str, xFromWord(i), xToWord(i))
        Next
        '
        Return str
    End Function


End Class
