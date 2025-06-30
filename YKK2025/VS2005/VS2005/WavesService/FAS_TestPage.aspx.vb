Imports System.Data

Partial Class FAS_TestPage
    Inherits System.Web.UI.Page

    Dim oCommon As New CommonService
    Dim oEDI As New EDIService

    Protected Sub BGetItemStructure_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGetItemStructure.Click
        Dim pItem(10) As String
        Dim pItemName(10) As String
        Dim pQty(10) As String
        Dim pCount, i As Integer
        DData.Text = ""

        DCode.Text = oCommon.GetItemStructure(IDepo.Text, IPItem.Text, IClass.Text, pItem, pItemName, pQty, pCount)

        For i = 1 To pCount
            If DData.Text <> "" Then
                DData.Text = DData.Text + Chr(13) + CStr(i) + ":" + _
                            pItem(i) + "/" + _
                            pItemName(i) + "/" + _
                            CStr(CDbl(pQty(i)) / 10000000)
            Else
                DData.Text = CStr(i) + ":" + _
                             pItem(i) + "/" + _
                             pItemName(i) + "/" + _
                             CStr(CDbl(pQty(i)) / 10000000)
            End If
        Next
    End Sub

    Protected Sub BKeepCodeInventory_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BKeepCodeInventory.Click
        Dim pQty As String
        DData.Text = ""

        DCode.Text = oCommon.GetKeepCodeInventory(IKDepo.Text, IKItem.Text, IKColor.Text, IKeepCode.Text, pQty)

        DData.Text = CStr(CDbl(pQty) / 10000000)
    End Sub

    Protected Sub BFreeInventory_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFreeInventory.Click
        Dim pQty As String
        DData.Text = ""

        DCode.Text = oCommon.GetFreeInventory(IFDepo.Text, IFItem.Text, IFColor.Text, pQty)

        DData.Text = CStr(CDbl(pQty) / 10000000)
    End Sub

    Protected Sub BProductionQty_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BProductionQty.Click
        Dim pQty As String
        Dim xMonProd(6) As String        ' N1 ~ N4 SCHE_PRODQTY, ON_PRODQTY
        DData.Text = ""

        DCode.Text = oCommon.GetProductionQty("01", "1709051", "", "AD-RB", 0, pQty, xMonProd)



        DData.Text = "[" + CStr(CDbl(pQty) / 10000000) + "]/[" + CStr(CDbl(xMonProd(1)) / 10000000) + "/" + CStr(CDbl(xMonProd(2)) / 10000000) + "/" + _
                     CStr(CDbl(xMonProd(3)) / 10000000) + "/" + CStr(CDbl(xMonProd(4)) / 10000000) + "/" + CStr(CDbl(xMonProd(5)) / 10000000) + "/" + _
                     CStr(CDbl(xMonProd(6)) / 10000000) + "]"


    End Sub

    Protected Sub BChangeColor_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BChangeColor.Click
        Dim pColor As String
        DData.Text = ""

        DCode.Text = oCommon.GetChangeColor("01", "8235250", "6368657", "TNP38", pColor)

        DData.Text = pColor

    End Sub

    Protected Sub BSpecialChain_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSpecialChain.Click
        DData.Text = ""

        DCode.Text = oCommon.GetSpecialChain("1906469", "T8/CFT8/CNT8/;T9/CFT9/CNT9/;T10/CFT10/CNT10/;DT1/CFDT1/CNDT1/;DT2/CFDT2/CNDT2/;BP/BP12/BP14/BP15/BP16/BP18/;PBR/PBR12/PBR14/PBR16/;")

    End Sub

    Protected Sub BGetChildItemStructure_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGetChildItemStructure.Click
        '@@    Public Function GetChildItemStructure(ByVal pDepo As String, ByVal pObjectProduct As String, ByVal pPItem As String, _
        '                                            ByRef pItem() As String, ByRef pItemName() As String, ByRef pQty() As String, ByRef pCount As Integer) As Integer
        Dim xItem(5), xItemName(5), xQty(5) As String
        Dim xcount As Integer

        'DCode.Text = CStr(oCommon.GetChildItemStructure("01", "PULLER- Z", "3800148", xItem, xItemName, xQty, xcount))
        '3800148

        'DCode.Text = CStr(oCommon.GetChildItemStructure("01", "CH-DYED", "3946315", xItem, xItemName, xQty, xcount))
        'DCode.Text = CStr(oCommon.GetChildItemStructure("01", "TAPE-SET", "6346064", xItem, xItemName, xQty, xcount))
        DCode.Text = CStr(oCommon.GetChildItemStructure("01", "TAPE-DYED", "6743463", xItem, xItemName, xQty, xcount))

        '3759172

        'DCode.Text = CStr(oCommon.GetChildItemStructure("01", "SLIDERPART-ZPULLER", "1746206", xItem, xItemName, xQty, xcount))

        'DCode.Text = CStr(oCommon.GetChildItemStructure("01", "SLIDERPART- Z", "3879892", xItem, xItemName, xQty, xcount))

        'DCode.Text = CStr(oCommon.GetChildItemStructure("01", "TAPE-NAT", "0262092", xItem, xItemName, xQty, xcount))


        'MsgBox("ITEM=[" + xItem(1) + "]")
        'MsgBox("ITEMNAME=[" + xItemName(1) + "]")
        'MsgBox("QTY=[" + xQty(1) + "]")

        For i As Integer = 1 To 5
            'MsgBox("ITEM=[" + xItem(i) + "]")
            'MsgBox("ITEMNAME=[" + xItemName(i) + "]")
            'MsgBox("QTY=[" + xQty(i) + "]")
        Next
    End Sub

    Protected Sub BGetMinStock_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGetMinStock.Click
        Dim xQty As String = ""

        DData.Text = ""
        DCode.Text = oCommon.GetMininumStock("01", "0273662", "", xQty)
        DData.Text = xQty
    End Sub

    Protected Sub BGetFreeByLocation_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGetFreeByLocation.Click
        Dim xDescr As String = ""

        DData.Text = ""
        DCode.Text = oCommon.GetFreeByLocation("01", "1456741", "", xDescr)

        '1060402 580:1645496

        DData.Text = xDescr
    End Sub

    Protected Sub BGetProductionInf_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGetProductionInf.Click
        Dim xProdInf(50) As String
        Dim xcount As Integer

        DCode.Text = oCommon.GetProductionInf("01", "1458867", "", "AD-RB", xProdInf, xcount)
        MsgBox("Count=" + CStr(xcount))

        For i As Integer = 0 To xcount - 1
            MsgBox(CStr(i) + "=[" + xProdInf(i) + "]")
        Next
    End Sub

    Protected Sub BPurchaseInf_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BPurchaseInf.Click
        Dim xProdInf(50) As String
        Dim xcount As Integer

        DCode.Text = oCommon.GetPurchaseInf("01", "1094947", "  030", "AD-RB", xProdInf, xcount)
        MsgBox("Count=" + CStr(xcount))

        For i As Integer = 0 To xcount - 1
            MsgBox(CStr(i) + "=[" + xProdInf(i) + "]")
        Next

    End Sub

    Protected Sub BSalesManCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSalesManCode.Click
        Dim xSalesManCode As String = ""

        DCode.Text = oCommon.GetSalesManCode("A4760", xSalesManCode)

        DData.Text = "[" + xSalesManCode + "]"
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim pQty As String
        DData.Text = ""

        DCode.Text = oCommon.GetKeepCodeInventoryZIP("01", "2218870", "  580", "6", "I", "SHEI-KP1", pQty)

        DData.Text = CStr(CDbl(pQty) / 10000000)
    End Sub

    Protected Sub BGetCostPrice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGetCostPrice.Click
        Dim xPriceA, xPriceB As String
        DData.Text = ""

        DCode.Text = oCommon.GetCostPrice("2147982", xPriceA, xPriceB)

        DData.Text = "A=" + CStr(CDbl(xPriceA) / 1000000) + " / B=" + CStr(CDbl(xPriceB) / 1000000)

    End Sub

    Protected Sub BUPDATEFCDATA_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BUPDATEFCDATA.Click
        DData.Text = ""

        DCode.Text = oEDI.UpdateVendorFC("F-VENDOR-SL022", "SL022")

        DData.Text = "OK"
        If DCode.Text <> "0" Then
            DData.Text = "NG"
        End If

    End Sub
    '----------------------------
    Protected Sub NewColor2Wings_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NewColor2Wings.Click
        DData.Text = ""

        DCode.Text = oCommon.NewColor2Wings("A2021020002")

        DData.Text = "OK"
        If DCode.Text <> "0" Then
            DData.Text = "NG"
        End If

    End Sub
    '----------------------------
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub BEDICheckItemCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BEDICheckItemCode.Click
        DData.Text = ""

        DCode.Text = oEDI.EDICheckItemCode("20210316111105", "EOES-SL013", "SL013", "OPXRXTPOO")

        DData.Text = "OK"
        If DCode.Text <> "0" Then
            DData.Text = "NG"
        End If

    End Sub

    Protected Sub Bcheckkeepcode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bcheckkeepcode.Click
        DData.Text = ""

        DCode.Text = oEDI.EDICheckKeepCode("20210421132933", "EOES-TC010", "TC010", "NPXPRTPOO")

        DData.Text = "OK"
        If DCode.Text <> "0" Then
            DData.Text = "NG"
        End If

    End Sub

    Protected Sub BCHECKCOLOR_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BCHECKCOLOR.Click
        DData.Text = ""

        DCode.Text = oEDI.EDICheckColorCode("20210421120436", "EOES-TC010", "TC010", "OPXRXTPOO")

        DData.Text = "OK"
        If DCode.Text <> "0" Then
            DData.Text = "NG"
        End If

    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        DData.Text = ""

        DCode.Text = oCommon.OldColor2Wings("L2021090079")

        DData.Text = "OK"
        If DCode.Text <> "0" Then
            DData.Text = "NG"
        End If

    End Sub

    Protected Sub StockInNew2Wings_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles StockInNew2Wings.Click
        DData.Text = ""

        DCode.Text = oCommon.StockInNew2Wings("PR14040732")

        DData.Text = "OK"
        If DCode.Text <> "0" Then
            DData.Text = "NG"
        End If

    End Sub

    Protected Sub StockOutNew2Wings_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles StockOutNew2Wings.Click
        DData.Text = ""

        DCode.Text = oCommon.StockOutNew2Wings("02130000000009")

        DData.Text = "OK"
        If DCode.Text <> "0" Then
            DData.Text = "NG"
        End If

    End Sub

    Protected Sub GetMaterialKeepCodeInventory_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles GetMaterialKeepCodeInventory.Click
        Dim str As String
        str = ""

        DCode.Text = oCommon.GetMaterialKeepCodeInventory("01", "6721700", "", str)

        DData.Text = str

    End Sub
End Class
