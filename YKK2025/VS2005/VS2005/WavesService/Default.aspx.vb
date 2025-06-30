Imports System.Data

Partial Class _Default
    Inherits System.Web.UI.Page

    Dim oCommon As New CommonService


    Protected Sub BGetCustomerList_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGetCustomerList.Click
        Dim ds As New DataSet
        '
        DCode.Text = CStr(oCommon.GetCustomerList(DCustomer.Text, ds))
    End Sub

    Protected Sub BGetStardandPrice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGetStardandPrice.Click
        Dim a, b, c, d As Single

        DCode.Text = CStr(oCommon.GetStandardPrice(DType.Text, DVersion.Text, DCurrency.Text, DItem.Text, a, b, c, d))

    End Sub

    Protected Sub BGetOrderPrice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGetOrderPrice.Click

        DCode.Text = oCommon.GetOrderPrice("A2970 ", _
                                                      "000013", _
                                                      "K120009919", _
                                                      1, _
                                                      DOrderNo1.Text, _
                                                      DItem1.Text, _
                                                      DColor1.Text, _
                                                      DPrice1.Text)
    End Sub

    Protected Sub BGetOrderPriceDetail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGetOrderPriceDetail.Click
        DCode.Text = oCommon.GetOrderPriceDetail("OR04011862", _
                                             1, _
                                             DPriceVersion.Text, _
                                             DOrderQty.Text, _
                                             DListPrice.Text, _
                                             DSalesPrice.Text, _
                                             DSalesAmount.Text)


    End Sub


    Protected Sub BGetStatus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGetStatus.Click
        DCode.Text = oCommon.GetDescriptionMaster("PSTC", "MM060", DPriceVersion.Text)

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub


    Protected Sub BCHECKDUPLICATION_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BCHECKDUPLICATION.Click
        DCode.Text = oCommon.CheckDuplicateData("111", "AAAA")
    End Sub
End Class
