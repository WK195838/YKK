
Partial Class _Default
    Inherits System.Web.UI.Page

    Dim oDB As New ForProject
    Dim oMapping As New MappingService
    Dim oCommon As New CommonService
    Dim oFMapping As New FMappingService
    Dim oFCommon As New FCommonService

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            DLogID.Text = Now.ToString("yyyyMMddHHmmss")
        End If
    End Sub



    Protected Sub BMakePONO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BMakePONO.Click
        Dim xCode As Integer = 9
        '        xCode = oCommon.MakePONO(DLogID.Text, DBuyer.SelectedValue, DUserID.Text, DGRBuyer.Text)
        xCode = oCommon.MakePONO("202003706170000", "EOES-TC020", "IT003", "NPXPRTPOO")

        '        xCode = oFCommon.NewLocalStockPlan(LogID, "FALL-000021", "IT003", "000021B", "91111XXXXX")
        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BCHECKPONO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BCHECKPONO.Click
        Dim xCode As Integer = 9
        xCode = oCommon.CheckPONO("202003706170000", "EOES-TC020", "IT003", "NPXPRTPOO")
        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BMAKEGRPC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BMAKEGRPC.Click
        Dim xCode As Integer = 9
        xCode = oCommon.MakeGRPC(DLogID.Text, DBuyer.SelectedValue, DUserID.Text, DGRBuyer.Text)
        DCode.Text = CStr(xCode)
    End Sub

    '######
    Protected Sub BCheckGRPC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BCheckGRPC.Click
        Dim xCode As Integer = 9
        xCode = oCommon.CheckGRPC("20200706170000", "EOES-TC020", "IT003", "NPXPRTPOO")
        DCode.Text = CStr(xCode)
    End Sub

    '######

    Protected Sub BCheckCompanyCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BCheckCompanyCode.Click
        Dim xCode As Integer = 9
        xCode = oCommon.CheckCompanyCode(DLogID.Text, DBuyer.SelectedValue, DUserID.Text, DGRBuyer.Text)
        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BCheckKeepCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BCheckKeepCode.Click
        Dim xCode As Integer = 9
        xCode = oCommon.CheckKeepCode("20211116170000", "EOES-IT003", "IT003", "NPXPRTPOO")


        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BMakeColorCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BMakeColorCode.Click
        Dim xCode As Integer = 9
        xCode = oCommon.MakeColorCode(DLogID.Text, DBuyer.SelectedValue, DUserID.Text, DGRBuyer.Text)
        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BCheckColorCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BCheckColorCode.Click
        Dim xCode As Integer = 9
        xCode = oCommon.CheckColorCode(DLogID.Text, DBuyer.SelectedValue, DUserID.Text, DGRBuyer.Text)
        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BCheckItemCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BCheckItemCode.Click
        Dim xCode As Integer = 9
        'xCode = oCommon.CheckItemCode(DLogID.Text, DBuyer.SelectedValue, DUserID.Text, DGRBuyer.Text)
        xCode = oCommon.CheckItemCode("20221003173839", "EOES-SL036", "SL036", "OPXPRTPOO")
        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BCheckDuplicateEDIData_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BCheckDuplicateData.Click
        Dim xCode As Integer = 9
        xCode = oCommon.CheckDuplicateData(DLogID.Text, DBuyer.SelectedValue, DUserID.Text, DGRBuyer.Text)
        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BCheckNikeVDP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BCheckNikeVDP.Click
        Dim xCode As Integer = 9
        xCode = oCommon.CheckNikeVDP(DLogID.Text, DBuyer.SelectedValue, DUserID.Text, DGRBuyer.Text)
        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BCheckPOStructure_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BCheckPOStructure.Click
        Dim xCode As Integer = 9
        xCode = oCommon.CheckPOStructure(DLogID.Text, DBuyer.SelectedValue, DUserID.Text, DGRBuyer.Text)
        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BEDI2Waves_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BEDI2Waves.Click
        Dim xCode As Integer = 9
        xCode = oCommon.EDI2Waves(DLogID.Text, DBuyer.SelectedValue, DUserID.Text, DGRBuyer.Text)
        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BGetFunctionCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGetFunctionCode.Click
        DCode.Text = oDB.GetFunctionCode(DGRBuyer.Text, 3)
    End Sub

    Protected Sub BGETPOSeqno_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGETPOSeqno.Click
        Dim xCode As Integer = 1
        xCode = oCommon.GetPOSeqNo(DBuyer.Text, DPONO.Text)
        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BPriceList_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BPriceList.Click
        Dim xCode As Integer = 1
        xCode = oCommon.PriceList(DLogID.Text, DBuyer.SelectedValue, DUserID.Text, DGRBuyer.Text)
        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BEDI2B2B_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BEDI2B2B.Click
        Dim xCode As Integer = 1
        xCode = oCommon.EDI2B2B(DLogID.Text, DBuyer.SelectedValue, DUserID.Text, DGRBuyer.Text)
        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BUpdateOrderProgress_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BUpdateOrderProgress.Click
        Dim xCode As Integer = 1
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oCommon.UpdateOrderProgress(LogID, "D0760-000013", "IT003", "DB12100603-SC")
        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BFASRule2Data_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFASRule2Data.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFMapping.Rule2Data(LogID, "FALL-000001", "IT003", "")
        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BFASLSPlan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFASLSPlan.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.LocalStockPlan(LogID, "FA4760-000001", "IT003", "000001", "1XXXXXXXXX")

        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BFASBULSPlan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFASBULSPlan.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.BuyerLocalStockPlan(LogID, "FALL-TP000013", "IT003", "000013T", "91112XXXXX")

        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub BFASLFLSPlan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFASLFLSPlan.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.LastFinalLocalStockPlan(LogID, "FALL-VENDOR", "IT003", "VENDORB", "9111XXXXXX")

        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub BFASCanRunLFLSPlan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFASCanRunLFLSPlan.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.CanRunLFLocalStockPlan("FA4760-000001", "000001")

        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub BResetFCTNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BResetFCTNo.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")
        xCode = oFCommon.ResetFCTNo("FA4760-000001", "000001")

        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BResetLSNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BResetLSNo.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")
        xCode = oFCommon.ResetLSNo("FA4760-000001", "000001")

        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub BResetBuyerLSNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BResetBuyerLSNo.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")
        xCode = oFCommon.ResetBuyerLSNo("FA4760-000001", "000001")

        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub BBYConvert_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BBYConvert.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")
        xCode = oFCommon.ConvertToBYFCT(LogID, "FALL-000001", "IT003", "000001B", "91111XXXXX")

        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub BBDataCheck_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BBDataCheck.Click

        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")
        xCode = oFCommon.BYFCTDataCheck(LogID, "FALL-000001", "IT003", "000001B", "91111XXXXX")

        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub FCTPLAN_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FCTPLAN.Click
        '################################
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.NewForcastPlan(LogID, "FALL-000001", "IT003", "000001B", "91111XXXXX", 0)


        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub NewLSPlan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NewLSPlan.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.NewLocalStockPlan(LogID, "FALL-000021", "IT003", "000021B", "91111XXXXX")

        DCode.Text = CStr(xCode)
    End Sub



    Protected Sub EDITransfer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles EDITransfer.Click
        Dim xCode As Integer = 9
        Dim LogID As String = "20180130115441"
        
        xCode = oFCommon.EDITransfer(LogID, "FALL-TP000013", "IT003", "000013T", "91112XXXXX")

        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub LSPlanInf_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LSPlanInf.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.LocalStockPlanProdInf(LogID, "FALL-000001", "IT003", "000001B", "911XXXXXXX")


        DCode.Text = CStr(xCode)

    End Sub


    Protected Sub BMaterialExpansion_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BMaterialExpansion.Click
        Dim xCode As Integer = 9

        oFCommon.MaterialExpansion()

        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub BDelivery_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BDelivery.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oMapping.Rule2Data(LogID, "A2972-000999", "IT003", "OXXXXTPOO")

        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub BRule2Data_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BRule2Data.Click
        Dim xCode As Integer = 9
        'xCode = oMapping.Rule2Data(DLogID.Text, DBuyer.SelectedValue, DUserID.Text, DGRBuyer.Text)
        xCode = oMapping.Rule2Data("20190523093641", "W9520-000999", "IT003", "OPXPRTPOX")
        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub Rule2DataEOES_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Rule2DataEOES.Click
        Dim xCode As Integer = 9
        'xCode = oMapping.Rule2Data(DLogID.Text, DBuyer.SelectedValue, DUserID.Text, DGRBuyer.Text)


        xCode = oMapping.Rule2DataEOES("20220116170000", "EOES-IT003", "IT003", "NPXPRTPOO")
        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.VendorForcastPlan(LogID, "F-VENDOR-SL022", "SL022", "VINPUTB", "911XXXXXXX", 0)


        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.MakeForcastNo(LogID, "FALL-VENDOR", "IT003", "VENDORB", "9111XXXXXX", 0)

        DCode.Text = CStr(xCode)


    End Sub

    Protected Sub BUpdateFC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BUpdateFC.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.MakeForcastNo(LogID, "FALL-VENDOR", "IT003", "VENDORB", "9111XXXXXX", 0)

        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        'xCode = oFCommon.MatForcastPlan(LogID, "FALL-TW0655", "IT003", "TW0655B", "91114XXXXX", 0)


        'xCode = oFCommon.MakeForcastNo(LogID, "FALL-TW0371T", "IT003", "TW0371T", "91114XXXXX", 0)

        xCode = oFCommon.MatForcastPlanUABAG(LogID, "FALL-TW0371T", "IT003", "TW0371T", "91115XXXXX", 0)


        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub BKPI_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BKPI.Click

        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.KPIExpansion(LogID, "FALL-000001", "IT003", "000001B", "91111XXXXX")


        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub ISOSFCTPLAN_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ISOSFCTPLAN.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.ISOS2FAS(LogID, "FALL-000021", "IT003", "000001B", "91111XXXXX", 0)

        'xCode = oFCommon.ForcastPlanISOS(LogID, "FALL-000021", "IT003", "000001B", "91111XXXXX", 0)

        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub ISOSLSPLAN_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ISOSLSPLAN.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.LocalStockPlanISOS(LogID, "FALL-000001", "IT003", "000001B", "91111XXXXX")

        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub SPRESET_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SPRESET.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.SPUpdateControlRecord("SHA-AD", "IT003", "Reset", 0, "Reset", 0)

        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub SPMakeForcastNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SPMakeForcastNo.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.SPMakeForcastNo(LogID, "SHA-AD", "IT003", "000071", "XXXXXXXXXX", 0)

        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub SPForcastPlan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SPForcastPlan.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.SPForcastPlan(LogID, "SHA-NK-B", "IT003", "000075", "XXXXXXXXXX", 0)


        DCode.Text = CStr(xCode)


    End Sub

    Protected Sub SPLocalStockPlan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SPLocalStockPlan.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.SPLocalStockPlan(LogID, "SHA-AD", "IT003", "000071", "XXXXXXXXXX")

        DCode.Text = CStr(xCode)

    End Sub


    Protected Sub SPUpdateControlRecord_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SPUpdateControlRecord.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.SPUpdateControlRecord("SHA-AD-B", "IT003", "Demand", 2, "ActPlan", 1)

        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub DeleteActionData_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DeleteActionData.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        'xCode = oFCommon.DeleteActionData("SHA-AD")
        xCode = oFCommon.DeleteActionData("FALL-000001")

        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub SPActionkPlan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SPActionkPlan.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        'xCode = oFCommon.SPActionkPlan(LogID, "SHA-AD", "IT003", "000071", "XXXXXXXXXX")
        xCode = oFCommon.SPActionkPlan(LogID, "SHA-AD", "IT003", "000071", "XXXXXXXXXX")

        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub SPKPInterface_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SPKPInterface.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.SPKPInterface(LogID, "SHA-AD-B", "IT003", "000071", "XXXXXXXXXX")

        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub SPFinalPlan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SPFinalPlan.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.SPFinalPlan(LogID, "SHA-AD", "IT003", "000071", "XXXXXXXXXX")

        DCode.Text = CStr(xCode)
    End Sub

    Protected Sub SPRule2Data_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SPRule2Data.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFMapping.Rule2Data(LogID, "THA-AD-B", "IT003", "XXXXXXXXX")

        DCode.Text = CStr(xCode)

    End Sub

    Protected Sub SPNo2Import_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SPNo2Import.Click
        Dim xCode As Integer = 9
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        xCode = oFCommon.SPNo2Import(LogID, "THA-AD-Y", "IT003", "000042", "XXXXXXXXXX", 0)

        DCode.Text = CStr(xCode)

    End Sub
End Class
