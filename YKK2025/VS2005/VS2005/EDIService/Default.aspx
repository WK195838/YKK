<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>測試頁</title>
</head>
<body>
    <form id="form1" runat="server">
    <div id="DIV1" runat="server">
        <br />
        <asp:Button ID="SPNo2Import" runat="server" Text="SP No2Import" />

        <br />
        <br />
        <asp:Button ID="SPRule2Data" runat="server" Text="SP Rule2Data" />

        <br />
        <asp:Button ID="SPRESET" runat="server" Text="SP RESET" />
        <br />
        <asp:Button ID="SPMakeForcastNo" runat="server" Text="SPMakeForcastNo" />
        <br />
        <asp:Button ID="SPForcastPlan" runat="server" Text="SPForcastPlan" />
        <br />
        <asp:Button ID="SPLocalStockPlan" runat="server" Text="SPLocalStockPlan" />
        <br />
        <asp:Button ID="SPUpdateControlRecord" runat="server" Text="SPUpdateControlRecord" />
        <br />
        <asp:Button ID="DeleteActionData" runat="server" Text="DeleteActionData" />
        <br />
        <asp:Button ID="SPActionkPlan" runat="server" Text="SPActionkPlan" />
        <br />
        <asp:Button ID="SPKPInterface" runat="server" Text="SPKPInterface" />
        <br />
        <asp:Button ID="SPFinalPlan" runat="server" Text="SPFinalPlan" />
        <br />
        <br />

        LogID=<asp:TextBox ID="DLogID" runat="server" BackColor="Yellow" Width="172px">20120912170000</asp:TextBox>    
        <br />
        Buyer=
 <asp:DropDownList ID="DBuyer" runat="server" BackColor="Yellow" ForeColor="Blue" Width="179px">
            <asp:ListItem  Value="E8200-KT-000999">STARITE-KT-OTHER</asp:ListItem>
            <asp:ListItem  Value="D0880-000013">FarEastern-NIKE</asp:ListItem>
            <asp:ListItem Value="A2970-000013">KC-NIKE</asp:ListItem>
            <asp:ListItem Value="A2971-000003">KC-COLUMBIA</asp:ListItem>
            <asp:ListItem Value="A2972-000151">KC-PUMA</asp:ListItem>
            <asp:ListItem Selected="True" Value="A2972-000999">KC-OTHER</asp:ListItem>
            <asp:ListItem Value="H4530-000013">QV-NIKE</asp:ListItem>
            <asp:ListItem Value="H4536-000001">QV-ADIDAS</asp:ListItem>
            <asp:ListItem Value="H4538-000016">QV-REEBOK</asp:ListItem>
            <asp:ListItem Value="H4537-000021">QV-TNF</asp:ListItem>
            <asp:ListItem Value="A4760-000001">SCI-ADIDAS</asp:ListItem>
            <asp:ListItem  Value="A4763-000021">SCI-TNF</asp:ListItem>
            
        </asp:DropDownList>        
        <br />
        UserID=<asp:TextBox ID="DUserID" runat="server" BackColor="Yellow" Width="172px">IT003</asp:TextBox> 
        <br />
        GroupBuyer=<asp:TextBox ID="DGRBuyer" runat="server" BackColor="Yellow" Width="172px">OPXPRTPOO</asp:TextBox> 
        <br />
        PONO<asp:TextBox ID="DPONO" runat="server" BackColor="Yellow" Width="172px">SC11C13020</asp:TextBox> 
        <br />
        Code=<asp:TextBox ID="DCode" runat="server" BackColor="Yellow" Width="172px"></asp:TextBox>&nbsp;


        <asp:Button ID="BCheckGRPC" runat="server" Text="CheckGRPC" /><br />

        <asp:Button ID="ISOSFCTPLAN" runat="server" Text="ISOS FCT PLAN" /><br />

        <asp:Button ID="ISOSLSPLAN" runat="server" Text="ISOS LocalStock PLAN" />        <br />

        <asp:Button ID="BKPI" runat="server" Text="KPI-OVERFLOW FCT" />

        <br />

        <asp:Button ID="BUpdateFC" runat="server" Text="Update FC" />

        <br />
        <asp:Button ID="BDelivery" runat="server" Text="Delivery Date" />
        
        <br />
        <asp:Button ID="EDITransfer" runat="server" Text="EDITransfer" />

        <br />
        <asp:Button ID="LSPlanInf" runat="server" Text="LocalStock PLAN-INF" />

        <br />
        <asp:Button ID="NewLSPlan" runat="server" Text="LocalStock PLAN" />
        <br /><asp:Button ID="Button2" runat="server" Text="MAKE FC-NO" /><br />

        <asp:Button ID="FCTPLAN" runat="server" Text="FCT PLAN" /><br />

        <asp:Button ID="Button3" runat="server" Text="Mat FCT PLAN" /><br />
        <asp:Button ID="Button1" runat="server" Text="VENDOR FC PLAN" />

        <br />
        <asp:Button ID="BBDataCheck" runat="server" Text="BY FCT DataCheck" />

        <br />
        <asp:Button ID="BBYConvert" runat="server" Text="BY FCT Convert" />

        <br />
        <asp:Button ID="BResetBuyerLSNo" runat="server" Text="ResetBuyerLSNo" />
        <br />
        <asp:Button ID="BResetLSNo" runat="server" Text="ResetLSNo" />
        <br />
        <asp:Button ID="BResetFCTNo" runat="server" Text="ResetFCTNo" />

        <br />
        <asp:Button ID="BFASCanRunLFLSPlan" runat="server" Text="FASCanRunLFLSPlan" />
        <br />
        <asp:Button ID="BFASLFLSPlan" runat="server" Text="FASLFLSPlan" />
        <br />
        <asp:Button ID="BFASBULSPlan" runat="server" Text="FASBULSPlan" />
        <br />
        <asp:Button ID="BFASLSPlan" runat="server" Text="FASLSPlan" />
        <br />
        <asp:Button ID="BFASRule2Data" runat="server" Text="FASRule2Data" />
        <br />
        <asp:Button ID="BRule2Data" runat="server" Text="Rule2Data" /><br />
        <asp:Button ID="Rule2DataEOES" runat="server" Text="Rule2DataEOES" />
        <br />
        <br />
        <asp:Button ID="BMakePONO" runat="server" Text="MAKEPONO" />
        <br />
        <asp:Button ID="BCHECKPONO" runat="server" Text="CHECKPONO" />
        <br />
        <asp:Button ID="BMAKEGRPC" runat="server" Text="MAKEGRPC" />&nbsp;<br />
        <br />
        &nbsp;<br />
        <asp:Button ID="BCheckCompanyCode" runat="server" Text="CheckCompanyCode" />
        <br />
        <asp:Button ID="BCheckKeepCode" runat="server" Text="CheckKeepCode" />
        <br />
        <asp:Button ID="BMakeColorCode" runat="server" Text="MakeColorCode" />
        <br />
        <asp:Button ID="BCheckColorCode" runat="server" Text="CheckColorCode" />
        <br />
        <asp:Button ID="BCheckItemCode" runat="server" Text="CheckItemCode" />
        <br />
        <asp:Button ID="BCheckDuplicateData" runat="server" Text="CheckDuplicateData" />
        <br />
        <asp:Button ID="BCheckNikeVDP" runat="server" Text="CheckNikeVDP" />
        <br />
        <asp:Button ID="BCheckPOStructure" runat="server" Text="CheckPOStructure" />
        <br />
        <asp:Button ID="BEDI2Waves" runat="server" Text="EDI2Waves" />
        <br />
        <asp:Button ID="BGetFunctionCode" runat="server" Text="FunctionCode" />
        <br />
        <asp:Button ID="BGETPOSeqno" runat="server" Text="GETPOSeqno" />
        <br />
        <asp:Button ID="BPriceList" runat="server" Text="PriceList" />
        <br />
        <asp:Button ID="BEDI2B2B" runat="server" Text="EDI2B2B" />
        <br />
        <asp:Button ID="BUpdateOrderProgress" runat="server" Text="UpdateOrderProgress" />
        <br />
        <asp:Button ID="BMaterialExpansion" runat="server" Text="MaterialExpansion" />
        <br />
        
        </div>
    </form>
</body>
</html>
