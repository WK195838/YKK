<%@ Page Language="VB" AutoEventWireup="false" CodeFile="FAS_TestPage.aspx.vb" Inherits="FAS_TestPage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        Return Code =<asp:TextBox ID="DCode" runat="server" BackColor="Yellow" Width="172px"></asp:TextBox><br />
        Return Data =<asp:TextBox ID="DData" runat="server" BackColor="Yellow" Width="579px" Height="33px" TextMode="MultiLine"></asp:TextBox><br />

        <br />
        ChangeColor<br />
        Depo<asp:TextBox ID="CCDepo" runat="server" BackColor="Yellow" Width="71px">01</asp:TextBox>
        Item<asp:TextBox ID="CCItem" runat="server" BackColor="Yellow" Width="71px">2259170</asp:TextBox>
        OriColor<asp:TextBox ID="CCOriColor" runat="server" BackColor="Yellow" Width="60px">THV16</asp:TextBox>
        <asp:Button ID="BChangeColor" runat="server" Text="FAS-GetChangeColor" />    <br />


        <br />
        GetMaterialKeepCodeInventory<asp:Button ID="GetMaterialKeepCodeInventory" runat="server" Text="GetMaterialKeepCodeInventory" />    
        <br />



        <br />
        StockOutNew2Wings<asp:Button ID="StockOutNew2Wings" runat="server" Text="StockOutNew2Wings" />    
        <br />


        <br />
        StockInNew2Wings<asp:Button ID="StockInNew2Wings" runat="server" Text="StockInNew2Wings" />    
        <br />


        <br />
        DTM_OLD2WINGS<asp:Button ID="Button2" runat="server" Text="DTM_OLD2WINGS" />    
        <br />


        <br />
        EDICheckCOLORCode<asp:Button ID="BCHECKCOLOR" runat="server" Text="checKCOLOR" />    
        <br />



        <br />
        EDICheckKeepCode<asp:Button ID="Bcheckkeepcode" runat="server" Text="checkkeepcode" />    
        <br />


        <br />
        EDICheckItemCode<asp:Button ID="BEDICheckItemCode" runat="server" Text="EDICheckItemCode" />    
        <br />


        <br />
        NEWCOLOR 2 WINGS FA300<asp:Button ID="NewColor2Wings" runat="server" Text="NewColor2Wings" />    
        <br />

        <br />
        UPDATE FC DATAT<asp:Button ID="BUPDATEFCDATA" runat="server" Text="UPDATEFCDATA" />    
        <br />

        <br />
        GetCostPrice<asp:Button ID="BGetCostPrice" runat="server" Text="GetCostPrice" />    

        <br />
        <br />
        GetSalesManCode<asp:Button ID="BSalesManCode" runat="server" Text="GetSalesManCode" />    <br />

        <br />
        GetPurchaseInf<asp:Button ID="BPurchaseInf" runat="server" Text="GetPurchaseInf" />    <br />

        <br />
        GetProductionInf<asp:Button ID="BGetProductionInf" runat="server" Text="GetProductionInf" />    <br />

        <br />
        GetFreeByLocation<asp:Button ID="BGetFreeByLocation" runat="server" Text="GetFreeByLocation" />    <br />

        <br />
        GetMinStock<asp:Button ID="BGetMinStock" runat="server" Text="GetMinStock" />    <br />

        <br />
        GetSpecialChain<asp:Button ID="BSpecialChain" runat="server" Text="GetSpecialChain" />    <br />

        <br />
        <asp:Button ID="BGetChildItemStructure" runat="server" Text="FAS-GetChildItemStructure" /><br />

        <br />
        Item Structure<br />
        Depo<asp:TextBox ID="IDepo" runat="server" BackColor="Yellow" Width="71px">01</asp:TextBox>
        父Item<asp:TextBox ID="IPItem" runat="server" BackColor="Yellow" Width="71px">0233012</asp:TextBox>
        Class(可指定/可不指定)<asp:TextBox ID="IClass" runat="server" BackColor="Yellow" Width="60px">CH</asp:TextBox>
        <asp:Button ID="BGetItemStructure" runat="server" Text="FAS-GetItemStructure" /><br />
        <br />

        Inventory(KeepCode)<br />
        Depo<asp:TextBox ID="IKDepo" runat="server" BackColor="Yellow" Width="71px">01</asp:TextBox>
        Item<asp:TextBox ID="IKItem" runat="server" BackColor="Yellow" Width="71px">1624696</asp:TextBox>
        Color<asp:TextBox ID="IKColor" runat="server" BackColor="Yellow" Width="60px"></asp:TextBox>
        KeepCode(可指定/可不指定)<asp:TextBox ID="IKeepCode" runat="server" BackColor="Yellow" Width="60px">AD-ST</asp:TextBox>
        <asp:Button ID="BKeepCodeInventory" runat="server" Text="FAS-GetKeepCodeInventory" /><br />
        <asp:Button ID="Button1" runat="server" Text="FAS-GetKeepCodeInventoryZIP" /><br />
        <br />

        Inventory(Free)<br />
        Depo<asp:TextBox ID="IFDepo" runat="server" BackColor="Yellow" Width="71px">01</asp:TextBox>
        Item<asp:TextBox ID="IFItem" runat="server" BackColor="Yellow" Width="71px">1624696</asp:TextBox>
        Color<asp:TextBox ID="IFColor" runat="server" BackColor="Yellow" Width="60px"></asp:TextBox>
        <asp:Button ID="BFreeInventory" runat="server" Text="FAS-GetFreeInventory" /><br />
        <br />
        

        Production<br />
        Depo<asp:TextBox ID="PDepo" runat="server" BackColor="Yellow" Width="71px">01</asp:TextBox>
        Item<asp:TextBox ID="PItem" runat="server" BackColor="Yellow" Width="71px">1647008</asp:TextBox>
        Color<asp:TextBox ID="PColor" runat="server" BackColor="Yellow" Width="60px"></asp:TextBox>
        KeepCode<asp:TextBox ID="PKeepCode" runat="server" BackColor="Yellow" Width="60px">AD-RB</asp:TextBox>
        Prod(0:SCHE/1:ON)<asp:TextBox ID="PProd" runat="server" BackColor="Yellow" Width="60px">1</asp:TextBox>
        <asp:Button ID="BProductionQty" runat="server" Text="FAS-GetProductionQty" />    <br />
        <br />
        


    </div>
    </form>
</body>
</html>
