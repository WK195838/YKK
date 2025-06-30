<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
            Customer=<asp:TextBox ID="DCustomer" runat="server" BackColor="Yellow" Width="172px">H45</asp:TextBox>    
        <br />
            Type=<asp:TextBox ID="DType" runat="server" BackColor="Yellow" Width="172px">A</asp:TextBox>    
            Version=<asp:TextBox ID="DVersion" runat="server" BackColor="Yellow" Width="172px">211</asp:TextBox>    
            Currency=<asp:TextBox ID="DCurrency" runat="server" BackColor="Yellow" Width="172px">TWD</asp:TextBox>    
            Item=<asp:TextBox ID="DItem" runat="server" BackColor="Yellow" Width="172px"></asp:TextBox>    
        <br />
        <asp:TextBox ID="DOrderNo1" runat="server" BackColor="Yellow" Width="172px">A</asp:TextBox>
        <asp:TextBox ID="DItem1" runat="server" BackColor="Yellow" Width="172px">A</asp:TextBox>
        <asp:TextBox ID="DColor1" runat="server" BackColor="Yellow" Width="172px">A</asp:TextBox>
        <asp:TextBox ID="DPrice1" runat="server" BackColor="Yellow" Width="172px">A</asp:TextBox>
        <br />
        <asp:TextBox ID="DPriceVersion" runat="server" BackColor="Yellow" Width="172px">A</asp:TextBox>
        <asp:TextBox ID="DOrderQty" runat="server" BackColor="Yellow" Width="172px">A</asp:TextBox>
        <asp:TextBox ID="DListPrice" runat="server" BackColor="Yellow" Width="172px">A</asp:TextBox>
        <asp:TextBox ID="DSalesPrice" runat="server" BackColor="Yellow" Width="172px">A</asp:TextBox>
        <asp:TextBox ID="DSalesAmount" runat="server" BackColor="Yellow" Width="172px">A</asp:TextBox>
        <br />
       
        Code=<asp:TextBox ID="DCode" runat="server" BackColor="Yellow" Width="172px"></asp:TextBox> 
        <br />
        <asp:Button ID="BGetCustomerList" runat="server" Text="GetCustomerList" />
        <br />
        <asp:Button ID="BGetStardandPrice" runat="server" Text="GetStardandPrice" />&nbsp;<br /><asp:Button ID="BGetOrderPrice" runat="server" Text="GetOrderPrice" />
        <br />
        <asp:Button ID="BGetOrderPriceDetail" runat="server" Text="GetOrderPriceDetail" />
        <br />
        <asp:Button ID="BGetOrderProgress" runat="server" Text="GetOrderProgress" />
        <br />
        <asp:Button ID="BGetStatus" runat="server" Text="GetStatus" />
        <br />
        <asp:Button ID="BGetPackList" runat="server" Text="GetPackList" /><br />
        <br />
        <asp:Button ID="BCHECKDUPLICATION" runat="server" Text="CHECKDUPLICATION" /></div>
    </form>
</body>
</html>
