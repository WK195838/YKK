<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CkLimitedPage.aspx.vb" Inherits="CkLimitedPage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>限定檢查</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="TextBox6" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 8px; position: absolute;
            top: 16px; text-align: left" Width="88px">CUSTOMER</asp:TextBox> 
        <asp:TextBox ID="DCustomerCode" runat="server" BackColor="White" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" Style="z-index: 126; left: 96px; position: absolute; 
            top: 16px; text-align: left" Width="64px" MaxLength="6" ReadOnly="True" Enabled="False"></asp:TextBox>
        <asp:TextBox ID="DCustomer" runat="server" BackColor="White" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" ReadOnly="True" Style="z-index: 105; left: 168px;
            position: absolute; top: 16px" Width="152px" Enabled="False"></asp:TextBox>
        <asp:TextBox ID="TextBox7" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 8px; position: absolute;
            top: 41px; text-align: left" Width="88px">BUYER</asp:TextBox> 
        <asp:TextBox ID="DBuyerCode" runat="server" BackColor="White" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" Style="z-index: 126; left: 96px; position: absolute; 
            top: 41px; text-align: left" Width="64px" MaxLength="6" ReadOnly="True" Enabled="False"></asp:TextBox>
        <asp:TextBox ID="DBuyer" runat="server" BackColor="White" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 108; left: 168px; position: absolute;
            top: 41px" Width="152px" Enabled="False"></asp:TextBox>
        <asp:TextBox ID="TextBox3" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 8px; position: absolute;
            top: 66px; text-align: left" Width="88px">單價設定</asp:TextBox> 
                <asp:CheckBox ID="DA211" runat="server" Font-Size="9pt" Style="z-index: 174;
            left: 192px; position: absolute; top: 64px" Text="A211" Width="48px" Enabled="False" />
        <asp:CheckBox ID="DK206" runat="server" Font-Size="9pt" Style="z-index: 174; left: 288px;
            position: absolute; top: 66px" Text="K206" Width="48px" Enabled="False" />
        <asp:CheckBox ID="DK211" runat="server" Font-Size="9pt" Style="z-index: 174;
            left: 336px; position: absolute; top: 66px" Text="K211" Width="48px" Enabled="False" />
        <asp:CheckBox ID="DA206" runat="server" Font-Size="9pt" Style="z-index: 174; left: 144px;
            position: absolute; top: 66px" Text="A206" Width="48px" Enabled="False" />
        <asp:CheckBox ID="DA999" runat="server" Font-Size="9pt" Style="z-index: 174; left: 240px;
            position: absolute; top: 66px" Text="A999" Width="48px" Enabled="False" />
        <asp:CheckBox ID="DA001" runat="server" Font-Size="9pt" Style="z-index: 174; left: 96px;
            position: absolute; top: 66px" Text="A001" Width="48px" Enabled="False" />
        <asp:TextBox ID="TextBox1" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 8px; position: absolute;
            top: 91px; text-align: left" Width="88px">用途</asp:TextBox>
        <asp:TextBox ID="DForUse" runat="server" BackColor="White" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" ReadOnly="True" Style="z-index: 126; left: 96px;
            position: absolute; top: 91px" Width="320px" Enabled="False"></asp:TextBox>
        <asp:TextBox ID="TextBox2" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 8px; position: absolute;
            top: 116px; text-align: left" Width="88px">備註</asp:TextBox>
        <asp:TextBox ID="DRemark" runat="server" BackColor="White" BorderStyle="Groove" ForeColor="Blue"
            Height="43px" MaxLength="100" Style="z-index: 126; left: 96px; position: absolute;
            top: 116px; text-align: left" TextMode="MultiLine" Width="328px" Enabled="False" ReadOnly="True"></asp:TextBox>
        <asp:HyperLink ID="ReferenceLink" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="https://ykkglobal.sharepoint.com/sites/asia_twn_discuss_mktsal/Lists/IRW1/DispForm.aspx?ID=15&e=B3W7Td"
            Style="z-index: 124; left: 8px; position: absolute; top: 172px" Target="_blank"
            Width="176px">BUYER限定販售ITEM表</asp:HyperLink>
            
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            BorderColor="Black" BorderStyle="Solid" BorderWidth="2px" CellPadding="4" Font-Size="9pt"
            PageSize="25" Style="z-index: 114; left: 8px; position: absolute; top: 200px"
            Width="424px">
            <RowStyle BackColor="White" ForeColor="Blue" />
            <Columns>
                <asp:BoundField DataField="Rno" HeaderText="Rno">
                    <ItemStyle Width="40px" />
                </asp:BoundField>
                <asp:BoundField DataField="SubNo" HeaderText="SubNo">
                    <ItemStyle Width="35px" />
                </asp:BoundField>
                <asp:BoundField DataField="MSG" HeaderText="Item限定/Buyer限定錯誤訊息" />
            </Columns>
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#FF8000" Font-Bold="True" ForeColor="White" BorderStyle="Solid" BorderWidth="2px" />
            <FooterStyle BorderStyle="Groove" />
        </asp:GridView>
        <asp:TextBox ID="DLimitedItem" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
            Font-Bold="False" Font-Size="Medium" ForeColor="Black" Height="616px" Style="z-index: 103;
            left: 440px; position: absolute; top: 208px" Width="296px" TextMode="MultiLine" Visible="False"></asp:TextBox>
        &nbsp;
    </form>
</body>
</html>
