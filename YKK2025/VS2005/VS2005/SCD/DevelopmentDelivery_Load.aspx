<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DevelopmentDelivery_Load.aspx.vb" Inherits="DevelopmentDelivery_Load" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>負荷工程明細</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>

        <asp:TextBox ID="TextBox10" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 15px; position: absolute;
            top: 7px; text-align: left" Width="55px">預定起迄</asp:TextBox>
        <asp:TextBox ID="DBStartTime" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="10" Style="z-index: 126; left: 72px;
            position: absolute; top: 7px; text-align: left" Width="82px"></asp:TextBox>
        <asp:TextBox ID="DBEndTime" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="10" Style="z-index: 126; left: 188px;
            position: absolute; top: 7px; text-align: left" Width="82px"></asp:TextBox>
        <asp:TextBox ID="TextBox13" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 165px; position: absolute;
            top: 7px; text-align: left" Width="16px">～</asp:TextBox>
        <asp:Button ID="BOK" runat="server" Height="25px" Style="z-index: 104; left: 297px;
            position: absolute; top: 7px" Text="GO" Width="62px" />

        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"
            BorderColor="#CC9966" BorderStyle="Groove" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            PageSize="20" Style="z-index: 114; left: 10px; position: absolute; top: 40px"
            Width="900px">
            <RowStyle BackColor="White" ForeColor="Blue" />
            <Columns>
                <asp:BoundField DataField="STSD" />
                <asp:BoundField DataField="OPPER" HeaderText="擔當" />
                <asp:BoundField DataField="NO" HeaderText="No." />
                <asp:BoundField DataField="DEVNO" HeaderText="開發No." />
                <asp:BoundField DataField="CODE" HeaderText="Code" />
                <asp:BoundField DataField="BUYER" HeaderText="Buyer" />
                <asp:BoundField DataField="BTIME" HeaderText="預定" />
                <asp:BoundField DataField="ATIME" HeaderText="實績" />
            </Columns>
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#FF8000" BorderStyle="Groove" Font-Bold="True" ForeColor="White" />
            <FooterStyle BorderStyle="Groove" />
        </asp:GridView>
    </div>
    </form>
</body>
</html>
