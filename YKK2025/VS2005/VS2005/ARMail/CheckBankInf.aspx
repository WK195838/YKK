<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CheckBankInf.aspx.vb" Inherits="CheckBankInf" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>檢測銀行資訊</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" BackColor="White"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Height="300px" Style="z-index: 126; left: 8px; position: absolute; top: 8px"
            Width="900px">
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <SelectedItemStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="#663399" />
            <ItemStyle BackColor="White" ForeColor="#330099" />
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
            <Columns>
                <asp:BoundColumn DataField="Cust" HeaderText="客戶無銀行資訊異常名單" ReadOnly="True">
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundColumn>
            </Columns>
        </asp:DataGrid>
    
    </div>
    </form>
</body>
</html>
