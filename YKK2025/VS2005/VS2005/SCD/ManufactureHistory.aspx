<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ManufactureHistory.aspx.vb" Inherits="ManufactureHistory" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>製造委託履歷</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" ShowFooter="True"
            Style="z-index: 101; left: 10px; position: absolute; top: 40px" Width="790px">
            <Columns>
                <asp:BoundField DataField="OP" HeaderText="工程" />
                <asp:BoundField DataField="OPPER" HeaderText="擔當" />
                <asp:BoundField DataField="OPBTIME" HeaderText="預定時間" />
                <asp:BoundField DataField="OPBHOURS" HeaderText="預定分數" />
                <asp:BoundField DataField="OPATIME" HeaderText="實際時間" />
                <asp:BoundField DataField="OPAHOURS" HeaderText="實際分數" />
                <asp:BoundField DataField="OPDELAYC1" HeaderText="遲納類別" />
                <asp:BoundField DataField="OPDELAYC2" HeaderText="原因類別" />
                <asp:BoundField DataField="OPREM" HeaderText="作業說明" />
            </Columns>
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
    
    </div>
    </form>
</body>
</html>
