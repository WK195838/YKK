<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ItemCheck.aspx.vb" Inherits="ItemCheck" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>ITEM CHECK</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        &nbsp;


        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 2px; position: absolute; top: 2px" Width="1200px">
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" Wrap="True" />
            <Columns>
                <asp:BoundField DataField="AN1" HeaderText="ITEM" ReadOnly="True"></asp:BoundField>
                <asp:BoundField DataField="STS" HeaderText="STATUS" ReadOnly="True"></asp:BoundField>
                <asp:BoundField DataField="A1" HeaderText="A1" ReadOnly="True"></asp:BoundField>
                <asp:BoundField DataField="B1" HeaderText="B1" ReadOnly="True"></asp:BoundField>
                <asp:BoundField DataField="C1" HeaderText="C1" ReadOnly="True"></asp:BoundField>
                <asp:BoundField DataField="D1" HeaderText="D1" ReadOnly="True"></asp:BoundField>
                <asp:BoundField DataField="E1" HeaderText="E1" ReadOnly="True"></asp:BoundField>
                <asp:BoundField DataField="F1" HeaderText="F1" ReadOnly="True"></asp:BoundField>
                <asp:BoundField DataField="G1" HeaderText="G1" ReadOnly="True"></asp:BoundField>
                <asp:BoundField DataField="H1" HeaderText="H1" ReadOnly="True"></asp:BoundField>
                <asp:BoundField DataField="I1" HeaderText="I1" ReadOnly="True"></asp:BoundField>
                <asp:BoundField DataField="J1" HeaderText="J1" ReadOnly="True"></asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>

    </div>
    </form>
</body>
</html>
