<%@ Page Language="VB" AutoEventWireup="false" CodeFile="OR2PR.aspx.vb" Inherits="OR2PR" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Wave's OR2PR</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            PageSize="20" Style="z-index: 114; left: -4px; position: absolute; top: 37px"
            Width="1100px">
            <RowStyle BackColor="White" ForeColor="Blue" />
            <Columns>
                <asp:BoundField DataField="XORNO" HeaderText="OR-NO." />
                <asp:BoundField DataField="PSCN9F" HeaderText="PP-NO." />
                <asp:BoundField DataField="PSHN9F" HeaderText="PR-NO." />
                <asp:BoundField DataField="LNGV9F" HeaderText="Length" />
                <asp:BoundField DataField="LUNC9F" HeaderText="U" />
                <asp:BoundField DataField="CLRC9F" HeaderText="Color" />
                <asp:BoundField DataField="ALMQ9F" HeaderText="Quantity" />
                <asp:BoundField DataField="PCPU9F" HeaderText="完成日期" />
                <asp:BoundField DataField="XF15" HeaderText="射出" />
                <asp:BoundField DataField="XF15DATA" HeaderText="射出完成日" />
                <asp:BoundField DataField="EMPCAB3" HeaderText="射出EMPID" />
                <asp:BoundField DataField="XF44" HeaderText="定吋" >
                    <HeaderStyle Width="50px" />
                </asp:BoundField>
                <asp:BoundField DataField="XF44DATE" HeaderText="定吋完成日" />
                <asp:BoundField DataField="EMPCAB1" HeaderText="定吋EMPID" >
                    <HeaderStyle Width="50px" />
                    <ItemStyle Width="50px" />
                </asp:BoundField>
                <asp:BoundField DataField="XF64" HeaderText="仕上" >
                    <HeaderStyle Width="50px" />
                </asp:BoundField>
                <asp:BoundField DataField="XF64DATE" HeaderText="仕上完成日" />
                <asp:BoundField DataField="EMPCAB2" HeaderText="仕上EMPID" />
            </Columns>
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#FF8000" BorderStyle="Groove" Font-Bold="True" ForeColor="White" />
            <FooterStyle BorderStyle="Groove" />
        </asp:GridView>
        <asp:TextBox ID="TextBox1" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 20px; position: absolute;
            top: 11px; text-align: left" Width="57px">訂單號碼</asp:TextBox>
        <asp:TextBox ID="DORNO" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" MaxLength="10" Style="z-index: 126; left: 81px; position: absolute;
            top: 7px; text-align: left" Width="135px"></asp:TextBox>
        <asp:Button ID="BORNO" runat="server" Height="25px" Style="z-index: 104; left: 223px;
            position: absolute; top: 6px" Text="GO" Width="45px" />
        &nbsp;&nbsp;&nbsp;<br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
    
    </div>
    </form>
</body>
</html>
