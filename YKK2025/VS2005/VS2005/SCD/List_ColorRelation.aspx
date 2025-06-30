<%@ Page Language="VB" AutoEventWireup="false" CodeFile="List_ColorRelation.aspx.vb" Inherits="List_ColorRelation" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="TextBox4" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 20px; position: absolute;
            top: 11px; text-align: left" Width="40px">供應商</asp:TextBox>
        <asp:DropDownList ID="DSup" runat="server" BackColor="Yellow" ForeColor="Blue" Height="24px"
            Style="z-index: 120; left: 67px; position: absolute; top: 8px" Width="170px">
        </asp:DropDownList>
        <asp:TextBox ID="TextBox1" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 244px; position: absolute;
            top: 11px; text-align: left" Width="40px">色號</asp:TextBox>
        <asp:TextBox ID="DColor" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" MaxLength="10" Style="z-index: 126; left: 288px; position: absolute;
            top: 8px; text-align: left" Width="166px"></asp:TextBox>
        
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            PageSize="50" Style="z-index: 114; left: 31px; position: absolute; top: 44px"
            Width="600px">
            <RowStyle BackColor="White" ForeColor="Blue" />
            <Columns>
                <asp:BoundField DataField="SUP" HeaderText="供應商" />
                <asp:BoundField DataField="COLNO" HeaderText="色號" />
                <asp:BoundField DataField="YCOLNO" HeaderText="YKK" />
            </Columns>
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#FF8000" BorderStyle="Groove" Font-Bold="True" ForeColor="White" />
            <FooterStyle BorderStyle="Groove" />
        </asp:GridView>
        
        
        <asp:Button ID="BOK" runat="server" Height="25px" Style="z-index: 104; left: 470px;
            position: absolute; top: 6px" Text="GO" Width="62px" />
        <asp:ImageButton ID="BExcel" runat="server" Height="21px" ImageUrl="Images\msexcel.gif"
            Style="z-index: 122; left: 540px; position: absolute; top: 6px" Width="21px" />
    
    </div>
    </form>
</body>
</html>
