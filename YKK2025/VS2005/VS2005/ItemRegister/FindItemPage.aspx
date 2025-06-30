<%@ Page Language="VB" AutoEventWireup="false" CodeFile="FindItemPage.aspx.vb" Inherits="FindItemPage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Wave's Item搜尋(AS400)</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="DCode" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" Style="z-index: 126; left: 101px; position: absolute;
            top: 5px; text-align: left" Width="98px" MaxLength="7"></asp:TextBox>
        <asp:TextBox ID="TextBox1" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 9px; position: absolute;
            top: 6px; text-align: left" Width="57px">CODE</asp:TextBox>
        <asp:TextBox ID="TextBox2" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 13px; position: absolute;
            top: 32px; text-align: left" Width="88px">NAME-1</asp:TextBox>
        <asp:TextBox ID="TextBox3" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 13px; position: absolute;
            top: 55px; text-align: left" Width="88px">AND NAME-2</asp:TextBox>
        <asp:TextBox ID="TextBox5" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 13px; position: absolute;
            top: 80px; text-align: left" Width="88px">AND NAME-3</asp:TextBox>
        <asp:TextBox ID="TextBox4" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 12px; position: absolute;
            top: 105px; text-align: left" Width="88px">AND NAME-4</asp:TextBox>
        <asp:TextBox ID="DName1" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" Style="z-index: 126; left: 101px; position: absolute; top: 29px;
            text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DName2" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" Style="z-index: 126; left: 101px; position: absolute; top: 54px;
            text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DName3" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" MaxLength="35" Style="z-index: 126; left: 101px; position: absolute;
            top: 79px; text-align: left" Width="289px"></asp:TextBox>
        <asp:Button ID="BFindItem" runat="server" Height="25px" Style="z-index: 104; left: 399px;
            position: absolute; top: 103px" Text="搜尋" Width="45px" />

        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            PageSize="25" Style="z-index: 114; left: 12px; position: absolute; top: 131px"
            Width="600px">
            <RowStyle BackColor="White" ForeColor="Blue" />
            <Columns>
                <asp:CommandField ShowSelectButton="True" />
                <asp:BoundField DataField="Code" HeaderText="Item Code" />
                <asp:BoundField DataField="Name1" HeaderText="Item Name-1" />
                <asp:BoundField DataField="Name2" HeaderText="Item Name-2" />
                <asp:BoundField DataField="Name3" HeaderText="Item Name-3" />
            </Columns>
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#FF8000" Font-Bold="True" ForeColor="White" BorderStyle="Groove" />
            <FooterStyle BorderStyle="Groove" />
        </asp:GridView>
        <asp:TextBox ID="DMessage1" runat="server" BackColor="White" BorderColor="White"
            BorderStyle="None" Font-Names="Times New Roman" Font-Size="11pt" Style="z-index: 100;
            left: 207px; position: absolute; top: 7px" Width="200px"> 只顯示前50筆資料</asp:TextBox>
        <asp:TextBox ID="DName4" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="35" Style="z-index: 126; left: 101px;
            position: absolute; top: 104px; text-align: left" Width="289px"></asp:TextBox>
    </div>
    </form>
</body>
</html>
