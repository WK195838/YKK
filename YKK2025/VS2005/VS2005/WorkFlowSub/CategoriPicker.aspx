<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CategoriPicker.aspx.vb" Inherits="CategoriPicker" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        &nbsp;<asp:TextBox ID="DCat" runat="server" Width="234px" BackColor="Yellow"></asp:TextBox>&nbsp;&nbsp;
        &nbsp;&nbsp; &nbsp;
        <asp:Button ID="BSerch" runat="server" Text="...." style="z-index: 100; left: 258px; position: absolute; width: 25px;top: 16px" />
        &nbsp;
        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="DCat"
            ErrorMessage="不允許空白"></asp:RequiredFieldValidator><br />
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 14px; position: absolute; top: 46px" Width="400px" AllowPaging="True" PageSize="20">
                <Columns>
                    <asp:CommandField ShowSelectButton="True" />
                    <asp:BoundField DataField="DKey" HeaderText="類別" />
                    <asp:BoundField DataField="Data" HeaderText="名稱" />
                </Columns>
                <RowStyle BackColor="White" ForeColor="Blue" />
                <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
                <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
                <HeaderStyle BackColor="#FF8000" Font-Bold="True" ForeColor="White" />
            </asp:GridView>
        &nbsp; &nbsp;
        &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;
    
    </div>
    </form>
</body>
</html>
