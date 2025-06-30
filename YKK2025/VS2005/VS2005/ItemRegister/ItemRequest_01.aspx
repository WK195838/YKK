<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ItemRequest_01.aspx.vb" Inherits="ItemRequest_01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Item Request</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
 
        <asp:DropDownList ID="DWaitRequest" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 7px; position: absolute; top: 7px" Width="188px">
            <asp:ListItem Selected="True" Value="徐徐徐">徐徐徐</asp:ListItem>
        </asp:DropDownList><asp:DropDownList ID="DSts" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 200px; position: absolute; top: 7px" Width="73px">
            <asp:ListItem Selected="True" Value="ALL">ALL</asp:ListItem>
            <asp:ListItem Value="0">Ｏ</asp:ListItem>
            <asp:ListItem Value="1">Ｘ</asp:ListItem>
            <asp:ListItem Value="2">Ｗ</asp:ListItem>
        </asp:DropDownList>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" PageSize="25"
            Style="z-index: 114; left: 9px; position: absolute; top: 42px">
            <RowStyle BackColor="White" ForeColor="Black" />
            <Columns>
                <asp:HyperLinkField DataNavigateUrlFields="UpdateURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="Field1" Target="_blank">
                </asp:HyperLinkField>
                <asp:BoundField DataField="StsD" HeaderText="Status" />
                <asp:BoundField DataField="ItemName" HeaderText="Item Name" />
                <asp:BoundField DataField="RCode" HeaderText="Ref.Code" />
                <asp:BoundField DataField="RItemName" HeaderText="Ref.Item Name" />
            </Columns>
            <FooterStyle BorderStyle="Groove" />
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#FF8000" BorderStyle="Groove" Font-Bold="True" ForeColor="White" />
        </asp:GridView><asp:Button ID="BExpandRefItem" runat="server" Height="25px" Style="z-index: 104; left: 676px;
            position: absolute; top: 4px" Text="展開" Width="72px" Visible="False" /><asp:Button ID="BRefresh" runat="server" Height="25px" Style="z-index: 104; left: 582px;
            position: absolute; top: 4px" Text="查詢 / 更新" Width="91px" /><asp:Button ID="BUpload" runat="server" Height="25px" Style="z-index: 104; left: 751px;
            position: absolute; top: 4px" Text="上傳" Width="72px" />
        <asp:TextBox ID="DItemName" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 111; left: 277px; position: absolute;
            top: 5px" Width="296px"></asp:TextBox>
    
    </div>
    </form>
</body>
</html>
