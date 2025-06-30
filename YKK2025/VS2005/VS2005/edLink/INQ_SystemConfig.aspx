<%@ Page Language="VB" AutoEventWireup="false" CodeFile="INQ_SystemConfig.aspx.vb" Inherits="INQ_SystemConfig" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>系統參數</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  篩選條件                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        類別：<asp:TextBox ID="DCat" runat="server" BackColor="#FFFF80"></asp:TextBox>
        <asp:Button ID="BCat" runat="server" Text="...." CausesValidation="False" Width="16px" />&nbsp;
        <br />
<!--  -->
        項目：<asp:TextBox ID="DDKey" runat="server" BackColor="#FFFF80"></asp:TextBox>
<!--  -->
        <asp:Button ID="BSearch" runat="server" Text="搜尋" Width="48px" />
<!--  -->
        <asp:HiddenField ID="HiddenField1" runat="server" />
        <br />
        <br />
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  資料                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="DGridView1" runat="server" AutoGenerateColumns="false" Style="z-index: 103; left: 16px; position: absolute; top: 72px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Cat" HeaderText="類別" />
                <asp:BoundField DataField="CatDescr" HeaderText="類別名" />
                <asp:BoundField DataField="DKey" HeaderText="項目" />
                <asp:BoundField DataField="Data" HeaderText="設定/名稱" />
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>
    </div>
    </form>
</body>
</html>
