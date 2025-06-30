<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfF_StockProdList.aspx.vb" Inherits="InfF_StockProdList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Stock & Prod Inf.</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  篩選欄位                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Label ID="Label3" runat="server" Style="z-index: 103; left: 10px; position: absolute; top: 8px">Buyer：</asp:Label>
        <asp:Label ID="Label1" runat="server" Style="z-index: 103; left: 155px; position: absolute; top: 8px">Keep Code：</asp:Label>
        <asp:Label ID="Label2" runat="server" Style="z-index: 103; left: 386px; position: absolute; top: 8px">Item：</asp:Label>

        <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="74px" Style="z-index: 103; left: 64px; position: absolute; top: 8px"></asp:TextBox>


        <asp:DropDownList id="DKeepCode" runat="server" Width="129px" ForeColor="Blue" BackColor="Yellow" Style="z-index: 103; left: 245px; position: absolute; top: 8px">
            <asp:ListItem>SS14</asp:ListItem>
        </asp:DropDownList>
        
        <asp:TextBox ID="DItem" runat="server" BackColor="Yellow" BorderStyle="Groove"
        ForeColor="Blue" Height="16px"  Width="74px" Style="z-index: 103; left: 432px; position: absolute; top: 8px"></asp:TextBox>
        &nbsp;<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ --><!-- ++  Action (CheckBox)                                                                   ++ --><!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ --><!-- 含 New-FCT -->

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  按鈕                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    	<asp:button id="BInq" runat="server" Height="24px" Width="40px" Style="z-index: 103; left: 527px; position: absolute; top: 10px" ForeColor="Black" BackColor="Silver" Text="Go"></asp:button>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 15px; position: absolute; top: 40px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Item" HeaderText="Item"  />
                <asp:BoundField DataField="ItemName" HeaderText="ItemName"  />
                <asp:BoundField DataField="Color" HeaderText="Color"  />
                <asp:BoundField DataField="KeepCode" HeaderText="Keep Code"  />
                <asp:BoundField DataField="KeepQty" HeaderText="Stock Qty" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="KeepQtyDescr"  HeaderText="LOC/QTY/LAST_USED/OR/C.ORDER"   /> 
                <asp:BoundField DataField="ProdQty"   HeaderText="Prod. Qty" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="ProdQtyDescr1"   HeaderText="SCHE-PROD/QTY, ON-PROD/QTY"  /> 
                <asp:BoundField DataField="ProdQtyDescr2"   HeaderText="OR/QTY/USING_AVAILABLE"  /> 
                <asp:BoundField DataField="DataTime"   HeaderText="資料時點"  /> 
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
