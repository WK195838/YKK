<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfF_FCTACTAnalysis_02.aspx.vb" Inherits="InfF_FCTACTAnalysis_02" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>ACT-FCT FCT or ACT INF.</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Order GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="OrderGridView" runat="server" AutoGenerateColumns="False" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Customer" HeaderText="成衣廠"  />
                <asp:BoundField DataField="Season" HeaderText="季"  />
                <asp:BoundField DataField="Month" HeaderText="MONTH"  />
                <asp:BoundField DataField="CustItem" HeaderText="客戶ITEM"  />

                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="Item" HeaderText="ITEM" Target="_blank">
                </asp:HyperLinkField>

                <asp:BoundField DataField="ItemName" HeaderText="ITEMN NAME"  />
                <asp:BoundField DataField="Color" HeaderText="COLOR"  />
                <asp:BoundField DataField="ActQty" HeaderText="QTY" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  FCT GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="FCTGridView" runat="server" AutoGenerateColumns="False" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Customer" HeaderText="成衣廠"  />
                <asp:BoundField DataField="Season" HeaderText="季"  />
                <asp:BoundField DataField="Month" HeaderText="MONTH"  />
                <asp:BoundField DataField="CustItem" HeaderText="ITEM"  />
                <asp:BoundField DataField="Color" HeaderText="COLOR"  />
                <asp:BoundField DataField="KeyData1" HeaderText="Y_ITEM"  />
                <asp:BoundField DataField="KeyData2" HeaderText="Y_COLOR"  />
                <asp:BoundField DataField="ItemName" HeaderText="ITEMN NAME"  />
                <asp:BoundField DataField="Style" HeaderText="STYLE"  />
                <asp:BoundField DataField="Article" HeaderText="INF-1"  />
                <asp:BoundField DataField="Part" HeaderText="INF-2"  />
                <asp:BoundField DataField="FCTQty" HeaderText="QTY" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
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