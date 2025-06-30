<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfF_FCTAnalysis_01.aspx.vb" Inherits="InfF_FCTAnalysis_01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>FCT-FCT ITEM</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="ITEMGridView" runat="server" AutoGenerateColumns="False" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Customer" HeaderText="成衣廠"  />
                <asp:BoundField DataField="Season" HeaderText="季"  />

                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="CustItem" HeaderText="ITEM" Target="_blank">
                </asp:HyperLinkField>
                
                <asp:BoundField DataField="Qty6"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Ratio6" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Qty5"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Ratio5" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Qty4"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Ratio4" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Qty3"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Ratio3" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Qty2"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Ratio2" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Qty1"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Ratio1" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
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