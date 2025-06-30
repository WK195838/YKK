<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfF_FCTACTAnalysis_03a.aspx.vb" Inherits="InfF_FCTACTAnalysis_03a" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>ACT-FCT ACT SLIDER INF.</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Order GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="OrderGridView" runat="server" AutoGenerateColumns="False" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="OrderNo" HeaderText="ORD.NO"  />
                <asp:BoundField DataField="OrderSubNo" HeaderText="SUB"  />
                <asp:BoundField DataField="PO" HeaderText="PO"  />
                <asp:BoundField DataField="Seqno" HeaderText="SEQ"  />
                <asp:BoundField DataField="CustWavesCode" HeaderText="CODE"  />
                <asp:BoundField DataField="Season" HeaderText="SEASON"  />
                <asp:BoundField DataField="V_Month" HeaderText="BUYMONTH"  />
                <asp:BoundField DataField="Item" HeaderText="ITEM"  />
                <asp:BoundField DataField="ItemName" HeaderText="ITEMNAME"  />
                <asp:BoundField DataField="Length" HeaderText="LENGTH"  />
                <asp:BoundField DataField="Unit" HeaderText="UNIT"  />
                <asp:BoundField DataField="Color" HeaderText="COLOR"  />
                <asp:BoundField DataField="OrderDate" HeaderText="ORD.DATE"  />
                <asp:BoundField DataField="ReqDate" HeaderText="REQ.DATE"  />
                <asp:BoundField DataField="CompletedDate" HeaderText="COM.DATE"  />
                <asp:BoundField DataField="OrderQty" HeaderText="QTY" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>

                <asp:BoundField DataField="sItem" HeaderText="ITEM"  />
                <asp:BoundField DataField="sItemName" HeaderText="ITEMNAME"  />
                <asp:BoundField DataField="sColor" HeaderText="COLOR"  />

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