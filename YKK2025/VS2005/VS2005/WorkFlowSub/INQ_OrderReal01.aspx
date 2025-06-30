<%@ Page Language="VB" AutoEventWireup="false" CodeFile="INQ_OrderReal01.aspx.vb" Inherits="INQ_OrderReal01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Order Real-00</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>


<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"  Width="800px" Style="z-index: 103; left: 8px; position: absolute; top: 8px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px">
            <Columns>

                
                <asp:BoundField DataField="SUMCODE" HeaderText="Summary Code"  />
                <asp:BoundField DataField="YYMM" HeaderText="Month" />
                <asp:BoundField DataField="DAY" HeaderText="Day" />
                <asp:BoundField DataField="KEYNAME" HeaderText="Buyer/Sales"  />
                <asp:BoundField DataField="AMOUNT" HeaderText="Amount" DataFormatString="{0:N}" ItemStyle-HorizontalAlign="Right"/>

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
