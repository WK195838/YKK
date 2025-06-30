<%@ Page Language="VB" AutoEventWireup="false" CodeFile="INQ_OPEnt2Etd_01.aspx.vb" Inherits="INQ_OPEnt2Etd_01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>OrderProgress Ent2Etd_01</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:imagebutton id="BExcel" style="Z-INDEX: 100; LEFT: 10px; POSITION: absolute; TOP: 10px" runat="server"
				ImageUrl="~/Images/Excel.jpg" Height="21px" Width="21px"></asp:imagebutton>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"  Width="1100px" Style="z-index: 103; left: 24px; position: absolute; top: 40px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px">
            <Columns>

                <asp:BoundField DataField="OrderNo" HeaderText="OrderNO" />
                <asp:BoundField DataField="OrderSubNo" HeaderText="SubNo"  />
                <asp:BoundField DataField="Customer" HeaderText="Customer"  />
                <asp:BoundField DataField="CustomerName" HeaderText=""  />
                <asp:BoundField DataField="Salesman" HeaderText="Sales"  />
                <asp:BoundField DataField="Buyer" HeaderText="Buyer"  />
                <asp:BoundField DataField="Progress" HeaderText="Progress" />
                <asp:BoundField DataField="Item" HeaderText="Item" />
                <asp:BoundField DataField="Color" HeaderText="Color" />
                <asp:BoundField DataField="QTY" HeaderText="Qty" />
                <asp:BoundField DataField="EntryDate" HeaderText="ConfirmDate" />
                <asp:BoundField DataField="RequestDate" HeaderText="RequestDate" />
                <asp:BoundField DataField="ETDDate" HeaderText="ETD" />
                <asp:BoundField DataField="Entry2ETDDays" HeaderText="Days" />
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
