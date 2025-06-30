<%@ Page Language="VB" AutoEventWireup="false" CodeFile="INQ_S5MReport.aspx.vb" Inherits="INQ_S5MReport" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>S5M00 DATA</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:imagebutton id="BExcel" style="Z-INDEX: 100; LEFT: 10px; POSITION: absolute; TOP: 10px" runat="server"
				ImageUrl="~/Images/Excel.jpg" Height="21px" Width="21px"></asp:imagebutton>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"  Width="800px" Style="z-index: 103; left: 24px; position: absolute; top: 40px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px">
            <Columns>
                    <asp:BoundField DataField="ORDN5M" HeaderText="OrderNo" />
                    <asp:BoundField DataField="UIDC5M" HeaderText="Modify User" />
                    <asp:BoundField DataField="HCTU5M" HeaderText="Modify Date" />
                    <asp:BoundField DataField="HCRT5M" HeaderText="Modify Time" />
                    <asp:BoundField DataField="HACC5M" HeaderText="Access Code" />

                    <asp:BoundField DataField="CSTC5M" HeaderText="Customer" />
                    <asp:BoundField DataField="BYRC5M" HeaderText="Buyer" />
                    <asp:BoundField DataField="RDLU5M" HeaderText="REQ.Date" />
                    <asp:BoundField DataField="SMPF5M" HeaderText="Bulk/Sample" />
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
