<%@ Page Language="VB" AutoEventWireup="false" CodeFile="INQ_REQStatus01.aspx.vb" Inherits="INQ_REQStatus01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >


<head  runat="server">
    <title>Adjust Status-01</title>
</head>
<body>
    <form id="form2" runat="server">
    <div>
        <asp:imagebutton id="BExcel" style="Z-INDEX: 100; LEFT: 10px; POSITION: absolute; TOP: 10px" runat="server"
				ImageUrl="~/Images/Excel.jpg" Height="21px" Width="21px"></asp:imagebutton>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"  Width="800px" Style="z-index: 103; left: 24px; position: absolute; top: 40px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px">
            <Columns>

                <asp:BoundField DataField="ORDN5C" HeaderText="Order NO" />
                <asp:BoundField DataField="CSTC5C" HeaderText="Customer"  />
                <asp:BoundField DataField="BYRC5C" HeaderText="Buyer"  />
                <asp:BoundField DataField="RDLU5M" HeaderText="新規登錄"  />
                <asp:BoundField DataField="FREMARK" HeaderText=""  />
                <asp:BoundField DataField="RDLU5C" HeaderText="現在"  />
                <asp:BoundField DataField="LREMARK" HeaderText="" />

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
