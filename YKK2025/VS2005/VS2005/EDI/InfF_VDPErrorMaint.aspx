<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfF_VDPErrorMaint.aspx.vb" Inherits="InfF_VDPErrorMaint" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<<head id="Head1" runat="server">
    <title>Maint. VDP Inf.</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="DKey1" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="100px" Style="z-index: 103; left: 2px; position: absolute; top: 8px"></asp:TextBox>
        <asp:TextBox ID="DKey2" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="100px" Style="z-index: 103; left: 111px; position: absolute; top: 8px"></asp:TextBox>

    	<asp:button id="BInq" runat="server" Height="24px" Width="40px" Style="z-index: 103; left: 230px; position: absolute; top: 6px" ForeColor="Black" BackColor="Silver" Text="Go"></asp:button>

        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  GridView                                                                   ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" Style="z-index: 103; left: 2px; position: absolute; top: 34px" AutoGenerateColumns="False" CellPadding="3" Font-Size="9pt"  GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                 <asp:CommandField ShowEditButton="True" UpdateText="儲存">
                            <ItemStyle Width="100px" HorizontalAlign="Center" />
                 </asp:CommandField>
                                  
                <asp:BoundField DataField="Completed" HeaderText="Status" ReadOnly="True"  />
                <asp:BoundField DataField="CustWavesCode" HeaderText="成衣廠" ReadOnly="True"   />
                <asp:BoundField DataField="OrderNo" HeaderText="Order" ReadOnly="True"   />

                <asp:BoundField DataField="Finished" HeaderText="修改狀態"   />
                <asp:BoundField DataField="ModifyUser" HeaderText="修改者ID"  />
                <asp:BoundField DataField="Month" HeaderText="Confirm年月"  />

                <asp:BoundField DataField="Season" HeaderText="Season Code"  />
                <asp:BoundField DataField="Year" HeaderText="Season Year"   />
                <asp:BoundField DataField="BuyMonth" HeaderText="Buy Month"   />
                <asp:BoundField DataField="PLM" HeaderText="PLM"  />
                <asp:BoundField DataField="Style" HeaderText="Style"   />
                <asp:BoundField DataField="TapeColor" HeaderText="Color" />

                <asp:BoundField DataField="Unique_ID" HeaderText="ID"   />

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>

    </div>
    </form>
</body></html>