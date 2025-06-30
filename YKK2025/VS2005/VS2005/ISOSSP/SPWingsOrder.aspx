<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SPWingsOrder.aspx.vb" Inherits="SPWingsOrder" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Wings Order Inf.</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++                                                                     ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<asp:TextBox style="Z-INDEX: 126; LEFT: 24px; POSITION: absolute; TOP: 10px; TEXT-ALIGN: left" id="DTitle" runat="server" Width="232px" Height="24px" ReadOnly="True" ForeColor="White" BorderStyle="None" BackColor="Black" Font-Bold="true" Font-Size="Larger">WINGS Order Inf</asp:TextBox>
<asp:TextBox style="Z-INDEX: 126; LEFT: 24px; POSITION: absolute; TOP: 290px; TEXT-ALIGN: left" id="DDetail" runat="server" Width="232px" Height="24px" ReadOnly="True" ForeColor="White" BorderStyle="None" BackColor="Black" Font-Bold="true" Font-Size="Larger">WINGS Order Inf</asp:TextBox>
    
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"   Style="z-index: 103; left: 24px; position: absolute; top: 40px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px">
            <Columns>
                    <asp:BoundField DataField="ORDN5E" HeaderText="OrderNo" />
                    <asp:BoundField DataField="OSBN5E" HeaderText="SubNo" />
                    <asp:BoundField DataField="CORN5C" HeaderText="C.OrderNo" />
                    <asp:BoundField DataField="OCNU5C" HeaderText="Cfm.D" />

                    <asp:BoundField DataField="RDLU5C" HeaderText="Req.D" />

                    <asp:BoundField DataField="CSTC5C" HeaderText="Customer" />
                    <asp:BoundField DataField="FL1I39" HeaderText="" />
                    <asp:BoundField DataField="BYRC5C" HeaderText="Buyer" />
                    <asp:BoundField DataField="BYRI35" HeaderText="" />
                    <asp:BoundField DataField="SPRC5C" HeaderText="Sales" />
                    <asp:BoundField DataField="FEMI05" HeaderText="" />

                    <asp:BoundField DataField="SMPF5C" HeaderText="Sample" />
                    <asp:BoundField DataField="NCMF5C" HeaderText="No.C" />

                    <asp:BoundField DataField="ITMC5E" HeaderText="Item" />
                    <asp:BoundField DataField="ITEMNAME" HeaderText="" />
                    <asp:BoundField DataField="CLRC5E" HeaderText="Color" />
                    <asp:BoundField DataField="LNGV5E" HeaderText="Length" />
                    <asp:BoundField DataField="LUNC5E" HeaderText="Length.U" />
                    <asp:BoundField DataField="ORRQ5E" HeaderText="Qty" DataFormatString="{0:#,0}" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>      
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False"   Style="z-index: 103; left: 24px; position: absolute; top: 320px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px">
            <Columns>
                    <asp:BoundField DataField="ORDN5E" HeaderText="OrderNo" />
                    <asp:BoundField DataField="OSBN5E" HeaderText="SubNo" />
                    <asp:BoundField DataField="CORN5C" HeaderText="C.OrderNo" />
                    <asp:BoundField DataField="OCNU5C" HeaderText="Cfm.D" />
                    <asp:BoundField DataField="RDLU5C" HeaderText="Req.D" />

                    <asp:BoundField DataField="LUNC5E" HeaderText="Length.U" DataFormatString="{0:#,0}" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField> 
                    <asp:BoundField DataField="CLRC5E" HeaderText="Color" DataFormatString="{0:#,0}" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>                     
                    <asp:BoundField DataField="ORRQ5E" HeaderText="Qty" DataFormatString="{0:#,0}" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>      

                    <asp:BoundField DataField="BYRC5C" HeaderText="Buyer" />
                    <asp:BoundField DataField="BYRI35" HeaderText="" />
                    <asp:BoundField DataField="SPRC5C" HeaderText="Sales" />
                    <asp:BoundField DataField="FEMI05" HeaderText="" />

                    <asp:BoundField DataField="SMPF5C" HeaderText="Sample" />
                    <asp:BoundField DataField="NCMF5C" HeaderText="No.C" />

                    <asp:BoundField DataField="ITMC5E" HeaderText="Item" />
                    <asp:BoundField DataField="ITEMNAME" HeaderText="" />
                    
                    <asp:BoundField DataField="LNGV5E" HeaderText="Length"/>
                    

                    <asp:BoundField DataField="CSTC5C" HeaderText="Customer" />
                    <asp:BoundField DataField="FL1I39" HeaderText="" />
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
