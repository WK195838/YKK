<%@ Page Language="VB" AutoEventWireup="false" CodeFile="EDXList01.aspx.vb" Inherits="EDXList01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>EDX Inf</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox style="Z-INDEX: 103; LEFT: 24px; POSITION: absolute; TOP: 16px" id="LSpec" runat="server" Width="96px" Height="24px" ForeColor="Black" BorderStyle="None" Font-Size="Medium" Font-Bold="False">型別組</asp:TextBox>
        <asp:TextBox style="Z-INDEX: 110; LEFT: 80px; POSITION: absolute; TOP: 16px" id="DSPEC" runat="server" Width="272px" Height="24px" BorderStyle="Groove" BackColor="#FFFF80" ForeColor="Black"></asp:TextBox>
    
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left:24px; position: absolute; top: 56px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>

                <asp:BoundField DataField="Seqno"  HeaderText="No."> </asp:BoundField>   
                <asp:BoundField DataField="Cat" HeaderText="Catogory"> </asp:BoundField>    
                
                <asp:BoundField DataField="Size"  HeaderText="Size"> </asp:BoundField>            
                <asp:BoundField DataField="Family"  HeaderText="Family"></asp:BoundField>            
                <asp:BoundField DataField="Body"  HeaderText="Body"> </asp:BoundField>            
                <asp:BoundField DataField="Puller"  HeaderText="Puller"> </asp:BoundField>            
                <asp:BoundField DataField="Color"  HeaderText="Color"> </asp:BoundField>            
                <asp:BoundField DataField="Finish"  HeaderText="Finish"> </asp:BoundField>            

                <asp:BoundField DataField="Supplier"  HeaderText="Supplier"> </asp:BoundField>            

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>    
        <asp:Button ID="Go" runat="server" BackColor="Aqua" ForeColor="Blue" Height="24px"
            Style="z-index: 121; left: 368px; position: absolute; top: 16px" Text="Go" Width="40px" />
    </div>
    </form>
</body>
</html>