<%@ Page Language="VB" AutoEventWireup="false" CodeFile="INQ_OrderReal.aspx.vb" Inherits="INQ_OrderReal" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >

<head runat="server">
    <title>Order Real-00</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
              
              

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView 1                                                                 ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Height="100px" Width="800px"
        Style="z-index: 103; left: 8px; position: absolute; top: 16px" CellPadding="3" Font-Size="16pt" 
        GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                
                <asp:BoundField DataField="YYMM" HeaderText="Month (WFS890R1)" ItemStyle-HorizontalAlign="Center" />

                <asp:BoundField DataField="AMOUNT" HeaderText="Amount" DataFormatString="{0:N}" ItemStyle-HorizontalAlign="Right"/> 

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>



<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView 2                                                              ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->    

        <asp:DropDownList ID="DCat" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 16px; position: absolute; top: 136px" Width="96px" Height="24px">
            <asp:ListItem Selected="True" Value="MONTH">月</asp:ListItem>
            <asp:ListItem Value="DAY">日</asp:ListItem>
        </asp:DropDownList>
                

        <asp:DropDownList ID="DSum" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 112px; position: absolute; top: 136px" Width="104px" Height="24px">
            <asp:ListItem Selected="True" Value="00">ALL</asp:ListItem>
            <asp:ListItem Value="20">EX-YKK</asp:ListItem>
            <asp:ListItem Value="21">YKK</asp:ListItem>
        </asp:DropDownList>
                
<asp:button id="BInq" runat="server" Height="32px" Width="64px" Style="z-index: 103; left: 224px; position: absolute; top: 128px" ForeColor="Black" BackColor="Silver" Text="Go" Font-Size="Larger"></asp:button>



<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView 2                                                              ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->    
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" Height="200px" Width="800px"
        Style="z-index: 103; left: 8px; position: absolute; top: 160px" CellPadding="3" Font-Size="16pt" 
        GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                
                <asp:BoundField DataField="SUMCODE" HeaderText="Summary Code" ItemStyle-HorizontalAlign="Center"/> 
                <asp:BoundField DataField="YYMM" HeaderText="Month" ItemStyle-HorizontalAlign="Center" />
                <asp:BoundField DataField="DAY" HeaderText="Day" ItemStyle-HorizontalAlign="Center" />
                <asp:BoundField DataField="AMOUNT" HeaderText="Amount" DataFormatString="{0:N}" ItemStyle-HorizontalAlign="Right"/> 
                <asp:BoundField DataField="Buyer" HeaderText="" ItemStyle-HorizontalAlign="Center"/> 
                <asp:BoundField DataField="Sales" HeaderText="" ItemStyle-HorizontalAlign="Center"/> 

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
