<%@ Page Language="VB" AutoEventWireup="false" CodeFile="IRWInqMoldCancel.aspx.vb" Inherits="IRWInqMoldCancel" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>模具報廢一覽</title>
</head>
<body>

    <form id="form1" runat="server">
    <div>
    
    <asp:Label style="Z-INDEX: 103; LEFT: 8px; POSITION: absolute; TOP: 16px" id="LSeaxh" runat="server" Font-Size="12pt" Width="300px">SEARCH KEY</asp:Label>
    <asp:TextBox style="Z-INDEX: 103; LEFT: 8px; POSITION: absolute; TOP: 40px" id="DSearchKey" runat="server" Font-Size="12pt" Height="24px" Width="296px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove"></asp:TextBox>
    <asp:button style="Z-INDEX: 103; LEFT: 320px; POSITION: absolute; TOP: 16px" id="BInq" runat="server" Font-Size="Larger" Height="48px" Width="112px" Text="Go" ForeColor="Black" BackColor="Silver"></asp:button>

    <asp:Label style="Z-INDEX: 103; LEFT: 440px; POSITION: absolute; TOP: 16px" id="Label1" runat="server" Font-Size="24pt" Width="200px" Height="32px" BackColor="Black" ForeColor="White">模具報廢一覽</asp:Label>
    
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" width="1100" Style="z-index: 103; left:0px; position: absolute; top: 80px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="canceldate"  HeaderText="">  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="item" HeaderText="">   <ItemStyle HorizontalAlign="left" />   </asp:BoundField>    
                <asp:BoundField DataField="itemname" HeaderText="" >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="size" HeaderText="" >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="chain" HeaderText="" >   <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="slider" HeaderText="" >    <ItemStyle HorizontalAlign="left" />   </asp:BoundField> 
                <asp:BoundField DataField="tape_" HeaderText="" >    <ItemStyle HorizontalAlign="left" />   </asp:BoundField> 
                <asp:BoundField DataField="finish" HeaderText="" >    <ItemStyle HorizontalAlign="left" />   </asp:BoundField> 
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
