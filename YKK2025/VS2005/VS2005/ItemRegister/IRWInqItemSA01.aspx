<%@ Page Language="VB" AutoEventWireup="false" CodeFile="IRWInqItemSA01.aspx.vb" Inherits="IRWInqItemSA01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>IRW-WINGS SA </title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <asp:TextBox style="Z-INDEX: 103; LEFT: 544px; POSITION: absolute; TOP: 8px" id="DKUser" runat="server" Font-Size="10pt" Height="24px" Width="128px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove"></asp:TextBox>
    <asp:button style="Z-INDEX: 103; LEFT: 760px; POSITION: absolute; TOP: 40px" id="BInq" runat="server" Font-Size="Larger" Height="56px" Width="112px" Text="Go" ForeColor="Black" BackColor="Silver"></asp:button>
    
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" width="1100" Style="z-index: 103; left:0px; position: absolute; top: 150px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Form"  HeaderText="">  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="No" HeaderText="">   <ItemStyle HorizontalAlign="left" />   </asp:BoundField>    
                <asp:BoundField DataField="Date" HeaderText="" >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="Code" HeaderText="" >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="ItemName" HeaderText="" >   <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="Status" HeaderText="" >    <ItemStyle HorizontalAlign="left" />   </asp:BoundField> 
                <asp:BoundField DataField="NoDisplay" HeaderText="" >    <ItemStyle HorizontalAlign="left" />   </asp:BoundField> 
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>

        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" width="400" Style="z-index: 103; left:0px; position: absolute; top: 2px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Name"  HeaderText="">  <ItemStyle HorizontalAlign="LEFT" />   </asp:BoundField>  
                <asp:BoundField DataField="FORM"  HeaderText="">  <ItemStyle HorizontalAlign="LEFT" />   </asp:BoundField>  
                <asp:BoundField DataField="AppCount" HeaderText="">  <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField>     
                <asp:BoundField DataField="ComCount" HeaderText="">   <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField>    
                <asp:BoundField DataField="SACount" HeaderText="">   <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField>    
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>   
        <asp:TextBox ID="DKNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
            Font-Size="10pt" ForeColor="Blue" Height="24px" Style="z-index: 103; left: 680px;
            position: absolute; top: 8px" Width="192px"></asp:TextBox>
        &nbsp;
    </div>
    </form>
</body>
</html>
