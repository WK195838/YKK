<%@ Page Language="VB" AutoEventWireup="false" CodeFile="INQ_DTMWPriorityList.aspx.vb" Inherits="INQ_DTMWPriorityList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>DTMW STATUS</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" width="1200" Style="z-index: 103; left:0px; position: absolute; top: 2px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Name"  HeaderText="">  <ItemStyle HorizontalAlign="LEFT" />   </asp:BoundField>  
                <asp:BoundField DataField="pNo"  HeaderText="">  <ItemStyle HorizontalAlign="LEFT" />   </asp:BoundField>
                <asp:BoundField DataField="Priority" HeaderText="">   <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField> 
                <asp:BoundField DataField="YKKColorCode" HeaderText="">  <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField>     
                <asp:BoundField DataField="OrderSts" HeaderText="">  <ItemStyle HorizontalAlign="LEFT" />   </asp:BoundField>     

                <asp:BoundField DataField="Sts" HeaderText="">   <ItemStyle HorizontalAlign="center" />   </asp:BoundField>    

                <asp:HyperLinkField DataNavigateUrlFields="WFUrl" DataNavigateUrlFormatString="{0}"
                    DataTextField="WFUrlDesc" HeaderText="" Target="_blank">  <ItemStyle HorizontalAlign="CENTER" /> 
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="ViewUrl" DataNavigateUrlFormatString="{0}"
                    DataTextField="ViewUrlDesc" HeaderText="" Target="_blank">  <ItemStyle HorizontalAlign="CENTER" /> 
                </asp:HyperLinkField>

                <asp:BoundField DataField="date" HeaderText="">   <ItemStyle HorizontalAlign="LEFT" />   </asp:BoundField>    
                <asp:BoundField DataField="customerD" HeaderText="">   <ItemStyle HorizontalAlign="LEFT" />   </asp:BoundField>    
                <asp:BoundField DataField="buyerD" HeaderText="" >   <ItemStyle HorizontalAlign="LEFT" />   </asp:BoundField>
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
