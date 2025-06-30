<%@ Page Language="VB" AutoEventWireup="false" CodeFile="IRWNoticePageRanking.aspx.vb" Inherits="IRWNoticePageRanking" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>無效工時 (Ranking) </title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView9" runat="server" WIDTH=600 AutoGenerateColumns="False" Style="z-index: 103; left: 0px; position: absolute; top: 30px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Rank" HeaderText="1"   >  <ItemStyle HorizontalAlign="center" />   </asp:BoundField>  
                <asp:BoundField DataField="DepName" HeaderText="1"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="Hour" HeaderText="10"   >  <ItemStyle HorizontalAlign="right"   />   </asp:BoundField>  
                <asp:BoundField DataField="HourTotal" HeaderText="10"   >  <ItemStyle HorizontalAlign="right"   />   </asp:BoundField>  
                <asp:BoundField DataField="HourPer" HeaderText="10"   >  <ItemStyle HorizontalAlign="right"   />   </asp:BoundField>  
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>

        <asp:GridView ID="GridView1" runat="server" width="1000" AutoGenerateColumns="False" Style="z-index: 103; left:0px; position: absolute; top: 150px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Description" HeaderText="TOP" >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
            
                <asp:HyperLinkField DataNavigateUrlFields="NoURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="No" HeaderText="" Target="_blank">
                       <ItemStyle HorizontalAlign="CENTER" />
                </asp:HyperLinkField>
                <asp:BoundField DataField="DepName" HeaderText="TOP" >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="ApplyQty" HeaderText="COUNT"  >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="OrderQty" HeaderText="COUNT"  >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="NoOrderQty" HeaderText="COUNT"  >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="Percen" HeaderText="COUNT"  >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="NoworkHour1" HeaderText="COUNT"  >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="NoworkHour2" HeaderText="COUNT"  >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
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
