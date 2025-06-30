<%@ Page Language="VB" AutoEventWireup="false" CodeFile="IRWNoticePage2310.aspx.vb" Inherits="IRWNoticePage2310" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >

<head id="Head1" runat="server">
    <title>NOTICE PAGE(2310)</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView4" runat="server" WIDTH=700 AutoGenerateColumns="False" Style="z-index: 103; left: 0px; position: absolute; top: 8px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Warning" HeaderText="1"   >  <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField>  

                <asp:BoundField DataField="DepName" HeaderText="1"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="A_YES" HeaderText="2" >  <ItemStyle HorizontalAlign="right" width=50 />   </asp:BoundField>  
                <asp:BoundField DataField="A_NO" HeaderText="3"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
                <asp:BoundField DataField="A_TOTAL" HeaderText="4"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
                <asp:BoundField DataField="A_PERCENT" HeaderText="5"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  

                <asp:BoundField DataField="S_PERCENT" HeaderText="6"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
                <asp:BoundField DataField="S_NO" HeaderText="7"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
                <asp:BoundField DataField="H_NO" HeaderText="8"   >  <ItemStyle HorizontalAlign="right" width=60  />   </asp:BoundField>  
                <asp:BoundField DataField="L_NO" HeaderText="9"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
                <asp:BoundField DataField="B_NO" HeaderText="10"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>
         
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView9 ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView9" runat="server" WIDTH=600 AutoGenerateColumns="False" Style="z-index: 103; left:0px; position: absolute; top: 300px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="DepName" HeaderText="1"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="EmpName" HeaderText="2"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="A_YES" HeaderText="3" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="A_NO" HeaderText="4"   >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="A_TOTAL" HeaderText="5"   >  <ItemStyle HorizontalAlign="right"   />   </asp:BoundField>  
                <asp:BoundField DataField="A_PERCENT" HeaderText="6"   >  <ItemStyle HorizontalAlign="right"   />   </asp:BoundField>  
                <asp:BoundField DataField="G_NO" HeaderText="3" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="L_NO" HeaderText="4"   >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="REMARK" HeaderText="7"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  

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
