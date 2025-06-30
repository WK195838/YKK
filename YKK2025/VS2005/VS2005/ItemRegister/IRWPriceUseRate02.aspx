<%@ Page Language="VB" AutoEventWireup="false" CodeFile="IRWPriceUseRate02.aspx.vb" Inherits="IRWPriceUseRate02" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >

<head id="Head1" runat="server">
    <title>Price Use Rate-02</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <asp:ImageButton ID="BExcel" runat="server" Height="24px" ImageUrl="Images\msexcel.gif"
            Style="z-index: 103; left: 6px; position: absolute; top: 2px" Width="24px" />
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"  Width="1200px" Style="z-index: 103; left: 0px; position: absolute; top: 25px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="200" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="DepName" HeaderText="" ItemStyle-HorizontalAlign="Left"/> 
                <asp:BoundField DataField="EmpName" HeaderText="" ItemStyle-HorizontalAlign="Left"/> 

                <asp:BoundField DataField="MM1_YES" HeaderText="" DataFormatString="{0:N0}" ItemStyle-HorizontalAlign="Right"/> 
                <asp:BoundField DataField="MM1_NO" HeaderText=""  DataFormatString="{0:N0}" ItemStyle-HorizontalAlign="Right"/> 
                <asp:BoundField DataField="MM1_PER" HeaderText=""  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>

                <asp:BoundField DataField="MM2_YES" HeaderText="" DataFormatString="{0:N0}" ItemStyle-HorizontalAlign="Right"/> 
                <asp:BoundField DataField="MM2_NO" HeaderText="" DataFormatString="{0:N0}" ItemStyle-HorizontalAlign="Right"/> 
                <asp:BoundField DataField="MM2_PER" HeaderText=""  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>

                <asp:BoundField DataField="MM3_YES" HeaderText=""  DataFormatString="{0:N0}" ItemStyle-HorizontalAlign="Right"/> 
                <asp:BoundField DataField="MM3_NO" HeaderText=""  DataFormatString="{0:N0}" ItemStyle-HorizontalAlign="Right"/> 
                <asp:BoundField DataField="MM3_PER" HeaderText=""  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>

                <asp:BoundField DataField="MM4_YES" HeaderText="" DataFormatString="{0:N0}" ItemStyle-HorizontalAlign="Right"/> 
                <asp:BoundField DataField="MM4_NO" HeaderText=""  DataFormatString="{0:N0}" ItemStyle-HorizontalAlign="Right"/> 
                <asp:BoundField DataField="MM4_PER" HeaderText=""  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>

                <asp:BoundField DataField="MM5_YES" HeaderText=""  DataFormatString="{0:N0}" ItemStyle-HorizontalAlign="Right"/> 
                <asp:BoundField DataField="MM5_NO" HeaderText=""  DataFormatString="{0:N0}" ItemStyle-HorizontalAlign="Right"/> 
                <asp:BoundField DataField="MM5_PER" HeaderText=""  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>

                <asp:BoundField DataField="MM6_YES" HeaderText=""  DataFormatString="{0:N0}" ItemStyle-HorizontalAlign="Right"/> 
                <asp:BoundField DataField="MM6_NO" HeaderText=""  DataFormatString="{0:N0}" ItemStyle-HorizontalAlign="Right"/> 
                <asp:BoundField DataField="MM6_PER" HeaderText=""  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>

                <asp:BoundField DataField="TOTAL_YES" HeaderText="" DataFormatString="{0:N0}" ItemStyle-HorizontalAlign="Right"/> 
                <asp:BoundField DataField="TOTAL_NO" HeaderText=""  DataFormatString="{0:N0}" ItemStyle-HorizontalAlign="Right"/> 
                <asp:BoundField DataField="TOTAL_PER" HeaderText=""  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>

                <asp:BoundField DataField="DepName" HeaderText=""  />

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="White" />
        </asp:GridView>
    
    </div>
    </form>
</body>

</html>
