<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SPActionPlan_03.aspx.vb" Inherits="SPActionPlan_03" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >

<head id="Head1" runat="server">
    <title>Shopping Action Plan</title>
	    <script language="javascript" src="ForProject.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
		<!-- ****************************************************************************************** -->
		<!-- ** CheckBox                                                                              -->
		<!-- ****************************************************************************************** -->
        <asp:CheckBox ID="AtClose" runat="server" style="z-index: 174; left: 550px; position: absolute; top: 20px" Font-Size="9pt" Text="Close Windows" Width="130px" AutoPostBack="True" Font-Bold="true" />

		<!-- ****************************************************************************************** -->
		<!-- ** (GridView1)                                                -->
		<!-- ****************************************************************************************** -->
            <asp:GridView ID="GridView1" runat="server" WIDTH="1700px" AutoGenerateColumns="False"  Style="z-index: 103; left: 8px; position: absolute; top: 40px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
                <Columns>
                    <asp:BoundField DataField="Item"  />
                    <asp:BoundField DataField="ItemName"  />
                    <asp:BoundField DataField="Color"  />
                    <asp:BoundField DataField="Keep"  />
                    <asp:BoundField DataField="YY"  />
                    <asp:BoundField DataField="ActionName"  />

                    <asp:BoundField DataField="01_Qty" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="60px" />   </asp:BoundField>   
                    <asp:BoundField DataField="02_Qty" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="60px" />   </asp:BoundField>   
                    <asp:BoundField DataField="03_Qty" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="60px" />   </asp:BoundField>   
                    <asp:BoundField DataField="04_Qty" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="60px" />   </asp:BoundField>   
                    <asp:BoundField DataField="05_Qty" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="60px" />   </asp:BoundField>   
                    <asp:BoundField DataField="06_Qty" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="60px" />   </asp:BoundField>   

                    <asp:BoundField DataField="07_Qty" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="60px" />   </asp:BoundField>   
                    <asp:BoundField DataField="08_Qty" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="60px" />   </asp:BoundField>   
                    <asp:BoundField DataField="09_Qty" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="60px" />   </asp:BoundField>   
                    <asp:BoundField DataField="10_Qty" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="60px" />   </asp:BoundField>   
                    <asp:BoundField DataField="11_Qty" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="60px" />   </asp:BoundField>   
                    <asp:BoundField DataField="12_Qty" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="60px" />   </asp:BoundField>   

                    <asp:BoundField DataField="Balance" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="80px" />   </asp:BoundField>   
                </Columns>
                <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
                <FooterStyle BackColor="#CCCCCC" />
                <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
                <SelectedRowStyle BackColor="Chartreuse" Font-Bold="False" ForeColor="Black" />
                <AlternatingRowStyle BackColor="#CCCCCC" />
            </asp:GridView>
		<!-- ****************************************************************************************** -->
		<!-- ** Detail Inf.                                                                                -->
		<!-- ****************************************************************************************** -->
        <asp:GridView ID="GridView2" runat="server" WIDTH="800px" AutoGenerateColumns="False" Style="z-index: 103; left: 150px; position: absolute; top: 250px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="Red" BorderStyle="Solid" BorderWidth="5px" >
            <Columns>
                <asp:BoundField DataField="Y_J1" HeaderText=""  />          
                <asp:BoundField DataField="FCTNo" HeaderText=""  />          
                <asp:BoundField DataField="C_Code" HeaderText=""  />          
                <asp:BoundField DataField="C_Color" HeaderText=""  />          
                <asp:BoundField DataField="C_SPECIALREQUEST" HeaderText=""  />        
                <asp:BoundField DataField="C_Insert" HeaderText=""  />      
                      
                <asp:BoundField DataField="C_Season" HeaderText=""  />          
                <asp:BoundField DataField="C_SHORTENLT" HeaderText=""  />  
                <asp:BoundField DataField="C_Buyer" HeaderText=""  />          
                
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#7FFF00" Font-Bold="false" ForeColor="black" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>
    
    </div>
    </form>
</body>

</html>
