<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SPActionPlan_01.aspx.vb" Inherits="SPActionPlan_01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Shopping Action Plan</title>
	    <script language="javascript" src="ForProject.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
		<!-- ****************************************************************************************** -->
		<!-- ** (GridView1)         table-layout:fixed;        width="1000px"                                                 -->
		<!-- ****************************************************************************************** -->
            <asp:GridView ID="GridView1" runat="server" WIDTH="2300px" AutoGenerateColumns="False"  Style="z-index: 103; left: 8px; position: absolute; top: 20px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
                <Columns>
                    <asp:BoundField DataField="Item"  />
                    <asp:BoundField DataField="ItemName"  />
                    <asp:BoundField DataField="Color"  />
                    <asp:BoundField DataField="Keep"  />
                    <asp:BoundField DataField="SPNo"  />
                    <asp:BoundField DataField="SPSubNo"  />

                    <asp:BoundField DataField="C_Code"  />
                    <asp:BoundField DataField="C_Color"  />
                    <asp:BoundField DataField="C_Special"  />
                    <asp:BoundField DataField="C_KeepCode"  />
                    <asp:BoundField DataField="C_Season"  />
                    <asp:BoundField DataField="C_Cust"  />

                    <asp:BoundField DataField="SchePQ"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                    <asp:BoundField DataField="OnPQ"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                    <asp:BoundField DataField="FreeQ"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                    <asp:BoundField DataField="KeepQ"   >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                    <asp:BoundField DataField="TotalQ"   >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                        
                    <asp:BoundField DataField="N_FC"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                    <asp:BoundField DataField="N_Action"  />
                    <asp:BoundField DataField="N1_FC"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                    <asp:BoundField DataField="N1_Action"  />
                    <asp:BoundField DataField="N2_FC"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                    <asp:BoundField DataField="N2_Action"  />
                    <asp:BoundField DataField="N3_FC"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                    <asp:BoundField DataField="N3_Action"  />
                    <asp:BoundField DataField="N4_FC"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                    <asp:BoundField DataField="N4_Action"  />

                    <asp:HyperLinkField DataNavigateUrlFields="PILURL" DataNavigateUrlFormatString="{0}" 
                    DataTextField="PIL" HeaderText="Link" Target="_blank"  >
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:HyperLinkField>

                </Columns>
                <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
                <FooterStyle BackColor="#CCCCCC" />
                <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
                <SelectedRowStyle BackColor="Chartreuse" Font-Bold="False" ForeColor="Black" />
                <AlternatingRowStyle BackColor="#CCCCCC" />
            </asp:GridView>
    
    </div>
    </form>
</body>
</html>
