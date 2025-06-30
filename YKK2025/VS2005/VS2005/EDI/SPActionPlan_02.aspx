<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SPActionPlan_02.aspx.vb" Inherits="SPActionPlan_02" %>

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
		<!-- ** Text                                                                              -->
		<!-- ****************************************************************************************** -->
        <asp:TextBox ID="TextBox8" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 16px; position: absolute;
            top: 8px; text-align: left" Width="72px">SP No.</asp:TextBox>
        <asp:TextBox ID="DKSPNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" MaxLength="16" Style="z-index: 126; left: 88px; position: absolute;
            top: 8px; text-align: left" Width="192px"></asp:TextBox>
		<!-- ****************************************************************************************** -->
		<!-- ** Button                                                                              -->
		<!-- ****************************************************************************************** -->
        <asp:Button ID="BFind" runat="server" Height="25px" Style="z-index: 104; left: 296px;
            position: absolute; top: 8px" Text="Go" Width="45px" />

		<!-- ****************************************************************************************** -->
		<!-- ** CheckBox                                                                              -->
		<!-- ****************************************************************************************** -->
        <asp:CheckBox ID="AtCloseCust" runat="server" style="z-index: 174; left: 550px; position: absolute; top: 20px" Font-Size="9pt" Text="Close Windows" Width="130px" AutoPostBack="True" Font-Bold="true" />
        <asp:CheckBox ID="AtCloseYKK" runat="server" style="z-index: 174; left: 550px; position: absolute; top: 20px" Font-Size="9pt" Text="Close Windows" Width="130px" AutoPostBack="True" Font-Bold="true" />

		<!-- ****************************************************************************************** -->
		<!-- ** (GridView1)         table-layout:fixed;        width="1000px"                                                 -->
		<!-- ****************************************************************************************** -->
            <asp:GridView ID="GridView1" runat="server" WIDTH="1700px" AutoGenerateColumns="False"  Style="z-index: 103; left: 8px; position: absolute; top: 40px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
                <Columns>
                    <asp:BoundField DataField="Item"  />
                    <asp:BoundField DataField="ItemName"  />
                    <asp:BoundField DataField="Color"  />
                    <asp:BoundField DataField="Keep"  />
                    
                    <asp:HyperLinkField DataNavigateUrlFields="SPNoURL" DataNavigateUrlFormatString="{0}" 
                    DataTextField="SPNo" HeaderText="Link" Target="_blank"  >
                        <HeaderStyle HorizontalAlign="left" />
                        <ItemStyle HorizontalAlign="left" />
                    </asp:HyperLinkField>

                    
                    <asp:CommandField ShowSelectButton="True" SelectText="...." >                                        
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:CommandField>
                   
                   <asp:CommandField ShowEditButton="True"  EditText="...." >                                        
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:CommandField>

                    <asp:BoundField DataField="ActionName"  />

                    <asp:BoundField DataField="N_Content"  >  <ItemStyle HorizontalAlign="left" Width="100px" />   </asp:BoundField> 
                    <asp:BoundField DataField="N_Qty" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="60px" />   </asp:BoundField>   
                    <asp:BoundField DataField="N_Yobi1"  >  <ItemStyle HorizontalAlign="left" Width="100px" />   </asp:BoundField> 
                    
                    <asp:BoundField DataField="N1_Content" >  <ItemStyle HorizontalAlign="left" Width="100px" />   </asp:BoundField> 
                    <asp:BoundField DataField="N1_Qty" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="60px"  />   </asp:BoundField>   
                    <asp:BoundField DataField="N1_Yobi1"  >  <ItemStyle HorizontalAlign="left" Width="100px" />   </asp:BoundField> 
                    
                    <asp:BoundField DataField="N2_Content" >  <ItemStyle HorizontalAlign="left" Width="100px" />   </asp:BoundField> 
                    <asp:BoundField DataField="N2_Qty" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="60px"  />   </asp:BoundField>  
                    <asp:BoundField DataField="N2_Yobi1"  >  <ItemStyle HorizontalAlign="left" Width="100px" />   </asp:BoundField> 
                     
                    <asp:BoundField DataField="N3_Content" >  <ItemStyle HorizontalAlign="left" Width="100px" />   </asp:BoundField> 
                    <asp:BoundField DataField="N3_Qty" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="60px"  />   </asp:BoundField>   
                    <asp:BoundField DataField="N3_Yobi1"  >  <ItemStyle HorizontalAlign="left" Width="100px" />   </asp:BoundField> 
                    
                    <asp:BoundField DataField="N4_Content" >  <ItemStyle HorizontalAlign="left" Width="100px" />   </asp:BoundField> 
                    <asp:BoundField DataField="N4_Qty" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right"  Width="60px" />   </asp:BoundField>   
                    <asp:BoundField DataField="N4_Yobi1"  >  <ItemStyle HorizontalAlign="left" Width="100px" />   </asp:BoundField> 

                    <asp:BoundField DataField="Balance" DataFormatString="{0:#,0.00}" >  <ItemStyle HorizontalAlign="right" Width="80px" />   </asp:BoundField>   

                    <asp:BoundField DataField="wSPNo"  />

                </Columns>
                <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
                <FooterStyle BackColor="#CCCCCC" />
                <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
                <SelectedRowStyle BackColor="Chartreuse" Font-Bold="False" ForeColor="Black" />
                <AlternatingRowStyle BackColor="#CCCCCC" />
            </asp:GridView>
		<!-- ****************************************************************************************** -->
		<!-- ** Custer Inf.                                                                                -->
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
		<!-- ****************************************************************************************** -->
		<!-- ** LS Inf.                                                                                -->
		<!-- ****************************************************************************************** -->
        <asp:GridView ID="GridView3" runat="server" WIDTH="900px" AutoGenerateColumns="False" Style="z-index: 103; left: 150px; position: absolute; top: 250px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="Red" BorderStyle="Solid" BorderWidth="5px" >
            <Columns>
                <asp:BoundField DataField="GR_09" HeaderText=""  />        
                <asp:BoundField DataField="LSNo" HeaderText=""  />        
                <asp:BoundField DataField="MinimumStock" HeaderText=""  />          
                <asp:BoundField DataField="N_ScheProd" HeaderText=""  />          
                <asp:BoundField DataField="N_OnProd" HeaderText=""  />          
                <asp:BoundField DataField="N_FreeInv" HeaderText=""  />    
                <asp:BoundField DataField="N_KeepInv" HeaderText=""  />          
                <asp:BoundField DataField="Description" HeaderText=""  />          
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
