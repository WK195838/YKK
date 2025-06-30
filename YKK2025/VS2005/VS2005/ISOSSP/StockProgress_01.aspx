<%@ Page Language="VB" AutoEventWireup="false" CodeFile="StockProgress_01.aspx.vb" Inherits="StockProgress_01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>SHOPPING 2024 Ver1.0</title>
	    <script language="javascript" src="ForProject.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
		<!-- ****************************************************************************************** -->
		<!-- ** System
		<!-- ****************************************************************************************** -->
        <img src="iMages/SPSLOGO.png" style="z-index: 1; left: 6px; position: absolute;top: 6px; width: 500px; height: 32px;"/>

		<!-- ****************************************************************************************** -->
		<!-- ** Button                                                                                  -->
		<!-- ****************************************************************************************** -->
	    <!-- KP -->
        <asp:HyperLink style="Z-INDEX: 104; LEFT: 770px; POSITION: absolute; TOP: 8px; TEXT-ALIGN: center" id="LKPTool" runat="server" 
            Width="80px" Height="24px" Target="_blank" ForeColor="Blue" Font-Size="Larger" Font-Bold="True">KP Check</asp:HyperLink>
	    <!-- Shopping List -->
        <asp:HyperLink style="Z-INDEX: 104; LEFT: 864px; POSITION: absolute; TOP: 8px; TEXT-ALIGN: center" id="LShoppingList" runat="server" 
            Width="128px" Height="24px" Target="_blank" ForeColor="Blue" Font-Size="Larger" Font-Bold="True">Shopping List</asp:HyperLink>

	    <!-- Go -->
        <asp:Button ID="BFind" runat="server" Height="25px" Style="z-index: 104; left: 688px;
                position: absolute; top: 48px" Text="Go" Width="45px" />
	    <!-- Material -->
        <asp:ImageButton ID="BMaterial" runat="server" style="z-index: 1; left: 1168px; position: absolute;top: 16px;" ImageUrl="~/iMages/MaterialPlan.png" Height="56px" Width="154px" />

		<!-- ****************************************************************************************** -->
		<!-- ** Puller Key                                                                              -->
		<!-- ****************************************************************************************** -->
            <asp:TextBox ID="TextBox1" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
                Height="18px" ReadOnly="True" Style="z-index: 126; left: 304px; position: absolute;
                top: 48px; text-align: left" Width="80px">Keep Code</asp:TextBox>
            <asp:TextBox ID="DKKeepCode" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 126; left: 384px; position: absolute;
                top: 48px; text-align: left" Width="96px" MaxLength="10"></asp:TextBox>

            <asp:TextBox ID="TextBox8" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
                Height="18px" ReadOnly="True" Style="z-index: 126; left:16px; position: absolute;
                top: 48px; text-align: left" Width="72px">Code</asp:TextBox>
            <asp:TextBox ID="DKCode" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 126; left: 88px; position: absolute;
                top: 48px; text-align: left" Width="72px" MaxLength="7"></asp:TextBox>

            <asp:TextBox ID="TextBox2" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
                Height="18px" ReadOnly="True" Style="z-index: 126; left:168px; position: absolute;
                top: 48px; text-align: left" Width="80px">Color</asp:TextBox>
            <asp:TextBox ID="DKColor" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 126; left: 248px; position: absolute;
                top: 48px; text-align: left" Width="48px" MaxLength="5"></asp:TextBox>

            <asp:TextBox ID="TextBox12" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
                Height="18px" ReadOnly="True" Style="z-index: 126; left:488px; position: absolute;
                top: 48px; text-align: left" Width="48px">Search</asp:TextBox>
            <asp:TextBox ID="DKOther" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 126; left: 544px; position: absolute;
                top: 48px; text-align: left" Width="128px" MaxLength="50"></asp:TextBox>

            <asp:TextBox ID="DKSource" runat="server" BackColor="white" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 126; left: 624px; position: absolute;
                top: 8px; text-align: left" Width="48px" MaxLength="7"></asp:TextBox>

            <asp:TextBox ID="DKPuller" runat="server" BackColor="white" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 126; left: 680px; position: absolute;
                top: 8px; text-align: left" Width="72px" MaxLength="7"></asp:TextBox>

            <asp:TextBox ID="DUpdateTime" runat="server" BackColor="white" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 126; left: 368px; position: absolute;
                top: 8px; text-align: left" Width="240px" MaxLength="7" Font-Bold="True">Data update time 2024/1212/ 23:23:23</asp:TextBox>

		<!-- ****************************************************************************************** -->
		<!-- ** CheckBox                                                                              -->
		<!-- ****************************************************************************************** -->
        <asp:CheckBox ID="AtCloseHistory" runat="server" style="z-index: 174; left: 744px; position: absolute; top: 48px" Font-Size="9pt" Text="Close Windows" Width="104px" AutoPostBack="True" Font-Bold="true" />
		
		
		<!-- ****************************************************************************************** -->
		<!-- ** Automate                                                                         -->
		<!-- ****************************************************************************************** -->


		<!-- ****************************************************************************************** -->
		<!-- ** (GridView1)         table-layout:fixed;        width="1000px"                                                 -->
		<!-- ****************************************************************************************** -->
            <asp:GridView ID="GridView1" runat="server" WIDTH="2300px" AutoGenerateColumns="False"  Style="z-index: 103; left: 8px; position: absolute; top: 80px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
                <Columns>
                    <asp:BoundField DataField="Item"  />
                    <asp:BoundField DataField="ItemName"  />
                    <asp:BoundField DataField="Color"  />
                    <asp:BoundField DataField="Keep"  />
                    <asp:BoundField DataField="ScheProcS"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                    <asp:BoundField DataField="OnProcS"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                    <asp:BoundField DataField="FreeS"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                    <asp:BoundField DataField="KeepS"   >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                    <asp:BoundField DataField="TotalS"   >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                        
                    <asp:HyperLinkField DataNavigateUrlFields="FCLUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="FCL" HeaderText="Link" Target="_blank"  >
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:HyperLinkField>
                        
                    <asp:HyperLinkField DataNavigateUrlFields="FDMUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="FDM" HeaderText="Link" Target="_blank"  >
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:HyperLinkField>
                        
                    <asp:HyperLinkField DataNavigateUrlFields="PRODSTSUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="PRODSTS" HeaderText="Link" Target="_blank"  >
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:HyperLinkField>

                    <asp:CommandField ShowEditButton="True" SelectText="...." EditText="...."   >
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:CommandField>

                    <asp:BoundField DataField="N21IH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                    <asp:BoundField DataField="N21OH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>              
                    <asp:BoundField DataField="N22IH"   >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>                
                    <asp:BoundField DataField="N22OH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>                
                    <asp:BoundField DataField="N23IH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>               
                    <asp:BoundField DataField="N23OH"   >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>               
                    <asp:BoundField DataField="N24IH"   >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>               
                    <asp:BoundField DataField="N24OH"   >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                                    
                    <asp:BoundField DataField="N11IH"   >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>              
                    <asp:BoundField DataField="N11OH"   >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>               
                    <asp:BoundField DataField="N12IH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>               
                    <asp:BoundField DataField="N12OH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>               
                    <asp:BoundField DataField="N13IH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>              
                    <asp:BoundField DataField="N13OH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>              
                    <asp:BoundField DataField="N14IH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>              
                    <asp:BoundField DataField="N14OH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>             
                                    
                    <asp:BoundField DataField="N01IH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>             
                    <asp:BoundField DataField="N01OH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>              
                    <asp:BoundField DataField="N02IH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>              
                    <asp:BoundField DataField="N02OH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>              
                    <asp:BoundField DataField="N03IH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>              
                    <asp:BoundField DataField="N03OH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>             
                    <asp:BoundField DataField="N04IH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>             
                    <asp:BoundField DataField="N04OH"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>             

                </Columns>
                <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
                <FooterStyle BackColor="#CCCCCC" />
                <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
                <SelectedRowStyle BackColor="Chartreuse" Font-Bold="False" ForeColor="Black" />
                <AlternatingRowStyle BackColor="#CCCCCC" />
            </asp:GridView>

		<!-- ****************************************************************************************** -->
		<!-- ** SPD List                                                                                -->
		<!-- ****************************************************************************************** -->
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 150px; position: absolute; top: 200px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="Red" BorderStyle="Solid" BorderWidth="5px" >
            <Columns>
                <asp:BoundField DataField="ITMC10"  />          
                <asp:BoundField DataField="CLRC10"  />          
                <asp:BoundField DataField="KEPC10"  />          
                
                
                <asp:HyperLinkField DataNavigateUrlFields="ORDN11URL" DataNavigateUrlFormatString="{0}" 
                DataTextField="ORDN11" HeaderText="Link" Target="_blank"  >
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:HyperLinkField>

                <asp:BoundField DataField="CORN11"  />     
                <asp:BoundField DataField="OCNU11"  />     
                
                <asp:BoundField DataField="IORRQ11"  >  <ItemStyle HorizontalAlign="Right"  />   </asp:BoundField>                      
                <asp:BoundField DataField="OORRQ11"  >  <ItemStyle HorizontalAlign="Right"  />   </asp:BoundField>                      
                <asp:BoundField DataField="BORRQ11"  >  <ItemStyle HorizontalAlign="Right"  />   </asp:BoundField>                      
                    
                <asp:BoundField DataField="REMK11"  />     
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="Chartreuse" Font-Bold="False" ForeColor="Black" />
            <AlternatingRowStyle BackColor="#FFF3DB" />
        </asp:GridView>

    </div>
    </form>
</body>
</html>
