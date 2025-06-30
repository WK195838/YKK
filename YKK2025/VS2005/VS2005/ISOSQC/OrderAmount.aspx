<%@ Page Language="VB" AutoEventWireup="false" CodeFile="OrderAmount.aspx.vb" Inherits="OrderAmount" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Magement-Order Amount</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  底圖                                                                                ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/OrderAmount.png" style="z-index: 1; left: 16px; position: absolute;top: 8px; width: 135px; height: 45px;" />
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:TextBox ID="TextBox1" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 48px; position: absolute;
            top: 56px; text-align: left" Width="80px">Puller Code</asp:TextBox>
        <asp:TextBox ID="DKPullerCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="7" Style="z-index: 126; left: 128px;
            position: absolute; top: 56px; text-align: left" Width="100px"></asp:TextBox>
        <asp:TextBox ID="TextBox11" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 232px; position: absolute;
            top: 56px; text-align: left" Width="80px">Color</asp:TextBox>
        <asp:TextBox ID="DKColor" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="7" Style="z-index: 126; left: 312px;
            position: absolute; top: 56px; text-align: left" Width="48px"></asp:TextBox>
        <asp:TextBox ID="TextBox8" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 360px; position: absolute;
            top: 56px; text-align: left" Width="80px">Buyer</asp:TextBox>
        <asp:TextBox ID="DKBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="7" Style="z-index: 126; left: 440px;
            position: absolute; top: 56px; text-align: left" Width="100px"></asp:TextBox>
        <asp:TextBox ID="TextBox12" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 544px; position: absolute;
            top: 56px; text-align: left" Width="80px">Search</asp:TextBox>
        <asp:TextBox ID="DKOther" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="50" Style="z-index: 126; left: 624px;
            position: absolute; top: 56px; text-align: left" Width="100px"></asp:TextBox>
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:Button ID="BFind" runat="server" Height="25px" Style="z-index: 104; left: 744px;
            position: absolute; top: 56px" Text="Go" Width="45px" />

		<!-- ****************************************************************************************** -->
		<!-- ** CheckBox                                                                              -->
		<!-- ****************************************************************************************** -->
        <asp:CheckBox ID="AtCloseW" runat="server" style="z-index: 174; left: 1100px; position: absolute; top: 56px" Font-Size="9pt" Text="Close Windows" Width="130px" AutoPostBack="True" Font-Bold="true" />
    
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:GridView ID="GridView1" runat="server" WIDTH="2000" AutoGenerateColumns="False" Style="z-index: 103; left: 48px; position: absolute; top: 88px" CellPadding="3" Font-Size="10pt" GridLines="Vertical"  ForeColor="Black" BackColor="White" >
            <Columns>
                    <asp:HyperLinkField DataNavigateUrlFields="PullerURL" DataNavigateUrlFormatString="{0}" 
                    DataTextField="Puller" HeaderText="色" Target="_blank"  >
                    </asp:HyperLinkField>

                    <asp:BoundField DataField="Color" HeaderText=""  />          
                    <asp:BoundField DataField="ColorName" HeaderText=""  />    
                    <asp:BoundField DataField="BYColorCode" HeaderText=""  />          
                    <asp:BoundField DataField="Buyer"  />            
                    <asp:BoundField DataField="BuyerName"  />            
                    <asp:BoundField DataField="Remark"  />            
                    <asp:BoundField DataField="DTM_YOBI1"  />            
                    <asp:BoundField DataField="DTM_YOBI2"  />            
                    <asp:BoundField DataField="IRW_YOBI1"  />            
                    <asp:BoundField DataField="IRW_YOBI2"  />            
                    <asp:BoundField DataField="OR_YOBI2"  />            

                    <asp:CommandField ShowEditButton="True" SelectText="..." EditText="Link" />
                    <asp:BoundField DataField="QTY" HeaderText="" DataFormatString="{0:#,0}">  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                    <asp:BoundField DataField="AMT" HeaderText="" DataFormatString="{0:#,0}" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                    <asp:BoundField DataField="QTY1Y" HeaderText="" DataFormatString="{0:#,0}" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                    <asp:BoundField DataField="AMT1Y" HeaderText="" DataFormatString="{0:#,0}">  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                    <asp:BoundField DataField="QTY2Y" HeaderText="" DataFormatString="{0:#,0}" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                    <asp:BoundField DataField="AMT2Y" HeaderText="" DataFormatString="{0:#,0}" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                    <asp:BoundField DataField="QTY3Y" HeaderText="" DataFormatString="{0:#,0}" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                    <asp:BoundField DataField="AMT3Y" HeaderText="" DataFormatString="{0:#,0}" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                    <asp:BoundField DataField="QTY4Y" HeaderText="" DataFormatString="{0:#,0}" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                    <asp:BoundField DataField="AMT4Y" HeaderText="" DataFormatString="{0:#,0}">  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                    <asp:BoundField DataField="QTY5Y" HeaderText="" DataFormatString="{0:#,0}" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                    <asp:BoundField DataField="AMT5Y" HeaderText="" DataFormatString="{0:#,0}" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                    <asp:BoundField DataField="Puller" HeaderText=""  />          
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#7FFF00" Font-Bold="false" ForeColor="black" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>    

    
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 900px; position: absolute; top: 200px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="Red" BorderStyle="Solid" BorderWidth="5px" >
            <Columns>
                    <asp:BoundField DataField="Puller" HeaderText=""  />          
                    <asp:BoundField DataField="Color" HeaderText=""  /> 
                    <asp:BoundField DataField="Month" HeaderText=""  />    
                    <asp:BoundField DataField="Slider" HeaderText=""  />    
                    <asp:BoundField DataField="Custmer" HeaderText=""  />          
                    <asp:BoundField DataField="CustmerName"  />            
                    <asp:BoundField DataField="Buyer"  />            
                    <asp:BoundField DataField="BuyerName"  />            
                    <asp:BoundField DataField="Quantity" HeaderText="" DataFormatString="{0:#,0}">  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                    <asp:BoundField DataField="Amount" HeaderText="" DataFormatString="{0:#,0}" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
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
