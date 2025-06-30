<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ISIP2IRW.aspx.vb" Inherits="ISIP2IRW" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Magement-ISIP-IRW</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  底圖                                                                                ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/ISIP2IRW.png" style="z-index: 1; left: 16px; position: absolute;top: 8px; width: 135px; height: 45px;" />

        <asp:TextBox ID="DOther" runat="server" BackColor="white" BorderStyle="None"
            Height="18px" MaxLength="7" Style="z-index: 126; left: 896px; position: absolute;
            top: 24px; text-align: left" Width="312px">[ALL/ISOS]件數過多時只顯示150件，請使用篩選</asp:TextBox>
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:CheckBox ID="AtALL" runat="server" AutoPostBack="True" Font-Bold="true" Font-Size="9pt"
            Style="z-index: 174; left: 824px; position: absolute; top: 8px" Text="ALL"
            Width="56px" />
        <asp:CheckBox ID="AtISOS" runat="server" AutoPostBack="True" Font-Bold="true" Font-Size="9pt"
            Style="z-index: 174; left: 824px; position: absolute; top: 32px" Text="ISOS"
            Width="56px" />
        <asp:CheckBox ID="AtISIP" runat="server" AutoPostBack="True" Font-Bold="true" Font-Size="9pt"
            Style="z-index: 174; left: 824px; position: absolute; top: 56px" Text="ISIP"
            Width="56px" Checked="True" />
        <asp:CheckBox ID="AtEDX" runat="server" AutoPostBack="True" Font-Bold="true" Font-Size="9pt"
            Style="z-index: 174; left: 736px; position: absolute; top: 56px" Text="EDX=OK"
            Width="80px" Checked="True" />
       
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
        <asp:Button ID="BFind" runat="server" Height="25px" Style="z-index: 104; left: 896px;
            position: absolute; top: 48px" Text="Go" Width="45px" />
    
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 48px; position: absolute; top: 88px" CellPadding="3" Font-Size="10pt" GridLines="Vertical"  ForeColor="Black" BackColor="White" >
            <Columns>

                <asp:BoundField DataField="RD" HeaderText=""  />          
                <asp:BoundField DataField="DTM" HeaderText=""  />    
                <asp:BoundField DataField="EDX" HeaderText=""  />          
                <asp:BoundField DataField="IRW"  />            
                <asp:BoundField DataField="ORDERS"  />      
                 <asp:BoundField DataField="IRW_NO"  />            
                     

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
                <asp:BoundField DataField="Yobi1"  />            
                <asp:BoundField DataField="Yobi2"  />            
           
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
