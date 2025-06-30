<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MagApplyHistory.aspx.vb" Inherits="MagApplyHistory" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Magement-Apply History</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  底圖                                                                                ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/MagApplyHistory.png" style="z-index: 1; left: 16px; position: absolute;top: 8px; width: 135px; height: 45px;" />

		<!-- ****************************************************************************************** -->
		<!-- **                                                                                        -->
		<!-- ****************************************************************************************** -->
        <asp:TextBox ID="TextBox4" runat="server" ForeColor="Black" 
            Height="18px" MaxLength="8" Style="z-index: 126; left: 152px;
            position: absolute; top: 48px; text-align: left" Width="100px" BorderStyle="None">20240201</asp:TextBox>

        <asp:TextBox ID="TextBox5" runat="server" 
            ForeColor="Black" Height="18px" MaxLength="8" Style="z-index: 126; left: 288px;
            position: absolute; top: 48px; text-align: left" Width="100px" BorderStyle="None">20240203</asp:TextBox>
            
		<!-- ****************************************************************************************** -->
		<!-- **                                                                                         -->
		<!-- ****************************************************************************************** -->
        <asp:TextBox ID="TextBox1" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 48px; position: absolute;
            top: 72px; text-align: left" Width="104px">Reg. Date</asp:TextBox>
        <asp:TextBox ID="TextBox3" runat="server"  BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 272px; position: absolute;
            top: 80px; text-align: left" Width="24px">～</asp:TextBox>
            
        <asp:TextBox ID="DKADDStart" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="8" Style="z-index: 126; left: 160px;
            position: absolute; top: 72px; text-align: left" Width="100px"></asp:TextBox>

        <asp:TextBox ID="DKADDEnd" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="8" Style="z-index: 126; left: 296px;
            position: absolute; top: 72px; text-align: left" Width="100px"></asp:TextBox>


        <asp:TextBox ID="TextBox2" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 416px; position: absolute;
            top: 72px; text-align: left" Width="104px">Search</asp:TextBox>

        <asp:TextBox ID="DKKey" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="20" Style="z-index: 126; left: 528px;
            position: absolute; top: 72px; text-align: left" Width="100px"></asp:TextBox>


        <asp:Button ID="BFind" runat="server" Height="25px" Style="z-index: 104; left: 768px;
            position: absolute; top: 72px" Text="Go" Width="45px" />
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:CheckBox ID="AtIncludeAdmin" runat="server" AutoPostBack="false" Font-Bold="true" Font-Size="9pt"
            Style="z-index: 174; left: 640px; position: absolute; top: 48px" Text="Include Admin."
            Width="112px" />
        <asp:CheckBox ID="AtIncludeSystem" runat="server" AutoPostBack="false" Font-Bold="true" Font-Size="9pt"
            Style="z-index: 174; left: 640px; position: absolute; top: 72px" Text="Include System"
            Width="112px" />
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 48px; position: absolute; top: 104px" CellPadding="3" Font-Size="10pt" GridLines="Vertical"  ForeColor="Black" BackColor="White" >
            <Columns>
                <asp:BoundField DataField="CreateUserName" HeaderText=""  />          
                <asp:BoundField DataField="ModifyUserName" HeaderText=""  />          
                <asp:BoundField DataField="CreateTimeDesc"  />            

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
