<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ISIPNotRegister.aspx.vb" Inherits="ISIPNotRegister" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Magement-ISIP Not Register</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  底圖                                                                                ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/ISIPNotRegister.png" style="z-index: 1; left: 16px; position: absolute;top: 8px; width: 135px; height: 45px;" />
        ;<!-- ****************************************************************************************** --><!-- **                                                                             --><!-- ****************************************************************************************** -->
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
        &nbsp;&nbsp;

		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:Button ID="BFind" runat="server" Height="25px" Style="z-index: 104; left: 552px;
            position: absolute; top: 56px" Text="Go" Width="45px" />
    
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Style="z-index: 103; width:1000px; left: 48px; position: absolute; top: 88px" CellPadding="3" Font-Size="10pt" GridLines="Vertical"  ForeColor="Black" BackColor="White" >
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
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 1056px; position: absolute; top: 88px" CellPadding="3" Font-Size="10pt" GridLines="Vertical"  ForeColor="Black" BackColor="White" >
            <Columns>

                <asp:BoundField DataField="NoData" HeaderText=""  />          

                <asp:HyperLinkField DataNavigateUrlFields="NoURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="No" HeaderText="色" Target="_blank"  >
                </asp:HyperLinkField>

                <asp:BoundField DataField="Date" HeaderText=""  />          
                <asp:BoundField DataField="Name" HeaderText=""  />    
                <asp:BoundField DataField="Code" HeaderText=""  />          
                <asp:BoundField DataField="Slider"  />            
                <asp:BoundField DataField="ItemName"  />            
           
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#7FFF00" Font-Bold="false" ForeColor="black" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>
        <asp:TextBox ID="DOther" runat="server" BackColor="white" BorderStyle="None" Height="18px"
            MaxLength="7" Style="z-index: 126; left: 608px; position: absolute; top: 56px;
            text-align: left" Width="312px">搜尋件數過多效能差，請使用篩選</asp:TextBox>
    </div>
    </form>
</body>
</html>
