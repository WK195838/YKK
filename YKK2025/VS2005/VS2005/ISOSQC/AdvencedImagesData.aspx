<%@ Page Language="VB" AutoEventWireup="false" CodeFile="AdvencedImagesData.aspx.vb" Inherits="AdvencedImagesData" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Images Data</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
		<!-- ****************************************************************************************** -->
		<!-- ** Text                                                                                -->
		<!-- ****************************************************************************************** -->
            <asp:TextBox ID="DLabel1" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
                Height="24px" ReadOnly="True" Style="z-index: 126; left:8px; position: absolute;
                top: 8px; text-align: left" Width="208px" Font-Bold="True" Font-Size="Larger" TextMode="SingleLine" Font-Overline="False" Font-Underline="True"></asp:TextBox>

		<!-- ****************************************************************************************** -->
		<!-- ** Images Data                                                                                -->
		<!-- ****************************************************************************************** -->
		<asp:Image ID="DImages" Runat="server" Height="230" Width="200" ImageUrl="" Style="z-index: 103; left: 8px; position: absolute; top: 40px" BorderStyle="Solid" BorderWidth="3px" BorderColor="Black" > </asp:Image>

		<!-- ****************************************************************************************** -->
		<!-- ** List                                                                                -->
		<!-- ****************************************************************************************** -->
            <asp:GridView ID="GridView1" runat="server"  AutoGenerateColumns="False"  width="1300" Style="z-index: 103; left: 8px; position: absolute; top: 288px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px">
            <Columns>
                <asp:BoundField DataField="MMS">
                <ItemStyle Width="30px" />
                </asp:BoundField>
              
                <asp:HyperLinkField DataNavigateUrlFields="NoUrl" DataNavigateUrlFormatString="{0}" 
                DataTextField="No" HeaderText="No" Target="_blank"  >
                    <ItemStyle Width="80px" />
                </asp:HyperLinkField>

                <asp:BoundField DataField="Status"  /> 
                
                <asp:BoundField DataField="Spec"  /> 
                
                <asp:BoundField DataField="Remark"  /> 
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="Chartreuse" Font-Bold="False" ForeColor="Black" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>
		<!-- ****************************************************************************************** -->    
    </div>
    </form>
</body>
</html>
