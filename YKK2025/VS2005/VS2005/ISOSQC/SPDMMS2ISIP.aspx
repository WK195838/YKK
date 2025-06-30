<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SPDMMS2ISIP.aspx.vb" Inherits="SPDMMS2ISIP" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>ISIP-SPD MMS</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  底圖                                                                                ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/SPDMMS2ISIP.png" style="z-index: 1; left: 16px; position: absolute;top: 8px; width: 135px; height: 45px;" />
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:CheckBox ID="AtYesMMS" runat="server" AutoPostBack="True" Font-Bold="true" Font-Size="9pt"
            Style="z-index: 174; left: 240px; position: absolute; top: 56px" Text="全模廢"
            Width="88px" Checked="True" />
        <asp:CheckBox ID="AtNoMMS" runat="server" AutoPostBack="True" Font-Bold="true" Font-Size="9pt"
            Style="z-index: 174; left: 312px; position: absolute; top: 56px" Text="無模廢"
            Width="88px" />
        <asp:CheckBox ID="AtYesNoMMS" runat="server" AutoPostBack="True" Font-Bold="true" Font-Size="9pt"
            Style="z-index: 174; left: 384px; position: absolute; top: 56px" Text="有些模廢"
            Width="88px" />
        <asp:CheckBox ID="AtNoSPD" runat="server" AutoPostBack="True" Font-Bold="true" Font-Size="9pt"
            Style="z-index: 174; left: 464px; position: absolute; top: 56px" Text="無SPD"
            Width="88px" />
        <asp:CheckBox ID="AtALL" runat="server" Font-Bold="true" Font-Size="9pt"
            Style="z-index: 174; left: 528px; position: absolute; top: 56px" Text="ALL"
            Width="56px" AutoPostBack="True" />

        <asp:CheckBox ID="AtCloseRDW" runat="server" style="z-index: 174; left: 424px; position: absolute; top: 152px" Font-Size="9pt" Text="Close RD Windows" Width="130px" AutoPostBack="True" Font-Bold="true" />


		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:TextBox ID="TextBox1" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 48px; position: absolute;
            top: 56px; text-align: left" Width="80px">Puller Code</asp:TextBox>
        <asp:TextBox ID="DKPullerCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="7" Style="z-index: 126; left: 128px;
            position: absolute; top: 56px; text-align: left" Width="100px"></asp:TextBox>
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:Button ID="BFind" runat="server" Height="25px" Style="z-index: 104; left: 592px;
            position: absolute; top: 48px" Text="Go" Width="45px" />
    
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:GridView ID="GridView1" runat="server" width=300 AutoGenerateColumns="False" Style="z-index: 103; left: 48px; position: absolute; top: 88px" CellPadding="3" Font-Size="10pt" GridLines="Vertical"  ForeColor="Black" BackColor="White" >
            <Columns>
                    <asp:CommandField ShowEditButton="True" SelectText="..." EditText="R&D" />
                <asp:BoundField DataField="Puller" HeaderText=""  />          
                <asp:BoundField DataField="RDCount" HeaderText="TOP" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="NoMMSCount" HeaderText="TOP" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="YesMMSCount" HeaderText="TOP" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#7FFF00" Font-Bold="false" ForeColor="black" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>         
		<!-- ****************************************************************************************** -->
		<!-- ** SPD List                                                                                -->
		<!-- ****************************************************************************************** -->
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 424px; position: absolute; top: 192px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="Red" BorderStyle="Solid" BorderWidth="5px" >
            <Columns>
                <asp:BoundField DataField="MMSDESC" HeaderText=""  />          
                <asp:BoundField DataField="Puller" HeaderText=""  />          
                <asp:BoundField DataField="Status" HeaderText=""  />          
                
                <asp:HyperLinkField DataNavigateUrlFields="RDNOUrl" DataNavigateUrlFormatString="{0}" 
                DataTextField="RDNOM" HeaderText="RDNo" Target="_blank"  >
                </asp:HyperLinkField>


                <asp:BoundField DataField="SliderCode" HeaderText=""  /> 
                <asp:BoundField DataField="Spec" HeaderText=""  />          
                <asp:BoundField DataField="SUPPLIER" HeaderText=""  />          

                <asp:HyperLinkField DataNavigateUrlFields="OPContactL" DataNavigateUrlFormatString="{0}" 
                DataTextField="OPContactM" HeaderText="型" Target="_blank"  >
                </asp:HyperLinkField>
                
                <asp:HyperLinkField DataNavigateUrlFields="ContactL" DataNavigateUrlFormatString="{0}" 
                DataTextField="ContactM" HeaderText="連" Target="_blank"  >
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="SliderDetailL" DataNavigateUrlFormatString="{0}" 
                DataTextField="SliderDetailM" HeaderText="細" Target="_blank"  >
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="SurfaceL" DataNavigateUrlFormatString="{0}" 
                DataTextField="SurfaceM" HeaderText="表" Target="_blank"  >
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="ColorAppendL" DataNavigateUrlFormatString="{0}" 
                DataTextField="ColorAppendM" HeaderText="色" Target="_blank"  >
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="YKKGroupCopyL" DataNavigateUrlFormatString="{0}" 
                DataTextField="YKKGroupCopyM" HeaderText="Y" Target="_blank"  >
                </asp:HyperLinkField>

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
