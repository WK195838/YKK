<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PS_INQ_COMBI01.aspx.vb" Inherits="PS_INQ_COMBI01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">

    <title>FAS COMBI INQ</title>
	    <script language="javascript" src="ForProject.js"></script>
	    
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  篩選欄位                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Label ID="Label2" runat="server" Style="z-index: 103; left: 8px; position: absolute; top: 32px"  Font-Size="10pt" >SEARCH-1</asp:Label>
        <asp:Label ID="Label3" runat="server" Style="z-index: 103; left: 8px; position: absolute; top: 56px"  Font-Size="10pt" >SEARCH-2</asp:Label>

        <asp:TextBox ID="DBUYER" runat="server" BackColor="#80FF80" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="120px" Style="z-index: 103; left: 72px; position: absolute; top: 8px"></asp:TextBox>

        <asp:TextBox ID="DKEY1" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="248px" Style="z-index: 103; left: 72px; position: absolute; top: 32px"  Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="DKEY2" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="248px" Style="z-index: 103; left: 72px; position: absolute; top: 56px"  Font-Size="10pt"></asp:TextBox>

        <asp:DropDownList ID="DColorType" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 8px; position: absolute; top: 88px" Width="168px" Height="24px">
            <asp:ListItem Selected="True" Value="TAPE">TAPE</asp:ListItem>
            <asp:ListItem Value="TAPETEETH">TAPE+TEETH</asp:ListItem>
            <asp:ListItem Value="ALL">ALL</asp:ListItem>
        </asp:DropDownList>
    

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Width="600px" Style="z-index: 103; left: 8px; position: absolute; top: 112px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:CommandField ShowSelectButton="True" SelectText="Ｓ" />

                <asp:BoundField DataField="CODE" HeaderText="Code"  />
                <asp:BoundField DataField="COLOR1" HeaderText="Tape"  />
                <asp:BoundField DataField="COLOR2" HeaderText="Teeth"  />
                <asp:BoundField DataField="COLOR3" HeaderText="Slder"  />
                <asp:BoundField DataField="SliderStatus" HeaderText="Other"  />
                <asp:BoundField DataField="YCODE" HeaderText="Item"  />
                <asp:BoundField DataField="YNAME" HeaderText="Item Name"  />
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  按鈕                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    	<asp:button id="BInq" runat="server" Height="40px" Width="56px" Style="z-index: 103; left: 328px; position: absolute; top: 32px" ForeColor="Black" BackColor="Silver" Text="Go" Font-Size="Larger"></asp:button>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- System Parameter  -->
        <asp:TextBox ID="HBuyerCode" runat="server" Height="16px" Style="z-index: 318; left: 8px;
            position: absolute; top: 192px;font-size:10px;background:transparent" Width="600px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 624px; position: absolute; top: 112px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Season" HeaderText="Season"  />
                <asp:BoundField DataField="Color1" HeaderText="Tape"  />
                <asp:BoundField DataField="Color2" HeaderText="Teeth / Slider"  />
                <asp:BoundField DataField="Green" HeaderText="Other"  />
                <asp:BoundField DataField="YColor" HeaderText="Color"  />
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>    
    
    </div>
    </form>
</body>
</html>
