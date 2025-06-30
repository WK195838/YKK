<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PS_INQ_Color01.aspx.vb" Inherits="PS_INQ_Color01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >


<head id="Head1" runat="server">
    <title>Color Inf.</title>
	    <script language="javascript" src="ForProject.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/ITEMSOURCE.png" style="z-index: 1; left: 6px; position: absolute;top: 0px; width: 282px; height: 80px;"/>

		<asp:hyperlink id="LSearch" style="Z-INDEX: 102; LEFT: 536px; POSITION: absolute; TOP: 8px" runat="server"
					Width="80px" Height="16px" Font-Size="12pt" NavigateUrl="iMages/SEARCHREMARK.pdf" Target="_blank">Search說明</asp:hyperlink>

		<asp:hyperlink id="LTeethPage" style="Z-INDEX: 102; LEFT: 630px; POSITION: absolute; TOP: 80px" runat="server"
					Width="136px" Height="16px" Font-Size="12pt" NavigateUrl="http://10.245.0.3/ISOS/Document/000013/TEETH/Metallic Vislon Teeth.xls" Target="_blank">Metallic Vislon Teeth</asp:hyperlink>

		<asp:hyperlink id="LTapeDesc" style="Z-INDEX: 102; LEFT: 40px; POSITION: absolute; TOP: 88px" runat="server"
					Width="88px" Height="16px" Font-Size="12pt" NavigateUrl="http://10.245.0.3/ISOS/iMages/TapeDesc_TW1741.png" Target="_blank" ForeColor="Red">*注意事項*</asp:hyperlink>
    
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Action (CheckBox)                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:CheckBox ID="AtDevlop" runat="server" style="z-index: 174; left: 36px; position: absolute; top: 12px" Font-Size="12pt" Text="" Width="20px" AutoPostBack="True" />
        <asp:CheckBox ID="AtBuyer" runat="server" style="z-index: 174; left: 114px; position: absolute; top: 12px" Font-Size="12pt" Text="" Width="20px" AutoPostBack="True"  />
        <asp:CheckBox ID="AtFAS" runat="server" style="z-index: 174; left: 208px; position: absolute; top: 12px" Font-Size="12pt" Text="" Width="20px" AutoPostBack="True" />

        <asp:CheckBox ID="AtReplace" runat="server" style="z-index: 174; left: 36px; position: absolute; top: 48px" Font-Size="12pt" Text="" Width="20px" AutoPostBack="True" />
        <asp:CheckBox ID="AtSales" runat="server" style="z-index: 174; left: 114px; position: absolute; top: 48px" Font-Size="12pt" Text="" Width="20px" AutoPostBack="True" />
    
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  篩選欄位                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Label ID="Label1" runat="server" Style="z-index: 103; left: 296px; position: absolute; top: 36px"  Font-Size="10pt" >SEARCH-F</asp:Label>
        <asp:Label ID="Label2" runat="server" Style="z-index: 103; left: 296px; position: absolute; top: 56px"  Font-Size="10pt" >SEARCH-1</asp:Label>
        <asp:Label ID="Label3" runat="server" Style="z-index: 103; left: 296px; position: absolute; top: 76px"  Font-Size="10pt" >SEARCH-2</asp:Label>
        <asp:Label ID="Label4" runat="server" Style="z-index: 103; left: 230px; position: absolute; top: 14px"  Font-Size="10pt" BackColor="Black" ForeColor="White" 
        Width="40px" Height="16px">_無_</asp:Label>

        <asp:TextBox ID="DBUYER" runat="server" BackColor="#80FF80" BorderStyle="Groove" 
                ForeColor="Blue" Height="20px"  Width="120px" Style="z-index: 103; left: 296px; position: absolute; top: 8px"></asp:TextBox>

        <asp:TextBox ID="DFKEY" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="20px"  Width="248px" Style="z-index: 103; left: 360px; position: absolute; top: 34px"  Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="DKEY1" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="20px"  Width="248px" Style="z-index: 103; left: 360px; position: absolute; top: 54px"  Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="DKEY2" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="20px"  Width="248px" Style="z-index: 103; left: 360px; position: absolute; top: 74px"  Font-Size="10pt"></asp:TextBox>

        <asp:DropDownList ID="DColorType" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 200px; position: absolute; top: 46px" Width="80px" Height="24px">
            <asp:ListItem Selected="True" Value="ALL">ALL</asp:ListItem>
            <asp:ListItem Value="TAPE">TAPE</asp:ListItem>
            <asp:ListItem Value="PULLER">PULLER</asp:ListItem>
        </asp:DropDownList>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  按鈕                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    	<asp:button id="BInq" runat="server" Height="64px" Width="64px" Style="z-index: 103; left: 624px; position: absolute; top: 8px" ForeColor="Black" BackColor="Silver" Text="Go" Font-Size="Larger"></asp:button>
        <asp:ImageButton ID="BExcel" runat="server" Height="24px" ImageUrl="Images\msexcel.gif" Style="z-index: 103; left: 6px; position: absolute; top: 80px" Width="24px" />
    
    	<asp:button id="BSPInq" runat="server" Height="64px" Width="80px" Style="z-index: 103; left: 696px; position: absolute; top: 8px" ForeColor="Black" BackColor="Silver" Text="特別篩選(B)"></asp:button>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- System Parameter  -->
        <asp:TextBox ID="HBuyerCode" runat="server" Height="16px" Style="z-index: 318; left: 8px;
            position: absolute; top: 192px;font-size:10px;background:transparent" Width="600px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DINQBUYERCOLORFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 6px; position: absolute; top: 112px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="A1" HeaderText=""  />
                <asp:BoundField DataField="B1" HeaderText=""  />
                <asp:BoundField DataField="C1" HeaderText=""  />
                <asp:BoundField DataField="D1" HeaderText=""  />
                <asp:BoundField DataField="E1" HeaderText=""  />
                <asp:BoundField DataField="F1" HeaderText=""  />
                <asp:BoundField DataField="G1" HeaderText=""  />
                <asp:BoundField DataField="H1" HeaderText=""  />
                <asp:BoundField DataField="I1" HeaderText=""  />
                <asp:BoundField DataField="J1" HeaderText=""  />
                <asp:BoundField DataField="K1" HeaderText=""  />
                <asp:BoundField DataField="L1" HeaderText=""  />
                <asp:BoundField DataField="M1" HeaderText=""  />
                <asp:BoundField DataField="N1" HeaderText=""  />
                <asp:BoundField DataField="O1" HeaderText=""  />
                <asp:BoundField DataField="P1" HeaderText=""  />
                <asp:BoundField DataField="Q1" HeaderText=""  />
                <asp:BoundField DataField="R1" HeaderText=""  />
                <asp:BoundField DataField="S1" HeaderText=""  />
                <asp:BoundField DataField="T1" HeaderText=""  />
                <asp:BoundField DataField="U1" HeaderText=""  />
                <asp:BoundField DataField="V1" HeaderText=""  />
                <asp:BoundField DataField="W1" HeaderText=""  />
                <asp:BoundField DataField="X1" HeaderText=""  />
                <asp:BoundField DataField="Y1" HeaderText=""  />
                <asp:BoundField DataField="Z1" HeaderText=""  />
                
                <asp:HyperLinkField DataNavigateUrlFields="DTMWURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="DTMW" HeaderText="" Target="_blank">
                </asp:HyperLinkField>
                
                <asp:HyperLinkField DataNavigateUrlFields="SLDCURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="SLDC" HeaderText="" Target="_blank">
                </asp:HyperLinkField>
                
                <asp:HyperLinkField DataNavigateUrlFields="IRWURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="IRW" HeaderText="" Target="_blank">
                </asp:HyperLinkField>
                
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
