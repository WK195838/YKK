<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PS_INQ_ListPrice01.aspx.vb" Inherits="PS_INQ_ListPrice01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >


<head id="Head1" runat="server">
    <title>List Price Inf.</title>
	    <script language="javascript" src="ForProject.js"></script>

</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- System Parameter  -->
        <asp:TextBox ID="HBuyerCode" runat="server" Height="16px" Style="z-index: 318; left: 8px;
            position: absolute; top: 192px;font-size:10px;background:transparent" Width="600px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

		<asp:hyperlink id="LSearch" style="Z-INDEX: 102; LEFT: 712px; POSITION: absolute; TOP: 8px" runat="server"
					Width="80px" Height="16px" Font-Size="12pt" NavigateUrl="iMages/SEARCHREMARK.pdf" Target="_blank">Search說明</asp:hyperlink>
    
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  篩選欄位                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Label ID="Label3" runat="server" Style="z-index: 103; left: 10px; position: absolute; top: 8px">BUYER：</asp:Label>
        <asp:Label ID="Label1" runat="server" Style="z-index: 103; left: 224px; position: absolute; top: 8px"  Font-Size="10pt" >SEARCH-F</asp:Label>
        <asp:Label ID="Label2" runat="server" Style="z-index: 103; left: 224px; position: absolute; top: 28px"  Font-Size="10pt" >SEARCH-1</asp:Label>
        <asp:Label ID="Label4" runat="server" Style="z-index: 103; left: 224px; position: absolute; top: 48px"  Font-Size="10pt" >SEARCH-2</asp:Label>
    
    <asp:Label ID="Label5" runat="server" Style="z-index: 103; left: 712px; position: absolute; top: 40px" ForeColor="Red" Font-Size="Smaller">Unit: USD$/per 100pcs</asp:Label>

        <asp:TextBox ID="DBUYER" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="120px" Style="z-index: 103; left: 64px; position: absolute; top: 8px"></asp:TextBox>
                
        <asp:TextBox ID="DFKEY" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="248px" Style="z-index: 103; left: 288px; position: absolute; top: 8px"  Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="DKEY1" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="248px" Style="z-index: 103; left: 288px; position: absolute; top: 28px"  Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="DKEY2" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="248px" Style="z-index: 103; left: 288px; position: absolute; top: 48px"  Font-Size="10pt"></asp:TextBox>

        <asp:DropDownList ID="DSeason" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 64px; position: absolute; top: 32px" Width="80px" Height="24px">
            <asp:ListItem Selected="True" Value="BUYERPRICE">當季</asp:ListItem>
            <asp:ListItem Value="BUYERPRICE-B">前季</asp:ListItem>
            <asp:ListItem Value="999999">ALL</asp:ListItem>
        </asp:DropDownList>

        <asp:DropDownList ID="DItemType" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 64px; position: absolute; top: 32px" Width="64px" Height="24px">
            <asp:ListItem Selected="True" Value="APP">APP</asp:ListItem>
            <asp:ListItem Value="BAG">BAG</asp:ListItem>
        </asp:DropDownList>
        
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Action (CheckBox)                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  按鈕                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    	<asp:button id="BInq" runat="server" Height="48px" Width="64px" Style="z-index: 103; left: 552px; position: absolute; top: 8px" ForeColor="Black" BackColor="Silver" Text="Go" Font-Size="Larger"></asp:button>
    	<asp:button id="BSPInq" runat="server" Height="48px" Width="80px" Style="z-index: 103; left: 624px; position: absolute; top: 8px" ForeColor="Black" BackColor="Silver" Text="特別篩選"></asp:button>
        <asp:ImageButton ID="BExcel" runat="server" Height="24px" ImageUrl="Images\msexcel.gif" Style="z-index: 103; left: 10px; position: absolute; top: 32px" Width="24px" />

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 24px; position: absolute; top: 76px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}" 
                DataTextField="Mark" HeaderText="Select" Target="_blank">
                </asp:HyperLinkField>
                
                <asp:BoundField DataField="A1" HeaderText="A"  />
                <asp:BoundField DataField="B1" HeaderText="B"  />
                <asp:BoundField DataField="C1" HeaderText="C"  />
                <asp:BoundField DataField="D1" HeaderText="D"  />
                <asp:BoundField DataField="E1" HeaderText="E"  />
                <asp:BoundField DataField="F1" HeaderText="F"  />
                <asp:BoundField DataField="G1" HeaderText="G"  />
                <asp:BoundField DataField="H1" HeaderText="H"  />
                <asp:BoundField DataField="I1" HeaderText="I"  />
                <asp:BoundField DataField="J1" HeaderText="J"  />
                <asp:BoundField DataField="K1" HeaderText="K"  />
                <asp:BoundField DataField="L1" HeaderText="L"  />
                <asp:BoundField DataField="M1" HeaderText="M"  />
                <asp:BoundField DataField="N1" HeaderText="N"  />
                <asp:BoundField DataField="O1" HeaderText="O"  />
                <asp:BoundField DataField="P1" HeaderText="P"  />
                <asp:BoundField DataField="Q1" HeaderText="Q"  />
                <asp:BoundField DataField="R1" HeaderText="R"  />
                <asp:BoundField DataField="S1" HeaderText="S"  />
                <asp:BoundField DataField="T1" HeaderText="T"  />
                <asp:BoundField DataField="U1" HeaderText="U"  />
                <asp:BoundField DataField="V1" HeaderText="V"  />
                <asp:BoundField DataField="W1" HeaderText="W"  />
                <asp:BoundField DataField="X1" HeaderText="X"  />
                <asp:BoundField DataField="Y1" HeaderText="Y"  />
                <asp:BoundField DataField="Z1" HeaderText="Z"  />

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- FileData  -->
        <asp:TextBox ID="DINQLISTPRICEFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

    </div>
    </form>
</body>

</html>
