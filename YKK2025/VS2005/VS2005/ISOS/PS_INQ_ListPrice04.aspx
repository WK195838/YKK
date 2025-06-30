<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PS_INQ_ListPrice04.aspx.vb" Inherits="PS_INQ_ListPrice04" %>

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
        <asp:TextBox ID="DINQLISTPRICEFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="HBuyerCode" runat="server" Height="16px" Style="z-index: 318; left: 8px;
            position: absolute; top: 192px;font-size:10px;background:transparent" Width="600px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

		<asp:hyperlink id="Hyperlink1" style="Z-INDEX: 102; LEFT: 664px; POSITION: absolute; TOP: 16px" runat="server"
					Width="80px" Height="16px" Font-Size="12pt" NavigateUrl="iMages/SEARCHREMARK.pdf" Target="_blank">Search說明</asp:hyperlink>
    
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  篩選欄位                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Label ID="Label3" runat="server" Style="z-index: 103; left: 10px; position: absolute; top: 10px"  >BUYER：</asp:Label>
        <asp:Label ID="Label1" runat="server" Style="z-index: 103; left: 170px; position: absolute; top: 8px"  Font-Size="10pt" >SEARCH-F</asp:Label>
        <asp:Label ID="Label2" runat="server" Style="z-index: 103; left: 170px; position: absolute; top: 28px"  Font-Size="10pt" >SEARCH-1</asp:Label>
        <asp:Label ID="Label4" runat="server" Style="z-index: 103; left: 170px; position: absolute; top: 48px"  Font-Size="10pt" >SEARCH-2</asp:Label>

        <asp:Label ID="Label5" runat="server" Style="z-index: 103; left: 32px; position: absolute; top: 152px"  Font-Size="10pt" Width="624px" >-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------</asp:Label>

        <asp:Label ID="Label6" runat="server" Style="z-index: 103; left: 1080px; position: absolute; top: 184px"  Font-Size="10pt" >L888(USD)(Exchange Rate=31.5)</asp:Label>

        <asp:TextBox ID="DBUYER" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="88px" Style="z-index: 103; left: 74px; position: absolute; top: 8px"></asp:TextBox>
                
        <asp:TextBox ID="DFKEY" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="248px" Style="z-index: 103; left: 234px; position: absolute; top: 8px"  Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="DKEY1" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="248px" Style="z-index: 103; left: 234px; position: absolute; top: 28px"  Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="DKEY2" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="248px" Style="z-index: 103; left: 234px; position: absolute; top: 48px"  Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="DZKEY" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="200px" Style="z-index: 103; left: 24px; position: absolute; top: 172px"  Font-Size="10pt"></asp:TextBox>
        <asp:TextBox ID="DPKEY" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="200px" Style="z-index: 103; left: 334px; position: absolute; top: 172px"  Font-Size="10pt"></asp:TextBox>

    <asp:Label ID="Label7" runat="server" Style="z-index: 103; left: 664px; position: absolute; top: 48px" ForeColor="Red" Font-Size="Smaller">Unit: USD$/per 100pcs</asp:Label>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  按鈕                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    	<asp:button id="BInq" runat="server" Height="56px" Width="64px" Style="z-index: 103; left: 488px; position: absolute; top: 8px" ForeColor="Black" BackColor="Silver" Text="Go" Font-Size="Larger"></asp:button>
        <asp:ImageButton ID="BExcel" runat="server" Height="24px" ImageUrl="Images\msexcel.gif" Style="z-index: 103; left: 16px; position: absolute; top: 140px" Width="24px" />
    
    	<asp:button id="BSPInq" runat="server" Height="56px" Width="96px" Style="z-index: 103; left: 560px; position: absolute; top: 8px" ForeColor="Black" BackColor="Silver" Text="特別篩選(B)" ></asp:button>

    	<asp:button id="BGetItem" runat="server" Height="40px" Width="528px" Style="z-index: 103; left: 24px; position: absolute; top: 112px" ForeColor="Black" BackColor="Silver" Text="搜尋ITEM" ></asp:button>
    	<asp:button id="BReset" runat="server" Height="40px" Width="88px" Style="z-index: 103; left: 560px; position: absolute; top: 112px" ForeColor="Black" BackColor="Silver" Text="RESET" ></asp:button>


    	<asp:button id="BZKEY" runat="server" Height="28px" Width="48px" Style="z-index: 103; left: 232px; position: absolute; top: 168px" ForeColor="Black" BackColor="Silver" Text="Go" ></asp:button>
    	<asp:button id="BPKEY" runat="server" Height="28px" Width="48px" Style="z-index: 103; left: 544px; position: absolute; top: 168px" ForeColor="Black" BackColor="Silver" Text="Go" ></asp:button>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  ITEM KEY 欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:TextBox ID="LItemKey1" runat="server" BackColor="Black" BorderStyle="Groove"
                ForeColor="White" Height="14px"  Width="250px" Style="z-index: 103; left: 24px; position: absolute; top: 72px" Font-Size="10pt">Z#</asp:TextBox>
        <asp:TextBox ID="DItemKey1" runat="server" BackColor="#80FFFF" BorderStyle="Groove"
                ForeColor="Red" Height="14px"  Width="250px" Style="z-index: 103; left: 24px; position: absolute; top: 88px" Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="LItemKey2" runat="server" BackColor="Black" BorderStyle="Groove"
                ForeColor="White" Height="14px"  Width="150px" Style="z-index: 103; left: 288px; position: absolute; top: 72px" Font-Size="10pt">P#</asp:TextBox>
        <asp:TextBox ID="DItemKey2" runat="server" BackColor="#80FFFF" BorderStyle="Groove"
                ForeColor="Red" Height="14px"  Width="150px" Style="z-index: 103; left: 288px; position: absolute; top: 88px" Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="LItemKey3" runat="server" BackColor="Black" BorderStyle="Groove"
                ForeColor="White" Height="14px"  Width="100px" Style="z-index: 103; left: 448px; position: absolute; top: 72px" Font-Size="10pt">Other</asp:TextBox>
        <asp:TextBox ID="DItemKey3" runat="server" BackColor="#80FFFF" BorderStyle="Groove"
                ForeColor="red" Height="14px"  Width="100px" Style="z-index: 103; left: 448px; position: absolute; top: 88px" Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="LZLength" runat="server" BackColor="Black" BorderStyle="Groove"
                ForeColor="White" Height="14px"  Width="100px" Style="z-index: 103; left: 560px; position: absolute; top: 72px" Font-Size="10pt">Z# Length</asp:TextBox>
        <asp:TextBox ID="DZLength" runat="server" BackColor="White" BorderStyle="Groove"
                ForeColor="red" Height="14px"  Width="100px" Style="z-index: 103; left: 560px; position: absolute; top: 90px" Font-Size="10pt">0</asp:TextBox>

        <asp:TextBox ID="LZBase" runat="server" BackColor="Black" BorderStyle="Groove"
                ForeColor="White" Height="14px"  Width="100px" Style="z-index: 103; left: 664px; position: absolute; top: 72px" Font-Size="10pt">Z# Base</asp:TextBox>
        <asp:TextBox ID="DZBase" runat="server" BackColor="White" BorderStyle="Groove"
                ForeColor="red" Height="14px"  Width="100px" Style="z-index: 103; left: 664px; position: absolute; top: 90px" Font-Size="10pt">0</asp:TextBox>

        <asp:TextBox ID="LZInchBase" runat="server" BackColor="Black" BorderStyle="Groove"
                ForeColor="White" Height="14px"  Width="100px" Style="z-index: 103; left: 776px; position: absolute; top: 72px" Font-Size="10pt">Z# add 1"</asp:TextBox>
        <asp:TextBox ID="DZInchBase" runat="server" BackColor="White" BorderStyle="Groove"
                ForeColor="red" Height="14px"  Width="100px" Style="z-index: 103; left: 776px; position: absolute; top: 90px" Font-Size="10pt">0</asp:TextBox>

        <asp:TextBox ID="LPBase" runat="server" BackColor="Black" BorderStyle="Groove"
                ForeColor="White" Height="14px"  Width="100px" Style="z-index: 103; left: 664px; position: absolute; top: 110px" Font-Size="10pt">P# Base</asp:TextBox>
        <asp:TextBox ID="DPBase" runat="server" BackColor="White" BorderStyle="Groove"
                ForeColor="red" Height="14px"  Width="100px" Style="z-index: 103; left: 664px; position: absolute; top: 128px" Font-Size="10pt">0</asp:TextBox>

        <asp:TextBox ID="LFBase" runat="server" BackColor="Black" BorderStyle="Groove"
                ForeColor="White" Height="14px"  Width="100px" Style="z-index: 103; left: 664px; position: absolute; top: 146px" Font-Size="10pt">F# Base</asp:TextBox>
        <asp:TextBox ID="LFBase1" runat="server" BackColor="Black" BorderStyle="Groove"
                ForeColor="White" Height="14px"  Width="500px" Style="z-index: 103; left: 770px; position: absolute; top: 146px" Font-Size="10pt"></asp:TextBox>
        <asp:TextBox ID="DFBase" runat="server" BackColor="White" BorderStyle="Groove"
                ForeColor="red" Height="14px"  Width="606px" Style="z-index: 103; left: 664px; position: absolute; top: 164px" Font-Size="10pt">0</asp:TextBox>
                
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView-1                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!--            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />  -->

        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Width="300px" Style="z-index: 103; left: 24px; position: absolute; top: 200px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
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
                
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}" 
                DataTextField="Mark" HeaderText="Select" Target="_blank"  >
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="FormURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="FormMark" HeaderText="Select" Target="_blank">
                </asp:HyperLinkField>

                <asp:CommandField ShowSelectButton="True" SelectText="Ｓ" />

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView-2                                                               ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!--            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />  -->

        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" Width="300px" Style="z-index: 103; left: 350px; position: absolute; top: 200px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
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
                
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}" 
                DataTextField="Mark" HeaderText="Select" Target="_blank"  >
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="FormURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="FormMark" HeaderText="Select" Target="_blank">
                </asp:HyperLinkField>

                <asp:CommandField ShowSelectButton="True" SelectText="Ｓ" />

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>        
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView-3                                                              ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!--            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />  -->

        <asp:GridView ID="GridView3" runat="server" AutoGenerateColumns="False" Width="400px" Style="z-index: 103; left: 664px; position: absolute; top: 200px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
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
                
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}" 
                DataTextField="Mark" HeaderText="Select" Target="_blank"  >
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="FormURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="FormMark" HeaderText="Select" Target="_blank">
                </asp:HyperLinkField>

                <asp:CommandField ShowSelectButton="True" SelectText="Ｓ" />

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>        
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  GridView-5                                                               ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!--            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />  -->

        <asp:GridView ID="GridView5" runat="server" AutoGenerateColumns="False" Width="700px" Style="z-index: 103; left: 1080px; position: absolute; top: 200px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="PriceInf" HeaderText="A"  />
                <asp:BoundField DataField="BaseLength" HeaderText="A"  />
                <asp:BoundField DataField="ListPrice" HeaderText="B"  />
                <asp:BoundField DataField="InchPrice" HeaderText="C"  />
                <asp:BoundField DataField="ItemInf" HeaderText="ItemInf"  />
                <asp:BoundField DataField="Code" HeaderText=""  />
                <asp:BoundField DataField="Name1" HeaderText=""  />
                <asp:BoundField DataField="Name2" HeaderText=""  />
                <asp:BoundField DataField="Name3" HeaderText=""  />
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>

    
    </div>
    </form>
</body>
</html>
