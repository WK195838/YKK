<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PS_INQ_UpdateMark.aspx.vb" Inherits="PS_INQ_UpdateMark" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Update Inf.</title>
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
        &nbsp;
    
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  篩選欄位                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Label ID="Label3" runat="server" Style="z-index: 103; left: 10px; position: absolute; top: 8px">BUYER：</asp:Label>
        &nbsp;
        <asp:Label ID="Label2" runat="server" Style="z-index: 103; left: 192px; position: absolute; top: 8px"  Font-Size="10pt" >SEARCH-1</asp:Label>
        <asp:Label ID="Label4" runat="server" Style="z-index: 103; left: 192px; position: absolute; top: 32px"  Font-Size="10pt" >SEARCH-2</asp:Label>

        <asp:TextBox ID="DBUYER" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="112px" Style="z-index: 103; left: 64px; position: absolute; top: 8px"></asp:TextBox>
        &nbsp;

        <asp:TextBox ID="DKEY1" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="248px" Style="z-index: 103; left: 256px; position: absolute; top: 8px"  Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="DKEY2" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="248px" Style="z-index: 103; left: 256px; position: absolute; top: 32px"  Font-Size="10pt"></asp:TextBox>

        <asp:DropDownList ID="DDataType" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 16px; position: absolute; top: 32px" Width="168px" Height="24px">
            <asp:ListItem Selected="True" Value="BUYERITEM">ITEM</asp:ListItem>
            <asp:ListItem Value="BUYERCOLOR-TAPE">COLOR(TAPE)</asp:ListItem>
            <asp:ListItem Value="BUYERCOLOR-PULLER">COLOR(PULLER)</asp:ListItem>
            <asp:ListItem Value="BUYERPRICE">PRICE</asp:ListItem>
        </asp:DropDownList>
        
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Action (CheckBox)                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  按鈕                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    	<asp:button id="BInq" runat="server" Height="40px" Width="72px" Style="z-index: 103; left: 512px; position: absolute; top: 8px" ForeColor="Black" BackColor="Silver" Text="Go" Font-Size="Larger"></asp:button>
        <asp:ImageButton ID="BExcel" runat="server" Height="24px" ImageUrl="Images\msexcel.gif" Style="z-index: 103; left: 600px; position: absolute; top: 24px" Width="24px" />

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 12px; position: absolute; top: 60px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="CAT" HeaderText=""  />
                <asp:BoundField DataField="ID" HeaderText="更新時間"  />
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
                <asp:BoundField DataField="Active" HeaderText=""  />

                <asp:BoundField DataField="WA1" HeaderText="WA"  />
                <asp:BoundField DataField="WB1" HeaderText="WB"  />
                <asp:BoundField DataField="WC1" HeaderText="WC"  />
                <asp:BoundField DataField="WD1" HeaderText="WD"  />
                <asp:BoundField DataField="WE1" HeaderText="WE"  />
                <asp:BoundField DataField="WF1" HeaderText="WF"  />
                <asp:BoundField DataField="WG1" HeaderText="WG"  />
                <asp:BoundField DataField="WH1" HeaderText="WH"  />
                <asp:BoundField DataField="WI1" HeaderText="WI"  />
                <asp:BoundField DataField="WJ1" HeaderText="WJ"  />
                <asp:BoundField DataField="WK1" HeaderText="WK"  />
                <asp:BoundField DataField="WL1" HeaderText="WL"  />
                <asp:BoundField DataField="WM1" HeaderText="WM"  />
                <asp:BoundField DataField="WN1" HeaderText="WN"  />
                <asp:BoundField DataField="WO1" HeaderText="WO"  />
                <asp:BoundField DataField="WP1" HeaderText="WP"  />
                <asp:BoundField DataField="WQ1" HeaderText="WQ"  />
                <asp:BoundField DataField="WR1" HeaderText="WR"  />
                <asp:BoundField DataField="WS1" HeaderText="WS"  />
                <asp:BoundField DataField="WT1" HeaderText="WT"  />
                <asp:BoundField DataField="WU1" HeaderText="WU"  />
                <asp:BoundField DataField="WV1" HeaderText="WV"  />
                <asp:BoundField DataField="WW1" HeaderText="WW"  />
                <asp:BoundField DataField="WX1" HeaderText="WX"  />
                <asp:BoundField DataField="WY1" HeaderText="WY"  />
                <asp:BoundField DataField="WZ1" HeaderText="WZ"  />


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

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- FileData  -->

    </div>
    </form>

</body>
</html>
