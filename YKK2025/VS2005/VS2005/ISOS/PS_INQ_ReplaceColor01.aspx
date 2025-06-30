<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PS_INQ_ReplaceColor01.aspx.vb" Inherits="PS_INQ_ReplaceColor01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >


<head id="Head1" runat="server">
    <title>Replace Item Inf.</title>
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
    
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  篩選欄位                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Label ID="Label3" runat="server" Style="z-index: 103; left: 10px; position: absolute; top: 8px">BUYER：</asp:Label>
        <asp:Label ID="Label1" runat="server" Style="z-index: 103; left: 155px; position: absolute; top: 8px">SEARCH KEY：</asp:Label>

        <asp:TextBox ID="DBUYER" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="74px" Style="z-index: 103; left: 64px; position: absolute; top: 8px"></asp:TextBox>

        <asp:TextBox ID="DKEY" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="456px" Style="z-index: 103; left: 272px; position: absolute; top: 8px"></asp:TextBox>
        
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Action (CheckBox)                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  按鈕                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    	<asp:button id="BInq" runat="server" Height="24px" Width="40px" Style="z-index: 103; left: 744px; position: absolute; top: 8px" ForeColor="Black" BackColor="Silver" Text="Go"></asp:button>
        <asp:ImageButton ID="BExcel" runat="server" Height="24px" ImageUrl="Images\msexcel.gif" Style="z-index: 103; left: 10px; position: absolute; top: 32px" Width="24px" />

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 24px; position: absolute; top: 60px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                
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

    </div>
    </form>
</body>
</html>
