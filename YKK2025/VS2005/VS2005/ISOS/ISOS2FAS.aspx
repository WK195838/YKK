<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ISOS2FAS.aspx.vb" Inherits="ISOS2FAS" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>ISOS2FAS</title>
	    <script language="javascript" src="ForProject.js"></script>

</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  篩選欄位                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Label ID="Label1" runat="server" Style="z-index: 103; left: 8px; position: absolute; top: 8px"  Font-Size="10pt" >BUYER</asp:Label>
        <asp:Label ID="Label2" runat="server" Style="z-index: 103; left: 104px; position: absolute; top: 8px"  Font-Size="10pt" >ITEM/COLOR</asp:Label>
        <asp:Label ID="Label3" runat="server" Style="z-index: 103; left: 264px; position: absolute; top: 8px"  Font-Size="10pt" >KEEP CODE</asp:Label>
        <asp:Label ID="Label4" runat="server" Style="z-index: 103; left: 368px; position: absolute; top: 8px"  Font-Size="10pt" >CUSTOMER PULLER COLOR</asp:Label>
        &nbsp;<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ --><!-- ++  Action (CheckBox)                                                                   ++ --><!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->

        <asp:TextBox ID="DBUYER" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="94px" Style="z-index: 103; left: 8px; position: absolute; top: 24px"></asp:TextBox>

        <asp:TextBox ID="DItem" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="74px" Style="z-index: 103; left: 104px; position: absolute; top: 24px"></asp:TextBox>

        <asp:TextBox ID="DColor" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="74px" Style="z-index: 103; left: 184px; position: absolute; top: 24px"></asp:TextBox>

        <asp:TextBox ID="DKeepCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="94px" Style="z-index: 103; left: 264px; position: absolute; top: 24px"></asp:TextBox>

        <asp:TextBox ID="DPuller" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="144px" Style="z-index: 103; left: 368px; position: absolute; top: 24px"></asp:TextBox>
    
        <asp:TextBox ID="DItemName1" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="200px" Style="z-index: 103; left: 104px; position: absolute; top: 48px"></asp:TextBox>
    
        <asp:TextBox ID="DItemName2" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="200px" Style="z-index: 103; left: 312px; position: absolute; top: 48px"></asp:TextBox>
   
        <asp:TextBox ID="DItemName3" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="150px" Style="z-index: 103; left: 520px; position: absolute; top: 48px"></asp:TextBox>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  按鈕                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    	<asp:button id="BInq" runat="server" Height="56px" Width="80px" Style="z-index: 103; left: 680px; position: absolute; top: 8px" ForeColor="Black" BackColor="Silver" Text="Go(單)" Font-Size="Larger"></asp:button>
    	<asp:button id="BManyInq" runat="server" Height="56px" Width="80px" Style="z-index: 103; left: 760px; position: absolute; top: 8px" ForeColor="Black" BackColor="Silver" Text="Go(多)" Font-Size="Larger"></asp:button>
    	<asp:button id="BInput" runat="server" Height="32px" Width="136px" Style="z-index: 103; left: 536px; position: absolute; top: 8px" ForeColor="Black" BackColor="Silver" Text="多筆輸入"></asp:button>
        <asp:ImageButton ID="BExcel1" runat="server" Height="24px" ImageUrl="Images\msexcel.gif" Style="z-index: 103; left: 8px; position: absolute; top: 72px" Width="24px" />
        <asp:ImageButton ID="BExcel2" runat="server" Height="24px" ImageUrl="Images\msexcel.gif" Style="z-index: 103; left: 448px; position: absolute; top: 72px" Width="24px" />

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Width="430px" Style="z-index: 103; left: 8px; position: absolute; top: 100px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="C_CODE" HeaderText="Item"  />
                <asp:BoundField DataField="ITEMNAME" HeaderText="Item Name"  />
                <asp:BoundField DataField="Y_COLOR" HeaderText="Color"  />
                <asp:BoundField DataField="C_COLOR" HeaderText="C.PullerColor"  />
                <asp:BoundField DataField="C_ShortenLT" HeaderText="Keep"  />
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 448px; position: absolute; top: 100px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="GR_02" HeaderText="Keep"  />
                <asp:BoundField DataField="GR_03" HeaderText="Item"  />
                <asp:BoundField DataField="GR_04" HeaderText="Item Name"  />
                <asp:BoundField DataField="GR_05" HeaderText=""  />
                <asp:BoundField DataField="GR_06" HeaderText=""  />
                <asp:BoundField DataField="GR_07" HeaderText="Color"  />
                <asp:BoundField DataField="GR_08" HeaderText=""  />
                <asp:BoundField DataField="MinimumStock" HeaderText="Minimum"  />
                <asp:BoundField DataField="N_ScheProd" HeaderText="Sche Prod"  />
                <asp:BoundField DataField="N_OnProd" HeaderText="On Prod"  />
                <asp:BoundField DataField="N_FreeInv" HeaderText="Free Inv"  />
                <asp:BoundField DataField="N_KeepInv" HeaderText="Keep Inv"  />
                <asp:BoundField DataField="N_Total" HeaderText="Toal"  />
                <asp:BoundField DataField="DataTime" HeaderText="Time"  />
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
        <asp:TextBox ID="DISOS2FASFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
            
    </div>
    </form>
</body>
</html>
