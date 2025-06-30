<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfF_OrderProgress.aspx.vb" Inherits="InfF_OrderProgress" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>LS Order Progress</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  篩選欄位                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Label ID="Label1" runat="server" Style="z-index: 103; left: 11px; position: absolute; top: 8px">Buyer：</asp:Label>
        <asp:Label ID="Label2" runat="server" Style="z-index: 103; left: 821px; position: absolute; top: 10px">FCT No：</asp:Label>
        <asp:Label ID="Label3" runat="server" Style="z-index: 103; left: 150px; position: absolute; top: 8px">BLS No：</asp:Label>
        <asp:Label ID="Label4" runat="server" Style="z-index: 103; left: 335px; position: absolute; top: 8px">Order No：</asp:Label>
        
        <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="74px" Style="z-index: 103; left: 64px; position: absolute; top: 8px"></asp:TextBox>
        <asp:TextBox ID="DFCTNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="100px" Style="z-index: 103; left: 891px; position: absolute; top: 10px"></asp:TextBox>
        <asp:TextBox ID="DBLSNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="100px" Style="z-index: 103; left: 219px; position: absolute; top: 8px"></asp:TextBox>
        <asp:TextBox ID="DOrderNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="100px" Style="z-index: 103; left: 412px; position: absolute; top: 8px"></asp:TextBox>
        <asp:TextBox ID="DTime" runat="server" BackColor="White" BorderStyle="None" ForeColor="#C00000"
            Height="16px" Style="z-index: 103; left: 577px; position: absolute; top: 9px"
            Width="138px">實際Order資料時點</asp:TextBox>
        <asp:TextBox ID="TextBox1" runat="server" BackColor="White" BorderStyle="None" ForeColor="#C00000"
            Height="16px" Style="z-index: 103; left: 695px; position: absolute; top: 9px"
            Width="109px">當日 7:00</asp:TextBox>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  按鈕                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    	<asp:button id="BInq" runat="server" Height="24px" Width="40px" Style="z-index: 103; left: 528px; position: absolute; top: 8px" ForeColor="Black" BackColor="Silver" Text="Go"></asp:button>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  OrderProgress GridView                                                             ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="PGGridView" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 15px; position: absolute; top: 40px" CellPadding="4" Font-Size="9pt" GridLines="None" PageSize="20" ForeColor="#333333" >
            <Columns>
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="OrderNo" HeaderText="OrderNo" Target="_blank">
                </asp:HyperLinkField>
                <asp:BoundField DataField="OrderSubNo" HeaderText="Sub"  />
                <asp:BoundField DataField="Status" HeaderText="Progress"  />
                <asp:BoundField DataField="Keep" HeaderText="Keep"  />
                <asp:BoundField DataField="Item" HeaderText="Item"  />
                <asp:BoundField DataField="ItemName" HeaderText="ItemName"  />
                <asp:BoundField DataField="Color" HeaderText="Color"  />
                <asp:BoundField DataField="Qty" HeaderText="Qty"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="OrderDate" HeaderText="OrderDate"  />
                <asp:BoundField DataField="ReqDate" HeaderText="Req.Date"  />
                <asp:BoundField DataField="PlanDate" HeaderText="PlanDate"  />
                <asp:BoundField DataField="CustOrderNo" HeaderText="C.OrderNo"  />
            </Columns>
            <HeaderStyle BackColor="#990000" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <RowStyle BackColor="#FFFBD6" ForeColor="#333333" />
            <FooterStyle BackColor="#990000" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#FFCC66" ForeColor="#333333" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="Navy" />
            <AlternatingRowStyle BackColor="White" />
        </asp:GridView>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Local Stock Plan GridView                                                           ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="LSGridView" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 15px; position: absolute; top: 560px" CellPadding="4" Font-Size="9pt" AllowPaging="True" PageSize="20" ForeColor="#333333" GridLines="None">
            <Columns>
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="BULSNo" HeaderText="LSNo" Target="_blank">
                </asp:HyperLinkField>
                <asp:BoundField DataField="BULSSubNo" HeaderText="Sub"  />
                <asp:BoundField DataField="GR_03" HeaderText="Item"  />
                <asp:BoundField DataField="ItemName" HeaderText="ItemName"  />
                <asp:BoundField DataField="GR_07" HeaderText="Color"  />
                <asp:BoundField DataField="GR_02" HeaderText="Keep"  />

                <asp:BoundField DataField="FreeAlloc" HeaderText="FreeAlloc">  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="N_ScheProd" HeaderText="SCHE">  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="N_OnProd" HeaderText="ON">  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="N_FreeInv" HeaderText="FREE">  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="N_KeepInv" HeaderText="KEEP">  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                
                <asp:BoundField DataField="BLANK" HeaderText="|"  />
                <asp:BoundField DataField="FS_00" HeaderText="N0_F">  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>

                <asp:BoundField DataField="BLANK" HeaderText="|"  />
                <asp:BoundField DataField="SP_01" HeaderText="N1_SCHE">  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="OP_01" HeaderText="N1_ON">  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="FS_01" HeaderText="N1_F">  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="PS_01" HeaderText="N1_P">  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="IS_01" HeaderText="N1_I">  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>

                <asp:BoundField DataField="BLANK" HeaderText="|"  />
                <asp:BoundField DataField="SP_02" HeaderText="N2_SCHE">  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="OP_02" HeaderText="N2_ON">  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="FS_02" HeaderText="N2_F">  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="PS_02" HeaderText="N2_P">  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="IS_02" HeaderText="N2_I">  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>

            </Columns>
            <HeaderStyle BackColor="#990000" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <RowStyle BackColor="#FFFBD6" ForeColor="#333333" />
            <FooterStyle BackColor="#990000" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#FFCC66" ForeColor="#333333" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="Navy" />
            <AlternatingRowStyle BackColor="White" />
        </asp:GridView>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Forecast Plan GridView                                                              ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="FCTGridView" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 17px; position: absolute; top: 698px" CellPadding="4" Font-Size="9pt" AllowPaging="True" ForeColor="#333333" GridLines="None" PageSize="20">
            <Columns>
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}" DataTextField="FCTNo" HeaderText="FCTNo" Target="_blank">
                </asp:HyperLinkField>
                <asp:BoundField HeaderText="Sub" DataField="FCTSubNo" />
                <asp:BoundField HeaderText="PLM" DataField="C_Code" />
                <asp:BoundField HeaderText="Style" DataField="C_Style" />
                <asp:BoundField HeaderText="Article" DataField="C_A1" />
                <asp:BoundField HeaderText="Part" DataField="C_B1" />
                <asp:BoundField HeaderText="Color" DataField="C_Color" />
                <asp:BoundField HeaderText="Keep" DataField="C_ShortenLT" />
                <asp:BoundField HeaderText="Factory" DataField="C_C1" />

                <asp:BoundField HeaderText="Code" DataField="Y_ItemCode" />
                <asp:BoundField HeaderText="ItemName" DataField="Y_ItemName" />
                <asp:BoundField HeaderText="Code" DataField="Y_Color" />

                <asp:BoundField DataField="BLANK" HeaderText="|"  />
                <asp:BoundField HeaderText="N0_F" DataField="N_F" >  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>                
                <asp:BoundField HeaderText="N1_F" DataField="N1_F" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>  
                <asp:BoundField HeaderText="N2_F" DataField="N2_F" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField> 
                <asp:BoundField HeaderText="N3_F" DataField="N3_F" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField> 
                <asp:BoundField HeaderText="N4_F" DataField="N4_F" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField> 
            </Columns>
            <HeaderStyle BackColor="#990000" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <RowStyle BackColor="#FFFBD6" ForeColor="#333333" />
            <FooterStyle BackColor="#990000" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#FFCC66" ForeColor="#333333" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="Navy" />
            <AlternatingRowStyle BackColor="White" />
        </asp:GridView>
    
    </div>
    </form>
</body>
</html>
