<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfF_TraceLSOrder.aspx.vb" Inherits="InfF_TraceLSOrder" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>LS Order Progress Detail Inf.</title>
</head>
<body style="font-size: 12pt">
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  按鈕                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- REMARK    	<asp:button id="BInq" runat="server" Height="24px" Width="40px" ForeColor="Black" BackColor="Silver" Text="查詢"></asp:button>      
REMARK -->

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Forecast Plan GridView                                                              ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
FCT Inf.
        <asp:GridView ID="FCTGridView" runat="server" AutoGenerateColumns="False"  CellPadding="3" Font-Size="9pt" PageSize="20" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" ForeColor="Black" GridLines="Vertical">
            <Columns>
                <asp:BoundField HeaderText="FCTNo" DataField="FCTNo" />
                <asp:BoundField HeaderText="Sub" DataField="FCTSubNo" />
                <asp:BoundField HeaderText="PLM" DataField="C_Code" />
                <asp:BoundField HeaderText="Style" DataField="C_Style" />
                <asp:BoundField HeaderText="INF-1" DataField="C_A1" />
                <asp:BoundField HeaderText="INF-2" DataField="C_B1" />
                <asp:BoundField HeaderText="Color" DataField="C_Color" />
                <asp:BoundField HeaderText="Keep" DataField="C_ShortenLT" />
                <asp:BoundField HeaderText="Factory" DataField="C_C1" />

                <asp:BoundField HeaderText="Code" DataField="Y_ItemCode" />
                <asp:BoundField HeaderText="ItemName" DataField="Y_ItemName" />
                <asp:BoundField HeaderText="Color" DataField="Y_Color" />

                <asp:BoundField DataField="BLANK" HeaderText="|"  />
                <asp:BoundField HeaderText="N0_F" DataField="N_F" >  <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>                
                <asp:BoundField HeaderText="N1_F" DataField="N1_F" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>  
                <asp:BoundField HeaderText="N2_F" DataField="N2_F" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField> 
                <asp:BoundField HeaderText="N3_F" DataField="N3_F" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField> 
                <asp:BoundField HeaderText="N4_F" DataField="N4_F" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField> 
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Local Stock Plan GridView                                                           ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
LS Inf.
        <asp:GridView ID="LSGridView" runat="server" AutoGenerateColumns="False" CellPadding="4" Font-Size="9pt" PageSize="20" ForeColor="#333333" 
GridLines="None">
            <Columns>
                <asp:BoundField DataField="BULSNo" HeaderText="LSNo"  />
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
<!-- ++  OrderProgress GridView                                                             ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
Order Inf.
        <asp:GridView ID="PGGridView" runat="server" AutoGenerateColumns="False"  CellPadding="4" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#DEDFDE" BorderStyle="None" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="OrderNo" HeaderText="OrderNo"  />
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
            <HeaderStyle BackColor="#6B696B" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <RowStyle BackColor="#F7F7DE" />
            <FooterStyle BackColor="#CCCC99" />
            <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
            <SelectedRowStyle BackColor="#CE5D5A" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="White" />
        </asp:GridView>        
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    </div>
    </form>
</body>
</html>
