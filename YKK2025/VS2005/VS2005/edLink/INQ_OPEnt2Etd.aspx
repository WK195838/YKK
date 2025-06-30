<%@ Page Language="VB" AutoEventWireup="false" CodeFile="INQ_OPEnt2Etd.aspx.vb" Inherits="INQ_OPEnt2Etd" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>OrderProgress Ent2Etd</title>
</head>
<body>

    <form id="form1" runat="server">
    <div>

        <asp:Label ID="Label1" runat="server" Style="z-index: 103; left: 32px; position: absolute; top: 16px"  Font-Size="16pt"  Width="160px" >Confirm-ETD日數</asp:Label>

        <asp:TextBox ID="Ddays" runat="server" BackColor="#80FF80" BorderStyle="Groove"
                ForeColor="Blue"  Width="74px" Style="z-index: 103; left: 200px; position: absolute; top: 16px" Font-Size="14pt"></asp:TextBox>

        <asp:Button ID="BSearch" runat="server" Text="GO" Width="72px" style="Z-INDEX: 100; LEFT: 288px; POSITION: absolute; TOP: 16px" Height="32px" />          

        <asp:TextBox ID="DRemark" runat="server" Style="z-index: 100; left: 376px; position: absolute; 
            top: 16px" BackColor="White" Width="472px" ReadOnly="true" BorderColor="Black" Font-Size="16pt"> BULK &amp; Order Complete Flag=未完成 &amp; ST1C=1 </asp:TextBox>

                
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Height="300px" Width="1000px"
        Style="z-index: 103; left: 32px; position: absolute; top: 72px" CellPadding="3" Font-Size="16pt" 
        GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                
                <asp:BoundField DataField="YYMM" HeaderText="Month" ItemStyle-HorizontalAlign="Center" />

                <asp:BoundField DataField="XBLUK" HeaderText="" ItemStyle-HorizontalAlign="Center" />

                <asp:BoundField DataField="WITHIN" HeaderText="後延" ItemStyle-HorizontalAlign="Center"/> 
                <asp:BoundField DataField="WITHINPer" HeaderText="" ItemStyle-HorizontalAlign="Center"/> 
                <asp:BoundField DataField="WITHINSel" HeaderText="" ItemStyle-HorizontalAlign="Center"/> 


                <asp:BoundField DataField="WITHOUT" HeaderText="前倒"  ItemStyle-HorizontalAlign="Center"/>
                <asp:BoundField DataField="WITHOUTPer" HeaderText="" ItemStyle-HorizontalAlign="Center"/> 
                <asp:BoundField DataField="WITHOUTSel" HeaderText="" ItemStyle-HorizontalAlign="Center"/> 

                <asp:BoundField DataField="EQUAL" HeaderText="同樣" ItemStyle-HorizontalAlign="Center" />
                <asp:BoundField DataField="EQUALPer" HeaderText="" ItemStyle-HorizontalAlign="Center"/> 
                <asp:BoundField DataField="EQUALSel" HeaderText="" ItemStyle-HorizontalAlign="Center"/> 

                <asp:BoundField DataField="Total" HeaderText="計" ItemStyle-HorizontalAlign="Center" />

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
