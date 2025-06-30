<%@ Page Language="VB" AutoEventWireup="false" CodeFile="QCEDXList.aspx.vb" Inherits="QCEDXList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>QC EDX LIST (5Y-BEFORE)</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  底圖                                                                                ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/QCEDXList.png" style="z-index: 1; left: 16px; position: absolute;top: 8px; width: 135px; height: 45px;" />

		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:TextBox ID="TextBox1" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 56px; position: absolute;
            top: 64px; text-align: left" Width="80px">Puller Code</asp:TextBox>
        <asp:TextBox ID="DKPullerCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="7" Style="z-index: 126; left: 136px;
            position: absolute; top: 64px; text-align: left" Width="100px"></asp:TextBox>
        <asp:TextBox ID="TextBox11" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 248px; position: absolute;
            top: 64px; text-align: left" Width="80px">Color</asp:TextBox>
        <asp:TextBox ID="DKColor" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="7" Style="z-index: 126; left: 328px;
            position: absolute; top: 64px; text-align: left" Width="48px"></asp:TextBox>
        &nbsp;&nbsp;
        <asp:TextBox ID="TextBox12" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 400px; position: absolute;
            top: 64px; text-align: left" Width="80px">Search</asp:TextBox>
        <asp:TextBox ID="DKOther" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="50" Style="z-index: 126; left: 480px;
            position: absolute; top: 64px; text-align: left" Width="100px"></asp:TextBox>
        <asp:Button ID="BFind" runat="server" Height="25px" Style="z-index: 104; left: 600px;
            position: absolute; top: 64px" Text="Go" Width="45px" />
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BackColor="White"
            CellPadding="3" Font-Size="10pt" ForeColor="Black" GridLines="Vertical" Style="z-index: 103;
            left: 56px; position: absolute; top: 96px">
            <Columns>

                <asp:BoundField DataField="Cat" HeaderText="Catogory"> </asp:BoundField>    
                
                <asp:BoundField DataField="Size"  HeaderText="Size"> </asp:BoundField>            
                <asp:BoundField DataField="Family"  HeaderText="Family"></asp:BoundField>            
                <asp:BoundField DataField="Body"  HeaderText="Body"> </asp:BoundField>            
                <asp:BoundField DataField="Puller"  HeaderText="Puller"> </asp:BoundField>            
                <asp:BoundField DataField="Color"  HeaderText="Color"> </asp:BoundField>            
                <asp:BoundField DataField="Finish"  HeaderText="Finish"> </asp:BoundField>            

                <asp:BoundField DataField="Supplier"  HeaderText="Supplier"> </asp:BoundField>            

                <asp:BoundField DataField="CreateTime"  HeaderText="EDX Date"> </asp:BoundField>            

            </Columns>
            <HeaderStyle BackColor="Black" Font-Bold="True" ForeColor="White" HorizontalAlign="Center"
                VerticalAlign="Middle" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#7FFF00" Font-Bold="false" ForeColor="black" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>
        &nbsp;&nbsp;
        <asp:TextBox ID="DOther" runat="server" BackColor="white" BorderStyle="None" Height="18px"
            MaxLength="7" Style="z-index: 126; left: 664px; position: absolute; top: 64px;
            text-align: left" Width="312px">顯示150件(件數過多)，請使用篩選</asp:TextBox>
    </div>
    </form>
</body>
</html>
