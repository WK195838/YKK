<%@ Page Language="VB" AutoEventWireup="false" CodeFile="INQ_REQStatus.aspx.vb" Inherits="INQ_REQStatus" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >


<head runat="server">
    <title>Adjust Status-00</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<asp:TextBox ID="DRemark" runat="server" Style="z-index: 100; left: 24px; position: absolute; 
            top: 24px" BackColor="White" Width="752px" ReadOnly="true" BorderColor="Black" Font-Size="16pt"> BULK &amp; ST1C=1 &amp; 顧客TYPE=一般(A國內, H輸出)/YKK GROUP(E國內,K輸出) </asp:TextBox>
                
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Height="300px" Width="1000px"
        Style="z-index: 103; left: 24px; position: absolute; top: 60px" CellPadding="3" Font-Size="16pt" 
        GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                
                <asp:BoundField DataField="YYMM" HeaderText="Month" ItemStyle-HorizontalAlign="Center" />

                <asp:BoundField DataField="BULK" HeaderText="" ItemStyle-HorizontalAlign="Center" />

                <asp:BoundField DataField="Postpone" HeaderText="後延" ItemStyle-HorizontalAlign="Center"/> 
                <asp:BoundField DataField="PPer" HeaderText="" ItemStyle-HorizontalAlign="Center"/> 
                <asp:BoundField DataField="PSelect" HeaderText="" ItemStyle-HorizontalAlign="Center"/> 


                <asp:BoundField DataField="Advance" HeaderText="前倒"  ItemStyle-HorizontalAlign="Center"/>
                <asp:BoundField DataField="APer" HeaderText="" ItemStyle-HorizontalAlign="Center"/> 
                <asp:BoundField DataField="ASelect" HeaderText="" ItemStyle-HorizontalAlign="Center"/> 

                <asp:BoundField DataField="Same" HeaderText="同樣" ItemStyle-HorizontalAlign="Center" />
                <asp:BoundField DataField="SPer" HeaderText="" ItemStyle-HorizontalAlign="Center"/> 

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
