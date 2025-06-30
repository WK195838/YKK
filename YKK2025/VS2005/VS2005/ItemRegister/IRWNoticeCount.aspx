<%@ Page Language="VB" AutoEventWireup="false" CodeFile="IRWNoticeCount.aspx.vb" Inherits="IRWNoticeCount" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>IRW實績-推移</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  GRID                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->


        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" width="500" Style="z-index: 103; left:0px; position: absolute; top: 0px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="YYMM"  HeaderText="">  <ItemStyle HorizontalAlign="Center" />   </asp:BoundField>  
                <asp:BoundField DataField="DayCount" HeaderText="">  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>   
                <asp:BoundField DataField="WorkDay" HeaderText="">   <ItemStyle HorizontalAlign="right" />   </asp:BoundField>    
                <asp:BoundField DataField="Records" HeaderText="" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="LYRecords" HeaderText="" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="Banlance" HeaderText="" >   <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="Remark" HeaderText="" >    <ItemStyle HorizontalAlign="left" />   </asp:BoundField> 
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>

        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" width="500" Style="z-index: 103; left:0px; position: absolute; top: 315px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="DATATYPEDESC"  HeaderText="">  <ItemStyle HorizontalAlign="LEFT" />   </asp:BoundField>  
                <asp:BoundField DataField="NAME" HeaderText="">  <ItemStyle HorizontalAlign="LEFT" />   </asp:BoundField>   
                <asp:BoundField DataField="WARN1" HeaderText="">   <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField>    
                <asp:BoundField DataField="WARN2" HeaderText="" >  <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField>  
                <asp:BoundField DataField="WARN3" HeaderText="" >  <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField>  
                <asp:BoundField DataField="WARN4" HeaderText="" >   <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField>  
                <asp:BoundField DataField="WARN5" HeaderText="" >    <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField> 
                <asp:BoundField DataField="WARN6" HeaderText="" >    <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField> 
                <asp:BoundField DataField="WARN7" HeaderText="" >    <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField> 
                <asp:BoundField DataField="WARN8" HeaderText="" >    <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField> 
                <asp:BoundField DataField="WARN9" HeaderText="" >    <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField> 
                <asp:BoundField DataField="WARN10" HeaderText="" >    <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField> 
                <asp:BoundField DataField="WARN11" HeaderText="" >    <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField> 
                <asp:BoundField DataField="WARN12" HeaderText="" >    <ItemStyle HorizontalAlign="CENTER" />   </asp:BoundField> 
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
