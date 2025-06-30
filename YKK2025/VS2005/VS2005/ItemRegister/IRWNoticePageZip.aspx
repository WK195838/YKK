<%@ Page Language="VB" AutoEventWireup="false" CodeFile="IRWNoticePageZip.aspx.vb" Inherits="IRWNoticePageZip" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>NOTICE PAGE(ZIP)</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  hyperlink                                                      ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->

<asp:hyperlink style="Z-INDEX: 102; LEFT: 8px; POSITION: absolute; TOP: 8px" id="LSearch" runat="server" Target="_blank"
NavigateUrl="" Font-Size="10.5pt" Height="16px" Width="480px" ></asp:hyperlink>
    

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖 & TEXT                                                                         ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Label ID="Label1" runat="server" Style="z-index: 103; left: 0px; position: absolute; top: 30px; TEXT-ALIGN:center"  Font-Size="10pt" Width="350px" BackColor="Blue" Font-Bold="True" ForeColor="White" Height="16px">LT</asp:Label>
        <asp:Label ID="Label2" runat="server" Style="z-index: 103; left: 351px; position: absolute; top: 30px; TEXT-ALIGN:center"  Font-Size="10pt" Width="146px" BackColor="Blue" Font-Bold="True" ForeColor="White" Height="16px">短縮 LT</asp:Label>
        <img src="iMages/CustomerLT.png" style="z-index: 1; left: 0px; position: absolute;top: 45px; width: 500px; height: 150px;"/>

        <asp:TextBox ID="DNowDays" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Red" Height="24px"  Width="46px" Style="z-index: 103; left: 68px; position: absolute; top: 149px; TEXT-ALIGN:center" BorderColor="Yellow" Font-Bold="False" Font-Size="Larger">31.5</asp:TextBox>

        <asp:TextBox ID="DIRWDays"  runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Red" Height="20px"  Width="165px" Style="z-index: 103; left: 139px; position: absolute; top: 139px; TEXT-ALIGN:center" BorderColor="Yellow" Font-Bold="False" Font-Size="Medium">31.5</asp:TextBox>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  GRID                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->

        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" width="500" Style="z-index: 103; left:0px; position: absolute; top: 200px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="YYMM"  HeaderText="YYMM">
                    <ItemStyle Width="80px" />
                    <ItemStyle HorizontalAlign="Center" /> 
                </asp:BoundField>                 
                <asp:BoundField DataField="DataTypeDesc" HeaderText="日">
                    <ItemStyle Width="50px"/>  <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>    
                <asp:BoundField DataField="SALES" HeaderText="SALES">
                    <ItemStyle Width="50px"/>  <ItemStyle HorizontalAlign="right" />
                </asp:BoundField>    
                <asp:HyperLinkField DataNavigateUrlFields="ZIPURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="ZIP" HeaderText="ZIP" Target="_blank">
                    <ItemStyle Width="50px" />   <ItemStyle HorizontalAlign="right" />
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="SLDURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="SLD" HeaderText="SLD" Target="_blank">
                    <ItemStyle Width="50px" />   <ItemStyle HorizontalAlign="right" />
                </asp:HyperLinkField>

                <asp:BoundField DataField="Total" HeaderText="實際" >  <ItemStyle Width="50px" />   <ItemStyle HorizontalAlign="right" />  </asp:BoundField> 
                <asp:BoundField DataField="Target" HeaderText="目標" >  <ItemStyle Width="50px" />  <ItemStyle HorizontalAlign="right" />   </asp:BoundField> 
                <asp:BoundField DataField="Banlance" HeaderText="差" >  <ItemStyle Width="50px" />  <ItemStyle HorizontalAlign="right" />   </asp:BoundField> 
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
