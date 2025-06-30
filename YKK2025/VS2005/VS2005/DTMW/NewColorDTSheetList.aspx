<%@ Page Language="VB" AutoEventWireup="false" CodeFile="NewColorDTSheetList.aspx.vb" Inherits="NewColorDTSheetList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>新色依賴完成表</title>
</head>
<body>
    <form id="form1" runat="server">
        <br />
        <asp:Panel ID="Panel1" runat="server"  Height="579px" Width="598px" ScrollBars="Auto" BorderStyle="Solid" BorderWidth="1px" style="z-index: 100; left: 10px; position: absolute; top: 34px">     
        <asp:GridView id="GridView1" runat="server" Height="34px"  CellPadding="4"  EmptyDataText="查無資料" Width="554px" Font-Size="Small" AutoGenerateColumns="False" AllowSorting="True" PageSize="15">
<RowStyle HorizontalAlign="Center" BackColor="#FFFF99"></RowStyle>
<Columns>
    <asp:HyperLinkField DataTextField="NO" HeaderText="NO" NavigateUrl="BoardEdit.aspx"
        Target="_blank" Text="NO" DataNavigateUrlFields="URL" />
    <asp:BoundField DataField="DTSheet" HeaderText="核可單種類" />
    <asp:BoundField DataField="newoldcolor" HeaderText="新色/舊色" />
    <asp:BoundField DataField="ykkcolortype" HeaderText="YKK色別" />
    <asp:BoundField DataField="ykkcolorcode" HeaderText="YKK色號" />
</Columns>

<HeaderStyle BackColor="#FFCC99"></HeaderStyle>

<AlternatingRowStyle BackColor="#CCFFCC"></AlternatingRowStyle>
</asp:GridView>
 </asp:Panel> 
        <br />
        &nbsp;&nbsp;&nbsp; &nbsp;&nbsp;<br />
        &nbsp;&nbsp;<br />
        <br />
        &nbsp;
            <br />
    </form>
</body>
</html>
