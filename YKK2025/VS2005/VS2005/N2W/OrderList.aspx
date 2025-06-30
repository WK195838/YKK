<%@ Page Language="VB" AutoEventWireup="false" CodeFile="OrderList.aspx.vb" Inherits="OrderList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>訂單資料</title>
</head>
<body>
    <form id="form1" runat="server">
        &nbsp; 訂單號碼：OR&nbsp;&nbsp;<asp:TextBox ID="DData" runat="server" style="z-index: 100; left: 120px; position: absolute; top: 8px;position: absolute;text-transform : uppercase;" Height="17px"></asp:TextBox><asp:Button ID="Button1" runat="server" Text="查詢" style="z-index: 101; left: 288px; position: absolute; top: 8px" />
        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="DData"
            ErrorMessage="請輸入訂單號碼：OR" Style="left: -14px; position: relative; top: 24px; z-index: 102;" Width="274px"></asp:RequiredFieldValidator>&nbsp;<br />
        <br />
        &nbsp;&nbsp;<asp:Panel ID="Panel1" runat="server"  Height="300px" Width="598px" ScrollBars="Auto" BorderStyle="Solid" BorderWidth="1px">     
        <asp:GridView id="GridView1" runat="server" Height="34px"  CellPadding="4"  EmptyDataText="查無資料" Width="554px" Font-Size="Small" AutoGenerateColumns="False" AllowSorting="True" PageSize="15">
<RowStyle HorizontalAlign="Center" BackColor="#FFFF99"></RowStyle>
<Columns>
<asp:CommandField ShowSelectButton="True"></asp:CommandField>
<asp:BoundField DataField="ORDN5C" HeaderText="訂單號碼"></asp:BoundField>
    <asp:BoundField DataField="DK1C5C" HeaderText="顧客名稱" />
<asp:BoundField DataField="BYRC5C" HeaderText="Buyer"></asp:BoundField>
</Columns>

<HeaderStyle BackColor="#FFCC99"></HeaderStyle>

<AlternatingRowStyle BackColor="#CCFFCC"></AlternatingRowStyle>
</asp:GridView>
 </asp:Panel> &nbsp; &nbsp;<br />
        &nbsp;&nbsp;<br />
        <br />
        &nbsp;
            <br />
    </form>
</body>
</html>
