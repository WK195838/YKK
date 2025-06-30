<%@ Page Language="VB" AutoEventWireup="false" CodeFile="YKKColorList.aspx.vb" Inherits="YKKColorList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>色番資料</title>
</head>
<body>
    <form id="form1" runat="server">
        YKK Color&nbsp;<asp:TextBox ID="DData" runat="server" style="z-index: 100; left: 101px; position: absolute; top: 13px" Height="17px"></asp:TextBox><asp:Button ID="Button1" runat="server" Text="查詢" style="z-index: 101; left: 265px; position: absolute; top: 15px" />
        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="DData"
            ErrorMessage="請輸入YKK COLOR" Style="left: 232px; position: relative; top: 0px; z-index: 102;" Width="274px"></asp:RequiredFieldValidator>&nbsp;<br />
        <br />
        &nbsp;&nbsp;<asp:Panel ID="Panel1" runat="server"  Height="300px" Width="563px" ScrollBars="Auto" BorderStyle="Solid" BorderWidth="1px">     
        <asp:GridView id="GridView1" runat="server" Height="34px"  CellPadding="4"  EmptyDataText="查無資料" Width="554px" Font-Size="Small" AutoGenerateColumns="False" AllowSorting="True" PageSize="15">
<RowStyle HorizontalAlign="Center" BackColor="#FFFF99"></RowStyle>
<Columns>
<asp:CommandField ShowSelectButton="True"></asp:CommandField>
<asp:BoundField DataField="Color" HeaderText="Color"></asp:BoundField>
<asp:BoundField DataField="Bwire" HeaderText="B.Wire"></asp:BoundField>
    <asp:BoundField DataField="OpenParts" HeaderText="Open Parts" />
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
