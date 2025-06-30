<%@ Page Language="VB" AutoEventWireup="false" CodeFile="EMPList.aspx.vb" Inherits="EMPList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>個人資料</title>
</head>
<body>
    <form id="form1" runat="server">
        尋找部門 / 姓名<asp:TextBox ID="DData" runat="server" style="z-index: 100; left: 128px; position: absolute;;text-transform : uppercase; top: 16px" Height="17px" Width="280px"></asp:TextBox><asp:Button ID="Button1" runat="server" Text="查詢" style="z-index: 101; left: 424px; position: absolute; top: 16px" />
        <br />
        <br />
        <asp:Panel ID="Panel1" runat="server"  Height="300px" Width="598px" ScrollBars="Auto" BorderStyle="Solid" BorderWidth="1px">    
         
        <asp:GridView id="GridView1" runat="server" Height="34px"  CellPadding="4"  EmptyDataText="查無資料" Width="554px" Font-Size="Small" AutoGenerateColumns="False" AllowSorting="True" PageSize="15">
<RowStyle HorizontalAlign="Center" BackColor="#FFFF99"></RowStyle>
<Columns>
<asp:CommandField ShowSelectButton="True"></asp:CommandField>
<asp:BoundField DataField="DepName" HeaderText="部門"></asp:BoundField>
<asp:BoundField DataField="EMPName" HeaderText="姓名"></asp:BoundField>
<asp:BoundField DataField="DepID" HeaderText="部門代碼"></asp:BoundField>
<asp:BoundField DataField="EMPID" HeaderText="員工ID"></asp:BoundField>
<asp:BoundField DataField="JobTitle" HeaderText="職稱"></asp:BoundField>
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
