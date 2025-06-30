<%@ Page Language="VB" AutoEventWireup="false" CodeFile="QCList.aspx.vb" Inherits="QCList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Item Code </title>
</head>
<body>
    <form id="form1" runat="server">
        &nbsp; 請輸入客訴編號<asp:TextBox ID="DData" runat="server" style="z-index: 100; left: 192px; position: absolute; top: 8px;position: absolute;text-transform : uppercase;" Height="17px"></asp:TextBox><asp:Button ID="Button1" runat="server" Text="查詢" style="z-index: 101; left: 360px; position: absolute; top: 8px" />
        /業務員<asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="DData"
            ErrorMessage="請輸入客訴編號/業務員" Style="left: 224px; position: relative; top: 0px; z-index: 102;" Width="274px"></asp:RequiredFieldValidator>&nbsp;<br />
        <br />
        &nbsp;&nbsp;<asp:Panel ID="Panel1" runat="server"  Height="300px" Width="598px" ScrollBars="Auto" BorderStyle="Solid" BorderWidth="1px">     
        <asp:GridView id="GridView1" runat="server" Height="34px"  CellPadding="4"  EmptyDataText="查無資料" Width="554px" Font-Size="Small" AutoGenerateColumns="False" AllowSorting="True" PageSize="15">
<RowStyle HorizontalAlign="Center" BackColor="#FFFF99"></RowStyle>
<Columns>
<asp:CommandField ShowSelectButton="True"></asp:CommandField>
<asp:BoundField DataField="No" HeaderText="No"></asp:BoundField>
<asp:BoundField DataField="appname" HeaderText="業務員"></asp:BoundField>
    <asp:BoundField DataField="qcno" HeaderText="客訴編號" />
    <asp:BoundField DataField="cpcontent" HeaderText="客訴內容" />
    <asp:BoundField DataField="formno" HeaderText="formno" />
    <asp:BoundField DataField="formsno" HeaderText="formsno" />
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
