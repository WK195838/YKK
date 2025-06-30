<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DTMW_CombiSheetCopy.aspx.vb" Inherits="DTMW_CombiSheetCopy" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Combi表單</title>
</head>
<body>
    <form id="form1" runat="server">
        <asp:TextBox ID="DData" runat="server" style="z-index: 100; left: 12px; position: absolute; top: 13px" Height="17px" Width="244px"></asp:TextBox><asp:Button ID="Button1" runat="server" Text="查詢" style="z-index: 101; left: 265px; position: absolute; top: 15px" />
        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="DData"
            ErrorMessage="請輸入No" Style="left: 303px; position: relative; top: 2px; z-index: 102;" Width="293px"></asp:RequiredFieldValidator>&nbsp;<br />
        <br />
        &nbsp;&nbsp;<asp:Panel ID="Panel1" runat="server"  Height="579px" Width="598px" ScrollBars="Auto" BorderStyle="Solid" BorderWidth="1px">     
        <asp:GridView id="GridView1" runat="server" Height="34px"  CellPadding="4"  EmptyDataText="查無資料" Width="554px" Font-Size="Small" AutoGenerateColumns="False" AllowSorting="True" PageSize="15">
<RowStyle HorizontalAlign="Center" BackColor="#FFFF99"></RowStyle>
<Columns>
<asp:CommandField ShowSelectButton="True"></asp:CommandField>
<asp:BoundField DataField="STS" HeaderText="狀態"></asp:BoundField>
<asp:BoundField DataField="NO" HeaderText="No"></asp:BoundField>
    <asp:BoundField DataField="CombiItem" HeaderText="項目" />
    <asp:BoundField DataField="Name" HeaderText="擔當者" />
    <asp:HyperLinkField DataNavigateUrlFields="FormNo,FormSno" DataNavigateUrlFormatString="~\DTMW_COMBISheet02.aspx?pFormNo={0}&amp;pFormSno={1}"
        DataTextFormatString="{0:c}" HeaderText="連結" Target="_blank" Text="連結" />
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
