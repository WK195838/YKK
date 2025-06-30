<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ReferenceListCL.aspx.vb" Inherits="ReferenceListCL" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >

<head id="Head1" runat="server">
    <title>參照經費</title>
</head>
<body>
    <form id="form1" runat="server">
        尋找&nbsp;
        <asp:TextBox ID="DData" runat="server" style="z-index: 100; left: 56px; position: absolute;;text-transform : uppercase; top: 10px" Height="24px" Width="304px"  BackColor="yellow"></asp:TextBox>
        <asp:Button ID="Button1" runat="server" Text="查詢" style="z-index: 101; left: 376px; position: absolute; top: 10px" Height="32px" Width="64px" />
        <br />
        <br />
        <asp:Panel ID="Panel1" runat="server"  Height="300px" Width="650px" ScrollBars="Auto" BorderStyle="Solid" BorderWidth="1px">    
         
        <asp:GridView id="GridView1" runat="server" Height="34px"  CellPadding="4"  EmptyDataText="查無資料" Width="650px" Font-Size="Small" AutoGenerateColumns="False" AllowSorting="True" PageSize="15">
<RowStyle HorizontalAlign="Left" BackColor="#FFFF99"></RowStyle>
<Columns>
<asp:CommandField ShowSelectButton="True"></asp:CommandField>

                <asp:BoundField DataField="ExpItem" HeaderText="類別用途"  />
                <asp:BoundField DataField="ADate" HeaderText="日期"  />
                <asp:BoundField DataField="TaxType" HeaderText="稅別"  />
                <asp:BoundField DataField="NetAmt" HeaderText="淨額" DataFormatString="{0:N0}"  />
                <asp:BoundField DataField="TaxAmt" HeaderText="稅額" DataFormatString="{0:N0}"  />
                <asp:BoundField DataField="Amt" HeaderText="總額"    DataFormatString="{0:N0}"  />
                <asp:BoundField DataField="Content" HeaderText="內容"   />
                <asp:BoundField DataField="Remark" HeaderText="備註"  />
    <asp:BoundField DataField="ACID" HeaderText="ACID" />


</Columns>

<HeaderStyle BackColor="#FFCC99"></HeaderStyle>

<AlternatingRowStyle BackColor="#CCFFCC"></AlternatingRowStyle>
</asp:GridView>
 </asp:Panel>
  &nbsp; &nbsp;<br />
        &nbsp;&nbsp;<br />
        <br />
        &nbsp;
            <br />
    </form>
</body>
</html>
