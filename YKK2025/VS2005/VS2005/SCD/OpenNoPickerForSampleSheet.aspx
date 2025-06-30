<%@ Page Language="VB" AutoEventWireup="false" CodeFile="OpenNoPickerForSampleSheet.aspx.vb" Inherits="OpenNoPickerForSampleSheet" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>選取開發資料</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="DKey" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" Style="z-index: 104; left: 2px; position: absolute; top: 8px" Width="318px">DKey</asp:TextBox>

        <asp:Button ID="BKey" runat="server" Height="25px" Style="z-index: 103; left: 326px;
            position: absolute; top: 8px" Text="...." Width="26px" />

        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            PageSize="20" Style="z-index: 114; left: 2px; position: absolute; top: 38px"
            Width="350px">
            <RowStyle BackColor="White" ForeColor="Blue" />
            <Columns>
                <asp:CommandField ShowSelectButton="True" />
                <asp:BoundField DataField="StsDesc" HeaderText="狀態" />
                <asp:BoundField DataField="No" HeaderText="No." />
                <asp:BoundField DataField="CodeNo" HeaderText="Code" />
                <asp:BoundField DataField="DevNo" HeaderText="開發No." />
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="NoDesc" HeaderText="連結" Target="_blank">
                </asp:HyperLinkField>
            </Columns>
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#FF8000" BorderStyle="Groove" Font-Bold="True" ForeColor="White" />
            <FooterStyle BorderStyle="Groove" />
        </asp:GridView>

    </div>
    </form>
</body>
</html>
