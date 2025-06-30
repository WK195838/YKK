<%@ Page Language="VB" AutoEventWireup="false" CodeFile="INQ_CustomerGroup.aspx.vb" Inherits="INQ_CustomerGroup" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>顧客群組(履歷)</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  篩選條件                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        客戶／客戶代碼 &nbsp;<asp:TextBox ID="DCust" runat="server" BackColor="#FFFF80"></asp:TextBox><br />
<!--  -->
        業務員 &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DSalesName" runat="server" BackColor="#FFFF80"></asp:TextBox>
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp;<br />
<!--  -->
        業務員代號 &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DSalesCode" runat="server" BackColor="#FFFF80"></asp:TextBox>
<!--  -->
        <asp:Button ID="BSearch" runat="server" Text="搜尋" Width="80px" />
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  顧客資料                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="DGridView1" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 15px; position: absolute; top: 100px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" Width="2000px" >
            <Columns>
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="GroupCode" HeaderText="顧客" Target="_blank">
                </asp:HyperLinkField>

                <asp:BoundField DataField="GroupName" HeaderText="顧客名" />
                <asp:BoundField DataField="Active" HeaderText="狀態" />
                <asp:BoundField DataField="GroupType" HeaderText="類型" />
                <asp:BoundField DataField="SalesCode" HeaderText="業務員" />
                <asp:BoundField DataField="SalesName" HeaderText="業務員名" />
                <asp:BoundField DataField="SalesAddress" HeaderText="業務員Mail" />
                <asp:BoundField DataField="UseType" HeaderText="使用區分" />

                <asp:BoundField DataField="OPActive" HeaderText="OP.啟用" />
                <asp:BoundField DataField="OPSendActive" HeaderText="OP.郵件" />
                <asp:BoundField DataField="OPRptType" HeaderText="OP.報表" />

                <asp:BoundField DataField="PIActive" HeaderText="PI.啟用" />
                <asp:BoundField DataField="PISendActive" HeaderText="PI.郵件" />
                <asp:BoundField DataField="PIRptType" HeaderText="PI.報表" />

                <asp:BoundField DataField="PIPActive" HeaderText="PIP.啟用" />
                <asp:BoundField DataField="PIPSendActive" HeaderText="PIP.郵件" />
                <asp:BoundField DataField="PIPRptType" HeaderText="PIP.報表" />

                <asp:BoundField DataField="SNPActive" HeaderText="SNP.啟用" />
                <asp:BoundField DataField="SNPSendActive" HeaderText="SNP.郵件" />
                <asp:BoundField DataField="SNPRptType" HeaderText="SNP.報表" />

                <asp:BoundField DataField="PLPActive" HeaderText="PLP.啟用" />
                <asp:BoundField DataField="PLPSendActive" HeaderText="PLP.郵件" />
                <asp:BoundField DataField="PLPRptType" HeaderText="PLP.報表" />

                <asp:BoundField DataField="IVPActive" HeaderText="IVP.啟用" />
                <asp:BoundField DataField="IVPSendActive" HeaderText="IVP.郵件" />
                <asp:BoundField DataField="IVPRptType" HeaderText="IVP.報表" />
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