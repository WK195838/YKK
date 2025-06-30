<%@ Page Language="VB" AutoEventWireup="false" CodeFile="INQ_MailCustomerGroupHistory.aspx.vb" Inherits="INQ_MailCustomerGroupHistory" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>郵件履歷</title>
    <script language="javascript">
    function OpenFile(doc)
    {
        window.open(doc);
        return false;
    }

    function GoBack()
    {
        history.go(-1);
    }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  顧客資料                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!--
                <asp:BoundField DataField="RCreateTime" HeaderText="報表作成時間" />
                <asp:BoundField DataField="RFile" HeaderText="報表" />

                <asp:BoundField DataField="WFile" HeaderText="放置處(W)" />
                <asp:BoundField DataField="HFile" HeaderText="放置處(H)" />

                <asp:BoundField DataField="CreateUser" HeaderText="作成者" />
                <asp:BoundField DataField="CreateTime" HeaderText="作成時間" />
                <asp:BoundField DataField="ModifyUser" HeaderText="修改者" />
                <asp:BoundField DataField="ModifyTime" HeaderText="修改時間" />
-->

        <asp:GridView ID="DGridView1" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 15px; position: absolute; top: 10px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" Width="1200px" >
            <Columns>
                <asp:BoundField DataField="GroupCode" HeaderText="顧客" />
                <asp:BoundField DataField="GroupName" HeaderText="顧客名" />
                <asp:BoundField DataField="RType" HeaderText="類型" />
                <asp:BoundField DataField="RID" HeaderText="ID" />

                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="RAttachFile" HeaderText="" Target="_blank">
                </asp:HyperLinkField>

                <asp:BoundField DataField="MCreateTime" HeaderText="郵件傳送時間" />
                <asp:BoundField DataField="MSender" HeaderText="送件者" />
                <asp:BoundField DataField="MSenderName" HeaderText="" />
                <asp:BoundField DataField="MTo" HeaderText="收件者" />
                <asp:BoundField DataField="MToName" HeaderText="" />
                <asp:BoundField DataField="MCc" HeaderText="CC者" />
                <asp:BoundField DataField="MBcc" HeaderText="BCC者" />
                <asp:BoundField DataField="MSubject" HeaderText="主旨" />
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