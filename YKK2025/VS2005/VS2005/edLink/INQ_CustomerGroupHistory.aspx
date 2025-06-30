<%@ Page Language="VB" AutoEventWireup="false" CodeFile="INQ_CustomerGroupHistory.aspx.vb" Inherits="INQ_CustomerGroupHistory" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>顧客履歷</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  顧客資料                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="DGridView1" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 15px; position: absolute; top: 10px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" Width="5000px" >
            <Columns>
                <asp:BoundField DataField="HistoryID" HeaderText="ID" />
                <asp:BoundField DataField="HistoryCode" HeaderText="動作" />
                <asp:BoundField DataField="GroupCode" HeaderText="顧客" />
                <asp:BoundField DataField="GroupName" HeaderText="顧客名" />
                <asp:BoundField DataField="Active" HeaderText="狀態" />
                <asp:BoundField DataField="GroupType" HeaderText="類型" />
                <asp:BoundField DataField="SalesCode" HeaderText="業務員" />
                <asp:BoundField DataField="SalesName" HeaderText="業務員名" />
                <asp:BoundField DataField="SalesAddress" HeaderText="業務員Mail" />

                <asp:BoundField DataField="OPActive" HeaderText="OP.啟用" />
                <asp:BoundField DataField="OPSendActive" HeaderText="OP.郵件" />
                <asp:BoundField DataField="OPResendActive" HeaderText="OP.再傳郵件" />
                <asp:BoundField DataField="OPRptType" HeaderText="OP.報表" />
                <asp:BoundField DataField="OPRptLastTime" HeaderText="OP.最後報表" />
                <asp:BoundField DataField="OPSendLastTime" HeaderText="OP.最後郵件" />

                <asp:BoundField DataField="PIActive" HeaderText="PI.啟用" />
                <asp:BoundField DataField="PISendActive" HeaderText="PI.郵件" />
                <asp:BoundField DataField="PIResendActive" HeaderText="PI.再傳郵件" />
                <asp:BoundField DataField="PIRptType" HeaderText="PI.報表" />
                <asp:BoundField DataField="PIRptLastTime" HeaderText="PI.最後報表" />
                <asp:BoundField DataField="PISendLastTime" HeaderText="PI.最後郵件" />

                <asp:BoundField DataField="SalesNote" HeaderText="SalesNote" />
                <asp:BoundField DataField="PINote" HeaderText="PINote" />
                <asp:BoundField DataField="Fax" HeaderText="FAXNo." />

                <asp:BoundField DataField="CreateUser" HeaderText="作成者" />
                <asp:BoundField DataField="CreateTime" HeaderText="作成時間" />
                <asp:BoundField DataField="ModifyUser" HeaderText="修改者" />
                <asp:BoundField DataField="ModifyTime" HeaderText="修改時間" />
                <asp:BoundField DataField="SystemUser" HeaderText="系統更新者" />
                <asp:BoundField DataField="SystemTime" HeaderText="系統更新時間" />
                
                <asp:BoundField DataField="ToList" HeaderText="M to" />
                <asp:BoundField DataField="CcList" HeaderText="M cc" />
                <asp:BoundField DataField="BccList" HeaderText="M bcc" />
                <asp:BoundField DataField="CCodeList" HeaderText="抬頭" />
                <asp:BoundField DataField="OPFieldList" HeaderText="OP.自由欄位" />
                <asp:BoundField DataField="PIFieldList" HeaderText="PI.自由欄位" />
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