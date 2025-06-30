<%@ Page Language="VB" AutoEventWireup="false" CodeFile="EDLMenu.aspx.vb" Inherits="EDLMenu" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>edLink System 2019 Ver1.0</title>
	    <script language="javascript" src="ForProject.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/EDLMenu.png" style="z-index: 1; left: 40px; position: absolute;top: 0px; width: 800px; height: 500px;"/>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  各Action按鈕                                                                        ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:HyperLink ID="LDebitNote" runat="server" style="z-index: 1; left: 124px; position: absolute;top: 94px;" Font-Bold="True" Font-Italic="False" Font-Names="標楷體" Font-Underline="False" NavigateUrl="http://10.245.1.10/portal/管理/SAPSupportSystem/ARMail/客戶主檔處理履歷/tabid/243/Default.aspx" Target="_parent" >A/R對帳單</asp:HyperLink>
        <asp:HyperLink ID="LOrderProgress" runat="server" style="z-index: 1; left: 264px; position: absolute;top: 206px;" Font-Bold="True" Font-Italic="False" Font-Names="標楷體" Font-Underline="False" NavigateUrl="http://10.245.1.10/Portal/GetPortalUser.aspx?pURL=http://10.245.0.3/edLink/OPPIMenu.aspx" Target="_self" >OP訂單進度</asp:HyperLink>
        <asp:HyperLink ID="LPackingInf" runat="server" style="z-index: 1; left: 408px; position: absolute;top: 94px;" Font-Bold="True" Font-Italic="False" Font-Names="標楷體" Font-Underline="False" NavigateUrl="http://10.245.1.10/Portal/GetPortalUser.aspx?pURL=http://10.245.0.3/edLink/OPPIMenu.aspx" Target="_self" >PI包裝清單</asp:HyperLink>

        <asp:HyperLink ID="HyperLink1" runat="server" style="z-index: 1; left: 550px; position: absolute;top: 206px;" Font-Bold="True" Font-Italic="False" Font-Names="標楷體" Font-Underline="False" Target="_self" Font-Size="Larger" >無</asp:HyperLink>
        <asp:HyperLink ID="HyperLink2" runat="server" style="z-index: 1; left: 694px; position: absolute;top: 94px;" Font-Bold="True" Font-Italic="False" Font-Names="標楷體" Font-Underline="False" Target="_self" Font-Size="Larger" >無</asp:HyperLink>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        </div>
    </form>
</body>
</html>
