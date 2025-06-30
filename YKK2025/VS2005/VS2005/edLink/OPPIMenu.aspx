<%@ Page Language="VB" AutoEventWireup="false" CodeFile="OPPIMenu.aspx.vb" Inherits="OPPIMenu" %>

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
        <img src="iMages/OPPIMenu.png" style="z-index: 1; left: 40px; position: absolute;top: 0px; width: 1000px; height: 500px;"/>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  各Action按鈕                                                                        ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
       <asp:ImageButton ID="BMCustomer" runat="server" style="z-index: 1; left: 376px; position: absolute;top: 80px;" ImageUrl="iMages/BMCustomer.png" Height="40px" Width="90px" />
       <asp:ImageButton ID="BMCustomerHistory" runat="server" style="z-index: 1; left: 376px; position: absolute;top: 120px;" ImageUrl="iMages/BMCustomerHistory.png" Height="40px" Width="90px" />
       <asp:ImageButton ID="BMSystem" runat="server" style="z-index: 1; left: 736px; position: absolute;top: 208px;" ImageUrl="iMages/BMSystem.png" Height="40px" Width="90px" />
       <asp:ImageButton ID="BInqReport" runat="server" style="z-index: 1; left: 584px; position: absolute;top: 344px;" ImageUrl="iMages/BInqReport.png" Height="40px" Width="90px" />
       <asp:ImageButton ID="BInqSendOffMail" runat="server" style="z-index: 1; left: 912px; position: absolute;top: 336px;" ImageUrl="iMages/BInqSendOffMail.png" Height="40px" Width="90px" />
       <asp:ImageButton ID="BPIReport" runat="server" style="z-index: 1; left: 296px; position: absolute;top: 400px;" ImageUrl="iMages/BPIReport.png" Height="40px" Width="90px" />

       <asp:ImageButton ID="BSpoolStore" runat="server" style="z-index: 1; left: 904px; position: absolute;top: 16px;" ImageUrl="iMages/BSpoolStore.png" Height="40px" Width="115px"  />
       <asp:ImageButton ID="BOPReportSys" runat="server" style="z-index: 1; left: 904px; position: absolute;top: 56px;" ImageUrl="iMages/BOPReportSys.png" Height="40px" Width="115px" />

       <asp:ImageButton ID="BRigisterSheet" runat="server" style="z-index: 1; left: 904px; position: absolute;top: 96px;" ImageUrl="iMages/BRigisterSheet.png" Height="40px" Width="115px" />
       <asp:ImageButton ID="BChangeSheet" runat="server" style="z-index: 1; left: 904px; position: absolute;top: 136px;" ImageUrl="iMages/BChangeSheet.png" Height="40px" Width="115px" />

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:TextBox ID="DUserID" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFileName" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath1" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath2" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath3" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
       </div>
    </form>
</body>
</html>
