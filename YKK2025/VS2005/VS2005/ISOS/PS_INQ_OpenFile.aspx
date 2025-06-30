<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PS_INQ_OpenFile.aspx.vb" Inherits="PS_INQ_OpenFile" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
	    <script language="javascript" src="ForProject.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- System Parameter  -->

        <asp:TextBox ID="DPOPUPIMAGEFile" runat="server" Height="16px" Style="z-index: 318; left: 32px;
            position: absolute; top: 24px;font-size:10px;background:transparent" Width="752px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
    
        <asp:TextBox ID="DPOPUPIMAGEFile1" runat="server" Height="16px" Style="z-index: 318; left: 32px;
            position: absolute; top: 44px;font-size:10px;background:transparent" Width="752px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:Button ID="BGO" runat="server" Height="25px" Style="z-index: 104; left: 568px;
            position: absolute; top: 104px" Text="GO" Width="45px" />

    </div>
    </form>
</body>
</html>
