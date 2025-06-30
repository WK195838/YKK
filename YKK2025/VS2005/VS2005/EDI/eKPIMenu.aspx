<%@ Page Language="VB" AutoEventWireup="false" CodeFile="eKPIMenu.aspx.vb" Inherits="eKPIMenu" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>eKPI System 2019 Ver1.0</title>
	    <script language="javascript" src="ForProject.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/eKPIMenu.png" style="z-index: 1; left: 40px; position: absolute;top: 0px; width: 1000px; height: 550px;"/>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  各Action按鈕                                                                        ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
       <asp:ImageButton ID="BMSTMaint" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 40px;" ImageUrl="~/iMages/eKPI_Btn_MST-M.png" Height="40px" Width="90px" />
       <asp:ImageButton ID="BMSTExpand" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 184px;" ImageUrl="~/iMages/eKPI_Btn_MST-E.png" Height="40px" Width="90px" />
       <asp:ImageButton ID="BOFCT" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 280px;" ImageUrl="iMages/eKPI_Btn_OFCT.png" Height="40px" Width="90px" />
       <asp:ImageButton ID="BOFCTCAL" runat="server" style="z-index: 1; left: 496px; position: absolute;top: 232px;" ImageUrl="iMages/eKPI_Btn_OFCTCAL.png" Height="40px" Width="90px" />

       <asp:ImageButton ID="BDecision" runat="server" style="z-index: 1; left: 608px; position: absolute;top: 152px;" ImageUrl="iMages/eKPI_Btn_Decision.png" Height="40px" Width="90px" />
       <asp:ImageButton ID="BFINAL" runat="server" style="z-index: 1; left: 864px; position: absolute;top: 136px;" ImageUrl="iMages/eKPI_Btn_FINAL.png" Height="40px" Width="90px" />
       <asp:ImageButton ID="BREPORT" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 392px;" ImageUrl="iMages/eKPI_Btn_REPORT.png" Height="40px" Width="90px" />
       <asp:ImageButton ID="BADCHECK" runat="server" style="z-index: 1; left: 392px; position: absolute;top: 392px;" ImageUrl="iMages/eKPI_Btn_ADCHECK.png" Height="40px" Width="90px" />

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  執行跑馬燈                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 1列  -->
        <marquee id="ProDecision" runat="server" style="z-index: 1; left: 488px; position: absolute;top: 130px; width: 310px; height: 20px; color: #0000ff; " >■■處理中■■處理中■■處理中■■</marquee>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- System Parameter  -->
        <asp:TextBox ID="DLogID" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DBuyer" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DUserID" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DGRBuyer" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFunList" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- FileData  -->
        <asp:TextBox ID="DFileName" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath1" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath2" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath3" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath4" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath5" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath6" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath7" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

    </div>
    </form>
</body>

</html>