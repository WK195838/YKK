<%@ Page Language="VB" AutoEventWireup="false" CodeFile="AdminMenu.aspx.vb" Inherits="AdminMenu" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>ISIP Admin Menu</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  底圖                                                                                ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/AdminMenu.png" style="z-index: 1; left: 40px; position: absolute;top: 10px; width: 760px; height: 710px;" />
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  MST button                                                                           ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:ImageButton ID="BSharePuller" runat="server" style="z-index: 1; left: 176px; position: absolute;top: 136px; width: 135px; height: 45px;" ImageUrl="iMages/SharePuller.png"  />
        <asp:ImageButton ID="BShortPuller" runat="server" style="z-index: 1; left: 320px; position: absolute;top: 136px; width: 135px; height: 45px;" ImageUrl="iMages/ShortPuller.png"  />
        <asp:ImageButton ID="BColorJump" runat="server" style="z-index: 1; left: 464px; position: absolute;top: 136px; width: 135px; height: 45px;" ImageUrl="iMages/ColorJump.png"  />
        <asp:ImageButton ID="BNextColorNo" runat="server" style="z-index: 1; left: 608px; position: absolute;top: 136px; width: 135px; height: 45px;" ImageUrl="iMages/NextColorNo.png"  />
        <asp:ImageButton ID="BRegisterAuthority" runat="server" style="z-index: 1; left: 176px; position: absolute;top: 184px; width: 135px; height: 45px;" ImageUrl="iMages/RegisterAuthority.png"  />
        <asp:ImageButton ID="BNonRubber" runat="server" style="z-index: 1; left: 320px; position: absolute;top: 184px; width: 135px; height: 45px;" ImageUrl="iMages/NonRubber.png"  />
        <asp:ImageButton ID="BBodyList" runat="server" style="z-index: 1; left: 464px; position: absolute;top: 184px; width: 135px; height: 45px;" ImageUrl="iMages/BodyList.png"  />
        <asp:ImageButton ID="BStatusRebuild" runat="server" style="z-index: 1; left: 608px; position: absolute;top: 184px; width: 135px; height: 45px;" ImageUrl="iMages/StatusRebuild.png"  />
        <asp:ImageButton ID="BQCEDXList" runat="server" style="z-index: 1; left: 176px; position: absolute;top: 232px; width: 135px; height: 45px;" ImageUrl="iMages/QCEDXList.png"  />
        <asp:ImageButton ID="BSPDStatusRebuild" runat="server" style="z-index: 1; left: 320px; position: absolute;top: 232px; width: 135px; height: 45px;" ImageUrl="iMages/SPDStatusRebuild.png"  />
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  STATUS button 45                                                                           ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:ImageButton ID="BRegisterHistory" runat="server" style="z-index: 1; left: 176px; position: absolute;top: 352px; width: 135px; height: 45px;" ImageUrl="iMages/MagApplyHistory.png"  />
        <asp:ImageButton ID="BDeplicatPuller" runat="server" style="z-index: 1; left: 320px; position: absolute;top: 352px; width: 135px; height: 45px;" ImageUrl="iMages/DeplicatPuller.png"  />
        <asp:ImageButton ID="BDrop" runat="server" style="z-index: 1; left: 464px; position: absolute;top: 352px; width: 135px; height: 45px;" ImageUrl="iMages/Drop.png"  />
        <asp:ImageButton ID="BWIP" runat="server" style="z-index: 1; left: 608px; position: absolute;top: 352px; width: 135px; height: 45px;" ImageUrl="iMages/WIPList.png"  />
        <asp:ImageButton ID="BManyBuyer" runat="server" style="z-index: 1; left: 176px; position: absolute;top: 400px; width: 135px; height: 45px;" ImageUrl="iMages/ManyBuyer.png"  />
        <asp:ImageButton ID="BManySupplier" runat="server" style="z-index: 1; left: 320px; position: absolute;top: 400px; width: 135px; height: 45px;" ImageUrl="iMages/ManySupplier.png"  />
        
        <asp:ImageButton ID="BISIP2RD" runat="server" style="z-index: 1; left: 464px; position: absolute;top: 400px; width: 135px; height: 45px;" ImageUrl="iMages/ISIP2RD.png"  />
        <asp:ImageButton ID="BISIP2EDX" runat="server" style="z-index: 1; left: 608px; position: absolute;top: 400px; width: 135px; height: 45px;" ImageUrl="iMages/ISIP2EDX.png"  />
        <asp:ImageButton ID="BISIP2IRW" runat="server" style="z-index: 1; left: 176px; position: absolute;top: 448px; width: 135px; height: 45px;" ImageUrl="iMages/ISIP2IRW.png"  />
        <asp:ImageButton ID="BNoOrderReceiver" runat="server" style="z-index: 1; left: 320px; position: absolute;top: 448px; width: 135px; height: 45px;" ImageUrl="iMages/NoOrderReceiver.png"  />
        <asp:ImageButton ID="BISIPNotRegister" runat="server" style="z-index: 1; left: 464px; position: absolute;top: 448px; width: 135px; height: 45px;" ImageUrl="iMages/ISIPNotRegister.png"  />
        <asp:ImageButton ID="BNotUsedColor" runat="server" style="z-index: 1; left: 608px; position: absolute;top: 448px; width: 135px; height: 45px;" ImageUrl="iMages/NotUsedColor.png"  />

        <asp:ImageButton ID="BSPD2EDX" runat="server" style="z-index: 1; left: 176px; position: absolute;top: 496px; width: 135px; height: 45px;" ImageUrl="iMages/SPD2EDX.png"  />
        <asp:ImageButton ID="BSPDMMS2ISIP" runat="server" style="z-index: 1; left: 320px; position: absolute;top: 496px; width: 135px; height: 45px;" ImageUrl="iMages/SPDMMS2ISIP.png"  />
        <asp:ImageButton ID="BColorNoDuplicate" runat="server" style="z-index: 1; left: 464px; position: absolute;top: 496px; width: 135px; height: 45px;" ImageUrl="iMages/ColorNoDuplicate.png"  />
        <asp:ImageButton ID="BOrderAmount" runat="server" style="z-index: 1; left: 608px; position: absolute;top: 496px; width: 135px; height: 45px;" ImageUrl="iMages/OrderAmount.png"  />
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  HISTORY button 45                                                                           ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:ImageButton ID="BAutoDRR" runat="server" style="z-index: 1; left: 176px; position: absolute;top: 608px; width: 135px; height: 45px;" ImageUrl="iMages/AutoLIST.png"  />
        <asp:ImageButton ID="BISOS" runat="server" style="z-index: 1; left: 320px; position: absolute;top: 608px; width: 135px; height: 45px;" ImageUrl="iMages/ISOS2ISIP.png"  />
        <asp:ImageButton ID="BBackupList" runat="server" style="z-index: 1; left: 464px; position: absolute;top: 608px; width: 135px; height: 45px;" ImageUrl="iMages/BackupList.png"  />
        <asp:ImageButton ID="BPullerHistoryList" runat="server" style="z-index: 1; left: 608px; position: absolute;top: 608px; width: 135px; height: 45px;" ImageUrl="iMages/PullerHistoryList.png"  />
    </div>
    </form>
</body>
</html>
