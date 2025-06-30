<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SliderImageMain.aspx.vb" Inherits="SliderImageMain" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Image Portal Main</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/ImagePortal_11.jpg" style="z-index: 1; left: 4px; position: absolute;top: 8px; width: 915px; height: 455px;" />
        <img src="iMages/ImagePortal_12.jpg" style="z-index: 1; left: 2px; position: absolute;top: 463px; width: 915px; height: 455px;" />

        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  各Action按鈕                                                                        ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:ImageButton ID="BRDImagePuller" runat="server" style="z-index: 1; left: 64px; position: absolute;top: 96px;" ImageUrl="~/iMages/RDImage_Puller.png" Height="45px" Width="170px" />
        <asp:ImageButton ID="BSUPPLIERImagePuller" runat="server" style="z-index: 1; left: 360px; position: absolute;top: 96px;" ImageUrl="~/iMages/SUPPLIERImage_Puller.png"  Height="45px" Width="170px" />
        <asp:ImageButton ID="BSLDImagePuller" runat="server" style="z-index: 1; left: 648px; position: absolute;top: 96px;" ImageUrl="~/iMages/SLDImage_Puller.png"  Height="45px" Width="170px" />

        <asp:ImageButton ID="BSLDImageITEM" runat="server" style="z-index: 1; left: 64px; position: absolute;top: 480px;" ImageUrl="~/iMages/SLDImage_ITEM.png"  Height="45px" Width="170px" />
        <asp:ImageButton ID="BGRImage" runat="server" style="z-index: 1; left: 360px; position: absolute;top: 480px;" ImageUrl="~/iMages/GRImage_01.png"  Height="45px" Width="170px" />
        <asp:ImageButton ID="BTWN2YOC" runat="server" style="z-index: 1; left: 648px; position: absolute;top: 480px;" ImageUrl="~/iMages/TWN2YOC_01.png"  Height="45px" Width="170px" />
    
    </div>
    </form>
</body>
</html>
