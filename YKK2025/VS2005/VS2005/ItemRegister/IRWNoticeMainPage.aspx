<%@ Page Language="VB" AutoEventWireup="false" CodeFile="IRWNoticeMainPage.aspx.vb" Inherits="IRWNoticeMainPage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>IRW Notice Main Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  ZIP & SLD                                                                             ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <iframe width=500 height=7400 Style="z-index: 103; left:0px; position: absolute; top: 0px" frameborder=0 scrolling=no src="http://10.245.1.6/IRW/IRWNoticePageZip.aspx"></iframe>

        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  實績&推移                                                                           ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->        
        <iframe width=500 height=900 Style="z-index: 103; left:0px; position: absolute; top: 1100px" frameborder=0 scrolling=no src="http://10.245.1.6/IRW/IRWNoticeCount.aspx"></iframe>
        
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  top 10                                                                              ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <iframe width=650 height=2000 Style="z-index: 103; left:510px; position: absolute; top: 0px" frameborder=0 scrolling=no src="http://10.245.1.6/IRW/IRWNoticePage.aspx"></iframe>

        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  無效ITEM-部門+人                                                                           ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <iframe width=1500 height=600 Style="z-index: 103; left:1140px; position: absolute; top: 0px" frameborder=0 scrolling=no src="http://10.245.1.6/IRW/IRWNOActivity.aspx"></iframe>

        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  無效ITEM-TOP-10 客戶 (PER) + (COUNT)                                                                          ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
       <iframe width=1500 height=600 Style="z-index: 103; left:1140px; position: absolute; top: 280px" frameborder=0 scrolling=no src="http://10.245.1.6/IRW/IRWNOActivityCust.aspx?pTop10=1&pFun=PER"></iframe>
       <iframe width=1500 height=600 Style="z-index: 103; left:1140px; position: absolute; top: 480px" frameborder=0 scrolling=no src="http://10.245.1.6/IRW/IRWNOActivityCust.aspx?pTop10=1&pFun=COUNT"></iframe>

        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  無效ITEM-TOP-10 BUYER  (PER) + (COUNT)                                                                 ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <iframe width=1500 height=600 Style="z-index: 103; left:1140px; position: absolute; top: 680px" frameborder=0 scrolling=no src="http://10.245.1.6/IRW/IRWNOActivityBuyer.aspx?pTop10=1&pFun=PER"></iframe>
        <iframe width=1500 height=600 Style="z-index: 103; left:1140px; position: absolute; top: 880px" frameborder=0 scrolling=no src="http://10.245.1.6/IRW/IRWNOActivityBuyer.aspx?pTop10=1&pFun=COUNT"></iframe>

        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  估價使用率                                                                          ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <iframe width=1500 height=600 Style="z-index: 103; left:1140px; position: absolute; top: 1080px" frameborder=0 scrolling=no src="http://10.245.1.6/IRW/IRWPriceUseRate.aspx"></iframe>

    </div>
    </form>
</body>
</html>
