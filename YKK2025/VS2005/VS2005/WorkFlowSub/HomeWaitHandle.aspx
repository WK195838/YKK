<%@ Page Language="VB" AutoEventWireup="false" CodeFile="HomeWaitHandle.aspx.vb" Inherits="HomeWaitHandle" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>待處理</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
            <table border="0" style="Top: 0px; left:6px; position: absolute;">
                <tr>
                    <td style="width: 60px; color: black; text-align: center;" >
                        <asp:TextBox ID="DHighTitle" runat="server" BackColor="White" BorderStyle="None"
                            Font-Bold="False" Font-Size="10pt" ForeColor="#0000C0" Width="60px"  Height="15px" Style="text-align:center">待處理件</asp:TextBox></td>
                            
                    <td style="width: 50px; color: blue " >
                        <asp:TextBox ID="DHighCount" runat="server" BackColor="White" BorderStyle="None"
                            ForeColor="Red" Width="50px" Height="15px" Style="text-align:Left" Font-Size="10pt">11</asp:TextBox></td>
                </tr>
            </table>
        <asp:HyperLink ID="DLabel1" style="Top: 26px; left:6px; position: absolute;text-align: right;" 
        Font-Size="10pt" runat="server" Font-Underline="False" NavigateUrl="http://LocalHost/Portal/工作流程/待處理/tabid/90/Default.aspx" Target="_parent" Width="126px">請馬上處理(連結)</asp:HyperLink>
        


        <asp:HyperLink style="LEFT: 6px; POSITION: absolute; TOP: 42px; TEXT-ALIGN: right" id="DLabel2" runat="server" Width="126px" 
        Font-Size="10pt" Target="_parent" NavigateUrl="http://LocalHost/portal/管理/經費系統/經費批簽/tabid/444/Default.aspx" Font-Underline="False">...</asp:HyperLink>    
        
        <asp:HyperLink style="LEFT: 6px; POSITION: absolute; TOP: 58px; TEXT-ALIGN: right" id="DLabel3" runat="server" Width="126px" 
        Font-Size="10pt" Target="_parent" NavigateUrl="http://LocalHost/portal/管理/經費系統/經費批簽/tabid/444/Default.aspx" Font-Underline="False">...</asp:HyperLink>    
        
        <asp:HyperLink style="LEFT: 6px; POSITION: absolute; TOP: 74px; TEXT-ALIGN: right" id="DLabel4" runat="server" Width="126px" 
        Font-Size="10pt" Target="_parent" NavigateUrl="http://LocalHost/portal/管理/經費系統/經費批簽/tabid/444/Default.aspx" Font-Underline="False">...</asp:HyperLink>    
        
        <asp:HyperLink style="LEFT: 6px; POSITION: absolute; TOP: 90px; TEXT-ALIGN: right" id="DLabel5" runat="server" Width="126px" 
        Font-Size="10pt" Target="_parent" NavigateUrl="http://LocalHost/portal/管理/經費系統/經費批簽/tabid/444/Default.aspx" Font-Underline="False">...</asp:HyperLink>    
        
        <asp:HyperLink style="LEFT: 6px; POSITION: absolute; TOP: 106px; TEXT-ALIGN: right" id="DLabel6" runat="server" Width="126px" 
        Font-Size="10pt" Target="_parent" NavigateUrl="http://LocalHost/portal/管理/經費系統/經費批簽/tabid/444/Default.aspx" Font-Underline="False">...</asp:HyperLink>    
                
        </div>
        

    </form>
</body>
</html>
