<%@ Page Language="VB" AutoEventWireup="false" CodeFile="RemarkListCL.aspx.vb" Inherits="RemarkListCL" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" ><head id="Head1" runat="server">
    <title>經費資料</title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <!-- -->
            <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
            <!-- ++  底圖                                                                                ++ -->
            <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
            <img src="images/Remark.jpg" style="z-index: 1; left: 10px; position: absolute;top: 34px; width: 586px; height: 319px;"/>
            <asp:Button ID="BDEL" runat="server" ForeColor="Red" Style="z-index: 101; left: 432px;
                position: absolute; top: 320px" Text="修改" Width="72px" />

            <asp:DropDownList style="Z-INDEX: 134; LEFT: 20px; POSITION: absolute; TOP: 64px" id="DCat" runat="server" Width="136px" Height="24px" BackColor="Yellow">
            </asp:DropDownList>
            
            <asp:TextBox ID="DContentList" runat="server" BackColor="Yellow" BorderStyle="None" ForeColor="Black"
                Height="80px" MaxLength="7" Style="z-index: 126; left: 160px; position: absolute;
                top: 64px; text-align: left" TextMode="MultiLine" Width="424px" Wrap="False"></asp:TextBox>
                
            <asp:TextBox ID="DRemarkList" runat="server" BackColor="Yellow" BorderStyle="None" ForeColor="Black"
                Height="96px" MaxLength="7" Style="z-index: 126; left: 16px; position: absolute;
                top: 216px; text-align: left" TextMode="MultiLine" Width="568px"></asp:TextBox>
           
            <asp:Button ID="BAdd" runat="server" Style="z-index: 101; left: 514px; position: absolute;
                top: 160px" Text="新追加" Width="72px" />
            <asp:Button ID="BClose" runat="server" Style="z-index: 101; left: 512px; position: absolute;
                top: 320px" Text="完成" Width="72px" />
            <asp:TextBox ID="DChk" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
                Height="24px" Style="z-index: 104; left: 952px; position: absolute; top: 104px"
                Width="142px"></asp:TextBox>

             </div>
    </form>
</body>
</html>