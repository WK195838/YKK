<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MaintItem.aspx.vb" Inherits="_MaintItem" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
    
		<script language="javascript" type="text/javascript">
            function IsConfirm(msg) {
                return confirm('是否確定' + msg + '?');
            }
		</script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="BID" runat="server" Style="z-index: 100; left: 127px; position: absolute;
            top: 97px" BackColor="LightGray"></asp:TextBox>
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/ControlTable.jpg" />
        <asp:TextBox ID="BDATA" runat="server" Style="z-index: 100; left: 126px; position: absolute;
            top: 195px" BackColor="GreenYellow" Width="548px"></asp:TextBox>
        <asp:TextBox ID="BDKEY" runat="server" Style="z-index: 100; left: 126px; position: absolute;
            top: 163px" BackColor="GreenYellow" Width="279px"></asp:TextBox>
        <asp:TextBox ID="BCAT" runat="server" Style="z-index: 100; left: 126px; position: absolute;
            top: 131px" BackColor="LightGray"></asp:TextBox>
        <input id="BBack" runat="server" style="z-index: 145; left: 521px;
            width: 80px; position: absolute; top: 228px; height: 28px" type="button" value="回上一頁" />
        <input id="BSave" runat="server" style="z-index: 145; left: 605px;
            width: 80px; position: absolute; top: 228px; height: 28px" type="button" value="儲存" />
    </div>
    </form>
</body>
</html>
