<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        FormNo:<asp:DropDownList ID="DropDownList1" runat="server">
            <asp:ListItem  Value="001151">Item登錄單</asp:ListItem>
            <asp:ListItem Value="001152">ZIP登錄單</asp:ListItem>
            <asp:ListItem Value="001153">SLD登錄單</asp:ListItem>
            <asp:ListItem Value="001154">CH登錄單</asp:ListItem>
            <asp:ListItem Value="001155">SLD-工廠登錄單</asp:ListItem>
            <asp:ListItem  Value="001161">單價承認委託單</asp:ListItem>
            <asp:ListItem Selected="True" Value="001101">案件委託</asp:ListItem>
        </asp:DropDownList>
        FormSNo<asp:TextBox ID="TextBox2" runat="server" Width="95px">2</asp:TextBox>Step<asp:TextBox ID="TextBox3" runat="server"  Width="43px">10</asp:TextBox>
        SeqNo<asp:TextBox ID="TextBox4" runat="server" Width="66px">1</asp:TextBox>
        ApplyID<asp:TextBox ID="TextBox5" runat="server"  Width="67px">it013</asp:TextBox>
        UserID<asp:TextBox ID="TextBox6" runat="server"  Width="67px">it013</asp:TextBox>
        <br />
        <asp:Button ID="ItemRegister" runat="server" Text="Item登錄申請單" />
        <asp:Button ID="ItemRegister2" runat="server" Text="Item登錄申請單_02" />
        <asp:Button ID="ItemRegisterHistory" runat="server" Text="履歷" />
        <br />
        <asp:Button ID="ZIPRegister" runat="server" Text="ZIP登錄申請單" />
        <asp:Button ID="ZIPRegister2" runat="server" Text="ZIP登錄申請單_02" />&nbsp;
        <asp:Button ID="ZIPRegisterHistory" runat="server" Text="履歷" />
        <br />
        <asp:Button ID="SLDRegister" runat="server" Text="SLD登錄申請單" />
        <asp:Button ID="SLDRegister2" runat="server" Text="SLD登錄申請單_02" />&nbsp;
        <asp:Button ID="SLDRegisterHistory" runat="server" Text="履歷" />
        <br />
        <asp:Button ID="CHRegister" runat="server" Text="CH登錄申請單" />
        <asp:Button ID="CHRegister2" runat="server" Text="CH登錄申請單_02" />&nbsp;
        <asp:Button ID="CHRegisterHistory" runat="server" Text="履歷" />
        <br />
        <asp:Button ID="FSLDRegister" runat="server" Text="SLD-工廠登錄申請單" />
        <asp:Button ID="FSLDRegister2" runat="server" Text="SLD-工廠登錄申請單_02" />&nbsp;
        <asp:Button ID="FSLDRegisterHistory" runat="server" Text="履歷" />
        <br />
        <asp:Button ID="Button18" runat="server" Style="position: relative; z-index: 100;" Text="單價計算01" />
        <asp:Button ID="Button19" runat="server" Style="position: relative; z-index: 101;" Text="單價計算02" /><br />
        <asp:Button ID="Inq_Commission" runat="server" Text="調閱資料" />
        <br />
        <asp:Button ID="CaseReviewSheet_01" runat="server" Text="案件委託_01" />
        <asp:Button ID="Button1" runat="server" Text="FAS_01" style="z-index: 103; left: 13px; position: absolute; top: 313px" /><asp:Button ID="Button2" runat="server" Text="FAS_02" style="z-index: 103; left: 97px; position: absolute; top: 314px" /><asp:Button ID="CaseReviewSheet_02"
            runat="server" Text="案件委託_02" /><asp:Button ID="CaseReviewSheet_H" runat="server"
                Text="履歷" /></div>
        
        
    </form>
</body>
</html>
