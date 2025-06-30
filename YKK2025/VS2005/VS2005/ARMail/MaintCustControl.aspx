<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MaintCustControl.aspx.vb" Inherits="MaintCustControl" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        客戶／客戶代碼 &nbsp;<asp:TextBox ID="DCust" runat="server" SkinID="txt_Yellow"></asp:TextBox>&nbsp;<br />
        業務員 &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DSalesMan" runat="server" SkinID="txt_Yellow"></asp:TextBox>
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp;<br />
        業務員代號 &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DSalesCode" runat="server" SkinID="txt_Yellow"></asp:TextBox>
        <asp:DropDownList ID="DropDownList1" runat="server" SkinID="DDL_Yellow">
            <asp:ListItem>ALL</asp:ListItem>
            <asp:ListItem Value="1">國內</asp:ListItem>
            <asp:ListItem Value="0">國外</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DropDownList2" runat="server" SkinID="DDL_Yellow">
            <asp:ListItem>ALL</asp:ListItem>
            <asp:ListItem Value="1">是</asp:ListItem>
            <asp:ListItem Value="0">否</asp:ListItem>
        </asp:DropDownList>
        <asp:Button ID="Button1" runat="server" Text="搜尋" />
        <asp:Button ID="Button2" runat="server" Text="新客戶" />
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            PageSize="50" SkinID="GV" Width="95%">
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:Button ID="Button4" runat="server" CommandArgument='<%# Eval("Unique_ID") %>'
                            CommandName="Upt" Text="變更" />
                        <asp:Button ID="Button3" runat="server" CommandArgument='<%# Eval("Unique_ID") %>'
                            CommandName="Del" Text="刪除" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="active" HeaderText="狀態" />
                <asp:TemplateField HeaderText="客戶">
                    <ItemTemplate>
                        <asp:LinkButton ID="LinkButton1" runat="server" CommandArgument='<%# Eval("CustCode") %>'
                            CommandName="Link" Text='<%# Eval("CustName") %>'></asp:LinkButton>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="mailtoname" HeaderText="收件者" />
                <asp:BoundField DataField="salesman" HeaderText="業務員" />
                <asp:BoundField DataField="pdf_create" HeaderText="PDF製作" />
                <asp:BoundField DataField="PDF_RECREATE" HeaderText="PDF再製作" />
                <asp:BoundField DataField="pdf_period" HeaderText="期間" />
                <asp:BoundField DataField="smtp_send" HeaderText="Mail傳送" />
                <asp:BoundField DataField="SMTP_RESEND" HeaderText="Mail再傳送" />
                <asp:BoundField DataField="smtp_period" HeaderText="期間" />
            </Columns>
        </asp:GridView>
    
    </div>
    </form>
</body>
</html>
