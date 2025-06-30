<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ViewMaintHistory.aspx.vb" Inherits="ViewMaintHistory" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="TextBox1" runat="server" SkinID="txt_Yellow"></asp:TextBox>
        <asp:Button ID="Button1" runat="server" SkinID="txt_Yellow" Text="搜尋" />
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            PageSize="50" SkinID="GV" Width="90%">
            <Columns>
                <asp:BoundField DataField="ProcProgram" HeaderText="處理程式" />
                <asp:TemplateField HeaderText="處理時間">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("ProcTime") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# CDate(Eval("ProcTime")).tostring("yyyy/MM/dd HH:mm:ss") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="ProcUser" HeaderText="處理者" />
                <asp:BoundField DataField="Function" HeaderText="動作" />
                <asp:BoundField DataField="Before" HeaderText="動作前" />
                <asp:BoundField DataField="After" HeaderText="動作後" />
            </Columns>
        </asp:GridView>
    
    </div>
    </form>
</body>
</html>
