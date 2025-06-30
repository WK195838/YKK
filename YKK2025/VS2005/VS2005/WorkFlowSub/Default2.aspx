<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default2.aspx.vb" Inherits="Default2" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
            <asp:Panel ID="Panel1" runat="server" Style="left: 0px; position: relative; top: 0px" Width="416px">
          <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="~/Images/WorkRatioInfor_B1.jpg" />
          <asp:ImageButton ID="ImageButton2" runat="server" ImageUrl="~/Images/WorkRatioInfor_B2.jpg" style="left: -6px; position: relative; top: 0px" />
          <asp:ImageButton ID="ImageButton3" runat="server" ImageUrl="~/Images/WorkRatioInfor_B3.jpg" style="left: -14px; position: relative; top: 0px" />
        </asp:Panel>

<asp:MultiView ID="MultiView1" runat="server" ActiveViewIndex="0">
  <asp:View ID="View1" runat="server">
    View 1<asp:DropDownList ID="DropDownList1" runat="server">
          <asp:ListItem>1</asp:ListItem>
          <asp:ListItem>2</asp:ListItem>
          <asp:ListItem>3</asp:ListItem>
          <asp:ListItem>4</asp:ListItem>
      </asp:DropDownList><br />
    <br />
    <asp:Button ID="Button1" runat="server" 
      CommandArgument="View2" 
      CommandName="SwitchViewByID"
      Text="Go to View2" />
  </asp:View>
  <asp:View ID="View2" runat="server">
    View 2<br />
    <br />
    <asp:Button ID="Button2" runat="server" 
      CommandArgument="View3" 
      CommandName="SwitchViewByID"
      Text="Go to View 3" />
  </asp:View>
  <asp:View ID="View3" runat="server">
    View 3<br />
    <br />
    <asp:Button ID="Button3" runat="server" 
      CommandArgument="View1" 
      CommandName="SwitchViewByID"
      Text="Go to View 1" />
  </asp:View>
</asp:MultiView>    
    </div>
    </form>
</body>
</html>
