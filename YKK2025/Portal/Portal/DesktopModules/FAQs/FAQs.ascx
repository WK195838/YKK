<%@ Control Inherits="DotNetNuke.Modules.FAQs.FAQs" CodeBehind="FAQs.ascx.vb" language="vb" AutoEventWireup="false" Explicit="True" %>
<asp:datalist id="lstFAQs" datakeyfield="ItemID" runat="server" cellpadding="4">
  <itemtemplate>
    <table cellpadding="4" width="100%">
      <tr>
        <td valign="top" width="100%" align="left">
          <asp:HyperLink NavigateUrl='<%# EditURL("ItemID",DataBinder.Eval(Container.DataItem,"ItemID")) %>' Visible="<%# IsEditable %>" runat="server" ID="Hyperlink1"><asp:Image ID=Hyperlink1Image Runat=server  ImageUrl="~/images/edit.gif" AlternateText="Edit" Visible="<%#IsEditable%>" resourcekey="Edit"/></asp:hyperlink>
          <asp:label id="Label2" resourcekey="Q" class="SubHead" runat="server">Q.</asp:Label>&nbsp;
          <asp:linkbutton id="Q" Text='<%# HtmlDecode(DataBinder.Eval(Container.DataItem, "question")) %>' CommandName="Select" class="SubHead" runat="server"/>
          <asp:panel id="pnl" visible="false" runat="server">
						<asp:label id="Label1" resourcekey="A" class="SubHead" runat="server">A.</asp:Label>&nbsp;
						<asp:label id="A" Text='<%# HtmlDecode(DataBinder.Eval(Container.DataItem, "answer")) %>' class="Normal" runat="server"/>
					</asp:panel>
        </td>
      </tr>
    </table>
  </itemtemplate>
</asp:datalist>
