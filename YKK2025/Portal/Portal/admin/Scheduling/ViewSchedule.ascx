<%@ Control CodeBehind="ViewSchedule.ascx.vb" language="vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.Admin.Scheduling.ViewSchedule" %>
<asp:datagrid id="dgSchedule" runat="server" autogeneratecolumns="false" cellpadding="4" datakeyfield="ScheduleID" enableviewstate="false" border="0" summary="This table shows the schedule of events for the portal." BorderStyle="None" BorderWidth="0px" GridLines="None">
<Columns>
<asp:TemplateColumn>
<ItemTemplate>
                <asp:hyperlink id="editLink" NavigateUrl='<%# EditURL("ScheduleID",DataBinder.Eval(Container.DataItem,"ScheduleID").ToString()) %>' Visible="<%# IsEditable %>" runat="server">
                    <asp:IMAGE id="editLinkImage" ImageUrl="~/images/edit.gif" Visible="<%# IsEditable %>" AlternateText="Edit" runat="server" resourcekey="Edit"/>
                </asp:hyperlink>
            
</ItemTemplate>
</asp:TemplateColumn>
<asp:BoundColumn DataField="TypeFullName" HeaderText="Type">
<HeaderStyle CssClass="NormalBold">
</HeaderStyle>

<ItemStyle CssClass="Normal">
</ItemStyle>
</asp:BoundColumn>
<asp:TemplateColumn HeaderText="Enabled">
<HeaderStyle CssClass="NormalBold">
</HeaderStyle>

<ItemStyle HorizontalAlign="Center">
</ItemStyle>

<ItemTemplate>
                <asp:IMAGE id="Image1" ImageUrl="~/images/checked.gif" Visible='<%# DataBinder.Eval(Container.DataItem,"Enabled")="true" %>' AlternateText="Enabled" runat="server" resourcekey="Enabled.Header" />
            
</ItemTemplate>
</asp:TemplateColumn>
<asp:TemplateColumn HeaderText="Frequency">
<HeaderStyle CssClass="NormalBold">
</HeaderStyle>

<ItemStyle CssClass="Normal">
</ItemStyle>

<ItemTemplate>
                <%# GetTimeLapse(DataBinder.Eval(Container.DataItem,"TimeLapse"), DataBinder.Eval(Container.DataItem,"TimeLapseMeasurement")) %>
            
</ItemTemplate>
</asp:TemplateColumn>
<asp:TemplateColumn HeaderText="RetryTimeLapse">
<HeaderStyle CssClass="NormalBold">
</HeaderStyle>

<ItemStyle CssClass="Normal">
</ItemStyle>

<ItemTemplate>
                <%# GetTimeLapse(DataBinder.Eval(Container.DataItem,"RetryTimeLapse"), DataBinder.Eval(Container.DataItem,"RetryTimeLapseMeasurement")) %>
            
</ItemTemplate>
</asp:TemplateColumn>
<asp:TemplateColumn HeaderText="NextStart">
<HeaderStyle CssClass="NormalBold">
</HeaderStyle>

<ItemStyle HorizontalAlign="Center">
</ItemStyle>

<ItemTemplate>
                <asp:label CssClass="Normal" ID="lblNextStart" Runat="server" Visible='<%# DataBinder.Eval(Container.DataItem,"Enabled")="true" %>'><%# DataBinder.Eval(Container.DataItem,"NextStart")%></asp:label>
            
</ItemTemplate>
</asp:TemplateColumn>
<asp:TemplateColumn>
<ItemTemplate>
                <asp:hyperlink CssClass="CommandButton" id="lnkHistory" resourcekey="lnkHistory" NavigateUrl='<%# EditURL("ScheduleID",DataBinder.Eval(Container.DataItem,"ScheduleID").ToString(), "History") %>' runat="server">History</asp:hyperlink>
            
</ItemTemplate>
</asp:TemplateColumn>
</Columns>
</asp:datagrid>
