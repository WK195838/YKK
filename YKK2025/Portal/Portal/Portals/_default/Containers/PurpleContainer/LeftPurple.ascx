<%@ Control language="vb" CodeBehind="~/admin/Containers/container.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Containers.Container" %>
<%@ Register TagPrefix="dnn" TagName="VISIBILITY" Src="~/Admin/Containers/Visibility.ascx" %>
<%@ Register TagPrefix="dnn" TagName="ACTIONS" Src="~/Admin/Containers/SolPartActions.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TITLE" Src="~/Admin/Containers/Title.ascx" %>
<TABLE WIDTH="100%" BORDER=0 align="center" CELLPADDING=0 CELLSPACING=0>
	<TR>
	  <TD HEIGHT=18 colspan="3"><img src="<%= SkinPath %>spacer.gif" width="1" height="18"></TD>
  </TR>
	<TR>
		<TD WIDTH=8 rowspan="4" align="right" background="<%= SkinPath %>03_04b.gif"><img src="<%= SkinPath %>spacer.gif" width="1" height="1"></TD>
		<TD BACKGROUND="<%= SkinPath %>03_04.jpg" HEIGHT=19 ALIGN=LEFT VALIGN=TOP></TD>
		<TD width="8" rowspan="4" ALIGN=LEFT VALIGN=top background="<%= SkinPath %>03_04b2.gif"></TD>
	</TR>
	<TR>
	  <TD ALIGN=LEFT VALIGN=TOP><table width="100%"  border="0" cellspacing="0" cellpadding="5">
        <tr>
          <td><TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
            <TR>
              <TD COLSPAN=3></TD>
            </TR>
            <TR>
              <TD VALIGN=TOP><dnn:VISIBILITY runat="server" id="dnnVISIBILITY" /></TD>
              <TD>&nbsp;&nbsp;<dnn:ACTIONS runat="server" id="dnnACTIONS" /></TD>
              <TD ALIGN=LEFT VALIGN=MIDDLE>&nbsp;&nbsp;<dnn:TITLE runat="server" id="dnnTITLE" /></TD>
            </TR>
          </TABLE></td>
        </tr>
        <tr>
          <td id="ContentPane" runat="server" ALIGN=LEFT VALIGN=TOP></td>
        </tr>
      </table></TD>
	</TR>
	<TR>
	  <TD ALIGN=LEFT VALIGN=TOP>
			<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
				<TR><TD></TD></TR>
	  </TABLE>	  </TD>
	</TR>
	<TR>
		<TD height="1" background="<%= SkinPath %>03_04bb.gif"><img src="<%= SkinPath %>spacer.gif" width="1" height="1"></TD>
	</TR>
</TABLE>


