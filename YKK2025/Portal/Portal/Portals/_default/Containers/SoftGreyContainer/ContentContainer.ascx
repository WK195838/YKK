<%@ Control language="vb" CodeBehind="~/admin/Containers/container.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Containers.Container" %>
<%@ Register TagPrefix="dnn" TagName="VISIBILITY" Src="~/Admin/Containers/Visibility.ascx" %>
<%@ Register TagPrefix="dnn" TagName="ACTIONS" Src="~/Admin/Containers/SolPartActions.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TITLE" Src="~/Admin/Containers/Title.ascx" %>
<TABLE WIDTH="100%" BORDER=0 align="center" CELLPADDING=0 CELLSPACING=0>
	<TR>
	  <TD HEIGHT=18 colspan="3"><img src="<%= SkinPath %>spacer.gif" width="1" height="12"></TD>
  </TR>
	<TR>
		<TD WIDTH=8 align="right"><img src="<%= SkinPath %>01_01.jpg" width="8" height="54"></TD>
		<TD BACKGROUND="<%= SkinPath %>03_03.jpg" HEIGHT=25 ALIGN=LEFT VALIGN=TOP><img src="<%= SkinPath %>02_02.jpg" width="24" height="54"></TD>
		<TD ALIGN=LEFT VALIGN=top><img src="<%= SkinPath %>04_04.jpg" width="9" height="54"></TD>
	</TR>
	<TR>
	  <TD BACKGROUND="<%= SkinPath %>05_05.jpg"></TD>
	  <TD ALIGN=LEFT VALIGN=TOP><TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
				<TR><TD COLSPAN=3></TD></TR>
				<TR><TD VALIGN=TOP><dnn:VISIBILITY runat="server" id="dnnVISIBILITY" /></TD><TD><dnn:ACTIONS runat="server" id="dnnACTIONS" /></TD><TD ALIGN=LEFT VALIGN=MIDDLE><dnn:TITLE runat="server" id="dnnTITLE" /></TD></TR>
	  </TABLE></TD>
	  <TD BACKGROUND="<%= SkinPath %>07_07.jpg"></TD>
  </TR>
	<TR>
		<TD BACKGROUND="<%= SkinPath %>05_05.jpg"></TD>
		<TD ALIGN=LEFT VALIGN=TOP>
			<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
				<TR><TD id="ContentPane" runat="server" ALIGN=LEFT VALIGN=TOP></TD></TR>
	  </TABLE>	  </TD>
		<TD BACKGROUND="<%= SkinPath %>07_07.jpg" WIDTH=12></TD>
	</TR>
	<TR>
		<TD WIDTH=8 HEIGHT=12 valign="top" ALT=""><img src="<%= SkinPath %>08_08.jpg" width="8" height="11"></TD>
		<TD BACKGROUND="<%= SkinPath %>09_09.jpg"></TD>
		<TD WIDTH=12 HEIGHT=12 valign="top" ALT=""><img src="<%= SkinPath %>10_10.jpg" width="9" height="11"></TD>
	</TR>
</TABLE>


