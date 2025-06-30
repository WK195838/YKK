<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ModifyDataSheet_02.aspx.vb" Inherits="SPD.ModifyDataSheet_02"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>ModifyDataSheet_02</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<style type="text/css">BODY { FONT-FAMILY: Verdana, Geneva, Sans-Serif; FONT-SIZE: 9pt }
	TABLE { FONT-FAMILY: Verdana, Geneva, Sans-Serif; FONT-SIZE: 9pt }
	TR { FONT-FAMILY: Verdana, Geneva, Sans-Serif; FONT-SIZE: 9pt }
	TD { FONT-FAMILY: Verdana, Geneva, Sans-Serif; FONT-SIZE: 9pt }
	UL { FONT-FAMILY: Verdana, Geneva, Sans-Serif; FONT-SIZE: 9pt }
	LI { FONT-FAMILY: Verdana, Geneva, Sans-Serif; FONT-SIZE: 9pt }
	.normal { FONT-FAMILY: Verdana, Geneva, Sans-Serif; FONT-SIZE: 9pt }
	H1 { FONT-FAMILY: Verdana, Geneva, Sans-Serif; COLOR: #666666; FONT-SIZE: 10.5pt; FONT-WEIGHT: 900 }
	.small { FONT-FAMILY: Verdana, Geneva, Sans-Serif; FONT-SIZE: 7.5pt }
	.error { COLOR: #ff0033 }
	.required { FONT-FAMILY: Verdana, Geneva, Sans-Serif; COLOR: #ff0033; FONT-WEIGHT: 900 }
	.success { MARGIN: 10px 0px; COLOR: #009933; FONT-SIZE: 11pt; FONT-WEIGHT: 900 }
		</style>
		<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
		    function SendToSpec() {
				wPop.document.SpecForm.DSpec.value=document.Form1.DSpec.value;
			}
		    function Button(F, MSG) {
				//alert(F);
				document.cookie="RunBOK=False";
				document.cookie="RunBNG1=False";
				document.cookie="RunBNG2=False";
				document.cookie="RunBSAVE=False";

				answer = confirm("是否執行[" + MSG + "]作業嗎？ 請確認....");
				if (answer) {
					//OK Button
					FOK=document.getElementById("BOK");
					if(FOK!=null) document.Form1.BOK.disabled=true;  	
					//NG-1 Button
					FNG1=document.getElementById("BNG1");
					if(FNG1!=null) document.Form1.BNG1.disabled=true;  	
					//NG-2 Button
					FNG2=document.getElementById("BNG2");
					if(FNG2!=null) document.Form1.BNG2.disabled=true;  	
					//Save Button
					FSAVE=document.getElementById("BSAVE");
					if(FSAVE!=null) document.Form1.BSAVE.disabled=true;  	
						
					if (F=="OK")   document.cookie="RunBOK=True";
					if (F=="NG1")  document.cookie="RunBNG1=True";
					if (F=="NG2")  document.cookie="RunBNG2=True";
					if (F=="SAVE") document.cookie="RunBSAVE=True";
				}
			}
		    function OpenPrintSheet(URL) {
				window.open(URL,'PrintSheet','width=630,height=580,resizable=yes,scrollbars=yes, menubar=yes');
			}
		    function OpenSimulationSheet(URL) {
				window.open(URL,'Simulation','width=630,height=580,resizable=yes,scrollbars=yes, menubar=yes');
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:image id="DModifyDataSheet1" style="Z-INDEX: 102; POSITION: absolute; TOP: 8px; LEFT: 8px"
					runat="server" ImageUrl="Images\ModifyDatatSheet_001.jpg" Width="593px" Height="684px"></asp:image>
				<asp:hyperlink id="LSheet1" style="Z-INDEX: 143; POSITION: absolute; TOP: 224px; LEFT: 528px" runat="server"
					Height="8px" Width="64px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">連結表單</asp:hyperlink><asp:dropdownlist id="DIPerson" style="Z-INDEX: 124; POSITION: absolute; TOP: 664px; LEFT: 440px"
					runat="server" Width="152px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="宋繼明" Selected="True">宋繼明</asp:ListItem>
					<asp:ListItem Value="徐滿霖">徐滿霖</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFDateTime" style="Z-INDEX: 123; POSITION: absolute; TOP: 664px; LEFT: 144px"
					runat="server" Width="152px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DFDateTime</asp:textbox><asp:textbox id="DIContent" style="Z-INDEX: 122; POSITION: absolute; TOP: 560px; LEFT: 144px"
					runat="server" Width="448px" Height="90px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DIContent</asp:textbox><asp:textbox id="DMContent2" style="Z-INDEX: 121; POSITION: absolute; TOP: 432px; LEFT: 144px"
					runat="server" Width="448px" Height="80px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DMContent2</asp:textbox><asp:textbox id="DMContent1" style="Z-INDEX: 120; POSITION: absolute; TOP: 352px; LEFT: 144px"
					runat="server" Width="448px" Height="80px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DMContent1</asp:textbox><asp:dropdownlist id="DWPerson" style="Z-INDEX: 119; POSITION: absolute; TOP: 188px; LEFT: 304px"
					runat="server" Width="162px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">徐滿</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DWDivision" style="Z-INDEX: 118; POSITION: absolute; TOP: 188px; LEFT: 144px"
					runat="server" Width="152px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True">
					<asp:ListItem Value="3">徐滿</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DMPerson" style="Z-INDEX: 117; POSITION: absolute; TOP: 154px; LEFT: 304px"
					runat="server" Width="162px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">徐滿</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DMDivision" style="Z-INDEX: 116; POSITION: absolute; TOP: 154px; LEFT: 144px"
					runat="server" Width="152px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True">
					<asp:ListItem Value="3">徐滿</asp:ListItem>
				</asp:dropdownlist><asp:imagebutton id="BFlow" style="Z-INDEX: 115; POSITION: absolute; TOP: 32px; LEFT: 8px" runat="server"
					ImageUrl="Images\Flow-01.gif" Width="16px" Height="16px"></asp:imagebutton><asp:dropdownlist id="DSheet" style="Z-INDEX: 113; POSITION: absolute; TOP: 222px; LEFT: 144px" runat="server"
					Width="152px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DMReasonType" style="Z-INDEX: 112; POSITION: absolute; TOP: 288px; LEFT: 144px"
					runat="server" Width="152px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Custom" Selected="True">客戶要求</asp:ListItem>
					<asp:ListItem Value="YKK">內部需求</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DComNo" style="Z-INDEX: 111; POSITION: absolute; TOP: 222px; LEFT: 432px" runat="server"
					Width="162px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" MaxLength="20">DComNo</asp:textbox><asp:textbox id="DMReason" style="Z-INDEX: 110; POSITION: absolute; TOP: 320px; LEFT: 16px" runat="server"
					Width="577px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DMReason</asp:textbox><asp:textbox id="DFormSno" style="Z-INDEX: 109; POSITION: absolute; TOP: 696px; LEFT: 16px" runat="server"
					Width="97px" Height="20px" BackColor="White" ForeColor="Blue" BorderStyle="None">單號：123</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 103; POSITION: absolute; TOP: 86px; LEFT: 144px" runat="server"
					Width="152px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DNo</asp:textbox></FONT><asp:textbox id="DDate" style="Z-INDEX: 104; POSITION: absolute; TOP: 86px; LEFT: 432px" runat="server"
				Width="136px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:button id="BDate" style="Z-INDEX: 105; POSITION: absolute; TOP: 86px; LEFT: 568px" runat="server"
				Width="20px" Height="20px" Text="....."></asp:button><asp:dropdownlist id="DStatus" style="Z-INDEX: 106; POSITION: absolute; TOP: 256px; LEFT: 144px" runat="server"
				Width="152px" Height="20px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="開發中" Selected="True">開發中</asp:ListItem>
				<asp:ListItem Value="OK完成">OK完成</asp:ListItem>
				<asp:ListItem Value="NG完成">NG完成</asp:ListItem>
				<asp:ListItem Value="取消完成">取消完成</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 107; POSITION: absolute; TOP: 120px; LEFT: 144px"
				runat="server" Width="152px" Height="20px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="3">IS</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DPerson" style="Z-INDEX: 108; POSITION: absolute; TOP: 120px; LEFT: 432px" runat="server"
				Width="162px" Height="20px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="3">徐滿</asp:ListItem>
			</asp:dropdownlist><asp:image id="DMapSheet3" style="Z-INDEX: 101; POSITION: absolute; TOP: 608px; LEFT: 8px"
				runat="server" ImageUrl="Images\MapSheet_003_A.jpg" Width="593px" Height="77px"></asp:image>
			<asp:imagebutton id="BPrint" style="Z-INDEX: 114; POSITION: absolute; TOP: 8px; LEFT: 8px" runat="server"
				ImageUrl="Images\Print.gif" Width="16px" Height="16px"></asp:imagebutton></form>
	</body>
</HTML>
