<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SufaceAppendSheet_01.aspx.vb" Inherits="SPD.SufaceAppendSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>表面處理委託-追加型別</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			var wPop;
			var val;
		
			function CalendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'CalendarPopup','width=250,height=190,resizable=yes');
			}
			function CodeNoPicker(strField,FormNo)
			{
				window.open('CodeNoPicker.aspx?field=' + strField + '&pFormNo=' + FormNo,'CodeNoPopup','width=250,height=360,resizable=yes');
			}
			//--Spec------------------------------------
			function SpecPicker(strField, Fun) {
			    val=0;
				wPop=window.open('SpecPicker.aspx?field=' + strField + '&fun=' + Fun, 'SpecPopup','width=330,height=250,resizable=yes');
				if (document.Form1.DSpec.value != "") {
					setTimeout("SendToSpec()",200);
				}
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
					//FOK=document.getElementById("BOK");
					//if(FOK!=null) document.Form1.BOK.disabled=true;  	
					//NG-1 Button
					//FNG1=document.getElementById("BNG1");
					//if(FNG1!=null) document.Form1.BNG1.disabled=true;  	
					//NG-2 Button
					//FNG2=document.getElementById("BNG2");
					//if(FNG2!=null) document.Form1.BNG2.disabled=true;  	
					//Save Button
					//FSAVE=document.getElementById("BSAVE");
					//if(FSAVE!=null) document.Form1.BSAVE.disabled=true;  	
						
					if (F=="OK")   document.cookie="RunBOK=True";
					if (F=="NG1")  document.cookie="RunBNG1=True";
					if (F=="NG2")  document.cookie="RunBNG2=True";
					if (F=="SAVE") document.cookie="RunBSAVE=True";
				}
			}

		    function OpenPrintSheet(URL) {
				//alert(URL);
				window.open(URL,'PrintSheet','width=630,height=580,resizable=yes,scrollbars=yes, menubar=yes');
			}
		    function OpenSimulationSheet(URL) {
				window.open(URL,'Simulation','width=630,height=580,resizable=yes,scrollbars=yes, menubar=yes');
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:image id="DAppendSpecSheet" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" Height="510px" Width="766px" ImageUrl="Images\SufaceAppendSheet_001_A.jpg"></asp:image><asp:button id="BCode" style="Z-INDEX: 142; LEFT: 344px; POSITION: absolute; TOP: 162px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button><asp:textbox id="DSpec" style="Z-INDEX: 141; LEFT: 120px; POSITION: absolute; TOP: 228px" runat="server"
					Height="56px" Width="614px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True" TextMode="MultiLine">DSpec</asp:textbox><asp:button id="BSpec" style="Z-INDEX: 109; LEFT: 736px; POSITION: absolute; TOP: 228px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button><asp:dropdownlist id="DQCNeed" style="Z-INDEX: 140; LEFT: 120px; POSITION: absolute; TOP: 376px" runat="server"
					Height="20px" Width="200px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DBuyer" style="Z-INDEX: 139; LEFT: 120px; POSITION: absolute; TOP: 196px" runat="server"
					Height="20px" Width="246px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DBuyer</asp:textbox><asp:hyperlink id="LAttachFile1" style="Z-INDEX: 138; LEFT: 128px; POSITION: absolute; TOP: 488px"
					runat="server" Height="8px" Width="64px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">參考附件</asp:hyperlink><INPUT id="DAttachFile1" style="Z-INDEX: 137; LEFT: 120px; WIDTH: 640px; POSITION: absolute; TOP: 488px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="87" name="File1" runat="server"><asp:textbox id="DSellVendor" style="Z-INDEX: 136; LEFT: 504px; POSITION: absolute; TOP: 196px"
					runat="server" Height="20px" Width="260px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DSellVendor</asp:textbox><asp:textbox id="DOPReadyDesc" style="Z-INDEX: 135; LEFT: 240px; POSITION: absolute; TOP: 672px"
					runat="server" Height="20px" Width="96px" BackColor="White" ForeColor="Red" BorderStyle="None">需閱讀工程履歷</asp:textbox><asp:textbox id="DOPReady" style="Z-INDEX: 134; LEFT: 336px; POSITION: absolute; TOP: 672px"
					runat="server" Height="20px" Width="48px" BackColor="GreenYellow" ForeColor="Red" BorderStyle="Groove" ReadOnly="True"></asp:textbox><asp:imagebutton id="BFlow" style="Z-INDEX: 133; LEFT: 8px; POSITION: absolute; TOP: 32px" runat="server"
					Height="16px" Width="16px" ImageUrl="Images\Flow-01.gif"></asp:imagebutton><asp:imagebutton id="BPrint" style="Z-INDEX: 132; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Height="16px" Width="16px" ImageUrl="Images\Print.gif"></asp:imagebutton><asp:hyperlink id="LONo" style="Z-INDEX: 127; LEFT: 720px; POSITION: absolute; TOP: 162px" runat="server"
					Height="8px" Width="34px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">連結</asp:hyperlink><asp:textbox id="DOFormNo" style="Z-INDEX: 126; LEFT: 408px; POSITION: absolute; TOP: -376px"
					runat="server" Height="20px" Width="72px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" AutoPostBack="True">DOFormNo</asp:textbox><asp:textbox id="DReasonDesc" style="Z-INDEX: 125; LEFT: 168px; POSITION: absolute; TOP: 744px"
					runat="server" Height="59px" Width="424px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" Visible="False" TextMode="MultiLine">DReasonDesc</asp:textbox><asp:textbox id="DReason" style="Z-INDEX: 124; LEFT: 240px; POSITION: absolute; TOP: 712px" runat="server"
					Height="20px" Width="352px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Visible="False">DReason</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 123; LEFT: 168px; POSITION: absolute; TOP: 712px"
					runat="server" Height="20px" Width="64px" BackColor="Gold" ForeColor="Blue" AutoPostBack="True" Visible="False">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LBefOP" style="Z-INDEX: 122; LEFT: 168px; POSITION: absolute; TOP: 672px" runat="server"
					Height="8px" Width="144px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">工程履歷</asp:hyperlink><asp:textbox id="DAEndTime" style="Z-INDEX: 121; LEFT: 440px; POSITION: absolute; TOP: 640px"
					runat="server" Height="20px" Width="152px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" Font-Size="9pt" ReadOnly="True">DAEndTime</asp:textbox><asp:textbox id="DAStartTime" style="Z-INDEX: 120; LEFT: 168px; POSITION: absolute; TOP: 640px"
					runat="server" Height="20px" Width="152px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" Font-Size="9pt" ReadOnly="True">DAStartTime</asp:textbox><asp:textbox id="DBEndTime" style="Z-INDEX: 119; LEFT: 440px; POSITION: absolute; TOP: 608px"
					runat="server" Height="20px" Width="152px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" Font-Size="9pt" ReadOnly="True">DBEndTime</asp:textbox><asp:textbox id="DBStartTime" style="Z-INDEX: 118; LEFT: 168px; POSITION: absolute; TOP: 608px"
					runat="server" Height="20px" Width="152px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" Font-Size="9pt" ReadOnly="True">DBStartTime</asp:textbox><asp:textbox id="DDecideDesc" style="Z-INDEX: 117; LEFT: 56px; POSITION: absolute; TOP: 528px"
					runat="server" Height="56px" Width="536px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:image id="DDelay" style="Z-INDEX: 116; LEFT: 8px; POSITION: absolute; TOP: 704px" runat="server"
					Height="107px" Width="593px" ImageUrl="Images\Sheet_Delay.jpg" Visible="False"></asp:image><asp:image id="DDelivery" style="Z-INDEX: 115; LEFT: 8px; POSITION: absolute; TOP: 592px" runat="server"
					Height="110px" Width="593px" ImageUrl="Images\Sheet_Delivery.jpg"></asp:image><asp:image id="DDescSheet" style="Z-INDEX: 114; LEFT: 8px; POSITION: absolute; TOP: 520px"
					runat="server" Height="75px" Width="593px" ImageUrl="Images\MapSheet_004.jpg"></asp:image><asp:textbox id="DFormSno" style="Z-INDEX: 113; LEFT: 8px; POSITION: absolute; TOP: 816px" runat="server"
					Height="20px" Width="97px" BackColor="White" ForeColor="Blue" BorderStyle="None">單號：123</asp:textbox>
				<asp:textbox id="DQCRemark" style="Z-INDEX: 112; LEFT: 120px; POSITION: absolute; TOP: 408px"
					runat="server" Height="56px" Width="640px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"
					TextMode="MultiLine">DQCRemark</asp:textbox><asp:textbox id="DAppendReason" style="Z-INDEX: 111; LEFT: 120px; POSITION: absolute; TOP: 296px"
					runat="server" Height="56px" Width="640px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DAppendReason</asp:textbox><asp:textbox id="DOFormSno" style="Z-INDEX: 110; LEFT: 520px; POSITION: absolute; TOP: -376px"
					runat="server" Height="20px" Width="86px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" AutoPostBack="True">DOFormSno</asp:textbox><asp:textbox id="DCode" style="Z-INDEX: 108; LEFT: 120px; POSITION: absolute; TOP: 162px" runat="server"
					Height="20px" Width="224px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DCode</asp:textbox><asp:textbox id="DONo" style="Z-INDEX: 107; LEFT: 504px; POSITION: absolute; TOP: 162px" runat="server"
					Height="20px" Width="210px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DONo</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 106; LEFT: 120px; POSITION: absolute; TOP: 96px" runat="server"
					Height="20px" Width="246px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DNo</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 103; LEFT: 504px; POSITION: absolute; TOP: 96px" runat="server"
					Height="20px" Width="240px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:button id="BDate" style="Z-INDEX: 105; LEFT: 744px; POSITION: absolute; TOP: 96px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button><asp:dropdownlist id="DPerson" style="Z-INDEX: 102; LEFT: 504px; POSITION: absolute; TOP: 128px" runat="server"
					Height="20px" Width="260px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 101; LEFT: 120px; POSITION: absolute; TOP: 128px"
					runat="server" Height="20px" Width="246px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist></FONT><INPUT id="BSAVE" style="Z-INDEX: 128; LEFT: 256px; WIDTH: 80px; POSITION: absolute; TOP: 816px; HEIGHT: 28px"
				onclick="Button('SAVE');" type="button" value="儲存" name="BSAVE" runat="server"><INPUT id="BNG2" style="Z-INDEX: 129; LEFT: 344px; WIDTH: 80px; POSITION: absolute; TOP: 816px; HEIGHT: 28px"
				onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 130; LEFT: 432px; WIDTH: 80px; POSITION: absolute; TOP: 816px; HEIGHT: 28px"
				onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><INPUT id="BOK" style="Z-INDEX: 131; LEFT: 520px; WIDTH: 80px; POSITION: absolute; TOP: 816px; HEIGHT: 28px"
				onclick="Button('OK');" type="button" value="OK" name="BOK" runat="server"></form>
	</body>
</HTML>
