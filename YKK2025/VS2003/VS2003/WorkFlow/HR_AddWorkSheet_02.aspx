<%@ Page Language="vb" AutoEventWireup="false" Codebehind="HR_AddWorkSheet_02.aspx.vb" Inherits="SPD.HR_AddWorkSheet_02"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>補工申請書</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			function ShowCardTime()
			{
				window.open('HR_CardTimeList.aspx?pWorkDate=' + document.Form1.DWorkDate.value + '&pDepoID=' + document.Form1.DDepoCode.value + '&pEmpID=' + document.Form1.DEmpID.value, 'CardTimeSheet','width=290,height=150,resizable=yes');
			}

			function ShowCardTimeA()
			{
				window.open('HR_CardTimeList.aspx?pWorkDate=' + document.Form1.DAWorkDate.value + '&pDepoID=' + document.Form1.DDepoCode.value + '&pEmpID=' + document.Form1.DEmpID.value, 'CardTimeSheet','width=290,height=150,resizable=yes');
			}

			function ShowOverTime()
			{
				window.open('http://10.245.1.10/WorkFlowSub/HR_OverTimeList.aspx?pMonth=' + document.Form1.DSalaryYM.value + '&pDepoID=' + document.Form1.DDepoCode.value + '&pEmpID=' + document.Form1.DEmpID.value, 'OverTimeSheet','width=320,height=400,resizable=yes');
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
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:textbox id="DWM" style="Z-INDEX: 110; LEFT: 600px; POSITION: absolute; TOP: 186px; TEXT-ALIGN: right"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove"
					ReadOnly="True">DWM</asp:textbox>
				<asp:textbox id="DAddWorkType" style="Z-INDEX: 131; LEFT: 456px; POSITION: absolute; TOP: 152px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="120px" Height="20px" BackColor="Yellow"
					ForeColor="Blue">DAddWorkType</asp:textbox>
				<asp:textbox id="DAEndM" style="Z-INDEX: 130; LEFT: 392px; POSITION: absolute; TOP: 250px; TEXT-ALIGN: right"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow"
					ForeColor="Blue">DAEndM</asp:textbox>
				<asp:textbox id="DAStartM" style="Z-INDEX: 126; LEFT: 206px; POSITION: absolute; TOP: 250px; TEXT-ALIGN: right"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow"
					ForeColor="Blue">DAStartM</asp:textbox>
				<asp:textbox id="DAEndH" style="Z-INDEX: 127; LEFT: 310px; POSITION: absolute; TOP: 250px; TEXT-ALIGN: right"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow"
					ForeColor="Blue">DAEndH</asp:textbox>
				<asp:textbox id="DAStartH" style="Z-INDEX: 129; LEFT: 120px; POSITION: absolute; TOP: 250px; TEXT-ALIGN: right"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow"
					ForeColor="Blue">DAStartH</asp:textbox>
				<asp:textbox id="DWEndM" style="Z-INDEX: 128; LEFT: 392px; POSITION: absolute; TOP: 185px; TEXT-ALIGN: right"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow"
					ForeColor="Blue">DWEndM</asp:textbox>
				<asp:textbox id="DWEndH" style="Z-INDEX: 124; LEFT: 306px; POSITION: absolute; TOP: 185px; TEXT-ALIGN: right"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow"
					ForeColor="Blue">DWEndH</asp:textbox>
				<asp:textbox id="DWStartM" style="Z-INDEX: 125; LEFT: 206px; POSITION: absolute; TOP: 185px; TEXT-ALIGN: right"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow"
					ForeColor="Blue">DWStartM</asp:textbox>
				<asp:textbox id="DWStartH" style="Z-INDEX: 123; LEFT: 120px; POSITION: absolute; TOP: 185px; TEXT-ALIGN: right"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow"
					ForeColor="Blue">DWStartH</asp:textbox>
				<asp:textbox id="DName" style="Z-INDEX: 121; LEFT: 456px; POSITION: absolute; TOP: 88px" runat="server"
					ReadOnly="True" BorderStyle="Groove" Width="160px" Height="20px" BackColor="Yellow" ForeColor="Blue">DName</asp:textbox>
				<asp:button id="BACardTime" style="Z-INDEX: 120; LEFT: 248px; POSITION: absolute; TOP: 218px"
					runat="server" Width="94px" Height="24px" BackColor="#FFFFC0" Text="刷卡記錄" CausesValidation="False"></asp:button><asp:textbox id="DAWorkDate" style="Z-INDEX: 119; LEFT: 120px; POSITION: absolute; TOP: 218px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="104px" BorderStyle="Groove" ReadOnly="True">2008/10/10</asp:textbox><asp:button id="BOverTime" style="Z-INDEX: 118; LEFT: 584px; POSITION: absolute; TOP: 152px"
					runat="server" BackColor="#FFFFC0" Height="24px" Width="96px" Text="加班記錄" CausesValidation="False"></asp:button><asp:button id="BCardTime" style="Z-INDEX: 117; LEFT: 248px; POSITION: absolute; TOP: 152px"
					runat="server" BackColor="#FFFFC0" Height="24px" Width="94px" Text="刷卡記錄" CausesValidation="False"></asp:button><asp:textbox id="DDepoCode" style="Z-INDEX: 116; LEFT: 496px; POSITION: absolute; TOP: 120px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="25px" BorderStyle="Groove" ReadOnly="True">01</asp:textbox><asp:textbox id="DDepoName" style="Z-INDEX: 115; LEFT: 456px; POSITION: absolute; TOP: 120px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True">台北</asp:textbox><asp:textbox id="DSalaryYM" style="Z-INDEX: 114; LEFT: 224px; POSITION: absolute; TOP: 88px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="64px" BorderStyle="Groove" ReadOnly="True">2008/10</asp:textbox><asp:textbox id="DJobCode" style="Z-INDEX: 113; LEFT: 224px; POSITION: absolute; TOP: 120px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="64px" BorderStyle="Groove" ReadOnly="True">DJobCode</asp:textbox><asp:textbox id="DAM" style="Z-INDEX: 112; LEFT: 600px; POSITION: absolute; TOP: 251px; TEXT-ALIGN: right"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove" ReadOnly="True">DAM</asp:textbox><asp:textbox id="DAH" style="Z-INDEX: 111; LEFT: 520px; POSITION: absolute; TOP: 251px; TEXT-ALIGN: right"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove" ReadOnly="True">DAH</asp:textbox><asp:textbox id="DWH" style="Z-INDEX: 109; LEFT: 520px; POSITION: absolute; TOP: 186px; TEXT-ALIGN: right"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove" ReadOnly="True">DWH</asp:textbox><asp:textbox id="DDivisionCode" style="Z-INDEX: 108; LEFT: 616px; POSITION: absolute; TOP: 120px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="67px" BorderStyle="Groove" ReadOnly="True">DDivisionCode</asp:textbox><asp:textbox id="DDivision" style="Z-INDEX: 107; LEFT: 528px; POSITION: absolute; TOP: 120px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="88px" BorderStyle="Groove" ReadOnly="True">DDivision</asp:textbox><asp:textbox id="DJobTitle" style="Z-INDEX: 106; LEFT: 120px; POSITION: absolute; TOP: 120px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="104px" BorderStyle="Groove" ReadOnly="True">DJobTitle</asp:textbox><asp:textbox id="DEmpID" style="Z-INDEX: 105; LEFT: 616px; POSITION: absolute; TOP: 88px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="67px" BorderStyle="Groove" ReadOnly="True">DEmpID</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 104; LEFT: 120px; POSITION: absolute; TOP: 88px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="104px" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:textbox id="DFReason" style="Z-INDEX: 102; LEFT: 120px; POSITION: absolute; TOP: 286px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="118px" Width="560px" BorderStyle="Groove" TextMode="MultiLine" MaxLength="240">DFReason</asp:textbox><asp:image id="DAddWorkSheet1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" Height="406px" Width="680px" ImageUrl="Images\HR_AddWorkSheet_02.jpg"></asp:image><asp:textbox id="DWorkDate" style="Z-INDEX: 103; LEFT: 120px; POSITION: absolute; TOP: 152px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="104px" BorderStyle="Groove" ReadOnly="True">2008/10/10</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 9px" runat="server"
					ForeColor="Black" BackColor="White" Height="16px" Width="216px" BorderStyle="None" ReadOnly="True" Font-Names="Times New Roman"> 123</asp:textbox></FONT></form>
		</FONT></FORM>
	</body>
</HTML>
