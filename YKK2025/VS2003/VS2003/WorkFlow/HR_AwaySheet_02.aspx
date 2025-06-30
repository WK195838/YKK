<%@ Page Language="vb" AutoEventWireup="false" Codebehind="HR_AwaySheet_02.aspx.vb" Inherits="SPD.HR_AwaySheet_02"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>外出申請書</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			//--Calendar------------------------------------
			function CalendarPicker(strField1, strField2, strField3) {
				window.open('VacationDatePicker.aspx?field1=' + strField1 + '&field2=' + strField2 + '&field3=' + strField3,'CalendarPopup','width=250,height=190,resizable=yes');
			}

			function ShowCardTime()
			{
				window.open('HR_CardTimeList01.aspx?pMonth=' + document.Form1.DSalaryYM.value + '&pDepoID=' + document.Form1.DDepoCode.value + '&pEmpID=' + document.Form1.DEmpID.value, 'CardTimeSheet','width=290,height=600,resizable=yes,scrollbar=yes');
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
				<asp:button id="BCardTime" style="Z-INDEX: 136; LEFT: 596px; POSITION: absolute; TOP: 150px"
					runat="server" CausesValidation="False" Width="85px" Height="24px" Text="刷卡記錄" BackColor="#FFFFC0"></asp:button>
				<asp:textbox id="DPlace" style="Z-INDEX: 142; LEFT: 120px; POSITION: absolute; TOP: 152px" runat="server"
					Width="104px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" MaxLength="20">DOtherPlace</asp:textbox>
				<asp:textbox id="DName" style="Z-INDEX: 141; LEFT: 456px; POSITION: absolute; TOP: 88px" runat="server"
					Width="144px" Height="20px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True" BorderStyle="Groove">DName</asp:textbox>
				<asp:textbox id="DBEndH" style="Z-INDEX: 140; LEFT: 400px; POSITION: absolute; TOP: 186px" runat="server"
					Width="48px" Height="20px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True" BorderStyle="Groove">2008/10/10</asp:textbox>
				<asp:textbox id="DAEndH" style="Z-INDEX: 139; LEFT: 400px; POSITION: absolute; TOP: 218px" runat="server"
					Width="48px" Height="20px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True" BorderStyle="Groove">2008/10/10</asp:textbox>
				<asp:textbox id="DAStartH" style="Z-INDEX: 138; LEFT: 208px; POSITION: absolute; TOP: 218px"
					runat="server" Width="48px" Height="20px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True"
					BorderStyle="Groove">2008/10/10</asp:textbox>
				<asp:textbox id="DBStartH" style="Z-INDEX: 137; LEFT: 208px; POSITION: absolute; TOP: 186px"
					runat="server" Width="48px" Height="20px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True"
					BorderStyle="Groove">2008/10/10</asp:textbox><asp:textbox id="DBDay" style="Z-INDEX: 135; LEFT: 536px; POSITION: absolute; TOP: 186px; TEXT-ALIGN: right"
					runat="server" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">0</asp:textbox><asp:textbox id="DADay" style="Z-INDEX: 134; LEFT: 536px; POSITION: absolute; TOP: 218px; TEXT-ALIGN: right"
					runat="server" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">0</asp:textbox><asp:textbox id="DAEndDate" style="Z-INDEX: 133; LEFT: 312px; POSITION: absolute; TOP: 218px"
					runat="server" Height="20px" Width="88px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">2008/10/10</asp:textbox><asp:textbox id="DAStartDate" style="Z-INDEX: 132; LEFT: 120px; POSITION: absolute; TOP: 218px"
					runat="server" Height="20px" Width="88px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">2008/10/10</asp:textbox><asp:textbox id="DBEndDate" style="Z-INDEX: 131; LEFT: 312px; POSITION: absolute; TOP: 186px"
					runat="server" Height="20px" Width="88px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">2008/10/10</asp:textbox><asp:textbox id="DBStartDate" style="Z-INDEX: 130; LEFT: 120px; POSITION: absolute; TOP: 186px"
					runat="server" Height="20px" Width="88px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">2008/10/10</asp:textbox><asp:textbox id="DAHour" style="Z-INDEX: 128; LEFT: 600px; POSITION: absolute; TOP: 218px; TEXT-ALIGN: right"
					runat="server" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">0</asp:textbox><asp:textbox id="DJobAgent" style="Z-INDEX: 127; LEFT: 456px; POSITION: absolute; TOP: 152px"
					runat="server" Height="20px" Width="136px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" MaxLength="20">DJobAgent</asp:textbox><asp:textbox id="DOtherPlace" style="Z-INDEX: 126; LEFT: 224px; POSITION: absolute; TOP: 152px"
					runat="server" Height="20px" Width="123px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" MaxLength="20">DOtherPlace</asp:textbox><asp:textbox id="DDepoCode" style="Z-INDEX: 124; LEFT: 496px; POSITION: absolute; TOP: 120px"
					runat="server" Height="20px" Width="25px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">01</asp:textbox><asp:textbox id="DDepoName" style="Z-INDEX: 123; LEFT: 456px; POSITION: absolute; TOP: 120px"
					runat="server" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">台北</asp:textbox><asp:textbox id="DSalaryYM" style="Z-INDEX: 122; LEFT: 264px; POSITION: absolute; TOP: 88px"
					runat="server" Height="20px" Width="81px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">2008/10</asp:textbox><asp:textbox id="DJobCode" style="Z-INDEX: 114; LEFT: 264px; POSITION: absolute; TOP: 120px"
					runat="server" Height="20px" Width="81px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">DJobCode</asp:textbox><asp:textbox id="DBHour" style="Z-INDEX: 112; LEFT: 600px; POSITION: absolute; TOP: 186px; TEXT-ALIGN: right"
					runat="server" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">0</asp:textbox><asp:textbox id="DDivisionCode" style="Z-INDEX: 111; LEFT: 616px; POSITION: absolute; TOP: 120px"
					runat="server" Height="20px" Width="65px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">DDivisionCode</asp:textbox><asp:textbox id="DDivision" style="Z-INDEX: 110; LEFT: 528px; POSITION: absolute; TOP: 120px"
					runat="server" Height="20px" Width="88px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">DDivision</asp:textbox><asp:textbox id="DJobTitle" style="Z-INDEX: 109; LEFT: 120px; POSITION: absolute; TOP: 120px"
					runat="server" Height="20px" Width="144px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">DJobTitle</asp:textbox><asp:textbox id="DEmpID" style="Z-INDEX: 108; LEFT: 600px; POSITION: absolute; TOP: 88px" runat="server"
					Height="20px" Width="81px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">DEmpID</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 107; LEFT: 120px; POSITION: absolute; TOP: 88px" runat="server"
					Height="20px" Width="144px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">DDate</asp:textbox><asp:textbox id="DFReason" style="Z-INDEX: 102; LEFT: 120px; POSITION: absolute; TOP: 250px"
					runat="server" Height="122px" Width="560px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" MaxLength="240" TextMode="MultiLine">DFReason</asp:textbox><asp:image id="DAwaySheet1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Height="372px" Width="680px" ImageUrl="Images\HR_AwaySheet_01.jpg"></asp:image><asp:textbox id="DNo" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 9px" runat="server"
					Height="16px" Width="216px" BorderStyle="None" ReadOnly="True" ForeColor="Black" BackColor="White" Font-Names="Times New Roman"> 123</asp:textbox></FONT></form>
		</FONT></FORM>
	</body>
</HTML>
