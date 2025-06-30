<%@ Page Language="vb" AutoEventWireup="false" Inherits="FinalcheckSheet_01" aspCompat="True" EnableEventValidation = "false"  CodeFile="FinalcheckSheet_01.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>最終再檢驗處理報告書</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(strField)
			{
				window.open('DatePicker1.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
			
			               function GetDep()
{
        
    window.open('DepList.aspx?','','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
}	
		
			               function GetError(Str)
{
        
    window.open('ErrorList.aspx?Str=' + Str,'','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
}	
			
			

     function ConfirmMe(btn) {
                if(Page_ClientValidate())   {
                    btn.disabled="disabled";
				    var answer = confirm("是否執行作業嗎？ 請確認....");
				    if (answer) {
                        document.forms[0].__EVENTTARGET.value = btn.name;
                        document.forms[0].__EVENTARGUMENT.value = '';
                        document.forms[0].submit();
				    }                    
                    else    {
                        btn.disabled="";
                        return false;   
                    }				    
                }
                else    {
                    return false;
                }
            }
            


function IMG1_onclick() {

}

		</script>
		
</head>
<body>

  	<form id="Form1" runat="server">
  	 <div>
        	<FONT face="新細明體"></FONT>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<div>
               <img src="images/FinalcheckSheet_01.jpg" style="z-index: 1; left: 8px; position: absolute;
                   top: 8px" id="IMG1" onclick="return IMG1_onclick()" />
               &nbsp;
               <asp:HyperLink ID="LComplete" runat="server" Height="1px" NavigateUrl="BoardEdit.aspx"
                   Style="z-index: 103; left: 1352px; position: absolute; top: 56px" Target="_blank"
                   Visible="False" Width="112px">有修改委託書</asp:HyperLink>
               <asp:DropDownList ID="DEREASON" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 1136px; position: absolute; top: 360px" Visible="False" Width="200px">
                   </asp:DropDownList>
               <asp:DropDownList ID="DECONTENT" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 152px; position: absolute; top: 360px" Visible="False" Width="264px">
               </asp:DropDownList><asp:DropDownList ID="DSITUATION" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 152px; position: absolute; top: 400px" Visible="False" Width="264px">
               </asp:DropDownList>
               <asp:DropDownList ID="DQCANSWER" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 152px; position: absolute; top: 448px" Visible="False" Width="264px">
               </asp:DropDownList>
               <asp:DropDownList ID="DUNIT2" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 232px; position: absolute; top: 140px" Visible="False" Width="64px">
               </asp:DropDownList>
               <asp:DropDownList ID="DUNIT3" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134; left: 568px; position: absolute;
                   top: 140px" Width="64px" Visible="False">
               </asp:DropDownList><asp:DropDownList ID="DUNIT1" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 1552px; position: absolute; top: 112px" Width="64px" Visible="False">
               </asp:DropDownList><asp:TextBox ID="DQCODATE3" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 113;
                   left: 1488px; position: absolute; top: 728px" Width="96px" BorderStyle="None" ForeColor="Black"></asp:TextBox><input id="BQCODATE3" runat="server" style="z-index: 133; left: 1592px; width: 24px;
                   position: absolute; top: 728px" type="button" value="..." />
               <input id="BQCODATE4" runat="server" style="z-index: 133; left: 1752px; width: 24px;
                   position: absolute; top: 728px" type="button" value="..." />
               <asp:TextBox ID="DQCODATE4" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 1648px; position: absolute;
                   top: 728px" Width="96px"></asp:TextBox>
               <asp:TextBox ID="TextBox4" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
                   Font-Names="Times New Roman" Font-Size="11pt" ForeColor="Black" Style="z-index: 100;
                   left: 1624px; position: absolute; top: 728px" Width="24px">～</asp:TextBox>
               <asp:TextBox ID="DQCIDATE4" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 1640px; position: absolute;
                   top: 592px" Width="96px"></asp:TextBox>
               <asp:TextBox ID="TextBox6" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
                   Font-Names="Times New Roman" Font-Size="11pt" ForeColor="Black" Style="z-index: 100;
                   left: 952px; position: absolute; top: 592px" Width="24px">～</asp:TextBox>
               <input id="BQCIDATE4" runat="server" style="z-index: 133; left: 1744px; width: 24px;
                   position: absolute; top: 592px" type="button" value="..." />
               <asp:TextBox ID="DQCIDATE1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 816px; position: absolute;
                   top: 592px" Width="96px"></asp:TextBox>
               <input id="BQCIDATE3" runat="server" style="z-index: 133; left: 1584px; width: 24px;
                   position: absolute; top: 592px" type="button" value="..." />
               <asp:TextBox ID="DQCIDATE3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 1480px; position: absolute;
                   top: 592px" Width="96px"></asp:TextBox>
               <asp:TextBox ID="DERRORQTY" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 480px;
                   position: absolute; top: 140px; text-align: left" Width="80px"></asp:TextBox>
               &nbsp;
               <asp:Button ID="DAttachfile2" runat="server" CausesValidation="False" Style="z-index: 274;
                   left: 160px; position: absolute; top: 840px" Text="開啟附檔資料夾" Width="261px" />
               <asp:Button ID="DAttachfile1" runat="server" CausesValidation="False" Style="z-index: 274;
                   left: 160px; position: absolute; top: 296px" Text="開啟附檔資料夾" Width="261px" />
               <asp:TextBox ID="DTYPE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 152px; position: absolute;
                   top: 114px" Width="208px"></asp:TextBox>
               <asp:TextBox ID="DCOLORITEM" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 1136px; position: absolute;
                   top: 115px" Width="208px"></asp:TextBox>
               <asp:TextBox ID="DCUSTOMER" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 480px; position: absolute;
                   top: 88px" Width="208px"></asp:TextBox>
               &nbsp;
               <asp:TextBox ID="DCHECKDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 480px; position: absolute;
                   top: 116px" Width="96px"></asp:TextBox>
               <asp:TextBox ID="DQCODATE1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 816px; position: absolute;
                   top: 720px" Width="96px"></asp:TextBox><input id="BQCODATE1" runat="server" style="z-index: 133; left: 920px; width: 24px;
                   position: absolute; top: 720px" type="button" value="..." />
               <asp:TextBox ID="DDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 1136px; position: absolute;
                   top: 144px" Width="80px"></asp:TextBox>
               <asp:TextBox ID="TextBox27" runat="server" BackColor="White" BorderColor="White"
                   BorderStyle="None" Font-Names="Times New Roman" Font-Size="11pt" ForeColor="Black"
                   Style="z-index: 100; left: 1616px; position: absolute; top: 592px" Width="24px">～</asp:TextBox>
               <asp:TextBox ID="DORNO" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 832px; position: absolute;
                   top: 115px" Width="184px"></asp:TextBox>
               <input id="BQCODATE2" runat="server" style="z-index: 133; left: 1080px; width: 24px;
                   position: absolute; top: 720px" type="button" value="..." />
               <asp:TextBox ID="DQCODATE2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 976px; position: absolute;
                   top: 720px" Width="96px"></asp:TextBox>
               <input id="BCHECKDATE" runat="server" style="z-index: 133; left: 584px; width: 24px;
                   position: absolute; top: 114px" type="button" value="..." />
               <asp:TextBox ID="TextBox26" runat="server" BackColor="White" BorderColor="White"
                   BorderStyle="None" Font-Names="Times New Roman" Font-Size="11pt" ForeColor="Black"
                   Style="z-index: 100; left: 952px; position: absolute; top: 720px" Width="24px">～</asp:TextBox>
             
               <input id="BQCIDATE2" runat="server" style="z-index: 133; left: 1080px; width: 24px;
                   position: absolute; top: 592px" type="button" value="..." />
               <asp:TextBox ID="DFINISHDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 264px; position: absolute;
                   top: 880px" Width="120px"></asp:TextBox>
               <input id="BFINISHDATE" runat="server" style="z-index: 133; left: 392px; width: 24px;
                   position: absolute; top: 880px" type="button" value="..." />
               <asp:TextBox ID="DNO" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 152px; position: absolute;
                   top: 88px" Width="208px"></asp:TextBox>
               <asp:TextBox ID="DQCIDATE2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 976px; position: absolute;
                   top: 592px" Width="96px"></asp:TextBox>
               <input id="BQCIDATE1" runat="server" style="z-index: 133; left: 920px; width: 24px;
                   position: absolute; top: 592px" type="button" value="..." />
               <asp:TextBox ID="DAPPNAME" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 808px; position: absolute;
                   top: 144px" Width="80px"></asp:TextBox>
               &nbsp;
               <asp:TextBox ID="DCHECKQTY" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 152px;
                   position: absolute; top: 142px; text-align: left" Width="80px"></asp:TextBox>
               <asp:TextBox ID="DQTY" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 1464px;
                   position: absolute; top: 115px; text-align: left" Width="80px"></asp:TextBox>
               &nbsp;
               <asp:TextBox ID="DECONTENT1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="24px" MaxLength="7" Style="z-index: 126; left: 424px;
                   position: absolute; top: 360px; text-align: left" TextMode="MultiLine" Width="584px"></asp:TextBox>
               &nbsp;&nbsp;
               <asp:TextBox ID="DQCANSWER1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="40px" MaxLength="7" Style="z-index: 126; left: 424px;
                   position: absolute; top: 440px; text-align: left" TextMode="MultiLine" Width="1344px"></asp:TextBox>
               <asp:TextBox ID="DQCANSWER2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="24px" MaxLength="7" Style="z-index: 126; left: 480px;
                   position: absolute; top: 802px; text-align: left" TextMode="MultiLine" Width="1304px"></asp:TextBox>
               <asp:TextBox ID="DEREASON1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="24px" MaxLength="7" Style="z-index: 126; left: 1344px;
                   position: absolute; top: 360px; text-align: left" TextMode="MultiLine" Width="440px"></asp:TextBox>
               <asp:TextBox ID="DERRORSTS" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="24px" MaxLength="7" Style="z-index: 126; left: 152px;
                   position: absolute; top: 256px; text-align: left" TextMode="MultiLine" Width="1632px"></asp:TextBox>
               &nbsp;
               <asp:DropDownList ID="DACCDEPNAME" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 808px; position: absolute; top: 88px" Visible="False" Width="208px" AutoPostBack="True">
               </asp:DropDownList>
               &nbsp; &nbsp;
               <asp:TextBox ID="DSITUATION1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="40px" MaxLength="7" Style="z-index: 126; left: 424px;
                   position: absolute; top: 392px; text-align: left" TextMode="MultiLine" Width="1344px"></asp:TextBox>
               <asp:TextBox ID="DERROR1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="48px" MaxLength="7" Style="z-index: 126; left: 152px;
                   position: absolute; top: 205px; text-align: left" TextMode="MultiLine" Width="296px"></asp:TextBox>
               <input id="BERROR1" runat="server" style="z-index: 133; left: 448px; width: 24px;
               position: absolute; top: 205px" type="button" value="..." />
               &nbsp;
               <asp:TextBox ID="DERROR3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="48px" MaxLength="7" Style="z-index: 126; left: 808px;
                   position: absolute; top: 205px; text-align: left" TextMode="MultiLine" Width="296px"></asp:TextBox>
               <asp:TextBox ID="DERROR4" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="48px" MaxLength="7" Style="z-index: 126; left: 1136px;
                   position: absolute; top: 205px; text-align: left" TextMode="MultiLine" Width="296px"></asp:TextBox>
               <input id="BERROR4" runat="server" style="z-index: 133; left: 1432px; width: 24px;
               position: absolute; top: 205px" type="button" value="..." />
               <asp:TextBox ID="DERROR5" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="48px" MaxLength="7" Style="z-index: 126; left: 1464px;
                   position: absolute; top: 205px; text-align: left" TextMode="MultiLine" Width="296px"></asp:TextBox>
               <input id="BERROR5" runat="server" style="z-index: 133; left: 1760px; width: 24px;
               position: absolute; top: 205px" type="button" value="..." />
               <input id="BERROR3" runat="server" style="z-index: 133; left: 1104px; width: 24px;
               position: absolute; top: 208px; height: 24px;" type="button" value="..." />
               <asp:TextBox ID="DERROR2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="48px" MaxLength="7" Style="z-index: 126; left: 480px;
                   position: absolute; top: 205px; text-align: left" TextMode="MultiLine" Width="296px"></asp:TextBox>
               <asp:TextBox ID="DQCO1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="120px" MaxLength="7" Style="z-index: 126; left: 152px;
                   position: absolute; top: 680px; text-align: left" TextMode="MultiLine" Width="320px"></asp:TextBox>
               <asp:TextBox ID="DQCO3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="120px" MaxLength="7" Style="z-index: 126; left: 1136px;
                   position: absolute; top: 680px; text-align: left" TextMode="MultiLine" Width="320px"></asp:TextBox>
               &nbsp;
               <asp:TextBox ID="DQCO2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="120px" MaxLength="7" Style="z-index: 126; left: 480px;
                   position: absolute; top: 680px; text-align: left" TextMode="MultiLine" Width="320px"></asp:TextBox>
               &nbsp;
               <asp:TextBox ID="DQCI1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="120px" MaxLength="7" Style="z-index: 126; left: 152px;
                   position: absolute; top: 544px; text-align: left" TextMode="MultiLine" Width="320px"> </asp:TextBox>
               <asp:TextBox ID="DQCI3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="120px" MaxLength="7" Style="z-index: 126; left: 1136px;
                   position: absolute; top: 544px; text-align: left" TextMode="MultiLine" Width="320px"></asp:TextBox>
               &nbsp;
               <input id="BERROR2" runat="server" style="z-index: 133; left: 776px; width: 24px;
               position: absolute; top: 205px" type="button" value="..." />
               &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:DropDownList ID="DRDIVISION" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 1464px; position: absolute; top: 88px" Visible="False" Width="216px">
               </asp:DropDownList>
               <asp:DropDownList ID="DSLDDIVISION" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 1136px; position: absolute; top: 88px" Visible="False" Width="208px" AutoPostBack="True">
               </asp:DropDownList>
               &nbsp;
               <asp:TextBox ID="DQCI2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="120px" MaxLength="7" Style="z-index: 126; left: 480px;
                   position: absolute; top: 544px; text-align: left" TextMode="MultiLine" Width="320px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp;
               &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DACCEMPNAME" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 1136px; position: absolute;
                   top: 880px; text-align: left" Width="216px"></asp:TextBox>
               &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 1; left: 12px;
                   position: absolute; top: 1034px" visible="false" />
               <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
                   Height="56px" Style="z-index: 132; left: 59px; position: absolute; top: 965px"
                   TextMode="MultiLine" Width="536px"></asp:TextBox>
               <asp:TextBox ID="DReason" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
                   Style="z-index: 133; left: 243px; position: absolute; top: 1039px" Visible="False"
                   Width="352px"></asp:TextBox>
               <asp:DropDownList ID="DReasonCode" runat="server" BackColor="Yellow" Height="20px"
                   Style="z-index: 134; left: 167px; position: absolute; top: 1042px" Visible="False"
                   Width="64px">
                   <asp:ListItem>01</asp:ListItem>
                   <asp:ListItem>02</asp:ListItem>
               </asp:DropDownList>
               <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
                   Height="56px" Style="z-index: 135; left: 171px; position: absolute; top: 1072px"
                   Visible="False" Width="424px"></asp:TextBox>
               <asp:Label ID="DHistoryLabel" runat="server" Font-Size="11pt" ForeColor="Blue" Height="20px"
                   Style="z-index: 137; left: 14px; position: absolute; top: 1157px" Text="核定履歷"></asp:Label>
               <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
                   BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
                   Style="z-index: 136; left: 16px; position: absolute; top: 1184px" Width="780px">
                   <RowStyle BackColor="White" Font-Size="9pt" ForeColor="#330099" />
                   <Columns>
                       <asp:BoundField DataField="StepNameDesc" HeaderText="工程">
                           <ItemStyle HorizontalAlign="Left" />
                       </asp:BoundField>
                       <asp:BoundField DataField="DecideName" HeaderText="擔當" />
                       <asp:BoundField DataField="AgentName" HeaderText="代理/兼職" />
                       <asp:BoundField DataField="FlowTypeDesc" HeaderText="類別" />
                       <asp:BoundField DataField="StsDesc" HeaderText="處理結果" />
                       <asp:BoundField DataField="DecideDescA" HeaderText="說明">
                           <ItemStyle HorizontalAlign="Left" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Description" HeaderText="核定時間">
                           <ItemStyle HorizontalAlign="Left" />
                       </asp:BoundField>
                   </Columns>
                   <HeaderStyle BackColor="#CC6600" Font-Size="9pt" ForeColor="#FFFFCC" HorizontalAlign="Center"
                       VerticalAlign="Middle" />
               </asp:GridView>
               <img id="DDescSheet" runat="server" src="images/Sheet_Remark.jpg" style="z-index: 1;
                   left: 13px; position: absolute; top: 959px" />
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               <asp:TextBox ID="D1" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
                   Height="24px" Style="z-index: 132; left: 768px; position: absolute; top: 300px"
                   Width="190px"></asp:TextBox>
               <asp:TextBox ID="D2" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
                   Height="24px" Style="z-index: 133; left: 769px; position: absolute; top: 337px"
                   Width="190px"></asp:TextBox>
               <!--************************************************************ 
            ** 防止重覆按 [BUTTON 2 次]                               **     
            ************************************************************ -->
               <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 126; left: -500px; position: absolute;
                   top: 100px; text-align: left">AAA</asp:TextBox>
               &nbsp;
               <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextBox1"
                   ErrorMessage="不可為空白" Width="96px"></asp:RequiredFieldValidator>
               &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp;
               <asp:Button ID="BSAVE" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 114; left: 354px; position: absolute; top: 1144px" Text="儲存"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BNG2" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 115; left: 445px; position: absolute; top: 1144px" Text="NG2"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BNG1" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 116; left: 537px; position: absolute; top: 1144px" Text="NG1"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BOK" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 117; left: 629px; position: absolute; top: 1144px" Text="OK"
                   UseSubmitBehavior="false" Width="80px" />
           </div>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="D3" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 133; left: 1584px; position: absolute; top: 16px"
               Width="190px"></asp:TextBox>
           <asp:TextBox ID="D4" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 133; left: 1584px; position: absolute; top: 56px"
               Width="190px"></asp:TextBox>
           <asp:HyperLink ID="HyperLink1" runat="server">HyperLink</asp:HyperLink></div>
    </form>
</body>
</html>
