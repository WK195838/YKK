<%@ Page Language="vb" AutoEventWireup="false" Inherits="ComplaintInSheet_02" aspCompat="True" EnableEventValidation = "false"  CodeFile="ComplaintInSheet_02.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>內部客訴處理報告書</title>
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
              <font face="新細明體"></font>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<div>
                  <img id="IMG1" onclick="return IMG1_onclick()" src="images/ComplaintInSheet_10.jpg"
                      style="z-index: 1; left: 0px; position: absolute; top: 8px" />
                  <asp:TextBox ID="DQCDate" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                      Height="20px" Style="z-index: 113; left: 488px; position: absolute; top: 576px"
                      Width="104px"></asp:TextBox>
                  <asp:CheckBox ID="DChkData3" runat="server" ForeColor="Red" Style="z-index: 183;
                      left: 600px; position: absolute; top: 1040px" Width="104px" />
                  <asp:Button ID="DAttachfile3" runat="server" CausesValidation="False" Style="z-index: 274;
                      left: 488px; position: absolute; top: 1032px" Text="開啟附檔資料夾" Width="112px" />
                  <asp:TextBox ID="DColorCode" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                      ForeColor="Black" Height="20px" Style="z-index: 113; left: 488px; position: absolute;
                      top: 176px" Width="192px"></asp:TextBox>
                  <asp:DropDownList ID="DStstusList" runat="server" BackColor="#C0FFFF" Height="20px"
                      Style="z-index: 134; left: 176px; position: absolute; top: 408px" Visible="False"
                      Width="192px">
                  </asp:DropDownList>
                  <asp:CheckBox ID="DChkData2" runat="server" ForeColor="Red" Style="z-index: 183;
                      left: 304px; position: absolute; top: 672px" Width="104px" />
                  <asp:CheckBox ID="DChkData1" runat="server" ForeColor="Red" Style="z-index: 182;
                      left: 608px; position: absolute; top: 544px" Width="104px" />
                  <asp:TextBox ID="DBUQTY" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                      Height="20px" onkeyup="if(isNaN(value))execCommand('undo')" Style="z-index: 113;
                      left: 496px; position: absolute; top: 1168px" Width="56px"></asp:TextBox>
                  <asp:TextBox ID="DERRORQTY" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
                      BorderStyle="None" ForeColor="Black" Height="20px" onkeyup="if(isNaN(value))execCommand('undo')"
                      Style="z-index: 113; left: 200px; position: absolute; top: 1136px" Width="56px"></asp:TextBox>
                  <asp:TextBox ID="DREMARK1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                      ForeColor="Black" Height="208px" Style="z-index: 113; left: 24px; position: absolute;
                      top: 744px" TextMode="MultiLine" Width="344px" ReadOnly="True"></asp:TextBox>
                  <asp:TextBox ID="DREMARK4" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                      ForeColor="Black" Height="20px" Style="z-index: 113; left: 488px; position: absolute;
                      top: 1072px" Width="200px"></asp:TextBox>
                  <asp:Image ID="LMapFile" runat="server" BorderStyle="Groove" Height="320px" Style="z-index: 100;
                      left: 384px; position: absolute; top: 640px" Width="304px" />
                  &nbsp;
                  <asp:HyperLink ID="LMapFile1" runat="server" ForeColor="Blue" Height="24px" Style="z-index: 190;
                      left: 384px; position: absolute; top: 936px" Target="_blank" ToolTip="簡圖放大" Width="304px">簡圖放大</asp:HyperLink>
                  &nbsp;&nbsp;
                  <asp:TextBox ID="DBUYER" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                      ForeColor="Black" Height="24px" Style="z-index: 110; left: 280px; position: absolute;
                      top: 280px" Width="408px"></asp:TextBox>
                  &nbsp;
                  <asp:TextBox ID="DBUYERCODE" runat="server" AutoCompleteType="Disabled" BackColor="#C0FFFF"
                      BorderStyle="Groove" ForeColor="Black" Height="24px" Style="z-index: 111; left: 176px;
                      position: absolute; top: 280px" Width="96px"></asp:TextBox>
                  <asp:TextBox ID="DCUSTOMER" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                      ForeColor="Black" Height="24px" Style="z-index: 112; left: 280px; position: absolute;
                      top: 248px" Width="408px"></asp:TextBox>
                  &nbsp;
                  <asp:TextBox ID="DCUSTOMERCODE" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                      ForeColor="Black" Height="24px" Style="z-index: 113; left: 176px; position: absolute;
                      top: 248px" Width="96px"></asp:TextBox>
                  <asp:TextBox ID="DSPEC" runat="server" AutoCompleteType="Disabled" BackColor="#C0FFFF"
                      BorderStyle="Groove" ForeColor="Black" Height="24px" Style="z-index: 127; left: 176px;
                      position: absolute; top: 216px" Width="96px"></asp:TextBox>
                  <asp:TextBox ID="DSPECNAME" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                      ForeColor="Black" Height="24px" Style="z-index: 112; left: 280px; position: absolute;
                      top: 216px" Width="408px"></asp:TextBox>
                  &nbsp;
                  <asp:TextBox ID="DSTATUS" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                      Height="24px" Style="z-index: 113; left: 176px; position: absolute; top: 440px"
                      TextMode="MultiLine" Width="192px"></asp:TextBox>
                  <asp:DropDownList ID="DUNIT2" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                      left: 496px; position: absolute; top: 1136px" Width="48px">
                  </asp:DropDownList>
                  <asp:DropDownList ID="DANSWER" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                      left: 200px; position: absolute; top: 1104px" Width="176px">
                  </asp:DropDownList>
                  <asp:DropDownList ID="DSITUATION" runat="server" BackColor="#C0FFFF" Height="20px"
                      Style="z-index: 134; left: 200px; position: absolute; top: 1072px"
                      Width="176px">
                  </asp:DropDownList>
                  <asp:DropDownList ID="DFREASON" runat="server" BackColor="#C0FFFF" Height="20px"
                      Style="z-index: 134; left: 200px; position: absolute; top: 1040px"
                      Width="176px">
                  </asp:DropDownList>
                  <asp:DropDownList ID="DFCONTENT" runat="server" BackColor="#C0FFFF" Height="20px"
                      Style="z-index: 134; left: 200px; position: absolute; top: 1008px"
                      Width="176px">
                  </asp:DropDownList>
                  &nbsp;
                  <asp:DropDownList ID="DSHIP" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                      left: 512px; position: absolute; top: 968px" Width="88px">
                  </asp:DropDownList>
                  <asp:DropDownList ID="DFCONFIRM" runat="server" BackColor="#C0FFFF" Height="20px"
                      Style="z-index: 134; left: 200px; position: absolute; top: 968px"
                      Width="112px">
                  </asp:DropDownList>
                  <asp:TextBox ID="DACCEMPNAME" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                      ForeColor="Black" Height="20px" Style="z-index: 113; left: 176px; position: absolute;
                      top: 640px" Width="200px"></asp:TextBox>
                  <asp:DropDownList ID="DACCDEPNAME" runat="server" BackColor="#C0FFFF" Height="20px"
                      Style="z-index: 134; left: 176px; position: absolute; top: 608px"
                      Width="200px">
                  </asp:DropDownList>
                  <asp:TextBox ID="DQCNO" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                      Height="20px" Style="z-index: 113; left: 176px; position: absolute; top: 576px"
                      Width="192px"></asp:TextBox>
                  <asp:DropDownList ID="DLOCATION" runat="server" BackColor="#C0FFFF" Height="20px"
                      Style="z-index: 134; left: 176px; position: absolute; top: 512px"
                      Width="192px">
                  </asp:DropDownList>
                  <asp:TextBox ID="DSTOCKSTS" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                      ForeColor="Black" Height="56px" Style="z-index: 113; left: 488px; position: absolute;
                      top: 408px" TextMode="MultiLine" Width="192px"></asp:TextBox>
                  <asp:TextBox ID="DDELIVERYDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                      ForeColor="Black" Height="20px" Style="z-index: 113; left: 176px; position: absolute;
                      top: 480px" Width="104px"></asp:TextBox>
                  &nbsp;
                  <asp:DropDownList ID="DSTOCKCHECK" runat="server" BackColor="#C0FFFF" Height="20px"
                      Style="z-index: 134; left: 488px; position: absolute; top: 480px"
                      Width="192px">
                  </asp:DropDownList>
                  &nbsp;
                  <asp:DropDownList ID="DUNIT1" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                      left: 496px; position: absolute; top: 312px" Width="48px">
                  </asp:DropDownList>
                  <asp:TextBox ID="DORQTY" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                      Height="20px" onkeyup="if(isNaN(value))execCommand('undo')" Style="z-index: 113;
                      left: 176px; position: absolute; top: 312px" Width="64px"></asp:TextBox>
                  &nbsp;
                  <asp:TextBox ID="DMATERIALNO" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                      ForeColor="Black" Height="56px" Style="z-index: 113; left: 176px; position: absolute;
                      top: 344px" TextMode="MultiLine" Width="192px"></asp:TextBox>
                  <asp:TextBox ID="DREMARK3" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                      Height="20px" Style="z-index: 113; left: 488px; position: absolute; top: 1008px"
                      Width="200px"></asp:TextBox>
                  &nbsp;
                  <asp:TextBox ID="DORNO" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                      Height="20px" MaxLength="7" Style="z-index: 126; left: 176px; position: absolute;
                      top: 176px; text-align: left" Width="200px"></asp:TextBox>
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
                  <asp:DropDownList ID="DCOMDEPNAME" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
                      Height="20px" Style="z-index: 134; left: 488px; position: absolute; top: 112px" Width="192px">
                  </asp:DropDownList>
                  <asp:Button ID="DAttachfile2" runat="server" CausesValidation="False" Style="z-index: 274;
                      left: 176px; position: absolute; top: 672px" Text="開啟附檔資料夾" Width="120px" />
                  &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
                  <asp:Button ID="DAttachfile1" runat="server" CausesValidation="False" Style="z-index: 274;
                      left: 488px; position: absolute; top: 536px" Text="開啟附檔資料夾" Width="112px" />
                  <asp:TextBox ID="DCOMDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                      ForeColor="Black" Height="20px" Style="z-index: 113; left: 176px; position: absolute;
                      top: 112px" Width="104px"></asp:TextBox>
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
                  <asp:TextBox ID="DCOMNAME" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                      ForeColor="Black" Height="20px" Style="z-index: 113; left: 176px; position: absolute;
                      top: 144px" Width="192px"></asp:TextBox>
                  &nbsp;&nbsp;&nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp;&nbsp;
                  <asp:TextBox ID="DNO" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                      Height="20px" Style="z-index: 113; left: 520px; position: absolute; top: 80px"
                      Width="160px"></asp:TextBox>
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
                  <asp:TextBox ID="D1" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
                      Height="24px" Style="z-index: 132; left: 2104px; position: absolute; top: 368px"
                      Width="190px"></asp:TextBox>
                  <asp:TextBox ID="D2" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
                      Height="24px" Style="z-index: 133; left: 2104px; position: absolute; top: 400px"
                      Width="190px"></asp:TextBox>
               <!--************************************************************ 
            ** 防止重覆按 [BUTTON 2 次]                               **     
            ************************************************************ -->
                  <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 126; left: -500px; position: absolute;
                      top: 100px; text-align: left">AAA</asp:TextBox>
                  &nbsp;
                  <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextBox1"
                      ErrorMessage="不可為空白" Width="96px"></asp:RequiredFieldValidator>
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp;
              </div>
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp;&nbsp;
              <asp:TextBox ID="D3" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
                  Height="24px" Style="z-index: 133; left: 1584px; position: absolute; top: 16px"
                  Width="190px"></asp:TextBox>
              <asp:TextBox ID="D4" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
                  Height="24px" Style="z-index: 133; left: 1584px; position: absolute; top: 56px"
                  Width="190px"></asp:TextBox>
              &nbsp;&nbsp;
              <asp:TextBox ID="DCOMADMIN" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                  ForeColor="Black" Height="20px" Style="z-index: 113; left: 488px; position: absolute;
                  top: 144px" Width="192px"></asp:TextBox>
              <asp:DropDownList ID="DCOMINTYPE" runat="server" BackColor="#C0FFFF" Height="20px"
                  Style="z-index: 134; left: 488px; position: absolute; top: 512px"
                  Width="192px">
              </asp:DropDownList>
              <asp:DropDownList ID="DLASTQC" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                  left: 176px; position: absolute; top: 544px" Width="192px">
              </asp:DropDownList>
              <asp:TextBox ID="DREMARK2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                  ForeColor="Black" Height="72px" Style="z-index: 113; left: 24px; position: absolute;
                  top: 1232px" TextMode="MultiLine" Width="664px"></asp:TextBox>
              <asp:TextBox ID="DERRORP" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                  Height="20px" Style="z-index: 113; left: 200px; position: absolute; top: 1168px"
                  Width="96px"></asp:TextBox>
              <asp:TextBox ID="DCODE" runat="server" BackColor="Transparent" BorderStyle="None"
                  ForeColor="White" Height="24px" Style="z-index: 138; left: 1816px; position: absolute;
                  top: 504px" Width="190px"></asp:TextBox>
              <asp:TextBox ID="DMappath" runat="server" BackColor="Transparent" BorderStyle="None"
                  ForeColor="White" Height="24px" Style="z-index: 189; left: 1384px; position: absolute;
                  top: 440px" Width="190px"></asp:TextBox>
              <asp:TextBox ID="chktemp" runat="server" BackColor="Transparent" BorderStyle="None"
                  ForeColor="White" Height="24px" Style="z-index: 187; left: 848px; position: absolute;
                  top: 296px" Width="190px"></asp:TextBox>
              <asp:CheckBox ID="DMach1" runat="server" Font-Size="Smaller" ForeColor="Red" Style="z-index: 183;
                  left: 376px; position: absolute; top: 1104px" Text="機台" Width="56px" />
              <asp:CheckBox ID="DMach2" runat="server" Font-Size="Smaller" ForeColor="Red" Style="z-index: 183;
                  left: 632px; position: absolute; top: 1104px" Text="非機台" Width="80px" />
              <asp:TextBox ID="DMACHNO" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                  Height="20px" Style="z-index: 157; left: 424px; position: absolute; top: 1104px;
                  text-align: left" Width="107px"></asp:TextBox>
              <asp:Label ID="Label1" runat="server" Font-Size="9pt" ForeColor="Black" Height="20px"
                  Style="z-index: 135; left: 536px; position: absolute; top: 1112px" Text="該工程POP機台號"
                  Width="104px"></asp:Label>
          </div>
    </form>
</body>
</html>
