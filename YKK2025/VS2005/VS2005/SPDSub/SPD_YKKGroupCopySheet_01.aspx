<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SPD_YKKGroupCopySheet_01.aspx.vb" Inherits="SPD_YKKGroupCopySheet_01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>姊妹社圖面複製履歷</title>
    
    <script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(field1)
			{
				window.open('DatePicker.aspx?field1=' + field1,'calendarPopup','width=250,height=190,resizable=yes');
			}
			
			   function GetMapNo()
{
        
    window.open('MapNoList.aspx?','','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
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

  </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp;
        <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 132; left: 19px;
            position: absolute; top: 601px" visible="false" />
        &nbsp;
        <asp:TextBox ID="DReason" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
            Style="z-index: 100; left: 241px; position: absolute; top: 604px" Visible="False"
            Width="352px"></asp:TextBox>
        <asp:DropDownList ID="DReasonCode" runat="server" Height="20px" Style="z-index: 101;
            left: 174px; position: absolute; top: 608px" Visible="False" Width="64px" BackColor="Yellow">
            <asp:ListItem>01</asp:ListItem>
            <asp:ListItem>02</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 102; left: 171px; position: absolute; top: 640px"
            Visible="False" Width="424px"></asp:TextBox>
        <asp:Label ID="DHistoryLabel" runat="server" Font-Size="11pt" ForeColor="Blue" Height="20px"
            Style="z-index: 103; left: 25px; position: absolute; top: 743px" Text="核定履歷"></asp:Label>
        <img id="DDescSheet" runat="server" src="images/Sheet_Remark.jpg" style="z-index: 99;
            left: 19px; position: absolute; top: 523px" />
        <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 104; left: 66px; position: absolute; top: 528px"
            TextMode="MultiLine" Width="536px"></asp:TextBox>
        &nbsp;&nbsp;
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
            Style="z-index: 105; left: 24px; position: absolute; top: 771px" Width="780px">
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
            <HeaderStyle BackColor="#CC6600" Font-Size="9pt" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
        <asp:Button ID="BSAVE" runat="server" Text="SAVE" Style="z-index: 106; left: 435px; position: absolute; top: 737px" Width="80px" />
        <asp:Button ID="BNG1" runat="server" Text="NG1" Style="z-index: 107; left: 526px; position: absolute; top: 737px" Width="80px" />
        <asp:Button ID="BNG2" runat="server" Text="NG2" Style="z-index: 108; left: 618px; position: absolute; top: 737px" Width="80px" />
        <asp:Button ID="BOK" runat="server" Text="OK" Style="z-index: 109; left: 710px; position: absolute; top: 737px" Width="80px"  OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false" />
        &nbsp;&nbsp;
        <asp:Image ID="Image1" runat="server" ImageUrl="~/images/YKKGroupCopySheet_01.jpg"
            Style="z-index: 110; left: 4px; position: absolute; top: 6px" />
        <asp:CheckBox ID="DCopyCheck3" runat="server" AutoPostBack="True" Style="z-index: 111;
            left: 212px; position: absolute; top: 362px" />
        <asp:TextBox ID="DMapNo" runat="server" BackColor="#C0FFFF" Style="z-index: 112;
            left: 211px; position: absolute; top: 171px" Width="159px"></asp:TextBox>
        <asp:TextBox ID="DOFormSno" runat="server" BackColor="#C0FFFF" Style="z-index: 113;
            left: 300px; position: absolute; top: 235px" Width="78px"></asp:TextBox>
        &nbsp;
        <input id="BMapNo" runat="server" style="z-index: 133; left: 381px; width: 24px; position: absolute;
            top: 170px" type="button" value="..." />
        &nbsp; &nbsp;
        <input id="BDate" runat="server" style="z-index: 134; left: 768px; width: 24px; position: absolute;
            top: 390px" type="button" value="..." />
        <asp:TextBox ID="DRemark" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="52px" MaxLength="100" Style="z-index: 114; left: 212px;
            position: absolute; top: 426px; text-align: left" TextMode="MultiLine" Width="580px"></asp:TextBox>
        <asp:TextBox ID="DSliderCode" runat="server" BackColor="#C0FFFF" Style="z-index: 115;
            left: 211px; position: absolute; top: 201px" Width="190px"></asp:TextBox>
        <asp:TextBox ID="DWaveCode" runat="server" BackColor="#C0FFFF" Style="z-index: 116;
            left: 601px; position: absolute; top: 202px" Width="190px"></asp:TextBox>
        <asp:TextBox ID="DProvideDate" runat="server" BackColor="#C0FFFF" Style="z-index: 117;
            left: 604px; position: absolute; top: 393px" Width="149px"></asp:TextBox>
        <asp:TextBox ID="DDate" runat="server" BackColor="#C0FFFF" Style="z-index: 118;
            left: 601px; position: absolute; top: 105px" Width="190px"></asp:TextBox>
        <asp:TextBox ID="DPerson" runat="server" BackColor="#C0FFFF" Style="z-index: 119;
            left: 601px; position: absolute; top: 136px" Width="190px"></asp:TextBox>
        <asp:TextBox ID="DBuyer" runat="server" BackColor="#C0FFFF" Style="z-index: 120;
            left: 601px; position: absolute; top: 168px" Width="190px"></asp:TextBox>
        &nbsp; &nbsp;
        <asp:TextBox ID="DNo" runat="server" BackColor="#C0FFFF" Style="z-index: 121;
            left: 213px; position: absolute; top: 106px" Width="190px"></asp:TextBox>
        <asp:TextBox ID="DDivision" runat="server" BackColor="#C0FFFF" Style="z-index: 122;
            left: 211px; position: absolute; top: 136px" Width="190px"></asp:TextBox>
    
        <asp:TextBox ID="DOFormNo" runat="server" BackColor="#C0FFFF" Style="z-index: 123;
            left: 212px; position: absolute; top: 233px" Width="72px"></asp:TextBox>
        &nbsp;&nbsp;
        <asp:TextBox ID="DForcast" runat="server" BackColor="#C0FFFF" Style="z-index: 124;
            left: 210px; position: absolute; top: 395px" Width="194px"></asp:TextBox>
        <asp:TextBox ID="DCopyReason" runat="server" BackColor="#C0FFFF" Style="z-index: 125;
            left: 422px; position: absolute; top: 361px" Width="367px"></asp:TextBox>
        &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <asp:HyperLink ID="LFormsno" runat="server" Style="z-index: 126; left: 401px; position: absolute;
            top: 235px" Width="103px" Target="_blank">原委託</asp:HyperLink>
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <asp:DropDownList ID="DYKKGroup" runat="server" Style="z-index: 127; left: 210px;
            position: absolute; top: 265px" Width="584px" BackColor="#C0FFFF">
        </asp:DropDownList>
        <asp:CheckBox ID="DCopyCheck1" runat="server" AutoPostBack="True" Style="z-index: 128;
            left: 212px; position: absolute; top: 297px" />
        <asp:CheckBox ID="DCopyCheck2" runat="server" AutoPostBack="True" Style="z-index: 129;
            left: 212px; position: absolute; top: 330px" />
        <br />
        <br />
            <!--************************************************************ 
            ** 防止重覆按 [BUTTON 2 次]                               **     
            ************************************************************ -->
            <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 130; left: -500px;position: absolute; top: 100px; text-align: left">AAA</asp:TextBox>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextBox1" ErrorMessage="不可為空白"></asp:RequiredFieldValidator>            
            
        </div>
    </form>
   

</body>
</html>


