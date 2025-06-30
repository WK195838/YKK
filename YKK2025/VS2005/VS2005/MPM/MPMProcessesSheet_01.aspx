<%@ Page Language="vb" AutoEventWireup="false" Inherits="MPMProcessesSheet_01" aspCompat="True" CodeFile="MPMProcessesSheet_01.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>機械加工依賴書</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(field1)
			{
				window.open('DatePicker.aspx?field1=' + field1,'calendarPopup','width=250,height=190,resizable=yes');
			}
			
			    function EmpDatePicker(userid)
{
        
    window.open('DivisionEmpList.aspx?userid='+userid,'','status=0,toolbar=0,width=500,height=500,top=10,resizable=yes');
   
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
		
</head>
<body>

  	<form id="Form1" runat="server">
  	 <div>
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/MPM-ProcessesSheet02.jpg" />
           <asp:HyperLink ID="LNO" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
               Style="z-index: 124; left: 461px; position: absolute; top: 159px" Target="_blank"
               Visible="False" Width="40px">履歷</asp:HyperLink>
           <asp:DropDownList ID="DDivisionCode" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 119; left: 349px; position: absolute;
            top: 229px" Width="147px">
           </asp:DropDownList>
           <input id="BDate2" runat="server" style="z-index: 144; left: 468px; width: 24px;
            position: absolute; top: 377px" type="button" value="..." />
           <asp:DropDownList ID="DEngine" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 100; left: 491px; position: absolute;
            top: 447px" Width="108px" AutoPostBack="True">
           </asp:DropDownList>
           &nbsp;&nbsp;
           <asp:TextBox ID="D1" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
               Style="z-index: 101; left: 720px; position: relative; top: -70px" Width="112px"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DEngine1" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="27px" Style="z-index: 102; left: 72px; position: absolute;
            top: 488px" Width="94px" Font-Size="Larger"></asp:TextBox>
           <asp:TextBox ID="DEngine2" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="27px" Style="z-index: 103; left: 177px; position: absolute;
               top: 484px" Width="94px" Font-Size="Larger"></asp:TextBox>
           <asp:TextBox ID="DEngine3" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="27px" Style="z-index: 104; left: 284px; position: absolute;
               top: 484px" Width="94px" Font-Size="Larger"></asp:TextBox>
           <asp:TextBox ID="DEngine4" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="27px" Style="z-index: 105; left: 387px; position: absolute;
               top: 484px" Width="94px" Font-Size="Larger"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DEngine11" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="27px" Style="z-index: 106; left: 387px; position: absolute;
               top: 520px" Width="94px" Font-Size="Larger"></asp:TextBox>
           &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="DEngine13" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="27px" Style="z-index: 107; left: 600px; position: absolute;
               top: 520px" Width="94px" Font-Size="Larger"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DEngine6" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="27px" Style="z-index: 108; left: 600px; position: absolute;
               top: 484px" Width="94px" Font-Size="Larger"></asp:TextBox>
           <asp:TextBox ID="DEngine12" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="27px" Style="z-index: 109; left: 493px; position: absolute;
               top: 520px" Width="94px" Font-Size="Larger"></asp:TextBox>
           <asp:TextBox ID="DEngine14" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="27px" Style="z-index: 110; left: 705px; position: absolute;
               top: 520px" Width="94px" Font-Size="Larger"></asp:TextBox>
           &nbsp;&nbsp;
           <asp:TextBox ID="DEngine5" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="27px" Style="z-index: 111; left: 493px; position: absolute;
               top: 484px" Width="94px" Font-Size="Larger"></asp:TextBox>
           <asp:TextBox ID="DEngine7" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="27px" Style="z-index: 112; left: 705px; position: absolute;
               top: 484px" Width="94px" Font-Size="Larger"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DClinter" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Style="z-index: 113; left: 605px; position: absolute;
               top: 193px" Width="148px"></asp:TextBox>
           <asp:Image ID="LMapFile" runat="server" BorderStyle="Groove" Height="219px"
              
               Style="z-index: 114; left: 26px; position: absolute; top: 184px" Width="223px" />
           &nbsp; &nbsp;&nbsp;
           <input id="DMapFile" runat="server" name="UploadFile" style="z-index: 141; left: 25px;
            width: 228px; position: absolute; top: 407px; height: 26px; background-color: #ffff00"
            type="file" visible="true" />
           <asp:HyperLink ID="LMapFile1" runat="server" Style="z-index: 115; left: 28px; position: absolute;
               top: 409px" Target="_blank" Width="167px">簡圖放大</asp:HyperLink>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
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
           &nbsp;
           <asp:Button ID="BEngineAdd" runat="server" CausesValidation="False" Style="z-index: 116;
               left: 24px; position: absolute; top: 607px" Text="加入工程" Visible="False" />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;
        <input id="BSAVE" runat="server" name="Button1" onclick="Button('SAVE');" style="z-index: 148;
            left: 318px; width: 80px; position: absolute; top: 1364px; height: 28px" type="button"
            value="儲存" />
        <input id="BNG2" runat="server" name="Button1" onclick="Button('NG2');" style="z-index: 147;
            left: 403px; width: 80px; position: absolute; top: 1364px; height: 28px" type="button"
            value="NG2" />
        <input id="BNG1" runat="server" name="Button1" onclick="Button('NG1');" style="z-index: 146;
            left: 487px; width: 80px; position: absolute; top: 1364px; height: 28px" type="button"
            value="NG1" />
        &nbsp;&nbsp;
        <input id="BOK" runat="server" name="BOK" onclick="Button('OK');" style="z-index: 145;
            left: 570px; width: 80px; position: absolute; top: 1364px; height: 28px" type="button"
            value="OK" />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <input id="BDate1" runat="server" style="z-index: 144; left: 722px; width: 24px;
            position: absolute; top: 230px" type="button" value="..." /><input id="BDate3" runat="server" style="z-index: 144; left: 723px; width: 24px;
            position: absolute; top: 153px" type="button" value="..." /><input id="BClienter" runat="server" style="z-index: 143; left: 759px; width: 20px;
            position: absolute; top: 193px" type="button" value="..." visible="false" />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;&nbsp; 
           <asp:Panel ID="DPanel" runat="server" BackColor="LightGreen" ForeColor="White" Height="30px"
               Style="z-index: 117; left: 23px; position: absolute; top: 630px" Width="780px" Visible="False">
               &nbsp;
               <asp:CheckBoxList ID="CheckBoxList1" runat="server" RepeatDirection="Horizontal"
                   Style="z-index: 102; left: 32px; position: absolute; top: 24px">
               </asp:CheckBoxList>
               <asp:RadioButtonList ID="RadioButtonList1" runat="server" Style="z-index: 101; left: 1px;
                   position: absolute; top: 3px">
               </asp:RadioButtonList>
           </asp:Panel>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:DropDownList ID="DType2" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 118; left: 416px; position: absolute;
            top: 267px" Width="82px">
           </asp:DropDownList>
           <asp:DropDownList ID="DType1" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 119; left: 344px; position: absolute;
            top: 267px" Width="63px">
           </asp:DropDownList>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:TextBox ID="DCode" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 120; left: 343px; position: absolute; top: 338px"
               Width="150px"></asp:TextBox>
           &nbsp; &nbsp;
           <asp:TextBox ID="DMaterial" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 120; left: 606px; position: absolute;
               top: 305px" Width="153px"></asp:TextBox>
           <asp:TextBox ID="DManufout" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 120; left: 606px; position: absolute;
               top: 338px" Width="153px"></asp:TextBox>
           <asp:TextBox ID="DDevNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 120; left: 728px; position: absolute;
               top: 378px" Width="72px"></asp:TextBox>
           &nbsp;
           &nbsp;&nbsp;
           <asp:TextBox ID="DMapNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 121; left: 342px; position: absolute; top: 304px"
               Width="150px"></asp:TextBox>
           <asp:TextBox ID="DAQty" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 121; left: 610px; position: absolute;
               top: 376px" Width="41px"></asp:TextBox>
           <asp:TextBox ID="DQty" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 122; left: 604px; position: absolute; top: 264px"
               Width="43px"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DWeight" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 123; left: 730px; position: absolute;
               top: 264px" Width="62px"></asp:TextBox>
           &nbsp;
           &nbsp;&nbsp;<input id="DFinishDate" runat="server" style="z-index: 142; left: 605px; width: 112px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 230px; background-color: yellow; border-bottom-style: groove"
            type="text" value="2009/12/12" /><input id="DAppDate" runat="server" style="z-index: 142; left: 605px; width: 112px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 157px; background-color: yellow; border-bottom-style: groove"
            type="text" value="2009/12/12" /><input id="DAFinishDate" runat="server" style="z-index: 142; left: 347px; width: 112px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 376px; background-color: yellow; border-bottom-style: groove"
            type="text" value="2009/12/12" />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="DAppper" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 124; left: 347px; position: absolute;
               top: 192px" Width="150px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 128; left: 347px; position: absolute;
            top: 156px" Width="111px"></asp:TextBox>
           &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="DEngine8" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="27px" Style="z-index: 129; left: 71px; position: absolute;
               top: 521px" Width="94px" Font-Size="Larger"></asp:TextBox>
           <asp:TextBox ID="DEngine9" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="27px" Style="z-index: 130; left: 177px; position: absolute;
               top: 520px" Width="94px" Font-Size="Larger"></asp:TextBox>
           <asp:TextBox ID="DEngine10" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="27px" Style="z-index: 131; left: 284px; position: absolute;
               top: 520px" Width="94px" Font-Size="Larger"></asp:TextBox>
           &nbsp; &nbsp;
           &nbsp;&nbsp;&nbsp; &nbsp;<br />
           &nbsp; &nbsp;<br />
        <br />
        <br />
        <br />
           &nbsp;
           <img id="DDescSheet" runat="server" src="images/MapSheet_004.jpg" style="z-index: 99;
               left: 24px; position: absolute; top: 1432px" />
           <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="58px" Style="z-index: 132; left: 71px; position: absolute; top: 1438px"
               TextMode="MultiLine" Width="536px"></asp:TextBox>
           &nbsp;&nbsp;&nbsp;<br />
        <br />
        <br />
        <br />
        <br />
        <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 140; left: 24px;
            width: 592px; position: absolute; top: 1504px; height: 112px" visible="false" />
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
           <br />
        <br />
        <br />
           <asp:Button ID="BEngineDel" runat="server" CausesValidation="False" Style="z-index: 133;
               left: 103px; position: absolute; top: 607px" Text="清除工程" Visible="False" />
        <br />
        <asp:Label ID="DHistoryLabel" runat="server" Font-Names="新細明體" Font-Size="11pt" ForeColor="Blue"
            Height="16px" Style="z-index: 134; left: 16px; position: absolute; top: 1704px"
            Width="64px">核定履歷</asp:Label>
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
            Style="z-index: 135; left: 16px; position: absolute; top: 1720px" Width="780px">
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
        <br />
        <asp:TextBox ID="DReason" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
            Style="z-index: 136; left: 248px; position: absolute; top: 1504px" Visible="False"
            Width="360px"></asp:TextBox>
        <asp:DropDownList ID="DReasonCode" runat="server" Height="20px" Style="z-index: 137;
            left: 184px; position: absolute; top: 1504px" Visible="False" Width="64px">
            <asp:ListItem>01</asp:ListItem>
            <asp:ListItem>02</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 138; left: 176px; position: absolute; top: 1544px"
            Visible="False" Width="432px"></asp:TextBox>
        &nbsp;&nbsp;<br />
         <asp:CustomValidator  ID="CustomValidator1" runat="server" ErrorMessage="CustomValidator"></asp:CustomValidator>
        <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="true" ShowSummary="false" />
        </div>
    </form>
</body>
</html>
