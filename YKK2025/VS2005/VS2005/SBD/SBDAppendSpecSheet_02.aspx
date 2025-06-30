<%@ Page Language="vb" AutoEventWireup="false" Inherits="SBDAppendSpecSheet_02" aspCompat="True" CodeFile="SBDAppendSpecSheet_02.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>S & B開發委託書</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(field1)
			{
				window.open('DatePicker.aspx?field1=' + field1,'calendarPopup','width=250,height=190,resizable=yes');
			}
		 
		 
		 			function SBDCommissionNoPicker(field1)
			{
				window.open('SBDCommissionNoPicker.aspx?field1=' + field1,'calendarPopup','width=300,height=600,resizable=yes');
			}

 			function SBDBSurfNoPicker(field1)
			{
				window.open('SBDSurfNoPicker.aspx?field1=' + field1,'calendarPopup','width=300,height=600,resizable=yes');
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
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/SBD-AppendSpecSheet01.jpg" style="z-index: 99; left: 10px; position: absolute; top: 15px" />&nbsp;
           <asp:Image ID="LFinalSampleFile" runat="server" BorderStyle="Groove" Height="146px"
               ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"
               Style="z-index: 101; left: 30px; position: absolute; top: 332px" Width="141px" />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:HyperLink ID="LForcastFile" runat="server" Style="z-index: 102; left: 461px; position: absolute;
               top: 504px" Target="_blank" Width="113px">報價單</asp:HyperLink>
           <asp:HyperLink ID="LQAFile" runat="server" Style="z-index: 103; left: 336px; position: absolute;
               top: 365px" Target="_blank" Width="241px">品測報告書</asp:HyperLink>
           &nbsp; &nbsp;
           <asp:HyperLink ID="LContactFile" runat="server" Style="z-index: 104; left: 462px; position: absolute;
               top: 535px" Target="_blank" Width="113px">切結書</asp:HyperLink>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:HyperLink ID="LManufFlowFile" runat="server" Style="z-index: 105; left: 156px;
               position: absolute; top: 501px" Target="_blank" Width="112px">製造流程表</asp:HyperLink>
           <asp:HyperLink ID="LOPManualFile" runat="server" Style="z-index: 106; left: 149px;
               position: absolute; top: 533px" Target="_blank" Width="122px">作業標準書</asp:HyperLink>
           &nbsp;
           <asp:HyperLink ID="LQCReqFile" runat="server" Style="z-index: 107; left: 336px; position: absolute;
               top: 336px" Target="_blank" Width="240px">品質依賴書</asp:HyperLink>
           &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp;&nbsp;&nbsp;
        <input id="BSAVE" runat="server" name="Button1" onclick="Button('SAVE');" style="z-index: 133;
            left: 312px; width: 80px; position: absolute; top: 1368px; height: 28px" type="button"
            value="儲存" />
        <input id="BNG2" runat="server" name="Button1" onclick="Button('NG2');" style="z-index: 132;
            left: 400px; width: 80px; position: absolute; top: 1368px; height: 28px" type="button"
            value="NG2" />
        <input id="BNG1" runat="server" name="Button1" onclick="Button('NG1');" style="z-index: 131;
            left: 488px; width: 80px; position: absolute; top: 1368px; height: 28px" type="button"
            value="NG1" />
        &nbsp;&nbsp;
        <input id="BOK" runat="server" name="BOK" onclick="Button('OK');" style="z-index: 130;
            left: 576px; width: 80px; position: absolute; top: 1368px; height: 28px" type="button"
            value="OK" />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
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
           <input id="DQCDate" runat="server" style="z-index: 122; left: 336px; width: 96px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 426px; background-color: yellow; border-bottom-style: groove"
            type="text" value="2009/12/12" readonly="readOnly" />
           &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        <asp:TextBox ID="DDivision" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 108; left: 152px; position: absolute;
            top: 130px" Width="171px" ReadOnly="True"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp;&nbsp;&nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DQCRemark" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 109; left: 337px; position: absolute;
            top: 459px" Width="300px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DQCResult" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" ReadOnly="True" Style="z-index: 110; left: 464px;
               position: absolute; top: 428px" Width="169px"></asp:TextBox>
        <asp:TextBox ID="DAppper" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 111; left: 462px; position: absolute;
            top: 128px" Width="168px" ReadOnly="True"></asp:TextBox>
        &nbsp; &nbsp;
        <asp:TextBox ID="DAppDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 112; left: 462px; position: absolute;
            top: 96px" Width="168px" ReadOnly="True"></asp:TextBox>
           &nbsp;
           <input id="DNo" runat="server" style="z-index: 123; left: 151px; width: 171px;
               color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
               position: absolute; top: 97px; background-color: yellow; border-bottom-style: groove"
               type="text" readonly="readOnly" /><input id="DBuyer" runat="server" style="z-index: 127; left: 152px; width: 170px;
               color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
               position: absolute; top: 161px; background-color: yellow; border-bottom-style: groove"
               type="text" readonly="readOnly" />
           <input id="DVendor" runat="server" style="z-index: 126; left: 460px; width: 170px;
               color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
               position: absolute; top: 158px; background-color: yellow; border-bottom-style: groove"
               type="text" readonly="readOnly" /><input id="DSupplier" runat="server" style="z-index: 128; left: 150px; width: 170px;
               color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
               position: absolute; top: 193px; background-color: yellow; border-bottom-style: groove"
               type="text" readonly="readOnly" /><input id="DSurfSupplier" runat="server" style="z-index: 124; left: 465px; width: 170px;
               color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
               position: absolute; top: 224px; background-color: yellow; border-bottom-style: groove"
               type="text" readonly="readOnly" /><input id="DCap" runat="server" style="z-index: 121; left: 150px; width: 170px;
               color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
               position: absolute; top: 257px; background-color: yellow; border-bottom-style: groove"
               type="text" readonly="readOnly" />
           <input id="DSchedule" runat="server" style="z-index: 125; left: 464px; width: 170px;
               color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
               position: absolute; top: 259px; background-color: yellow; border-bottom-style: groove"
               type="text" readonly="readOnly" /><input id="DSurfSheetNo" runat="server" style="z-index: 129; left: 151px; width: 168px;
               color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
               position: absolute; top: 225px; background-color: yellow; border-bottom-style: groove"
               type="text" readonly="readOnly" />
           &nbsp; &nbsp;&nbsp;<br />
           <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="58px" Style="z-index: 113; left: 72px; position: absolute; top: 1432px"
               TextMode="MultiLine" Width="536px"></asp:TextBox>
        <br />
        <br />
        <br />
        <br />
           &nbsp;
           <img id="DDescSheet" runat="server" src="images/MapSheet_004.jpg" style="z-index: 119;
               left: 24px; position: absolute; top: 1424px" />
        <br />
        <br />
        <br />
        <br />
        <br />
        <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 120; left: 24px;
            width: 592px; position: absolute; top: 1504px; height: 112px" visible="false" />
        &nbsp;<br />
        <br />
        <br />
        <br />
        <asp:Label ID="DHistoryLabel" runat="server" Font-Names="新細明體" Font-Size="11pt" ForeColor="Blue"
            Height="16px" Style="z-index: 114; left: 16px; position: absolute; top: 1704px"
            Width="64px">核定履歷</asp:Label>
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
            Style="z-index: 115; left: 16px; position: absolute; top: 1720px" Width="780px">
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
            Style="z-index: 116; left: 248px; position: absolute; top: 1504px" Visible="False"
            Width="360px"></asp:TextBox>
        <asp:DropDownList ID="DReasonCode" runat="server" Height="20px" Style="z-index: 117;
            left: 184px; position: absolute; top: 1504px" Visible="False" Width="64px">
            <asp:ListItem>01</asp:ListItem>
            <asp:ListItem>02</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 118; left: 176px; position: absolute; top: 1544px"
            Visible="False" Width="432px"></asp:TextBox>
        &nbsp;&nbsp;<br />
         <asp:CustomValidator  ID="CustomValidator1" runat="server" ErrorMessage="CustomValidator"></asp:CustomValidator>
        <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="true" ShowSummary="false" />
        </div>
    </form>
</body>
</html>
