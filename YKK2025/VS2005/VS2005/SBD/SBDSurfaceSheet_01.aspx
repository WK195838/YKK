<%@ Page Language="vb" AutoEventWireup="false" Inherits="SBDSurfaceSheet_01" aspCompat="True" CodeFile="SBDSurfaceSheet_01.aspx.vb" %>
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
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/SBD-SufaceSheet01.jpg" style="z-index: 100; left: 10px; position: absolute; top: 15px" />
           <asp:Image ID="LCustSampleFile" runat="server" BorderStyle="Groove" Height="240px"
               ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"
               Style="z-index: 101; left: 24px; position: absolute; top: 128px" Width="200px" />
           <asp:Image ID="LFinalSampleFile" runat="server" BorderStyle="Groove" Height="176px"
               ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"
               Style="z-index: 102; left: 24px; position: absolute; top: 432px" Width="200px" />
           &nbsp;
           <input id="DQCReqFile" runat="server" name="UploadFile" style="z-index: 138; left: 344px;
            width: 456px; position: absolute; top: 568px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
           <input id="DCustSampleFile" runat="server" name="UploadFile" style="z-index: 139; left: 24px;
            width: 200px; position: absolute; top: 368px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
        <input id="DFinalSampleFile" runat="server" name="UploadFile" style="z-index: 137; left: 24px;
            width: 200px; position: absolute; top: 608px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
        <input id="DForcastFile" runat="server" name="UploadFile" style="z-index: 148; left: 528px;
            width: 264px; position: absolute; top: 720px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
          &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:HyperLink ID="LForcastFile" runat="server" Style="z-index: 103; left: 528px; position: absolute;
               top: 720px" Target="_blank" Width="168px">報價單</asp:HyperLink>
           <asp:HyperLink ID="LQAFile" runat="server" Style="z-index: 104; left: 344px; position: absolute;
               top: 608px" Target="_blank" Width="368px">品測報告書</asp:HyperLink>
           &nbsp; &nbsp;
           <asp:HyperLink ID="LContactFile" runat="server" Style="z-index: 105; left: 528px; position: absolute;
               top: 752px" Target="_blank" Width="168px">切結書</asp:HyperLink>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:HyperLink ID="LManufFlowFile" runat="server" Style="z-index: 106; left: 136px;
               position: absolute; top: 720px" Target="_blank" Width="176px">製造流程表</asp:HyperLink>
           <asp:HyperLink ID="LOPManualFile" runat="server" Style="z-index: 107; left: 136px;
               position: absolute; top: 752px" Target="_blank" Width="184px">作業標準書</asp:HyperLink>
           &nbsp;
           <asp:HyperLink ID="LQCReqFile" runat="server" Style="z-index: 108; left: 344px; position: absolute;
               top: 568px" Target="_blank" Width="360px">品質依賴書</asp:HyperLink>
           &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;&nbsp;
        <input id="DQAFile" runat="server" name="UploadFile" style="z-index: 149; left: 344px;
            width: 456px; position: absolute; top: 608px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
           &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          
        <input id="DContactFile" runat="server" name="UploadFile" style="z-index: 152; left: 528px;
            width: 264px; position: absolute; top: 752px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp;&nbsp;&nbsp;
        <input id="BSAVE" runat="server" name="Button1" onclick="Button('SAVE');" style="z-index: 156;
            left: 312px; width: 80px; position: absolute; top: 1368px; height: 28px" type="button"
            value="儲存" />
        <input id="BNG2" runat="server" name="Button1" onclick="Button('NG2');" style="z-index: 155;
            left: 400px; width: 80px; position: absolute; top: 1368px; height: 28px" type="button"
            value="NG2" />
        <input id="BNG1" runat="server" name="Button1" onclick="Button('NG1');" style="z-index: 154;
            left: 488px; width: 80px; position: absolute; top: 1368px; height: 28px" type="button"
            value="NG1" />
        &nbsp;&nbsp;
        <input id="BOK" runat="server" name="BOK" onclick="Button('OK');" style="z-index: 153;
            left: 576px; width: 80px; position: absolute; top: 1368px; height: 28px" type="button"
            value="OK" />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp;
        <input id="BDate3" runat="server" style="z-index: 144; left: 208px; width: 20px;
            position: absolute; top: 672px" type="button" value="..." />
        <input id="BDate1" runat="server" style="z-index: 147; left: 744px; width: 20px;
            position: absolute; top: 264px" type="button" value="..." /><input id="BDate4" runat="server" style="z-index: 145; left: 448px; width: 20px;
            position: absolute; top: 200px" type="button" value="..." /><input id="BDate2" runat="server" style="z-index: 146; left: 440px; width: 20px;
            position: absolute; top: 472px" type="button" value="..." />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <input id="DManufFlowFile" runat="server" name="UploadFile" style="z-index: 151; left: 136px;
            width: 256px; position: absolute; top: 720px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
           <input id="DOPManualFile" runat="server" name="UploadFile" style="z-index: 150; left: 136px;
            width: 256px; position: absolute; top: 752px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:DropDownList ID="DBuyer" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 109; left: 336px; position: absolute;
            top: 164px" Width="168px">
           </asp:DropDownList>
           <asp:DropDownList ID="DAttachSample" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 110; left: 632px; position: absolute;
            top: 230px" Width="96px">
           </asp:DropDownList>
           <asp:DropDownList ID="DSampleQty" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 111; left: 624px; position: absolute;
            top: 328px" Width="112px">
           </asp:DropDownList>
           <asp:DropDownList ID="DAllowSample" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 112; left: 336px; position: absolute;
            top: 436px" Width="112px">
           </asp:DropDownList><asp:DropDownList ID="DSupplier" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 113; left: 336px; position: absolute;
            top: 404px" Width="464px">
           </asp:DropDownList>
           &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:DropDownList ID="DQCResult" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 114; left: 240px; position: absolute;
            top: 672px" Width="80px">
        </asp:DropDownList>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <input id="DQCDate" runat="server" style="z-index: 143; left: 112px; width: 96px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 672px; background-color: yellow; border-bottom-style: groove"
            type="text" value="2009/12/12" />
           <input id="DBFinalDate" runat="server" style="z-index: 141; left: 336px; width: 104px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 472px; background-color: yellow; border-bottom-style: groove"
            type="text" value="2009/12/12" />
           &nbsp; &nbsp; &nbsp; &nbsp;
        <asp:TextBox ID="DSellVendor" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 115; left: 632px; position: absolute;
            top: 164px" Width="168px"></asp:TextBox>
        <asp:TextBox ID="DDivision" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 116; left: 336px; position: absolute;
            top: 128px" Width="176px"></asp:TextBox>
        <asp:TextBox ID="DReqQty" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="24px" Style="z-index: 117; left: 632px; position: absolute; top: 196px"
            Width="168px"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DCode" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 118; left: 336px; position: absolute; top: 230px"
               Width="176px"></asp:TextBox>
           <asp:TextBox ID="DORNO" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 119; left: 336px; position: absolute; top: 264px"
               Width="176px"></asp:TextBox>
           <asp:TextBox ID="DColor" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 120; left: 336px; position: absolute; top: 328px"
               Width="176px"></asp:TextBox>
           &nbsp;&nbsp;<input id="DOrderTime" runat="server" style="z-index: 142; left: 632px; width: 112px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 264px; background-color: yellow; border-bottom-style: groove"
            type="text" value="2009/12/12" /><input id="DReqDelDate" runat="server" style="z-index: 140; left: 336px; width: 112px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 200px; background-color: yellow; border-bottom-style: groove"
            type="text" value="2009/12/12" />
           &nbsp;&nbsp;
        <asp:TextBox ID="DDevReason" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 121; left: 336px; position: absolute;
            top: 296px" Width="464px"></asp:TextBox>
        <asp:TextBox ID="DCap" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 122; left: 624px; position: absolute;
            top: 436px" Width="120px"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DSchedule" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 123; left: 624px; position: absolute;
               top: 472px" Width="120px"></asp:TextBox>
           &nbsp;&nbsp;
           <asp:TextBox ID="DEnglishName" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 124; left: 336px; position: absolute;
               top: 360px" Width="464px"></asp:TextBox>
           &nbsp;
        <asp:TextBox ID="DQCRemark" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 125; left: 336px; position: absolute;
            top: 672px" Width="464px"></asp:TextBox>
        <asp:TextBox ID="DAppper" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 126; left: 632px; position: absolute;
            top: 128px" Width="168px"></asp:TextBox>
        &nbsp; &nbsp;
        <asp:TextBox ID="DAppDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 127; left: 632px; position: absolute;
            top: 96px" Width="168px"></asp:TextBox>
        <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 128; left: 336px; position: absolute;
            top: 96px" Width="176px"></asp:TextBox>
           &nbsp;<br />
           <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="58px" Style="z-index: 129; left: 72px; position: absolute; top: 1432px"
               TextMode="MultiLine" Width="536px"></asp:TextBox>
        <br />
        <br />
        <br />
        <br />
           &nbsp;
           <img id="DDescSheet" runat="server" src="images/MapSheet_004.jpg" style="z-index: 99;
               left: 24px; position: absolute; top: 1424px" />
        <br />
        <br />
        <br />
        <br />
        <br />
        <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 136; left: 24px;
            width: 592px; position: absolute; top: 1504px; height: 112px" visible="false" />
        &nbsp;<br />
        <br />
        <br />
        <br />
        <asp:Label ID="DHistoryLabel" runat="server" Font-Names="新細明體" Font-Size="11pt" ForeColor="Blue"
            Height="16px" Style="z-index: 130; left: 16px; position: absolute; top: 1704px"
            Width="64px">核定履歷</asp:Label>
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
            Style="z-index: 131; left: 16px; position: absolute; top: 1720px" Width="780px">
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
            Style="z-index: 132; left: 248px; position: absolute; top: 1504px" Visible="False"
            Width="360px"></asp:TextBox>
        <asp:DropDownList ID="DReasonCode" runat="server" Height="20px" Style="z-index: 133;
            left: 184px; position: absolute; top: 1504px" Visible="False" Width="64px">
            <asp:ListItem>01</asp:ListItem>
            <asp:ListItem>02</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 134; left: 176px; position: absolute; top: 1544px"
            Visible="False" Width="432px"></asp:TextBox>
        &nbsp;&nbsp;<br />
         <asp:CustomValidator  ID="CustomValidator1" runat="server" ErrorMessage="CustomValidator"></asp:CustomValidator>
        <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="true" ShowSummary="false" />
        </div>
    </form>
</body>
</html>
