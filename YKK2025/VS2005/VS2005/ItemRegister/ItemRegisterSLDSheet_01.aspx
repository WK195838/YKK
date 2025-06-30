<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ItemRegisterSLDSheet_01.aspx.vb" Inherits="ItemRegisterSLDSheet_01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Wave's Item登錄申請單(SLD)</title>
    <script language="javascript" src="RegisterItem.js"></script>   
		<script language="javascript" type="text/javascript">
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
        <img src="images/RegisterSLDSheet_001.jpg" style="z-index: 1; left: 6px; position: absolute;top: 7px" />
        <asp:TextBox ID="DDate" runat="server" AutoPostBack="True" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 116px;
            position: absolute; top: 74px; text-align: left" Width="226px"></asp:TextBox>
        <asp:TextBox ID="DJobTitle" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126;
            left: 116px; position: absolute; top: 100px; text-align: left" Width="226px"></asp:TextBox>
        <asp:TextBox ID="DDivision" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126;
            left: 452px; position: absolute; top: 100px; text-align: left" Width="246px"></asp:TextBox>
        <asp:TextBox ID="DName" runat="server" AutoPostBack="True" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 452px;
            position: absolute; top: 74px; text-align: left" Width="246px"></asp:TextBox>
        <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
            Font-Names="Times New Roman" Font-Size="11pt" Style="z-index: 100; left: 18px;
            position: absolute; top: 14px">DNo</asp:TextBox>
        <img src="images/RegisterSLDSheet_002.jpg" style="z-index: 1; left: 6px; position: absolute;top: 733px" id="DRemarkSheet" runat="server" />
        <asp:FileUpload ID="DAttachfile1" runat="server" style="z-index: 121; left: 114px;
            width: 588px; position: absolute; top: 790px; height: 24px; background-color: #C0FFFF"/>
        <asp:HyperLink ID="LAttachfile1" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 124; left: 122px; position: absolute; top: 794px" Target="_blank"
            Width="100px">參考文件</asp:HyperLink>
        <asp:TextBox ID="DRemark" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="43px" Style="z-index: 126; left: 115px;
            position: absolute; top: 740px; text-align: left" Width="583px" MaxLength="100" TextMode="MultiLine"></asp:TextBox>
        <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 1; left: 7px;
            position: absolute; top: 893px" visible="false" />
        <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 132; left: 54px; position: absolute; top: 825px"
            TextMode="MultiLine" Width="536px"></asp:TextBox>
        <asp:TextBox ID="DReason" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
            Style="z-index: 133; left: 238px; position: absolute; top: 899px" Visible="False"
            Width="352px"></asp:TextBox>
        <asp:DropDownList ID="DReasonCode" runat="server" Height="20px" Style="z-index: 134;
            left: 162px; position: absolute; top: 902px" Visible="False" Width="64px" BackColor="Yellow">
            <asp:ListItem>01</asp:ListItem>
            <asp:ListItem>02</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 135; left: 166px; position: absolute; top: 932px"
            Visible="False" Width="424px"></asp:TextBox>
        <asp:Label ID="DHistoryLabel" runat="server" Font-Size="11pt" ForeColor="Blue" Height="20px"
            Style="z-index: 137; left: 9px; position: absolute; top: 1016px" Text="核定履歷"></asp:Label>
        <img id="DDescSheet" runat="server" src="images/Sheet_Remark.jpg" style="z-index: 1;
            left: 7px; position: absolute; top: 819px" />
           
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#400040"
            BorderStyle="Groove" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 9px; position: absolute; top: 129px" Width="850px" HorizontalAlign="Left">
            <Columns>
                <asp:HyperLinkField DataNavigateUrlFields="FNoURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="FNo" HeaderText="No." Target="_blank">
                </asp:HyperLinkField>
                <asp:BoundField DataField="ApplyType" HeaderText="動作" >
                </asp:BoundField>
                <asp:BoundField DataField="Code" HeaderText="Code" >
                </asp:BoundField>
                <asp:BoundField DataField="ItemName" HeaderText="ItemName" >
                </asp:BoundField>
                <asp:BoundField DataField="DTCode" HeaderText="Code" >
                </asp:BoundField>
                <asp:BoundField DataField="DTItemName" HeaderText="ItemName(SLD)" >
                </asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
            Style="z-index: 136; left: 8px; position: absolute; top: 1038px" Width="780px">
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
        <asp:Button ID="BSAVE" runat="server" Text="SAVE" Style="z-index: 130; left: 349px; position: absolute; top: 1002px" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <asp:Button ID="BNG1" runat="server" Text="NG1" Style="z-index: 130; left: 440px; position: absolute; top: 1002px" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <asp:Button ID="BNG2" runat="server" Text="NG2" Style="z-index: 130; left: 532px; position: absolute; top: 1002px" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <asp:Button ID="BOK" runat="server" Text="OK" Style="z-index: 130; left: 624px; position: absolute; top: 1002px" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        &nbsp;

        <!--************************************************************ 
            ** 防止重覆按 [BUTTON 2 次]                               **     
            ************************************************************ -->
            <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 126; left: -500px;position: absolute; top: 100px; text-align: left">AAA</asp:TextBox>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextBox1" ErrorMessage="不可為空白"></asp:RequiredFieldValidator>            

        </div>
    </form>
</body>
</html>
