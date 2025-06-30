<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PriceInforSheet_01.aspx.vb" Inherits="PriceInforSheet_01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>未命名頁面</title>
<script language="javascript" type="text/javascript">
<!--
            function ConfirmMe(btn) {
                if(Page_ClientValidate())   {
                    btn.disabled="disabled";
				    var answer = confirm("是否執行作業嗎？ 請確認....");
				    if (answer) {
		                document.cookie="RunBOK=True";
		                document.cookie="RunBNG1=True";
		                document.cookie="RunBNG2=True";
		                document.cookie="RunBSAVE=True";

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
// -->
</script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
            Font-Names="Times New Roman" Font-Size="11pt" Style="z-index: 100; left: 24px;
            position: relative; top: 8px">DNo</asp:TextBox>
        <img src="images/WFS_PriceInfor.PNG" style="z-index: 99; left: 16px; position: absolute;
            top: 56px" id="IMG1" language="javascript" onclick="return IMG1_onclick()" />
        <asp:TextBox ID="DEmpID" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Style="z-index: 102; left: 619px; position: absolute; top: 128px" Text="EmpID"
            Width="93px"></asp:TextBox>
        <asp:TextBox ID="Dremark" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="32px" Style="z-index: 101; left: 144px; position: absolute;
            top: 184px" Width="567px"></asp:TextBox>
        <asp:TextBox ID="DDate" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Style="z-index: 102; left: 144px; position: absolute; top: 128px" Text="Date"
            Width="229px"></asp:TextBox>
        <asp:TextBox ID="DName" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Style="z-index: 103; left: 496px; position: absolute; top: 128px" Text="Name"
            Width="116px"></asp:TextBox>
        <asp:TextBox ID="DJobTitle" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Style="z-index: 104; left: 144px; position: absolute; top: 152px"
            Text="DJobTitle" Width="150px"></asp:TextBox>
        <asp:TextBox ID="DJobCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Style="z-index: 104; left: 301px; position: absolute; top: 152px"
            Text="JobCode" Width="72px"></asp:TextBox>
        &nbsp;&nbsp;
        <asp:TextBox ID="DDepoName" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Style="z-index: 106; left: 496px; position: absolute; top: 152px"
            Text="DepoName" Width="32px"></asp:TextBox>
        <asp:TextBox ID="DDepoCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Style="z-index: 107; left: 535px; position: absolute; top: 152px"
            Text="DepoCode" Width="24px"></asp:TextBox>
        <asp:TextBox ID="DDivision" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Style="z-index: 108; left: 566px; position: absolute; top: 152px"
            Text="Division" Width="88px"></asp:TextBox>
        <asp:TextBox ID="DDivisionCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Style="z-index: 109; left: 661px; position: absolute; top: 152px"
            Text="DivisionCode" Width="50px"></asp:TextBox>
        &nbsp; &nbsp; &nbsp; &nbsp;
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
        <br />
        <asp:GridView ID="grd" runat="server" AllowSorting="True" AutoGenerateColumns="False"
            BackColor="White" Font-Bold="False" Font-Names="Arial" Font-Size="9pt" Height="120px"
            ShowHeader="False" Style="z-index: 106; left: 24px; position: absolute; top: 296px"
            Width="696px">
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <table border="0">
                            <tr>
                                <td style="width: 38px;">
                                    <asp:TextBox ID="ItemNo" runat="server" ReadOnly="true" Text='<%# Eval("No") %>' Width="115px"></asp:TextBox></td>
                                <td style="width: 38px;">
                                    <asp:TextBox ID="Code" runat="server" ReadOnly="true" Text='<%# Eval("Code") %>'
                                        Width="55px"></asp:TextBox></td>
                                <td colspan="8" style="width: 500px;">
                                    <asp:TextBox ID="ItemName" runat="server" ReadOnly="true" Text='<%# Eval("ItemName") %>'
                                        Width="490px"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:TextBox ID="ApplyName" runat="server" ReadOnly="true" Text='<%# Eval("ApplyName") %>'
                                        Width="115px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="Version" runat="server" ReadOnly="true" Text='<%# Eval("Version") %>'
                                        Width="55px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="A1" runat="server" ReadOnly="true" Text='<%# Eval("A1") %>' Width="50px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="B1" runat="server" ReadOnly="true" Text='<%# Eval("B1") %>' Width="50px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="B2" runat="server" ReadOnly="true" Text='<%# Eval("B2") %>' Width="50px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="B4" runat="server" ReadOnly="true" Text='<%# Eval("B4") %>' Width="50px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="C1" runat="server" ReadOnly="true" Text='<%# Eval("C1") %>' Width="50px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="C2" runat="server" ReadOnly="true" Text='<%# Eval("C2") %>' Width="50px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="C3" runat="server" ReadOnly="true" Text='<%# Eval("C3") %>' Width="50px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="D1" runat="server" ReadOnly="true" Text='<%# Eval("D1") %>' Width="50px"></asp:TextBox></td>
                            </tr>
                        </table>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <HeaderStyle BackColor="White" />
            <AlternatingRowStyle BackColor="#FFFF80" />
        </asp:GridView>
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
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <asp:Button ID="BSAVE" runat="server" CausesValidation="False" Height="28px" Style="z-index: 107;
            left: 368px; position: absolute; top: 968px" Text="儲存" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <asp:Button ID="BNG2" runat="server" CausesValidation="False" Height="28px" Style="z-index: 108;
            left: 456px; position: absolute; top: 968px" Text="NG2" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <asp:Button ID="BNG1" runat="server" CausesValidation="False" Height="28px" Style="z-index: 109;
            left: 552px; position: absolute; top: 968px" Text="NG1" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <asp:Button ID="BOK" runat="server" CausesValidation="False" Height="28px" Style="z-index: 110;
            left: 640px; position: absolute; top: 968px" Text="OK" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        &nbsp;<br />
        &nbsp; &nbsp;
        <asp:Label ID="DHistoryLabel" runat="server" Font-Size="11pt" ForeColor="Blue" Height="1px"
            Style="z-index: 115; left: 24px; position: absolute; top: 1264px" Text="核定履歷"></asp:Label>
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
            Style="z-index: 116; left: 16px; position: absolute; top: 1288px" Width="780px">
            <RowStyle BackColor="White" Font-Size="9pt" ForeColor="#330099" />
            <Columns>
                <asp:BoundField DataField="StepNameDesc" HeaderText="工程">
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="DecideName" HeaderText="擔當" />
                <asp:BoundField DataField="AgentName" HeaderText="代理" />
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
      
        <asp:DropDownList ID="DReasonCode" runat="server" Height="20px" Style="z-index: 115;
            left: 184px; position: absolute; top: 1152px" Visible="False" Width="64px">
            <asp:ListItem>01</asp:ListItem>
            <asp:ListItem>02</asp:ListItem>
        </asp:DropDownList>
        <img id="DDescSheet" runat="server" src="images/Sheet_Remark.jpg" style="z-index: 1;
            left: 24px; position: absolute; top: 1064px" />
        <asp:TextBox ID="DReason" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
            Style="z-index: 133; left: 248px; position: absolute; top: 1152px" Visible="False"
            Width="352px"></asp:TextBox>
        <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 1; left: 24px;
            position: absolute; top: 1144px" visible="false" />
        <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 132; left: 72px; position: absolute; top: 1072px"
            TextMode="MultiLine" Width="536px"></asp:TextBox>
        <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 135; left: 176px; position: absolute; top: 1184px"
            Visible="False" Width="424px"></asp:TextBox>

        <!--************************************************************ 
            ** 防止重覆按 [BUTTON 2 次]                               **     
            ************************************************************ -->
            <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 126; left: -500px;position: absolute; top: 100px; text-align: left">AAA</asp:TextBox>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextBox1" ErrorMessage="不可為空白"></asp:RequiredFieldValidator>            
            
    
    </div>
       <asp:CustomValidator  ID="CustomValidator1" runat="server" ErrorMessage="CustomValidator"></asp:CustomValidator>
        <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="true" ShowSummary="false" />
    </form>
</body>
</html>
