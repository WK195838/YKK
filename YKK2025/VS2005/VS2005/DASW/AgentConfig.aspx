<%@ Page Language="VB" AutoEventWireup="false" CodeFile="AgentConfig.aspx.vb" Inherits="AgentConfig" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>代理人設定</title>
    <style type="text/css">
    body { font-size:12px;}
    </style>
		<script language="javascript" type="text/javascript">
            function OpenDatePicker(xDepo, xField)
            {
                window.open('http://10.245.1.10/SCD/DatePicker.aspx?field1='+xField+'&pDepo='+xDepo,'','status=0,toolbar=0,width=300,height=200');
            }
            
            	function calendarPicker(xDepo,field1)
			{
				window.open('DatePicker.aspx?field1=' + field1+'&pDepo='+xDepo,'','status=0,toolbar=0,width=300,height=200');
			}
		</script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <!-- 表單Image -->
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/Agent_01.jpg"
             Style="z-index: 99; left: 2px; position: absolute; top: 2px" BorderStyle="None" />
        <!-- -->
        <!-- 個人 -->
        <asp:TextBox ID="DName" runat="server" BackColor="#C0FFFF" BorderStyle="Double"
            ForeColor="Black" Height="22px" Style="z-index: 126; left: 98px;
            position: absolute; top: 57px; text-align: left" Width="130px" BorderWidth="1px"></asp:TextBox>
        <asp:TextBox ID="DUserID" runat="server" BackColor="#C0FFFF" BorderStyle="Double"
            BorderWidth="1px" ForeColor="Black" Height="22px" Style="z-index: 126; left: 233px;
            position: absolute; top: 57px; text-align: left" Width="130px"></asp:TextBox>
        <!-- -->
        <!-- 最終設定時間 -->
        <asp:TextBox ID="DLastDate" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            BorderWidth="1px" ForeColor="Black" Height="22px" Style="z-index: 126; left: 459px;
            position: absolute; top: 57px; text-align: left" Width="130px"></asp:TextBox>
        <asp:TextBox ID="DLastTime" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            BorderWidth="1px" ForeColor="Black" Height="22px" Style="z-index: 126; left: 459px;
            position: absolute; top: 84px; text-align: left" Width="130px"></asp:TextBox>
        <!-- -->
        <!-- 代理人 -->
        <asp:TextBox ID="DAgentName" runat="server" BackColor="#C0FFFF" BorderStyle="Double"
            ForeColor="Black" Height="22px" Style="z-index: 126; left: 98px;
            position: absolute; top: 101px; text-align: left" Width="130px" BorderWidth="1px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DAgentUserID" runat="server" BackColor="#C0FFFF" BorderStyle="Double"
            ForeColor="Black" Height="22px" Style="z-index: 126; left: 233px;
            position: absolute; top: 101px; text-align: left" Width="130px" BorderWidth="1px"></asp:TextBox>
        <!-- -->
        <!-- 期間(起) -->
        <asp:TextBox ID="DFromDate" runat="server" BackColor="LightGray" BorderStyle="Double"
            ForeColor="Black" Height="18px" Style="z-index: 126; left: 98px;
            position: absolute; top: 150px; text-align: left" Width="70px" BorderWidth="1px" EnableViewState="False">2011/12/12</asp:TextBox>
        <asp:Button ID="BFromDate" runat="server" Height="24px" Style="z-index: 111; left: 172px;
            position: absolute; top: 149px" Text="....." Width="24px" />
        <asp:DropDownList ID="DFromHH" runat="server" BackColor="LightGray" ForeColor="Black"
            Height="24px" Style="z-index: 266; left: 196px; position: absolute; top: 150px"
            Width="43px">
            <asp:ListItem Selected="True">00</asp:ListItem>
            <asp:ListItem>01</asp:ListItem>
            <asp:ListItem>02</asp:ListItem>
            <asp:ListItem>03</asp:ListItem>
            <asp:ListItem>04</asp:ListItem>
            <asp:ListItem>05</asp:ListItem>
            <asp:ListItem>06</asp:ListItem>
            <asp:ListItem>07</asp:ListItem>
            <asp:ListItem>08</asp:ListItem>
            <asp:ListItem>09</asp:ListItem>
            <asp:ListItem>10</asp:ListItem>
            <asp:ListItem>11</asp:ListItem>
            <asp:ListItem>12</asp:ListItem>
            <asp:ListItem>13</asp:ListItem>
            <asp:ListItem>14</asp:ListItem>
            <asp:ListItem>15</asp:ListItem>
            <asp:ListItem>16</asp:ListItem>
            <asp:ListItem>17</asp:ListItem>
            <asp:ListItem>18</asp:ListItem>
            <asp:ListItem>19</asp:ListItem>
            <asp:ListItem>20</asp:ListItem>
            <asp:ListItem>21</asp:ListItem>
            <asp:ListItem>22</asp:ListItem>
            <asp:ListItem>23</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DFromMM" runat="server" BackColor="LightGray" ForeColor="Black"
            Height="24px" Style="z-index: 266; left: 240px; position: absolute; top: 150px"
            Width="43px">
            <asp:ListItem Selected="True">00</asp:ListItem>
            <asp:ListItem>10</asp:ListItem>
            <asp:ListItem>20</asp:ListItem>
            <asp:ListItem>30</asp:ListItem>
            <asp:ListItem>40</asp:ListItem>
            <asp:ListItem>50</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DFromSS" runat="server" BackColor="LightGray" ForeColor="Black"
            Height="24px" Style="z-index: 266; left: 283px; position: absolute; top: 150px"
            Width="43px">
            <asp:ListItem Selected="True">00</asp:ListItem>
            <asp:ListItem>10</asp:ListItem>
            <asp:ListItem>20</asp:ListItem>
            <asp:ListItem>30</asp:ListItem>
            <asp:ListItem>40</asp:ListItem>
            <asp:ListItem>50</asp:ListItem>
        </asp:DropDownList>
        <!-- -->
        
        <asp:TextBox ID="DMark" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
            Height="12px" ReadOnly="True" Style="z-index: 126; left: 328px; position: absolute;
            top: 154px; text-align: left" Width="16px">～</asp:TextBox>

        <!-- 期間(迄) -->
        <asp:TextBox ID="DToDate" runat="server" BackColor="LightGray" BorderStyle="Double"
            ForeColor="Black" Height="18px" Style="z-index: 126; left: 349px;
            position: absolute; top: 150px; text-align: left" Width="70px" BorderWidth="1px" EnableViewState="False">2011/12/12</asp:TextBox>
        <asp:Button ID="BToDate" runat="server" Height="24px" Style="z-index: 111; left: 423px;
            position: absolute; top: 149px" Text="....." Width="24px" />
        <asp:DropDownList ID="DToHH" runat="server" BackColor="LightGray" ForeColor="Black"
            Height="24px" Style="z-index: 266; left: 447px; position: absolute; top: 150px"
            Width="43px">
            <asp:ListItem Selected="True">00</asp:ListItem>
            <asp:ListItem>01</asp:ListItem>
            <asp:ListItem>02</asp:ListItem>
            <asp:ListItem>03</asp:ListItem>
            <asp:ListItem>04</asp:ListItem>
            <asp:ListItem>05</asp:ListItem>
            <asp:ListItem>06</asp:ListItem>
            <asp:ListItem>07</asp:ListItem>
            <asp:ListItem>08</asp:ListItem>
            <asp:ListItem>09</asp:ListItem>
            <asp:ListItem>10</asp:ListItem>
            <asp:ListItem>11</asp:ListItem>
            <asp:ListItem>12</asp:ListItem>
            <asp:ListItem>13</asp:ListItem>
            <asp:ListItem>14</asp:ListItem>
            <asp:ListItem>15</asp:ListItem>
            <asp:ListItem>16</asp:ListItem>
            <asp:ListItem>17</asp:ListItem>
            <asp:ListItem>18</asp:ListItem>
            <asp:ListItem>19</asp:ListItem>
            <asp:ListItem>20</asp:ListItem>
            <asp:ListItem>21</asp:ListItem>
            <asp:ListItem>22</asp:ListItem>
            <asp:ListItem>23</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DToMM" runat="server" BackColor="LightGray" ForeColor="Black"
            Height="24px" Style="z-index: 266; left: 491px; position: absolute; top: 150px"
            Width="43px">
            <asp:ListItem Selected="True">00</asp:ListItem>
            <asp:ListItem>10</asp:ListItem>
            <asp:ListItem>20</asp:ListItem>
            <asp:ListItem>30</asp:ListItem>
            <asp:ListItem>40</asp:ListItem>
            <asp:ListItem>50</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DToSS" runat="server" BackColor="LightGray" ForeColor="Black"
            Height="24px" Style="z-index: 266; left: 535px; position: absolute; top: 150px"
            Width="43px">
            <asp:ListItem Selected="True">00</asp:ListItem>
            <asp:ListItem>10</asp:ListItem>
            <asp:ListItem>20</asp:ListItem>
            <asp:ListItem>30</asp:ListItem>
            <asp:ListItem>40</asp:ListItem>
            <asp:ListItem>50</asp:ListItem>
        </asp:DropDownList>
        <!-- -->
        <!-- 說明 -->
        <asp:TextBox ID="DDescription" runat="server" BackColor="LightGray" BorderStyle="Double"
            ForeColor="Black" Height="35px" Style="z-index: 126; left: 98px; position: absolute;
            top: 181px; text-align: left" TextMode="MultiLine" Width="496px" BorderWidth="1px" MaxLength="200"></asp:TextBox>
        <!-- -->
        <!-- 按鈕-啟動設定 -->
        <asp:Button ID="BSave" runat="server" Height="23px" Style="z-index: 131; left: 521px;
            position: absolute; top: 224px" Text="啟動設定" Width="80px" />
        <!-- -->
    </div>
    </form>
</body>
</html>
