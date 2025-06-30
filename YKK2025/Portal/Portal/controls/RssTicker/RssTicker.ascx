<%@ Control Language="c#" AutoEventWireup="false" Codebehind="RssTicker.ascx.cs" Inherits="Speerio.DNN.Controls.UI.RssTicker.RssTicker" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ OutputCache Duration="600" Shared="true" VaryByParam="None" %>
<style>
.Ticker {
  background-color: #78A2B0;
  width: 100%;
  border: #e0e0e0 1px solid;
  font-family: Verdana,Tahoma, Arial, Helvetica; 
  font-size: 8pt; 
  font-style: normal; 
  padding: 4px;
  height: 24px;
}
</style>
<input type="text" id="dnn_ticker" onclick="goURL();" class="Ticker" onmouseover="this.style.cursor='hand'" onmouseout="this.style.cursor='default'">
<asp:literal id="ItemData" runat="server" />
<script language="Javascript">
count=0;
var ticker_speed=5;
var ticker = document.getElementById("dnn_ticker");
function startTicker()
{
  if(count<ticker_msg.length)
  {
    ticker.value=ticker_msg[count]; 
    count++;
    if(count==ticker_msg.length) count=0;
    setTimeout("startTicker();",ticker_speed*1000);
  }
}

function goURL()
{
  if(ticker.value==ticker_msg[ticker_msg.length-1])
     location.href=ticker_url[ticker_msg.length-1];
  else
     location.href=ticker_url[count-1];
}

startTicker();
</script>
