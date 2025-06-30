<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        表 單=<asp:TextBox ID="DFormNo" runat="server" BackColor="Yellow" Width="172px">002002</asp:TextBox>
        &nbsp;
        單 號=<asp:TextBox ID="DFormSno" runat="server" BackColor="Yellow" Width="172px">1</asp:TextBox>&nbsp;
        工 程=<asp:TextBox ID="DStep" runat="server" BackColor="Yellow" Width="172px">35</asp:TextBox>&nbsp;
        序 號=<asp:TextBox ID="DSeqno" runat="server" BackColor="Yellow" Width="172px">1</asp:TextBox><br />
        Ｄｅｐｏ =<asp:TextBox ID="DDepo" runat="server" BackColor="Yellow" Width="172px"></asp:TextBox>&nbsp;
        簽核者 =<asp:TextBox ID="DUser" runat="server" BackColor="Yellow" Width="172px">ea009</asp:TextBox>&nbsp;
        下一簽核者=<asp:TextBox ID="DNextUser" runat="server" BackColor="Yellow" Width="172px"></asp:TextBox><br />
        被代理者 =<asp:TextBox ID="DAgentID" runat="server" BackColor="Yellow" Width="172px"></asp:TextBox>&nbsp;
        申請者 =<asp:TextBox ID="DApplyID" runat="server" BackColor="Yellow" Width="172px">it003</asp:TextBox>&nbsp;
        預定開始日=<asp:TextBox ID="DStartTime" runat="server" BackColor="Yellow" Width="172px">2010/1/4 16:00:00</asp:TextBox><br />
        預定完成日=<asp:TextBox ID="DEndTime" runat="server" BackColor="Yellow" Width="172px">2010/1/4 16:30:00</asp:TextBox>&nbsp;
        重要性 =<asp:TextBox ID="DImportant" runat="server" BackColor="Yellow" Width="172px">0</asp:TextBox>&nbsp;
        被指定者 =<asp:TextBox ID="DAllocateID" runat="server" BackColor="Yellow" Width="172px"></asp:TextBox><br />
        多工程No =<asp:TextBox ID="DMultiJob" runat="server" BackColor="Yellow" Width="172px">1</asp:TextBox>&nbsp;
        動作 =
        <asp:TextBox ID="DAction" runat="server" BackColor="Yellow" Width="172px">0</asp:TextBox>&nbsp;
        難易度 =<asp:TextBox ID="DLevel" runat="server" BackColor="Yellow" Width="172px"></asp:TextBox><br />
        QC-L/T =<asp:TextBox ID="DQCLT" runat="server" BackColor="Yellow" Width="172px">0</asp:TextBox>&nbsp;
        送件者=<asp:TextBox ID="DFrom" runat="server" BackColor="Yellow" Width="172px">0</asp:TextBox>&nbsp;
        收件者=<asp:TextBox ID="DTo" runat="server" BackColor="Yellow" Width="172px"></asp:TextBox>&nbsp;
        <br />
        訊息類別=<asp:TextBox ID="DMsgType" runat="server" BackColor="Yellow" Width="172px"></asp:TextBox>&nbsp;
        No=<asp:TextBox ID="DNo" runat="server" BackColor="Yellow" Width="172px">0</asp:TextBox><br />
        <br />
        Return Code=<asp:TextBox ID="DRtnCode" runat="server" BackColor="Yellow" Width="172px"></asp:TextBox><br />
        Error &nbsp; Code=<asp:TextBox ID="DErrorCode" runat="server" BackColor="Yellow" Width="172px"></asp:TextBox><br />
        <br />
        <asp:Button ID="BUPDATEFC" runat="server" Text="BUPDATEFC" /><br />


        <asp:Button ID="BGetFormSeqNo" runat="server" Text="GetFormSeqNo" />表單<br />
        <asp:Button ID="BNewFlow" runat="server" Text="NewFlow" />表單,單號,工程,序號,Depo,簽核者,申請者<br />
        <asp:Button ID="BNextFlow" runat="server" Text="NextFlow" />表單,單號,工程,序號,Depo,簽核者,下一簽核者,被代理者,申請者,預定開始日,預定完成日,重要性
        <br />
        <asp:Button ID="BEndFlow" runat="server" Text="EndFlow" />表單,單號,工程,序號,Depo,簽核者,申請者<br />
        <asp:Button ID="Button1" runat="server" Text="NextGate" />表單,工程,簽核者,申請者,被代理者,被指定者,多工程No,動作<br />
        <asp:Button ID="Button2" runat="server" Text="GetLeadTime" />'表單,工程,難易度,QC-L/T,現在時間,
        預定開始日時, 預定完成日時, Depo<br />
        <asp:Button ID="Button3" runat="server" Text="GetReadLeadTime" />預定開始日, Depo<br />
        <asp:Button ID="Button4" runat="server" Text="WriteWaitcheck" />表單,單號,預定開始日<br /><asp:Button ID="Button5" runat="server" Text="WriteWaitSend" />送件者, 收件者, 申請者, 表單,
        單號, 工程, 訊息類別<br />
        <asp:Button ID="Button6" runat="server" Text="CheckNO" />表單,單號,工程,No<br />
        <asp:Button ID="Button7" runat="server" Text="GetCalendarGroup" />申請者<br />
        <asp:Button ID="Button8" runat="server" Text="BatchFlowControl" /><br />
        <asp:Button ID="Button9" runat="server" Text="GetDevelopTime" /><br />
        <asp:Button ID="Button10" runat="server" Text="GetActualStartTime" /><br />
        <asp:Button ID="Button11" runat="server" Text="GetLastTimeCustom" /><br />
        <asp:Button ID="Button12" runat="server" Text="ArrangFlowControl" /><br />

        <asp:Button ID="Button14" runat="server" Text="AGENTAPPROV" /><br />
        <asp:Button ID="Button13" runat="server" Text="AGENTAPPROV" /><br />

        <asp:Button ID="BRecoveryFile" runat="server" Text="RecoveryFile" /><br />

        </div>
    </form>
</body>
</html>