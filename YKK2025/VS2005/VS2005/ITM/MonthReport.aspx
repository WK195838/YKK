<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MonthReport.aspx.vb" Inherits="MonthReport" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>MonthReport</title>
    <script language="javascript" src="ForProject.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" PageSize="20" Style="z-index: 114; left: 8px; position: absolute;
            top: 83px" Width="1100px" BorderColor="Black" BorderStyle="Groove">
            <Columns>
                <asp:BoundField DataField="ReportDate" HeaderText="月/日" >
                    <ItemStyle HorizontalAlign="Center"  Font-Size="10pt" BorderColor="#804040" BorderStyle="Groove"/>
                    <HeaderStyle  BorderColor="Black" BorderStyle="Groove" />
                </asp:BoundField>

                <asp:BoundField DataField="WeekDesc" HeaderText="星期" >
                    <ItemStyle HorizontalAlign="Center"  Font-Size="10pt" BorderColor="#804040" BorderStyle="Groove"/>
                    <HeaderStyle  BorderColor="Black" BorderStyle="Groove"/>
                </asp:BoundField>

                <asp:TemplateField HeaderText="GROUP">
                    <ItemTemplate>
                        <a href='MonthReport.aspx?pUserID=<%#Eval("pUserID")%>&pReportYear=<%#Eval("pNowYear")%>&pReportMonth=<%#Eval("pNowMonth")%>&pKeyField=Division&pKeyData=<%#Server.UrlEncode(Eval("Division"))%>'><%#Eval("Division")%></a>
                    </ItemTemplate>
                        <ItemStyle HorizontalAlign="Left"  Font-Size="10pt" BorderColor="#804040" BorderStyle="Groove"/>
                        <HeaderStyle  BorderColor="Black" BorderStyle="Groove"/>
                </asp:TemplateField>


                <asp:TemplateField HeaderText="TEAM">
                    <ItemTemplate>
                        <a href='MonthReport.aspx?pUserID=<%#Eval("pUserID")%>&pReportYear=<%#Eval("pNowYear")%>&pReportMonth=<%#Eval("pNowMonth")%>&pKeyField=Section&pKeyData=<%#Server.UrlEncode(Eval("Section"))%>'><%#Eval("Section")%></a>
                    </ItemTemplate>
                        <ItemStyle HorizontalAlign="Left"  Font-Size="10pt" BorderColor="#804040" BorderStyle="Groove"/>
                        <HeaderStyle  BorderColor="Black" BorderStyle="Groove"/>
                </asp:TemplateField>

                <asp:TemplateField HeaderText="成員">
                    <ItemTemplate>
                        <a href='MonthReport.aspx?pUserID=<%#Eval("pUserID")%>&pReportYear=<%#Eval("pNowYear")%>&pReportMonth=<%#Eval("pNowMonth")%>&pKeyField=Name&pKeyData=<%#Server.UrlEncode(Eval("Name"))%>'><%#Eval("Name")%></a>
                    </ItemTemplate>
                        <ItemStyle HorizontalAlign="Left"  Font-Size="10pt" BorderColor="#804040" BorderStyle="Groove"/>
                        <HeaderStyle BorderColor="Black" BorderStyle="Groove" />
                </asp:TemplateField>
                
                <asp:BoundField DataField="wType" HeaderText="作業" >
                    <ItemStyle HorizontalAlign="Left"  Font-Size="10pt" BorderColor="#804040" BorderStyle="Groove" />
                    <HeaderStyle  BorderColor="Black" BorderStyle="Groove" />
                </asp:BoundField>
                
                <asp:BoundField DataField="wContent" HeaderText="內容" HtmlEncode="false">
                    <ItemStyle HorizontalAlign="Left"  Font-Size="10pt" BorderColor="#804040" BorderStyle="Groove"/>
                    <HeaderStyle  BorderColor="Black" BorderStyle="Groove" />
                </asp:BoundField>
                
                <asp:BoundField DataField="wRemark" HeaderText="其它說明" HtmlEncode="false">
                    <ItemStyle HorizontalAlign="Left"  Font-Size="10pt" BorderColor="#804040" BorderStyle="Groove"/>
                    <HeaderStyle BorderColor="Black" BorderStyle="Groove" />
                </asp:BoundField>

                <asp:BoundField DataField="Vacation" HeaderText="@休假" >
                </asp:BoundField>

                <asp:BoundField DataField="Division" HeaderText="@GROUP" >
                </asp:BoundField>

                <asp:BoundField DataField="Section" HeaderText="@Section" >
                </asp:BoundField>

                <asp:BoundField DataField="Name" HeaderText="@Name" >
                </asp:BoundField>

                <asp:BoundField DataField="wType" HeaderText="@Type" >
                </asp:BoundField>
               
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <a href='NewReport.aspx?pID=<%#Eval("wID")%>&pUserID=<%#Eval("pUserID")%>&pFun=MOD' target="_blank" ><%#Eval("Maint")%></a>
                    </ItemTemplate>
                        <ItemStyle HorizontalAlign="Left"  Font-Size="10pt" BorderColor="#804040" BorderStyle="Groove"/>
                        <HeaderStyle  BorderColor="Black" BorderStyle="Groove"/>
                </asp:TemplateField>


            </Columns>
            <FooterStyle BackColor="Silver" BorderStyle="Groove" />
            <RowStyle BackColor="#FFFFC0" BorderStyle="Groove" ForeColor="Blue" />
            <PagerStyle BackColor="#FFE0C0" />
            <HeaderStyle BackColor="LightGray" BorderStyle="Groove" BorderColor="Maroon" Font-Bold="False" Font-Italic="False" Font-Strikeout="False" Font-Underline="False" ForeColor="Black" />
        </asp:GridView>
        <asp:Button ID="BUp" runat="server" Style="z-index: 114; left: 8px; position: absolute; vertical-align: medium;
            top: 55px" Text="上一月" Width="100px" />
         <asp:Button ID="BDown" runat="server" Style="z-index: 114; left: 1009px; position: absolute;
            top: 55px" Text="下一月" Width="100px" />
        <asp:Button ID="BThis" runat="server" Style="z-index: 114; left: 536px; position: absolute;
            top: 55px" Text="當月" Width="100px" />

        <asp:Button ID="BWeek" runat="server" Style="z-index: 114; left: 1006px; position: absolute; 
            top: 8px" Text="週表示" Width="100px" Height="25px"/>

        <asp:TextBox ID="DReportYear" runat="server" BackColor="White" BorderStyle="None"
            Font-Bold="True" Font-Overline="True" Font-Size="14pt" Font-Underline="True"
            ForeColor="Black" Height="20px" Style="z-index: 108; left: 475px; position: absolute;
            top: 54px" Width="50px">2008</asp:TextBox>
        &nbsp;&nbsp;

        <asp:HyperLink ID="DFunLink1" runat="server" Style="z-index: 114; left: 12px; position: absolute;
            top: 15px" Font-Bold="True" Font-Size="14pt" ForeColor="Maroon" Width="120px">全體行程</asp:HyperLink>
        <asp:HyperLink ID="DFunLink2" runat="server" Font-Bold="True" Font-Size="14pt" ForeColor="Maroon"
            Style="z-index: 114; left: 163px; position: absolute; top: 15px" Width="250px">北部營業部</asp:HyperLink>
        <asp:Label ID="DFunMark1" runat="server" Style="z-index: 114; left: 136px; position: absolute; top: 15px; text-align:center" Font-Size="14pt" ForeColor="Maroon">→</asp:Label>

        <asp:Label ID="DLabel1" runat="server" Style="z-index: 114; left: 714px; position: absolute; top: 55px; text-align:center" Font-Size="10pt" ForeColor="Maroon">SEARCH</asp:Label>
        <asp:TextBox ID="DSearch" runat="server" Style="z-index: 114; left: 769px; position: absolute; top: 55px; text-align:LEFT" Font-Size="10pt" ForeColor="Maroon" BackColor="#FFFF80" Height="16px" Width="224px"></asp:TextBox>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:TextBox ID="DUserID" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        </div>
    </form>
</body>
</html>
