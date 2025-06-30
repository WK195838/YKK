<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ViewProcHistory.aspx.vb" Inherits="ViewProcHistory" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
    <script language="javascript">
    function GoBack()
    {
        history.go(-1);
    }
    
    function OpenFile(doc)
    {
        window.open(doc);
        return false;
    
    }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="TextBox1" runat="server" ReadOnly="True" SkinID="txt_Yellow" Columns="50"></asp:TextBox>
        <input id="Button2" type="button" value="回上頁" onclick="GoBack();" />
        
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            PageSize="50" SkinID="GV" Width="100%">
            <Columns>
                <asp:BoundField DataField="ProcPeriod" HeaderText="期間" />
                <asp:BoundField DataField="PDF_ID" HeaderText="PDF ID" />
                <asp:TemplateField HeaderText="作成時間">
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# ConvertDate(Eval("PDF_CreateTime")) %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="PDF_PDFFile" HeaderText="檔案名" />
                <asp:TemplateField HeaderText="附件">
                    <ItemTemplate>
                        <asp:ImageButton ID="ImageButton1" runat="server" CommandArgument='<%# Eval("PDF_PDFFile") & "|" & Eval("Mail_ID") %>'
                            ImageUrl="~/images/pdf.jpg" CommandName="attchfile" Visible="False" />
                        <asp:Image ID="Image1" runat="server" ImageUrl="~/images/pdf.jpg" style="cursor:hand;"/>
                    </ItemTemplate>
                </asp:TemplateField>
                
                <asp:TemplateField HeaderText="傳送時間">
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# ConvertDate(Eval("Mail_CreateTime")) %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="Mail_SenderName" HeaderText="傳送者" />
                <asp:BoundField DataField="Mail_ToName" HeaderText="收件者" />
                <asp:BoundField DataField="Mail_Cc" HeaderText="CC者" />
                <asp:BoundField DataField="Mail_Subject" HeaderText="主旨" />
            </Columns>
        </asp:GridView>
    
    </div>
    </form>
</body>
</html>
