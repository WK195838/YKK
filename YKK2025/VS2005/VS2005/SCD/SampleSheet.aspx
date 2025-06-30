<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SampleSheet.aspx.vb" Inherits="SampleSheet" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
                    <asp:Image ID="Image4" runat="server" ImageUrl="~/Images/DevelopmentSample_01.png" Style="z-index: 99;
                    left: 5px; position: absolute; top: 2px" />

        <asp:FileUpload ID="D3QCFILE2" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 525px; position: absolute; top: 722px; background-color: #c0ffff"
            Width="267px" />
        <asp:FileUpload ID="D3QCFILE4" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 525px; position: absolute; top: 752px; background-color: #c0ffff"
            Width="267px" />
        <asp:FileUpload ID="D3SAMPLEFILE" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 20px; position: absolute; top: 198px; background-color: #c0ffff"
            Width="771px" />
        <asp:FileUpload ID="D3QCFILE1" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 146px; position: absolute; top: 722px; background-color: #c0ffff"
            Width="248px" />
        <asp:FileUpload ID="D3QCFILE3" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 146px; position: absolute; top: 752px; background-color: #c0ffff"
            Width="248px" />
        <asp:FileUpload ID="D3QCFILE5" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 146px; position: absolute; top: 780px; background-color: #c0ffff"
            Width="248px" />

				<asp:hyperlink id="LQCFILE5" style="Z-INDEX: 178; POSITION: absolute; TOP: 781px; LEFT: 148px"
					runat="server" Font-Size="12pt" Height="8px" NavigateUrl="" Target="_blank">其它</asp:hyperlink><asp:label ID="D31Other" style="Z-INDEX: 173; POSITION: absolute; TOP: 966px; LEFT: 170px" runat="server"
					Font-Size="12px" Width="116px"> </asp:label><asp:label ID="D32Other" style="Z-INDEX: 172; POSITION: absolute; TOP: 990px; LEFT: 170px" runat="server"
					Font-Size="12px" Width="116px"> </asp:label><asp:button id="BCREATE" style="Z-INDEX: 170; POSITION: absolute; TOP: 97px; LEFT: 543px" runat="server"
					Width="20px" Height="20px" CausesValidation="False" Text="....."></asp:button><asp:dropdownlist ID="D3WF7Name" style="Z-INDEX: 169; POSITION: absolute; TOP: 1080px; LEFT: 674px"
					runat="server" Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="LightGray">
				</asp:dropdownlist><asp:dropdownlist ID="D3WF6Name" style="Z-INDEX: 168; POSITION: absolute; TOP: 1080px; LEFT: 569px"
					runat="server" Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="LightGray">
				</asp:dropdownlist><asp:dropdownlist ID="D3WF5Name" style="Z-INDEX: 167; POSITION: absolute; TOP: 1080px; LEFT: 464px"
					runat="server" Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="LightGray">
				</asp:dropdownlist><asp:dropdownlist ID="D3WF4Name" style="Z-INDEX: 166; POSITION: absolute; TOP: 1080px; LEFT: 359px"
					runat="server" Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="LightGray">
				</asp:dropdownlist><asp:dropdownlist ID="D3WF3Name" style="Z-INDEX: 165; POSITION: absolute; TOP: 1080px; LEFT: 254px"
					runat="server" Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="LightGray">
				</asp:dropdownlist><asp:hyperlink id="LQCFILE1" style="Z-INDEX: 160; POSITION: absolute; TOP: 724px; LEFT: 147px"
					runat="server" Font-Size="12pt" Height="8px" NavigateUrl="" Target="_blank">吋法檔案</asp:hyperlink><asp:dropdownlist ID="D3WF7" style="Z-INDEX: 159; POSITION: absolute; TOP: 1107px; LEFT: 674px" runat="server"
					Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="LightGray">
				</asp:dropdownlist><asp:dropdownlist ID="D3WF6" style="Z-INDEX: 158; POSITION: absolute; TOP: 1107px; LEFT: 569px" runat="server"
					Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="LightGray">
				</asp:dropdownlist><asp:dropdownlist ID="D3WF5" style="Z-INDEX: 157; POSITION: absolute; TOP: 1107px; LEFT: 464px" runat="server"
					Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="LightGray">
				</asp:dropdownlist><asp:dropdownlist ID="D3WF4" style="Z-INDEX: 156; POSITION: absolute; TOP: 1107px; LEFT: 359px" runat="server"
					Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="LightGray">
				</asp:dropdownlist><asp:dropdownlist ID="D3WF3" style="Z-INDEX: 155; POSITION: absolute; TOP: 1107px; LEFT: 254px" runat="server"
					Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="LightGray">
				</asp:dropdownlist><asp:dropdownlist ID="D3WF2" style="Z-INDEX: 154; POSITION: absolute; TOP: 1107px; LEFT: 149px" runat="server"
					Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="LightGray">
				</asp:dropdownlist><asp:dropdownlist ID="D3WF1" style="Z-INDEX: 109; POSITION: absolute; TOP: 1107px; LEFT: 44px" runat="server"
					Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="LightGray">
				</asp:dropdownlist><asp:textbox ID="D3CITEM" style="Z-INDEX: 147; POSITION: absolute; TOP: 1011px; LEFT: 295px" runat="server"
					Width="324px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3O2ITEM" style="Z-INDEX: 145; POSITION: absolute; TOP: 987px; LEFT: 295px" runat="server"
					Width="324px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3O1ITEM" style="Z-INDEX: 146; POSITION: absolute; TOP: 963px; LEFT: 295px" runat="server"
					Width="324px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3CDITEM" style="Z-INDEX: 144; POSITION: absolute; TOP: 939px; LEFT: 295px" runat="server"
					Width="324px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3CSITEM" style="Z-INDEX: 143; POSITION: absolute; TOP: 915px; LEFT: 295px" runat="server"
					Width="324px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TDRITEM" style="Z-INDEX: 142; POSITION: absolute; TOP: 867px; LEFT: 525px"
					runat="server" Width="94px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TSRITEM" style="Z-INDEX: 141; POSITION: absolute; TOP: 843px; LEFT: 525px"
					runat="server" Width="94px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TNRITEM" style="Z-INDEX: 140; POSITION: absolute; TOP: 819px; LEFT: 525px"
					runat="server" Width="94px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3CNITEM" style="Z-INDEX: 139; POSITION: absolute; TOP: 891px; LEFT: 295px" runat="server"
					Width="324px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TDLITEM" style="Z-INDEX: 138; POSITION: absolute; TOP: 867px; LEFT: 295px"
					runat="server" Width="92px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TSLITEM" style="Z-INDEX: 137; POSITION: absolute; TOP: 843px; LEFT: 295px"
					runat="server" Width="92px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TNLITEM" style="Z-INDEX: 136; POSITION: absolute; TOP: 819px; LEFT: 295px"
					runat="server" Width="92px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox>
				<asp:hyperlink id="LQCFILE2" style="Z-INDEX: 161; POSITION: absolute; TOP: 724px; LEFT: 524px"
					runat="server" Font-Size="12pt" Height="8px" Target="_blank" Width="121px">強度檔案</asp:hyperlink><asp:hyperlink id="LQCFILE3" style="Z-INDEX: 163; POSITION: absolute; TOP: 754px; LEFT: 147px"
					runat="server" Font-Size="12pt" Height="8px" NavigateUrl="" Target="_blank">往覆測試檔案</asp:hyperlink><asp:hyperlink id="LQCFILE4" style="Z-INDEX: 162; POSITION: absolute; TOP: 751px; LEFT: 524px"
					runat="server" Font-Size="12pt" Height="8px" Target="_blank" Width="188px">仕樣書.組織圖檔案</asp:hyperlink>
        &nbsp;&nbsp;
				<asp:textbox ID="D3OTHER" style="Z-INDEX: 130; POSITION: absolute; TOP: 634px; LEFT: 147px" runat="server"
					Width="640px" Height="76px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" MaxLength="240"
					TextMode="MultiLine" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3THCOL" style="Z-INDEX: 129; POSITION: absolute; TOP: 578px; LEFT: 147px" runat="server"
					Width="640px" Height="45px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" MaxLength="240" TextMode="MultiLine" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3CCOL" style="Z-INDEX: 128; POSITION: absolute; TOP: 547px; LEFT: 147px" runat="server"
					Width="640px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3ECOL" style="Z-INDEX: 127; POSITION: absolute; TOP: 518px; LEFT: 147px" runat="server"
					Width="640px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TACOL" style="Z-INDEX: 126; POSITION: absolute; TOP: 402px; LEFT: 147px" runat="server"
					Width="640px" Height="48px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" MaxLength="240" TextMode="MultiLine" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3DEVPRD" style="Z-INDEX: 125; POSITION: absolute; TOP: 372px; LEFT: 609px" runat="server"
					Width="178px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3DEVNO" style="Z-INDEX: 124; POSITION: absolute; TOP: 372px; LEFT: 378px" runat="server"
					Width="115px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TAWIDTH" style="Z-INDEX: 123; POSITION: absolute; TOP: 372px; LEFT: 147px"
					runat="server" Width="85px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox>
				<asp:image id="LSAMPLEFILE" style="Z-INDEX: 121; POSITION: absolute; TOP: 220px; LEFT: 21px"
					runat="server" Width="770px" Height="146px" BorderStyle="Groove" ImageUrl="F:\DMF04006-DS2W.jpg"></asp:image><asp:textbox ID="D3CODENO" style="Z-INDEX: 120; POSITION: absolute; TOP: 144px; LEFT: 566px" runat="server"
					Width="220px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3ITEM" style="Z-INDEX: 119; POSITION: absolute; TOP: 144px; LEFT: 356px" runat="server"
					Width="200px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3SIZENO" style="Z-INDEX: 118; POSITION: absolute; TOP: 144px; LEFT: 146px" runat="server"
					Width="200px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TALINE" style="Z-INDEX: 103; POSITION: absolute; TOP: 460px; LEFT: 147px" runat="server"
					Width="640px" Height="48px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" MaxLength="240" TextMode="MultiLine" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3DATE" style="Z-INDEX: 164; POSITION: absolute; TOP: 91px; LEFT: 671px" runat="server"
					Width="115px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3APPBUYER" style="Z-INDEX: 104; POSITION: absolute; TOP: 91px; LEFT: 147px"
					runat="server" Width="390px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox>&nbsp;

        <asp:HyperLink ID="LOldSSheet" runat="server" Font-Size="12pt" Height="8px"
            Style="z-index: 261; left: 128px; position: absolute; top: 30px" Target="_blank"
            Width="104px">舊-開發見本</asp:HyperLink>

        <asp:HyperLink ID="LOldSImage" runat="server" Font-Size="12pt" Height="8px"
            Style="z-index: 261; left: 18px; position: absolute; top: 30px" Target="_blank"
            Width="104px">舊-樣品圖</asp:HyperLink>

    
    </div>
    </form>
</body>
</html>
