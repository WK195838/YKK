


<%@ Page Language="VB" AutoEventWireup="false" CodeFile="GentaniSheet_04.aspx.vb" Inherits="GentaniSheet_04" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>原單位</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
                    <asp:Image ID="Image3" runat="server" ImageUrl="~/Images/DevelopmentGentani_06.jpg" Style="z-index: 99;
                    left: 8px; position: absolute; top: 8px" />
        <asp:Label ID="D4VMCLDLUNIT" runat="server" Style="z-index: 126; left: 760px; position: absolute;
            top: 496px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TCLDRITEM" runat="server" Style="z-index: 126; left: 632px; position: absolute;
            top: 496px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TCLDLITEM" runat="server" Height="16px" Style="z-index: 126; left: 632px;
            position: absolute; top: 472px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4VMCLDRUNIT" runat="server" Style="z-index: 126; left: 760px; position: absolute;
            top: 472px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CSUNIT" runat="server" Style="z-index: 145; left: 560px; position: absolute;
            top: 904px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O2SUM" runat="server" Style="z-index: 145; left: 800px; position: absolute;
            top: 904px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4COUNIT" runat="server" Style="z-index: 145; left: 632px; position: absolute;
            top: 904px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O1SUM" runat="server" Style="z-index: 145; left: 708px; position: absolute;
            top: 904px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CLSUM" runat="server" Style="z-index: 145; left: 488px; position: absolute;
            top: 904px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O1THRLOUN" runat="server" Style="z-index: 145; left: 708px; position: absolute;
            top: 885px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4COTHRUPUN" runat="server" Style="z-index: 205; left: 632px; position: absolute;
            top: 816px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4COTHLUPUN" runat="server" Style="z-index: 185; left: 632px; position: absolute;
            top: 800px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O1THLUPUN" runat="server" Style="z-index: 185; left: 708px; position: absolute;
            top: 800px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O2THLUPUN" runat="server" Style="z-index: 185; left: 800px; position: absolute;
            top: 800px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CSTHLUPUN" runat="server" Style="z-index: 185; left: 560px; position: absolute;
            top: 800px" Font-Size="Smaller"></asp:Label>
        &nbsp;
        <asp:Label ID="D4XMLCOL" runat="server" Style="z-index: 107; left: 186px; position: absolute;
            top: 195px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4XMRLEN" runat="server" Style="z-index: 234; left: 760px; position: absolute;
            top: 195px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4XMRITEM" runat="server" Style="z-index: 235; left: 672px; position: absolute;
            top: 195px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4XMRCOLNO" runat="server" Style="z-index: 247; left: 592px; position: absolute;
            top: 195px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4XMRCOL" runat="server" Style="z-index: 258; left: 520px; position: absolute;
            top: 195px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4XMLLEN" runat="server" Style="z-index: 269; left: 416px; position: absolute;
            top: 195px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4XMLITEM" runat="server" Style="z-index: 282; left: 336px; position: absolute;
            top: 195px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4XMLCOLNO" runat="server" Style="z-index: 293; left: 264px; position: absolute;
            top: 195px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TALCOL" runat="server" Style="z-index: 106; left: 186px; position: absolute;
            top: 176px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TALCOLNO" runat="server" Style="z-index: 108; left: 264px; position: absolute;
            top: 176px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TALITEM" runat="server" Style="z-index: 109; left: 336px; position: absolute;
            top: 176px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TALLEN" runat="server" Style="z-index: 110; left: 416px; position: absolute;
            top: 176px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TARCOL" runat="server" Style="z-index: 111; left: 520px; position: absolute;
            top: 176px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TARCOLNO" runat="server" Style="z-index: 112; left: 592px; position: absolute;
            top: 176px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TARITEM" runat="server" Style="z-index: 113; left: 672px; position: absolute;
            top: 176px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TARLEN" runat="server" Style="z-index: 115; left: 760px; position: absolute;
            top: 176px" Font-Size="Smaller"></asp:Label>
                    
     <asp:Label ID="D4NO" runat="server" Style="z-index: 100; left: 32px; position: absolute;
            top: 88px" Width="144px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4DEVNO" runat="server" Style="z-index: 101; left: 192px; position: absolute;
            top: 88px" Width="106px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4SIZENO" runat="server" Style="z-index: 102; left: 296px; position: absolute;
            top: 88px" Width="9px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4ITEM" runat="server" Style="z-index: 103; left: 376px; position: absolute;
            top: 88px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CODENO" runat="server" Style="z-index: 104; left: 448px; position: absolute;
            top: 88px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4LINENO1" runat="server" Style="z-index: 105; left: 528px; position: absolute;
            top: 88px" Font-Size="Smaller" Height="16px"></asp:Label>
        &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:Label ID="D4TCRLEN" runat="server" Style="z-index: 114; left: 760px; position: absolute;
            top: 344px" Font-Size="Smaller"></asp:Label>
        &nbsp;
        <asp:Label ID="D4YAMRSUM" runat="server" Style="z-index: 116; left: 760px; position: absolute;
            top: 360px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4PICOL" runat="server" Style="z-index: 117; left: 192px; position: absolute;
            top: 424px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CO1ITEM" runat="server" Style="z-index: 118; left: 336px; position: absolute;
            top: 448px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TSLITEM" runat="server" Style="z-index: 119; left: 632px; position: absolute;
            top: 424px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4VMCORUNIT" runat="server" Style="z-index: 120; left: 760px; position: absolute;
            top: 536px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TDRITEM" runat="server" Style="z-index: 121; left: 632px; position: absolute;
            top: 536px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4VMCOLUNIT" runat="server" Style="z-index: 122; left: 760px; position: absolute;
            top: 512px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TDLITEM" runat="server" Style="z-index: 123; left: 632px; position: absolute;
            top: 512px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4VMSETRUNIT" runat="server" Style="z-index: 124; left: 760px; position: absolute;
            top: 448px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4VMSETLUNIT" runat="server" Style="z-index: 125; left: 760px; position: absolute;
            top: 424px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TSRITEM" runat="server" Style="z-index: 126; left: 632px; position: absolute;
            top: 448px" Font-Size="Smaller"></asp:Label>
        &nbsp;&nbsp;
        <asp:Label ID="D4CO3LEN" runat="server" Style="z-index: 129; left: 416px; position: absolute;
            top: 496px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CO2LEN" runat="server" Style="z-index: 130; left: 416px; position: absolute;
            top: 472px" Font-Size="Smaller"></asp:Label>
        &nbsp;&nbsp;
        <asp:Label ID="D4CO3ITEM" runat="server" Style="z-index: 133; left: 336px; position: absolute;
            top: 496px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CO1LEN" runat="server" Style="z-index: 134; left: 416px; position: absolute;
            top: 448px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4PILEN" runat="server" Style="z-index: 135; left: 416px; position: absolute;
            top: 424px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CO2ITEM" runat="server" Style="z-index: 136; left: 336px; position: absolute;
            top: 472px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4PIITEM" runat="server" Style="z-index: 137; left: 336px; position: absolute;
            top: 424px" Font-Size="Smaller"></asp:Label>
        &nbsp;&nbsp;
        <asp:Label ID="D4CO3COLNO" runat="server" Style="z-index: 140; left: 264px; position: absolute;
            top: 496px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CO2COLNO" runat="server" Style="z-index: 141; left: 264px; position: absolute;
            top: 472px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CO1COLNO" runat="server" Style="z-index: 142; left: 264px; position: absolute;
            top: 448px" Font-Size="Smaller"></asp:Label>
        &nbsp;
        <asp:Label ID="D4ECOL1" runat="server" Style="z-index: 144; left: 192px; position: absolute;
            top: 712px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THRLOITEM" runat="server" Style="z-index: 145; left: 416px; position: absolute;
            top: 885px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THRLOCOLNO" runat="server" Style="z-index: 146; left: 336px; position: absolute;
            top: 885px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THLLOITEM" runat="server" Style="z-index: 147; left: 416px; position: absolute;
            top: 862px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THLLOCOLNO" runat="server" Style="z-index: 148; left: 336px; position: absolute;
            top: 862px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THLOITEM" runat="server" Style="z-index: 149; left: 416px; position: absolute;
            top: 840px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THLOCOLNO" runat="server" Style="z-index: 150; left: 336px; position: absolute;
            top: 840px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THRUPITEM" runat="server" Style="z-index: 151; left: 416px; position: absolute;
            top: 816px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THRUPCOLNO" runat="server" Style="z-index: 152; left: 336px; position: absolute;
            top: 816px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THLUPITEM" runat="server" Style="z-index: 153; left: 416px; position: absolute;
            top: 800px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THLUPCOLNO" runat="server" Style="z-index: 154; left: 336px; position: absolute;
            top: 800px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THUPITEM" runat="server" Style="z-index: 155; left: 416px; position: absolute;
            top: 776px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THUPCOLNO" runat="server" Style="z-index: 156; left: 336px; position: absolute;
            top: 776px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THRLOCOL" runat="server" Style="z-index: 157; left: 192px; position: absolute;
            top: 885px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THLLOCOL" runat="server" Style="z-index: 158; left: 192px; position: absolute;
            top: 840px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THLOCOL" runat="server" Style="z-index: 159; left: 192px; position: absolute;
            top: 862px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THRUPCOL" runat="server" Style="z-index: 160; left: 192px; position: absolute;
            top: 816px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THLUPCOL" runat="server" Style="z-index: 161; left: 192px; position: absolute;
            top: 800px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4THUPCOL" runat="server" Style="z-index: 162; left: 192px; position: absolute;
            top: 776px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CCOL" runat="server" Style="z-index: 163; left: 192px; position: absolute;
            top: 755px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CITEM" runat="server" Style="z-index: 164; left: 336px; position: absolute;
            top: 755px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4EITEM2" runat="server" Style="z-index: 165; left: 336px; position: absolute;
            top: 736px" Visible="False" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4EITEM1" runat="server" Style="z-index: 166; left: 336px; position: absolute;
            top: 712px" Font-Size="Smaller"></asp:Label>
        &nbsp;
        <asp:Label ID="D4CNITEM" runat="server" Style="z-index: 168; left: 488px; position: absolute;
            top: 608px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CLTAUN" runat="server" Style="z-index: 169; left: 488px; position: absolute;
            top: 688px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4COTAUN" runat="server" Style="z-index: 170; left: 632px; position: absolute;
            top: 688px" Font-Size="Smaller"></asp:Label>
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:Label ID="D4CLTHLUPUN" runat="server" Style="z-index: 185; left: 488px; position: absolute;
            top: 800px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O2CUN" runat="server" Style="z-index: 186; left: 800px; position: absolute;
            top: 752px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O2MONOUN" runat="server" Style="z-index: 187; left: 800px; position: absolute;
            top: 736px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O1CUN" runat="server" Style="z-index: 188; left: 708px; position: absolute;
            top: 755px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O1MONOUN" runat="server" Style="z-index: 189; left: 708px; position: absolute;
            top: 736px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4COCUN" runat="server" Style="z-index: 190; left: 632px; position: absolute;
            top: 755px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O2EUN" runat="server" Style="z-index: 191; left: 800px; position: absolute;
            top: 712px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O1EUN" runat="server" Style="z-index: 192; left: 708px; position: absolute;
            top: 712px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4COEUN" runat="server" Style="z-index: 193; left: 632px; position: absolute;
            top: 712px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O2TAUN" runat="server" Style="z-index: 194; left: 800px; position: absolute;
            top: 688px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O1TAUN" runat="server" Style="z-index: 196; left: 708px; position: absolute;
            top: 688px" Font-Size="Smaller"></asp:Label>
        &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:Label ID="D4CLTHRUPUN" runat="server" Style="z-index: 205; left: 560px; position: absolute;
            top: 816px" Font-Size="Smaller"></asp:Label>
        &nbsp;
        <asp:Label ID="D4CSCUN" runat="server" Style="z-index: 207; left: 560px; position: absolute;
            top: 755px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CSEUN" runat="server" Style="z-index: 208; left: 560px; position: absolute;
            top: 712px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CLCUN" runat="server" Style="z-index: 209; left: 488px; position: absolute;
            top: 755px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CLMONOUN" runat="server" Style="z-index: 210; left: 488px; position: absolute;
            top: 736px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CLEUN" runat="server" Style="z-index: 211; left: 488px; position: absolute;
            top: 712px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CSTAUN" runat="server" Style="z-index: 212; left: 560px; position: absolute;
            top: 688px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CSITEM" runat="server" Style="z-index: 213; left: 560px; position: absolute;
            top: 608px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CDITEM" runat="server" Style="z-index: 214; left: 632px; position: absolute;
            top: 608px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O1ITEM" runat="server" Style="z-index: 215; left: 708px; position: absolute;
            top: 608px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O2ITEM" runat="server" Style="z-index: 216; left: 800px; position: absolute;
            top: 608px" Font-Size="Smaller"></asp:Label>
        &nbsp;
        <asp:Label ID="D4ECOL2" runat="server" Style="z-index: 219; left: 192px; position: absolute;
            top: 736px" Visible="False" Font-Size="Smaller"></asp:Label>
        &nbsp;
        <asp:Label ID="D4CO3COL" runat="server" Style="z-index: 221; left: 192px; position: absolute;
            top: 496px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CO2COL" runat="server" Style="z-index: 222; left: 192px; position: absolute;
            top: 472px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CO1COL" runat="server" Style="z-index: 223; left: 192px; position: absolute;
            top: 448px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4PICOLNO" runat="server" Style="z-index: 224; left: 264px; position: absolute;
            top: 424px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4LYRLEN" runat="server" Style="z-index: 225; left: 760px; position: absolute;
            top: 320px" Font-Size="Smaller"></asp:Label>
        &nbsp; &nbsp;
        <asp:Label ID="D4EMRLEN" runat="server" Style="z-index: 229; left: 760px; position: absolute;
            top: 304px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4DMRLEN" runat="server" Style="z-index: 230; left: 760px; position: absolute;
            top: 280px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CMRLEN" runat="server" Style="z-index: 231; left: 760px; position: absolute;
            top: 264px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4BMRLEN" runat="server" Style="z-index: 232; left: 760px; position: absolute;
            top: 240px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4AMRLEN" runat="server" Style="z-index: 233; left: 760px; position: absolute;
            top: 216px" Font-Size="Smaller"></asp:Label>
        &nbsp;&nbsp;
        <asp:Label ID="D4AMRITEM" runat="server" Style="z-index: 236; left: 672px; position: absolute;
            top: 216px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4BMRITEM" runat="server" Style="z-index: 237; left: 672px; position: absolute;
            top: 240px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CMRITEM" runat="server" Style="z-index: 238; left: 672px; position: absolute;
            top: 264px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4DMRITEM" runat="server" Style="z-index: 239; left: 672px; position: absolute;
            top: 280px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4EMRITEM" runat="server" Style="z-index: 240; left: 672px; position: absolute;
            top: 304px" Font-Size="Smaller"></asp:Label>
        &nbsp; &nbsp;
        <asp:Label ID="D4LYRITEM" runat="server" Style="z-index: 244; left: 672px; position: absolute;
            top: 320px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TCRITEM" runat="server" Style="z-index: 245; left: 672px; position: absolute;
            top: 344px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TNRITEM1" runat="server" Style="z-index: 246; left: 624px; position: absolute;
            top: 360px" Font-Size="Smaller"></asp:Label>
        &nbsp;
        <asp:Label ID="D4AMRCOLNO" runat="server" Style="z-index: 248; left: 592px; position: absolute;
            top: 216px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4BMRCOLNO" runat="server" Style="z-index: 249; left: 592px; position: absolute;
            top: 240px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CMRCOLNO" runat="server" Style="z-index: 250; left: 592px; position: absolute;
            top: 260px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4DMRCOLNO" runat="server" Style="z-index: 251; left: 592px; position: absolute;
            top: 280px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4EMRCOLNO" runat="server" Style="z-index: 252; left: 592px; position: absolute;
            top: 304px" Font-Size="Smaller"></asp:Label>
        &nbsp; &nbsp;
        <asp:Label ID="D4LYRCOLNO" runat="server" Style="z-index: 256; left: 592px; position: absolute;
            top: 320px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TCRCOLNO" runat="server" Style="z-index: 257; left: 592px; position: absolute;
            top: 344px" Font-Size="Smaller"></asp:Label>
        &nbsp;
        <asp:Label ID="D4AMRCOL" runat="server" Style="z-index: 259; left: 520px; position: absolute;
            top: 216px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4BMRCOL" runat="server" Style="z-index: 260; left: 520px; position: absolute;
            top: 240px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CMRCOL" runat="server" Style="z-index: 261; left: 520px; position: absolute;
            top: 260px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4DMRCOL" runat="server" Style="z-index: 262; left: 520px; position: absolute;
            top: 280px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4EMRCOL" runat="server" Style="z-index: 263; left: 520px; position: absolute;
            top: 304px" Font-Size="Smaller"></asp:Label>
        &nbsp; &nbsp;
        <asp:Label ID="D4LYRCOL" runat="server" Style="z-index: 267; left: 520px; position: absolute;
            top: 320px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TCRCOL" runat="server" Style="z-index: 268; left: 520px; position: absolute;
            top: 344px" Font-Size="Smaller"></asp:Label>
        &nbsp;
        <asp:Label ID="D4AMLLEN" runat="server" Style="z-index: 270; left: 416px; position: absolute;
            top: 216px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4BMLLEN" runat="server" Style="z-index: 271; left: 416px; position: absolute;
            top: 240px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CMLLEN" runat="server" Style="z-index: 272; left: 416px; position: absolute;
            top: 260px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4DMLLEN" runat="server" Style="z-index: 273; left: 416px; position: absolute;
            top: 280px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4EMLLEN" runat="server" Style="z-index: 274; left: 416px; position: absolute;
            top: 304px" Font-Size="Smaller"></asp:Label>
        &nbsp; &nbsp;
        <asp:Label ID="D4LYLLEN" runat="server" Style="z-index: 278; left: 416px; position: absolute;
            top: 320px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TCLLEN" runat="server" Style="z-index: 279; left: 416px; position: absolute;
            top: 344px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4YAMLSUM" runat="server" Style="z-index: 280; left: 416px; position: absolute;
            top: 360px" Height="1px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TNLITEM1" runat="server" Style="z-index: 281; left: 288px; position: absolute;
            top: 360px" Height="1px" Font-Size="Smaller"></asp:Label>
        &nbsp;
        <asp:Label ID="D4AMLITEM" runat="server" Style="z-index: 283; left: 336px; position: absolute;
            top: 216px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4BMLITEM" runat="server" Style="z-index: 284; left: 336px; position: absolute;
            top: 240px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CMLITEM" runat="server" Style="z-index: 285; left: 336px; position: absolute;
            top: 260px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4DMLITEM" runat="server" Style="z-index: 286; left: 336px; position: absolute;
            top: 280px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4EMLITEM" runat="server" Style="z-index: 287; left: 336px; position: absolute;
            top: 304px" Font-Size="Smaller"></asp:Label>
        &nbsp; &nbsp;
        <asp:Label ID="D4LYLITEM" runat="server" Style="z-index: 291; left: 336px; position: absolute;
            top: 320px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TCLITEM" runat="server" Style="z-index: 292; left: 336px; position: absolute;
            top: 344px" Font-Size="Smaller"></asp:Label>
        &nbsp;
        <asp:Label ID="D4AMLCOLNO" runat="server" Style="z-index: 294; left: 264px; position: absolute;
            top: 216px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4BMLCOLNO" runat="server" Style="z-index: 295; left: 264px; position: absolute;
            top: 240px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CMLCOLNO" runat="server" Style="z-index: 296; left: 264px; position: absolute;
            top: 260px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4DMLCOLNO" runat="server" Style="z-index: 297; left: 264px; position: absolute;
            top: 280px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4EMLCOLNO" runat="server" Style="z-index: 298; left: 264px; position: absolute;
            top: 304px" Font-Size="Smaller"></asp:Label>
        &nbsp; &nbsp;
        <asp:Label ID="D4LYLCOLNO" runat="server" Style="z-index: 302; left: 264px; position: absolute;
            top: 320px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TCLCOLNO" runat="server" Style="z-index: 303; left: 264px; position: absolute;
            top: 344px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4AMLCOL" runat="server" Style="z-index: 304; left: 184px; position: absolute;
            top: 216px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4BMLCOL" runat="server" Style="z-index: 305; left: 184px; position: absolute;
            top: 240px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CMLCOL" runat="server" Style="z-index: 306; left: 184px; position: absolute;
            top: 260px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4DMLCOL" runat="server" Style="z-index: 307; left: 184px; position: absolute;
            top: 280px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4EMLCOL" runat="server" Style="z-index: 308; left: 184px; position: absolute;
            top: 304px" Font-Size="Smaller"></asp:Label>
        &nbsp; &nbsp;
        <asp:Label ID="D4LYLCOL" runat="server" Style="z-index: 312; left: 184px; position: absolute;
            top: 320px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4TCLCOL" runat="server" Style="z-index: 313; left: 184px; position: absolute;
            top: 344px" Font-Size="Smaller"></asp:Label>
        &nbsp;&nbsp;&nbsp;
        <asp:Label ID="D4CSTHRUPUN" runat="server" Style="z-index: 205; left: 488px; position: absolute;
            top: 816px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O1THRUPUN" runat="server" Style="z-index: 205; left: 708px; position: absolute;
            top: 816px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O2THRUPUN" runat="server" Style="z-index: 205; left: 800px; position: absolute;
            top: 816px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CLTHLLOUN" runat="server" Style="z-index: 205; left: 488px; position: absolute;
            top: 862px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CSTHLLOUN" runat="server" Style="z-index: 205; left: 560px; position: absolute;
            top: 862px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O1THLLOUN" runat="server" Style="z-index: 205; left: 708px; position: absolute;
            top: 862px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4COTHLLOUN" runat="server" Style="z-index: 205; left: 632px; position: absolute;
            top: 862px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O2THLLOUN" runat="server" Style="z-index: 205; left: 800px; position: absolute;
            top: 864px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4O2THRLOUN" runat="server" Style="z-index: 145; left: 800px; position: absolute;
            top: 885px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4COTHRLOUN" runat="server" Style="z-index: 145; left: 632px; position: absolute;
            top: 885px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CLTHRLOUN" runat="server" Style="z-index: 145; left: 488px; position: absolute;
            top: 885px" Font-Size="Smaller"></asp:Label>
        <asp:Label ID="D4CSTHRLOUN" runat="server" Style="z-index: 145; left: 560px; position: absolute;
            top: 885px" Font-Size="Smaller"></asp:Label>
                    
           <asp:TextBox ID="D4OTHER2" runat="server" Height="38px" Style="z-index: 314; left: 792px;
            position: absolute; top: 624px;font-size:10px;background:transparent" TextMode="MultiLine" Width="72px" BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="D4OTHER1" runat="server" Height="38px" Style="z-index: 318; left: 704px;
            position: absolute; top: 624px;font-size:10px;background:transparent" TextMode="MultiLine" Width="72px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
          
    </div>
    </form>
</body>
</html>
