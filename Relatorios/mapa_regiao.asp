<!--#include file="RelatoriosCla.asp"-->
<%

QtdeRioGSul=1
QtdeParana=2
QtdeSantaCatarina=3
QtdeSaoPaulo=4
QtdeBeloHorizonte=5
QtdeRiodeJaneiro=6
QtdeMatoGSul=7
QtdeEspiritoSanto=8
QtdeGoias=9
QtdeMatoGrosso=10
QtdeBahia=11
QtdeDistritoFeredal=12
QtdeTocantins=13
QtdeRondonia=14
QtdeAcre=15
QtdeAmazonas=16
QtdeRoraima=17
QtdePara=18
QtdeAmapa=19
QtdeMaranhao=20
QtdePiaui=21
QtdeCeara=22
QtdeRioGNorte=23
QtdeParaiba=24
QtdePernanbuco=25
QtdeAlagoas=26
QtdeSergipe=27

%>
<html>
<head>
<title>mapa do brasil</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<style type="text/css">
ul { /*a área do menu*/
position:relative /*contexto de posicionamento*/
list-style:none; /*retira bullets da lista*/
padding:0; /*zerando paddings*/
margin:0;  /*zerando margins*/
  
/*dimensões da área do menu*/
width:360px; 
height:253px;
  
/*Posicionando o crânio*/
background:#fff1f1 url('mapabr.gif') no-repeat 20px 10px;
border:1px solid #e4ddd7
}
/*As regras para a (âncora) criam o efeito visual
para os bullets sobre o crânio, 
são os pontos clicáveis*/
a {
position:absolute;
display:block;
text-decoration:none;
text-align:center;
color:#f00;
border:2px solid #f00;/*simula o 
quadradinho vermelho*/
}
/*As regras a seguir posicionam os 
bullets sobre o crânio*/
/*Rio grande do sul*/
a.menu1{
top:394px;
left:351px;
}
/*Santa catarina*/
a.menu2{
top:358px;
left:382px;
}
/*Parana*/
a.menu3{
top:316px;
left:367px;
}
/*Sao Paulo*/
a.menu4{
top:316px;
left:395px;
}
/*Belo Horizonte*/
a.menu5{
top:285px;
left:430px;
}
/*Rio de janeiro*/
a.menu6{
top:303px;
left:450px;
}
/*Mato Grosso do Sul*/
a.menu7{
top:305px;
left:338px;
}

/*Espirito Santo*/
a.menu8{
top:285px;
left:462px;
}

/* Goias */
a.menu9{
top:265px;
left:369px;
}

/* Mato Grosso */
a.menu10{
top:210px;
left:348px;
}

/* Bahia */
a.menu11{
top:222px;
left:458px;
}

/* Distrito Federal */
a.menu12{
top:240px;
left:402px;
}

/* Tocantins */
a.menu13{
top:180px;
left:396px;
}

/* Rondonia */
a.menu14{
top:208px;
left:270px;
}

/* Acre */
a.menu15{
top:185px;
left:196px;
}

/* Amazonas*/
a.menu16{
top:148px;
left:250px;
}

/* Roraima */
a.menu17{
top:75px;
left:275px;
}

/* Para */
a.menu18{
top:149px;
left:350px;
}

/* Amapa*/
a.menu19{
top:88px;
left:359px;
}

/* Maranhão */
a.menu20{
top:148px;
left:424px;
}

/* Piaui */
a.menu21{
top:148px;
left:454px;
}

/* Ceara */
a.menu22 {
top:150px;
left:475px;
}

/* Rio Grande do Norte */
a.menu23 {
top:146px;
left:503px;
}

/* Paraiba */
a.menu24 {
top:159px;
left:493px;
}

/* Pernanbuco */
a.menu25 {
top:171px;
left:492px;
}

/* Alagoas */
a.menu26 {
top:184px;
left:504px;
}

/* Sergipe */
a.menu27 {
top:192px;
left:497px;
}

a span {display:none} /*esconde caixa tooltip*/

/*mostra e estiliza a caixa tooltip*/
a.menu1:hover span, a.menu2:hover span,  
a.menu3:hover span, a.menu4:hover span,  
a.menu5:hover span, a.menu6:hover span, 
a.menu7:hover span, a.menu8:hover span,
a.menu9:hover span, a.menu10:hover span,
a.menu11:hover span, a.menu12:hover span,
a.menu13:hover span, a.menu14:hover span,
a.menu15:hover span, a.menu16:hover span,
a.menu17:hover span, a.menu18:hover span,
a.menu19:hover span, a.menu20:hover span,
a.menu21:hover span, a.menu22:hover span,
a.menu23:hover span, a.menu24:hover span,
a.menu25:hover span, a.menu26:hover span,
a.menu27:hover span
{ 

width:160px; 
display:block;
position:absolute;
font:11px arial, verdana, helvetica, sans-serif; 
text-align:center;
padding:5px; 
border:1px solid #f00;
background:#fff; 
color:#000;
text-decoration:none;
}
/*box model para browsers conformes*/
li>a.menu1:hover span, li>a.menu2:hover span,  
li>a.menu3:hover span, li>a.menu4:hover span,  
li>a.menu5:hover span,  li>a.menu6:hover span { 
width:148px;
}
/*posiciona as caixas tooltip*/
/*Rio grande do sul*/
a.menu1:hover  { 
border:none;
top:394px;
left:351px;
}
/*Santa catarina*/
a.menu2:hover { 
border:none;
top:358px;
left:382px;
}
/*Parana*/
a.menu3:hover {
border:none;
top:316px;
left:367px;
}
/*Sao Paulo*/
a.menu4:hover {
border:none;
top:316px;
left:395px;
}
/*Belo Horizonte*/
a.menu5:hover { 
border:none;
top:285px;
left:430px;
}
/*Rio de janeiro*/
a.menu6:hover  { 
border:none;
top:303px;
left:450px;
}
/*Mato Grosso do Sul*/
a.menu7:hover  { 
border:none;
top:305px;
left:338px;
}

/*Espirito Santo*/
a.menu8:hover  { 
border:none;
top:285px;
left:462px;
}

/* Goias */
a.menu9:hover  { 
border:none;
top:265px;
left:369px;
}

/* Mato Grosso */
a.menu10:hover  { 
border:none;
top:210px;
left:348px;
}

/* Bahia  */
a.menu11:hover  { 
border:none;
top:222px;
left:458px;
}

/* Distrito Federal */
a.menu12:hover  { 
border:none;
top:240px;
left:402px;
}

/* Tocantins */
a.menu13:hover  { 
border:none;
top:180px;
left:396px;
}

/* Rondonia */
a.menu14:hover  { 
border:none;
top:208px;
left:270px;
}

/* Acre */
a.menu15:hover  { 
border:none;
top:185px;
left:196px;
}

/* Amazonas */
a.menu16:hover  { 
border:none;
top:148px;
left:250px;
}

/* Roraima */
a.menu17:hover  { 
border:none;
top:75px;
left:275px;
}

/* Para */
a.menu18:hover  { 
border:none;
top:149px;
left:350px;
}

/* Amapa */
a.menu19:hover  { 
border:none;
top:88px;
left:359px;
}

/* Maranhão */
a.menu20:hover  { 
border:none;
top:148px;
left:424px;
}

/* Piaui */
a.menu21:hover  { 
border:none;
top:148px;
left:454px;
}

/* Ceara */
a.menu22:hover  { 
border:none;
top:150px;
left:475px;
}

/* Rio Grande do Norte */
a.menu23:hover  { 
border:none;
top:146px;
left:503px;
}

/* Paraiba */
a.menu24:hover  { 
border:none;
top:159px;
left:493px;
}

/*Pernanbuco*/
a.menu25:hover  { 
border:none;
top:171px;
left:492px;
}

/* Alagoas */
a.menu26:hover  { 
border:none;
top:184px;
left:504px;
}

/*Sergipe */
a.menu27:hover  { 
border:none;
top:192px;
left:497px;
}
/*estiliza o título do menu*/
.title h1 {
position:absolute;
right:8px;
top:-10px;
font:bold 14px  verdana, arial,helvetica, sans-serif; 
color:#9c816b;
background:none;
border:none;
}
</style>

</head>

<body bgcolor="#FFFFFF">
<Form name="FormRelat" method="Post" action="detalhe_acesso_endereco.asp" target="_self" >
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<table border="0" width="91%" aling="center" >
  <tr>
    <td width="74%" align="center"><map name="FPMap0">
  
  
    <a href="consrios.htm" class="menu1"><span><%=QtdeRioGSul%></span><area href="consrios.htm" shape="polygon" coords="192, 362, 202, 357, 207, 362, 210, 369, 215, 365, 217, 370, 235, 385, 232, 390, 236, 395, 244, 384, 244, 379, 252, 374, 266, 353, 263, 345, 251, 336, 230, 333, 226, 334"></a>
    <a href="conssc.htm" class="menu2"><span><%=QtdeParana%></span>  <area href="conssc.htm" shape="polygon" coords="227, 332, 243, 331, 255, 335, 262, 343, 266, 344, 266, 350, 273, 345, 279, 340, 279, 322, 272, 322, 264, 323, 256, 325, 251, 325, 245, 325, 238, 324, 230, 323"></a>
    <a href="conspara.htm" class="menu3"><span><%=QtdeSantaCatarina%></span> <area href="conspara.htm" shape="polygon" coords="223, 307, 222, 317, 230, 320, 254, 325, 266, 320, 277, 321, 278, 312, 271, 310, 269, 297, 262, 295, 255, 293, 241, 290"></a>
    <a href="conssaop.htm" class="menu4"><span><%=QtdeSaoPaulo%></span> <area href="conssaop.htm" shape="polygon" coords="281, 314, 276, 306, 271, 295, 264, 290, 252, 286, 243, 287, 249, 275, 256, 266, 265, 263, 270, 265, 273, 268, 288, 265, 292, 275, 296, 280, 299, 291, 315, 292, 312, 296, 308, 299, 300, 300"></a>
    <a href="consbelo.htm" class="menu5"><span><%=QtdeBeloHorizonte%></span><area href="consbelo.htm" shape="polygon" coords="353, 231, 343, 223, 327, 218, 322, 212, 308, 220, 298, 221, 294, 234, 291, 248, 282, 250, 264, 254, 256, 260, 287, 262, 294, 265, 301, 288, 316, 283, 334, 280, 340, 268, 352, 249, 351, 242"></a>
    <a href="consrioj.htm" class="menu6"><span><%=QtdeRiodeJaneiro%></span><area  href="consrioj.htm" shape="polygon" coords="339, 292, 330, 291, 323, 293, 317, 287, 334, 283, 340, 275, 348, 280, 349, 284"></a>
    <a href="consmats.htm" class="menu7"><span><%=QtdeMatoGSul%></span><area href="consmats.htm" shape="polygon" coords="188, 286, 203, 286, 209, 285, 217, 303, 222, 300, 223, 303, 236, 289, 248, 273, 252, 263, 250, 259, 233, 248, 228, 241, 221, 242, 209, 240, 203, 241, 193, 246, 193, 253"></a>
    <a href="consespi.htm" class="menu8"><span><%=QtdeEspiritoSanto%></span><area href="consespi.htm" shape="polygon" coords="361, 260, 360, 251, 354, 248, 341, 273, 351, 275"></a>
    <a href="consgoias.htm" class="menu9"><span><%=QtdeGoias%></span><area href="consgoias.htm"  shape="polygon" coords="261, 198, 270, 201, 276, 200, 285, 201, 299, 198, 300, 214, 294, 217, 273, 220, 275, 230, 291, 235, 289, 245, 281, 247, 274, 249, 259, 254, 253, 254, 239, 248, 236, 240, 242, 229, 251, 220, 258, 206, 260, 194"></a>
    <a href="consmato.htm" class="menu10"><span><%=QtdeMatoGrosso%></span><area href="consmato.htm" shape="polygon" coords="160, 162, 161, 174, 170, 181, 173, 187, 173, 194, 164, 203, 166, 209, 165, 217, 168, 226, 184, 227, 189, 236, 192, 241, 198, 235, 205, 234, 214, 238, 222, 237, 230, 235, 232, 241, 234, 239, 232, 233, 241, 225, 254, 207, 257, 191, 254, 180, 261, 169, 254, 166, 198, 161, 193, 154, 188, 147, 186, 156, 182, 160, 164, 159"></a>
    <a href="consbahi.htm" class="menu11"><span><%=QtdeBahia%></span><area href="consbahi.htm" shape="polygon" coords="305, 177, 302, 187, 305, 216, 323, 208, 335, 214, 360, 228, 356, 240, 362, 247, 369, 218, 372, 195, 382, 188, 373, 180, 377, 173, 373, 161, 366, 157, 355, 165, 346, 160, 337, 169, 327, 168, 324, 176, 318, 183"></a>   
    <a href="consbras.htm" class="menu12"><span><%=QtdeDistritoFeredal%></span><area  href="consbras.htm"  shape="polygon" coords="290, 223, 280, 221, 280, 229, 289, 230, 293, 225, 293, 228" ></a>
    <a href="constins.htm" class="menu13"><span><%=QtdeTocantins%></span><area href="constins.htm" shape="polygon" coords="279, 125, 287, 127, 287, 146, 289, 151, 296, 151, 293, 159, 301, 174, 296, 181, 299, 192, 292, 197, 277, 195, 268, 197, 266, 192, 260, 184, 268, 160, 274, 146, 283, 132"></a>
    <a href="consrond.htm" class="menu14"><span><%=QtdeRondonia%></span><area href="consrond.htm" shape="polygon" coords="138, 153, 158, 162, 158, 176, 168, 186, 168, 196, 163, 202, 153, 199, 143, 194, 133, 191, 124, 188, 121, 183, 119, 176, 111, 167, 122, 165, 131, 158"></a>
    <a href="consacre.htm" class="menu15"><span><%=QtdeAcre%></span><area href="consacre.htm" shape="polygon" coords="38, 148, 51, 161, 46, 163, 54, 165, 57, 170, 64, 170, 72, 162, 70, 176, 77, 178, 84, 178, 90, 181, 94, 175, 99, 176, 105, 171, 73, 152"></a>
    <a href="consamaz.htm" class="menu16"><span><%=QtdeAmazonas%></span><area href="consamaz.htm" shape="polygon" coords="77, 60, 85, 69, 77, 71, 78, 78, 83, 94, 75, 117, 49, 129, 43, 142, 76, 153, 107, 168, 110, 161, 120, 164, 135, 152, 141, 146, 154, 158, 184, 157, 188, 143, 182, 135, 198, 95, 188, 92, 179, 74, 173, 75, 169, 84, 164, 87, 160, 81, 156, 89, 145, 80, 146, 70, 138, 56, 123, 68, 104, 67, 100, 56, 94, 60, 85, 61"></a>
    <a href="consrora.htm" class="menu17"><span><%=QtdeRoraima%></span><area href="consrora.htm" shape="polygon" coords="125, 36, 147, 41, 153, 37, 165, 31, 167, 23, 171, 34, 177, 39, 172, 44, 172, 50, 176, 57, 182, 67, 174, 69, 168, 77, 164, 78, 154, 79, 153, 82, 149, 77, 149, 66, 145, 56, 139, 51, 131, 51, 128, 47"></a>
    <a href="consbele.htm" class="menu18"><span><%=QtdePara%></span><area href="consbele.htm" shape="polygon" coords="181, 66, 185, 77, 191, 90, 201, 91, 200, 102, 186, 136, 196, 153, 203, 161, 262, 165, 269, 152, 269, 138, 280, 131, 274, 124, 294, 103, 298, 88, 285, 82, 280, 90, 276, 95, 279, 81, 265, 77, 258, 79, 248, 89, 233, 77, 226, 65, 217, 54, 208, 54"></a>
    <a href="consamap.htm" class="menu19"><span><%=QtdeAmapa%></span><area href="consamap.htm" shape="polygon" coords="248, 36, 240, 57, 233, 57, 222, 55, 233, 64, 246, 85, 268, 60, 262, 58, 254, 45, 249, 39, 251, 36"></a>
    <a href="consmara.htm" class="menu20"><span><%=QtdeMaranhao%></span> <area href="consmara.htm" shape="polygon" coords="303, 90, 318, 96, 316, 104, 329, 100, 340, 104, 332, 113, 329, 137, 320, 137, 306, 148, 303, 159, 302, 172, 294, 161, 293, 157, 301, 150, 291, 145, 289, 126, 281, 122, 297, 105"></a>
    <a href="conspiau.htm" class="menu21"><span><%=QtdePiaui%></span><area href="conspiau.htm" shape="polygon" coords="343, 108, 352, 145, 349, 155, 336, 163, 324, 165, 317, 175, 306, 174, 309, 168, 308, 155, 311, 148, 324, 143, 333, 144, 334, 127, 337, 114"></a>
    <a href="conscear.htm" class="menu22"><span><%=QtdeCeara%></span><area href="conscear.htm" shape="polygon" coords="361, 106, 385, 121, 377, 127, 370, 136, 372, 146, 368, 147, 364, 145, 354, 145, 357, 141, 354, 137, 349, 123, 346, 111, 346, 106"></a>
    <a href="consnorte.htm" class="menu23"><span><%=QtdeRioGNorte%></span><area href="consnorte.htm" shape="polygon" coords="387, 125, 382, 125, 376, 134, 385, 133, 389, 137, 393, 132, 405, 136, 400, 127"></a>
    <a href="conspaba.htm" class="menu24"><span><%=QtdeParaiba%></span><area href="conspaba.htm" shape="polygon" coords="375, 138, 374, 148, 384, 145, 392, 150, 406, 144, 403, 139, 394, 139, 391, 144, 385, 141, 378, 138, 380, 138"></a>
    <a href="conspern.htm" class="menu25"><span><%=QtdePernanbuco%></span><area href="conspern.htm" shape="polygon" coords="353, 148, 351, 159, 356, 164, 366, 155, 378, 159, 389, 162, 406, 159, 408, 148, 400, 150, 393, 152, 383, 151, 381, 151, 372, 149, 368, 151"></a>
    <a href="consalag.htm" class="menu26"><span><%=QtdeAlagoas%></span><area href="consalag.htm" shape="polygon" coords="380, 165, 389, 166, 397, 163, 403, 163, 396, 172"></a>
    <a href="consserg.htm" class="menu27"><span><%=QtdeSergipe%></span><area href="consserg.htm" shape="polygon" coords="380, 172, 380, 181, 384, 184, 393, 175"></a>

   </map>
    <img polygon="(303,90) (318,96) (316,104) (329,100) (340,104) (332,113) (329,137) (320,137) (306,148) (303,159) (302,172) (294,161) (293,157) (301,150) (291,145) (289,126) (281,122) (297,105) consmara.htm" polygon="(343,108) (352,145) (349,155) (336,163) (324,165) (317,175) (306,174) (309,168) (308,155) (311,148) (324,143) (333,144) (334,127) (337,114) conspiau.htm" polygon="(361,106) (385,121) (377,127) (370,136) (372,146) (368,147) (364,145) (354,145) (357,141) (354,137) (349,123) (346,111) (346,106) conscear.htm" polygon="(380,172) (380,181) (384,184) (393,175) consserg.htm" polygon="(380,165) (389,166) (397,163) (403,163) (396,172) consalag.htm" polygon="(353,148) (351,159) (356,164) (366,155) (378,159) (389,162) (406,159) (408,148) (400,150) (393,152) (383,151) (381,151) (372,149) (368,151) conspern.htm" polygon="(375,138) (374,148) (384,145) (392,150) (406,144) (403,139) (394,139) (391,144) (385,141) (378,138) (380,138) conspaba.htm" polygon="(387,125) (382,125) (376,134) (385,133) (389,137) (393,132) (405,136) (400,127) consnorte.htm" polygon="(305,177) (302,187) (305,216) (323,208) (335,214) (360,228) (356,240) (362,247) (369,218) (372,195) (382,188) (373,180) (377,173) (373,161) (366,157) (355,165) (346,160) (337,169) (327,168) (324,176) (318,183) consbahi.htm" polygon="(361,260) (360,251) (354,248) (341,273) (351,275) consespi.htm" polygon="(339,292) (330,291) (323,293) (317,287) (334,283) (340,275) (348,280) (349,284) consrioj.htm" polygon="(353,231) (343,223) (327,218) (322,212) (308,220) (298,221) (294,234) (291,248) (282,250) (264,254) (256,260) (287,262) (294,265) (301,288) (316,283) (334,280) (340,268) (352,249) (351,242) consbelo.htm" polygon="(281,314) (276,306) (271,295) (264,290) (252,286) (243,287) (249,275) (256,266) (265,263) (270,265) (273,268) (288,265) (292,275) (296,280) (299,291) (315,292) (312,296) (308,299) (300,300) conssaop.htm" polygon="(261,198) (270,201) (276,200) (285,201) (299,198) (300,214) (294,217) (273,220) (275,230) (291,235) (289,245) (281,247) (274,249) (259,254) (253,254) (239,248) (236,240) (242,229) (251,220) (258,206) (260,194) consgoias.htm" polygon="(279,125) (287,127) (287,146) (289,151) (296,151) (293,159) (301,174) (296,181) (299,192) (292,197) (277,195) (268,197) (266,192) (260,184) (268,160) (274,146) (283,132) constins.htm" polygon="(192,362) (202,357) (207,362) (210,369) (215,365) (217,370) (235,385) (232,390) (236,395) (244,384) (244,379) (252,374) (266,353) (263,345) (251,336) (230,333) (226,334) consrios.htm" polygon="(227,332) (243,331) (255,335) (262,343) (266,344) (266,350) (273,345) (279,340) (279,322) (272,322) (264,323) (256,325) (251,325) (245,325) (238,324) (230,323) conssc.htm" polygon="(223,307) (222,317) (230,320) (254,325) (266,320) (277,321) (278,312) (271,310) (269,297) (262,295) (255,293) (241,290) conspara.htm" polygon="(188,286) (203,286) (209,285) (217,303) (222,300) (223,303) (236,289) (248,273) (252,263) (250,259) (233,248) (228,241) (221,242) (209,240) (203,241) (193,246) (193,253) consmats.htm" polygon="(160,162) (161,174) (170,181) (173,187) (173,194) (164,203) (166,209) (165,217) (168,226) (184,227) (189,236) (192,241) (198,235) (205,234) (214,238) (222,237) (230,235) (232,241) (234,239) (232,233) (241,225) (254,207) (257,191) (254,180) (261,169) (254,166) (198,161) (193,154) (188,147) (186,156) (182,160) (164,159) consmato.htm" polygon="(138,153) (158,162) (158,176) (168,186) (168,196) (163,202) (153,199) (143,194) (133,191) (124,188) (121,183) (119,176) (111,167) (122,165) (131,158) consrond.htm" polygon="(125,36) (147,41) (153,37) (165,31) (167,23) (171,34) (177,39) (172,44) (172,50) (176,57) (182,67) (174,69) (168,77) (164,78) (154,79) (153,82) (149,77) (149,66) (145,56) (139,51) (131,51) (128,47) consrora.htm" polygon="(181,66) (185,77) (191,90) (201,91) (200,102) (186,136) (196,153) (203,161) (262,165) (269,152) (269,138) (280,131) (274,124) (294,103) (298,88) (285,82) (280,90) (276,95) (279,81) (265,77) (258,79) (248,89) (233,77) (226,65) (217,54) (208,54) consbele.htm" polygon="(38,148) (51,161) (46,163) (54,165) (57,170) (64,170) (72,162) (70,176) (77,178) (84,178) (90,181) (94,175) (99,176) (105,171) (73,152) consacre.htm" polygon="(77,60) (85,69) (77,71) (78,78) (83,94) (75,117) (49,129) (43,142) (76,153) (107,168) (110,161) (120,164) (135,152) (141,146) (154,158) (184,157) (188,143) (182,135) (198,95) (188,92) (179,74) (173,75) (169,84) (164,87) (160,81) (156,89) (145,80) (146,70) (138,56) (123,68) (104,67) (100,56) (94,60) (85,61) consamaz.htm" src="mapabr-regiao.gif" border="0" usemap="#FPMap0" width="450" height="434"></td>
    <td width="26%"><font size="2" color="#FF4242" face="Arial"><strong>Clique sobre&nbsp; o
    Estado.</strong></font> </td>
  </tr>
</table>
</form>


</body>
</html>