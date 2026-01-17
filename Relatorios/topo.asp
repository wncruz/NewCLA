<%



%>
<html>

<head>
<SCRIPT LANGUAGE="JavaScript">

<!--
     NS4 = (document.layers);
     IE4 = (document.all);
    ver4 = (NS4 || IE4);
   isMac = (navigator.appVersion.indexOf("Mac") != -1);
  isMenu = (NS4 || (IE4 && !isMac));
  function popUp(){return};
  function popDown(){return};
  function startIt(){return};
  if (!ver4) event = null;

  if (isMenu) {
menuVersion = 3;
menuWidth = 120;
borWid = 1;
borSty = "double";
borCol = "#c0c0c0";
separator = 1;
separatorCol = "#c0c0c0";
fntFam = "Trebuchet MS";
fntBold = false;
fntItal = false;
fntSiz = 9;
fntCol = "#FFFFFF";
overFnt = "#FAEB00";
itemPad = 3;
backCol = "#1778bc";
overCol = "#1778bc";
imgSrc = "tri.gif";
imgSiz = 10;

childOffset = 6;
childOverlap = 60;
perCentOver = null;

clickStart = false;
clickKill = false;
secondsVisible = 0.5;

keepHilite = false;
NSfontOver = false;

isFrames = false;
navFrLoc = "top";
mainFrName = "main";
}
function abre(n){
	WD=open(n,'nova','toolbar=no,location=no,directories=no,menubar=yes,status=no,scrollbars=yes,resizable=yes,copyhistory=no,width=779,height=380');
	WD.focus();
}  

function abre_senha(n){
	WD=open(n,'nova','toolbar=no,location=no,directories=no,menubar=no,status=no,scrollbars=no,resizable=no,copyhistory=no,width=360,height=250');
	WD.focus();
}  
  
function selec_evt(){
	document.form_agencia.submit();
}
//-->
</SCRIPT>


<SCRIPT LANGUAGE="JavaScript1.2" SRC="exemplos/links-organizador.js" DEFER></SCRIPT>
<SCRIPT LANGUAGE="JavaScript1.2" SRC="exemplos/topnav.js" DEFER></SCRIPT>

<STYLE type ="text/css">
<!--
A.lnkmenu:link {font-family:Trebuchet MS, Arial; font-size:8pt; color:#000000; text-decoration:none}
A.lnkmenu:visited {font-family:Trebuchet MS, Arial; font-size:8pt; color:#000000; text-decoration:none}
A.lnkmenu:hover {font-family:Trebuchet MS, Arial; font-size:8pt; color:#1778bc; text-decoration:none}
-->
</style>


<title>Relatórios CLA</title>
<base target="conteudo">
</head>
<body topmargin="0" leftmargin="0">
<img border="0" src="Imagens/topo_embratel.jpg" width="100%" height="80">
<form method="POST" action="exemplos/testemenu.htm" name="form_agencia">
<table border="1" cellspacing="0" style="border-collapse: collapse; border: 1px solid #C0C0C0; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1" bordercolor="#C0C0C0" width="784" id="AutoNumber1" bgcolor="#E8ECF0" height="10" cellpadding="2">
  <tr>

     <td align="center" bordercolor="#000000" style="border-style: solid; border-width: 1" width="10">
    </td>
    <td align="center" bordercolor="#000000" style="border-style: solid; border-width: 1" width="55">
    <a class="lnkmenu" target="superior" href="topo.asp" onMouseOut="popDown('eMenu1')" onMouseOver="popUp('eMenu1',event)">
    Acessos</a></td>
    
    <!--
    <td align="center" bordercolor="#000000" style="border-style: solid; border-width: 1" width="71">
    <a class="lnkmenu" target="superior" href="menu-organizador.asp" onMouseOut="popDown('eMenu2')" onMouseOver="popUp('eMenu2',event)">
    funcionário</a></td>
    
    <td align="center" bordercolor="#000000" style="border-style: solid; border-width: 1" width="89">
    <a class="lnkmenu" target="superior" href="menu-organizador.asp" onMouseOut="popDown('eMenu3')" onMouseOver="popUp('eMenu3',event)">
    agência oficial</a></td>
    
    <td align="center" bordercolor="#000000" style="border-style: solid; border-width: 1" width="43">
    <a class="lnkmenu" target="superior" href="menu-organizador.asp" onMouseOut="popDown('eMenu4')" onMouseOver="popUp('eMenu4',event)">
    evento</a></td>
    
    <td align="center" bordercolor="#000000" style="border-style: solid; border-width: 1" width="56">
    <a class="lnkmenu" target="superior" href="menu-organizador.asp" onMouseOut="popDown('eMenu5')" onMouseOver="popUp('eMenu5',event)">
    inscrição</a></td>
    
    <td align="center" bordercolor="#000000" style="border-style: solid; border-width: 1" width="79">
    <a class="lnkmenu" target="superior" href="menu-organizador.asp" onMouseOut="popDown('eMenu6')" onMouseOver="popUp('eMenu6',event)">
    gerar boleto</a></td>
    
    <td align="center" bordercolor="#000000" style="border-style: solid; border-width: 1" width="60">
    <a class="lnkmenu" target="superior" href="organizador-consultas.asp">
    consultas</a></td>
    !-->
    <td align="center" bordercolor="#000000" style="border-style: solid; border-width: 1" width="275">
    <p align="left">
    <a href="Javascript:window.close();">
    
    sair</a></p>
     </td>
  </tr>
</table>
</form>
</body>

</html>