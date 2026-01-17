<%
'Mensagem de bloqueio de sistema
'var_homologacao = true '<--------
IF strLoginRede <> "PRSS" and strLoginRede <> "DAVIF" and strLoginRede <> "FMAG#" and strLoginRede <> "ISASV" and strLoginRede <> "JOAOFNS#" THEN
	msg = "<p align=center><b><font color=#000080 face=Arial Black size=6>Sistema NewCLA</font></b></p>"
	msg = msg & "<p align=center><b><font color=#000080 face=Arial Black size=4>Relatórios em manutenção</font></b></p><br><br><br><center><h4><a href='http://ntspo920/newcla/Asp/main.asp'>Voltar</a>"
	Response.write msg
	response.end
END IF
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
menuWidth = 160;
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

function sair(){
	window.close();
} 

//-->
</SCRIPT>
<STYLE type ="text/css">
<!--
A.lnkmenu:link {font-family:Trebuchet MS, Arial; font-size:8pt; color:#000000; text-decoration:none}
A.lnkmenu:visited {font-family:Trebuchet MS, Arial; font-size:8pt; color:#000000; text-decoration:none}
A.lnkmenu:hover {font-family:Trebuchet MS, Arial; font-size:8pt; color:#1778bc; text-decoration:none}
-->
</style>


<title>CLA - Relatório de Acesso</title>
<base target="conteudo">
</head>
<body topmargin="0" leftmargin="0">
<form method="post" action="RelatorioSCla.asp" name="form_cla">
<A href="RelatorioSCla.asp" target=_self><IMG style="WIDTH: 100%; HEIGHT: 76px" height=100 src="Imagens/topo_embratel_relatorios.jpg" width="100%" border=0 ></A>
<table border="1" cellspacing="0" style="BORDER-RIGHT: #c0c0c0 1px solid; PADDING-RIGHT: 4px; BORDER-TOP: #c0c0c0 1px solid; PADDING-LEFT: 4px; PADDING-BOTTOM: 1px; BORDER-LEFT: #c0c0c0 1px solid;  PADDING-TOP: 1px; BORDER-BOTTOM: #c0c0c0 1px solid; BORDER-COLLAPSE: collapse" bordercolor="#c0c0c0" width="100%" id="AutoNumber1" bgcolor="#e8ecf0" cellpadding="0">
   <tr>
    <td align="middle" bordercolor="#000000" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" 
    width="100%">
    <p align="center">
    <A href="http://<%=Request.ServerVariables("SERVER_NAME")%>/newcla/Asp/main.asp" target=_self>
    
    Voltar ao Menu</a></p>
     </td>
  </tr>
</table>
</form>
</body>

</html>