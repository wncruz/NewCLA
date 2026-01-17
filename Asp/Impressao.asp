<%
Response.Expires = -1
Response.CacheControl = "no-cache"
%>
<HTML>
<HEAD>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<TITLE>CLA - Controle Local de Acesso</TITLE>
<SCRIPT LANGUAGE=javascript>
<!--
function Sair()
{
	window.close()
}
function Imprimir()
{
	window.print()
}
//-->
</SCRIPT>
</HEAD>
<BODY topmargin=5 leftmargin=0 class=TA>
<span id=spnLabelCons></span>
<span id=spnLista></span>
<SCRIPT LANGUAGE=javascript>
<!--
var objAryPram = window.dialogArguments
var strHtml = new String(objAryPram[1])
strHtml = Replace(objAryPram[1],"width=760","width=100% ")
strHtml = Replace(strHtml,"border=0","border=1 class=TableLine borderColor=black ")
spnLabelCons.innerHTML = "<p align=center><font face=Arial size=2><b>"+objAryPram[0]+"</b></font></p>"
spnLista.innerHTML = strHtml
//-->
</SCRIPT>
<table border=0 width=100% >
<tr>
	<td align=center><br>
		<input type="button" class="button" name="btnImprimir" value="Imprimir" onClick="Imprimir()" accesskey="W" onmouseover="showtip(this,event,'Imprimir (Alt+W)');">&nbsp;
		<input type="button" class="button" name="btnSair" value="Sair" onClick="Sair()" >
	</td>
</tr>
</table>
</BODY>
</HTML>