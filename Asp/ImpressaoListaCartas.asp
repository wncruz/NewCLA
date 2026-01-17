<%
Response.Expires = -1
Response.CacheControl = "no-cache"
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ImpressaoListaCartas.ASP
'	- Responsável		: Vital
'	- Descrição			: Imprime a lista de cartas enviadas ao provedor
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
var objXmlGeral = window.dialogArguments

var strHtml = new String("<table border=1 borderColor=black cellspacing=1 cellpadding=1 width=100% class=TableLine>")

strProvedor = RequestNode(objXmlGeral,"Provedor")
strHtml += "<tr><th colspan=10>Lista de Email(s) Enviado(s) ao Provedor " + strProvedor + " </tr>"
strHtml += "<tr>"
strHtml += "<th>&nbsp;Cliente</th>"
strHtml += "<th>&nbsp;Pedido</th>"
strHtml += "<th>&nbsp;Data de Envio</th>"
strHtml += "<th>&nbsp;Processo</th>"
strHtml += "<th>&nbsp;Nº Acesso</th>"
strHtml += "<th>&nbsp;CCTO Prov</th>"
strHtml += "<th>&nbsp;Acesso Logico(678)</th>"
strHtml += "<th>&nbsp;CNL Cliente</th>"

strHtml += "</tr>"

var objNode = objXmlGeral.selectNodes("//Carta[Acao='I']")

for(var intIndex=0;intIndex<objNode.length;intIndex++){
	strHtml += "<tr class=clsSilver><td>" + objNode[intIndex].childNodes[1].text + "</td>"
	strHtml += "<td>" + objNode[intIndex].childNodes[2].text + "</td>"
	strHtml += "<td>" + objNode[intIndex].childNodes[3].text + "</td>"
	strHtml += "<td>" + objNode[intIndex].childNodes[4].text + "</td>"
	strHtml += "<td>" + objNode[intIndex].childNodes[6].text + "</td>"
	strHtml += "<td>" + objNode[intIndex].childNodes[7].text + "</td>"

	strHtml += "<td>" + objNode[intIndex].childNodes[8].text + "</td>"
	strHtml += "<td>" + objNode[intIndex].childNodes[9].text + "</td>"

	strHtml += "</tr>"

}
strHtml += "</table>"

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