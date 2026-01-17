<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ConfirmacaoAlocacaoFac.asp
'	- Responsável		: Vital
'	- Descrição			: Mostra mensagem da alocação de Facilidade
%>
<HTML>
<HEAD>
<TITLE>CLA - Controle Local de Acesso</TITLE>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
<SCRIPT LANGUAGE=javascript>
<!--
var objXml = window.dialogArguments
var strAtualiza = new String("<table border=0 cellspacing=1 cellpadding=1 width=100%>")

var objNode = objXml.selectNodes("//CLA_RetornoTmp")

var blnDet = false
var blnNDet = false
var blnAde = false
var intAtualiza = 0

strAtualiza += "<tr><th colspan=10>&nbsp;•&nbsp;Informações</tr>"
strAtualiza += "<tr>"
strAtualiza += "<th>&nbsp;Nº</th>"
strAtualiza += "<th>&nbsp;Ponto</th>"
strAtualiza += "<th>&nbsp;Mensagem</th>"
strAtualiza += "<th>&nbsp;Valor</th>"
strAtualiza += "<th width=25 align=center>&nbsp;Sts</th>"
strAtualiza += "</tr>"

for(var intIndex=0;intIndex<objNode.length;intIndex++){
	strAtualiza += "<tr class=clsSilver><td>" + objNode[intIndex].attributes[0].value + "</td>"
	var intPonto = objNode[intIndex].attributes[1].value
	strAtualiza += "<td>" + IIf(parseInt(intPonto)>0,intPonto,"") + "</td>"
	strAtualiza += "<td>" + objNode[intIndex].childNodes[0].attributes[0].value + "</td>"
	strAtualiza += "<td>" + objNode[intIndex].attributes[2].value + "</td>"
	switch (parseInt(objNode[intIndex].attributes[3].value)){
		case 0:
			strAtualiza += "<td align=center><img src='../imagens/info.gif' border=0 alt='Informativo.'></td>"
			break
		case 1:
			strAtualiza += "<td align=center><img src='../imagens/erro.gif' border=0 alt='Erro.'></td>"
			break
	}
	strAtualiza += "</tr>"
}

strAtualiza += "</table>"

function Imprimir()
{
	window.print()
}

objNode = objXml.selectNodes("//CLA_RetornoTmp")
var intRet = new String("")
if (objNode.length > 0){
	intRet = objNode[0].attributes[0].value
}
//-->
</SCRIPT>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function window_onunload() {
	window.returnValue = intRet
}
//-->
</SCRIPT>
</HEAD>
<BODY leftmargin=5 topmargin=5 class=TA LANGUAGE=javascript onunload="return window_onunload()">
<span id=spnFac></span>
<table width="100%" class=tableLine border=1>
<tr><th colspan=2>&nbsp;•&nbsp;Legenda</th></tr>
<tr>
	<td align="center" width=25>
		<img src="../imagens/info.gif" border=0>
	</td>
	<td nowrap>
		Informativo.
	</td>
</tr>
<tr>
	<td align="center" width=25>
		<img src="../imagens/erro.gif" border=0>
	</td>
	<td>
		Erro.
	</td>
</tr>
</table>	

<Form name=Form1 method=post>
<table width="100%" border=0>
<tr>
	<td align="center" height=25>
		<input type="button" class="button" name="btnImprimir" value="Imprimir" onClick="Imprimir()">&nbsp;
		<input type="button" class="button" name="btnSair" value="Sair" onClick="javascript:window.returnValue=intRet;window.close()" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
</table>
</form>
<SCRIPT LANGUAGE=javascript>
<!--
spnFac.innerHTML = strAtualiza;
//-->
</SCRIPT>
</BODY>
</HTML>