<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ConfirmacaoManutFac.asp
'	- Descrição			: Lista as facilidade que poderão ser incluídas.
%>
<HTML>
<HEAD>
<TITLE>Confirmação para Cadastro de Facilidades</TITLE>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
<SCRIPT LANGUAGE=javascript>
<!--
var objXml = window.dialogArguments
var strAtualiza = new String("<table border=0 cellspacing=1 cellpadding=1 width=100%>")

var objNode = objXml.selectNodes("//Facilidade")

var blnDet = false
var blnNDet = false
var blnAde = false
var intAtualiza = 0
var objNodeTeste


for(var intIndex=0;intIndex<objNode.length;intIndex++){
	switch (parseInt(objNode[intIndex].attributes[1].value)){
		case 1:
			if (!blnDet){
				strAtualiza += "<tr><th colspan=10>Confirmação para Cadastro de Facilidades</tr>"
				strAtualiza += "<tr>"
				strAtualiza += "<th>&nbsp;Fila</th>"
				strAtualiza += "<th>&nbsp;Bastidor</th>"
				strAtualiza += "<th>&nbsp;Régua</th>"
				strAtualiza += "<th>&nbsp;Posição</th>"
				strAtualiza += "<th>&nbsp;Domínio</th>"
				strAtualiza += "<th>&nbsp;Nó</th>"
				strAtualiza += "<th>&nbsp;Slot</th>"
				strAtualiza += "<th>&nbsp;Porta</th>"
				strAtualiza += "<th>&nbsp;Timeslot</th>"
				strAtualiza += "<th width=25 align=center>&nbsp;Sts</th>"
				blnDet = true
			}
			strAtualiza += "<tr class=clsSilver><td>" + objNode[intIndex].attributes[10].value + "</td>"
			strAtualiza += "<td>" + objNode[intIndex].attributes[2].value + "</td>"
			strAtualiza += "<td>" + objNode[intIndex].attributes[3].value + "</td>"
			strAtualiza += "<td>" + objNode[intIndex].attributes[4].value + "</td>"
			strAtualiza += "<td>" + objNode[intIndex].attributes[6].value + "</td>"
			strAtualiza += "<td>" + objNode[intIndex].attributes[7].value + "</td>"
			strAtualiza += "<td>" + objNode[intIndex].attributes[8].value + "</td>"
			strAtualiza += "<td>" + objNode[intIndex].attributes[9].value + "</td>"
			strAtualiza += "<td>" + objNode[intIndex].attributes[5].value + "</td>"
			if (parseInt(objNode[intIndex].attributes[0].value) == 1){
				strAtualiza += "<td align=center nowrap><img src='../imagens/table-.gif' border=0 alt='Já existe uma facilidade cadastrada essa caracteristica. Facilidade não pode ser incluída.'></td>"
			}else{	
				strAtualiza += "<td align=center><img src='../imagens/tableOk.gif' border=0 alt='A facilidade pode ser incluída.'></td>"
				intAtualiza += 1
			}	
			strAtualiza += "</tr>"
			break
		case 2:
			if (!blnDet){
				strAtualiza += "<tr><th colspan=3>Confirmação para Cadastro de Facilidades</tr>"
				strAtualiza += "<tr>"
				strAtualiza += "<th>&nbsp;Tronco</th>"
				strAtualiza += "<th>&nbsp;Par</th>"
				strAtualiza += "<th width=25 align=center>&nbsp;Sts</th>"
				blnDet = true
			}
			strAtualiza += "<tr class=clsSilver><td>" + objNode[intIndex].attributes[2].value + "</td>"
			strAtualiza += "<td>" + objNode[intIndex].attributes[3].value + "</td>"
			if (parseInt(objNode[intIndex].attributes[0].value) == 1){
				strAtualiza += "<td align=center nowrap><img src='../imagens/table-.gif' border=0 alt='Já existe uma facilidade cadastrada essa caracteristica. Facilidade não pode ser incluída.'></td>"
			}else{	
				strAtualiza += "<td align=center><img src='../imagens/tableOk.gif' border=0 alt='A facilidade pode ser incluída.'></td>"
				intAtualiza += 1
			}	
			strAtualiza += "</tr>"
			break
		case 3:
			if (!blnDet){
				strAtualiza += "<tr><th colspan=6>Confirmação para Cadastro de Facilidades</tr>"
				strAtualiza += "<tr>"
				strAtualiza += "<th>&nbsp;Cabo</th>"
				strAtualiza += "<th>&nbsp;Par</th>"
				strAtualiza += "<th>&nbsp;Tipo do Cabo</th>"
				strAtualiza += "<th>&nbsp;Derivação</th>"
				strAtualiza += "<th>&nbsp;PADE</th>"
				strAtualiza += "<th width=25 align=center>&nbsp;Sts</th>"
				blnDet = true
			}
			
				
			strAtualiza += "<tr class=clsSilver><td>" + objNode[intIndex].attributes[2].value + "</td>"
			strAtualiza += "<td>" + objNode[intIndex].attributes[3].value + "</td>"
					
			objNodeTeste = objNode[intIndex].getAttributeNode("TipoCabo")
			if  (objNodeTeste != null)
				strAtualiza += "<td>" + objNodeTeste.value + "</td>"
			else
				strAtualiza += "<td></td>"
				
			objNodeTeste = objNode[intIndex].getAttributeNode("Lateral")
			if  (objNodeTeste != null)
				strAtualiza += "<td>" + objNodeTeste.value + "</td>"
			else
				strAtualiza += "<td></td>"

			objNodeTeste = objNode[intIndex].getAttributeNode("CaixaEmenda")
			if  (objNodeTeste != null)
				strAtualiza += "<td>" + objNodeTeste.value + "</td>"
			else
				strAtualiza += "<td></td>"
						
			if (parseInt(objNode[intIndex].attributes[0].value) == 1){
				strAtualiza += "<td align=center nowrap><img src='../imagens/table-.gif' border=0 alt='Já existe uma facilidade cadastrada essa caracteristica. Facilidade não pode ser incluída.'></td>"
			}else{	
				strAtualiza += "<td align=center><img src='../imagens/tableOk.gif' border=0 alt='A facilidade pode ser incluída.'></td>"
				intAtualiza += 1
			}	
			strAtualiza += "</tr>"
			break
		case 999: //Interligacao
			if (!blnDet){
				strAtualiza += "<tr><th colspan=8>Confirmação para Cadastro de Facilidades</tr>"
				strAtualiza += "<tr>"
				strAtualiza += "<th>&nbsp;Origem</th>"
				strAtualiza += "<th>&nbsp;Esc Origem</th>"
				strAtualiza += "<th>&nbsp;Dist Origem</th>"
				strAtualiza += "<th>&nbsp;Destino</th>"
				strAtualiza += "<th>&nbsp;Esc Destino</th>"
				strAtualiza += "<th>&nbsp;Dist Destino</th>"
				strAtualiza += "<th width=25 align=center>&nbsp;Sts</th>"
				blnDet = true
			}
			strAtualiza += "<tr class=clsSilver><td nowrap>" + objNode[intIndex].attributes[2].value + "</td>"
			strAtualiza += "<td nowrap>" + objNode[intIndex].attributes[6].value + "</td>"
			strAtualiza += "<td nowrap>" + objNode[intIndex].attributes[4].value + "</td>"
			strAtualiza += "<td nowrap>" + objNode[intIndex].attributes[3].value + "</td>"
			strAtualiza += "<td nowrap>" + objNode[intIndex].attributes[7].value + "</td>"
			strAtualiza += "<td nowrap>" + objNode[intIndex].attributes[5].value + "</td>"

			if (parseInt(objNode[intIndex].attributes[0].value) == 1){
				strAtualiza += "<td align=center nowrap><img src='../imagens/table-.gif' border=0 alt='Já existe uma facilidade cadastrada essa caracteristica. Facilidade não pode ser incluída.'></td>"
			}else{	
				strAtualiza += "<td align=center><img src='../imagens/tableOk.gif' border=0 alt='A facilidade pode ser incluída.'></td>"
				intAtualiza += 1
			}	
			strAtualiza += "</tr>"
			break
	}
}	

strAtualiza += "</table>"

function Imprimir()
{
	window.print()
}
//-->
</SCRIPT>
</HEAD>
<BODY leftmargin=5 topmargin=5 class=TA >
<span id=spnFac></span>
<table width="100%" class=tableLine border=1>
<tr><th colspan=2>&nbsp;•&nbsp;Legenda</th></tr>
<tr>
	<td align="center" width=25>
		<img src="../imagens/tableOk.gif" border=0>
	</td>
	<td nowrap>
		A facilidade será incluída
	</td>
</tr>
<tr>
	<td align="center" width=25>
		<img src="../imagens/table-.gif" border=0>
	</td>
	<td nowrap>
		Já existe uma facilidade cadastrada com essas características. Facilidade não será incluída.
	</td>
</tr>
</table>	

<Form name=Form1 method=post>
<table width="100%" border=0>
<tr>
	<td align="center" height=25>
		<input type="button" class="button" name="btnAtualizar" value="Confirmar" onClick="window.returnValue=1;window.close();" accesskey="G" onmouseover="showtip(this,event,'Atualizar (Alt+G)');">&nbsp;
		<input type="button" class="button" name="btnImprimir" value="Imprimir" onClick="Imprimir()">&nbsp;
		<input type="button" class="button" name="btnSair" value="Sair" onClick="javascript:window.returnValue=0;window.close()" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>

</table>
</form>
<SCRIPT LANGUAGE=javascript>
<!--
spnFac.innerHTML = strAtualiza;
if (intAtualiza ==0) document.forms[0].btnAtualizar.disabled = true;
//-->
</SCRIPT>
</BODY>
</HTML>