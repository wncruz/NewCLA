<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
<SCRIPT LANGUAGE=javascript>
<!--

function ResgatarAcessoFisComp(intChave,objXmlAcessoFisComp)
{
	with (document.Form2)
	{
		hdnChaveAcessoFis.value = intChave
		hdnTecnologia.value = RequestNodeAcesso(objXmlAcessoFisComp,"cboTecnologia",intChave)
		hdnVelAcessoFisSel.value = RequestNodeAcesso(objXmlAcessoFisComp,"cboVelAcesso",intChave)
		hdnProvedor.value = RequestNodeAcesso(objXmlAcessoFisComp,"cboProvedor",intChave)
		strPropAcesso = RequestNodeAcesso(objXmlAcessoFisComp,"rdoPropAcessoFisico",intChave)

		if (strPropAcesso == "EBT")
		{
			divTecnologia.style.display = ""
		}
		else
		{
			if (divTecnologia.style.display == "")
			{
				cboTecnologia.value = ""
				divTecnologia.style.display = "none"
			}	
		}
		var strVel = RequestNodeAcesso(objXmlAcessoFisComp,"cboVelAcessoText",intChave)
		if (strVel == "2M" || strVel == "34M" || strVel == "155M" || strVel == "622M"){
			divTipoVel.style.display = ""
		}else{
			cboTipoVel.value = ""
			divTipoVel.style.display = "none"
		}	
		//Sinaliza que o id físico esta compartinhado par uma edição
		hdnNodeCompartilhado.value = RequestNodeAcesso(objXmlAcessoFisComp,"hdnCompartilhamento",intChave)
		hdnCompartilhamento.value = RequestNodeAcesso(objXmlAcessoFisComp,"hdnCompartilhamento",intChave)
		hdnNovoPedido.value = RequestNodeAcesso(objXmlAcessoFisComp,"hdnNovoPedido",intChave)
 		
		hdnAcao.value = "EditarAcessoFisComp"
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
}	 


function EditarAcessoFisComp(intChave,objXmlAcessoFisComp)
{
	var	objNode = objXmlAcessoFisComp.selectNodes("//xDados/Acesso[intIndice="+parseInt(intChave)+"]")
	//Refaz a lista de Ids no IFRAME
	for (var intIndex=0;intIndex<objNode.length;intIndex++){
		for (var intIndexII=0;intIndexII<objNode[intIndex].childNodes.length;intIndexII++){
			try{
				//Caso de radio button
				if (objNode[intIndex].childNodes[intIndexII].nodeName == "rdoPropAcessoFisico"){
					eval("document.Form2."+objNode[intIndex].childNodes[intIndexII].nodeName+"["+RequestNodeAcesso(objXmlAcessoFisComp,"rdoPropAcessoFisicoIndex",objNode[intIndex].childNodes[0].text)+"].checked = true")
				}
				else{
					//eval("document.Form2."+objNode[intIndex].childNodes[intIndexII].nodeName+".value='"+objNode[intIndex].childNodes[intIndexII].text+"'")
					var objChildForm = new Object(eval("document.Form2."+objNode[intIndex].childNodes[intIndexII].nodeName))
					//alert(objChildForm.value + " = " + objNode[intIndex].childNodes[intIndexII].nodeName)
					objChildForm.value = objNode[intIndex].childNodes[intIndexII].text
				}	
			}catch(e){}
		}	
	}	

	if (objNode.length > 0)
	{
		var intChaveFis = RequestNodeAcesso(objXmlAcessoFisComp,"Aec_Id",intChave)
		if (intChaveFis == "") intChaveFis = 0
		var objNodeFis = objXmlAcessoFisComp.selectNodes("//xDados/Acesso/IdFisico[Aec_Id="+intChaveFis+"]")
		var strAcessoIdFis =  new String("<table border=0 width=759 cellspacing=0 cellpadding=0>")
		//Refaz a lista de Ids no IFRAME
		for (var intIndex=0;intIndex<objNodeFis.length;intIndex++)
		{
			var intAcfId = objNodeFis[intIndex].childNodes[0].text
			var intAecId = objNodeFis[intIndex].childNodes[2].text
			var objNodePed = objXmlAcessoFisComp.selectNodes("//xDados/Acesso/Pedido[Acf_Id="+intAcfId+"]")
			strAcessoIdFis += "<tr class=clsSilver>"
			strAcessoIdFis += "<td >&nbsp;Pedido</td>"
			strAcessoIdFis += "<td>&nbsp;"+objNodePed[0].childNodes[1].text +"</td>"
			strAcessoIdFis += "<td >&nbsp;ID Físico</td>"
			strAcessoIdFis += "<td >"
			try{
				strAcessoIdFis +=objNodeFis[intIndex].childNodes[3].text
			}catch(e){}	
			strAcessoIdFis +="</td>"
			strAcessoIdFis += "<td >&nbsp;Nº Acesso</td>"
			strAcessoIdFis += "<td >"
			try{
				strAcessoIdFis += objNodeFis[intIndex].childNodes[4].text
			}catch(e){}	
			strAcessoIdFis +="</td>"
			strAcessoIdFis += "</tr>"
			strAcessoIdFis += "<tr></tr>"
		}
		strAcessoIdFis += "</table>"
		//Qtde circuitos
		var objNodeCircuitos = objXmlAcessoFisComp.selectNodes("//xDados/Acesso/Pedido[Aec_Id="+intAecId+"]")
		if (objNodeCircuitos.length > 0)
		{
			var intQtdeCir = 1
			var intQtdeCirAtual = RequestNodeAcesso(objXmlAcessoFisComp,"txtQtdeCircuitos",intChave)
			if (intQtdeCirAtual == ""){intQtdeCirAtual = 1}
			var intPedId = objNodeCircuitos[0].childNodes[0].text
			for (var intIndex=0;intIndex<objNodeCircuitos.length;intIndex++)
			{
				if (intPedId != objNodeCircuitos[intIndex].childNodes[0].text)
				{
					intQtdeCir += 1
				}
			}
			if (intQtdeCirAtual == intQtdeCir)
			{
				document.Form2.txtQtdeCircuitos.value = intQtdeCir
			}else
			{
				document.Form2.txtQtdeCircuitos.value = intQtdeCirAtual
			}	
		}
	}
	else{
		strAcessoIdFis = ""
	}	

	//document.Form2.hdnIntIndice.value = intChave //Chave Atual no Html não será editável mas item novo
	//Acerta o disabled para o provedor quando temos EBT
	if (RequestNodeAcesso(objXmlAcessoFisComp,"rdoPropAcessoFisico",intChave) != "EBT"){
		document.forms[1].cboProvedor.disabled = false
	}
	else{
		document.forms[1].cboProvedor.disabled = true
	}
	parent.spnListaIdFis.innerHTML = strAcessoIdFis

}
//-->
</SCRIPT>


</HEAD>
<BODY>

<P>&nbsp;</P>

</BODY>
</HTML>
