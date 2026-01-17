function ResgatarPromocaoRegime(obj)
{
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarPromocaoRegime"
		hdnProvedor.value = obj.value
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
}

function ResgatarDistribuicao(obj,intDistrib)
{
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarDistribuicao"
		hdnEstacaoAtual.value = obj.value

		if (intDistrib == "1")
		{
			hdnNomeLocal.value = "DistLocalInstala"
		}
		else
		{
			hdnNomeLocal.value = "DistLocalConfig"
		}
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
}

function LimparCamposFac()
{
	with (document.forms[0]) 
	{
		if (cboRede.value != "")
		{
			switch (cboRede.value) 
			{
				case "1": //Det
					cboCodProv.value = ''
					txtNroAcessoEbt.value = ''
					txtBastidor.value = ''
					txtRegua.value = ''
					txtPorta.value = ''
					txtTimeslot.value = ''
					for (var intIndex=0;intIndex<rdoUrbano.length;intIndex++){
						rdoUrbano[intIndex].checked = false
					}	
					txtCCTOPro.value = ''
					txtNroAcessoCli.value = ''
					txtCNLPtaA.value = ''
					txtCNLPtaB.value = ''
					txtFila.value = ''
					//Limpar posições do array
					objAryFac[0] = new Array("","")
					objAryFac[1] = new Array("","")
					objAryFac[2] = new Array("","")
					objAryFac[3] = new Array("","")
					objAryFac[4] = new Array("","")
					break

				case "2": //NDet
					cboCodProv.value = ''
					txtNroAcessoEbt.value = ''
					txtTronco.value = ''
					txtPar.value = ''
					//eval("spnStsNDet.innerHTML = '&nbsp;'")
					for (var intIndex=0;intIndex<rdoUrbano.length;intIndex++){
						rdoUrbano[intIndex].checked = false
					}	
					txtCCTOPro.value = ''
					txtNroAcessoCli.value = ''
					txtCNLPtaA.value = ''
					txtCNLPtaB.value = ''
					//Limpar posições do array
					objAryFac[0] = new Array("","")
					objAryFac[1] = new Array("","")
					objAryFac[2] = new Array("","")
					objAryFac[3] = new Array("","")
					objAryFac[4] = new Array("","")
					break

				case "3": //ADE
					txtNroAcessoEbt.value = ''
					txtCabo.value = ''
					txtLateral.value = ''
					cboTipoCabo.value = ''
					txtCaixaEmenda.value = ''
					txtPar.value = ''
					//eval("spnStsAde.innerHTML = '&nbsp;'")
					//Limpar posições do array
					objAryFac[0] = new Array("","")
					objAryFac[1] = new Array("","")
					objAryFac[2] = new Array("","")
					objAryFac[3] = new Array("","")
					objAryFac[4] = new Array("","")
					break
			}	
		}
	}	

	for (var intIndexII=0;intIndexII<IFrmFacilidade.document.Form1.elements.length;intIndexII++)
	{
		var elemento = IFrmFacilidade.document.Form1.elements[intIndexII];
		if (elemento.type == 'radio')
		{
			elemento.checked = false
		}
	}
}

function ApagarNroAcesso(intIndex)
{
	with (document.forms[0]) 
	{
		if (cboRede.value != "")
		{
			switch (cboRede.value) 
			{
				case "1": //Det
					txtNroAcessoEbt.value = ''
					break

				case "2": //NDet
					txtNroAcessoEbt.value = ''
					break

				case "3": //ADE
					txtNroAcessoEbt.value = ''
					break
			}	
		}
	}	
}


function ResgatarPadraoProvedor(obj,intProd)
{
	with (document.forms[0])
	{
		/*
		if (intProd == 0)
		{
			for (var intIndex=1;intIndex<5;intIndex++)
			{
				ApagarNroAcesso(intIndex)
			}
		}
		*/
		if (obj == '[object]'){
			hdnProvedor.value = obj.value
		}else{
			hdnProvedor.value = obj
		}	
		hdnAcao.value = "ResgatarPadraoProvedor"
		target = "IFrmProcesso1"
		action = "ProcessoCla.asp"
		submit()	
	}
}	

function ProcurarIDFis(intID)
{
	with (document.forms[0])
	{
		switch (intID)
		{
			case 1:
				hdnIdAcessoFisico.value = ""
				target = "IFrmIDFis1"
				action = "Estoque.asp?intEnd=1"
				submit()
				break
			case 2:	
				hdnIdAcessoFisico.value = ""
				target = "IFrmIDFis1"
				action = "AcessoCompartilhadoFac.asp?intEnd=1"
				submit()
				break
		}
	}	
}

function SelIDFisComp(obj,intEnd,IdAcFis )
{
	with (document.forms[0])
	{
		switch (intEnd)
		{
			case 1: //id's Físicos 
				hdnIdAcessoFisico.value=obj.value
				hdnPropIdFisico.value = obj.prop
				hdnCompartilhamento.value = "0"
				hdnAcao.value = "AutorizarCompartilhamento"
				hdnIdAcessoFisico.value = IdAcFis 				
				hdnSubAcao.value = "IdFisEndInstala"
				target = "IFrmProcesso"
				action = "ProcessoCla.asp"
				submit()
				break
			case 2:	//Estoque
				hdnIdAcessoFisico1.value=obj.value
				hdnPropIdFisico1.value = obj.prop
				hdnIdAcessoFisico.value = IdAcFis 				
				hdnAcao.value = "ResgatarEstoque"
				target = "IFrmProcesso"
				action = "ProcessoResgatarFac.asp"
				submit()
				break
		}
	}	
}

function EnviarEmailProvedor(intTipo)
{
	with (document.forms[0])
	{
		if (cboProvedor.value != "")
		{
			switch (parseInt(intTipo))
			{
				case 0:
					target = self.name 
					action = "emailprovedor.asp"
					submit()
					break
				case 1:
					target = self.name 
					action = "emailatt.asp"
					submit()
					break
			}
		}
		else
		{
			alert("Selecione o Provedor!")
			cboProvedor.focus()
			return
		}
			
	}	
}
function AbilitarEmail()
{
	with (document.forms[0])
	{
		if (cboProvedor.value != 2)
		{
			btnEmailPro.disabled = false
			btnEmailAtt.disabled = true
		}	
		else
		{	
			btnEmailAtt.disabled = false
			btnEmailPro.disabled = true
		}
	}
}

function limparIDFisico(intID)
{
	switch (parseInt(intID))
	{
		case 1:
			with (IFrmIDFis1.document.forms[0])
			{
				if (rdoIDFis1 == '[object]')
				{
					rdoIDFis1.checked = false
				}
				for (var intIndex=0;intIndex<rdoIDFis1.length;intIndex++)
				{
					rdoIDFis1[intIndex].checked = false	
				}
			}	
			break
		case 2:
			with (IFrmIDFis2.document.forms[0])
			{
				if (rdoIDFis2 == '[object]')
				{
					rdoIDFis2.checked = false
				}
				for (var intIndex=0;intIndex<rdoIDFis2.length;intIndex++)
				{
					rdoIDFis2[intIndex].checked = false	
				}
			}	
			break
	}	
}

function AdicionarCNL(obj)
{
	with (document.forms[0])
	{
		if (rdoUrbano[1].checked)
		{
			if (cboLocalConfig.value != "")
			{
				txtCNLPtaA.value = cboLocalConfig[cboLocalConfig.selectedIndex].text.split(" ")[0]
				txtCNLPtaB.value = cboLocalConfig[cboLocalConfig.selectedIndex].text.split(" ")[0]
			}	
		}
		else
		{
			txtCNLPtaA.value = cboLocalConfig[cboLocalConfig.selectedIndex].text.split(" ")[0]
			txtCNLPtaB.value = ''
		}
	}	
}

function PadraoMax(obj,intTam)
{
	if (event.keyCode > 46)
		event.returnValue = (obj.value.length < intTam);
}

function ReenviarSolicitacao(intRetASP,intRetJS)
{
	with (document.forms[0])
	{
		switch (parseInt(intRetASP))
		{
			case 138: //Endereco de instalação
				switch (parseInt(intRetJS))
				{
					case 1: //Aceito
						hdnCompartilhamento.value = "1"
						//hdnIdAcessoFisico.value=obj.value
						//hdnPropIdFisico.value = obj.prop
						hdnAcao.value = "ResgatarIdFisicoComp"
						target = "IFrmProcesso"
						action = "ProcessoResgatarFac.asp"
						submit()
						cboLocalInstala.disabled = true
						cboLocalConfig.disabled = true
						cboDistLocalInstala.disabled = true
						cboRede.disabled = true
						cboProvedor.disabled = true
						consultar.disabled = true
						btnConsEstoque.disabled = true
						ocupados.disabled = true
						btnConsIdFis.disabled = true
						cboRegimeCntr.disabled = true
						cboPromocao.disabled = true
						btnEmailAtt.disabled = true
						btnEmailPro.disabled = true
						break
					
					case 2: // Não Aceito
						hdnIdAcessoFisico.value = ""
						hdnPropIdFisico.value  = ""
						hdnCompartilhamento.value = "0"
						limparIDFisico(1)
						AbrirParam(cboRede)
						break
				}
				break
		
			case 139: //Endereco do ponto intermediário
				switch (parseInt(intRetJS))
				{
					case 1: //Aceito
						hdnCompartilhamento1.value = "1"
						break
					
					case 2: // Não Aceito
						hdnIdAcessoFisico1.value = ""
						hdnPropIdFisico1.value  = ""
						hdnCompartilhamento1.value = "0"
						btnIDFis2.focus()
						limparIDFisico(2)
						break
				}
				break
		}	
	}
}
/*
=====================================================================================
Objetivo		: Cria/Atualiza nodes no Xml de Acessos
Premissas		: Existir objeto XML na página que faz o include
Entradas		: blnLimpar - booleana se limpa os campos do acesso ou não
Retorno			: Atualiza lista com item adicionado
Data - Autor	: 20/06/2002 - Vital
=====================================================================================
*/
function xmlFacUpd(blnLimpar,intRede)
{
	for (var intIndex=0;intIndex<objAryObjs[intRede].length;intIndex++)
	{
		var elemento = new Object(eval("document.forms[0]."+objAryObjs[intRede][intIndex]))
		if (elemento.type == "radio" || elemento.type == "checkbox"){
			if (elemento.checked){
				if (document.forms[0].hdnIntIndice.value != ""){
					AdicionarNodeFac(objXmlGeral,elemento.name,elemento.value,document.forms[0].hdnIntIndice.value)
					AdicionarNodeFac(objXmlGeral,elemento.name+"Index",elemento.Index,document.forms[0].hdnIntIndice.value)
				}
				else{
					AdicionarNodeFac(objXmlGeral,elemento.name,elemento.value,intIndice)
					AdicionarNodeFac(objXmlGeral,elemento.name+"Index",elemento.Index,intIndice)
				}	
			}	
		}
		else{
			if (document.forms[0].hdnIntIndice.value != ""){
				AdicionarNodeFac(objXmlGeral,elemento.name,elemento.value,document.forms[0].hdnIntIndice.value)
			}
			else{
				AdicionarNodeFac(objXmlGeral,elemento.name,elemento.value,intIndice)
			}	
		}	
		//Cria um node com o text dos combos
		if (elemento.type == "select-one"){
			if (document.forms[0].hdnIntIndice.value != ""){
				AdicionarNodeFac(objXmlGeral,elemento.name+"Text",elemento[elemento.selectedIndex].text,document.forms[0].hdnIntIndice.value)
			}
			else{	
				AdicionarNodeFac(objXmlGeral,elemento.name+"Text",elemento[elemento.selectedIndex].text,intIndice)
			}
		}

		if (blnLimpar){
			//Não pode limpar o conteúdo do radio somente fazer checked=false
			if (elemento.type != 'button' && elemento.type != 'hidden' && elemento.type != 'radio' && elemento.name != 'cboRede') elemento.value = ""
			if (elemento.type == 'radio')
			{
				elemento.checked = false
			}
		}	
	}
	if (document.forms[0].hdnIntIndice.value != ""){
		if (RequestNodeFac(objXmlGeral,"Fac_Id",document.forms[0].hdnIntIndice.value) != ""){
			AdicionarNodeFac(objXmlGeral,'TipoAcao','A',document.forms[0].hdnIntIndice.value)
		}else{
			AdicionarNodeFac(objXmlGeral,'TipoAcao','N',document.forms[0].hdnIntIndice.value)
		}	
	}
	else{	
		AdicionarNodeFac(objXmlGeral,'TipoAcao','N',intIndice)
	}
	intIndice += 1
	AtualizarListaFac(document.forms[0].Acf_Id.value)
	document.forms[0].hdnIntIndice.value = ""
}

/*
=====================================================================================
Objetivo		: Adiciona um Node a um objeto XML. Se existir atualiza
Premissas		: Enviar o objeto XML
Entradas		: objXMLProcesso - objeto XML
				  strNomeNode - Nome do Node
				  strValorNode - Valor do node 
				  intChave - Chave para identificação/criação do node
Retorno			: 
Data - Autor	: 20/06/2002 - Vital
=====================================================================================
*/
function AdicionarNodeFac(objXMLProcesso,strNomeNode,strValorNode,intChave)
{	
	var objElemento
	var objNodeFilho
	var intIndex
	var objNodeList
	var objNode
    var objXML = objXMLProcesso
    if (objXML.xml == "")
    {
	   objXML.loadXML("<xDados></xDados>")
	}   
	//Verifica se já existe
	objNodeList = objXML.selectNodes("//xDados/Facilidade/Interligacao[Int_Id="+parseInt(intChave)+"]")
	if (objNodeList.length == 0)
	{
		//Cria o Node Chave
		objNode = objXML.createNode("element", "Facilidade", "")
		objXML.documentElement.appendChild (objNode)

		objNodeFilho = objXML.createNode("element", "intIndice", "")
		objNodeFilho.text = parseInt(intChave)
		objNode.appendChild (objNodeFilho)
	}	
	else
	{
		objNode = objXML.selectSingleNode("//xDados/Facilidade/Interligacao[Int_Id="+parseInt(intChave)+"]")
	}

	var objNodeListAtual = objNode.getElementsByTagName(strNomeNode)
	if (objNodeListAtual.length == 0)
	{
		//Cria
		objNodeFilho = objXML.createNode("element", strNomeNode, "")
		objNodeFilho.text = strValorNode
		objNode.appendChild (objNodeFilho)
	}	
	else
	{
		//Atualiza
		objNodeListAtual.item(0).text = strValorNode
	}	
}
/*
=====================================================================================
Objetivo		: Popula os campos de acesso físico
Premissas		: Existir objeto XML na página que faz o include
Entradas		: intChave -  Chave identificadora do item editado
Retorno			: Atualiza campos do formulário com informção
Data - Autor	: 20/06/2002 - Vital
=====================================================================================
*/
function EditarFac(obj){
	
	if (obj == '[object]'){
		var intChave = obj.value
	}else{
		var intChave = obj
	}	
	var	objNode = objXmlGeral.selectNodes("//xDados/Facilidade/Interligacao[Int_Id="+parseInt(intChave)+"]")
	//Refaz a lista de Ids no IFRAME
	for (var intIndex=0;intIndex<objNode.length;intIndex++){
		for (var intIndexII=0;intIndexII<objNode[intIndex].childNodes.length;intIndexII++){
			try{
				//Caso de radio button
				if (objNode[intIndex].childNodes[intIndexII].nodeName == "rdoUrbano"){
					eval("document.forms[0]."+objNode[intIndex].childNodes[intIndexII].nodeName+"["+RequestNodeFac(objXmlGeral,"rdoUrbanoIndex",objNode[intIndex].childNodes[0].text)+"].checked = true")
				}
				else{
					//eval("document.forms[0]."+objNode[intIndex].childNodes[intIndexII].nodeName+".value='"+objNode[intIndex].childNodes[intIndexII].text+"'")
					var objChildForm = new Object(eval("document.forms[0]."+objNode[intIndex].childNodes[intIndexII].nodeName))
					objChildForm.value = objNode[intIndex].childNodes[intIndexII].text
				}	
			}catch(e){}
		}	
	}	
	document.forms[0].hdnIntIndice.value = intChave //Chave Atual no Html
	if (document.forms[0].cboRede.value == 1 || document.forms[0].cboRede.value == 2){
		ResgatarPadraoProvedor(document.forms[0].cboCodProv,0)
	}else{
		ResgatarPadraoProvedor(document.forms[0].cboProvedor,0)
	}
}
/*
=====================================================================================
Objetivo		: Remove acesso do XML
Premissas		: Existir objeto XML na página que faz o include
Entradas		: intIndice - Chave o item a ser removido
				  blnLimpar - booleana para limpar os campos do formulário depois de remover
Retorno			: Atualiza lista do Iframe de acessos adicionados
Data - Autor	: 20/06/2002 - Vital
=====================================================================================
*/
function RemoverFac(intIndice,blnLimpar,intRede)
{

	if (RequestNodeFac(objXmlGeral,"Fac_Id",intIndice) != ""){
		AdicionarNodeFac(objXmlGeral,'TipoAcao','R',intIndice)
	}
	else
	{
		var	objNode = objXmlGeral.selectNodes("//xDados/Facilidade/Interligacao[Int_Id="+parseInt(intIndice)+"]")
		if (objNode.length > 0){
			objNode[0].parentNode.removeChild(objNode[0])
		}
		else
		{
			alert("Facilidade não encontrada.")
			return
		}	

	}

	AtualizarListaFac(document.forms[0].Acf_Id.value)
	document.forms[0].hdnIntIndice.value = ""
	for (var intIndex=0;intIndex<objAryObjs[intRede].length;intIndex++)
	{
		var elemento = new Object(eval("document.forms[0]."+objAryObjs[intRede][intIndex]))
		if (blnLimpar){
			if (elemento.type != 'button' && elemento.type != 'hidden' && elemento.type != 'radio' && elemento.name != 'cboRede')	elemento.value = ""
			if (elemento.type == 'radio')
			{
				elemento.checked = false
			}
		}	
	}
}
/*
=====================================================================================
Objetivo		: Adiciona item a lista que acessos adicionados
Premissas		: Existir objeto XML na página que faz o include
Entradas		: 
Retorno			: Mostra Html no Iframe de acessos adicionados
Data - Autor	: 20/06/2002 - Vital
=====================================================================================
*/
function AtualizarListaFac(intChave){

	var objNode = objXmlGeral.selectNodes("//xDados/Facilidade/Interligacao[Acf_Id="+intChave+"]")
	var strFacilidade =  new String("<table cellspacing=1 cellpadding=0 width=100% border=0>")
	//Refaz a lista de Ids no IFRAME
	for (var intIndex=0;intIndex<objNode.length;intIndex++){
		if ((intIndex%2)!=0){
			strFacilidade += "<tr class=clsSilver>"
		}
		else{
			strFacilidade += "<tr class=clsSilver2>"
		}	
		var intChave = objNode[intIndex].childNodes[0].text
		strFacilidade += "<td width=30 align=center><input type=radio name=rdoFacilidade value=" + intChave + " onClick='parent.EditarFac(this)'></td>"
		strFacilidade += "<td width=150px>" + objNode[intIndex].childNodes[2].text + "</td>"
		strFacilidade += "<td width=150px> " + objNode[intIndex].childNodes[1].text  + "</td>"
		strFacilidade += "</tr>"
	}	
	strFacilidade += "</table>"
	parent.IFrmFacilidade.spnFacilidade.innerHTML = strFacilidade
}
/*
=====================================================================================
Objetivo		: Retorna o valor de um node se não existir retorna ""
Premissas		: Enviar objeto XML 
Entradas		: objXml - objeto XML
				  strNomeNode - Nome do node
				  intChave - Indentificador do no	
Retorno			: Valor do node ou ""
Data - Autor	: 20/06/2002 - Vital
=====================================================================================
*/
function RequestNodeFac(objXML,strNomeNode,intChave)
{
	objNode = objXML.selectSingleNode("//xDados/Facilidade/Interligacao[Int_Id="+parseInt(intChave)+"]")
	var objNodeRequest = objNode.getElementsByTagName(strNomeNode)

	if (objNodeRequest.length != 0)
	{
		return objNodeRequest.item(0).text
		
	}	
	else
	{
		return ""		
	}
}

function RemoverFacLista(){
	with (document.forms[0]){
		if (hdnIntIndice.value != ""){
			var intRede = cboRede.value
			if (alertbox('Deseja permanecer com os dados?','Sim','Não')==1)
				RemoverFac(hdnIntIndice.value,false,intRede)
			else	
				RemoverFac(hdnIntIndice.value,true,intRede)	
		}
		else{
			alert("Selecione um item para remover em \"Facilidades do Pedido\".")
			return
		}
	}
}

function AbrirParam(obj)
{
	with (document.forms[0])
	{
		switch (obj.value)
		{
			case "1":
				spnDet.innerHTML = strDet
				spnNDet.innerHTML = ""
				spnAde.innerHTML = ""
				spnCampos.innerHTML = strCamposDet
				break
			case "2":
				spnDet.innerHTML = ""
				spnNDet.innerHTML = strNDet
				spnAde.innerHTML = ""
				spnCampos.innerHTML = strCamposNDet
				break
			case "3":
				spnDet.innerHTML = ""
				spnNDet.innerHTML = ""
				spnAde.innerHTML = strAde
				spnCampos.innerHTML = strCamposAde
				break
			default:	
				spnDet.innerHTML = ""
				spnNDet.innerHTML = ""
				spnAde.innerHTML = ""
				spnCampos.innerHTML = ""
		}
	}
}

function SelecionarRede(obj)
{
	with (document.forms[0]) 
	{
		if (obj.value == "7")
		{
			cboRede.value = 1
			if (cboProvedor.value != ""){ResgatarProvedoresAssociados(document.forms[0].cboProvedor)}
			spnDet.innerHTML = strDet
			spnNDet.innerHTML = ""
			spnAde.innerHTML = ""
			spnCampos.innerHTML = strCamposDet
		}
	}
}

function LimparFacSel()
{
	objAryFac[0] = new Array("","")
	objAryFac[1] = new Array("","")
	objAryFac[2] = new Array("","")
	objAryFac[3] = new Array("","")
	objAryFac[4] = new Array("","")
	//Limpa XML
	objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
	objXmlRoot = objXmlGeral.createNode("element","xDados","")
	objXmlGeral.appendChild (objXmlRoot)
	parent.IFrmFacilidade.spnFacilidade.innerHTML = ""
}

function PosicoesLivre(strPagina)
{
	with (document.forms[0])
	{
		if (!ValidarCampos(cboLocalInstala,"Local de Entrega")) return
		if (!ValidarCampos(cboDistLocalInstala,"Distribuidor")) return
		if (!ValidarCampos(cboRede,"Rede")) return
		if (!ValidarCampos(cboProvedor,"Provedor")) return

		switch (strPagina)
		{
			case "L": //livres
				objAryFacRet = window.showModalDialog('ConsultarFacilidades.asp?hdnAcao=Posicoes&strStsFac=L&cboLocalInstala='+cboLocalInstala.value+'&cboDistLocalInstala='+cboDistLocalInstala.value+'&cboProvedor='+cboProvedor.value+'&strPagina='+strPagina+'&cboSistema='+cboRede.value,objAryFac,"dialogHeight: 450px; dialogWidth: 550px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
				break
			case "O": //Ocupada
				objAryFacRet = window.showModalDialog('ConsultarFacilidades.asp?hdnAcao=Posicoes&strStsFac=O&cboLocalInstala='+cboLocalInstala.value+'&cboDistLocalInstala='+cboDistLocalInstala.value+'&cboProvedor='+cboProvedor.value+'&strPagina='+strPagina+'&cboSistema='+cboRede.value,objAryFac,"dialogHeight: 450px; dialogWidth: 550px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
				if (objAryFac[0][0] == "detalhar")
				{	
					hdnSolId.value = objAryFac[0][1]
					DetalharFac();
					return
				}
				break
		}

		try{
			for (var intIndex=0;intIndex<objAryFacRet.length;intIndex++)
			{
				switch (cboRede.value) 
				{
					case "1": //Det
						if (objAryFac[intIndex][0] != "")
						{
							txtFila.value		= objAryFac[intIndex][0]
							txtBastidor.value	= objAryFac[intIndex][1]
							txtRegua.value		= objAryFac[intIndex][2]
							txtPorta.value		= objAryFac[intIndex][3]
							txtTimeslot.value	= objAryFac[intIndex][4]
						}	
						break

					case "2": //NDet
						if (objAryFac[intIndex][0] != "")
						{
							txtTronco.value = objAryFac[intIndex][0]
							txtPar.value	= objAryFac[intIndex][1]
						}						
						break

					case "3": //ADE
						if (objAryFac[intIndex][0] != "")
						{
							txtCabo.value		= objAryFac[intIndex][0]
							txtPar.value		= objAryFac[intIndex][1]
							txtLateral.value	= objAryFac[intIndex][2]
							cboTipoCabo.value	= objAryFac[intIndex][3]
							txtCaixaEmenda.value= objAryFac[intIndex][4]
						}	
						break
				}	
			}
		}
		catch(e){}	
	}	
}

function VerificarRecurso(strPagina)
{
	with (document.forms[0])
	{
		if (!ValidarCampos(cboLocalConfig,"Local de Configuração")) return
		if (!ValidarCampos(cboRegimeCntr,"Prazos de Contratação de Acesso")) return
		if (!ValidarCampos(cboDistLocalInstala,"Distribuidor")) return
		if (!ValidarCampos(cboRede,"Rede")) return

		hdnAcao.value = "VerificarRecurso"
		target = "IFrmProcesso"
		action = 'ProcessoResgatarFac.asp?strEscConfig='+cboLocalConfig.value+'&strRegId='+cboRegimeCntr.value+'&strDistrib='+cboDistLocalInstala.value+'&strSis='+cboRede.value+'&strPagina='+strPagina
		submit()

	}

}

function ValidaProvedor(obj)
{
	if (IFrmProcesso1.document.body.innerHTML == "")
	{
		alert("Padrão do provedor não encontrado.\nPara Rede Deterministica/Tronco Par selecione\no código do provedor para ADE selecione o provedor.")
		return false
	}
	return true
}

//Remover daqui para baixo depois que a Dll estiver registrada
function ValidarNTipoPro()
{
    //Argumentos
    //ObjTextbox,Tipo,Tam,ProxTipo,Tam,ProxTipo,Tam,...
    //var checkOK = str;
    //var checkStr = Campo.value;
    var allValid = true;
    var decPoints = 0;
    var allNum = '';
    if (arguments.length > 0 )
    {
        var Campo = arguments[0]
        var strValue = arguments[0].value
        var intTamIni = 0
        var intTam = 0
        var intTipo = 0
        var intIndex = 1
        while (intIndex < arguments.length)
        {
            intTipo = arguments[intIndex]
            intTamIni = intTamIni + intTam
            intTam  = arguments[intIndex+1]
            checkStr = strValue.substring(intTamIni, intTamIni + intTam)
            if (!ValidarTipoPro(checkStr,intTipo))
            {
                intIndex = arguments.length
                return false;
            }
            else
            {
                intIndex = intIndex + 2
            }
        }
    }
    return true;
}

function ValidarTipoPro(Campo,intTipo)
{
 if (Campo == '[object]')
 {
     var checkStr = Campo.value;
 }
 else
 {
     var checkStr = Campo;
 }   
 var allValid = true;
 var decPoints = 0;
 var allNum = '';
 switch (intTipo)
 {
     case 0:
         var checkOK = '0123456789' //int,smallint,bit
         break
     case 1:
         var checkOK = 'QWERTYUIOPASDFGHJKLÇZXCVBNMqwertyuiopasdfghjklçzxcvbnm '
         break
     case 2:
         var checkOK = 'QWERTYUIOPASDFGHJKLÇZXCVBNMqwertyuiopasdfghjklçzxcvbnm01234546789 '
         break
     case 4:
         var checkOK = ' - ' //Traço do cep
         break
     case 5:
         var checkOK = ' ' //Em banco
         break
     case 6:
         var checkOK = '0123456789 ' //int,smallint,bit com espaço
         break
     default:
         var checkOK = intTipo //Recebe o próprio valor
         break
 }
 for (i = 0;  i < checkStr.length;  i++)
 {
     ch = checkStr.charAt(i);
     for (j = 0;  j < checkOK.length;  j++)
         if (ch == checkOK.charAt(j))
         break;
     if (j == checkOK.length)
     {
         allValid = false;
         break;
     }
     if (ch != ' ')
     allNum += ch;
 }
 if (!allValid)
 {
     return (false);
 }
 return (true);
}
