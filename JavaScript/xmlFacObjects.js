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
		target = "IFrmProcesso2"
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
					txtFila.value = ''
					txtBastidor.value = ''
					txtRegua.value = ''
					txtPosicao.value = ''
					txtTimeslot.value = ''
					txtCCTOPro.value = ''
					txtNroAcessoCli.value = ''
					txtCNLPtaA.value = ''
					txtCNLPtaB.value = ''
					txtQtdeModem.value = ''
					cboPropModem.value = ''
					txtObsFac.value = ''	
					for (var intIndex=0;intIndex<rdoUrbano.length;intIndex++){
						rdoUrbano[intIndex].checked = false
					}	
					try{
					for (var intIndex=0;intIndex<rdoFatura.length;intIndex++){
						rdoFatura[intIndex].checked = false
					}}catch(e){}
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
					txtCCTOPro.value = ''
					txtNroAcessoCli.value = ''
					txtCNLPtaA.value = ''
					txtCNLPtaB.value = ''
					txtQtdeModem.value = ''
					cboPropModem.value = ''
					txtObsFac.value = ''
					//eval("spnStsNDet.innerHTML = '&nbsp;'")
					for (var intIndex=0;intIndex<rdoUrbano.length;intIndex++){
						rdoUrbano[intIndex].checked = false
					}
					try{
					for (var intIndex=0;intIndex<rdoFatura.length;intIndex++){
						rdoFatura[intIndex].checked = false
					}}catch(e){}
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
					txtQtdeModem.value = ''
					cboPropModem.value = ''
					txtObsFac.value = ''
					try{
					for (var intIndex=0;intIndex<rdoFatura.length;intIndex++){
						rdoFatura[intIndex].checked = false
					}}catch(e){}
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
		if (elemento.type == 'radio' || elemento.type == 'checkbox')
		{
			elemento.checked = false
		}
	}
}

function ApagarNroAcesso()
{
	with (document.forms[0]) 
	{
		if (cboRede.value != "")
		{
			txtNroAcessoEbt.value = ''
		}
	}	
}


function ResgatarPadraoProvedor(obj,intProd)
{
	with (document.forms[0])
	{
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
				hdnIdAcessoFisico.value = "0"
				target = "IFrmIDFis1"
				action = "Estoque.asp?intEnd=1"
				submit()
				break
			case 2:	
				hdnIdAcessoFisico.value = "0"
				target = "IFrmIDFis1"
				action = "AcessoCompartilhadoSol.asp?intEnd=1"
				submit()
				break
		}
	}	
}

function SelIDFisComp(obj,intEnd)
{
	with (document.forms[0])
	{
		var objNodeE = objXmlGeral.selectNodes("//Facilidade[strTipoConsulta='E' || strTipoConsulta='C']")
		var objNode = objXmlGeral.selectNodes("//Facilidade")

		if (objNode.length > 0 )
		{
			if (objNodeE.length == 0)
			{
				alert("Para compartilhar um acesso físico ou selecioná-lo de estoque, favor remover as facilidades atualmente alocadas/adicionadas.")
				limparIDFisico(1)
				return
			}	
		}

		switch (intEnd)
		{
			case 1: //id's Físicos 
				hdnIdAcessoFisico.value=obj.value
				hdnPropIdFisico.value = obj.prop
				hdnCompartilhamento.value = "0"
				hdnAcao.value = "ResgatarIdFisicoComp"
				target = "IFrmProcesso"
				action = "ProcessoResgatarFac.asp"
				submit()
				break
			case 2:	//Estoque
				hdnIdAcessoFisico1.value=obj.value
				hdnPropIdFisico1.value = obj.prop
				hdnCompartilhamento1.value = 1
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
					hdnRede.value = cboRede.value
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
				txtCNLPtaA.value = txtCNLLocalConfig.value
				txtCNLPtaB.value = txtCNLLocalConfig.value
			}	
		}
		else
		{
			txtCNLPtaA.value = txtCNLLocalConfig.value
			txtCNLPtaB.value = ''
		}
	}	
}

function AdicionarCNLOtimiz(obj)
{
	with (document.forms[0])
	{
		if (rdoUrbano[1].checked)
		{
			if (cboLocalConfig.value != "")
			{
				txtCNLPtaA.value = txtCNLLocalConfig.value
				txtCNLPtaB.value = txtCNLLocalConfig.value
			}	
		}
		else
		{
			txtCNLPtaA.value = txtCNLLocalConfig.value
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
						AdicionarNode(objXmlGeral,"hdnIdAcessoFisico",hdnIdAcessoFisico.value)
						AdicionarNode(objXmlGeral,"hdnCompartilhamento",hdnCompartilhamento.value)
						AdicionarNode(objXmlGeral,"hdnIdAcessoFisico1",hdnIdAcessoFisico1.value)
						AdicionarNode(objXmlGeral,"hdnCompartilhamento1",hdnCompartilhamento1.value)
						var objNode = objXmlGeral.selectNodes("//Facilidade")
						if (objNode.length > 0){
							hdnAcao.value = "GravarFacilidade"
							hdnRede.value = cboRede.value
							hdnXml.value = objXmlGeral.xml 
							target = "IFrmProcesso"  
							action = "ProcessoFac.asp"
							method = "Post"
							submit()
						}else{
							alert("Favor informar as facilidades do pedido ou adicionar facilidades a lista.")
							return
						}	
						break
					
					case 2: // Não Aceito
						hdnCompartilhamento.value = "0"
						AdicionarNode(objXmlGeral,"hdnIdAcessoFisico",hdnIdAcessoFisico.value)
						AdicionarNode(objXmlGeral,"hdnCompartilhamento",hdnCompartilhamento.value)
						AdicionarNode(objXmlGeral,"hdnIdAcessoFisico1",hdnIdAcessoFisico1.value)
						AdicionarNode(objXmlGeral,"hdnCompartilhamento1",hdnCompartilhamento1.value)
						return
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
			if (elemento.type != 'button' && elemento.type != 'hidden' && elemento.type != 'radio' && elemento.type != 'checkbox' && elemento.name != 'cboRede') elemento.value = ""
			if (elemento.type == 'radio')
			{
				elemento.checked = false
			}
		}	
	}
	if (document.forms[0].hdnIntIndice.value != ""){
		var strParTimeSlot = new String(RequestNodeFac(objXmlGeral,"txtPar",document.forms[0].hdnIntIndice.value))
		if (strParTimeSlot == ""){
			strParTimeSlot = RequestNodeFac(objXmlGeral,"txtTimeslot",document.forms[0].hdnIntIndice.value)
		}
		if (strParTimeSlot.length == 4){strParTimeSlot = "0000-" + strParTimeSlot}
		AdicionarNodeFac(objXmlGeral,'OrderBy',strParTimeSlot,document.forms[0].hdnIntIndice.value)

		if (RequestNodeFac(objXmlGeral,"Fac_Id",document.forms[0].hdnIntIndice.value) != ""){
			AdicionarNodeFac(objXmlGeral,'TipoAcao','A',document.forms[0].hdnIntIndice.value)
		}else{
			AdicionarNodeFac(objXmlGeral,'TipoAcao','N',document.forms[0].hdnIntIndice.value)
		}	
	}
	else{	
		var strParTimeSlot = new String(RequestNodeFac(objXmlGeral,"txtPar",intIndice))
		if (strParTimeSlot == ""){
			strParTimeSlot = RequestNodeFac(objXmlGeral,"txtTimeslot",intIndice)
		}
		if (strParTimeSlot.length == 4){strParTimeSlot = "0000-" + strParTimeSlot}
		AdicionarNodeFac(objXmlGeral,'OrderBy',strParTimeSlot,intIndice)
		AdicionarNodeFac(objXmlGeral,'TipoAcao','N',intIndice)
	}
	intIndice += 1
	document.forms[0].hdnIntIndice.value = ""
	AtualizarListaFac()
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
	objNodeList = objXML.selectNodes("//xDados/Facilidade[intIndice="+parseInt(intChave)+"]")
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
		objNode = objXML.selectSingleNode("//xDados/Facilidade[intIndice="+parseInt(intChave)+"]")
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
	var	objNode = objXmlGeral.selectNodes("//xDados/Facilidade[intIndice="+parseInt(intChave)+"]")
	//Refaz a lista de Ids no IFRAME
	for (var intIndex=0;intIndex<objNode.length;intIndex++){
		for (var intIndexII=0;intIndexII<objNode[intIndex].childNodes.length;intIndexII++){
			try{
				var strNomdeName = new String(objNode[intIndex].childNodes[intIndexII].nodeName)
				var objChildForm = new Object(eval("document.forms[0]."+strNomdeName))
				if (strNomdeName.indexOf("rdo") != -1 || strNomdeName.indexOf("chk") != -1){
					var intIndexSelected = RequestNodeFac(objXmlGeral,strNomdeName+"Index",objNode[intIndex].childNodes[0].text)
					eval("document.forms[0]."+strNomdeName+"["+parseInt(intIndexSelected)+"].checked = true")
				}
				else{
					//eval("document.forms[0]."+objNode[intIndex].childNodes[intIndexII].nodeName+".value='"+objNode[intIndex].childNodes[intIndexII].text+"'")
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
		var	objNode = objXmlGeral.selectNodes("//xDados/Facilidade[intIndice="+parseInt(intIndice)+"]")
		if (objNode.length > 0){
			objNode[0].parentNode.removeChild(objNode[0])
		}
		else
		{
			alert("Facilidade não encontrada.")
			return
		}	

	}

	AtualizarListaFac()
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
function AtualizarListaFac(){

	var objNode = objXmlGeral.selectNodes("//xDados/Facilidade[TipoAcao='N' || TipoAcao='A']")
	var strFacilidade =  new String("<table cellspacing=1 cellpadding=0 width=740 border=0>")
	//Refaz a lista de Ids no IFRAME
	for (var intIndex=0;intIndex<objNode.length;intIndex++){
		if ((intIndex%2)!=0){
			strFacilidade += "<tr class=clsSilver>"
		}
		else{
			strFacilidade += "<tr class=clsSilver2>"
		}	
		var intChave = objNode[intIndex].childNodes[0].text
		var intRede = RequestNodeFac(objXmlGeral,"cboRede",intChave)
		
		switch (parseInt(intRede))
		{
			case 1:
				//objAryObjs[2] = new Array("txtNroAcessoEbt","txtTronco","txtPar","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB","rdoUrbano[0]","rdoUrbano[1]","cboRede")
				//ADE
				//objAryObjs[3] = new Array("txtNroAcessoEbt","txtCabo","txtPar","txtLateral","cboTipoCabo","txtCaixaEmenda","cboRede")
				strFacilidade += "<td width=30 ><input type=radio name=rdoFacilidade value=" + intChave + " onClick='parent.EditarFac(this)'></td>"
				strFacilidade += "<td width=100>" + RequestNodeFac(objXmlGeral,"cboCodProvText",intChave)  + "</td>"
				strFacilidade += "<td width=100>" + RequestNodeFac(objXmlGeral,"txtNroAcessoEbt",intChave)  + "</td>"
				strFacilidade += "<td width=80>" + RequestNodeFac(objXmlGeral,"txtFila",intChave)  + "</td>"
				strFacilidade += "<td width=80>" + RequestNodeFac(objXmlGeral,"txtBastidor",intChave)  + "</td>"
				strFacilidade += "<td width=80>" + RequestNodeFac(objXmlGeral,"txtRegua",intChave)  + "</td>"
				strFacilidade += "<td width=80>" + RequestNodeFac(objXmlGeral,"txtPosicao",intChave)  + "</td>"
				strFacilidade += "<td width=80>" + RequestNodeFac(objXmlGeral,"txtTimeslot",intChave)  + "</td>"
				strFacilidade += "<td width=110>" + RequestNodeFac(objXmlGeral,"txtNroAcessoCli",intChave)  + "</td>"
				strFacilidade += "</tr>"
				//RetornaCboPlataforma(document.forms[0].cboRede.value, RequestNodeFac(objXmlGeral,"cboPlataforma",intChave) )
				break
			case 2:
				strFacilidade += "<td width=30 align=center><input type=radio name=rdoFacilidade value=" + intChave + " onClick='parent.EditarFac(this)'></td>"
				strFacilidade += "<td width=100>" + RequestNodeFac(objXmlGeral,"cboCodProvText",intChave)  + "</td>"
				strFacilidade += "<td width=100>" + RequestNodeFac(objXmlGeral,"txtNroAcessoEbt",intChave)  + "</td>"
				strFacilidade += "<td width=100>" + RequestNodeFac(objXmlGeral,"txtTronco",intChave)  + "</td>"
				strFacilidade += "<td width=100>" + RequestNodeFac(objXmlGeral,"txtPar",intChave)  + "</td>"
				strFacilidade += "<td width=310>" + RequestNodeFac(objXmlGeral,"txtNroAcessoCli",intChave)  + "</td>"
				strFacilidade += "</tr>"
				break
			case 3:
				strFacilidade += "<td width=30 align=center><input type=radio name=rdoFacilidade value=" + intChave + " onClick='parent.EditarFac(this)'></td>"
				strFacilidade += "<td width=100>" + RequestNodeFac(objXmlGeral,"txtNroAcessoEbt",intChave)  + "</td>"
				strFacilidade += "<td width=100>" + RequestNodeFac(objXmlGeral,"txtCabo",intChave)  + "</td>"
				strFacilidade += "<td width=100>" + RequestNodeFac(objXmlGeral,"txtPar",intChave)  + "</td>"
				strFacilidade += "<td width=100>" + RequestNodeFac(objXmlGeral,"txtCaixaEmenda",intChave)  + "</td>"
				strFacilidade += "<td width=97>" + RequestNodeFac(objXmlGeral,"txtLateral",intChave)  + "</td>"
				strFacilidade += "<td width=85>" + RequestNodeFac(objXmlGeral,"cboTipoCabo",intChave)  + "</td>"
				strFacilidade += "</tr>"
				break
			case 4:
				strFacilidade += "<td width=30 align=center><input type=radio name=rdoFacilidade value=" + intChave + " onClick='parent.EditarFac(this)'></td>"
				//strFacilidade += "<td width=200>" + RequestNodeFac(objXmlGeral,"txtDesignacao",intChave)  + "</td>"
				strFacilidade += "<td width=200>" + RequestNodeFac(objXmlGeral,"txtNroAcessoEbt",intChave)  + "</td>"
				strFacilidade += "<td width=510>" + RequestNodeFac(objXmlGeral,"txtVlan",intChave)  + "</td>"
				strFacilidade += "</tr>"
				break
			case 5:
				strFacilidade += "<td width=30 align=center><input type=radio name=rdoFacilidade value=" + intChave + " onClick='parent.EditarFac(this)'></td>"
				strFacilidade += "<td width=200>" + RequestNodeFac(objXmlGeral,"txtNroAcessoEbt",intChave)  + "</td>"
				strFacilidade += "<td width=80>" + RequestNodeFac(objXmlGeral,"txtVlan",intChave)  + "</td>"
				strFacilidade += "<td width=200>" + RequestNodeFac(objXmlGeral,"txtPE",intChave)  + "</td>"
				strFacilidade += "<td width=140>" + RequestNodeFac(objXmlGeral,"txtPorta",intChave)  + "</td>"
				strFacilidade += "<td width=90>" + RequestNodeFac(objXmlGeral,"txtLink",intChave)  + "</td>"
				strFacilidade += "</tr>"
				break
			case 13:
				//alert('estou aqui!')
				strFacilidade += "<td width=30 align=center><input type=radio name=rdoFacilidade value=" + intChave + " onClick='parent.EditarFac(this)'></td>"
				strFacilidade += "<td width=170>" + RequestNodeFac(objXmlGeral,"txtNroAcessoEbt",intChave)  + "</td>"
				strFacilidade += "<td width=80>" + RequestNodeFac(objXmlGeral,"txtVlan",intChave)  + "</td>"
				strFacilidade += "<td width=150>" + RequestNodeFac(objXmlGeral,"txtPE",intChave)  + "</td>"
				strFacilidade += "<td width=140>" + RequestNodeFac(objXmlGeral,"txtPorta",intChave)  + "</td>"
				strFacilidade += "<td width=90>" + RequestNodeFac(objXmlGeral,"txtSvlan",intChave)  + "</td>"
				strFacilidade += "<td width=80>" + RequestNodeFac(objXmlGeral,"cboTipoPortaText",intChave)  + "</td>"
				strFacilidade += "</tr>"
				break				
			case 10:
				strFacilidade += "<td width=30 align=center><input type=radio name=rdoFacilidade value=" + intChave + " onClick='parent.EditarFac(this)'></td>"
				//strFacilidade += "<td width=200>" + RequestNodeFac(objXmlGeral,"txtDesignacao",intChave)  + "</td>"
				strFacilidade += "<td width=200>" + RequestNodeFac(objXmlGeral,"txtNroAcessoEbt",intChave)  + "</td>"
				strFacilidade += "<td width=150>" + RequestNodeFac(objXmlGeral,"txtIP",intChave)  + "</td>"
				strFacilidade += "<td width=150>" + RequestNodeFac(objXmlGeral,"txtGateway",intChave)  + "</td>"
				strFacilidade += "<td width=190>" + RequestNodeFac(objXmlGeral,"txtMascara",intChave)  + "</td>"
				strFacilidade += "</tr>"
				break
			case 11:
				strFacilidade += "<td width=30 align=center><input type=radio name=rdoFacilidade value=" + intChave + " onClick='parent.EditarFac(this)'></td>"
				//strFacilidade += "<td width=200>" + RequestNodeFac(objXmlGeral,"txtDesignacao",intChave)  + "</td>"
				strFacilidade += "<td width=200>" + RequestNodeFac(objXmlGeral,"txtNroAcessoEbt",intChave)  + "</td>"
				strFacilidade += "<td width=80>" + RequestNodeFac(objXmlGeral,"txtVlan_FO",intChave)  + "</td>"
				strFacilidade += "<td width=200>" + RequestNodeFac(objXmlGeral,"txtPE_FO",intChave)  + "</td>"
				strFacilidade += "<td width=80>" + RequestNodeFac(objXmlGeral,"txtPorta_FO",intChave)  + "</td>"
				strFacilidade += "<td width=150>" + RequestNodeFac(objXmlGeral,"txtSvlan_FO",intChave)  + "</td>"
				strFacilidade += "</tr>"
				break
				
		}		
	}	
	strFacilidade += "</table>"
	//alert('strFacilidade='+strFacilidade)
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
	objNode = objXML.selectSingleNode("//xDados/Facilidade[intIndice="+parseInt(intChave)+"]")
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
		if (hdnTipoProcesso.value == 4)
		{
			alert("Para o processo de cancelamento não é possível efetuar alterações nas facilidades do pedido.")
			return	
		}
		if (!ValidarAlteracaoFacilidade()) return
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
				spnBsod.innerHTML = ""
				spnNDet.innerHTML = ""
				spnAde.innerHTML = ""
				spnBsodNet.innerHTML = ""
				spnBsodVia.innerHTML = ""
				spnBsodLight.innerHTML = ""
				spnFoEtherNet.innerHTML = ""
				spnCampos.innerHTML = strCamposDet
				spnCamposEthernetInterconexao.innerHTML = ""
				spnCamposEthernetMetro.innerHTML = ""
				spnCamposEthernetEDD.innerHTML = ""
				spnCamposEthernetPE.innerHTML = ""
				spnCamposEthernetVlan.innerHTML = ""
				break
			case "2":
				spnDet.innerHTML = ""
				spnBsod.innerHTML = ""
				spnNDet.innerHTML = strNDet
				spnAde.innerHTML = ""
				spnBsodNet.innerHTML = ""
				spnBsodVia.innerHTML = ""
				spnBsodLight.innerHTML = ""
				spnFoEtherNet.innerHTML = ""
				spnCampos.innerHTML = strCamposNDet
				spnCamposEthernetInterconexao.innerHTML = ""
				spnCamposEthernetMetro.innerHTML = ""
				spnCamposEthernetEDD.innerHTML = ""
				spnCamposEthernetPE.innerHTML = ""
				spnCamposEthernetVlan.innerHTML = ""
				break
			case "3":
				spnDet.innerHTML = ""
				spnBsod.innerHTML = ""
				spnNDet.innerHTML = ""
				spnAde.innerHTML = strAde
				spnBsodNet.innerHTML = ""
				spnBsodVia.innerHTML = ""
				spnBsodLight.innerHTML = ""
				spnFoEtherNet.innerHTML = ""
				spnCampos.innerHTML = strCamposAde
				spnCamposEthernetInterconexao.innerHTML = ""
				spnCamposEthernetMetro.innerHTML = ""
				spnCamposEthernetEDD.innerHTML = ""
				spnCamposEthernetPE.innerHTML = ""
				spnCamposEthernetVlan.innerHTML = ""
				break
			case "4":
				spnDet.innerHTML = ""
				spnBsod.innerHTML = ""
				spnNDet.innerHTML = ""
				spnAde.innerHTML = ""
				spnBsodNet.innerHTML = strBsodNet
				spnBsodVia.innerHTML = ""
				spnBsodLight.innerHTML = ""
				spnFoEtherNet.innerHTML = ""
				spnCampos.innerHTML = strCamposBsodNet
				spnCamposEthernetInterconexao.innerHTML = ""
				spnCamposEthernetMetro.innerHTML = ""
				spnCamposEthernetEDD.innerHTML = ""
				spnCamposEthernetPE.innerHTML = ""
				spnCamposEthernetVlan.innerHTML = ""
				break
			case "5":
				spnDet.innerHTML = ""
				spnBsod.innerHTML = ""
				spnNDet.innerHTML = ""
				spnAde.innerHTML = ""
				spnBsodNet.innerHTML = ""
				spnBsodVia.innerHTML = strBsodVia
				spnBsodLight.innerHTML = ""
				spnFoEtherNet.innerHTML = ""
				spnCampos.innerHTML = strCamposBsodVia
				spnCamposEthernetInterconexao.innerHTML = ""
				spnCamposEthernetMetro.innerHTML = ""
				spnCamposEthernetEDD.innerHTML = ""
				spnCamposEthernetPE.innerHTML = ""
				spnCamposEthernetVlan.innerHTML = ""
				break
				case "13":
				spnDet.innerHTML = ""
				spnNDet.innerHTML = ""
				spnAde.innerHTML = ""
				spnBsodNet.innerHTML = ""
				spnBsodVia.innerHTML = ""
				spnBsod.innerHTML = strBsod
				spnBsodLight.innerHTML = ""
				spnFoEtherNet.innerHTML = ""
				spnCampos.innerHTML = strCamposBsod
				spnCamposEthernetInterconexao.innerHTML = ""
				spnCamposEthernetMetro.innerHTML = ""
				spnCamposEthernetEDD.innerHTML = ""
				spnCamposEthernetPE.innerHTML = ""
				spnCamposEthernetVlan.innerHTML = ""
				break			
			case "10":
				spnDet.innerHTML = ""
				spnBsod.innerHTML = ""
				spnNDet.innerHTML = ""
				spnAde.innerHTML = ""
				spnBsodNet.innerHTML = ""
				spnBsodVia.innerHTML = ""
				spnFoEtherNet.innerHTML = ""
				spnBsodLight.innerHTML = strBsodLight
				spnCampos.innerHTML = strCamposBsodLight
				spnCamposEthernetInterconexao.innerHTML = ""
				spnCamposEthernetMetro.innerHTML = ""
				spnCamposEthernetEDD.innerHTML = ""
				spnCamposEthernetPE.innerHTML = ""
				spnCamposEthernetVlan.innerHTML = ""
				break
				
			case "11":
				spnDet.innerHTML = ""
				spnBsod.innerHTML = ""
				spnNDet.innerHTML = ""
				spnAde.innerHTML = ""
				spnBsodNet.innerHTML = ""
				spnBsodVia.innerHTML = ""
				spnFoEtherNet.innerHTML = strFoEtherNet
				spnBsodLight.innerHTML = ""
				spnCampos.innerHTML = strCamposFoEtherNet
				spnCamposEthernetInterconexao.innerHTML = ""
				spnCamposEthernetMetro.innerHTML = ""
				spnCamposEthernetEDD.innerHTML = ""
				spnCamposEthernetPE.innerHTML = ""
				spnCamposEthernetVlan.innerHTML = ""
				break
			
			case "12":
				spnDet.innerHTML = ""
				spnBsod.innerHTML = ""
				spnNDet.innerHTML = ""
				spnAde.innerHTML = ""
				spnBsodNet.innerHTML = ""
				spnBsodVia.innerHTML = ""
				spnFoEtherNet.innerHTML = ""
				spnBsodLight.innerHTML = ""
				spnCampos.innerHTML = strCamposBsodRadioIP
				spnCamposEthernetInterconexao.innerHTML = ""
				spnCamposEthernetMetro.innerHTML = ""
				spnCamposEthernetEDD.innerHTML = ""
				spnCamposEthernetPE.innerHTML = ""
				spnCamposEthernetVlan.innerHTML = ""
				break
			
			case "14":
				spnDet.innerHTML = ""
				spnBsod.innerHTML = ""
				spnNDet.innerHTML = ""
				spnAde.innerHTML = ""
				spnBsodNet.innerHTML = ""
				spnBsodVia.innerHTML = ""
				spnFoEtherNet.innerHTML = ""
				spnBsodLight.innerHTML = ""
				spnCampos.innerHTML = ""
				spnCamposEthernetInterconexao.innerHTML = strCamposEthernetInterconexao
				spnCamposEthernetMetro.innerHTML = strCamposEthernetMetro
				spnCamposEthernetEDD.innerHTML = strCamposEthernetEDD
				spnCamposEthernetPE.innerHTML = strCamposEthernetPE
				spnCamposEthernetVlan.innerHTML = strCamposEthernetVlan
				break
			
			default:	
				spnDet.innerHTML = ""
				spnBsod.innerHTML = ""
				spnNDet.innerHTML = ""
				spnAde.innerHTML = ""
				spnBsodNet.innerHTML = ""
				spnBsodVia.innerHTML = ""
				spnBsodLight.innerHTML = ""
				spnCampos.innerHTML = ""
				spnFoEtherNet.innerHTML = ""
				spnCamposEthernetInterconexao.innerHTML = ""
				spnCamposEthernetMetro.innerHTML = ""
				spnCamposEthernetEDD.innerHTML = ""
				spnCamposEthernetPE.innerHTML = ""
				spnCamposEthernetVlan.innerHTML = ""
				
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

function HabilitaDesignacao(obj , campo)
{
	//alert(obj.value);
	//alert(obj.name);
	
	campo = obj.name;
	with (document.forms[0]) 
	{
		
		if (obj.value == "2")
		{
			if (obj.name == "cboPropriedadeInter")
			{
				cboSwitchInterconexao.disabled = false;
				
			}
			if (obj.name == "cboPropriedadeEDD")
			{
				cboSwitchEDD.disabled = false;
				
			}
			if (obj.name == "cboPropriedadeMetro")
			{
				cboSwitchMetro.disabled = false;
				
			}
						
			campo = "";
		}
		if (obj.value == "1" || obj.value == "3")
		{
			if (obj.name == "cboPropriedadeInter")
			{
				cboSwitchInterconexao.disabled = true;
				
			}
			if (obj.name == "cboPropriedadeEDD")
			{
				cboSwitchEDD.disabled = true;
				
			}
			if (obj.name == "cboPropriedadeMetro")
			{
				cboSwitchMetro.disabled = true;
				
			}
			
			
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
	var TipoPlataforma
	with (document.forms[0])
	{
		if (!ValidarCampos(cboLocalInstala,"Local de Entrega")) return
		if (!ValidarCampos(cboDistLocalInstala,"Distribuidor")) return
		if (!ValidarCampos(cboRede,"Rede")) return
		if (!ValidarCampos(cboProvedor,"Provedor")) return
		
		switch (strPagina)
		{
			case "L": //livres
				try{
					objAryFacRet = window.showModalDialog('ConsultarFacilidades.asp?hdnAcao=Posicoes&strStsFac=L&cboLocalInstala='+cboLocalInstala.value+'&cboDistLocalInstala='+cboDistLocalInstala.value+'&cboProvedor='+cboProvedor.value+'&strPagina='+strPagina+'&cboSistema='+cboRede.value+'&cboPlataforma='+cboPlataforma.value ,objAryFac,"dialogHeight: 450px; dialogWidth: 570px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
				}
				catch(e){
					objAryFacRet = window.showModalDialog('ConsultarFacilidades.asp?hdnAcao=Posicoes&strStsFac=L&cboLocalInstala='+cboLocalInstala.value+'&cboDistLocalInstala='+cboDistLocalInstala.value+'&cboProvedor='+cboProvedor.value+'&strPagina='+strPagina+'&cboSistema='+cboRede.value, objAryFac,"dialogHeight: 450px; dialogWidth: 570px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
				}
				break
			case "O": //Ocupada
				try{  
					objAryFacRet = window.showModalDialog('ConsultarFacilidades.asp?hdnAcao=Posicoes&strStsFac=O&cboLocalInstala='+cboLocalInstala.value+'&cboDistLocalInstala='+cboDistLocalInstala.value+'&cboProvedor='+cboProvedor.value+'&strPagina='+strPagina+'&cboSistema='+cboRede.value+'&cboPlataforma='+cboPlataforma.value ,objAryFac,"dialogHeight: 450px; dialogWidth: 570px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
				}
				catch(e){
					objAryFacRet = window.showModalDialog('ConsultarFacilidades.asp?hdnAcao=Posicoes&strStsFac=O&cboLocalInstala='+cboLocalInstala.value+'&cboDistLocalInstala='+cboDistLocalInstala.value+'&cboProvedor='+cboProvedor.value+'&strPagina='+strPagina+'&cboSistema='+cboRede.value,objAryFac,"dialogHeight: 450px; dialogWidth: 570px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
				}
				
				if (objAryFac[0][0] == "detalhar")
				{	
					hdnSolId.value = objAryFac[0][1]
					try{
						PopularXml(objXmlGeralReturn)
					}catch(e){}
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
							txtPosicao.value	= objAryFac[intIndex][3]
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

function ListarInterligacoes(intChave){

	var objNode = objXmlGeral.selectNodes("//xDados/Facilidade/Interligacao[Ped_Id="+intChave+"]")
	var strFacilidade =  new String("<table cellspacing=1 cellpadding=1 width=100% border=0>")
	//Refaz a lista de Ids no IFRAME
	for (var intIndex=0;intIndex<objNode.length;intIndex++){
		if ((intIndex%2)!=0){
			strFacilidade += "<tr class=clsSilver>"
		}
		else{
			strFacilidade += "<tr class=clsSilver2>"
		}	
		var intChave = objNode[intIndex].childNodes[0].text
		strFacilidade += "<td width=185px>" + objNode[intIndex].childNodes[7].text + "</td>"
		strFacilidade += "<td width=170px>" + objNode[intIndex].childNodes[1].text + "</td>"
		strFacilidade += "<td width=150px> " + objNode[intIndex].childNodes[3].text  + "</td>"
		strFacilidade += "</tr>"
	}	
	strFacilidade += "</table>"
	parent.IFrmFacilidade.spnFacilidade.innerHTML = strFacilidade
}

function ValidarAlteracaoFacilidade()
{
	with (document.forms[0])
	{
		if (hdnPodeAlterar.value == "N")
		{
			alert("Favor efetivar a alocação das facilidades de estoque ou compartilhadas para poder alterá-las. Pedido com acessos compartilhados serão enviados para entrega do acesso ao serviço.")
			return false
		}
	}	
	return true
}

function AddFac(){
	with (document.forms[0])
	{
//alert(cboTipoPorta.value)
		
		if (hdnTipoProcesso.value == 4)
		{
			alert("Para o processo de cancelamento não é possível efetuar alterações nas facilidades do pedido.")
			return	
		}

		if (!ValidarAlteracaoFacilidade()) return
		if (!ValidarCampos(cboRede,"Rede")) return
		//	SO-EMPLN-77275WBC - Correção para poder validar os campos da facilidade
		//
		// Edar
		//
		//if (cboRede.value != "" && cboRede.disabled == false)
		if (cboRede.value != "")
		{
			switch (parseInt(cboRede.value)) 
			{
				case 1: //Det
					if (!ValidarCampos(cboCodProv,"Código do provedor")) return false
					if (!ValidarCampos(txtBastidor,"Bastidor")) return false
					if (!ValidarCampos(txtRegua,"Posição")) return false
					if (!ValidarCampos(txtPosicao,"Porta")) return false
					if (!ValidarCampos(txtTimeslot,"Timeslot")) return false
					if (rdoUrbano[1].checked)
					{
						txtCNLPtaA.value = txtCNLLocalConfig.value
						txtCNLPtaB.value = txtCNLLocalConfig.value
					}
					else
					{
					    if (IsEmpty(txtCNLPtaA.value)) //Adicionado PRSS 22/02/2007
						{
						  txtCNLPtaA.value = txtCNLLocalConfig.value
						}
					}	
					break

				case 2: //NDet
					if (!ValidarCampos(cboCodProv,"Código do provedor")) return false
					if (!ValidarCampos(txtTronco,"Tronco")) return false
					if (!ValidarCampos(txtPar,"Par")) return false
					if (rdoUrbano[1].checked)
					{
						txtCNLPtaA.value = txtCNLLocalConfig.value
						txtCNLPtaB.value = txtCNLLocalConfig.value
					}
					else
					{
					    if (IsEmpty(txtCNLPtaA.value)) //Adicionado PRSS 22/02/2007
						{
						  txtCNLPtaA.value = txtCNLLocalConfig.value
						}
					}	
					break

				case 3: //ADE
					if (!ValidarCampos(txtCabo,"Cabo")) return false
					if (!ValidarCampos(txtPar,"Número do Cabo de Acesso")) return false
					break
					
				case 4: //BsodNet
					//if (!ValidarCampos(txtDesignacao,"Designação do Acesso")) return false
					if (!ValidarCampos(txtVlan,"VLAN")) return false
					break
					
				case 5: //BsodVia
					//if (!ValidarCampos(txtDesignacao,"Designação do Acesso")) return false
					if (!ValidarCampos(txtVlan,"VLAN")) return false
					if (!ValidarCampos(txtPE,"PE")) return false
					if (!ValidarCampos(txtPorta,"Porta")) return false
					// retirado na UAT CLA SAGP BSOD 24052011
					//if (!ValidarCampos(txtLink,"Link")) return false
				    break
					
				case 13: //HFC Bsod
					//if (!ValidarCampos(txtDesignacao,"Designação do Acesso")) return false
					if (!ValidarCampos(txtNroAcessoEbt,"Designação do Acesso")) return false
					if (!ValidarCampos(txtVlan,"CVLAN")) return false
					if (!ValidarCampos(txtPE,"PE")) return false
					if (!ValidarCampos(txtPorta,"Porta")) return false
					if (!ValidarCampos(cboTipoPorta,"TIPO PORTA")) return false
					// retirado na UAT CLA SAGP BSOD 24052011
					//if (!ValidarCampos(txtLink,"Link")) return false
				    break
				    					
				case 10: 
				
					if (!ValidarCampos(txtNroAcessoEbt,"Designação do Acesso")) return false
					if (!ValidarCampos(txtIP,"IP")) return false
					if (!ValidarCampos(txtGateway,"Gateway")) return false
					if (!ValidarCampos(txtMascara,"Máscara")) return false
					break	
									
				case 11: 
				
					if (!ValidarCampos(txtNroAcessoEbt,"Designação do Acesso")) return false
					if (!ValidarCampos(txtVlan_FO,"VLAN")) return false
					if (!ValidarCampos(txtPE_FO,"PE")) return false
					if (!ValidarCampos(txtPorta_FO,"Porta")) return false
					if (!ValidarCampos(txtSvlan_FO,"SVLAN")) return false
					
				    break

			}	
			if (!IsEmpty(txtNroAcessoEbt.value) )
			{
				if (!ValidarPadraoProvedor()) return false
			}	

			if ( cboRede.value == 3 || cboRede.value == 2 || cboRede.value == 1 )
			{
				
				if (!ValidarCampos(cboPropModem,"Proprietário do Modem")) return false
				if (!ValidarCampos(txtQtdeModem,"Quantidade de Modens")) return false
				try{
					if (!rdoFatura[0].checked && !rdoFatura[1].checked){
						alert("Fatura é um Campo Obrigatório.")
						return false
					}
				}catch(e){}	
			}
		}

		

		var blnMessage = false
		if (arguments.length>0){
			blnMessage = arguments[0]
			intRet = 1
		}
		if (!blnMessage){
			var intRet=alertbox('Deseja permanecer com os dados?','Sim','Não','Sair')
		}	
//alert('cboRede.value='+cboRede.value)
		switch (parseInt(intRet))
		{
			case 1:
				xmlFacUpd(false,cboRede.value)
				break
			case 2:	
				xmlFacUpd(true,cboRede.value)
				break
		}
		return true
	}	
}

function ValidarPadraoProvedor()
{
	with (document.forms[0])
	{
		var padraoMin = parent.IFrmProcesso1.document.forms[0].hdnPadraoMin.value;
		var padraoMax = parent.IFrmProcesso1.document.forms[0].hdnPadraoMax.value;

		//Valida o tamanho do padrão mínimo/máximo permitido ao provedor
		if (txtNroAcessoEbt.value.length != parent.IFrmProcesso1.document.forms[0].hdnIntPadraoMin.value && txtNroAcessoEbt.value.length != parent.IFrmProcesso1.document.forms[0].hdnIntPadraoMax.value)
		{
			alert("Número do Acesso Pta Ebt fora do padrão.")
			alert("Padrão de Designação:\nMin.:  "+ padraoMin + "\nMax.: " + padraoMax + ".");
			txtNroAcessoEbt.focus()
			return false
		}
		if (!parent.IFrmProcesso1.ValidarPadrao(document.forms[0].txtNroAcessoEbt)) return false
	}
	return true
}

function GravarFacilidade()
{
	
	//MSCAPRI - 30/08/2013
	//CH-43910FYZ
	
	//var objNode_tp_proc = objXmlGeral.selectSingleNode("//xDados/hdnTipoProcesso").text
	
	//DIFERENTE DE CANCELAMENTO
	if(document.forms[0].hdnTipoProcesso.value != 4){	
	
		var objNode_rede = objXmlGeral.selectSingleNode("//xDados/Facilidade/cboRede").text

		if(objNode_rede == 11)//FO ETHERNET
		{
			objNode = objXmlGeral.selectSingleNode("//xDados/Facilidade")
			var objNodeRequest2 = objNode.getElementsByTagName("txtVlan_FO")
			
			var objNodeRequest3 = objNode.getElementsByTagName("txtPE_FO")
			
			var objNodeRequest4 = objNode.getElementsByTagName("txtPorta_FO")
			
			var objNodeRequest5 = objNode.getElementsByTagName("txtSvlan_FO")

			if (objNodeRequest2.length == 0)
			{
				alert("'VLAN' obrigatório. Selecione o acesso e insira a 'VLAN'");	
				return		
			}
		
			if (objNodeRequest3.length == 0){
				alert("'PE' obrigatório. Selecione o acesso e insira a 'PE'");	
				return
			}

			if (objNodeRequest4.length == 0){
				alert("'PORTA' obrigatório. Selecione o acesso e insira a 'PORTA'");	
				return
			}

			if (objNodeRequest5.length == 0){
				alert("'SVLAN' obrigatório. Selecione o acesso e insira a 'SVLAN'");
				return
			}
		}
	}
	//
	
	/** retirada monetanea
	if (document.Form2.hdnTecnologia.value != "FO EDD")
	{
		if (IFrmEntregaProv.Form1.txtDtPrev.value=="" || IFrmEntregaProv.Form1.cboProjetoEspecial.value=="" || IFrmEntregaProv.Form1.hdnGravouPrevisao.value != "Sim")
		{
		  alert("Previsões de Entrega do Provedor é um item obrigatório!");
		  IFrmEntregaProv.Form1.txtDtPrev.focus();
		  return
		}
	}
**/
	with(document.forms[0])
	{

		
		if ( cboRede.value == 1 ){
			hdnPlataforma.value =  cboPlataforma.value
		}
		//Foi alterado a ordem do fluxo de envio de email.
		//Primeiro há a gravação da facilidade e só depois o envio do email com campo Acf_NroAcessoPtaEbt preenchido.
		/*
		if (hdnTipoProcesso.value != 4)
		{
			if (hdnIdAcessoFisico.value != "0" && hdnIdAcessoFisico.value != "")
			{ //Compartilhamento não faz verificações desse tipo
			}
			else
			{
				if (!ValidarCampos(cboLocalInstala,"Local de Entrega")) return false
				if (!ValidarCampos(cboLocalConfig,"Local de Configuração")) return false
				if (!ValidarCampos(cboDistLocalInstala,"Distribuidor")) return false
				if (!ValidarCampos(cboRede,"Rede")) return false
				if (!ValidarCampos(cboProvedor,"Provedor")) return false
				if (!ValidarCampos(cboRegimeCntr,"Prazos de Contratação de Acesso")) return false
			}
			
		var intRet=alertbox('Deseja envia e-mail para o provedor?','Sim','Não','Sair')
			switch (parseInt(intRet))
			{
				case 1:
					if (!ValidarCampos(cboProvedor,"Provedor")) return false
					EnviarEmail(0,document.forms[0].cboProvedor(document.forms[0].cboProvedor.selectedIndex).tag_provedor)
					break

				case 3:
					return
					break
			}	
		}else
		{
			var intRet=alertbox('Deseja envia e-mail para o provedor?','Sim','Não','Sair')
			switch (parseInt(intRet))
			{
				case 1:
					if (!ValidarCampos(cboProvedor,"Provedor")) return false
					hdnEnviarEmail.value = "EnviarEmailProvedor"
					break

				case 3:
					return
					break
			}
		}
		*/
		AdicionarNode(objXmlGeral,"cboLocalInstala",cboLocalInstala.value)
		AdicionarNode(objXmlGeral,"cboLocalConfig",cboLocalConfig.value)
		AdicionarNode(objXmlGeral,"cboDistLocalInstala",cboDistLocalInstala.value)
		AdicionarNode(objXmlGeral,"cboProvedor",cboProvedor.value)
		AdicionarNode(objXmlGeral,"cboRegimeCntr",cboRegimeCntr.value)
		AdicionarNode(objXmlGeral,"cboPromocao",cboPromocao.value)
		AdicionarNode(objXmlGeral,"hdnIdLog",hdnIdLog.value)
		AdicionarNode(objXmlGeral,"hdnIdAcessoFisico",hdnIdAcessoFisico.value)
		AdicionarNode(objXmlGeral,"hdnCompartilhamento",hdnCompartilhamento.value)
		AdicionarNode(objXmlGeral,"hdnIdAcessoFisico1",hdnIdAcessoFisico1.value)
		AdicionarNode(objXmlGeral,"hdnCompartilhamento1",hdnCompartilhamento1.value)
		AdicionarNode(objXmlGeral,"hdnAlteracao",hdnAlteracao.value)
		AdicionarNode(objXmlGeral,"hdnPedId",hdnPedId.value)
		AdicionarNode(objXmlGeral,"hdnTipoProcesso",hdnTipoProcesso.value)
		AdicionarNode(objXmlGeral,"hdnFacDetid",hdnFacDetid.value)
		if (hdnTipoProcesso.value != 3 && hdnTipoAcao.value !=1)
		{
			AdicionarNode(objXmlGeral,"txtObsProvedor",txtObsProvedor.value)
		}
		//@@Davif 
		//AdicionarNode(objXmlGeral,"txtDtPrevEntregaProvedor",txtDtPrevEntregaProvedor.value)
		//AdicionarNode(objXmlGeral,"txtDtReaprPrevEntregaProvedor",txtDtReaprPrevEntregaProvedor.value)
		//

		if (hdnTipoProcesso.value == 4)
		{
			//Coloca o tipo de ação de todas asn facilide como R
			var objNodeCan = objXmlGeral.selectNodes("//Facilidade[strTipoConsulta='R']")
			if (objNodeCan.length > 0){
				for (var intIndex=0;intIndex<objNodeCan.length;intIndex++)
				{
					var objNode = objNodeCan[intIndex].getElementsByTagName("TipoAcao")
					if (objNode.length > 0){
						objNode[0].childNodes[0].text = "R"
					}	
				}	
			}
			hdnAcao.value = "GravarFacilidade"
			hdnRede.value = cboRede.value
			//alert('objXmlGeral.xml...')
			//alert(objXmlGeral.xml)
			hdnXml.value = objXmlGeral.xml 
			
			target = "IFrmProcesso"  
			action = "ProcessoFac.asp"
			method = "Post"
			submit()
		}
		else
		{	
			var objNodeE = objXmlGeral.selectNodes("//Facilidade[strTipoConsulta='C']")
			if (objNodeE.length > 0){
				//Confirmação de Compartilhamento do Id'Físico
				hdnAcao.value = "AutorizarCompartilhamento"
				hdnRede.value = cboRede.value
				hdnSubAcao.value = "IdFisEndInstala"
				target = "IFrmProcesso"
				action = "ProcessoCla.asp"
				submit()
			}else{
				//Verifica se o usuário editou e não atualizou a lista
				if (FacilidadeAlteradaNaoAtualizada()){
					var intRet=alertbox('As informações da facilidade atualmente editada foram alteradas. Deseja atualizá-la e prosseguir com a alocação?','Sim','Não','Sair')
					switch (parseInt(intRet))
					{
						case 1:
							if (!AddFac(true)) return
							break
						case 3:
							return
							break
					}		
				}

				var objNode = objXmlGeral.selectNodes("//Facilidade")
				if (objNode.length > 0){
					hdnAcao.value = "GravarFacilidade"
					hdnRede.value = cboRede.value
					hdnXml.value = objXmlGeral.xml 
					target = "IFrmProcesso"  
					action = "ProcessoFac.asp"
					method = "Post"
					submit()
				}else{
					//alert("Favor informar as facilidades do pedido ou adicionar facilidades a lista.")
					if (!AddFac(true)) return
					hdnAcao.value = "GravarFacilidade"
					hdnRede.value = cboRede.value
					hdnXml.value = objXmlGeral.xml 
					target = "IFrmProcesso"  
					action = "ProcessoFac.asp"
					method = "Post"
					submit()
				}	
			}	
		}
	
			if (hdnTipoProcesso.value != 4)
		{
			if (hdnIdAcessoFisico.value != "0" && hdnIdAcessoFisico.value != "")
			{ //Compartilhamento não faz verificações desse tipo
			}
			else
			{
				if (!ValidarCampos(cboLocalInstala,"Local de Entrega")) return false
				if (!ValidarCampos(cboLocalConfig,"Local de Configuração")) return false
				if (!ValidarCampos(cboDistLocalInstala,"Distribuidor")) return false
				if (!ValidarCampos(cboRede,"Rede")) return false
				if (!ValidarCampos(cboProvedor,"Provedor")) return false
				if (!ValidarCampos(cboRegimeCntr,"Prazos de Contratação de Acesso")) return false
			}
			//Foi alterado a ordem do fluxo de envio de email.
			//Primeiro há a gravação da facilidade e só depois o envio do email com campo Acf_NroAcessoPtaEbt preenchido.
			//Inicio alteração.
				//if ( cboRede.value != 10 ){
				if (hdnEILD.value != "S")
				{
					  var intRet=alertbox('Deseja envia e-mail para o provedor?','Sim','Não','Sair')
						switch (parseInt(intRet))
						{
							case 1:
								if (!ValidarCampos(cboProvedor,"Provedor")) return false
								EnviarEmail(0,document.forms[0].cboProvedor(document.forms[0].cboProvedor.selectedIndex).tag_provedor)
								break
			
							case 3:
								return
								break
						}	
				}
				//}
		}else
		{
			if (hdnEILD.value != "S")
			{
			//if ( cboRede.value != 10 ){
				var intRet=alertbox('Deseja envia e-mail para o provedor?','Sim','Não','Sair')
	
				switch (parseInt(intRet))
				{
					case 1:
						if (!ValidarCampos(cboProvedor,"Provedor")) return false
						hdnEnviarEmail.value = "EnviarEmailProvedor"
						break
	
					case 3:
						return
						break
				}
			}
			//}
		}
		//Fim alteração.
	}	 
}

function ResgatarProvedoresAssociados(obj){
	with(document.forms[0])
	{
		hdnProvedor.value = obj.value
		hdnAcao.value = "ResgatarProvedoresAssociados"
		target = "IFrmProcesso3"  
		action = "ProcessoCla.asp"
		method = "Post"
		submit()
	}	
}

function AbrirInterligacoes()
{
	with(document.forms[0]){
		if (Ped_Id.value != ""){
			AdicionarNode(objXmlGeral,"Ped_Id",Ped_Id.value)
			var objAry = new Array(0,hdnIntIndice.value,Ped_Id.value)
			var intRet = window.showModalDialog('InterligacaoPedido.asp?hdnPedId='+Ped_Id.value,objAry,'dialogHeight: 200px; dialogWidth: 350px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
			if (intRet == 1){
				intRet = 0
				//Alterar(1) 
			}
		}else{
			alert("Número do pedido não encontrado.")
			return
		}	
	}
}

//Verifica se o acesso foi alterado e não foi recolocado a lista
function FacilidadeAlteradaNaoAtualizada()
{
	with (document.forms[0])
	{
		if (hdnIntIndice.value != ""){
			intChave = hdnIntIndice.value
		}else{
			return false
		}

		if (hdnPropAcesso.value  != "EBT" )
		{
			switch (cboRede.value)
			{
				case "1":
					objAry = new Array("txtNroAcessoEbt","txtFila","txtBastidor","txtRegua","txtPosicao","txtTimeslot","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB","cboPropModem","txtQtdeModem","txtObsFac")
					break
				case "2":	
					objAry = new Array("txtNroAcessoEbt","txtTronco","txtPar","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB","cboPropModem","txtQtdeModem","txtObsFac")
					break
				case "3":
					objAry  = new Array("txtNroAcessoEbt","txtCabo","txtPar","txtLateral","cboTipoCabo","txtCaixaEmenda","cboPropModem","txtQtdeModem","txtObsFac")
					break
				case "4":
					objAry  = new Array("txtNroAcessoEbt","txtVlan")
					break
				case "5":
					objAry  = new Array("txtNroAcessoEbt","txtVlan","txtPE","txtPorta","txtLink")
					break				
				case "13":
					objAry  = new Array("txtNroAcessoEbt","txtVlan","txtPE","txtPorta","txtSvlan","cboTipoPorta")
					break							
				case "10":
					objAry  = new Array("txtNroAcessoEbt")
					break	
				case "11":
					objAry  = new Array("txtNroAcessoEbt","txtVlan_FO","txtPE_FO","txtPorta_FO","txtSvlan_FO" ,"cboRede" , "hdnOrigem" , "hdnId_AcessoAsms")
					break			
			}		
		}
		else
		{
			switch (cboRede.value)
			{
				case "1":
					objAry = new Array("txtNroAcessoEbt","txtFila","txtBastidor","txtRegua","txtPosicao","txtTimeslot","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB","cboPropModem","txtQtdeModem","txtObsFac")
					break
				case "2":
					objAry = new Array("txtNroAcessoEbt","txtTronco","txtPar","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB","cboPropModem","txtQtdeModem","txtObsFac")
					break
				case "3":
					objAry = new Array("txtNroAcessoEbt","txtCabo","txtPar","txtLateral","cboTipoCabo","txtCaixaEmenda","cboPropModem","txtQtdeModem","txtObsFac")
					break
				case "4":
					objAry  = new Array("txtNroAcessoEbt","txtVlan")
					break
				case "5":
					objAry  = new Array("txtNroAcessoEbt","txtVlan","txtPE","txtPorta","txtLink")
					break	
				case "13":
					objAry  = new Array("txtNroAcessoEbt","txtVlan","txtPE","txtPorta","txtSvlan","cboTipoPorta")
					break						
				case "10":
					objAry  = new Array("txtNroAcessoEbt")
					break	
				case "11":
					objAry  = new Array("txtNroAcessoEbt","txtVlan_FO","txtPE_FO","txtPorta_FO","txtSvlan_FO" ,"cboRede" , "hdnOrigem" , "hdnId_AcessoAsms")
					break	
			}		
		}

		if (cboRede.value != 3)
		{
			var intIndexRb = RequestNodeFac(objXmlGeral,"rdoUrbanoIndex",intChave)
			if (intIndexRb != ""){
				if (!rdoUrbano[intIndexRb].checked){
					return true
				} 
			}	
		}

		if (hdnPropAcesso.value  != "EBT" )
		{
			var intIndexRb = RequestNodeFac(objXmlGeral,"rdoFaturaIndex",intChave)
			if (intIndexRb != ""){
				if (!rdoFatura[intIndexRb].checked){
					return true
				} 
			}	
		}

		for (var intIndex=0;intIndex<objAry.length;intIndex++)
		{
			var objChildForm = new Object(eval("document.forms[0]."+objAry[intIndex]))
			if (objChildForm.value != RequestNodeFac(objXmlGeral,objAry[intIndex],intChave))
			{
				return true
			}
		}
		return false
	}
}
function JanelaRedeDet()
{
	var blnDisabled 
	
	if (document.forms[0].cboRede.value ==1){ 
		document.forms[0].hdnPlataforma.value =  document.forms[0].cboPlataforma.value
	}
	PopularXml(objXmlReturn)
	objXmlReturn = window.showModalDialog('RedeDet.asp',objXmlReturn,'dialogHeight: 380px; dialogWidth: 780px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	with (document.forms[0])
	{
		if (RequestNode(objXmlReturn,"txtRedDetBastidor") != "")
		{
			//Recurso
			cboRede.value = 1
			ControleRede(cboRede)
			cboLocalInstala.value = RequestNode(objXmlReturn,"txtRedDetEstacao")
			cboDistLocalInstala.value = RequestNode(objXmlReturn,"txtRedDetDistribuidor")
			hdnDstId.value = RequestNode(objXmlReturn,"txtRedDetDistribuidor")
			//ResgatarDistribuicao(cboLocalInstala,1)
			cboProvedor.value = RequestNode(objXmlReturn,"txtRedDetProvedor")
			ControleProvedor(cboProvedor)
			//Facilidade
			txtBastidor.value	= RequestNode(objXmlReturn,"txtRedDetBastidor")
			txtRegua.value		= RequestNode(objXmlReturn,"txtRedDetRegua")
			txtPosicao.value	= RequestNode(objXmlReturn,"txtRedDetPosicao")
			txtTimeslot.value	= RequestNode(objXmlReturn,"txtRedDetTimeslot")
			txtFila.value		= RequestNode(objXmlReturn,"txtRedDetFila")
			hdnFacDetid.value   = RequestNode(objXmlReturn,"hdnFacDetid")
			
			if (cboPlataforma.disabled == true) blnDisabled = true 
			RetornaCboPlataforma(1,RequestNode(objXmlReturn,"txtRedDetPlataforma"))
			if (blnDisabled == true) cboPlataforma.disabled = true
			
			AdicionarNode(objXmlReturn,"txtRedDetBastidor","")
		}
	}
}
function ControleProvedor(obj)
{
	with(document.forms[0])
	{
		if (hdnProvedor.value != obj.value)
		{
			ApagarNroAcesso()
			ResgatarPromocaoRegime(obj)
			ResgatarProvedoresAssociados(obj)
			if (cboRede.value ==3)
			{
				ResgatarPadraoProvedor(obj,0)
			}
		}	
		hdnProvedor.value = obj.value
	}
}
function ControleRede(obj)
{
	with(document.forms[0])
	{
		if (hdnRede.value != obj.value)
		{
			AbrirParam(obj); // <--
			LimparFacSel();
			
			if (obj.value ==3)
			{
				ResgatarPadraoProvedor(document.forms[0].cboProvedor,1)
				cboProvedor.value = "11"
				hdnProvedor.value = "11"
				cboProvedor.disabled = true

				hdnAcao.value = "ResgatarPromocaoRegime"
				hdnProvedor.value = 11
				target = "IFrmProcesso"
				action = "ProcessoCla.asp"
				submit()
			} 
			if (obj.value == 1 || obj.value == 2)
			{
				ResgatarProvedoresAssociados(document.forms[0].cboProvedor)
				//document.forms[0].cboProvedor.value = ""
				//document.forms[0].hdnProvedor.value = ""
				document.forms[0].cboProvedor.disabled = false
				//ResgatarPromocaoRegime(0)
			}
			if (obj.value == 4 || obj.value == 5 || obj.value == 10 || obj.value == 11 || obj.value == 13)
			{
				//hdnProvedor.value = document.forms[0].cboProvedor
				ResgatarPadraoProvedor(document.forms[0].cboProvedor,1)
				
				document.forms[0].cboProvedor.disabled = false
								
			} 
		}
		hdnRede.value = obj.value
	}	
}
function JanelaConfirmacaoFac(objXmlGeral){
	var intRet = window.showModalDialog('ConfirmacaoAlocacaoFac.asp',objXmlGeral,'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	switch (intRet)
	{
		case "159":
			with (document.forms[0])
			{
				if (hdnCompartilhamento.value == "1" || hdnCompartilhamento1.value == "1")
				{
					target = self.name 
					action = "Facilidade_Main.asp"
					method = "Post"
					submit()
				}
			}
			break
		case "165":
			with (document.forms[0])
			{
				if (hdnEnviarEmail.value != "")
				{
					EnviarEmail(0,'') //Envia e-mail para provedor após um cancelamento
				}else{
					target = self.name 
					action = "Facilidade_Main.asp"
					method = "Post"
					submit()
				}
			}	
	}
} 

function liberarServico()
{
	if (confirm('Confirma a liberação para serviço?'))
	{
		with (document.forms[0])
		{
			hdnAcao.value = "LiberaServico"
			target = "IFrmProcesso"
			action = "ProcessoFac.asp"
			submit()	
		}
	 }
}

function RealizaCadastroOS(FormatoPed,AcfID,PedID,DtEnvioEmail,DtAtual)
{  
	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
	var repl , strXML , strRetorno
	var strData 
	var checkStr = ""
	var checkStr = FormatoPed
	with (document.forms[0])
	{
		var checkOK = TxtCodigoOS.value
		strData = TxtDtEmissao.value
		var DtHoje = new Date(DtAtual.substring(6,10),DtAtual.substring(3,5)-1,DtAtual.substring(0,2),0,0,0,0);
		var DtEnvio = new Date(DtEnvioEmail.substring(6,10),DtEnvioEmail.substring(3,5)-1,DtEnvioEmail.substring(0,2),0,0,0,0);
		var DtCheck = new Date(strData.substring(6,10),strData.substring(3,5)-1,strData.substring(0,2),0,0,0,0);
		if (TxtDtEmissao.value == "")
	    {
			alert("Data de Emissão OS não pode ser vazio");
	  		TxtDtEmissao.focus();
	  		return;
        }
		if (!ValidarData(strData))
		{
			TxtDtEmissao.focus();
			return;
		}
		//strData = strData.substring(3,5) + "/" + strData.substring(0,2) + "/" + strData.substring(6,10);
		if (TxtCodigoOS.value == "")
	    {
			alert("Código OS não pode ser vazio");
	  		TxtCodigoOS.focus();
	  		return;
        }
        if (DtEnvio > DtCheck)
		{
			alert("A data deve ser maior ou igual a data de envio da carta");
	 		TxtDtEmissao.focus();
	  		return;
		}
		if (DtCheck > DtHoje)
		{
			alert("A data deve ser menor ou igual a data atual");
	 		TxtDtEmissao.focus();
	 		return;
		}
		
        if(checkStr != "")
        {
			if(checkStr.length > checkOK.length || checkOK.length > checkStr.length )
			{
				alert("Campo fora do padrão.") 
				return;
			}
			for (var i = 0;  i < checkStr.length;  i++)
			{
				if(checkStr.charAt(i) == "N") //Int
				{
					if(!ValidarTipo(checkOK.charAt(i),0))
					return;
				}
				else if(checkStr.charAt(i) == "A") //AlfaNumerico
				{
					if(!ValidarTipo(checkOK.charAt(i),2))
					return;
				}
				else if(checkStr.charAt(i) == "L") //Alfa
				{
					if(!ValidarTipo(checkOK.charAt(i),1))
					return;
				}
				else if(checkStr.charAt(i) != checkOK.charAt(i))
				{
					alert("Campo fora do padrão.") 
					return;
				}
			}
		}
		if (!window.confirm('Confirma OS informada?'))
		{
			return;
		}
		strXML = "<root>"
		strXML = strXML + "<Acf>" +  AcfID + "</Acf>"
		strXML = strXML + "<Ped>" + PedID + "</Ped>" 
		strXML = strXML + "<OS>" + TxtCodigoOS.value + "</OS>" 
		strXML = strXML + "<DtEmissao>" + strData + "</DtEmissao>"
		strXML = strXML + "</root>"
		xmlDoc.loadXML(strXML);
			
		strRetorno = null;
		xmlhttp.Open("POST","RealizaCadastroOS.asp" , false);
		xmlhttp.Send(xmlDoc.xml);
							
		strRetorno = xmlhttp.responseText;
		
		strRetorno = strRetorno.substring(strRetorno.lenght -200,500)
				
		alert(strRetorno);
		
		window.returnValue = 'true';
		
		window.close();
	}	
}

function RealizaCadastroManobra(varAcfID,varRecID)
{  
	var i;
	var strDados;
	var arrayCampos;
	var varDelete;
	var deuErro = 0;
	var retorno;
	retorno = "";
	with (document.forms[0])
	{
		if (!window.confirm('Confirma Alocação das facilidades alteradas?'))
		{
			return;
		}
		varDelete = hdnDelete.value 
		if (hdnListRep.value != "")
		{
			hdnDelete.value = hdnListRep.value
			if(!ApagaFacilidadesManobra())
				deuErro = 1
		}
		if(deuErro != 1)
		{
			//if (document.getElementById('tblFac').rows.length < 2)
			//	{
			//	   alert("Para alocar é necessário informar uma Facilidade.")
			//	   deuErro = 1
			//	   return
			//	}		
		
			for(i = 0; i < document.getElementById('tblFac').rows.length;i++)
			{
				if((document.getElementById('hdnDados' + i ) != null) && (document.getElementById('hdnDados'+ i ).value != ""))
				{
					retorno = AlocaFacilidadeManobra(i,varAcfID,varRecID)
					if(retorno != "Sucesso")
					  {
						deuErro = 1
				}
			}
		}
		}
		if(deuErro != 1)
		{
			hdnDelete.value = varDelete 
		if(hdnDelete.value != "")	
		{
			if(!ApagaFacilidadesManobra())
				deuErro = 1
		}
		}
		if(deuErro == 1)
		{
			if(retorno != "Erro" || retorno != ''){
				alert(retorno);
				parent.window.ProcurarNroAcesso();
			}
			else
				alert("Ocorreu um erro ao alocar a facilidade");
		}
		else
		{
			alert("Facilidade alocada com sucesso");
			parent.window.ProcurarNroAcesso();
		}

	}

}

function ApagaFacilidadesManobra()
{
	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
	var arrayCampos;
	var  strXML , strRetorno
	var i;
	with (document.forms[0])
	{
		arrayCampos = hdnDelete.value.split(",");
		for (i = 0; i < arrayCampos.length; i++)
		{
			
		strXML = "<root>"
			strXML = strXML + "<FacID>" + arrayCampos[i]  + "</FacID>"
		strXML = strXML + "<PedDM>" + hdnPedDM.value  + "</PedDM>"
		strXML = strXML + "</root>"
		xmlDoc.loadXML(strXML);
		strRetorno = null;
		xmlhttp.Open("POST","RemoveFacManobra.asp" , false);
		xmlhttp.Send(xmlDoc.xml);
				
		strRetorno = xmlhttp.responseText;
		}
		if(strRetorno != "Facilidades Removidas com Sucesso!")
			return false;
		return true;
	}
	
}

function AlocaFacilidadeManobra(Row,varAcfID,varRecID)
{
	var strDados = new String();
	var arrayCampos;
	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
	var strXML , strRetorno
	var strData
	with (document.forms[0])
	{
		strDados = document.getElementById('hdnDados' + Row).value
		arrayCampos = strDados.split("&&");
		strXML = "<root>"
		strXML = strXML + "<FacID>" +  document.getElementById('hdnFacID' + Row).value + "</FacID>"
		strXML = strXML + "<PedID>" + hdnPedDM.value  + "</PedID>"
		strXML = strXML + "<AcfID>" + varAcfID  + "</AcfID>"
		strXML = strXML + "<ProID>" + hdnProvedor2.value + "</ProID>"
		if( hdnRepresentacao.value == "") 
			strXML = strXML + "<Representacao></Representacao>"
		else
			strXML = strXML + "<Representacao>" + hdnRepresentacao.value  + "</Representacao>"
			
		
		strXML = strXML + "<Rec>" + varRecID  + "</Rec>"

		if(txtRede.value == "DETERMINISTICO")
		{
			strXML = strXML + "<Fila>" +arrayCampos[2]  + "</Fila>"
			strXML = strXML + "<Bastidor>" +arrayCampos[3]  + "</Bastidor>"
			strXML = strXML + "<Regua>" +arrayCampos[4]  + "</Regua>"
			strXML = strXML + "<Posicao>" +arrayCampos[5]  + "</Posicao>"
			strXML = strXML + "<Timeslot>" +arrayCampos[6]  + "</Timeslot>"
			strXML = strXML + "<NumAcessoPtaEbt>" +arrayCampos[1]  + "</NumAcessoPtaEbt>"
			strXML = strXML + "<NumAcessoCLI>" +arrayCampos[8]  + "</NumAcessoCLI>"
			strXML = strXML + "<CodProv>" + arrayCampos[0] + "</CodProv>"
			strXML = strXML + "<CCTOProvedor>" +arrayCampos[7]  + "</CCTOProvedor>"
			strXML = strXML + "<CNLPontaA>" +arrayCampos[9]  + "</CNLPontaA>"
			strXML = strXML + "<CNLPontaB>" +arrayCampos[10]  + "</CNLPontaB>"
			strXML = strXML + "<SisID>1</SisID>"
			strXML = strXML + "<PropModem>" +arrayCampos[12]  + "</PropModem>"
			strXML = strXML + "<QtdeModem>" +arrayCampos[11]  + "</QtdeModem>"
			if (arrayCampos[16] == "S")
				strXML = strXML + "<Fatura>S</Fatura>"
			else
				strXML = strXML + "<Fatura>N</Fatura>"
			if (arrayCampos[15] == "U")
				strXML = strXML + "<CCTOTipo>U</CCTOTipo>"
			else
				strXML = strXML + "<CCTOTipo>I</CCTOTipo>"
			strXML = strXML + "<AreaObs>" +arrayCampos[14]  + "</AreaObs>"
			strXML = strXML + "</root>"
		}
		else if(txtRede.value == "NAO DETERMINISTICO")
		{
			strXML = strXML + "<NumAcessoPtaEbt>" +arrayCampos[1]  + "</NumAcessoPtaEbt>"
			strXML = strXML + "<CodProv>" + arrayCampos[0] + "</CodProv>"
			strXML = strXML + "<Tronco>" + arrayCampos[2] + "</Tronco>"
			strXML = strXML + "<Par>" + arrayCampos[3] + "</Par>"
			strXML = strXML + "<CCTOProvedor>" + arrayCampos[4] + "</CCTOProvedor>"
			strXML = strXML + "<NumAcessoCLI>" +arrayCampos[5]  + "</NumAcessoCLI>"
			strXML = strXML + "<CNLPontaA>" +arrayCampos[6]  + "</CNLPontaA>"
			strXML = strXML + "<CNLPontaB>" +arrayCampos[7]  + "</CNLPontaB>"
			strXML = strXML + "<SisID>2</SisID>"
			strXML = strXML + "<PropModem>" +arrayCampos[9]  + "</PropModem>"
			strXML = strXML + "<QtdeModem>" +arrayCampos[8]  + "</QtdeModem>"
			strXML = strXML + "<AreaObs>" +arrayCampos[10]  + "</AreaObs>"		
			if (arrayCampos[12] == "S")
				strXML = strXML + "<Fatura>S</Fatura>"
			else
				strXML = strXML + "<Fatura>N</Fatura>"
			if (arrayCampos[11] == "U")
				strXML = strXML + "<CCTOTipo>U</CCTOTipo>"
			else
				strXML = strXML + "<CCTOTipo>I</CCTOTipo>"
			strXML = strXML + "</root>"
		}
		else
		{
			strXML = strXML + "<NumAcessoPtaEbt>" +arrayCampos[0]  + "</NumAcessoPtaEbt>"
			strXML = strXML + "<Cabo>" + arrayCampos[1] + "</Cabo>"
			strXML = strXML + "<Par>" + arrayCampos[2] + "</Par>"
			strXML = strXML + "<Derivacao>" + arrayCampos[4] + "</Derivacao>"
			strXML = strXML + "<Pade>" + arrayCampos[3] + "</Pade>"
			strXML = strXML + "<TipoCabo>" + arrayCampos[6] + "</TipoCabo>"
			strXML = strXML + "<SisID>3</SisID>"
			strXML = strXML + "<PropModem>" +arrayCampos[5]  + "</PropModem>"
			strXML = strXML + "<QtdeModem>" +arrayCampos[7]  + "</QtdeModem>"
			strXML = strXML + "<AreaObs>" +arrayCampos[8]  + "</AreaObs>"		
			strXML = strXML + "</root>"	
		}
		xmlDoc.loadXML(strXML);
					
		strRetorno = null;
		xmlhttp.Open("POST","RealizaCadastroManobra.asp" , false);
		xmlhttp.Send(xmlDoc.xml);
							
		strRetorno = xmlhttp.responseText;
		
		strRetorno = strRetorno.substring(strRetorno.lenght -200,500)
		return strRetorno
	}	
}

function ExcluirSenha()
{  
	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
	var repl , strXML , strRetorno
	var strData 
	strXML = "<root>"
	with (document.forms[0])
	{
		if(hdnCheck.value == "")
		{
			alert('Selecione pelo menos um registro para excluir');
			return;
		}
		if (!window.confirm('Confirma a Exclusão?'))
		{
			return;
		}
		strXML = strXML + "<Aprov>" + hdnCheck.value + "</Aprov>"
		strXML = strXML + "</root>"
		xmlDoc.loadXML(strXML);
			
		strRetorno = null;
		xmlhttp.Open("POST","ExcluirAutorizarAcesso.asp" , false);
		xmlhttp.Send(xmlDoc.xml);
							
		strRetorno = xmlhttp.responseText;
					
		alert(strRetorno);
		
		parent.parent.window.ProcurarSenha();
						
	}	
}

function EnviarEmailConfiguracao(hdnIdLog)
{  
		var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
		var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
		var repl , strXML , strRetorno
		
		strXML = "<root>"
		strXML = strXML + "<acl>" + hdnIdLog + "</acl>" 
		strXML = strXML + "</root>"

		//alert(strXML);
		
		xmlDoc.loadXML(strXML);
		
		xmlhttp.Open("POST","gponRip.asp" , false);		
		
		xmlhttp.Send(xmlDoc.xml);
		strXML  = xmlhttp.responseText;
		
		//alert(strXML);
		xmlDoc.loadXML(strXML);
		xmlhttp.Open("POST", "RetornaCartaConfiguracao.asp", false);
		xmlhttp.Send(xmlDoc.xml);
		
		strXML = xmlhttp.responseText;
		//alert(strXML);
		objWindow = window.open("About:blank", null, "status=no,toolbar=no,menubar=no,location=no,resizable=Yes,scrollbars = Yes");
		objWindow.document.write(strXML);		
		objWindow.document.close();
		
			
}


function EnviarEmail(intTipo, strCartaProv)
{   //alert("xmlFacObjects.js(): ")
	with (document.forms[0])
	{		

                if (cboRede.value == "")
	        {
		  //alert("Selecione a Rede!")
	  	  //cboRede.focus()
	  	  //return
        	}
		
		if (cboProvedor.value != "")
		{
			switch (parseInt(intTipo))
			{
				case 0:
					hdnRede.value = cboRede.value
					var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
					var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
					var repl , strXML , strRetorno
									
					with(document.forms[0]){
						hdnAcao.value = "GravaObsPed"
						target = "IFrmProcesso"
						action = "ProcessoCla.asp"
						submit()		
						
					}	
					
					strXML = "<root>"
					strXML = strXML + "<AcfID>" +  hdnAcfIdRadio.value + "</AcfID>"
					
					strXML = strXML + "</root>"

					//alert('strXML1='+strXML);
					xmlDoc.loadXML(strXML);
	
					xmlhttp.Open("POST","RetornaCheck.asp" , false);
					xmlhttp.Send(xmlDoc.xml);	
					
					strRetorno = xmlhttp.responseText;
					
					//alert('strRetorno= ' + strRetorno);
					if (strRetorno=="780") {
						alert('FAVOR EFETUAR A GRAVAÇÃO DOS DADOS TÉCNICOS. ');
						return
					}
					
			
					strXML = "<root>"
					strXML = strXML + "<ped>" +  hdnPedId.value + "</ped>"
					strXML = strXML + "<sol>" + hdnSolId.value + "</sol>" 
					strXML = strXML + "<Prov>" + cboProvedor.value + "</Prov>"
					strXML = strXML + "<Esc>" + cboLocalInstala.value + "</Esc>"
					strXML = strXML + "<ndTipo>" +  hdnTipoProcesso.value + "</ndTipo>"
					strXML = strXML + "<Rede>" + hdnRede.value + "</Rede>"
					strXML = strXML + "</root>"

					xmlDoc.loadXML(strXML);
	
					xmlhttp.Open("POST","RetornaModelo.asp" , false);
					xmlhttp.Send(xmlDoc.xml);	
					
					strRetorno = xmlhttp.responseText;		
//*					alert('strRetorno='+strRetorno);
					xmlhttp.Open("POST",strRetorno , false);		
					
					xmlhttp.Send(xmlDoc.xml);
					
					strXML  = xmlhttp.responseText;
//alert('strXML1='+strXML);
					var posFound = strXML.search("http_404.htm")
					if (posFound != -1)
					{
						if  (strRetorno == "naoencontrado.asp"){
							alert("Não foi possível enviar o e-mail, verifique o cadastro do provedor contrato.")
							return 
						}
						alert("Não foi possível enviar o e-mail, verifique o cadastro do provedor. - Formulário padrão.")
						return 
					}
					
					if (strRetorno != "CartaPadrao.asp" && strRetorno != "ProcessoEmailProvedorRadioIP.asp"  && strRetorno != "ProcessoEmailProvedorPadrao.asp"  && strRetorno != "" &&  strXML.substring(0,6) != "<table")
					{
						repl = /&/g	
						strXML = strXML.replace(repl,"&amp;");
						xmlDoc.loadXML(strXML);
						xmlhttp.Open("POST", "RetornaCarta.asp", false);
						xmlhttp.Send(xmlDoc.xml);
					}
					
					strXML = xmlhttp.responseText;
					//alert('strXML2='+strXML);
					objWindow = window.open("About:blank", null, "status=no,toolbar=no,menubar=no,location=no,resizable=Yes,scrollbars = Yes");
					objWindow.document.write(strXML);		
					objWindow.document.close();
					
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
function EnviarEmailLiberacao(provedor,sol_id,usunome)
{

  
	document.Form3.btnAlterar.disabled = true
	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
	var repl ,rep2, strXML , strRetorno,NAcfId
	var sol_obs  = document.forms[0].txtObs.value
	

	with(document.forms[0]){
		
		hdnAcao.value = "GravaObs"
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()		
		
	}	
	strXML = "<root>"
	strXML = strXML + "<sol>" +  sol_id + "</sol>"
	strXML = strXML + "<user>" +  usunome + "</user>"
	strXML = strXML + "</root>"
	
	
	xmlDoc.loadXML(strXML);
	
	// Envia os dados sol_id e user para a pagina RetornaDados.asp
	xmlhttp.Open("POST","RetornaDados.asp" , false);
	xmlhttp.Send(xmlDoc.xml);
		
	strXML = xmlhttp.responseText;

		
	xmlDoc.loadXML(strXML);
		
	NAcfId = xmlDoc.getElementsByTagName("Acf")[0].firstChild.nodeValue
		
	xmlhttp.Open("POST","RetornaModeloLiberacao.asp" , false);
	xmlhttp.Send(xmlDoc.xml);
	
	
	
	strRetorno = xmlhttp.responseText;
	

	xmlhttp.Open("POST",strRetorno , false);
		
	xmlhttp.Send(xmlDoc.xml);
	
	strXML  = xmlhttp.responseText;	
	
		
	var posFound = strXML.search("http_404.htm")
	
	if (posFound != -1)
	{

		if  (strRetorno == "naoencontrado.asp"){
			alert("Não foi possível enviar o e-mail, verifique o cadastro do provedor contrato.")
			return 
		}
		alert("Não foi possível enviar o e-mail, verifique o cadastro do provedor. - Formulário padrão.")
		return 
	}
	
	if (strRetorno != "ProcessoEmailProvedorPadraoLiberacao.asp"  && strRetorno != "" &&  strXML.substring(0,6) != "<table")
	{
		repl = /&/g	
		strXML = strXML.replace(repl,"&amp;");
		
		//var rep2 = new RegExp("</ROOT>", gi);
		rep2 = /<\/root>/i
		
		var nacf = "<acfid>" + NAcfId + "</acfid></ROOT>"
		strXML = strXML.replace(rep2,nacf);
		rep2 = /<root>/i
		nacf = "<ROOT>"
		strXML = strXML.replace(rep2,nacf);
		//return(false)
		//strXML = strXML.replace(rep2,nacf);
		//strXML = strXML.replace("</ROOT>",nacf);
				
		xmlDoc.loadXML(strXML);
		xmlhttp.Open("POST", "RetornaCartaliberacao.asp?solid=" + sol_id, false);
		xmlhttp.Send(xmlDoc.xml);
	}
	
	
	strXML = xmlhttp.responseText;
	//var mod = window.showModalDialog("recebecartaliberacao.asp" ,strXML,"dialogHeight: 550px; dialogWidth: 750px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: yes; status: yes;");
	//window.showModalDialog('ConsultarFacilidades.asp?hdnAcao=Posicoes&strStsFac=L&cboLocalInstala='+cboLocalInstala.value+'&cboDistLocalInstala='+cboDistLocalInstala.value+'&cboProvedor='+cboProvedor.value+'&strPagina='+strPagina+'&cboSistema='+cboRede.value+'&cboPlataforma='+cboPlataforma.value ,objAryFac,"dialogHeight: 450px; dialogWidth: 570px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
	//window.showModalDialog("About:blank" ,,"dialogHeight: 450px; dialogWidth: 570px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
	objWindow = window.open("About:blank", null, "status=no,toolbar=no,menubar=no,location=no,resizable=Yes,scrollbars = Yes");
	objWindow.document.write(strXML);
	objWindow.document.close();
	//var mod = window.showModalDialog("recebecartaliberacao.asp" ,strXML,"dialogHeight: 550px; dialogWidth: 750px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: yes; status: yes;scroll:yes;");
	
		document.Form3.btnAlterar.disabled = false
		document.location.replace("DesativacaoLote.asp") 
}