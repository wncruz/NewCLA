function SelVelAcesso(obj)
{
	with (document.forms[1])
	{
		if (cboVelAcesso.value == "")
		{
			cboVelAcesso.value = obj.value
		}
	}
}

function ResgatarDesigServicoGravado(obj)
{
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarPadraoServico"
		if (obj == '[object]')
		{
			hdnCboServico.value = obj.value
		}
		else
		{
			hdnCboServico.value = obj + ",0"
		}	
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}	
}

function VerificaPropAcesso(strProp){
	var objNode = objXmlGeral.selectNodes("//xDados/Acesso")
	//Refaz a lista de Ids no IFRAME
	for (var intIndex=0;intIndex<objNode.length;intIndex++){
		strPropAcesso = RequestNodeAcesso(objXmlGeral,"rdoPropAcessoFisico",objNode[intIndex].childNodes[0].text)
		if (strPropAcesso == strProp){
			return true
		}
	}
	return false
}

function ResgatarGLA()
{
	var strPropAcesso = new String("")
	var blnAchouTER
	var blnAchouCLI
	var blnAchouEBT

	blnAchouTER = false
	blnAchouCLI = false
	blnAchouEBT = false
	
	//if (document.forms[1].hdnObrigaGla.value == "0") return false
	
	
	if (!IsEmpty(document.forms[0].txtRazaoSocial.value))
	{
		document.forms[1].hdnRazaoSocial.value = document.forms[0].txtRazaoSocial.value
		blnAchouTER = VerificaPropAcesso('TER')
		blnAchouCLI = VerificaPropAcesso('CLI')
		blnAchouEBT = VerificaPropAcesso('EBT')
		
		if (blnAchouTER || blnAchouCLI || blnAchouEBT){
			if (arguments.length > 0) {
				document.forms[1].hdnAcao.value = "ResgatarGLA&Gravar"
			}else{
				document.forms[1].hdnAcao.value = "ResgatarGLA"
			}	
		
			document.forms[1].target = "IFrmProcesso3"
			document.forms[1].action = "ProcessoSolic.asp"
			//alert("vou executar Pagina Processosolic")
			document.forms[1].submit()
			return true
		}	
		else{
			document.forms[1].hdntxtGLA.value = ""
			spnGLA.innerHTML = ""
			return false
		}
	}
	else
	{
		return false
	}
}

function EsconderTecnologia(intProcede)
{
	with (document.forms[1])
	{
		try{
			ReenviarSolicitacao(138,2)//limpa o acesso físico compartilhado
			divIDFis1.style.display = 'none'
			spnBtnLimparIdFis1.innerHTML =''
		}catch(e){}

		if (rdoPropAcessoFisico[1].checked)
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
		
		//Seleciona provedor embratel
		if (rdoPropAcessoFisico[1].checked)
		{
			SelProvedorEBT()
		}
		else
		{
			if (parseInt("0"+intProcede) != 1)
			{
				cboProvedor.disabled = false;
				cboProvedor.value = ""
				spnRegimeCntr.innerHTML = "<select name=cboRegimeCntr style=width:170px><Option></Option></select>"
				spnPromocao.innerHTML = "<select name=cboPromocao style=width:170px><Option></Option></select>"
			}
		}
	}
}

function ResgatarCidade(obj,intCid,objCNL)
{
	with (document.forms[1])
	{
		if (objCNL.value == "" ) return
		if (obj.value == "")
		{
			alert("Selecione a UF.")
			objCNL.value = ""
			if (intCid == 1) cboUFEnd.focus()
			else cboUFEndDest.focus()
			return
		}
				
		hdnAcao.value = "ResgatarCidadeCNL"
		hdnCNLNome.value = objCNL.name
		hdnUFAtual.value = obj.value
	
		if (intCid == 1){
			//if (hdnCNLAtual.value == objCNL.value) return
			hdnCNLAtual.value = objCNL.value
			hdnNomeCboCid.value = "EndCid"
			hdnNomeTxtCidDesc.value = "txtEndCidDesc"
			
		}else{
			if (hdnCNLAtual1.value == objCNL.value) return
			hdnCNLAtual1.value = objCNL.value
			hdnNomeCboCid.value = "EndCidDest"
			hdnNomeTxtCidDesc.value = "txtEndCidDescDest"
		}

		
		target = "IFrmProcesso"
		action = "ProcessoSolic.asp"
		submit()
	}
}

function ResgatarSev()
{
	with (document.forms[0])
	{
		if (txtNroSev.value != "")
		{
			hdnAcao.value = "ResgatarSev"
			target = "IFrmProcesso"
			action = "ProcessoSolic.asp"
			submit()
		}
		else
		{
			alert("Informe o número da sev.")
			txtNroSev.focus()
			return
		}
	}
}

function VerificarCidadeSev()
{
	with (document.forms[1])
	{
		if (txtEndCidDesc.value == "")
		{
			alert("Cidade da SEV não encontrada  para usuário atual.")
			return
		}
	}

}

function NovoCliente()
{
	document.forms[0].txtRazaoSocial.value = ""
	document.forms[0].txtSubContaSev.value = ""
	document.forms[0].txtNomeFantasia.value = ""
	document.forms[0].txtContaSev.value = ""
	document.forms[1].cboLogrEnd.value = ""
	document.forms[1].txtEnd.value = ""
	document.forms[1].txtNroEnd.value = ""
	document.forms[1].cboUFEnd.value = ""
	document.forms[1].txtEndCid.value = ""
	document.forms[1].txtCepEnd.value = ""
	document.forms[1].txtComplEnd.value = ""
	document.forms[1].txtBairroEnd.value = ""
	document.forms[1].txtContatoEnd.value = ""
	document.forms[1].txtTelEnd.value = ""
	document.forms[1].txtEndCidDesc.value = ""
	document.forms[1].txtCNPJ.value = ""
	document.forms[1].txtIE.value = ""
	document.forms[1].txtIM.value = ""
}

function SelecionarLocalConfig(obj)
{
	with (document.forms[2])
	{
		if (cboLocalConfig.value == "")
		{
			cboLocalConfig.value = obj.value 
		}
	}
}

function ResgatarUserCoordenacao(obj)
{
	if (obj.value != eval("document.forms[2].hdn"+obj.name+".value"))
	{
		with (document.forms[2])
		{
			eval("document.forms[2].hdn"+obj.name+".value = '"+obj.value+"'")
			hdnCoordenacaoAtual.value = obj.name
			hdnAcao.value = "ResgatarUserCoordenacao"
			target = "IFrmProcesso"
			action = "ProcessoSolic.asp"
			submit()
		}	
	}
}

function SistemaOrderEntry()
{
	with (document.forms[0])
	{
		if (cboSistemaOrderEntry.value == "")
		{
			if (!txtOrderEntry[0].readOnly)			
			{
				txtOrderEntry[0].readOnly = true
				txtOrderEntry[1].readOnly = true
				txtOrderEntry[2].readOnly = true
				txtOrderEntry[0].value = ""
				txtOrderEntry[1].value = ""
				txtOrderEntry[2].value = ""
			}	

		}	
		else
		{
			txtOrderEntry[0].readOnly = false
			txtOrderEntry[1].readOnly = false
			txtOrderEntry[2].readOnly = false
		}
	} 
}

function ProcurarCEP(intTipo,intObj)
{
	with (document.forms[1])
	{
		hdnAcao.value = "ProcurarCEP"
		hdnTipoCEP.value = intTipo
		if (intTipo == 1){ hdnCEP.value = txtCepEnd.value}
		else {hdnCEP.value = txtCepEndDest.value}
		if (hdnCEP.value.length < 5 && intObj == 1)
		{
			alert("CEP deve ser maior que cinco caracteres.")
			return
		}
		switch (intTipo)
		{
			case 1:
				txtNroEnd.value = ""
				txtComplEnd.value = ""	
				break
			case 2:
				txtNroEndDest.value = ""
				txtComplEndDest.value = ""	
				break
		}	
		target = "IFrmProcesso"
		action = "ProcessoSolic.asp"
		submit()
	}	
}

function ResgatarEstacaoDestino(objCNL,objCompl)
{
	with (document.forms[1])
	{
		if (objCNL.value != "" && objCompl.value != "")
		{
			//if (objCNL.value + objCompl.value != hdnEstacaoDestino.value)
			//{
				hdnEstacaoDestino.value = objCNL.value + objCompl.value
				hdnAcao.value = "ResgatarEstacaoDestino"
				target = "IFrmProcesso2"
				action = "ProcessoSolic.asp"
				submit()
			//}	
		}
	}
}

function ResgatarEstacaoOrigem(objCNL,objCompl)
{
	with (document.forms[1])
	{
		if (objCNL.value != "" && objCompl.value != "")
		{
			if (objCNL.value + objCompl.value != hdnEstacaoOrigem.value)
			{
				hdnEstacaoOrigem.value = objCNL.value + objCompl.value
				hdnAcao.value = "ResgatarEstacaoOrigem"
				target = "IFrmProcesso2"
				action = "ProcessoSolic.asp"
				submit()
			}	
		}
	}
}

function ResgatarEnderecoEstacao(obj)
{
	with (document.forms[2])
	{
		if (obj.value != ""){
			hdnAcao.value = "ResgatarEnderecoEstacao"
			hdnEstacaoAtual.value = obj.value
			target = "IFrmProcesso2"
			action = "ProcessoSolic.asp"
			submit()
		}
		else{
			spnContEndLocalInstala.innerHTML = ''
			spnTelEndLocalInstala.innerHTML = ''
		}	
	}
}

function ProcurarCliente()
{
	with (document.forms[0])
	{
		if (!IsEmpty(txtRazaoSocial.value))
		{
			
			hdnAcao.value = "ProcurarCliente"
			target = "IFrmProcesso"
			action = "ProcessoSolic.asp"
			submit()
			
		}
		else
		{
			alert("Informe a Razão Social.")
			return
		}	
	}	
}

function SelProvedorEBT()
{
	with (document.forms[1])
	{
		cboProvedor.value = 11
		hdnAcao.value = "ResgatarPromocaoRegime"
		hdnProvedor.value = 11
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
		cboProvedor.disabled = true
	}
}

function ResgatarPromocaoRegime(obj)
{
	with (document.forms[1])
	{
		if (obj == '[object]'){
			strValue = obj.value
		}
		else{
			strValue = obj
		}
		if (strValue != "")
		{
			hdnAcao.value = "ResgatarPromocaoRegime"
			hdnProvedor.value = strValue
			target = "IFrmProcesso"
			action = "ProcessoCla.asp"
			submit()
		}
		else
		{
			spnRegimeCntr.innerHTML = "<select name=cboRegimeCntr style=width:170px><Option></Option></select>"
			spnPromocao.innerHTML = "<select name=cboPromocao style=width:170px><Option></Option></select>"
		}	
	}
}

function ValidarNroContrato(obj)
{
	with (document.forms[0])
	{
		
		if (rdoNroContrato[0].checked)
		{
			if (!IsEmpty(obj.value) && obj.value != "")
			{
				if (!ValidarNTipo(obj,1,3,4,1,6,2,5,1,1,3,0,4,0,4)) return false
				
				if (obj.value.length > 4)
				{ 
					objAry = obj.value.split("-")
					if (objAry.length > 0)
					{
						switch (objAry[0].toUpperCase())
						{
							case "VES":
								if (!ValidarRangeCntr(obj,4,5,1,9))	return false
								break
							case "VEM":
								if (!ValidarRangeCntr(obj,4,5,1,11))return false
								break
							case "VMM":
								if (!ValidarRangeCntr(obj,4,5,1,5))	return false
								break
							default:
								alert("Tipo de Contrato Inválido(VES/VEM/VMM).")
								obj.value = ""
								return false
						}
					}
				}	
			}
		}
		else
		{
		}			
	}	
}

//Validar um determinado range
function ValidarRangeCntr(obj,intPosIni,intPosFim,intInterIni,intInterFim)
{
	var strValor = new String("")
	if (obj == '[object]')
		var checkStr = new String(obj.value);
	else
		var checkStr =  new String(obj);
	
	for (var intIndex = 0;intIndex<checkStr.length;intIndex++)
	{
		if (intIndex > parseInt(intPosIni-1) && intIndex < parseInt(intPosFim+1))
		{
			strValor += checkStr.charAt(intIndex);
		}
	}		
	
	if (parseInt(strValor)	< parseInt(intInterIni) || parseInt(strValor)	> parseInt(intInterFim))
	{
		alert("Valor fora do intervalo " + intInterIni + " a " + intInterFim + ".")
		if (obj == '[object]') obj.value = obj.value.substring(0,intPosIni)
		return false
	}
	else
	{
		return true
	}
}

function ProcurarIDFis(intID)
{
	document.forms[1].hdnRazaoSocial.value = document.forms[0].txtRazaoSocial.value
	with (document.forms[1])
	{
		switch (intID)
		{
			case 1:
				hdnIdAcessoFisico.value = ""
				target = "IFrmIDFis1"
				action = "AcessoCompartilhadoSol.asp?intEnd=1&strtipo=T"
				submit()
				break
			case 2:	
				hdnIdAcessoFisico1.value = ""
				target = "IFrmIDFis2"
				action = "AcessoCompartilhadoSol.asp?intEnd=2"
				submit()
				break
			case 3: //Editando o Id'Físico
				target = "IFrmIDFis1"
				action = "AcessoCompartilhadoSol.asp?intEnd=1"
				submit()
				break
		}
	}	
}


function SelIDFisComp(obj,intEnd,IdAcFis)
{
	with (document.forms[1])
	{
		hdnIdAcessoFisico.value = obj.value 
		hdnAcfId.value	= IdAcFis
		switch (parseInt(intEnd))
		{
			case 1:
				if (rdoPropAcessoFisico[0].checked)
				{
					hdnPropIdFisico.value = rdoPropAcessoFisico[0].value
				}
				if (rdoPropAcessoFisico[1].checked)
				{
					hdnPropIdFisico.value = rdoPropAcessoFisico[1].value
				}
				if (rdoPropAcessoFisico[2].checked)
				{
					hdnPropIdFisico.value = rdoPropAcessoFisico[2].value
				}
				hdnAecIdFis.value = obj.Aec_IdFis
				break
			case 2:	
				if (rdoPropAcessoFisico[0].checked)
				{
					hdnPropIdFisico1.value = rdoPropAcessoFisico[0].value
				}
				if (rdoPropAcessoFisico[1].checked)
				{
					hdnPropIdFisico1.value = rdoPropAcessoFisico[1].value
				}
				if (rdoPropAcessoFisico[2].checked)
				{
					hdnPropIdFisico1.value = rdoPropAcessoFisico[3].value
				}
				break
		}	

		switch (parseInt(intEnd))
		{
			case 1:
				hdnIdAcessoFisico.value=obj.value
				//hdnPropIdFisico.value = obj.prop
				//Verifica se o usuário quer compartilhar ou não o ID Físico selecionado 
				hdnCompartilhamento.value = "0"
				hdnChaveAcessoFis.value = IdAcFis 				
				hdnAcao.value = "AutorizarCompartilhamento"
				hdnSubAcao.value = "IdFisEndInstala"
				target = "IFrmProcesso"
				action = "ProcessoCla.asp"
				submit()

				break
			case 2:	
				hdnIdAcessoFisico1.value=obj.value
				//hdnPropIdFisico1.value = obj.prop
				//Verifica se o usuário quer compartilhar ou não o ID Físico selecionado
				hdnCompartilhamento1.value = "0"
				hdnChaveAcessoFis.value = IdAcFis 				
				hdnAcao.value = "AutorizarCompartilhamento"
				hdnSubAcao.value = "IdFisEndPtoInterme"
				target = "IFrmProcesso"
				action = "ProcessoCla.asp"
				submit()
				break
		}
	}	
}

function ReenviarSolicitacao(intRetASP,intRetJS)
{
	with (document.forms[1])
	{
		switch (parseInt(intRetASP))
		{
			case 138: //Endereco de instalação
				switch (parseInt(intRetJS))
				{
					case 1: //Aceito
						hdnCompartilhamento.value = "1"
						//Resgatar demais campos do Id Físico
						//hdnAecIdFis já foi populado no onClick do Radio button função SelIDFisComp()
						hdnAcao.value = "ResgatarAcessoFisComp"
						target = "IFrmProcesso"
						action = "ProcessoSolic.asp"
						submit()
						break
					
					case 2: // Não Aceito
						hdnIdAcessoFisico.value = ""
						hdnPropIdFisico.value  = ""
						hdnCompartilhamento.value = "0"
						btnIDFis1.focus()
						limparIDFisico(1)
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

function limparIDFisico(intID)
{
	switch (parseInt(intID))
	{
		case 1:
			//parent.IFrmIDFis1.
			try{
			with (IFrmIDFis1.document.forms[0])
			{
				if (rdoIDFis1 == '[object]')
				{
					rdoIDFis1.checked = false
					try{
						parent.document.Form2.hdnNovoPedido.value = ""
					}catch(e){}	
				}
				for (var intIndex=0;intIndex<rdoIDFis1.length;intIndex++)
				{
					rdoIDFis1[intIndex].checked = false	
					try{
						parent.document.Form2.hdnNovoPedido.value = ""
					}catch(e){}		
				}
			}	
			}
			catch(e){}
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

function Gravar()
{
	with (document.forms[0])
	{
		if (!ValidarCampos(txtRazaoSocial,"Nome do Cliente/Razão Social")) return
		if (!ValidarCampos(txtNomeFantasia,"Nome Fantasia")) return
		if (!ValidarCampos(txtContaSev,"Conta Corrente")) return
		if (!ValidarCampos(txtSubContaSev,"Sub Conta")) return
		
//LPEREZ - 24/10/2005					
/*		
		if (cboGrupo.value != "")
		{
			if(!ValidarCampos(cboOrigemSol,"Origem Solicitação")) return	
		}else{
			cboOrigemSol.value = null;
		}
*/				
//LP

		if (!IsEmpty(cboSistemaOrderEntry.value))
		{
			if (IsEmpty(txtOrderEntry[0].value) || IsEmpty(txtOrderEntry[1].value) || IsEmpty(txtOrderEntry[2].value))
			{ 
				alert("Order Entry incompleta.Favor preencher Sistema/Ano/Nro/Item.")
				cboSistemaOrderEntry.focus()
				return
			}
			else
			{	
				if (parseInt(txtOrderEntry[0].value) < 1964)
				{
					alert("Ano da Order Entry inválido. Ano dever ser maior ou igual a 1964.")
					return
				}
				else
				{
					hdnOrderEntry.value = cboSistemaOrderEntry.value + txtOrderEntry[0].value + txtOrderEntry[1].value + txtOrderEntry[2].value
				}	
			}	
		}	
		if (!MontarDesigServico()) return
		if (!ValidarCampos(cboServicoPedido,"Serviço")) return
		if (!ValidarCampos(cboVelServico,"Velocidade do Serviço")) return

		if (!ValidarCampos(txtNroContrServico,"Nº do Contrato Serviço")) return

		//Designação do Acesso Principal
		hdnDesigAcessoPri.value = ""
		if (txtDesigAcessoPri.value != "" && txtDesigAcessoPri.value.length < 7)
		{
			alert("Designação do Acesso Principal(678) fora de padrão 678N7.")
			txtDesigAcessoPri.focus()
			return
		}
		
		if (txtDesigAcessoPri.value != "" && txtDesigAcessoPri.value.length == 7)
		{
			hdnDesigAcessoPri.value = txtDesigAcessoPri0.value + txtDesigAcessoPri.value
		}

		if (!ValidarTipoInfo(txtDtIniTemp,1,"Data Início Temporário")) return;
		if (!ValidarTipoInfo(txtDtFimTemp,1,"Data Fim Temporário")) return;
		if (!ValidarTipoInfo(txtDtDevolucao,1,"Data Devolução Temporário")) return;
		if (!ValidarCampos(txtDtEntrAcesServ,"Data Desejada de Entrega do Acesso ao Serviço")) return;
		if (!ValidarTipoInfo(txtDtEntrAcesServ,1,"Data Desejada de Entrega do Acesso ao Serviço")) return ;
		if (!ValidarTipoInfo(txtDtPrevEntrAcesProv,1,"Data Prevista de Entrega do Acesso pelo Provedor")) return ;

	}
	//Verifica se tem acesso adicionado
	var objNode = objXmlGeral.selectNodes("//xDados/Acesso")
	if (objNode.length == 0){
		//alert("Informações do Acesso são Obrigatórias.")
		if (!AdicionarAcessoLista(true)) return
		if (document.Form2.rdoPropAcessoFisico[1].checked)
		{
			ContinuarGravacao()
		
		}else
		{
			//alert('vou resgatar gla')
			ResgatarGLA(true)
		
		}	
	}else
	{   
		ContinuarGravacao()
		
	}

}

function ContinuarGravacao()
{
	//alert('davif vou ContinuarGravacao()')

	var objNodeAux = objXmlGeral.selectNodes("//Acesso[cboTipoPonto='I']")
	if (objNodeAux.length == 0){
		alert("É necessário pelo menos um ponto de instalação.")
		return
	}

	with (document.forms[1]){
		
		//GLA é um campo obrigarório para TER/CLLI
		var blnAchou = VerificaPropAcesso('TER')
		if (blnAchou && IsEmpty(hdntxtGLA.value) && hdnObrigaGla.value == "1"){
			alert("GLA é um campo obrigatório.")
			return
		}	
		blnAchou = VerificaPropAcesso('CLI')
		if (blnAchou && IsEmpty(hdntxtGLA.value) && hdnObrigaGla.value == "1"){
			alert("GLA é um campo obrigatório.")
			return
		}	
	}	

	with (document.forms[2]) //Form3
	{
		
		var blnAchouSatelite = false
		if (IsEmpty(hdntxtGICL.value)){alert('GIC-L é um campo obrigatório.');return}
		if (!ValidarCampos(cboLocalEntrega,"Estação do Local de Entrega")) return
		if (!ValidarCampos(cboLocalConfig,"Estação do Local de Configuração")) return

		var objNode = objXmlGeral.selectNodes("//xDados/Acesso")
		if (objNode.length == 1){
			for (var intIndex=0;intIndex<objNode.length;intIndex++){
				var intChave = objNode[intIndex].childNodes[0].text
				var intTec = RequestNodeAcesso(objXmlGeral,"cboTecnologia",intChave)
				if (intTec == 4) blnAchouSatelite = true
			}
		}	
		if (!blnAchouSatelite){
			//if (!ValidarCampos(cboInterfaceEbt,"Interface EBT")) return
		}	
		if (!ValidarCampos(cboOrgao,"Orgão")) return
		if (IsEmpty(hdntxtGICN.value)){alert('GIC-N é um campo obrigatório.');txtGICN.focus();return}

	}

	document.Form4.hdnXml.value = ""
	for (var intIndex=0;intIndex<document.Form1.elements.length;intIndex++){
		var elemento = document.Form1.elements[intIndex];
		if (elemento.type != 'button'){
			AdicionarNode(objXmlGeral,elemento.name,elemento.value)
		}	
	}
	//Outros campos
	AdicionarNode(objXmlGeral,"hdnServico",document.Form1.cboServicoPedido.value)
	if (document.Form3.hdntxtGICL.value != "")	AdicionarNode(objXmlGeral,"hdntxtGICL",document.Form3.hdntxtGICL.value)
	if (document.Form2.hdntxtGLA.value != "")	AdicionarNode(objXmlGeral,"hdntxtGLA",document.Form2.hdntxtGLA.value)
	if (document.Form2.hdntxtGLAE.value != "")	AdicionarNode(objXmlGeral,"hdntxtGLAE",document.Form2.hdntxtGLAE.value)

	//Tipo do Contrato
	if (document.forms[0].rdoNroContrato[0].checked) AdicionarNode(objXmlGeral,"intTipoContrato",1)
	if (document.forms[0].rdoNroContrato[1].checked) AdicionarNode(objXmlGeral,"intTipoContrato",2)
	if (document.forms[0].rdoNroContrato[2].checked) AdicionarNode(objXmlGeral,"intTipoContrato",3)

	for (var intIndex=0;intIndex<document.Form3.elements.length;intIndex++){
		var elemento = document.Form3.elements[intIndex];
		if (elemento.type != 'button'){
			AdicionarNode(objXmlGeral,elemento.name,elemento.value)
		}	
	}
	//Verifica se o usuário editou e não atualizou a lista
	if (AcessoAlteradoNaoAtualizado()){
		var intRet=alertbox('As informações do acesso atualmente editado foram alteradas. Deseja atualizá-la e prosseguir com a gravação da solicitação?','Sim','Não','Sair')
		switch (parseInt(intRet))
		{
			case 1:
				if (!AdicionarAcessoLista(true)) return
				break
			case 3:
				return
				break
		}		
	}

	with (document.Form4){
		if (hdnTipoAcao.value == "Alteracao")
		{
			/*
			var objNodeComp = objXmlGeral.selectNodes("//Acesso[hdnCompartilhamento=0]")
			//Deve ser apresentada a tela de taxa de serviço
			if (hdnVelIdServicoOld.value != document.Form1.cboVelServico.value && objNodeComp.length > 0)
			{
				var objNodeComp = objXmlGeral.selectNodes("//Acesso[hdnCompartilhamento=1]")
				var strIdAcessoFisComp = new String("")
				for (var intIndex=0;intIndex<objNodeComp.length;intIndex++)
				{
					strIdAcessoFisComp += objNodeComp[intIndex].getElementsByTagName("Acf_Id").item(0).text
					if (intIndex<parseInt(objNodeComp.length-1)) strIdAcessoFisComp += ","
				}	
				
				if (strIdAcessoFisComp != "")
				{
					var strAry = hdnIdAcessoLogico.value+","+hdnSolId.value+","+hdnTipoProcesso.value + "," + strIdAcessoFisComp
					var objAry = strAry.split(",")
				}else
				{
					var objAry = new Array(hdnIdAcessoLogico.value,hdnSolId.value,hdnTipoProcesso.value)
				}	
				//var objAry = new Array(hdnIdAcessoLogico.value,hdnSolId.value,hdnTipoProcesso.value)
				var objXmlTaxaServico = window.showModalDialog('AlteracaoTaxaServico.asp',objAry,'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
				var objNode = objXmlTaxaServico.selectNodes("//Ret")
				if (objNode.length > 0 )
				{
					if (objNode[0].childNodes[0].text != 999)
					{
						AddTaxaServico(objXmlGeral,objXmlTaxaServico)
					}
					else
					{
						return
					}	
				}else
				{
					return
				}
			}
			*/
			hdnAcao.value = "Alteracao"
		}else
		{
			hdnAcao.value = "GravarSolicitacao"
		} 
		
		hdnXml.value = objXmlGeral.xml
		target = "IFrmProcesso"
		action = "ProcessoSolic.asp"
		
		//alert('davif antes ProcessoSolic()')
		//document.Form3.btnGravar.disabled=true; //PRSS - 05/01/2006
		submit()
	}	

}

function NovoPedido(obj)
{
	if (obj.checked){
		document.Form2.hdnNovoPedido.value = 1
	}else{
		document.Form2.hdnNovoPedido.value = ""
	}	
}

function AdicionarAcessoLista(){

	
	//alert('adicionar acesso a lista')
	with (document.forms[1]) //Form2
	{
		var blnAchou = false
		for (var intIndex=0;intIndex<rdoPropAcessoFisico.length;intIndex++)
		{
			if (rdoPropAcessoFisico[intIndex].checked)
			{
				blnAchou = true
			}
		}
		if (!blnAchou)
		{
			alert("Proprietário do Acesso Físico é um Campo Obrigatório.")	
			rdoPropAcessoFisico[0].focus()
			return false
		}

		if ((rdoPropAcessoFisico[1].checked) && (cboTecnologia.value == ""))
		{
			cboTecnologia.focus();
			return false
		}
		
		if (cboTecnologia[cboTecnologia.selectedIndex].innerText == "RADIO")
		{
			try{
				if (cboTipoRadio.value != ""){
					if (!ValidarCampos(cboVersaoRadio,"Versao do Radio")) return false
				}
				else
				{
					if (cboVersaoRadio.value != ""){
						alert('A Versão do Rádio não deve ser preenchida sem o preenchimento do Tipo de Rádio.')
						return false
					}
				}
				
			}
			catch(e){}
		}
		
		
		var objNodeAux = objXmlGeral.selectNodes("//Acesso[cboTipoPonto='I' && TipoAcao != 'R']")
		if (objNodeAux.length == 1 && cboTipoPonto.value == 'I'){
			if (objNodeAux[0].childNodes[0].text != hdnIntIndice.value && !ValidarPontoInstalacao()){
				alert("Não é possível adicionar mais que um Ponto de Instalação para endereços diferentes.")
				return false
			}	
		}

		if (!ValidarCampos(cboVelAcesso,"Velocidade do Acesso Físico")) return false

		//if (rdoPropAcessoFisico[1].checked && parseInt("0"+cboTecnologia.value) != 3)
		//{
		//	if (!ValidarCampos(cboRegimeCntr,"Para EBT com a Tecnologia Difente de ADE o Prazo de Contratação de //Acesso")) return false
//		}

		var strVel = cboVelAcesso[cboVelAcesso.selectedIndex].text
		if (strVel == "2M" || strVel == "34M" || strVel == "155M" || strVel == "622M"){
			if (!ValidarCampos(cboTipoVel,"Para as Velocidades de 2M/34M/155M/622M o Tipo de Velocidade")) return false
		}
		if (!ValidarCampos(txtQtdeCircuitos,"Quantidade de Circuitos")) return false
		if (txtQtdeCircuitos.value == 0){alert("Quantidade de Circuitos dever ser maior ou igual a um.");return false}
		if (!ValidarCampos(cboProvedor,"Provedor")) return false
		//Endereço do Acesso Físico
		//Endereço Origem
		if (!ValidarCampos(cboUFEnd,"Estado Origem")) return false
		if (!ValidarCampos(txtEndCid,"Cidade Origem")) return false
		if (!ValidarCampos(cboLogrEnd,"Logradouro Origem")) return false
		if (!ValidarCampos(txtEnd,"Nome do Logradouro Origem")) return false
		if (!ValidarCampos(txtNroEnd,"Número Origem")) return false
		if (!ValidarCampos(txtBairroEnd,"Bairro Origem")) return false
		if (!ValidarCampos(txtCepEnd,"CEP Origem")) return false
		if (!ValidarCampos(txtEndEstacaoEntrega,"Endereço de entrega")) return false
		
		if (!ValidarTipoInfo(txtCepEnd,2,"CEP Origem")) return false
		//Endereço Destino
		/*
		if (!ValidarCampos(cboUFEndDest,"Estado Destino")) return false
		if (!ValidarCampos(txtEndCidDest,"Cidade Destino")) return false
		if (!ValidarCampos(cboLogrEndDest,"Logradouro Destino")) return false
		if (!ValidarCampos(txtEndDest,"Nome do Logradouro Destino")) return false
		if (!ValidarCampos(txtNroEndDest,"Número Destino")) return false
		if (!ValidarCampos(txtBairroEndDest,"Bairro Destino")) return false
		if (!ValidarCampos(txtCepEndDest,"CEP Destino")) return false
		if (!ValidarTipoInfo(txtCepEndDest,2,"CEP Destino")) return false
		txtEndEstacaoEntrega*/

		if (!ValidarCampos(txtContatoEnd,"Contato")) return false
		if (!ValidarCampos(txtTelEndArea,"Telefone")) return false
		if (txtTelEndArea.value.length != 2)
		{ 
			alert("Código de area do telefone inválido.")
			txtTelEndArea.focus()
			return false
		}	
		if (!ValidarCampos(txtTelEnd,"Telefone")) return false
		if (!ValidarCampos(txtCNPJ,"CNPJ do Endereço de Instalação")) return false
		if (!VerificarCpfCnpj(txtCNPJ,2)) return false
		
		//if (rdoPropAcessoFisico[1].checked && (parseInt("0"+cboTecnologia.value) == 1 || parseInt("0"+cboTecnologia.value) == 2 || parseInt("0"+cboTecnologia.value) == 4))
		//{
		//	if (!ValidarCampos(txtCodSAP,"Para tecnologia igual a Rádio/Fibra Óptica/Satélite o código SAP")) return false;
		//}	
		//if (rdoPropAcessoFisico[1].checked && parseInt("0"+cboTecnologia.value) == 1 && hdnIntIndice.value == "")
		//{
		//	if (!ValidarCampos(txtNroPI,"Para tecnologia igual a Rádio e novo acesso\nfísico o número PI")) return false;
		//}	

		

		if (!ValidarCampos(cboTipoPonto,"Tipo do Ponto(Instalação/Intermediário)")) return false

		if (rdoPropAcessoFisico[1].checked && parseInt("0"+cboTecnologia.value) != 4)
		{
			if (!ValidarCampos(cboInterFaceEnd,"Interface")) return false
		}
		else
		{
			if (rdoPropAcessoFisico[0].checked || rdoPropAcessoFisico[2].checked)
			{
				if (!ValidarCampos(cboInterFaceEnd,"Interface")) return false
				if (!ValidarCampos(cboInterFaceEndFis,"Interface")) return false
			}
		}
		
//		if (rdoPropAcessoFisico[1].checked && (parseInt("0"+cboTecnologia.value) == 1))
//		{
//			//Verificar se existe o compartinhamento de ID'Físico
//			if (cboTipoPonto.value = "I"){
//				if (!ValidarCampos(txtCNLSiglaCentroCli,"CNL da Sigla do Centro do Cliente")) return false;
//				if (!ValidarCampos(txtComplSiglaCentroCli,"Complemento da Sigla do Centro do Cliente")) return false;
//				//if (!ValidarCampos(txtCNLSiglaCentroCliDest,"Destino")) return false;
//				//if (!ValidarCampos(txtComplSiglaCentroCliDest,"Complemento do Destino")) return false;
//			}else
//			{
//				if (!ValidarCampos(txtCNLSiglaCentroCli,"Origem")) return false;
//				if (!ValidarCampos(txtComplSiglaCentroCli,"Complemento da Origem")) return false;
//				//if (!ValidarCampos(txtCNLSiglaCentroCliDest,"Destino")) return false;
//				//if (!ValidarCampos(txtComplSiglaCentroCliDest,"Complemento do Destino")) return false;
//			}	
//		}	

		var blnMessage = false
		if (arguments.length>0){
			blnMessage = arguments[0]
			intRet = 1
		}
		if (!blnMessage){
			var intRet=alertbox('Deseja permanecer com os dados?','Sim','Não','Sair')
		}	
		switch (parseInt(intRet))
		{
			case 1:
				xmlUpd(false)
				updOrdemXml()
				AtualizarLista()
				//Para o compartilhamento
				try{
					ReenviarSolicitacao(138,2)//limpa o acesso físico compartilhado
					divIDFis1.style.display = 'none'
					spnBtnLimparIdFis1.innerHTML =''
					AdicionarNode(objXmlGeral,"hdnCompartilhamento",document.Form2.hdnCompartilhamento.value)
					AdicionarNode(objXmlGeral,"hdnIdAcessoFisico",document.Form2.hdnIdAcessoFisico.value)
					AdicionarNode(objXmlGeral,"hdnNovoPedido",document.Form2.hdnNovoPedido.value)
					Form2.hdnNovoPedido.value = ""
				}catch(e){}	
				break		
			case 2:
				//@@
				xmlUpd(true)
				updOrdemXml()
				AtualizarLista()
				TipoOrigem("T")
				//Para o compartilhamento
				try{
					ReenviarSolicitacao(138,2)//limpa o acesso físico compartilhado
					divIDFis1.style.display = 'none'
					spnBtnLimparIdFis1.innerHTML =''
					AdicionarNode(objXmlGeral,"hdnCompartilhamento",document.Form2.hdnCompartilhamento.value)
					AdicionarNode(objXmlGeral,"hdnIdAcessoFisico",document.Form2.hdnIdAcessoFisico.value)
					AdicionarNode(objXmlGeral,"hdnNovoPedido",document.Form2.hdnNovoPedido.value)
					Form2.hdnNovoPedido.value = ""
				}catch(e){}	
				break
		}

		DesabilitarCamposAlt(false,"")

		rdoPropAcessoFisico[0].focus()

		ResgatarGLA()
	}
	document.Form2.btnAddAcesso.value = "Adicionar"
	return true
}	

function DesabilitarCamposAlt(blnAcao,intChave)
{	
	try{
		var	objNode = objXmlGeral.selectNodes("//xDados/Acesso[intIndice="+parseInt(intChave)+"]")

		if ((document.forms[1].hdnTipoProcesso.value == "3") || (document.forms[1].hdnTipoProcesso.value == "1"))
		{
			if (intChave != "")
			{
				if (objNode.length > 0)
				{
					var strTipoAcao = objNode[0].getElementsByTagName("TipoAcao").item(0).text
					if (strTipoAcao != "N")
					{
						document.forms[1].rdoPropAcessoFisico[0].disabled = blnAcao
						document.forms[1].rdoPropAcessoFisico[1].disabled = blnAcao
						document.forms[1].rdoPropAcessoFisico[2].disabled = blnAcao
						document.forms[1].cboProvedor.disabled = blnAcao
					}else{
						if (intComp == "1")
						{
							document.forms[1].rdoPropAcessoFisico[0].disabled = blnAcao
							document.forms[1].rdoPropAcessoFisico[1].disabled = blnAcao
							document.forms[1].rdoPropAcessoFisico[2].disabled = blnAcao
							document.forms[1].cboProvedor.disabled = blnAcao
						}	
					}	
				}	
			}else{
				document.forms[1].rdoPropAcessoFisico[0].disabled = blnAcao
				document.forms[1].rdoPropAcessoFisico[1].disabled = blnAcao
				document.forms[1].rdoPropAcessoFisico[2].disabled = blnAcao
				document.forms[1].cboProvedor.disabled = blnAcao
			}
		}
	}catch(e){}
}

function RemoverAcessoLista(){

	with (document.forms[1]){
		DesabilitarCamposAlt(false,hdnIntIndice.value)
		if (hdnIntIndice.value != ""){
			RemoverAcesso(hdnIntIndice.value,true)
			LimparInfoAcesso()
			updOrdemXml()
			AtualizarLista()
			rdoPropAcessoFisico[0].focus()			
			ResgatarGLA()
			try{
				ReenviarSolicitacao(138,2)//limpa o acesso físico compartilhado
				divIDFis1.style.display = 'none'
				spnBtnLimparIdFis1.innerHTML =''
			}catch(e){}
			btnAddAcesso.value = "Adicionar"
		}
		else{
			alert("Selecione um item para remover em \"Acesso Adicionados\".")
			return
		}
	}
}

function EditarAcessoLista(objChave){
	with (document.Form2){
		hdnChaveAcessoFis.value = objChave.value
		hdnTecnologia.value = RequestNodeAcesso(objXmlGeral,"cboTecnologia",objChave.value)
		hdnVelAcessoFisSel.value = RequestNodeAcesso(objXmlGeral,"cboVelAcesso",objChave.value)
		hdnProvedor.value = RequestNodeAcesso(objXmlGeral,"cboProvedor",objChave.value)
		strPropAcesso = RequestNodeAcesso(objXmlGeral,"rdoPropAcessoFisico",objChave.value)
		TipoOrigem(RequestNodeAcesso(objXmlGeral,"cboTipoPonto",objChave.value))
		document.Form2.btnAddAcesso.value = "Alterar"

		try{
			ReenviarSolicitacao(138,2)//limpa o acesso físico compartilhado
			divIDFis1.style.display = 'none'
			spnBtnLimparIdFis1.innerHTML =''
		}catch(e){}

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
		var strVel = RequestNodeAcesso(objXmlGeral,"cboVelAcessoText",objChave.value)
		if (strVel == "2M" || strVel == "34M" || strVel == "155M" || strVel == "622M"){
			divTipoVel.style.display = ""
		}else{
			cboTipoVel.value = ""
			divTipoVel.style.display = "none"
		}	
		//Sinaliza que o id físico esta compartinhado par uma edição
		hdnNodeCompartilhado.value = RequestNodeAcesso(objXmlGeral,"hdnCompartilhamento",objChave.value)
		hdnCompartilhamento.value = RequestNodeAcesso(objXmlGeral,"hdnCompartilhamento",objChave.value)
		hdnNovoPedido.value = RequestNodeAcesso(objXmlGeral,"hdnNovoPedido",objChave.value)
		
		
		
		RetornaCboTipoRadio('RADIO', RequestNodeAcesso(objXmlGeral,"cboTecnologia",objChave.value) ,RequestNodeAcesso(objXmlGeral,"cboTipoRadio",objChave.value) , RequestNodeAcesso(objXmlGeral,"cboVersaoRadio",objChave.value))
		
		
		hdnAcao.value = "EditarAcessoFisico"
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
}	 

function LimparInfoAcesso(){

	try
	{
		spnTipoRadio.innerHTML = ""
		parent.tdRadio.innerHTML = ""
		document.forms[1].rdoPropAcessoFisico[0].disabled = false
		document.forms[1].rdoPropAcessoFisico[1].disabled = false
		document.forms[1].rdoPropAcessoFisico[2].disabled = false
		document.forms[1].cboProvedor.disabled = false
		document.Form2.btnAddAcesso.value = "Adicionar"
		document.Form2.hdnCNLAtual.value = ""
		document.Form2.hdnCNLAtual1.value = ""
		document.Form2.hdnEstacaoOrigem.value = ""
		document.Form2.hdnEstacaoDestino.value = ""
		
	}catch(e){}	

	TipoOrigem("T")

	for (var intIndex=0;intIndex<document.Form2.elements.length;intIndex++)
	{
		var elemento = document.Form2.elements[intIndex];
		if (elemento.type != 'button' && elemento.type != 'hidden' && elemento.type != 'radio')	elemento.value = ""
		if (elemento.type == 'radio')
		{
			elemento.checked = false
		}
	}
	for (var intIndexII=0;intIndexII<IFrmAcessoFis.document.Form1.elements.length;intIndexII++)
	{
		var elemento = IFrmAcessoFis.document.Form1.elements[intIndexII];
		if (elemento.type == 'radio')
		{
			elemento.checked = false
		}
	}
	try{
		ReenviarSolicitacao(138,2)//limpa o acesso físico compartilhado
		divIDFis1.style.display = 'none'
		spnBtnLimparIdFis1.innerHTML =''
	}catch(e){}

	document.Form2.cboProvedor.disabled = false 
	document.Form2.hdnIntIndice.value = "" //Chave para poder adicionar um novo
	document.Form2.txtQtdeCircuitos.value = 1
	divTecnologia.style.display = "none"
	document.Form2.cboTipoVel.value = ""
	divTipoVel.style.display = "none"
	document.Form2.rdoPropAcessoFisico[0].focus()
	parent.spnListaIdFis.innerHTML = ""
}

//Atualiza a ordem de entrada para o xml
function updOrdemXml(){
 var objNode = objXmlGeral.selectNodes("//Acesso")
 if (objNode.length > 0){
	for (var intIndex=0;intIndex<objNode.length;intIndex++){
		intChave = objNode[intIndex].childNodes[0].text
		AdicionarNodeAcesso(objXmlGeral,"intOrdem",parseInt(intIndex+1),intChave)
	}
 }
}

function ResgatarTecVel(){
	with (document.forms[1]){
		hdnAcao.value = "ResgatarTecVel"
		target = "IFrmProcesso2"
		action = "ProcessoCla.asp"
		hdnVelAcessoFisSel.value = ""
		if (rdoPropAcessoFisico[0].checked || rdoPropAcessoFisico[2].checked){
			cboTecnologia.value = ""
			submit()
		}else{
			if (cboTecnologia.value != ""){
				submit()				
			}else{
				spnVelAcessoFis.innerHTML = "<Select name=cboVelAcesso style='width:150px'></select>"
			}
		}
		cboVelAcesso.value = ""
		cboTipoVel.value = ""
		divTipoVel.style.display =  "none"
	}
}

function MostrarTipoVel(obj){
	var strVel = obj[obj.selectedIndex].text
	if (strVel == "2M" || strVel == "34M" || strVel == "155M" || strVel == "622M"){
		divTipoVel.style.display = ""
	}else{
		document.Form2.cboTipoVel.value = ""
		document.Form2.cboTipoVel.selectedIndex = 0
		divTipoVel.style.display = "none"
	}	
	
}

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
		document.Form2.txtQtdeCircuitos.value = 1
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

	DesabilitarCamposAlt(true,intChave)

	parent.spnListaIdFis.innerHTML = strAcessoIdFis
	
}

function TipoOrigem(strTipoOrig)
{
	if (strTipoOrig == "I")
	{
		spnOrigem.innerHTML = "&nbsp;&nbsp;&nbsp;Sigla do Centro do Cliente"
	}else{
		spnOrigem.innerHTML = "&nbsp;&nbsp;&nbsp;Sigla Estação Origem"
	}
}


function RetornaCboTipoRadio(strTec,TecID, TrdID, strVersao)
{
	if (strTec != "RADIO") {
		spnTipoRadio.innerHTML = ""
		parent.tdRadio.innerHTML  = ""	
		return 
	}
	

	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
	var strXML
	
	if (TrdID == "" ) TrdID = 0 
	
	strXML = "<root>"
	strXML = strXML + "<tecid>" +  TecID + "</tecid>"
	strXML = strXML + "<trdid>" +  TrdID + "</trdid>"
	strXML = strXML + "<funcao></funcao>"
	strXML = strXML + "<versao>" +  strVersao + "</versao>"
	strXML = strXML + "</root>" 
		
	xmlDoc.loadXML(strXML);
	xmlhttp.Open("POST","RetornaTipoRadio.asp" , false);
	xmlhttp.Send(xmlDoc.xml);
		
	strXML = xmlhttp.responseText;
	
	if (strXML != ""){
		parent.tdRadio.innerHTML  = "&nbsp;&nbsp;&nbsp;Tipo de Radio"	
	}
	else{
		parent.tdRadio.innerHTML  = ""	
	}
	spnTipoRadio.innerHTML = strXML
	
}
