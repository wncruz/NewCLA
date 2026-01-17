function AdicionarAcessoListaAlteracao(){

	with (document.forms[1]) //Form2
	{
		if (hdnIntIndice.value == "")
		{
			alert("Favor editar um acesso para alterá-lo")
			return
		}
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
			return
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
		if ((rdoPropAcessoFisico[1].checked) && (cboTecnologia.value == ""))
		{
			alert("Tecnologia é um Campo Obrigatório.")	
			cboTecnologia.focus();
			return
		}

		var objNodeAux = objXmlGeral.selectNodes("//Acesso[cboTipoPonto='I']")
		if (objNodeAux.length == 1 && cboTipoPonto.value == 'I'){
			if (objNodeAux[0].childNodes[0].text != hdnIntIndice.value){
				alert("Não é possível adicionar mais que um Ponto de Instalação.")
				return
			}	
		}

		if (!ValidarCampos(cboVelAcesso,"Velocidade do Acesso Físico")) return

		var strVel = cboVelAcesso[cboVelAcesso.selectedIndex].text
		if (strVel == "2M" || strVel == "34M" || strVel == "155M" || strVel == "622M"){
			if (!ValidarCampos(cboTipoVel,"Para as Velocidades de 2M/34M/155M/622M o Tipo de Velocidade")) return
		}
		if (!ValidarCampos(txtQtdeCircuitos,"Quantidade de Circuitos")) return
		if (txtQtdeCircuitos.value == 0){alert("Quantidade de Circuitos dever ser maior ou igual a um.");return}
		if (!ValidarCampos(cboProvedor,"Provedor")) return
		//Endereço do Acesso Físico
		if (!ValidarCampos(cboUFEnd,"Estado")) return
		if (!ValidarCampos(txtEndCid,"Cidade")) return
		if (!ValidarCampos(cboLogrEnd,"Logradouro")) return
		if (!ValidarCampos(txtEnd,"Nome do Logradouro")) return
		if (!ValidarCampos(txtNroEnd,"Número")) return
		if (!ValidarCampos(txtBairroEnd,"Bairro")) return
		if (!ValidarCampos(txtCepEnd,"CEP")) return
		if (!ValidarTipoInfo(txtCepEnd,2,"CEP")) return false;
		if (!ValidarCampos(txtContatoEnd,"Contato")) return
		if (!ValidarCampos(txtTelEndArea,"Telefone")) return false
		if (txtTelEndArea.value.length != 2)
		{ 
			alert("Código de area do telefone inválido.")
			txtTelEndArea.focus()
			return false
		}	
		if (!ValidarCampos(txtTelEnd,"Telefone")) return
		if (!ValidarCampos(txtCNPJ,"CNPJ do Endereço de Instalação")) return
		if (!VerificarCpfCnpj(txtCNPJ,2)) return false
		
		if (!ValidarCampos(cboTipoPonto,"Tipo do Ponto(Instalação/Intermediário)")) return false
		if (rdoPropAcessoFisico[1].checked && parseInt("0"+cboTecnologia.value) != 4)
		{
			if (!ValidarCampos(cboInterFaceEnd,"Interface")) return
		}	

		var intRet=alertbox('Deseja permanecer com os dados?','Sim','Não','Sair')
		switch (parseInt(intRet))
		{
			case 1:
				xmlUpd(false)
				updOrdemXml()
				AtualizarLista()
				break		
			case 2:
				xmlUpd(true)
				updOrdemXml()
				AtualizarLista()
				break
		}
		cboVelAcesso.focus()
	}
}	

function RemoverAcessoListaAlteracao(){   
	with (document.forms[1]){		
		DesabilitarCamposAlt(false,hdnIntIndice.value)		
		if (hdnIntIndice.value != ""){			
			RemoverAcesso(hdnIntIndice.value,true)			
			LimparInfoAcesso()			
			updOrdemXml()			
			AtualizarLista()			
			rdoPropAcessoFisico[0].focus()						
			//ResgatarGLA()			
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


//Alteração Aline 
//Rotina : AprovarAvaliacao
//Descrição: Rotina Criada para aprovar a avaliação
//Data 21/09/2006

function AprovarAvaliacao(Sol_ID)
{
		
		//if (!GravarAlteracao()) return
		if (!AprovarAltGic(Sol_ID)) return
		
		self.setTimeout

		var intRet=alertbox('Confirma Aprovação da Avaliação ?','Sim','Não')
		switch (parseInt(intRet))
		{
			case 1:
				
				with (document.Form4)
				{	
					//Depois Finaliza a Solicitação na fase de Avaliação para a próxima fase
					hdnAcao.value = "AprovarAcesso"					
					hdnXml.value = objXmlGeral.xml					
					target = "IFrmProcesso"
					action = "ProcessoSolic.asp"
					submit()
					
					return
				    break					
				    
				}
				break

			case 2:
				
				return
				break
		}

}


//Alteração Aline 
//Rotina : AprovarAvaliacao
//Descrição: alteração da chamada da rotina em asp
//Data 21/09/2006

function AprovarAlteracao(Sol_ID)
{
		
		if (!GravarAlteracao()) return
		
		self.setTimeout

		var intRet=alertbox('Confirma Alteração da Avaliação ?','Sim','Não')
		switch (parseInt(intRet))
		{
			case 1:
				
				with (document.Form4)
				{
					hdnDesigServ.value = document.forms[0].hdnDesigServ.value;
					hdnCboServico.value = document.forms[0].hdnCboServico.value;
										
					hdnAcao.value = "AlterarInfoAcesso"					
					hdnXml.value = objXmlGeral.xml					
					target = "IFrmProcesso"
					action = "ProcessoSolic.asp"
					submit()
					return
				    break					
				    					
				}
				break

			case 2:
				
				return
				break
		}

}

function DevolverParaGIC(Sol_ID)
{		
		if (!AprovarAltGic(Sol_ID)) return
				
		var intRet=alertbox('Confirma Devolução para o GIC ?','Sim','Não')
		switch (parseInt(intRet))
		{
			case 1:
				
				with (document.Form4)
				{									
					hdnAcao.value = "DevolverAcesso"
					target = "IFrmProcesso"
					action = "ProcessoSolic.asp"
					submit()					
				}
				return	
				break

			case 2:
				
				return
				break
		}

}


function AprovarAltGic(Sol_ID)
{
	if (!GravarAlteracao()) return false
		
		with (document.Form4)
		{	
						
			hdnAcao.value = "AlterarInfoAcesso"					
			hdnXml.value = objXmlGeral.xml					
			target = "IFrmProcesso"
			action = "ProcessoSolic.asp"
			submit()
			return true	  			
		}
}

function GravarAlteracao()
{
	with (document.forms[0])
	{
		
		
		if (!ValidarCampos(txtRazaoSocial,"Nome do Cliente/Razão Social")) return
		if (!ValidarCampos(txtNomeFantasia,"Nome Fantasia")) return false
		if (!ValidarCampos(txtContaSev,"Conta Corrente")) return false
		if (!ValidarCampos(txtSubContaSev,"Sub Conta")) return false
		
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
				return false
			}
			else
			{	
				if (parseInt(txtOrderEntry[0].value) < 1964)
				{
					alert("Ano da Order Entry inválido. Ano dever ser maior ou igual a 1964.")
					return false
				}
				else
				{
					hdnOrderEntry.value = cboSistemaOrderEntry.value + txtOrderEntry[0].value + txtOrderEntry[1].value + txtOrderEntry[2].value
				}	
			}	
		}	
		if (!MontarDesigServico()) return false
		
		if (!ValidarCampos(cboServicoPedido,"Serviço")) return false
		if (!ValidarCampos(cboVelServico,"Velocidade do Serviço")) return false

		if (!ValidarCampos(txtNroContrServico,"Nº do Contrato Serviço")) return false

		//Designação do Acesso Principal
		hdnDesigAcessoPri.value = ""
		if (txtDesigAcessoPri.value != "" && txtDesigAcessoPri.value.length < 7)
		{
			alert("Designação do Acesso Principal(678) fora de padrão 678N7.")
			txtDesigAcessoPri.focus()
			return false
		}
		
		if (txtDesigAcessoPri.value != "" && txtDesigAcessoPri.value.length == 7)
		{
			hdnDesigAcessoPri.value = txtDesigAcessoPri0.value + txtDesigAcessoPri.value
		}

		if (!ValidarTipoInfo(txtDtIniTemp,1,"Data Início Temporário")) return false;
		if (!ValidarTipoInfo(txtDtFimTemp,1,"Data Fim Temporário")) return false;
		if (!ValidarTipoInfo(txtDtDevolucao,1,"Data Devolução Temporário")) return false;
		//if (!ValidarCampos(txtDtEntrAcesServ,"Data Desejada de Entrega do Acesso ao Serviço")) return false;
		if (!ValidarTipoInfo(txtDtEntrAcesServ,1,"Data Desejada de Entrega do Acesso ao Serviço")) return false;
		if (!ValidarTipoInfo(txtDtPrevEntrAcesProv,1,"Data Prevista de Entrega do Acesso pelo Provedor")) return false;

	}
	
	
	//Verifica se tem acesso adicionado
	var objNode = objXmlGeral.selectNodes("//xDados/Acesso")
	if (objNode.length == 0){
		alert("Informações do Acesso são Obrigatórias.")
		document.forms[1].rdoPropAcessoFisico[0].focus()
		return false
	}

	var objNodeAux = objXmlGeral.selectNodes("//Acesso[cboTipoPonto='I']")
	if (objNodeAux.length == 0){
		alert("É necessário pelo menos um ponto de instalação.")
		return false
	}

	with (document.forms[1]){
		
		//GLA é um campo obrigarório para TER/CLLI
		//var blnAchou = VerificaPropAcesso('TER')
		//if (blnAchou && IsEmpty(hdntxtGLA.value) && hdnObrigaGla.value == "1" ){
		//	alert("GLA é um campo obrigatório.")
		//	return false
		//}	
		blnAchou = VerificaPropAcesso('CLI')
		if (blnAchou && IsEmpty(hdntxtGLA.value) && hdnObrigaGla.value == "1" ){
			alert("GLA é um campo obrigatório.")
		//	return false
		}	
	}	

	


	with (document.forms[2]) //Form3
	{


		var blnAchouSatelite = false


		if (IsEmpty(document.Form2.hdnUserGICL.value)){alert('GIC-L é um campo obrigatório.');return false}
		
		if (!ValidarCampos(cboLocalEntrega,"Estação do Local de Entrega")) return false
		if (!ValidarCampos(cboLocalConfig,"Estação do Local de Configuração")) return false

		
		var objNode = objXmlGeral.selectNodes("//xDados/Acesso")
		if (objNode.length == 1){
			for (var intIndex=0;intIndex<objNode.length;intIndex++){
				var intChave = objNode[intIndex].childNodes[0].text
				var intTec = RequestNodeAcesso(objXmlGeral,"cboTecnologia",intChave)
				if (intTec == 4) blnAchouSatelite = true
			}
		}	
		
		
		if (!ValidarCampos(cboOrgao,"Orgão")) return
		if (IsEmpty(hdntxtGICN.value)){alert('GIC-N é um campo obrigatório.');txtGICN.focus();return false}

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
	if (document.Form2.hdnUserGICL.value != "")	AdicionarNode(objXmlGeral,"hdntxtGICL",document.Form1.hdnUserGICL.value)
	if (document.Form2.hdntxtGLA.value != "")	AdicionarNode(objXmlGeral,"hdntxtGLA",document.Form2.hdntxtGLA.value)
	if (document.Form2.hdntxtGLAE.value != "")	AdicionarNode(objXmlGeral,"hdntxtGLAE",document.Form2.hdntxtGLAE.value)


	with (document.forms[0]) //Form3
	{
		//Tipo do Contrato
		if (document.forms[0].rdoNroContrato[0].checked) AdicionarNode(objXmlGeral,"intTipoContrato",1)
		if (document.forms[0].rdoNroContrato[1].checked) AdicionarNode(objXmlGeral,"intTipoContrato",2)
		if (document.forms[0].rdoNroContrato[2].checked) AdicionarNode(objXmlGeral,"intTipoContrato",3)
	}



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
				if (!AdicionarAcessoLista(true)) return false
				break
			case 3:
				return 
				break
		}		
	}
	
	return true	

//Acerto Aline Avaliador

// Alterado por PSOUTO
//	ACERTO NA RN DE AVALIADOR

//	with(document.Form4){	    
//		hdnAcao.value = "AlterarInfoAcesso"
//		hdnXml.value = objXmlGeral.xml
//		target = "IFrmProcesso"
//		action = "ProcessoSolic.asp"		
//		submit()
//		return true
//	}	

	
	//return false
//	with (document.Form4){		
//		hdnAcao.value = "AlterarInfoAcesso"
//		var valores = 'gravaavaliacao.asp?hdnAcao=' + hdnAcao.value + '&hdnTipoAcao=' + hdnTipoAcao.value
//		valores = valores + '&hdnIdAcessoLogico=' + hdnIdAcessoLogico.value + '&hdnSolId=' + hdnSolId.value + '&hdn678=' + hdn678.value + '&hdnTipoProcesso='+ hdnTipoProcesso.value
//		Retornas = window.showModalDialog(valores, objXmlGeral, 'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;')
//		return Retornas
//		
//		}

	
}

