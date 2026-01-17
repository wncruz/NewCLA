<%@ CodePage=65001 %>   
<%
Response.ContentType = "text/html; charset=utf-8"
Response.Charset = "UTF-8"

'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: solicitacao.ASP
'	- Descrição			: Cadastra/Altera uma solicitação no sistema CLA
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%

''@@ LPEREZ 13/10/2005
Dim strVisada			'Tipo Visada
Dim strGrupo			'Grupo Cliente
Dim strOriSol			'Origem Solicitacao
Dim strProjEspecial	' Projeto Especial
''LP

'@@Davif
Dim strNecessitaRecurso		'Se será necessário recurso (Sim/Não)
Dim strEmiteOTS				'Se será necessário emitir OTS (Sim/Não)
Dim strReaprFisico			'Reaproveitar Fisico (Sim/Não)
Dim strUsernamecadastrador	'UserName do Cadastrador
Dim strtelefonecadastrador	'Telefone do Cadastrador
Dim strorigem				'OrigemSolicitacao (4=APG/x=SGAPlus, etc..)
Dim strOrderEntryNumSis		'Numero de Identificação da OE no Sistema Origem (APG, etc..)
Dim bbloqueia				'Variavel de controle para Bloquear campos
Dim bdesbloqueia			'Variavel de controle para para Desbloquear campos
Dim dblSolAPGId				'IDentificador da Solicitação APG
Dim readOnly				'Variavel de controle para Travar campos
bbloqueia = "disabled=true"
bdesbloqueia =" "
readOnly = "true"

Dim intAno				'Ano
Dim strUserNameGICLAtual'UserName GicL
Dim dblNroSev			'Número da Sev do sistema SSA
Dim strRazaoSocial		'Razão social
Dim strNomeFantasia 	'Nome fantazia
Dim strContaSev			'Conta corrente
Dim strSubContaSev		'Sub conta
Dim strIE				'IE
Dim strIM				'IM
Dim dblCNPJ				'CNPJ
Dim strOrder			'Order Entry
Dim intTamSis			'Tamanho da OrderEntry utilizado para quebrar o campo
Dim strOrderEntrySis	'Sistema da OrderEntry
Dim strOrderEntryAno	'Ano da OrderEntry
Dim strOrderEntryNro	'Número da OrderEntry
Dim strOrderEntryItem	'Item da OrderEntry
Dim strDtPedido			'Data do pedido
Dim dblVelServico		'Id da Velocidade do serviço
Dim strTipoContratoServico'Tipo do cantrato
Dim strNroContrServico	'Número do contrato
Dim dblIdLogico			'Número do acesso lógico
Dim dblDesigAcessoPri	'Designação do acesso principal
Dim strDtEntrAcesServ	'Data de entrega do acesso ao serviço
Dim strDtPrevEntrAcesProv 'Data prevista de entrega do acesso pelo provedor
Dim strHtmlGla			'Html com username/nome e ramal so GLA
Dim strUserNameGLA		'UserName do GLA
Dim strNomeGLA			'Nome do GLA
Dim strRamalGLA			'Ramal do GLA
Dim strUserNameGICN		'UserName do GICN
Dim strNomeGICN			'Nome do GicN
Dim strRamalGICN		'Ramal do GICN
Dim strUserNameGICL		'UserName do GICL
Dim strNomeGICL			'Nome do GICL
Dim strRamalGICL		'Ramal do GICL
Dim strUserNameGLAE 	'UserName do GLAE
Dim strNomeGLAE			'Nome do GLAE
Dim strRamalGLAE		'Ramal do GLAE

Dim strUserNameAVL 		'UserName do AVL
Dim strNomeAVL			'Nome do AVL
Dim strRamalAVL			'Ramal do AVL

Dim dblSolId			'ID da Solicitação SOL_ID
Dim strItemSel			'Controle para o item que esta selecionado nos combos (selected)
Dim dblSerId			'ID do serviço (Ser_Id)
Dim strPropEnd			'Proprietário do endereço
Dim strEndCid			'Sigla da cidade
Dim dblAcaId			'ID da Ação Aca_ID
Dim strPropAcessoFisico 'Proprietário do acesso físico
Dim dblTecId			'ID da tecnologia Tec_id
Dim dblProId			'ID do Provedor Pro_id
Dim dblRegId			'ID do Regime de contrato Reg_Id
Dim dblPrmId			'ID da promoção Prm_Id
Dim strObsProvedor		'Obeservações para o provedor
Dim strEnd				'Nome do logradouro
Dim strComplEnd			'Complemento do logradouro
Dim strBairroEnd		'Bairro do logradouro
Dim strCepEnd			'CEP do logradouro
Dim strContatoEnd		'Conotao do logradouro
Dim strTelEnd			'Telefone  do logradouro
Dim strTelEndArea		'Area do Telefone  do logradouro
Dim strUFEnd			'UF do logradouro
Dim strNroEnd			'Número do logradouro
Dim strLogrEnd			'Sigla do logradouro
Dim strInterFaceEnd 	'Interface do logradouro
Dim strEndCidDesc		'Decrição da cidade do logradouro
Dim dblOrgId			'ID do orgão Org_Id
Dim dblStsId			'Id do Status Sts_Id
Dim strHistoricoSol		'Histórico da solicitação
Dim strPropAcessoFis	'Proprietário do acesso para o id físico gravado (Instalação)
Dim strVelAcesso		'Velocidade do acesso para o id físico gravado (Instalação)
Dim strDtIniTemp		'Data de inicio do acesso temporário
Dim strDtFimTemp		'Data de fim do acesso temporário
Dim strDtDevolucao		'Data de entrega do acesso temporário
Dim strDtTesteFimaFim	'Data Prevista do Teste Fim a Fim
Dim dblLocalEntrega 	'ID do Local de Entrega Esc_Id
Dim dblLocalConfig		'ID do Local de Configuração Esc_Id
Dim strInterfaceEbt		'Interface na EBT
Dim strContEscEntrega	'Contato no local de entrega
Dim strTelEscEntrega	'Telefone do contato no local de entrega
Dim objRSSolic			'Dados da solicitacão em edição
Dim DBAction1			'Ação auxiliar
Dim objRSFis			'Acessos físicos
Dim strIdAcessoFisicoInst 'Id do Acesso físico de instalação
Dim strVelDescAcessoFisicoInst'Velocidade do Acesso físico de instalação
Dim objRSDatas			'Datas
Dim strIdAcessoFisicoPtoI
Dim strVelDescAcessoFisicoPtoI
Dim dblCtfcId
Dim strCodSap
Dim dblNroPI
Dim strSiglaCliente
Dim strCNLSiglaCli
Dim strTipoPonto
Dim strTipoVel
Dim intIndice
Dim objDicProp
Dim strTipoAlteracao
Dim strProgramacaoAlteracao
Dim strPossuiAvaliador
dim strSelecao
Dim strcomp_troncoInicio
Dim	strcomp_tronco2m
Dim	strcomp_rota
Dim	strcomp_troncoInterface
Dim	strcomp_troncoFim
Dim	strcomp_rotaCNG
Dim	strcomp_contrato
Dim	strcomp_codServico
Dim	strcomp_servico
Dim	strcomp_obs 
DIm strXmlInfoAcesso

Set objDicProp = Server.CreateObject("Scripting.Dictionary")

'Monta o Xml de Acessos
%>
<!--#include file="../inc/xmlAcessos.asp"-->
<script>
var xmlhttp = null;

function ServicosPovoamento() {
	//document.forms[0].hdnOrderEntrySis.value="";
	//cboSistemaOrderEntry.value = "";
	//document.getElementById('cboSistemaOrderEntry').value = ""; 
	//alert(document.forms[0].cboSistemaOrderEntry.value);
	
	//document.forms[0].cboSistemaOrderEntry.value = ""
	
    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
        xmlhttp.onreadystatechange = processadorMudancaEstado;
        xmlhttp.open("POST", "../Ajax/AJX_Resgatar_Servicos.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<cboOrigemSol>"+document.getElementById('cboOrigemSol').value+"</cboOrigemSol>"
        xmlhttp.send(strXML);
		if (document.getElementById('cboOrigemSol').value == 3)
		//258
		{
		document.getElementById('cboStatusSolic').value = 177;
		}
		else
		{
		document.getElementById('cboStatusSolic').value = 0;
		}
		
    }
	
}

function processadorMudancaEstado () {

    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 	
          spnServico = document.getElementsByTagName("cboServicoPedido")[0]
		  valspnServico = spnServico.value   
		  document.getElementById("spncboServico").innerHTML = xmlhttp.responseText;
		  spnServico.value=valspnServico
			
        } else { 
            alert( "Erro: " + xmlhttp.statusText );  
        } 
    }
}

function Resgatar_SistemaOrderEntry() {
	//document.forms[0].hdnOrderEntrySis.value="";
	//cboSistemaOrderEntry.value = "";
	//document.getElementById('cboSistemaOrderEntry').value = ""; 
	//alert(document.getElementById('cboOrigemSol').value);
	
	//document.forms[0].cboSistemaOrderEntry.value = ""
	
    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
        xmlhttp.onreadystatechange = processadorMudancaEstadoOrder;
        xmlhttp.open("POST", "../Ajax/AJX_Resgatar_SistemaOrderEntry.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<cboOrigemSol>"+document.getElementById('cboOrigemSol').value+"</cboOrigemSol>"
        xmlhttp.send(strXML);
				
    }
}

function processadorMudancaEstadoOrder () {
    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 
			document.getElementById("spncboSistemaOrderEntry").innerHTML = xmlhttp.responseText;
        } else { 
            alert( "Erro: " + xmlhttp.statusText );  
        } 
    }
}

function Resgatar_SistemaID() {
	//document.forms[0].hdnOrderEntrySis.value="";
	//cboSistemaOrderEntry.value = "";
	//document.getElementById('cboSistemaOrderEntry').value = ""; 
	//alert(document.getElementById('cboOrigemSol').value);
	
	//document.forms[0].cboSistemaOrderEntry.value = ""
	
    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
        xmlhttp.onreadystatechange = processadorResgatar_SistemaID;
        xmlhttp.open("POST", "../Ajax/AJX_Resgatar_SistemaID.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<cboSistemaOrderEntry>"+document.getElementById('cboSistemaOrderEntry').value+"</cboSistemaOrderEntry>"
        xmlhttp.send(strXML);
				
    }
}
function processadorResgatar_SistemaID () {
    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 
			document.getElementById("spnResgatar_SistemaID").innerHTML = xmlhttp.responseText;
        } else { 
            alert( "Erro: " + xmlhttp.statusText );  
        } 
    }
}

function CompletarCampoIA(obj)
{
	
	if (obj.value != "" && obj.value != 0 )
	{
		var intLen = parseInt(obj.size) - parseInt(obj.value.length)
		
		
		if (!obj.TIPO){return;}
	
		switch (obj.TIPO.toUpperCase())
		{
			case "N":
				for (var intIndex=0;intIndex<intLen;intIndex++)
				{
					obj.value = "0" + obj.value
				}
				break
			default :
				for (var intIndex=0;intIndex<intLen;intIndex++)
				{
					obj.value = obj.value + " "
					//alert(obj.value)
				}
		}
	}	
}


function ServicosSEV() {

    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
        xmlhttp.onreadystatechange = processadorMudancaEstado;
        xmlhttp.open("POST", "../Ajax/AJX_Resgatar_Servicos_SEV.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<SEV>"+document.getElementById('txtNroSev').value+"</SEV>"
        xmlhttp.send(strXML);
				
    }
}

//GPON
function ResgatarFabONT() {
    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
		param = document.getElementById('cboTecnologia').value
		if (param==0){param=0}
        xmlhttp.onreadystatechange = processadorMudancaEstadoFabONT;
        xmlhttp.open("POST", "../Ajax/AJX_Resgatar_FabONT.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<param>"+param+"</param>"
        xmlhttp.send(strXML);
    }
}

function processadorMudancaEstadoFabONT () {
    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 
			document.getElementById("spncboFabricante").innerHTML = xmlhttp.responseText;
        } else { 
            alert( "Erro: " + xmlhttp.statusText ); 
			return 
        } 
    }
}



//GPON
function ResgatarTecnologiaSev() {
	//alert( "1" ); 
    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
		param = document.getElementById('txtNroSev').value
		//alert( param  ); 
		if (param==0){param=0}
        xmlhttp.onreadystatechange = processadorMudancaEstadoTecnologiaSev;
        xmlhttp.open("POST", "../Ajax/AJX_ResgatarTecnologiaSev.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<param>"+param+"</param>"
        xmlhttp.send(strXML);
    }
}

function processadorMudancaEstadoTecnologiaSev () {
    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 
			document.getElementById("strdivTecnologia").innerHTML = xmlhttp.responseText;
        } else { 
            alert( "Erro: " + xmlhttp.statusText ); 
			return 
        } 
    }
}


function ResgatarTipoONT() {
    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
		param = document.getElementById('cboFabricanteONT').value
		if (param==0){param=0}
        xmlhttp.onreadystatechange = processadorMudancaEstadoTipoONT;
        xmlhttp.open("POST", "../Ajax/AJX_Resgatar_TipoONT.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<dados><param>"+param+"</param><texto>S</texto></dados>"
        xmlhttp.send(strXML);
    }
}

function processadorMudancaEstadoTipoONT () {
    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 
			document.getElementById("spncboTipoONT").innerHTML = xmlhttp.responseText;
        } else { 
            alert( "Erro: " + xmlhttp.statusText ); 
			return 
        } 
    }
}

function ValidarItemOE(campo)
  {
    if (campo.value == "0")
    {
	  campo.value = "001"
	}							    
}

function isNumber(n) {   return !isNaN(parseFloat(n)) && isFinite(n); } 
//-------	Good inicio
function consultaCLE()
{
    if (isNumber(document.getElementById("txtConta15").value))
    {		

		var strConta15 = document.getElementById("txtConta15").value;
		document.Form1.txtContaSev.value = strConta15.substring(0,11)
		document.Form1.txtSubContaSev.value = strConta15.substring(11)		
		//xmlDoc = new ActiveXObject("Microsoft.XMLDOM");

        var xmlhp = new ActiveXObject("Msxml2.XMLHTTP"); 	
					
		if (xmlhp ) 
		{		
			xmlhp.onreadystatechange = function() 
			{
				if (xmlhp.readyState === 4 && xmlhp.status === 200) 
				{
					//console.log("Atualização realizada com sucesso: " + xmlhp.responseText);					
					 responseXML = xmlhp.responseText.split("|");				
                    if ( trim(responseXML[0]) != "*")
					{
						document.getElementById("txtRazaoSocial").value  = responseXML[0]
						document.getElementById("txtNomeFantasia").value = responseXML[1]					
					}
                    else{
						//document.getElementById("txtContaSev").value         = "" 
						document.getElementById("txtSubContaSev").value      = "" 
						document.getElementById("txtRazaoSocial").value      = ""
						document.getElementById("txtNomeFantasia").value     = ""	
						//document.getElementById("txtSegmento").value      	 = ""
						//document.getElementById("txtPorte").value     		 = ""			
						alert("Conta Corrente [" + strConta15 + "] não encontrada no CLE.")
						document.getElementById("txtConta15").focus()					
					}											
				}
				else{
					console.log("xmlhp.readyState: " + xmlhp.readyState + " xmlhp.status: " + xmlhp.status );
				}				
			}		
			xmlhp.open("POST", "retornaContaCLE.asp", true);		
			xmlhp.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded'); // Set content type for POST
			//xmlhp.setRequestHeader('encoding','ISO-8859-1');						
			xmlhp.send("conta15=" + strConta15)
								
		};			
								
	}			
}

function consultaCLE_old()
{
    if (isNumber(document.getElementById("txtConta15").value))
    {
		/**
	   //	alert("Consultar no CLE")
			var strConta15 = document.getElementById("txtConta15").value
			var xmlDoc  = new ActiveXObject("Microsoft.XMLDOM");
			var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
			var strXML
			strXML      = "<root>"
			strXML      = strXML + "<conta15>" +  document.getElementById("txtConta15").value + "</conta15>"
			strXML      = strXML + "</root>"
			xmlDoc.loadXML(strXML);
			xmlhttp.Open("POST","retornaContaCLE.asp" , false);
			xmlhttp.Send(xmlDoc.xml);
      strXML      = xmlhttp.responseText;				
//alert(strXML);
      xmlDoc.loadXML(strXML);
		
		**/
		
		//-------	Good
		var strConta15 = document.getElementById("txtConta15").value;
		var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
		var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
		var strXML = "<root>" +
						"<conta15>" + strConta15 + "</conta15>" +
						"</root>";
		xmlDoc.loadXML(strXML);
		 
		// Open a synchronous POST request to the ASP page
		xmlhttp.Open("POST", "retornaContaCLE.asp", false);
		 
		// Set the appropriate headers (if needed)
		xmlhttp.setRequestHeader("Content-Type", "text/xml");
		 
		try {
			xmlhttp.Send(xmlDoc.xml);
			// Check if the response is successful
			if (xmlhttp.status === 200) {
				var responseXML = xmlhttp.responseText;
				// Optionally check if the response is valid XML
				try {
					xmlDoc.loadXML(responseXML);
					alert("Response received: " + responseXML);
				} catch (e) {
					alert("Invalid XML received: " + e.message);
				}
			} else {
				alert("Error occurred: " + xmlhttp.status + " - " + xmlhttp.statusText);
			}
		} catch (e) {
			alert("Error occurred: " + e.message);
		}
			var ndCodRetorno   = xmlDoc.getElementsByTagName("codRetorno")[0].firstChild.nodeValue
			var ndRazaoSocial    = xmlDoc.getElementsByTagName("razaoSocial")[0].firstChild.nodeValue
			var ndNomeFantasia   = xmlDoc.getElementsByTagName("nomeFantasia")[0].firstChild.nodeValue
			var segmento   = xmlDoc.getElementsByTagName("segmento")[0].firstChild.nodeValue
			var porte    = xmlDoc.getElementsByTagName("porte")[0].firstChild.nodeValue
			var msg   = xmlDoc.getElementsByTagName("msg")[0].firstChild.nodeValue
			//alert('ndCodRetorno=' + ndCodRetorno)


//***
			if (ndCodRetorno=="*"){			
				document.getElementById("txtContaSev").value         = "" 
				document.getElementById("txtSubContaSev").value      = "" 
				document.getElementById("txtRazaoSocial").value      = ""
				document.getElementById("txtNomeFantasia").value     = ""		
				document.getElementById("txtSegmento").value      	 = ""
				document.getElementById("txtPorte").value     		 = ""			
				alert("Não foi possível consultar o CLE. Tente novamente.")
				document.getElementById("txtConta15").focus()
			}
			else if (ndCodRetorno=="2"){					
				document.getElementById("txtContaSev").value         = "" 
				document.getElementById("txtSubContaSev").value      = "" 
				document.getElementById("txtRazaoSocial").value      = ""
				document.getElementById("txtNomeFantasia").value     = ""	
				document.getElementById("txtSegmento").value      	 = ""
				document.getElementById("txtPorte").value     		 = ""	
				alert("SubConta não encontrada no CLE.")		
				document.getElementById("txtConta15").focus()	
			}
			else if(ndRazaoSocial!="*"){  
				document.getElementById("txtRazaoSocial").value      = ndRazaoSocial
				document.getElementById("txtNomeFantasia").value     = ndNomeFantasia
				document.getElementById("txtContaSev").value         = document.getElementById("txtConta15").value.substr(0,11)
				document.getElementById("txtSubContaSev").value      = document.getElementById("txtConta15").value.substr(11,4)
				document.getElementById("txtSegmento").value      	 = segmento
				document.getElementById("txtPorte").value     		 = porte
			}
			else
			{			
				document.getElementById("txtContaSev").value         = "" 
				document.getElementById("txtSubContaSev").value      = "" 
				document.getElementById("txtRazaoSocial").value      = ""
				document.getElementById("txtNomeFantasia").value     = ""	
				document.getElementById("txtSegmento").value      	 = ""
				document.getElementById("txtPorte").value     		 = ""			
				alert("Conta Corrente [" + strConta15 + "] não encontrada no CLE.")
				document.getElementById("txtConta15").focus()
			}
//***
/*			if (ndCodRetorno==0){
				document.getElementById("txtRazaoSocial").value      = xmlDoc.getElementsByTagName("razaoSocial")[0].firstChild.nodeValue
				document.getElementById("txtNomeFantasia").value     = xmlDoc.getElementsByTagName("nomeFantasia")[0].firstChild.nodeValue
				document.getElementById("txtContaSev").value         = document.getElementById("txtConta15").value.substr(0,11);
				document.getElementById("txtSubContaSev").value      = document.getElementById("txtConta15").value.substr(11,4);			
				ProcurarClienteCLE(true)
			}
			else if (ndCodRetorno==1)
			{
				document.getElementById("txtContaSev").value         = "" 
				document.getElementById("txtSubContaSev").value      = "" 
				document.getElementById("txtConta15").value          = ""
				document.getElementById("txtRazaoSocial").value      = ""
				document.getElementById("txtNomeFantasia").value     = ""
				document.getElementById("txtConta15").focus();				
//				alert('nfound!');
				alert("Conta Corrente [" + strConta15 + "] não encontrada no CLE.")
			}else
			{
				alert("Não foi possível realizar a comunicação com o CLE para a busca da Razão Social/Nome Fantasia. Tente novamente.")
			}
*/
		}
		else{
				document.getElementById("txtContaSev").value         = "" 
				document.getElementById("txtSubContaSev").value      = "" 
				document.getElementById("txtConta15").value          = ""
				document.getElementById("txtRazaoSocial").value      = ""
				document.getElementById("txtNomeFantasia").value     = ""
				document.getElementById("txtSegmento").value      	 = ""
				document.getElementById("txtPorte").value     		 = ""
				//document.getElementById("txtConta15").focus();			
		}							    
}
//-------	Good fim
function ProcurarClienteCLE(montaCboEnd)
{	//	  alert('montaCboEnd=' + montaCboEnd)
	with (document.forms[0])
	{   
			if (montaCboEnd == true) 
			{
				hdnMontaCboEnd.value = "S"
		  }else
		  {
		    hdnMontaCboEnd.value = "N"
		  }
	//	 alert('hdnMontaCboEnd.value=' + hdnMontaCboEnd.value)
			hdnAcao.value = "ProcurarClienteCLE"
			target = "IFrmProcesso"
			action = "ProcessoSolic.asp"
			submit()
	}	
}

function EsconderFabONT()
{
	document.getElementById("spncboFabricante").innerHTML = "";
}

function EsconderTipoONT()
{
	document.getElementById("spncboTipoONT").innerHTML = "";
}
</script>
<script language="javascript">
function MostraTec(idFac, strTecnologia) {
    var cboTecnologia = document.getElementsByName("cboTecnologia")[0];
    cboTecnologia.options.length = 0; 
    var arrTec = strTecnologia.split(";"); 
	var optionTec = new Option(':: TECNOLOGIA ','')
	cboTecnologia.options.add(optionTec); 	
    for (var i = 0; i < arrTec.length; i++) {
        var arrOpt = arrTec[i].split(",");        
        if (arrOpt[2] == idFac) {        
            var option = new Option(arrOpt[1], arrOpt[0]); 
            cboTecnologia.options.add(option); 
        }
    }
}
</script>

<%
Function indexPropAcesso(strProp)
	Select Case strProp
		Case "TER"
			indexPropAcesso = 0
		Case "EBT"
			indexPropAcesso = 1
		Case "CLI"
			indexPropAcesso = 2
	End Select
End Function



intAno = Year(Now)
strUserNameGICLAtual = Trim(strUserName)

'Perfis que podem acessar o essa págna E-GICL,GE-Ger.Usuario,GAT-GLA
For Each Perfil in objDicCef
	if Perfil = "E" then dblCtfcId = objDicCef(Perfil)
Next
if dblCtfcId = "" then
	For Each Perfil in objDicCef
		if Perfil = "GE" then dblCtfcId = objDicCef(Perfil)
	Next
End if
if dblCtfcId = "" then
	For Each Perfil in objDicCef
		if Perfil = "GAT" then dblCtfcId = objDicCef(Perfil)
	Next
End if

'Response.Write "DBLCTFCid: " & DBLCTFCid

'POVOAMENTO PARA SGAV
if objDicCef.Exists("POV") then
	strPOV = 1
else	
	strPOV = 0
end if 

'POVOAMENTO PARA SGAP
if objDicCef.Exists("POP") then
	strPOP = 1
else	
	strPOP = 0
end if 


Set objRS = db.execute("CLA_sp_sel_ConfigCtf null," & dblCtfcId)

'response.write dblCtfcId

if not objRS.Eof then
	if trim(objRS("Cfg_RedirecionamentoCarteira")) = "1" Then
		strObrigaGla = objRS("Cfg_RedirecionamentoCarteira")
	else
		strObrigaGla = 0 'objRS("Cfg_RedirecionamentoCarteira")
	end if
Else
	strObrigaGla = 0
End if


'Response.Write "<script language=javascript>alert('Cfg_RedirecionamentoCarteira " & strObrigaGla &   "');</script>"

if not objRS.Eof and not objRS.Bof then
	strPossuiAvaliador = objRS("Cfg_Avaliador")
Else
	strPossuiAvaliador = 0
End if

'Response.Write "Avaliador: " & strPossuiAvaliador
''@@Davif - Obtem se é de Origem - APG
strOrigem = Request.Form("hdnOEOrigem")
dblSolAPGId = Trim(Request.Form("hdnSolAPGId"))
'response.write "<script>alert('"&dblSolAPGId&"')</script>"
'response.end

	dblSolId = ""

	dblSolId = Trim(Request.Form("hdnSolId"))

If Trim(Request.Form("hdnAcao")) = "Alteracao"  then

	If Trim(Request.Form("hdnAcao")) = "Alteracao"  And (dblSolId = "" or dblSolId = "0") Then
	
		Response.Write "<script language=javascript>alert('Solicitação indisponível.');window.location.replace('main.asp');</script>"
		Response.End
	
	End if
	
	if dblSolId <> "" then
		strSQL = "select Ctf_ID from cla_solicitacaoctf where sol_id = " & dblSolId
		Set objRSCtfc = db.execute(strSQL)
			dblCtfcId = objRSCtfc("Ctf_ID")
	end if

		if dblSolId <> "" then
			Set objRSSolic = db.execute("CLA_sp_view_solicitacaomin " & dblSolId)
		End if
	

		if objRSSolic.Eof then
			Response.Write "<script language=javascript>alert('Solicitação indisponível.');window.location.replace('main.asp');</script>"
			Response.End
		End if

	
	If Trim(Request.Form("hdnAcao")) = "Alteracao" Then

		dblIdLogico			=	Trim(objRSSolic("Acl_IDAcessoLogico"))

	End If
	'Xml com os pontos
	Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
	objXmlDados.loadXml("<xDados/>")

	varRSVerFis = ""
	Set objRSVerFis = db.execute("SELECT ACF_ID FROM CLA_ACESSOLOGICOFISICO WHERE ACL_IDACESSOLOGICO = " & dblIdLogico)
	
	While Not objRSVerFis.Eof	
		varRSVerFis = varRSVerFis &","&objRSVerFis("ACF_ID")
		objRSVerFis.movenext
	Wend
	
	Vetor_Campos(1)="adInteger,4,adParamInput,"
	Vetor_Campos(2)="adInteger,4,adParamInput,"
	Vetor_Campos(3)="adDouble,8,adParamInput," & dblIdLogico
	Vetor_Campos(4)="adInteger,4,adParamInput,"
	Vetor_Campos(5)="adInteger,4,adParamInput,A"
	strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_VIEW_PONTO",5,Vetor_Campos)

	Set objRSFis = db.Execute(strSqlRet)
	if Not objRSFis.EOF and not objRSFis.BOF then
		Set objXmlDados = MontarXmlAcesso(objXmlDados,objRSFis,"")
		strXmlAcesso = FormatarXml(objXmlDados)
		intAcesso = 1
	End if
	
	'Dados do Acesso lógico
	if Not objRSSolic.Eof then
		
		dblNroSev		= Trim(objRSSolic("Sol_SevSeq"))
			
		'Localiza informações da solução SSA
		Vetor_Campos(1)="adInteger,2,adParamInput," & dblNroSev
		Vetor_Campos(2)="adInteger,2,adParamOutput,0"
		Call APENDA_PARAM("CLA_sp_sel_historico_solucao_ssa",2,Vetor_Campos)
		
		Set objRSCli = ObjCmd.Execute
		DBAction = ObjCmd.Parameters("RET").value
		
		'response.write "<script>alert('"&DBAction&"')</script>"
		
		if Cint("0" & DBAction) <> 0 then
			Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
			'response.write "<script>alert('"&DBAction&"')</script>"
			desativar_gravar = true
			
			strHigienizado = "NÃO"
			
		Else
			desativar_gravar = false
			
			strHigienizado = "SIM"
			
			strEnd			= Trim(objRSCli("PRE_NOMELOGR"))			'Nome do logradouro
			strComplEnd		= Trim(objRSCli("ACP_COMPLEND"))		'Complemento do logradouro
			strBairroEnd	= Trim(objRSCli("PRE_BAIRRO"))			'Bairro do logradouro
			strCepEnd		= Trim(objRSCli("PRE_COD_CEP"))				'CEP do logradouro
			
			'''strContatoEnd	= Trim(objRSCli("SEV_CONTATOLOCAL"))			'Contato do logradouro
			'''if len(Trim(objRSCli("SEV_TELCONTLOCAL"))) = 10 then
			'''	strTelEnd		= Right(Trim(objRSCli("SEV_TELCONTLOCAL")),8)		
			'''else
			'''	strTelEnd		= Right(Trim(objRSCli("SEV_TELCONTLOCAL")),9) 'Telefone  do logradouro
			'''end if 			
				
			'strTelEndArea   = mid(Trim(objRSSolic("Aec_Telefone")),1,2)'Area Telefone  do logradouro
			strUFEnd		= Trim(objRSCli("EST_SIGLA"))	 			'UF do logradouro
			strNroEnd		= Trim(objRSCli("PRE_NROLOGR"))			'Número do logradouro
			strLogrEnd		= Trim(objRSCli("TPL_SIGLA"))				'Sigla do logradouro
			strEndCid		= Trim(objRSCli("CID_SIGLA"))					'Sigla da cidade do logradouro
			strEndCidDesc	= Trim(objRSCli("LOCALIDADE"))				'Decrição da cidade do logradouro
			
			'response.write "<script>alert('"&strEndCidDesc&"')</script>"
		
		End if
				
		strGrupo				= Trim(objRSSolic("GCli_ID"))
		
		strRazaoSocial	= Trim(objRSSolic("Cli_Nome"))
		strNomeFantasia = Trim(objRSSolic("Cli_NomeFantasia"))
		strContaSev		= Trim(objRSSolic("Cli_CC"))
		strSubContaSev	= Trim(objRSSolic("Cli_SubCC"))

		if Trim(objRSSolic("Sol_OrderEntry")) <> "" then
			strOrder			= Trim(objRSSolic("Sol_OrderEntry"))
			intTamSis			= len(strOrder)-14
			strOrderEntrySis	= Ucase(Trim(Left(strOrder,intTamSis)))
			intTamSis			= intTamSis + 1
			strOrderEntryAno	= Mid(strOrder,intTamSis,4)
			intTamSis			= intTamSis + 4
			strOrderEntryNro	= Mid(strOrder,intTamSis,7)
			intTamSis			= intTamSis + 7
			strOrderEntryItem	= Right(strOrder,3)
		End if

		
		strOriSol					= Trim(objRSSolic("OriSol_ID"))
		strProjEspecial		= Trim(objRSSolic("Sol_IndProjEspecial"))
		
		strDtPedido				= Formatar_Data(Trim(objRSSolic("Sol_Data")))
		dblVelServico			= Trim(objRSSolic("IDVelAcessoLog"))
		strTipoContratoServico	= Trim(objRSSolic("Acl_TipoContratoServico"))
		strNroContrServico		= Trim(objRSSolic("Acl_NContratoServico"))
		dblDesigAcessoPriFull	= Trim(objRSSolic("Acl_IDAcessoLogicoPrincipal"))
		if dblDesigAcessoPriFull <> "" then
			dblDesigAcessoPri		= Right(dblDesigAcessoPriFull,len(dblDesigAcessoPriFull)-3)
		End if
		
		strDtIniTemp		= Formatar_Data(Trim(objRSSolic("Acl_DtIniAcessoTemp")))
		strDtFimTemp		= Formatar_Data(Trim(objRSSolic("Acl_DtFimAcessoTemp"))) '@@JKNUP: Correção. BO 50886
		strDtDevolucao		= Formatar_Data(Trim(objRSSolic("Acl_DtDevolAcessoTemp")))

		dblSerId			= Trim(objRSSolic("Ser_ID"))
		'Response.Write "<script language=javascript>alert('dblSerId=" & dblSerId & "');</script>"
		strObsProvedor		= Trim(objRSSolic("Sol_Obs"))

		strTipoAlteracao = Trim(objRSSolic("Sol_TipoAlteracao"))
		strProgramacaoAlteracao = Trim(objRSSolic("Sol_ProgramacaoAlteracao"))

		dblLocalEntrega = Trim(objRSSolic("Esc_IDEntrega"))
		'Endereço do local de instalação
		if Trim(dblLocalEntrega) <> "" then
			Set objRS = db.execute("CLA_sp_sel_estacao " & dblLocalEntrega)
			if Not objRS.Eof And Not objRS.Bof then
				strContEscEntrega	=	Replace(Trim(Cstr("" & objRS("Esc_Contato"))),"'","´")
				strTelEscEntrega	=	Replace(Trim(Cstr("" & objRS("Esc_Telefone"))),"'","´")
			End if
		End if
		dblLocalConfig = Trim(objRSSolic("Esc_IDConfiguracao"))
		strInterfaceEbt = Trim(objRSSolic("Acl_InterfaceEst"))


		strTipoAlteracao = Trim(objRSSolic("Sol_TipoAlteracao"))
		strProgramacaoAlteracao = Trim(objRSSolic("Sol_ProgramacaoAlteracao"))

		
	    if dblSolId <> "" then
		  Set objRS = db.execute("CLA_sp_view_agentesolicitacao " & dblSolId)
        End if
		if Not objRS.Eof then
			While Not objRS.Eof
				Select Case Trim(Ucase(objRS("Age_Desc")))
					Case "GLA"
						strUserNameGLA = Trim(objRS("Usu_Username"))
						strNomeGLA = Trim(objRS("Usu_Nome"))
						strRamalGLA = Trim(objRS("Usu_Ramal"))
					Case "GICN"
						strUserNameGICN = Trim(objRS("Usu_Username"))
						strNomeGICN = Trim(objRS("Usu_Nome"))
						strRamalGICN = Trim(objRS("Usu_Ramal"))
					Case "GICL"
						'strUserNameGICL = Trim(objRS("Usu_Username"))
						'strNomeGICL = Trim(objRS("Usu_Nome"))
						'strRamalGICL = Trim(objRS("Usu_Ramal"))
						'if Trim(objRS("Agp_Origem")) = "P" then
						'	strUserNameGICLAtual = strUserNameGICL
						'End if
					Case "GLAE"
						strUserNameGLAE = Trim(objRS("Usu_Username"))
						strNomeGLAE = Trim(objRS("Usu_Nome"))
						strRamalGLAE = Trim(objRS("Usu_Ramal"))

						'@@Davif - AVL - Avaliador
					Case "AVL"
						strUserNameAVL = Trim(objRS("Usu_Username"))
						strNomeAVL = Trim(objRS("Usu_Nome"))
						strRamalAVL = Trim(objRS("Usu_Ramal"))

				End Select
				objRS.MoveNext
			Wend
		End if

		dblOrgId = Trim(objRSSolic("Org_id"))
		dblStsId = Trim(objRSSolic("Sts_id"))
		strHistoricoSol = Trim(objRSSolic("StsSol_Historico"))

		Vetor_Campos(1)="adWChar,1,adParamInput," & Left(Trim(strRazaoSocial),1) 'Letra
		Vetor_Campos(2)="adInteger,4,adParamInput,null"
		Vetor_Campos(3)="adInteger,4,adParamInput," & dblCtfcId 'Ctf_Id

		Set objRSGLA = DB.Execute(APENDA_PARAMSTR("CLA_sp_check_usuario_redirsolicitacao",3,Vetor_Campos))

		if Not objRSGLA.Eof and Not objRSGLA.Bof then
			strNomeGlaAtual = objRSGLA("Usu_Nome")
			strRamalGlaAtual = objRSGLA("Usu_Ramal")
			strUserNameGlaAtual = objRSGLA("Usu_UserName")
		End if

		strHtmlGla	= "<table cellspacing=1 cellpadding=0 width=760px border=0 ><tr class=clsSilver >"
		strHtmlGla	= strHtmlGla & "<td width=170px ><font class=clsObrig>:: </font>UserName GLA</td>"
		strHtmlGla	= strHtmlGla & "<td colspan=5 >"
		strHtmlGla	= strHtmlGla & "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=2 bordercolorlight=#003388 bordercolordark=#ffffff width=355px >"
		strHtmlGla	= strHtmlGla & "<tr><td class=lightblue>&nbsp;"
		if strUserNameGlaAtual <> "" then
			strHtmlGla	= strHtmlGla & "<span id=spnUserNameGLA onclick=""ResgatarGLA()"">" & strUserNameGlaAtual &  "</span>"
		Else
			'strHtmlGla	= strHtmlGla & "<span id=spnUserNameGLA><font color=red>Não Encontrado</font></span>"
			strHtmlGla	= strHtmlGla & "<span id=spnUserNameGLA><font color=red></font></span>"
		End if
		strHtmlGla	= strHtmlGla & "</td></tr>"
		strHtmlGla	= strHtmlGla & "</table>"
		strHtmlGla	= strHtmlGla & "</td>"
		strHtmlGla	= strHtmlGla & "</tr>"
		strHtmlGla	= strHtmlGla & "<tr class=clsSilver>"
		strHtmlGla	= strHtmlGla & "<td width=170px>&nbsp;&nbsp;&nbsp;&nbsp;GLA</td>"
		strHtmlGla	= strHtmlGla & "<td width=355px>"
		strHtmlGla	= strHtmlGla & "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=2 bordercolorlight=#003388 bordercolordark=#ffffff width=100% >"
		strHtmlGla	= strHtmlGla & "<tr><td class=lightblue >&nbsp;"
		if strNomeGlaAtual <> "" then
			strHtmlGla	= strHtmlGla & "<span id=spnNomeGLA>" & strNomeGlaAtual &  "</span>"
		Else
			'strHtmlGla	= strHtmlGla & "<span id=spnNomeGLA><font color=red>**********************</font></span>"
			strHtmlGla	= strHtmlGla & "<span id=spnNomeGLA><font color=red></font></span>"
		End if
		strHtmlGla	= strHtmlGla & "</td></tr>"
		strHtmlGla	= strHtmlGla & "</table>"
		strHtmlGla	= strHtmlGla & "</td>"
		strHtmlGla	= strHtmlGla & "<td align=right >Ramal&nbsp;</td>"
		strHtmlGla	= strHtmlGla & "<td colspan=3 align=left >"
		strHtmlGla	= strHtmlGla & "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=2 bordercolorlight=#003388 bordercolordark=#ffffff width=100px >"
		strHtmlGla	= strHtmlGla & "<tr><td class=lightblue>&nbsp;"
		if strRamalGlaAtual <> "" then
			strHtmlGla	= strHtmlGla & "<span id=spnRamalGLA>" & strRamalGlaAtual & "</span>"
		Else
			'strHtmlGla	= strHtmlGla & "<span id=spnRamalGLA><font color=red>*******</font></span>"
			strHtmlGla	= strHtmlGla & "<span id=spnRamalGLA><font color=red></font></span>"
		End if
		strHtmlGla	= strHtmlGla & "</td></tr>"
		strHtmlGla	= strHtmlGla & "</table>"
		strHtmlGla	= strHtmlGla & "</td>"
		strHtmlGla	= strHtmlGla & "</tr></table>"

	
	End If
Else
	strDtPedido = right("0" & day(now),2) & "/" & right("0" & month(now),2) & "/" & year(now)
	strDtPrevEntrAcesProv = now() +  30
	strDtPrevEntrAcesProv = right("0" & day(strDtPrevEntrAcesProv),2) & "/" & right("0" & month(strDtPrevEntrAcesProv),2) & "/" & year(strDtPrevEntrAcesProv)
End if


%>

<script language='javascript' src="../javascript/solicitacao.js"></script>
<script language='javascript' src="../javascript/xmlObjects.js"></script>
<SCRIPT LANGUAGE=javascript>

var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
var objXmlAcessoFisComp = new ActiveXObject("Microsoft.XMLDOM")
var objXmlApg = new ActiveXObject("Microsoft.XMLDOM")
var objXmlFisicoAntigo = new ActiveXObject("Microsoft.XMLDOM")

objXmlGeral.preserveWhiteSpace = true
objXmlApg.preserveWhiteSpace = true
<%if Trim(Request.Form("hdnAcao")) = "Alteracao" then%>
	var intIndice = <%=intIndice%>
<%Else%>
	var intIndice = 0
<%End If%>


function Message(objXmlRet){

	var intRet = window.showModalDialog('Message.asp',objXmlRet,'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	if (intRet != "")
	{
		spnSolId.innerHTML = intRet
		document.Form3.txtGICN.value = ""
		document.Form3.hdntxtGICN.value = ""

		//Qdo. for processo de alteração Volta para tela Inicial da solictação
		if	(document.Form4.hdnTipoProcesso.value == 3 || document.Form4.hdnOrigem.value == "APG")
		{
			VoltarOrigem()
		}
	}
}
function Messagenew(objXmlRet){

	var intRet = window.showModalDialog('Messagenew.asp',objXmlRet,'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	if (intRet != "")
	{
		spnSolId.innerHTML = intRet
		document.Form3.txtGICN.value = ""
		document.Form3.hdntxtGICN.value = ""

		//Qdo. for processo de alteração Volta para tela Inicial da solictação
		if	(document.Form4.hdnTipoProcesso.value == 3 || document.Form4.hdnOrigem.value == "APG")
		{
			VoltarOrigem()
		}
	}
}


function CarregarLista()
{

	objXmlGeral.onreadystatechange = CheckStateXml;
	objXmlGeral.resolveExternals = false;
	if ('<%=intAcesso%>' != ''){
		objXmlGeral.loadXML("<%=strXmlAcesso%>")
	}else{
		var objXmlRoot = objXmlGeral.createNode("element","xDados","")
		objXmlGeral.appendChild (objXmlRoot)
	}
}

//Verifica se o Xml já esta carregado
function CheckStateXml()
{
  var state = objXmlGeral.readyState;

  if (state == 4)
  {
    var err = objXmlGeral.parseError;
    if (err.errorCode != 0)
    {
      alert(err.reason)
    }
    else
    {
		CarregarDoc()
	}
  }
}


function mostraWait(esconde){
	with (document.forms[0])
	{
		if(esconde == '1'){
			divWait.style.display = 'none';
		}else{
			divWait.style.display = 'inline';
		}
	}
}

function mostraSistemaWait(esconde){
	with (document.forms[0])
	{
		if(esconde == '1'){
			divSistemaWait.style.display = 'none';
		}else{
			divSistemaWait.style.display = 'inline';
		}
	}
}

function CarregarDoc()
{
	document.onreadystatechange = CheckStateDoc;
	document.resolveExternals = false;
}

function CheckStateDoc()
{
  var state = document.readyState;

  if (state == "complete")
  {
    var err = objXmlGeral.parseError;
    if (err.errorCode != 0)
    {
      alert(err.reason)
    }
    else
    {
		AtualizarLista()

		<%if Trim(Request.Form("hdnAcao")) = "Alteracao" or Request.Form("hdnOEOrigem")<>"APG" then%>
			ResgatarDesigServicoGravado(<%=dblSerId%>)
		<%End if%>
	}
  }
}

/* 
	=================================================
	
		Eduardo Araujo Analista Programador 
		Alteração realizada no dia 03/04/07
		
	=================================================
*/
function CarregarListaAnterior()
{

	objXmlFisicoAntigo.onreadystatechange = CheckStateXmlAnterior;
	objXmlFisicoAntigo.resolveExternals = false;
	if ('<%=intAcesso%>' != ''){
		objXmlFisicoAntigo.loadXML("<%=strXmlAcesso%>")
	}else{
		var objXmlRoot = objXmlFisicoAntigo.createNode("element","xDados","")
		objXmlFisicoAntigo.appendChild (objXmlRoot)
	}
}

//Verifica se o Xml já esta carregado
function CheckStateXmlAnterior()
{
  var state = objXmlFisicoAntigo.readyState;

  if (state == 4)
  {
    var err = objXmlFisicoAntigo.parseError;
    if (err.errorCode != 0)
    {
      alert(err.reason)
    }
    else
    {
		CarregarDocAnterior()
	}
  }
}



function CarregarDocAnterior()
{
	document.onreadystatechange = CheckStateDocAnterior;
	document.resolveExternals = false;
}

function CheckStateDocAnterior()
{
  var state = document.readyState;

  if (state == "complete")
  {
    var err = objXmlFisicoAntigo.parseError;
    if (err.errorCode != 0)
    {
      alert(err.reason)
    }
    else
    {
		ListaAnterior()
		<%if Trim(Request.Form("hdnAcao")) = "Alteracao"  or Request.Form("hdnOEOrigem")<>"APG" then%>
			ResgatarDesigServicoGravado(<%=dblSerId%>)
		<%End if%>
	}
  }
}


function CarregarListaApg()
{

	objXmlGeral.onreadystatechange = CheckStateXmlApg;
	objXmlGeral.resolveExternals = false;
	//if ('<%=intAcesso%>' != ''){
		//objXmlApg.loadXML("<%=strXmlAcesso%>")
	//}else{
		var objXmlRoot = objXmlApg.createNode("element","xDados","")
		objXmlGeral.appendChild (objXmlRoot)
	//}
}

//Verifica se o Xml já esta carregado
function CheckStateXmlApg()
{
  var state = objXmlGeral.readyState;

  if (state == 4)
  {
    var err = objXmlGeral.parseError;
    if (err.errorCode != 0)
    {
      alert(err.reason)
    }
    else
    {
		CarregarDocApg()
	}
  }
}

function CarregarDocApg()
{
	document.onreadystatechange = CheckStateDocApg;
	document.resolveExternals = false;
}

function CheckStateDocApg()
{
  var state = document.readyState;

  if (state == "complete")
  {
    var err = objXmlGeral.parseError;
    if (err.errorCode != 0)
    {
      alert(err.reason)
    }
    else
    {
		AtualizarListaApg()
		<%if Trim(Request.Form("hdnAcao")) = "Alteracao" or Request.Form("hdnOEOrigem")<>"APG" then %>
			ResgatarDesigServicoGravado(<%=dblSerId%>)
		<%End if%>
	}
  }
}

/*
	===========================================	
		fim da alteração do dia 03/04/2007
	===========================================
*/





function VoltarOrigem()
{
	with (document.forms[0])
	{
		target = self.name
		action = "<%=Request.Form("hdnPaginaOrig")%>"
		submit()
	}
}


// LPEREZ - 24/10/2005
function CheckGrupo()
{
	with (document.forms[0])
	{
		if (cboGrupo.value == 1)
		{
			divOrigemSol.style.display = '';
		}else{
			divOrigemSol.style.display = 'none';
		}
	}
}
//LP


//-->
</SCRIPT>

<form method="post" name="Form1">
<!-- Good Início -->
<input type="hidden" name="hdncboTecnologia" value="">
<input type="hidden" name="hdnTecnologia1">
<input type="hidden" name="hdnTecnologia2">
<input type="hidden" name="hdntxtFacilidade1" value="">
<input type="hidden" name="hdntxtFacilidade2" value="">
<input type="hidden" name="hdnAcaoMain" value ='<%= substring(upper(Trim(Request.Form("hdnAcao"))),1,3)%>'>

<!-- Good Fim -->

<input type="hidden" name="hdnOriSol_ID" value="<%=strOriSol%>">
<input type=hidden name=hdnMontaCboEnd>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUsuario>
<input type=hidden name=hdnOrigemEst>
<input type=hidden name=hdnCNLEstUsu>
<input type=hidden name=hdnComplEstUsu>
<input type=hidden name=hdnCboServico>
<input type=hidden name=hdnNomeCbo>
<input type=hidden name=hdnNomeLocal>
<input type=hidden name=hdnUserGICL value="<%=strUserNameGICLAtual%>">
<input type=hidden name=hdnDesigServ>
<input type=hidden name=hdnOrderEntry>

<input type=hidden name=hdnOrderEntrySis value="<%=strOrderEntrySis%>">
<input type=hidden name=hdnOrderEntryAno value="<%=strOrderEntryAno%>">
<input type=hidden name=hdnOrderEntryNro value="<%=strOrderEntryNro%>">
<input type=hidden name=hdnOrderEntryItem value="<%=strOrderEntryItem%>">

<input type=hidden name=hdnAclDesigSite value="<%=strAclDesigSite%>">

<input type=hidden name=hdnIdEnd>
<input type=hidden name=hdnIdEndInterme>
<input type=hidden name=hdnCNLAtual2>
<input type=hidden name=hdnDesigAcessoPri>
<input type=hidden name=hdnDesigAcessoPriDB value="<%=dblDesigAcessoPriFull%>">

<input type=hidden name=hdnPossuiAvaliador value="<%=strPossuiAvaliador%>">

<input type=hidden name=hdnIdAcessoLogico value="<%=dblIdLogico%>">
<input type=hidden name=hdnSolId value="<%=dblSolId%>">
<input type=hidden name=hdnDtSolicitacao value="<%=strDtPedido%>">
<input type=hidden name=hdnPadraoDesignacao >
<input type=hidden name=hdnCtfcId value="<%=dblCtfcId%>" >
<input type=hidden name=hdnSubAcao>
<input type=hidden name=hdnXmlReturn value="<%=Request.Form("hdnXmlReturn")%>">
<input type=hidden name=hdnAcfId >
<input type=hidden name=hdnStrAcfId >
<!-- LPEREZ - 21/10/2005 -->
<input type=hidden name=hdnProjEsp>
<%if strOrigem="APG" then%> 
	<input type=hidden name=idTarefaApg value="<%=objRSSolic("id_tarefa_Apg")%>">
 <% End if%>
<input type=hidden name=hdnCLINOME>
<input type=hidden name=hdnCLINOMEFANTASIA>
<input type=hidden name=hdnCliCC>
<input type=hidden name=hdnCLISUBCC>
<input type=hidden name=hdnSEGMENTO>
<input type=hidden name=hdnPORTE> 
<input type=hidden name=hdnSERDESC>
<input type=hidden name=hdnVELDESC>
<input type=hidden name=hdnOBSSEV>

<input type=hidden name=hdnVarRSVerFis value="<%=varRSVerFis%>">
<input type=hidden name=hdnstrPOP value="<%=strPOP%>" >
<input type=hidden name=hdnstrPOV value="<%=strPOV%>" >

<input type=hidden name=hdnobjChave value="">
<input type=hidden name=hdntxtFacilidade value="">


<!-- LP -->

<div id="divWait" style="background-color:#dcdcdc; width:300px; height:100px; float:left; margin:800px 0 0 280px; position:absolute; border:1px solid #0f1f5f; padding:40px 0 0 20px; display:none;">
	<p align="center" style="font-size: 12px; font-family:Arial, Helvetica; font-weight: bold; color:#003366;">Aguarde. Estamos consultando o CSL ...</p>
</div>

<div id="divSistemaWait" style="background-color:#dcdcdc; width:300px; height:100px; float:left; margin:800px 0 0 280px; position:absolute; border:1px solid #0f1f5f; padding:40px 0 0 20px; display:none;">
	<p align="center" style="font-size: 12px; font-family:Arial, Helvetica; font-weight: bold; color:#003366;">Aguarde. Sistema Processando ...</p>
</div>


<tr><td>
<table cellspacing="1" cellpadding="0" border=0 width="760">
	<tr >
		<th width=25%>&nbsp;•&nbsp;Solicitação de Acesso</th>
		<th width=25%>&nbsp;Nº&nbsp;:&nbsp;<span id=spnSolId><%=dblSolId%></Span></th>
		<%if Trim(Request.Form("hdnAcao")) = "Alteracao" then%>
			<th width=25%>&nbsp;Tipo&nbsp;:&nbsp;Alteração</th>
			<th width=25%>&nbsp;Acesso Lógico&nbsp;:&nbsp;<%=dblIdLogico%></th>
		<%Else%>
			<th width=25%>&nbsp;Tipo&nbsp;:&nbsp;Ativação</th>
		<%End if%>
		<th width=25%>&nbsp;Data&nbsp;:&nbsp;<%=strDtPedido%></th>
	</tr>
</table>



<table border=0 cellspacing="1" cellpadding="0" width="760">
	<tr >
		<th colspan=4>&nbsp;•&nbsp;Informações do Cliente</th>
	</tr>
	
   <tr class="clsSilver">
		<td width="170">&nbsp;&nbsp;&nbsp;&nbsp;Sev para procura</td>
		<td colspan="4">
			<input type="hidden" name="hdnOriSol_ID_N" value="3">
			<input type="text" class="text" name="txtNroSev" value="<%=dblNroSev%>" maxlength="8" size="10" onkeyup="ValidarTipo(this,0)" onblur="ResgatarSev()">&nbsp;
				
		</td>
	</tr>	
	
	<tr>
		</td>
		
		<td  class="clsSilver">
			
		&nbsp;
		Origem Solicitação <span align=right>
		</td>
<%

strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))
%>
         <td colspan="4" class="clsSilver">
				<!--<select name="cboOrigemSol" onchange="ServicosPovoamento();Resgatar_SistemaOrderEntry();">-->
				<select name="cboOrigemSol" id="cboOrigemSol"  onchange="ServicosPovoamento();">  
				<!--<select name="cboOrigemSol" > -->
					<option value="" ></option>
					<%
					set os = db.execute("CLA_sp_sel_OrigSolicitacao " & trim(strLoginRede))
					do while not os.eof
				        if os("OriSol_ID") <> "9" then 'ASMS (Não considerar para solicitações sem interface					
					%>
						<option value=<%=os("OriSol_ID")%>
					<%
						if strOriSol <> "" then
							if trim(strOriSol) = trim(os("OriSol_ID")) then
								response.write "selected"
							end if
						end if
					%>
						><%=ucase(os("OriSol_Descricao"))%></option>
					<%
				        end if					
						os.movenext
					loop
					%>
				</select>
		</td>
	</tr>

	<tr class="clsSilver">
		<td width="180px" ><font class="clsObrig">:: </font>Conta Corrente + SubConta</td>
		<td colspan="3" >
		   <input type=text class="text" name=txtConta15 id=txtConta15 size=15 maxlength=15	onKeyUp="ValidarTipo(this,0)" 
		   onblur="CompletarCampo(this);consultaCLE();" TIPO="N" value="<%=strContaSev%><%=left(strSubContaSev,4)%>"><span>   Ex.: 00000000000009999</span>  			
		   <input type=hidden name=txtContaSev id=txtContaSev value="<%=right(strContaSev,11)%>">			 
			<input type=hidden name=txtSubContaSev id=txtSubContaSev value="<%=left(strSubContaSev,4)%>"> 	
		<!--	<input type=hidden name=txtContaSev id=txtContaSev value="<%=strContaSev%>">			 
			<input type=hidden name=txtSubContaSev id=txtSubContaSev value="<%=strSubContaSev%>"> 		-->	 
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Razão Social</td>
		<td colspan="3" >
			<input type="text" class="text" id="txtRazaoSocial"  name="txtRazaoSocial"  maxlength="55" size="55" value="<%=strRazaoSocial%>"
			<%=bbloqueia%>>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Nome Fantasia</td>
		<td>
			<input type="text" class="text" id="txtNomeFantasia" name="txtNomeFantasia"  maxlength="20" size="25" value="<%=strNomeFantasia%>" <%=bbloqueia%> >
		</td> 
	</tr>
	
<!-- ''@@ 	<tr class="clsSilver" nowrap>
		<td width="170" nowrap><font class="clsObrig">:: </font>Segmento</td>
		<td width=279>
			<input type=text class="text" name=txtSegmento size=22 maxlength=22
			<%=bbloqueia%>
			 value="<%=SEGMENTO%>">
		</td>
		<td align=right width=89 >
		<p align="left"><font class="clsObrig">:: </font>Porte&nbsp;</td>
		<td width="189">
			<input type=text name=txtPorte class="text" size=22 maxlength=22
			<%=bbloqueia%>
			value="<%=PORTE%>">
		</td>
	</tr>
	-->
	
	<tr class="clsSilver">
		<td width="170px"><span id="spnLabelCliente"></span></td>
		<td colspan="3"><span id="spnCliente"></span></td>
	</tr>		
</table>
<table border=0 cellspacing="1" cellpadding="0" width="760">
	<tr >
		<th colspan="4" >
			&nbsp;•&nbsp;Informações do Serviço&nbsp;
		</th>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Order Entry</td>
		<td colSpan=3>
			<table border=0 border=0 cellspacing="0" cellpadding="0">
				<tr align=center class=clsSilver>
					<td>Sistema</td>
				</tr>
				<tr class=clsSilver>
					<td>
						<span id=spncboSistemaOrderEntry>
						<%
							set Order = db.execute("CLA_sp_sel_SistemaOrderEntry " & trim(OriSol_ID))
						%>
						<select name="cboSistemaOrderEntry"	<%if strOrigem="APG" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%> onChange="SistemaOrderEntry(this);hdnOrderEntrySis.value=this.value;Resgatar_SistemaID();"" >
						
							<option value="" ></option>
							<%
							do while not Order.eof
											
							%>
								<option value=<%=Order("SisOrderEntry_desc")%>
							<%
									if strOrderEntrySis <> "" then
										if trim(strOrderEntrySis) = trim(Order("SisOrderEntry_desc")) then
											response.write "selected"
										end if
									end if
							%>
								><%=ucase(Order("SisOrderEntry_desc"))%></option>
							<%
												
								Order.movenext
							loop
							%>
						</select>
						</span>
					
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px"></td>
		<td colspan="3"><span id=spnResgatar_SistemaID></span></td>
	</tr>
		
		
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Serviço</td>
		<td >
		<span id=spncboServico>
		<%
		'Response.Write "<script language=javascript>alert('Dados Carregados.');</script>"
		'seleciona servico
		set objRS = db.execute("CLA_sp_sel_servico")
'**********************************************
' Good início
'**********************************************		
		%>

        <select name="cboServicoPedido" id="cboServicoPedido" <%if strOrigem="APG" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%> onchange="ResgatarServico(this)" >
			<option></option>
			<%
'**********************************************
' Good fim
'**********************************************					
			While Not objRS.eof
				strItemSel = ""
				if Trim(dblSerId) = Trim(objRS("Ser_ID")) then strItemSel = " Selected " End if
				Response.Write "<Option value='" & objRS("Ser_ID") & "'" & strItemSel & ">" & objRS("Ser_Desc") & "</Option>"
				objRS.MoveNext
			Wend
			strItemSel = ""
			%>
		</select>
		</span>
		</td>
		<td width="150px" align=right><font class="clsObrig">:: </font>Velocidade&nbsp;</td>
		<td width="200px"><span id=spnVelServico>
				<select name="cboVelServico" <%if strOrigem="APG" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%> onChange="SelVelAcesso(this);ResgatarServico(this);" style="width:200px">
					<option ></option>
					<%  'if Trim(dblSerId) <> "" then
						set objRS = db.execute("CLA_sp_sel_AssocServVeloc null " )
						While Not objRS.eof
							strItemSel = ""
							
							if Trim(dblVelServico) = Trim(objRS("Vel_ID")) then strItemSel = " Selected " End if
							Response.Write "<Option value='" & objRS("Vel_ID") & "'" & strItemSel & ">" & Trim(objRS("Vel_Desc")) & "</Option>"
							objRS.MoveNext
						Wend
						strItemSel = ""
					'End if

					'Response.Write "<script language=javascript>alert('Dados Carregados.');</script>"
					%>
					</select>
			</span>
		</td>

	<%if strOrigem="APG" then%>

	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Designação do Serviço</td>
		<td colspan="3">
				<input type="text" class="text" name="txtdesignacaoServico" id="txtdesignacaoServico"
				<%if strOrigem="APG" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>
				value="<%=strDesignacaoServico%>" maxlength="22" size="30"><br>
		</td>
	</tr>
    
	<%else%>
	
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Designação do Serviço</td>
		<td colspan="3">		
			<span id=spnServico></span>
			<input type="hidden" id="spnServico_old" name="spnServico_old">
			
		</td>
	</tr>
		 <% End if%>
	 
	 <tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;Modalidade</td>
		<td colspan="3"><span id=spnModalidade></span></td>
	</tr>

	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Nº Contrato Serviço</td>
		<td colspan=3>
			<table rules="groups" cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="70%" >
				<tr><td nowrap width=200px >
					<input type=radio name=rdoNroContrato value=1
					<%if strOrigem="APG" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>
					onClick="spnDescNroContr.innerHTML= 'Ex.: VEM-11 XXX000012003'" checked <%if strTipoContratoServico = "1" then Response.Write " checked " End if%>>Contrato de Serviço</td><td></td></tr>
				<tr>
					<td nowrap><input type=radio name=rdoNroContrato onClick="spnDescNroContr.innerHTML= 'Padrão: A22'"
					<%if strOrigem="APG" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>
					value=2 <%if strTipoContratoServico = "2" then Response.Write " checked " End if%>>Contrato de Referência</td>
					<td nowrap>
						<input type="text" class="text" name="txtNroContrServico" value="<%=strNroContrServico %>"
						<%if strOrigem="APG" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>
						maxlength="22" size="30"><br>
						<span id=spnDescNroContr>Ex.: VEM-11 XXX00012003</span>
					</td>
				</tr>
				<tr><td nowrap><input type=radio name=rdoNroContrato
				<%if strOrigem="APG" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>
				onClick="spnDescNroContr.innerHTML= 'Padrão: A22'" value=3 <%if strTipoContratoServico = "3" then Response.Write " checked " End if%> >Carta de Compromisso</td><td></td></tr>
			</table>
		</td>
	</tr>

	

<% ''@@davif - Inclusão dos Campos Tipo de Alteração e Tipo de Programação da Alteração

if Trim(Request.Form("hdnAcao")) = "Alteracao" then %>

	<tr><th colspan=4>&nbsp;•&nbsp;Informações de Alteração</th></tr>

	<tr class="clsSilver">
		<td width=170px>&nbsp;&nbsp;&nbsp;&nbsp;Tipo de Alteração</td>
		<td >
			<Select name="cboTipoAlteracao" style="width:150px" >
			<option ></option>	&nbsp;
			<option value="Endereco"
			<% if trim(strTipoAlteracao) = "Endereco" Then %>
				 Selected > Endereço </option>
			<%Else %>
				> Endereço </option>
			<%End If %>

			<option Value="Velocidade"
			<% if trim(strTipoAlteracao) = "Velocidade" Then %>
				Selected> Velocidade </option>
			<%Else %>
				> Velocidade </option>
			<%End If %>

			<option Value="Velocidade+Facilidade"
			<% if trim(strTipoAlteracao) = "Velocidade+Facilidade" Then %>
				Selected> Velocidade e Facilidade</option>
			<%Else %>
				> Velocidade e Facilidade</option>
			<%End If %>

			</Select>
		</td>
		<td >&nbsp;&nbsp;&nbsp;Programação &nbsp;da &nbsp;&nbsp;&nbsp; Alteração&nbsp;</td>
		<td>
			<select name="cboProgAlteracao" style="width:150px" >
			<option ></option>	&nbsp;&nbsp;
			<option Value="Programada"
			<% if trim(strProgramacaoAlteracao) = "Programada" Then %>
				Selected > Programada  </option>
			<%Else %>
				 > Programada  </option>
			<%End If %>


			<option value= "Nao Programada"

			<% if trim(strProgramacaoAlteracao) = "Nao Programada" Then %>
				Selected > Não Programada  </option>
			<%Else %>
				> Não Programada  </option>
			<%End If %>

			</select>
		</td>
	</tr>
<%end if %>


</Form>
</table>

<table border=0 cellspacing="1" cellpadding="0" width="760">
	<tr >
		<th colspan=4 >
			&nbsp;•&nbsp;Análise do estudo de viabilidade (SEV   SSA) &nbsp;
		</th>
	</tr>
	<TR>
		<TD colspan=4 >
			<!--<span id=strProvedorSelSev LANGUAGE=javascript onactivate="return strProvedorSelSev_onactivate()">&nbsp;</span> -->
			<span id=strProvedorSelSev>
				
				<%
				if dblNroSev <> "" then

					Vetor_Campos(1)="adInteger,4,adParamInput," & dblNroSev
					Vetor_Campos(2)="adWChar,20,adParamInput,CLA" 
					Vetor_Campos(3)="adInteger,4,adParamOutput,0"
				
					Call APENDA_PARAM("CLA_sp_VALIDA_PRAZO_SEV ",3,Vetor_Campos)
					ObjCmd.Execute'pega dbaction
					'DBAction = ObjCmd.Parameters("RET3").value
					Set objRSSev = ObjCmd.Execute
					%>

					<table cellspacing=1 cellpadding=0 border=0 width=760> 
						<tr class=clsSilver>
						  <td width=321>Prazo:  	 <b> <%=objRSSev("Mensagem")%>   </b></td>
						  
						</tr>
					</table>
				<%
					'if DBAction = "0" then
						
						'Localiza informações da solução SSA
						Vetor_Campos(1)="adInteger,2,adParamInput," & dblNroSev
						Vetor_Campos(2)="adInteger,2,adParamOutput,0"
						'Call APENDA_PARAM("CLA_sp_sel_solucao_ssa",2,Vetor_Campos)
						Call APENDA_PARAM("CLA_sp_sel_historico_solucao_ssa",2,Vetor_Campos)
						
						
						Set objRSCli = ObjCmd.Execute
						if DBAction = 0 then
							If Not objRSCli.eof and  Not objRSCli.bof then
						%>
						<table cellspacing=1 cellpadding=0 border=0 width=760> 
								<tr class=clsSilver>
								  <td width=321>Status do Pedido:  	 <b> <%=objRSCli("STATUS_PEDIDO")%>   </b></td>
								  <td width=241>Status da SEV:  <b> <%=objRSCli("STATUS_SEV")%>   </b></td>
								</tr>
								<tr class=clsSilver>
								  		<td width=321>Nº do pedido de Venda do SNOA:  	 <b> <%= TratarAspasJS(Trim(objRSCli("Pedido_Venda_SNOA"))) %>  </b></td> 
							    		<td width=241>Caracteristíca do EILD :  <b> <%= TratarAspasJS(Trim(objRSCli("C_TIPO_EILD"))) %>  </b></td> 
								</tr> 
			
								<tr class=clsSilver>
										<td width=321>Ponta:  	 <b> <%= TratarAspasJS(Trim(objRSCli("Ponta"))) %>  </b></td> 
										<td width=241>Nº SEV Ponta A :  <b><%= TratarAspasJS(Trim(objRSCli("sevPontaA_seq"))) %>  </b></td> 
								</tr> 
			
								<tr class=clsSilver>
										<td width=321>Operadora Solicitante:  	 <b> <%= TratarAspasJS(Trim(objRSCli("OPR_DES"))) %>  </b></td> 
										<td width=321>Pabx Virtual:  	 <b> <%= TratarAspasJS(Trim(objRSCli("pabx_des"))) %>  </b> | Migração do Pabx Virtual:  	 <b> <%= TratarAspasJS(Trim(objRSCli("SEV_MIGRA_PABX"))) %>  </td> 
								</tr> 
																
								  		<table cellpadding=0 cellspacing=1 width=760 > 
											<tr> 
												<th nowrap>&nbsp;Facilidade</th>
												<th nowrap>&nbsp;Tecnologia</th> 
												<th>&nbsp;Provedor</th> 
												<th>&nbsp;Estação</th> 
												<th>&nbsp;Usuário</th> 
												<th>&nbsp;Data Vencimento</th> 
												<th>&nbsp;OBS</th> 
											</tr> 
											
									<%
									 While Not objRSCli.eof
									 %>
											<tr class=clsSilver > 
													<td width=15% nowrap >&nbsp; <%=objRSCli("FAC_DES")%>  </td>
													<td width=15% nowrap >&nbsp; <%=objRSCli("TEC_DES")%>  </td>	
													<td >&nbsp;  <%=objRSCli("FOR_DES")%> </td> 
													<td >&nbsp;  <%=objRSCli("ESTACAO")%> </td> 
													<td nowrap >&nbsp;  <%=objRSCli("SEV_USRNAMESOLUC")%>  </td> 
													<td nowrap >&nbsp; <%=objRSCli("data_vencimento")%> </td> 
													<td width=20% >&nbsp;  <%=objRSCli("SOL_OBS")%>  </td> 
												</tr> 
									<%
									objRSCli.MoveNext
									Wend
							End If
						 End If
					'End If
				End if
						%>
								</table> 
					
					</table> 
			
			</span>
		</TD >
	</TR>
	
	
</table>



<table ID= tbRadio cellspacing=1 cellpadding=0 width=760 border=0>
<Form name=Form2 method=Post>
<!--JCARTUS@ - Acessos Físicos sem CNPJ/IE/IM vinculado na base do CLA (CH-69613HSK)-->

<!-- Good 22/01/2025 inicio -->
<input type="hidden" name="hdnRdoAcesso" value="">
<input type="hidden" name="hdnCNLSiglaCentroCliDest" >
<input type="hidden" name="hdnComplSiglaCentroCliDest">
<input type="hidden" name="hdnPropIdFisBkp" value="">
<input type="hidden" name="hdnNomeFacilidade" value="">
<input type="hidden" name="hdnFacilidade" value="">

<!-- Good 22/01/2025 fim -->


<input type=hidden name=hdnCNPJ value="<%=dblCNPJ%>">
<input type="hidden" name="hdnOriSol_ID" value="<%=strOriSol%>">
<input type=hidden name=hdnIE value="<%=strIE%>">
<input type=hidden name=hdnIM value="<%=strIM%>">
<input type=hidden name=hdnIntIndice>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUsuario>
<input type=hidden name=hdnOrigemEst>
<input type=hidden name=hdnCNLEstUsu>
<input type=hidden name=hdnComplEstUsu>
<input type=hidden name=hdnSubAcao>
<input type=hidden name=hdnProvedor>
<input type=hidden name=hdnCboProvedor>
<input type=hidden name=hdnstrAcessoTipoRede >
<input type=hidden name=hdnSegmento value="">
<input type=hidden name=hdnPorte value=""> 
<input type=hidden name=hdnNroSev>
<input type=hidden name=hdnOrigemSol>
<input type=hidden name=hdnTipoAcaoAPG value="<%=strAcao%>">
<input type=hidden name=hdnTipoCEP>
<input type=hidden name=hdnCEP>
<input type=hidden name=hdnCNLNome>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnCNLAtual>
<input type=hidden name=hdnCNLAtual1>
<input type=hidden name=hdnNomeTxtCidDesc>
<input type=hidden name=hdnNomeCboCid>
<input type=hidden name=hdnUserGICL value="<%=strUserNameGICLAtual%>">
<input type=hidden name=hdntxtGLA value="<%=strUserNameGLA%>">
<input type=hidden name=hdntxtGLAE value="">
<input type=hidden name=hdnCtfcId value="<%=dblCtfcId%>" >
<input type=hidden name=hdnRazaoSocial>
<input type=hidden name=hdnChaveAcessoFis>
<input type=hidden name=hdnIdAcessoFisico	>
<input type=hidden name=hdnIdAcessoFisico1	>
<input type=hidden name=hdnPropIdFisico>
<input type=hidden name=hdnPropIdFisico1>
<input type=hidden name=hdnCompartilhamento		value="0">
<input type=hidden name=hdnNodeCompartilhado	value="0">
<input type=hidden name=hdnCompartilhamento1	value="0">
<input type=hidden name=hdnNovoPedido>
<input type=hidden name=hdnTecnologia>
<input type=hidden name=hdnVelAcessoFisSel>
<input type=hidden name=hdnAecIdFis>
<input type=hidden name=hdnEstacaoOrigem>
<input type=hidden name=hdnEstacaoDestino>
<input type=hidden name=hdnObrigaGla value="<%=strObrigaGla%>">
<input type=hidden name=hdnAcfId >
<input type=hidden name=strOrigemAPG value="<%=strOrigem%>">
<input type=hidden name=hdnUsuID value="<%=dblUsuID %>">
<input type=hidden name=hdnStrAcfId >

<%if strOrigem = "APG" then%>
  <input type=hidden name=cboUFEndAPG value="<%=strUFEnd%>">
  <input type=hidden name=txtEndCidAPG value="<%=strEndCid%>">
  <input type=hidden name=txtEndCidDescAPG value="<%=strEndCidDesc%>">
  <input type=hidden name=cboLogrEndAPG value="<%=strLogrEnd%>">
  <input type=hidden name=txtEndAPG value="<%=strEnd%>">
  <input type=hidden name=txtNroEndAPG value="<%=strNroEnd%>">
  <input type=hidden name=txtComplEndAPG value="<%=strComplEnd%>">
  <input type=hidden name=txtBairroEndAPG value="<%=strBairroEnd%>">
  <input type=hidden name=txtCepEndAPG value="<%=strCepEnd%>">
<%end if%>

<%if Trim(Request.Form("hdnAcao")) = "Alteracao" then%>
<input type=hidden name=hdnTipoProcesso value=3>
<%Else%>
<input type=hidden name=hdnTipoProcesso value=1>
<%End if%>
<input type=hidden name=hdnTipoTec>
<input type=hidden name=hdnCNLCliente>

	<!--	Eduardo Araujo Analista Programdor 
		Alteração realizada no dia 02/04/2007 	
	-->

			<tr><th colspan=4>&nbsp;•&nbsp;Acessos Físicos Utilizados</th></tr>
			<tr><td colspan=4>
					<table border=0 width=758 cellspacing=1 cellpadding=0>
						<tr>
							<th  width=15>&nbsp;</th>
							<th  width=35>&nbsp;Editar</th>
							<th  width=50>&nbsp;Prop Fis</th>
							<th  width=185>&nbsp;Provedor</th>
							<th  width=200>&nbsp;Velocidade</th>
							<th	 width=273>&nbsp;Endereço</th>
						</tr>
					</table>
				</td>
			</tr>
			<tr class=clsSilver>
				<td colSpan=4>
					<iframe id=IFrmAcessoFis
							name=IFrmAcessoFis
							align=left
							src="AcessosFisicos.asp"
							frameBorder=0
							width="100%"
							BORDER=0
							height=40>
					</iframe>
				</td>
			</tr>		
		
	<!-- novo ReCLA-->

	<tr>
		<th colSpan=4>&nbsp;•&nbsp;Informações do Acesso&nbsp;</th>
	</tr>
	<tr class="clsSilver">
			<td width=170px ><font class="clsObrig">:: </font>Prop do Acesso Físico</td>
			<td nowrap >	
				<%
				if dblNroSev <> "" then
					RSsolProv= ""
					Set objRSsolProv = db.execute("select ltrim(rtrim(for1.pro_id)) as pro_id from ssa_1.dbo.solucao sol inner join ssa.dbo.fornecedor for1 on sol.for_seq = for1.for_seq where sol.sev_seq = " & dblNroSev  )
			
					While Not objRSsolProv.Eof	
						RSsolProv = objRSsolProv("pro_id")  
					   objRSsolProv.movenext
					Wend
					
				end if
				'response.write "<script>alert('"&RSsolProv&"')</script>"
				%>		

				<input type=radio name=rdoPropAcessoFisico value="TER"  Index=0 onclick="EsconderTecnologia(0);ResgatarTecVel();">Terceiro&nbsp;&nbsp;&nbsp;
				<input type=radio name=rdoPropAcessoFisico value="EBT"  Index=1 onclick="EsconderTecnologia(0);ResgatarTecVel();">CLARO Brasil&nbsp;&nbsp;&nbsp;
				
				<!--Alterado por Fabio Pinho em 29/04/2016 - ver 1.0 - Inicio-->
				
				<input type=radio name=rdoPropAcessoFisico  value="CLI" disabled Index=2 onclick="EsconderTecnologia(0);ResgatarTecVel();">Cliente&nbsp;&nbsp;&nbsp;
				
				
				
				<!--
				<input type=radio name=rdoPropAcessoFisico  value="CLI" <%if strAcessoTipoRede = "10" then%>disabled<%end if%>  <%if strOriSol = "9" then%>disabled<%end if%> <%if strAcessoTipoRede = "4" then%>disabled<%end if%> Index=2 onclick="EsconderTecnologia(0);ResgatarTecVel();EsconderFabONT();EsconderTipoONT();MostrarVlan();">Cliente&nbsp;&nbsp;&nbsp;
				-->
				<!--Alterado por Fabio Pinho em 29/04/2016 - ver 1.0 - Fim-->

<!--
				<input type=radio name=rdoPropAcessoFisico value="TER"  <%if strAcessoTipoRede = "4" then%>checked<%end if%> Index=0	<%if strPropAcessoFisico = "TER" then Response.Write " checked " End if%> onclick="EsconderTecnologia(0);ResgatarTecVel();EsconderFabONT();EsconderTipoONT();MostrarVlan();">Terceiro&nbsp;&nbsp;&nbsp;
				<input type=radio name=rdoPropAcessoFisico value="EBT" <%if strAcessoTipoRede = "4" then%>disabled<%end if%> Index=1	<%if strPropAcessoFisico = "EBT" then Response.Write " checked " End if%> onclick="EsconderTecnologia(0);ResgatarTecVel();EsconderFabONT();EsconderTipoONT();MostrarVlan();">CLARO&nbsp;&nbsp;&nbsp;
				<input type=radio name=rdoPropAcessoFisico value="CLI" <%if strOriSol = "9" then%>disabled<%end if%> <%if strAcessoTipoRede = "4" then%>disabled<%end if%> Index=2	<%if strPropAcessoFisico = "CLI" then Response.Write " checked " End if%> onclick="EsconderTecnologia(0);ResgatarTecVel();EsconderFabONT();EsconderTipoONT();MostrarVlan();">Cliente&nbsp;&nbsp;&nbsp;
-->			
					
<%
' Execute the stored procedure to get the recordset
'set objRS = db.execute("CLA_sp_sel_SevFacilidadeTecnologia " & dblNroSev)
 sSql ="select cla_newtecnologia.newtec_id,cla_newtecnologia.newtec_nome,cla_newfacilidade.newfac_id,cla_newfacilidade.newfac_nome " 
     sSql = sSql + "from cla_assoc_tecnologiaFacilidade inner join cla_newtecnologia on cla_assoc_tecnologiaFacilidade.newtec_id = cla_newtecnologia.newtec_id " 
	 sSql = sSql + "inner join cla_newfacilidade	on cla_assoc_tecnologiaFacilidade.newfac_id = cla_newfacilidade.newfac_id where cla_newtecnologia.newtec_ativo = 'S' "
set objRS = db.execute(sSql)

' Initialize an array to hold the data
Dim dataArray()
Dim rowCount
rowCount = 0

' First, count the number of records
If Not objRS.Eof Then
    objRS.MoveFirst
    Do While Not objRS.Eof
        rowCount = rowCount + 1
        objRS.MoveNext
    Loop
End If

' Resize the array to hold the data
ReDim dataArray(rowCount - 1)

' Populate the array with data from the recordset
Dim i
dim strarr,strarr1,strarr2,strf
i = 0
strarr=""
strarr2=""
objRS.MoveFirst
For i = 0 To UBound(dataArray)
    ' Concatenate the values into a single string'
	if strarr <>  CStr(objRS("newtec_id").value)  or  strarr2 <>  CStr(objRS("newfac_id").value) then
		strarr = CStr(objRS("newtec_id").value)
		strarr1 = CStr(objRS("newtec_Nome").value)
		strarr2 = CStr(objRS("newfac_id").value)
		strf= strarr & "," & strarr1 & "," & strarr2 & ";"
	    dataArray(i) = strf	
	end if 	
    ' Move to the next record
    objRS.MoveNext
Next  

 Clean up
objRS.Close
Set objRS = Nothing

' Convert the array to a simple string for JavaScript
Dim dataString 
dataString = ""
For i = 0 To rowCount - 1 'UBound(dataArray)   	
    If i < UBound(dataArray) Then'
	     dataString =  dataString & dataArray(i)    		
    End If
Next
%>			
	<td nowrap colspan=2> <span ID =spnFacilidadeTecnologia  name=spnFacilidadeTecnologia >
					<select name="txtFacilidade" onchange='MostraTec(this.value, "<%=Replace(dataString, """", "\""")%>")'>
						<Option value="">:: FACILIDADE </Option>  
								<%
								  'set objRS = db.execute("CLA_sp_sel_SevFacilidade " & dblNroSev )
								  sSql ="select distinct cla_newfacilidade.newfac_id,cla_newfacilidade.newfac_nome " 
                                  sSql = sSql + "from cla_assoc_tecnologiaFacilidade " 
	                              sSql = sSql + "inner join cla_newfacilidade on cla_assoc_tecnologiaFacilidade.newfac_id = cla_newfacilidade.newfac_id where cla_newfacilidade.newfac_ativo = 'S' "
								  sSql = sSql + " order by cla_newfacilidade.newfac_nome "
                                  set objRS = db.execute(sSql)

								  dim regconta 
								  Dim lastFacID, lastFacName
								  Dim firstFacID
								  regconta = 0				  
								  firstFacID = ""
								  While not objRS.Eof
										if firstFacID <> objRS("newFac_id")  then
										   firstFacID = objRS("newFac_id")
										   lastFacID = objRS("newFac_id")
										   lastFacName = objRS("newFac_Nome")
										  Response.Write "<option value=""" & objRS("newFac_id") & """>" & objRS("newFac_Nome") & "</option>"
										 end if 
										objRS.MoveNext
								  Wend										   
								%>
					</Select>
					<br/>
				   <!-- <input type="text" class="text" name="txtFacilidade"  readonly="TRUE" value="<%=strTxtFacilidade %>">	-->
				
				<Select name=cboTecnologia id=cboTecnologia>
<!--					<Option value="">:: TECNOLOGIA </Option> -->
										<%
					'set objRS = db.execute("CLA_sp_sel_newTecnologia null,null,null")
					'set objRS = db.execute("CLA_sp_sel_SevFacilidadeTecnologia " & dblNroSev )
					'While not objRS.Eof
					'	strItemSel = ""
					'	'if Trim(dblTecId) = Trim(objRS("newTec_id")) and   then strItemSel = " Selected " End if
					'	 if firstFacID = objRS("newFac_id") then
					'	     Response.Write "<Option value=" & objRS("newTec_id") & strItemSel & ">" & objRS("newTec_Nome") & "</Option>"
					'	   end if
					'	objRS.MoveNext
					'Wend
					'strItemSel = ""
					%>
				</Select>				
			 </span>				
            <div id=divTecnologia name=divTecnologia style="display:none;POSITION:relative"></div>			  
			</td>

	</tr>
	<tr  class="clsSilver">
		<td id = tdRadio width=170px></td>
		<td colspan = 3><span ID =spnTipoRadio></span></td>
	</tr>
	
	<%'GPON%>
	<tr  class="clsSilver">
		<td id = tdONT width=170px></td>
		<td colspan = 2><span ID=spncboFabricante></span></td>
		<td colspan = 1><span ID=spncboTipoONT></span></td>
	</tr>

	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Vel do Acesso Físico</td>
		<td colspan=3><span id=spnVelAcessoFis name=spnVelAcessoFis>
			<!-- <select name="cboVelAcesso" style="width:150px" onChange="MostrarTipoVel(this)"> -->
			<select name="cboVelAcesso" style="width:150px" >
				<option ></option>
				<%
					'if Trim(dblTecId) <> "" then
					'	Set objRS = db.execute("CLA_sp_sel_AssocTecVeloc null," & dblTecId)
					'Else
						set objRS = db.execute("CLA_sp_sel_velocidade")
					'End if
					While Not objRS.eof
						strItemSel = ""
						if Trim(strVelAcesso) = Trim(objRS("Vel_ID")) then strItemSel = " Selected " End if
						Response.Write "<Option value='" & Trim(objRS("Vel_ID")) & "'" & strItemSel & ">" & objRS("Vel_Desc") & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				%>
			</select></span>
			
		</td>
	</tr>
	
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Provedor</td>
		<td colspan="4">
			<select name="cboProvedor">
				<option value=""></option>
				<%	
					'response.write "<script>alert('"&strAcessoTipoRede&"')</script>"
					'if isnull(strAcessoTipoRede) then
					'	strAcessoTipoRede = "null"
						'response.write "<script>alert('"&strAcessoTipoRede&"')</script>"
					'end if 
					'set objRS = db.execute("CLA_sp_sel_provedor 0,null,1, " & strAcessoTipoRede & ", " & strOriSol )
					'if dblNroSev <> "" then
						set objRS = db.execute("CLA_sp_sel_provedorTecnologia " & dblNroSev )
						
						While not objRS.Eof
							strItemSel = ""
							if Trim(dblProId) = Trim(objRS("Pro_ID")) then strItemSel = " Selected " End if
							'if  ( ( strAcessoTipoRede = "4" and Trim(objRS("Pro_ID")) = 136 ) or ( strAcessoTipoRede = "10" and Trim(objRS("Pro_ID")) = 143 ) )then 
							'strItemSel = " Selected " 
							'End if
							Response.Write "<Option value='" & Trim(objRS("Pro_ID")) & "'" & strItemSel & ">" & objRS("Pro_Nome") & "</Option>"
							objRS.MoveNext
						Wend
						strItemSel = ""
					'end if 
				%>
			</select>
		</td>
		
	</tr>
	
	<tr class=clsSilver2>
		<td width=170px >&nbsp;Endereço Origem&nbsp;</td>
		<td nowrap colspan=3>
			<font class=clsObrig>:: </font>PONTO&nbsp;
				<select name="cboTipoPonto" onChange="TipoOrigem(this.value)">
					<option value=""></option>
					<option value="I" <%if Trim(strTipoPonto) = "I" then Response.Write " selected " %>>CLIENTE</option>
					<option value="T" <%if Trim(strTipoPonto) = "T" then Response.Write " selected " %>>INTERMEDIÁRIO</option>
				</select>			
		</td>
	</tr>
	
	<%'Incluir tratamento para Bloquear apenas os dados do Endereço do Cliente %>
	<tr class="clsSilver">
		<td width=170px nowrap><span id=spnOrigem>&nbsp;&nbsp;&nbsp;Sigla Estação Origem(CNL)</span></td>
		<td colspan=3>
			<input type="text" class="text" name="txtCNLSiglaCentroCli"  maxlength="4" disabled="true"
			
			size="8" TIPO="A" value="<%=strCidSigla%>">

			&nbsp;Complemento
			<input type="text" class="text" name="txtComplSiglaCentroCli"  maxlength="3" disabled="true"
			
			size="10" TIPO="A" value="<%=strEscSigla%>">&nbsp;
						
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>UF</td>
		<td>
                        
			     <input type="text" class="text" name="cboUFEnd"  maxlength="2"
			    size=3" readonly="TRUE" value="<%=strUFEnd%>">
                       
                        
		
		</td>


			
				<td nowrap right><font class="clsObrig">:: </font>CEP&nbsp;</td>
				<td> <input type="text" class="text" name="txtCepEnd" readOnly="readOnly" value="<%=strCepEnd%>" maxlength="9" size="12">

			
			
			<!--Alterado por Fabio Pinho em 28/04/2016 - ver 1.0 - Fim-->

		
		
		
		 	<td nowrap right></td>
			<td> </td>
		 
		
			
	</tr>

	<tr class="clsSilver">
		<td nowrap><font class="clsObrig">:: </font>Município (CNL - ANATEL)</td>
		<td nowrap colspan=3 >
                    
			<input type=text size=5 maxlength=4 class=text name="txtEndCid" readonly="TRUE" value="<%=strEndCid%>">&nbsp;

			<input type=text size=70 readonly class=text readonly="TRUE" name="txtEndCidDesc" value="<%=strEndCidDesc%>">
		     
                    

                </td>
	</tr>
	
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Tipo do Logradouro</td>
		<td colspan=3>
                   
			<input type="text" class="text" name="cboLogrEnd"  maxlength="15" readonly="TRUE"
			size="15"  value="<%=strLogrEnd%>">
                    
                   
			
		</td>
		
	</tr>
	<tr class="clsSilver">
		<td><font class="clsObrig">:: </font>Nome Logr</td>
		<td nowrap>
			<input type="text" class="text" name="txtEnd"  readonly="TRUE" 
			
			value="<%=strEnd%>" maxlength="60" size="40">
		</td>
		<td width=170px><font class="clsObrig">:: </font> Número</td>
		<td>
			<input type="text" class="text" name="txtNroEnd" value="<%=strNroEnd%>"  readonly="TRUE" 
			 maxlength="10" size="10">
		</td>
	</tr>
	
	<tr class="clsSilver">
		<td>&nbsp;&nbsp;&nbsp;Complemento</td>
		<td colspan= 3 >
			<input type="text" class="text" name="txtComplEnd"  value="<%=strComplEnd%>"  maxlength="80" size="80" >
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Bairro</td>
		<td colspan = 3 >
			<input type="text" class="text" name="txtBairroEnd"
			 
			value="<%=strBairroEnd%>" maxlength="120" size="40" readonly="TRUE" >&nbsp;
		</td>			
			
		
	</tr>
	<tr class="clsSilver">
		<td colspan=4 align=right><span id=spnCEPSInstala></span></td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Contato</td>
		<td>
			<input type="text" class="text" name="txtContatoEnd" value="<%=strContatoEnd%>" maxlength="30" size="30">
		</td>
		<td><font class="clsObrig">:: </font>Telefone</td>
		<td >
			<input type="text" class="text" name="txtTelEndArea" value="<%=strTelEndArea%>"	maxlength="3" size="3" onkeyUp="ValidarTipo(this,0)">&nbsp;

			<input type="text" class="text" name="txtTelEnd" value="<%=strTelEnd%>"	maxlength="9" size="11" onkeyUp="ValidarTipo(this,0)">
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>CNPJ</td>
		<td colspan="4">
			<input type="text" class="text" name="txtCNPJ"  maxlength="14"
			
			size="16" onKeyUp="ValidarTipo(this,0)" value="<%=dblCNPJ%>" >&nbsp;(99999999999999)
		</td>
		<!--<td width=170px>&nbsp;&nbsp;&nbsp;</td>
	
		<td width=170px>&nbsp;&nbsp;&nbsp;</font>Email Contato Técnico</td>
	    <td>
	      <input type="text" class="text" name="txtEmailTec"  value="<%=strBairroEnd%>" maxlength="30" size="30">&nbsp;
	    </td>
		-->
	</tr>
	<tr class="clsSilver">
		<td width=170px>&nbsp;&nbsp;&nbsp;&nbsp;I.E.</td>
		<td >
			<input type="text" class="text" name="txtIE"  maxlength="15"
			<%if strIE <> "" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>
			size="17" onKeyUp="ValidarTipo(this,0)" value="<%=strIE%>" >
		</td>
		<td >&nbsp;&nbsp;&nbsp;I.M&nbsp;</td>
		<td>
			<input type="text" class="text" name="txtIM"  maxlength="15"
			<%if strIM <> "" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <%End if%>
			size="17" onKeyUp="ValidarTipo(this,0)" value="<%=strIM%>" >
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px nowrap>&nbsp;&nbsp;&nbsp;Proprietário do Endereço</td>
		
		<!--
		<td colspan="3">
			<input type="text" class="text" name="txtPropEnd"
			<%if strPropEnd <> "" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <%End if%>
			maxlength="55" size="50" value="<%=strPropEnd%>" >
		</td>-->

		<!--Alterado por Fabio Pinho em 28/04/2016 - ver 1.0 - Inicio-->
		<td colspan="3">
			<input type="text" class="text" name="txtPropEnd"
			<%if strPropEnd <> "" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <%End if%>
			maxlength="55" size="81" value="<%=strPropEnd%>" >
		</td>
		<!--Alterado por Fabio Pinho em 28/04/2016 - ver 1.0 - Fim-->

	
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class=clsObrig>:: </font><font color="#FF0000">Interface Cliente</font></td>
		<td colspan="3">
			<Select name="cboInterFaceEnd" >
				<Option value=""></Option>
				<%
				if strOriSol = "9" then
			        set objRS = db.execute("CLA_sp_sel_interface null , null , " & strOriSol)
				else
			      set objRS = db.execute("CLA_sp_sel_interface")
				end if 
					While not objRS.Eof
						strItemSel = ""
						if Trim(strInterFaceEnd) = Trim(objRS("ITF_Nome")) then strItemSel = " Selected " End if
						Response.Write "<Option value=""" & Trim(objRS("ITF_Nome")) &""" " & strItemSel & ">" & Trim(objRS("ITF_Nome")) & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				%>
			</Select>
		</td>
	</tr>
	<tr class="clsSilver2">
		<td width=170px><span id=spnDestino>&nbsp;&nbsp;&nbsp;Sigla Estação Entrega<br>&nbsp;&nbsp;&nbsp;do Acesso Físico</span></td>
		<td colspan=3 nowrap>
			<table border=0 cellspacing=0 cellpadding=0>
				<tr>
					<td>&nbsp;CNL</td>
					<td>&nbsp;Complemento</td>
					<td>&nbsp;Endereço de Entrega do Acesso Físico</td>
				</tr>
				<tr>
					<td><input type="text" class="text" name="txtCNLSiglaCentroCliDest"  maxlength="4" size="8" onKeyUp="ValidarTipo(this,1)"	value="<%=strCNLSiglaCli%>" onblur="CompletarCampo(this)" TIPO="A">&nbsp;</td>
					<td>&nbsp;<input type="text" class="text" name="txtComplSiglaCentroCliDest"  maxlength="6" size="10" onKeyUp="ValidarTipo(this,7)" value="<%=strCNLSiglaCli%>" onblur="CompletarCampo(this);ResgatarEstacaoDestino(document.Form2.txtCNLSiglaCentroCliDest,document.Form2.txtComplSiglaCentroCliDest)" TIPO="A">&nbsp;</td>
					<td>&nbsp;<TEXTAREA rows=2 cols=66 name="txtEndEstacaoEntrega" readonly tabIndex=-1></TEXTAREA></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class=clsObrig>:: </font><font color="#FF0000">Interface Claro Brasil</font></td>
		<td colspan="3">
			<Select name="cboInterFaceEndFis">
				<Option value=""></Option>
				<%'PRSS - 07/09/2005
				    if strOriSol = "9" then
				        set objRS = db.execute("CLA_sp_sel_interface null , null , " & strOriSol)
					else
				  set objRS = db.execute("CLA_sp_sel_interface")
					end if 
					While not objRS.Eof
						strItemSel = ""
						if Trim(strInterFaceEnd) = Trim(objRS("ITF_Nome")) then strItemSel = " Selected " End if
						Response.Write "<Option value=""" & Trim(objRS("ITF_Nome")) &""" " & strItemSel & ">" & Trim(objRS("ITF_Nome")) & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				%>
			</Select>
		</td>
	</tr>
	<tr>
		<td colspan=4>
			<span id=spnListaIdFis></span>
		</td>
	</tr>
	<tr class="clsSilver">
		<td colspan="4">
			<div id=divIDFis1 style="DISPLAY: none">
				<table width=100%>
					<tr>
						<td colspan=7>
							<iframe	id			= "IFrmIDFis1"
									name		= "IFrmIDFis1"
									width		= "100%"
									height		= "65px"
									frameborder	= "0"
									scrolling	= "auto"
									align		= "left">
							</iFrame>
						</td>
					</tr>
				</table>
			</div>
		</td>
	</tr>
	<tr class="clsSilver">
		<td colspan=4 align=right bgColor=#dcdcdc>
			<input type=button name=btnAddAcesso class=button value="Adicionar" onmouseover="showtip(this,event,'Adicionar/Atualizar um acesso da lista (Alt+A)');" onClick="CheckSevMestra()" accesskey="A" <%if strMudFisico="N" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>&nbsp;
			<input type=button name=btnRemAcesso class=button value="Remover" onmouseover="showtip(this,event,'Remove um acesso da lista (Alt+R)');" onClick="RemoverAcessoLista()" accesskey="R" <%if strMudFisico="N" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>  >&nbsp;
			
				<input type=button name=btnIDFis1 class=button value="Acesso Físico"  onClick="javascript:window.open('AcessoCompartilhar.asp?FlagOrigem=CLA2&txtNroSev='+document.Form1.txtNroSev.value,'janela','toolbar=no,statusbar=no,resizable=yes,scrollbars=YES,width=900,height=800,top=100,left=100')" onmouseover="showtip(this,event,'Procurar um id físico pré existente para o endereço atual (Alt+F)');" accesskey="F" <%if strMudFisico="N" or dblOriSolID = 9 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>  <%=bbloqueiaTipoRede%> >&nbsp; 
			
				<input type=button name=btnIDFis1 class=button value="Compartilhar" onClick="ProcurarIDFis(1)" onmouseover="showtip(this,event,'Procurar um id físico pré existente para o endereço atual (Alt+F)');" accesskey="F" <%if strMudFisico="N" or dblOriSolID = 9 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>  <%=bbloqueiaTipoRede%> >&nbsp;
						
			<span id="spnBtnLimparIdFis1"></span>&nbsp;
			<input type=button name=btnLimparAcesso class=button value="Limpar" onClick="LimparInfoAcesso()" accesskey="L" onmouseover="showtip(this,event,'Limpar dados do Acesso (Alt+L)');" <%if strMudFisico="N" then%> <%=readOnlyASMS%> <%else%> <%=bdesbloqueia%> <% End if%>  >&nbsp;
		</td>
	</tr>
	<tr class="clsSilver">
		<td colspan="4">
			<div id=divTronco2M style="DISPLAY: none">
				<table width=100%>
					<tr>
						<td colspan=7>
							<iframe	id			= "IFrmTronco2M"
									name		= "IFrmTronco2M"
									width		= "100%"
									height		= "65px"
									frameborder	= "0"
									scrolling	= "auto"
									align		= "left">
							</iFrame>
						</td>
					</tr>
				</table>
			</div>
		</td>
	</tr>
</Form>
</table>
<!-- fim novo ReCLA --> 

	
	
	
<table border=0 cellspacing="1" cellpadding="0" width="760" >
<Form name=Form3 method=Post>
<input type="hidden" name="hdnOriSol_ID" value="<%=strOriSol%>">
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUsuario>
<input type=hidden name=hdnOrigemEst>
<input type=hidden name=hdnCNLEstUsu>
<input type=hidden name=hdnComplEstUsu>
<input type=hidden name=hdnEstacaoAtual>
<input type=hidden name=hdnCtfcId value="<%=dblCtfcId%>" >
<input type=hidden name=hdntxtGICN value="<%=strUserNameGICN%>">
<input type=hidden name=hdntxtGICL value="<%=strUserNameGICLAtual%>">
<input type=hidden name=hdnCoordenacaoAtual>
<input type=hidden name=hdnNecessitaRecurso value="S"> <!-- Na ativação será sempre SIM -->
<input type=hidden name=hdnEmiteOTS>
<input type=hidden name=hdnReaproveitarFisico value="N"> <!-- Na ativação será sempre Não. Será modificado na Alteração -->
<input type=hidden name=hdnStrAcfId >





<table border=0 cellspacing="1" cellpadding="0" width="760" >
	<tr>
		<th colspan="4" >&nbsp;•&nbsp;Informações da Embratel</th>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Local de Entrega<br>&nbsp;&nbsp;&nbsp;do Acesso Lógico</td>
			<input type="Hidden" name="cboLocalEntrega">
		<td><input type="text" class="text" name="txtCNLLocalEntrega"  maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A">&nbsp;
		&nbsp;<input type="text" class="text" name="txtComplLocalEntrega"  maxlength="3" size="6" onKeyUp="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoUsu(document.Form3.txtCNLLocalEntrega,document.Form3.txtComplLocalEntrega,<%=dblUsuId%>,1);" TIPO="A">
		</td>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px" nowrap><font class="clsObrig">:: </font>Local de Configuração</td>
			<input type="Hidden" name="cboLocalConfig">
		<td><input type="text" class="text" name="txtCNLLocalConfig"  maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A">&nbsp;
		&nbsp;<input type="text" class="text" name="txtComplLocalConfig"  maxlength="3" size="6" onKeyUp="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoUsu(document.Form3.txtCNLLocalConfig,document.Form3.txtComplLocalConfig,<%=dblUsuId%>,2);" TIPO="A">
		</td>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Contato</td>
		<td width=50% >
			<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width=355px>
				<tr><td class="lightblue">&nbsp;
					<span id=spnContEndLocalInstala><%=strContEscEntrega%></span>
				</td></tr>
			</table>
		</td>
		<td align=right>Telefone</td>
		<td width=20%>
			<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="80%" >
				<tr><td class="lightblue">&nbsp;
					<span id=spnTelEndLocalInstala><%=strTelEscEntrega%></span>
				</td></tr>
			</table>
		</td>
	</tr>
</table>
<table  border=0 cellspacing="1" cellpadding="0" width="760" >
	<tr class="clsSilver">
		<th colspan=7 >&nbsp;•&nbsp;Coordenação Embratel</th>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Órgão de Venda</td>
		<td colspan="7" >
			<select name="cboOrgao">
				<option value=""></option>
				<%
					set objRS = db.execute("CLA_sp_sel_orgaovendas 0")
					While not objRS.Eof
						strItemSel = ""
						if Trim(dblOrgId) = Trim(objRS("Org_ID")) then strItemSel = " Selected " End if
						Response.Write "<Option value=" & objRS("Org_ID") & strItemSel & ">" & objRS("Org_Nome") & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				%>
			</select>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>UserName GIC-N</td>
		<td colspan="7">
			<input type="text" class="text" name="txtGICN"  value="<%=strUserNameGICN%>" maxlength="20" size="20" onblur="ResgatarUserCoordenacao(this)" >
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Nome GIC-N</td>
		<td colspan=2>
			<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width=355px >
				<tr><td class="lightblue">&nbsp;
					<span id=spnNomeGICN><%=strNomeGICN%></span>
				</td></tr>
			</table>
		</td>
		<td align=right >Ramal&nbsp;</td>
		<td colspan=3 align=left >

			<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="100px" >
				<tr><td class="lightblue">&nbsp;
					<span id=spnRamalGICN><%=strRamalGICN%></span>
				</td></tr>
			</table>


		</td>

	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>UserName GIC-L</td>
		<td colspan=6>
			<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="20%" >
				<tr>
					<td class="lightblue">&nbsp;<%=strUserNameGICLAtual%>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Nome GIC-L</td>
		<td width=355px colspan=2>
			<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="100%" >
				<tr><td class="lightblue">&nbsp;
					<span id=spnNomeGICL><%=strNomeGICL%></span>
				</td></tr>
			</table>
		</td>
		<td align=right >Ramal&nbsp;</td>
		<td colspan=3 align=left >
			<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="100px" >
				<tr><td class="lightblue">&nbsp;
					<span id=spnRamalGICL><%=strRamalGICL%></span>
				</td></tr>
			</table>
		</td>
	</tr>
</table>

<span id=spnGLA>
	<%if (Trim(Request.Form("hdnAcao")) = "Alteracao") and (objDicProp.Exists("CLI") or objDicProp.Exists("TER")) then
		Response.Write strHtmlGla
	  End if
	%>
</span>


<table border=0 cellspacing="1" cellpadding="0"width="760">
<tr><th colspan=2 >&nbsp;•&nbsp;Comunicação Interna</th></tr>
 <tr class=clsSilver>
	 <td width=170px >Status</td>
	 <td>
		 <select name="cboStatusSolic" style="width:320px">
		 	<option value=""></option>
			<%	Set objRS = db.execute("CLA_sp_sel_Status null,1")
				While Not objRS.Eof
				%>
				  <option value="<%=objRS("Sts_id")%>"
				
				<%if ucase(Trim(objRS("Sts_Desc"))) = "ENVIADO PARA O GLA" then
					if (Trim(Request.Form("hdnAcao")) = "Alteracao") or (Trim(Request.Form("hdnAcao")) = "Desativacao") or (Trim(Request.Form("hdnAcao")) = "Cancelamento") then
						Response.Write "selected"
					End If
				
				End If%>
				
					><%=ucase(objRS("Sts_Desc"))%>
			<%
				objRS.movenext
				Wend				
			%>
		 </select>
	</td>
</tr>
<tr>
	<th colspan="2">&nbsp;•&nbsp;Histórico</th>
</tr>
<tr class=clsSilver>
	<td width=170px>Motivo</td>
	<td>
		<textarea name="txtMotivo" cols="50" rows="3" onkeydown="MaxLength(this,300);"></textarea>
	</td>
</tr>
</table>
<%if Trim(Request.Form("hdnAcao")) = "Alteracao" then%>
<table cellspacing=1 cellpadding=1  width=760 border=0>
<%
  if dblSolId <> "" then
	'Set objRS = db.Execute("CLA_sp_sel_Status null," & dblSolId)
	Vetor_Campos(1)="adInteger,2,adParamInput,"
	Vetor_Campos(2)="adInteger,2,adParamInput," & dblSolId
	Vetor_Campos(3)="adInteger,2,adParamInput,1"
	Vetor_Campos(4)="adInteger,2,adParamInput,"

	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_StatusSolicitacao",4,Vetor_Campos)





	Set objRS = db.Execute(strSqlRet)
  End if
  
	blnCor = true
	strHtml = strHtml &  ""
	While Not objRS.Eof
		if blnCor then
			strHtml = strHtml &  "<tr class=clsSilver >"
			blnCor = false
		Else
			strHtml = strHtml &  "<tr class=clsSilver2>"
			blnCor = true
		End if
		strHtml = strHtml &  "<td width=15% nowrap >"& Formatar_Data(objRS("StsSol_Data")) &"</td>"
		strHtml = strHtml &  "<td nowrap >&nbsp;" & objRS("Pedido") & "</td>"
		strHtml = strHtml &  "<td width=20% >" & objRS("Usu_UserName") & "</td>"
		strHtml = strHtml &  "<td width=30% >"& objRS("Sts_Desc") & "</td>"
		strHtml = strHtml &  "<td width=35% >"& objRS("StsSol_Historico") &"</td>"
		strHtml = strHtml &  "</tr>"
		objRS.MoveNext
	Wend
	Response.Write strHtml%>
</table>
<%End if%>
<table  border=0 cellspacing="1" cellpadding="0" width="760" >
	<tr >
		<td align=center>
			<input type="button" class="button"  style="color:darkred;;font-weight:bold;width:180px" name="btnGravar" value=".::Gravar::." onclick="Gravar()" accesskey="I" onmouseover="showtip(this,event,'Gravar uma solicitação (Alt+I)');" <%if strOrigem="APG" and (trim(dblIdLogico) = "" or isnull(dblIdLogico)) and (strprocesso = "ATV" and stracao = "ALT") then%>Disabled<%end if%>>&nbsp;
			<input type="button" class="button" name="btnNovaSolic" value="Nova Solicitação" style="width:100px" onclick="javascript:window.location.replace('solicitacao.asp')" accesskey="N" onmouseover="showtip(this,event,'Nova solicitação (Alt+N)');">&nbsp;
			<%if Trim(Request.Form("hdnAcao")) = "Alteracao" then%>
			<input type=button	class="button" name=btnVoltar value=Voltar onclick="VoltarOrigem()" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">&nbsp;
			<%End if%>
			<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
		</td>
	</tr>
	<tr>
		<td>
			<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
		</td>
	</tr>
	<tr>
		<td>
			<font class="clsObrig">:: </font>Legenda: A - Alfanumérico;  N - Numérico;  L - Letra
		</td>
	</tr>
</table>

<%if Trim(Request.Form("hdnAcao")) <> "Alteracao" then%>
	<input type=hidden name="hdnStatus" value="38">
<%Else%>
	<input type=hidden name="hdnStatus" value="<%=dblStsId%>">
<%End if%>
</form>
</td>
</tr>
</table>
<iframe	id			= "IFrmProcesso"
		name        = "IFrmProcesso"
		width       = "0"
		height      = "0"
		frameborder = "0"
		scrolling   = "no"
		align       = "left">
</iFrame>
<iframe	id			= "IFrmProcesso2"
		name        = "IFrmProcesso2"
		width       = "0"
		height      = "0"
		frameborder = "0"
		scrolling   = "no"
		align       = "left">
</iFrame>
<iframe	id			= "IFrmProcesso3"
		name        = "IFrmProcesso3"
		width       = "0"
		height      = "0"
		frameborder = "0"
		scrolling   = "no"
		align       = "left">
</iFrame>

<SCRIPT LANGUAGE="JavaScript">
<%
Set objRS = db.execute("CLA_sp_sel_usuario 0,'" & Trim(strUserName) & "'")
if Not 	objRS.Eof And Not objRS.Bof then
	Response.Write "spnNomeGICL.innerHTML = '" & Replace(Trim(objRS("Usu_Nome")),"'","´")  & "';"
	Response.Write "spnRamalGICL.innerHTML = '" & Replace(Trim(Cstr("" & objRS("Usu_Ramal"))),"'","´")  & "';"
	Response.Write "document.forms[2].hdntxtGICL.value = '" & Trim(strUserName) & "';"
End if
%>
//Geral
with (document.forms[0])
{
	<%if Trim(Request.Form("hdnAcao")) <> "Alteracao" or Request.Form("hdnOEOrigem")="APG"then %>
		/*txtOrderEntry[0].readOnly = true
		txtOrderEntry[1].readOnly = true
		txtOrderEntry[2].readOnly = true*/
	<%Else%>
		
		ResgatarDesigServicoGravado(<%=dblSerId%>)
	<%End if%>
}
</script>
<!--Form que envia os dados para gravação-->
<TABLE border=0>
<tr><td>
<form method="post" name="Form4">
<input type=hidden id=strAcessoTipoRede name=strAcessoTipoRede>
<input type="hidden" name="hdnOriSol_ID" value="<%=strOriSol%>">
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUsuario>
<input type=hidden name=hdnOrigemEst>
<input type=hidden name=hdnCNLEstUsu>
<input type=hidden name=hdnComplEstUsu>
<input type=hidden name=hdnAntAcesso>
<input type=hidden name=hdnCboServico>
<input type=hidden name=hdnDesigServ>
<input type=hidden name=hdnTipoAcao value="<%=Request.Form("hdnAcao")%>" >
<input type=hidden name=hdnXml>
<input type=hidden name=hdnApg value="<%=Request.Form("hdnAcao")%>">
<input type=hidden name=hdnOrigem value="<%=strOrigem%>">
<input type=hidden name=hdnIdAcessoLogico value="<%=dblIdLogico%>">
<input type=hidden name=hdnSolId value="<%=dblSolId%>">
<input type=hidden name=hdn678 value="<%=Request.Form("hdn678")%>">
<%if Trim(Request.Form("hdnAcao")) = "Alteracao" then%>
<input type=hidden name=hdnTipoProcesso value=3>
<%Else%>
<input type=hidden name=hdnTipoProcesso value=1>
<%End if%>
<input type=hidden name=hdnVelIdServicoOld value="<%=dblVelServico%>">
<!-- DAVIF - 16/01/2006 -->
<input type=hidden name=hdnIdInterfaceAPG value="<%=dblSolAPGId%>">
<input type=hidden name=hdnStrAcfId >
<!-- LP -->
</form>
</td>
</tr>
</table>
<%
'*********** MIGRAÇÂO HFC BSoD **********
'response.write "<script>alert('hdnAcao=" & Request.Form("hdnAcao") & "')</script>"
strAcao = Request.Form("hdnAcao")
If (Not IsNull(dblIdLogico) or Trim(dblIdLogico) <> "") and strAcao="Alteracao" Then '1=1 then'

			'dblIdLogico = "6787336597"
			Set objMigr = db.execute("select TOP 1 cla_provedor.Pro_Nome,cla_acessofisico.Acf_ID,Rec_Id,cla_acessofisico.Pro_ID,Reg_ID,Acf_NroAcessoPtaEbt," &_
			"Acf_IDAcessoFisico,Fac_SVlan,Fac_PE,Fac_Vlan,Fac_Porta from cla_facilidade inner join cla_acessofisico on " &_
			"cla_facilidade.acf_id=cla_acessofisico.acf_id inner join cla_acessologicofisico on cla_acessologicofisico.acf_id=cla_acessofisico.acf_id " &_
			"inner join cla_provedor on  cla_provedor.pro_id=cla_acessofisico.pro_id where cla_acessofisico.pro_id in(136,137) and " &_
			"acf_dtdesatacessofis is null and acf_dtcancacessofis is null and alf_flag is null and acl_idacessologico='" & dblIdLogico & "'")
			
			If Not objMigr.Eof Then
			'response.write "<script>alert('strAcessoTipoRede=" & strAcessoTipoRede & "')</script>"
			  'Solicitações oriundas do SGA-Plus com marcação para BSOD NET deverão ser liberadas.
				'If Not (strAcessoTipoRede = 4 and objMigr("Pro_ID") = 136) Then
					response.write "<script>alert('ATENÇÃO \nEstão bloqueadas novas alterações para provedores HFC BSoD VIA e HFC BSoD NET. \nAssim, é necessário realizar a migração do acesso físico para HFC BSoD, através da funcionalidade de Migração, que será exibida após o OK desta mensagem. \nEm caso de dúvidas, contatar a equipe do SOE.')</script>"

'			    Set objRSMigr_Compart = db.execute("select count(acfl.acl_idacessologico) As Qtd_ACL from cla_acessofisico acf inner join cla_acessologicofisico acfl on acf.acf_id=acfl.acf_id " &_
'			                                       "where alf_flag is null and acf.acf_idacessofisico='" & objMigr("Acf_IDAcessoFisico") & "'")			
'			    If CInt(objRSMigr_Compart("Qtd_ACL"))>1 Then					
'				     response.write "<script>alert('Migração não permitida para Acesso Físico com compartilhamento.\n\nPara migrar o Acesso Físico, você deverá desativar ou remanejar os Acessos Lógicos vinculados, deixando apenas um Acesso Lógico vinculado ao Acesso Físico [" & objMigr("Acf_IDAcessoFisico") & "].')</script>" 
'			    End If		

   				response.write "<script>"   			
   				response.write "window.location='migracaoBsod.asp?acao=consulta&txtAcf_IdAcessoFisico=" & objMigr("Acf_IDAcessoFisico") & "'"
   				response.write "</script>"
   			'End If
			End If
			
'***********************lembrar de atualizar o regime de contrato do acf: cboRegimeCntr !!!
End If%>
</body>
<%DesconectarCla()%>
</html>
	