<%@ CodePage=65001 %>
<%

	Response.ContentType = "text/html; charset=utf-8"
	Response.Charset = "UTF-8"
'	- Sistema			: CLA
'	- Arquivo			: solicitacaoAprov.ASP
'	- Descrição			: Cadastra/Altera uma solicitação do Aprovisionamento no sistema CLA
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%

''@@ LPEREZ 13/10/2005
Dim strVisada			'Tipo Visada
Dim strGrupo			'Grupo Cliente
Dim strOriSol			'Origem Solicitacao
Dim strProjEspecial	' Projeto Especial
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
Dim strMudFisico            'Variavel para mudança no acesso físico
Dim readOnlyASMS
Dim bbloqueiaTipoRede

bbloqueia = "disabled=true"
bdesbloqueia =" "
readOnly = "true"
strIDLogico = ""
strMudFisico = "S"
readOnlyASMS = "readonly=""readonly"""

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
Dim strAcf_id
Dim strOrderEntry
'*********************************
' Good Início
''*********************************
DIm strXmlInfoAcesso
'*********************************
' Good fim
''*********************************

'<!--CH-41083NDL - inicio-->
Dim bbloqueiaSev				'Variavel de controle para Bloquear SEV quando a tela nao pode liberar a gravacao
bbloqueiaSev = "false"
'<!--CH-41083NDL - Fim-->

Set objDicProp = Server.CreateObject("Scripting.Dictionary")

'Monta o Xml de Acessos
%>
<!--#include file="../inc/xmlAcessos.asp"-->
<script>

function VerificaFacilidade(lstfac, facText) {
    var achou = false;
    var arrFac = lstfac.split(";");
    for (var f = 0; f < arrFac.length; f++) {
        if (arrFac[f] === facText) {
            achou = true;
        }
    }
    if (!achou) {
        alert("Facilidade não encontrada na Sev.");
        var fact = document.getElementsByName("txtFacilidade")[0];
        fact.focus();
    }
}


var xmlhttp = null;
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

function EsconderFabONT()
{
	document.getElementById("spncboFabricante").innerHTML = "";
}

function EsconderTipoONT()
{
	document.getElementById("spncboTipoONT").innerHTML = "";
}
</script>
<!--'*********************************
' Good Início
''*********************************-->
<script language="javascript">
    function MostraTec(idFac, strTecnologia) {
        var cboTecnologia = document.getElementsByName("cboTecnologia")[0];
        cboTecnologia.options.length = 0;
        var arrTec = strTecnologia.split(";");
        var optionTec = new Option(':: TECNOLOGIA ', '')
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
Function MontarXmlInfoAcesso(objXml, objRSCli, objRSolic)
    Dim objNodeAcesso
    Dim objElemento
    Dim objDicAcf
    Dim objDictionary
    Dim intIndice
    Dim strTel, strTelLocal, strTelArea
    Dim strEndEscEntrega
    Dim objRSEscEntrega


    Set objDicAcf = Server.CreateObject("Scripting.Dictionary")
    Set objDictionary = Server.CreateObject("Scripting.Dictionary")
	    
    ' Create the root element if it doesn't exist
    If objXml Is Nothing Then
        Set objXml = Server.CreateObject("Microsoft.XMLDOM")
        'objXml.appendChild objXml.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
        'objXml.appendChild objXml.createElement("Root") ' Create a root element
    End If
    objXml.async = False
    objXml.loadXML "<?xml version='1.0' encoding='UTF-8'?><xDados></xDados>" ' Initialize with a root element

    ' Create child elements for xmlDados
    intIndice = 1
		
	
    ' Create the InfoAcesso node using createNode
    Set objNodeAcesso = objXml.createNode(1, "Acesso", "")

    ' Append the new node to the document element
    objXml.documentElement.appendChild objNodeAcesso	

    'While Not objRS.EOF
	Call AddElemento(objXml, objNodeAcesso, "intIndice", intIndice)		
	dim valor 
	valor = ""
	if Trim(objRSSolic("cid_sigla")) <> "" then
	valor = Trim(objRSSolic("cid_sigla"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "txtCNLSiglaCentroCli", valor)

    valor = ""
	if Trim(objRSSolic("esc_sigla")) <> "" then
	valor = Trim(objRSSolic("esc_sigla"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "txtComplSiglaCentroCli",valor)
		
    valor = ""
	if Trim(objRSCli("PRE_NOMELOGR")) <> "" then
	valor = Trim(objRSCli("PRE_NOMELOGR"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "txtEnd",valor)

    valor = ""
	if Trim(objRSCli("ACP_COMPLEND")) <> "" then
	valor = Trim(objRSCli("ACP_COMPLEND"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "txtComplEnd", valor)

	valor = ""
	if Trim(objRSCli("PRE_BAIRRO")) <> "" then
	valor = Trim(objRSCli("PRE_BAIRRO"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "txtBairroEnd", valor)
	
	valor = ""
	if Trim(objRSCli("PRE_COD_CEP"))<> "" then
	valor = Trim(objRSCli("PRE_COD_CEP"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "txtCepEnd", valor)
	strXml = FormatarXml(objXml)

	valor = ""
	if Trim(objRSCli("EST_SIGLA")) <> "" then
	valor = Trim(objRSCli("EST_SIGLA"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "cboUFEnd", valor)
	
	valor = ""
	if Trim(objRSCli("PRE_NROLOGR")) <> "" then
	valor = Trim(objRSCli("PRE_NROLOGR"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "txtNroEnd", valor)
	
	valor = ""
	if Trim(objRSCli("TPL_SIGLA")) <> "" then
	valor = Trim(objRSCli("TPL_SIGLA"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "cboLogrEnd", valor)

	valor = ""
	if Trim(objRSCli("CID_SIGLA")) <> "" then
	valor = Trim(objRSCli("CID_SIGLA"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "txtEndCid", Trim(objRSCli("CID_SIGLA")))
	
	valor = ""
	if Trim(objRSCli("LOCALIDADE")) <> "" then
	valor = Trim(objRSCli("LOCALIDADE"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "txtEndCidDesc", valor)		

	' objRSsolic

	valor = ""
	if Trim(objRSSolic("Aec_CNPJ")) <> "" then
	valor = Trim(objRSSolic("Aec_CNPJ"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "txtCNPJ", valor)
	
	valor = ""
	if Trim(objRSSolic("Aec_Contato")) <> "" then
	valor = Trim(objRSSolic("Aec_Contato"))
	end if		
	Call AddElemento(objXml, objNodeAcesso, "txtContatoEnd",valor)
	
	valor = ""
	if Trim(objRSSolic("Aec_Telefone"))<> "" then
	valor = mid(Trim(objRSSolic("Aec_Telefone")),1,2)
	end if		
	Call AddElemento(objXml, objNodeAcesso, "txtTelEndArea",valor)

	'#Raio X#'
    valor = ""
	if Trim(objRSSolic("Aec_Telefone")) <> "" then
	valor = Trim(objRSSolic("Aec_Telefone"))
	end if

	if len(valor) = 10 then
	'#Raio X#'	
	Call AddElemento(objXml, objNodeAcesso, "txtTelEnd",Right(valor,8))					
	'#Raio X#'
	else
	'#Raio X#'	
	Call AddElemento(objXml, objNodeAcesso, "txtTelEnd",Right(valor,9))										
	'#Raio X#'
	end if 		

	valor = ""
	if Trim(objRSSolic("Aec_IE")) <> "" then
	valor = Trim(objRSSolic("Aec_IE"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "txtIE", valor)
	
	valor = ""
	if Trim(objRSSolic("Aec_IM"))<> "" then
	valor = Trim(objRSSolic("Aec_IM"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "txtIM", valor)

	if Trim(objRSSolic("Interface_Cliente"))<> "" then
	valor = Trim(objRSSolic("Interface_Cliente"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "cboInterFaceEnd", valor)
	
        if Trim(objRSSolic("Interface_Embratel"))<> "" then
	   valor = Trim(objRSSolic("Interface_Embratel"))
	else
	   valor = Trim(objRSSolic("Interface_Cliente"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "cboInterFaceEndFis", valor)	
	
	valor = ""
	if Trim(objRSSolic("Aec_PropEnd")) <> "" then
	valor = Trim(objRSSolic("Aec_PropEnd"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "txtPropEnd", valor)

	valor=""
	if Trim(objRSSolic("Acl_IDAcessoLogico")) <>"" Then
	valor =Trim(objRSSolic("Acl_IDAcessoLogico"))
	end if
	Call AddElemento(objXml, objNodeAcesso, "hdnAcfId", valor)

	set objDictionary = Nothing
	Set objNodeAcesso = Nothing
	Set objElemento = Nothing
		
	Set MontarXmlInfoAcesso = objXml

End Function
'*********************************
' Good Fim
''*********************************

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

Set objRS = db.execute("CLA_sp_sel_ConfigCtf null," & dblCtfcId)

if not objRS.Eof then
	if trim(objRS("Cfg_RedirecionamentoCarteira")) = "1" Then
		strObrigaGla = objRS("Cfg_RedirecionamentoCarteira")
	else
		strObrigaGla = 0 
	end if
Else
	strObrigaGla = 0
End if

if not objRS.Eof and not objRS.Bof then
	strPossuiAvaliador = objRS("Cfg_Avaliador")
Else
	strPossuiAvaliador = 0
End if

	strOrigem = Request.Form("hdnOEOrigem")
	dblAprovisiId = Trim(Request.Form("hdnAprovisiId"))
	
	dblSolId = ""
	
	If trim(strOrigem) = "Aprov" Then
		if Trim(Request.Form("hdnOrigem")) = "Recadastro" then
			Set objRSSolic = db.execute("CLA_sp_sel_AprovisionadorREC " & dblAprovisiId)
			
		else
			
			Set objRSSolic = db.execute("CLA_sp_sel_Aprovisionador " & dblAprovisiId)			
			
		end if 
'*********************************
' Good Início
''*********************************
		set objXmlInfoAcesso = Server.CreateObject("Microsoft.XMLDOM")
'*********************************
' Good Fim
''*********************************		
		If Not objRSSolic.eof or Not objRSSolic.bof Then
			'Povoamento:
			if Trim(objRSSolic("Acl_IDAcessoLogico")) <> "" then
				strIDLogico 		= Trim(Cstr(objRSSolic("Acl_IDAcessoLogico")))
			end if
			strIDSol 			= Trim(objRSSolic("Sol_ID"))
			strOriSol 			= Trim(objRSSolic("Orisol_ID"))
			strOriDesc 			= Trim(objRSSolic("Orisol_Descricao"))
			strPov_Razao 		= Ucase(Trim(objRSSolic("Cli_Nome")))
		 	strPov_CC 			= Trim(objRSSolic("Cli_CC"))
			strPov_SubCC 		= Trim(objRSSolic("Cli_SubCC"))
			strPov_ContratoServ = Trim(objRSSolic("Acl_TipoContratoServico"))
			strPov_ContatoCli 	= Ucase(Trim(objRSSolic("Aec_Contato")))
			strPov_Tel 			= Trim(objRSSolic("Aec_Telefone"))
			strPov_CNPJ 		= Trim(objRSSolic("Aec_CNPJ"))
			strPov_PropEnd 		= Ucase(Trim(objRSSolic("Aec_PropEnd")))
			strPov_End 			= Ucase(Trim(objRSSolic("End_NomeLogr")))
			strPov_VelServ		= Ucase(Trim(objRSSolic("vel_desc")))
			strPov_DesigServ 	= Trim(objRSSolic("Acl_DesignacaoServico"))
			dblOriSolID 		= Trim(objRSSolic("OriSol_ID"))
			
			strEndereco_padronizado = Trim(objRSSolic("ENDERECO_PADRONIZADO"))
			strTipo_Alteracao		= Trim(objRSSolic("TIPO_ALTERACAO"))
			strCorrecao_OTS		    = Trim(objRSSolic("CORRECAO_OTS"))
			
			Indicador_MUD_CADASTRAL = Trim(objRSSolic("Indicador_MUD_CADASTRAL")) 
			
			'response.write "<script>alert('"&strTipo_Alteracao&"')</script>"
			
			'Demais Dados
			dblSolId = trim(objRSSolic("sol_id"))
			dblIdLogico = trim(objRSSolic("acl_idacessologico"))
			strAcao = objRSSolic("Acao")
			strSolAcessoID = objRSSolic("Aprovisi_Id")
			
			dblNroSev			= Trim(objRSSolic("Sol_SevSeq"))
			
			'response.write "<script>alert('"&dblOriSolID&"')</script>"
			
			'''if dblOriSolID = "10" then 
				'Localiza informações da solução SSA
				Vetor_Campos(1)="adInteger,2,adParamInput," & dblNroSev
				Vetor_Campos(2)="adInteger,2,adParamOutput,0"
				'Vetor_Campos(3)="adWChar,30,adParamOutput,null"
				'Call APENDA_PARAM("CLA_sp_sel_solucao_ssa",2,Vetor_Campos)
				Call APENDA_PARAM("CLA_sp_sel_historico_solucao_ssa",2,Vetor_Campos)
				
				Set objRSCli = ObjCmd.Execute
				DBAction = ObjCmd.Parameters("RET").value
				
				'response.write "<script>alert('"& Cint("0" & DBAction) &"')</script>"
				
				
				if Cint("0" & DBAction) <> "0" then
					Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
					
					desativar_gravar = true
					'response.write "<script>alert('"&desativar_gravar&"')</script>"
					
					
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
					
					'response.write "<script>alert('"&strEnd&"')</script>"
					'response.write "<script>alert('"&strComplEnd&"')</script>"
					'response.write "<script>alert('"&strBairroEnd&"')</script>"
					'response.write "<script>alert('"&strCepEnd&"')</script>"
					'response.write "<script>alert('"&strUFEnd&"')</script>"
					'response.write "<script>alert('"&strNroEnd&"')</script>"
					'response.write "<script>alert('"&strEndCid&"')</script>"
					'response.write "<script>alert('"&strEndCidDesc&"')</script>"
					
					'response.write "<script>alert('"&strEnd&"')</script>"
				
				End if
			
			''''else
			
			'''		desativar_gravar = false
					
			'''		strHigienizado = "SIM"
					
					'#Raio X#'
			'''		strEnd			= Trim(objRSSolic("End_NomeLogr"))			'Nome do logradouro
			'''		strComplEnd		= Trim(objRSSolic("Aec_Complemento"))		'Complemento do logradouro
			'''		strBairroEnd	= Trim(objRSSolic("End_Bairro"))			'Bairro do logradouro
			'''		strCepEnd		= Trim(objRSSolic("End_CEP"))				'CEP do logradouro
					
					'#Raio X#'
					strContatoEnd	= Trim(objRSSolic("Aec_Contato"))			'Contato do logradouro
					'#Raio X#'
					if len(Trim(objRSSolic("Aec_Telefone"))) = 10 then
					'#Raio X#'	
						strTelEnd		= Right(Trim(objRSSolic("Aec_Telefone")),8)		
					'#Raio X#'
					else
					'#Raio X#'	
						strTelEnd		= Right(Trim(objRSSolic("Aec_Telefone")),9) 'Telefone  do logradouro
					'#Raio X#'
					end if 			
						
					strTelEndArea   = mid(Trim(objRSSolic("Aec_Telefone")),1,2)'Area Telefone  do logradouro
			'''		strUFEnd		= Trim(objRSSolic("Est_Sigla"))	 			'UF do logradouro
			'''		strNroEnd		= Trim(objRSSolic("End_NroLogr"))			'Número do logradouro
			'''		strLogrEnd		= Trim(objRSSolic("Tpl_Sigla"))				'Sigla do logradouro
			'''		strEndCid		= Trim(objRSSolic("CNL"))					'Sigla da cidade do logradouro
			'''		strEndCidDesc	= Trim(objRSSolic("Cidade"))				'Decrição da cidade do logradouro
	
			
			
			''' end if 
			
			
			strOrderEntrySis 	= StrOrigem
			strRazaoSocial		= Trim(objRSSolic("Cli_Nome"))
			strNomeFantasia 	= Trim(objRSSolic("Cli_NomeFantasia"))
			strContaSev			= Trim(objRSSolic("Cli_CC"))
			strSubContaSev		= Trim(objRSSolic("Cli_SubCC"))
			dblCNPJ 			= Trim(objRSSolic("Aec_CNPJ"))
			strAclDesigSite		= Trim(objRSSolic("ACL_DESIG_SITE"))
			
			'PABX VIRTUAL BROADSOFT / IMS
			strPabx_Virtual     = Trim(objRSSolic("PABX_VIRTUAL"))
			
			IF ISNULL(strPabx_Virtual) THEN
				strPabx_Virtual = ""
			END IF

			IF strPabx_Virtual = "NULL" THEN
				strPabx_Virtual = ""
			END IF
			
			
			strconta15 = strContaSev + strSubContaSev
			
				'Set obj035 = CreateObject("EOL.CLE035")
				Set Tronco = CreateObject("EOL.CLESP22O")	
  				Tronco.ServerAddress = "ETBPRDBA.NT.EMBRATEL.COM.BR:1971@RPC/SRVPCLE/CALLNAT" '10.4.2.1:5026@RPC/SRVPCLE/CALLNAT' //Prod  //Desenv 10.2.4.83:5024@RPC/SRVCLE/CALLNAT
  				Tronco.Logon
			

				'Set Tronco = CreateObject("EOL.CLESP22O")	
				Tronco.CLESP220 strconta15, strRazaoSocial , strNomeFantasia , SEGMENTO	, PORTE , COD-RETORNO , TXT-MSG
			
			'response.write "<script>alert('"&strRazaoSocial&"')</script>"
			'response.write "<script>alert('"&strNomeFantasia&"')</script>"
			'response.write "<script>alert('"&SEGMENTO&"')</script>"
			'response.write "<script>alert('"&PORTE&"')</script>"
			'response.write "<script>alert('"&COD-RETORNO&"')</script>"
			'response.write "<script>alert('"&TXT-MSG&"')</script>"
			
			
			
			strIE  				= Trim(objRSSolic("Aec_IE"))
			strIM   			= Trim(objRSSolic("Aec_IM"))
			
			'Código SAP
			strCodSap = Trim(objRSSolic("esc_cod_sap"))
			
			'Informações da OE
			strOrderEntryAno	= Trim(objRSSolic("OE_Ano"))
			strOrderEntryNro	= Trim(objRSSolic("OE_Numero"))
			strOrderEntryItem	= Trim(objRSSolic("OE_Item"))
			
			strIA				= Trim(objRSSolic("IA"))
 
			if dblOriSolID = 9 then
				strOrderEntrySis = "ASMS"
			end if 
						
			strOrderEntry = strOrderEntrySis + strOrderEntryAno + strOrderEntryNro + strOrderEntryItem
			
			if dblOriSolID =  10 then
				strOrderEntrySis = "CFD"
				strOrderEntry = strIA
			end if 
						
			
			
			Set objOriSol = db.execute("select oriSol_id from cla_origemSolicitacao where OriSol_Descricao = '" & Trim(objRSSolic("OriSol_Descricao")) & "'")
			
			If Not objOriSol.Eof Then
				strOriSol				= Trim(objOriSol("OriSol_ID"))
			End If
			
			''Busca Codigo da Velocidade do Servico
			Set objVelServ = db.execute("CLA_sp_sel_Velocidade null,'" & Trim(objRSSolic("vel_desc")) & "'")
			
			If Not objVelServ.Eof Then
				dblVelServico	= Trim(objVelServ("Vel_ID"))
			End If
			''
			strDesignacaoServico    = Trim(objRSSolic("acl_DesignacaoServico"))
	
			strTipoContratoServico	= Trim(objRSSolic("Acl_TipoContratoServico"))
			strNroContrServico		= Trim(objRSSolic("Acl_NContratoServico"))
	
			if dblIdLogico <> "" then
				dblDesigAcessoPri		= Right(dblIdLogico,len(dblIdLogico)-3)
			End if
	
			strDtEntrAcesServ	= Formatar_Data(Trim(objRSSolic("Acl_DtDesejadaEntregaAcessoServico")))
			strDtIniTemp		= Formatar_Data(Trim(objRSSolic("Acl_DtIniAcessoTemp")))
			strDtFimTemp		= Formatar_Data(Trim(objRSSolic("Acl_DtFimAcessoTemp")))
	
			'response.write "<script>alert('"&Trim(objRSSolic("Ser_Desc"))&"')</script>"
			
			''Busca Codigo do Servico
			Set objServ = db.execute("CLA_sp_sel_Servico null, null, '" & Trim(objRSSolic("Ser_Desc")) & "'")
	
			If Not objServ.Eof Then
				dblSerId	= Trim(objServ("Ser_ID"))
			End If
			
			'response.write "<script>alert('"&dblSerId&"')</script>"
			
			strDtPedido = right("0" & day(now),2) & "/" & right("0" & month(now),2) & "/" & year(now)
			strDtPrevEntrAcesProv = now() +  30
			strDtPrevEntrAcesProv = right("0" & day(strDtPrevEntrAcesProv),2) & "/" & right("0" & month(strDtPrevEntrAcesProv),2) & "/" & year(strDtPrevEntrAcesProv)
	
			strUsernamecadastrador = Trim(objRSSolic("Username_Cadastrador"))
			strTelefonecadastrador = Trim(objRSSolic("Telefone_Cadastrador"))
	
			strTipoPonto	= "I"
			
			
			
			strInterFaceEnd = Trim(objRSSolic("Interface_Cliente"))		'Interface do logradouro(Cliente)
			strInterfaceEbt = Trim(objRSSolic("Interface_Embratel"))	'Interface Embratel
	
			strPropEnd 		= Trim(objRSSolic("Aec_PropEnd"))
			
			' cidade da estacao
			strCidSigla		= Trim(objRSSolic("cid_sigla"))
			
			' sigla da estacao
			strEscSigla		= Trim(objRSSolic("esc_sigla"))
			
			' velocidade total
			strVelocidadeTotal		= Trim(objRSSolic("velocidade_total"))
			
			'BSOD:
			strAcessoMigracao	= Trim(objRSSolic("Acesso_Migracao"))
			strAcessoTipoRede	= Trim(objRSSolic("Acesso_TipoRede"))
			Indicador_Alterar   = Trim(objRSSolic("Indicador_Alterar"))
			Id_Acesso			= Trim(objRSSolic("Id_Acesso"))

			if ( strAcessoTipoRede = "10" ) then
				bbloqueiaTipoRede = "disabled=true"
			else 
				bbloqueiaTipoRede = " "
			end if 
			
			if objRSSolic("OriSol_ID") = 9 then
			'response.write "<script>alert('"&objRSSolic("Id_Acesso")&"')</script>"
			  Set objRS = db.execute("CLA_sp_sel_realidadeIP_compartilhamento """ & objRSSolic("Id_Acesso") & """")
			  if not objRS.Eof then	
			  		strAcf_id = Trim(objRS("acf_id")) 
					strMudFisico = "N"
					'response.write "<script>alert('"& trim(strAcf_id) &"')</script>"
					'response.write "<script>alert('"&strAcf_id&"')/script>"
			  end if
			  
			  if ( strAcao = "ALT" and ( isnull(Indicador_Alterar) or Indicador_Alterar = "" ) ) then
'			  'response.write "<script>alert('yyy="&objRSSolic("Id_Acesso")&"')</script>"
			  		strMudFisico = "N"
			  end if
			End if
			'GPON
			strVel_Voz	= Trim(objRSSolic("Vel_Voz"))
			strDados	= Trim(objRSSolic("Dados"))
			
			'''
			''' Inicio da implementação redirecionamento GLA para Interface
			'''
			
			
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
			
			
			'''
			''' Fim da implementação redirecionamento GLA para Interface
			'''
'*********************************
' Good Início
''*********************************
				Set objXmlDadosAcesso = MontarXmlInfoAcesso(objXmlInfoAcesso,objRSCli,objRSSolic)
				strXmlInfoAcesso = FormatarXml(objXmlDadosAcesso)
				' Output the XML string to JavaScript
				Response.Write("<script type='text/javascript'>")
				Response.Write("var xmlData = '" & Replace(strXmlInfoAcesso, "'", "\'") & "';") ' Escape single quotes
				Response.Write("</script>")
				
				'Response.Write("<h2>Serialized XML Output:</h2>")
				'Response.Write("<pre>" & Server.HTMLEncode(strXmlInfoAcesso) & "</pre>")	
				'Response.Write("<pre> cli:" & objRSSolic("Aec_PropEnd") & "</pre>")	
				'Response.End					
				
'*********************************
' Good Fim
''*********************************			
						
		End If
		
	End If


	If Trim(Request.Form("hdnAcao")) = "Alteracao" then
	
		If Trim(Request.Form("hdnAcao")) = "Alteracao"  And (dblSolId = "" or dblSolId = "0") Then
			Response.Write "<script language=javascript>alert('Solicitação indisponível.');window.location.replace('main.asp');</script>"
			Response.End	
		End if
	
		if not isnull(dblSolId) then
			'Adicionado PRSS 29/01/2007:
			strSQL = "select Ctf_ID from cla_solicitacaoctf WITH (NOLOCK) where sol_id = " & dblSolId
		    Set objRSCtfc = db.execute(strSQL)
			dblCtfcId = objRSCtfc("Ctf_ID")
		end if
	
		if objRSSolic.Eof then
			Response.Write "<script language=javascript>alert('Solicitação indisponível.');window.location.replace('main.asp');</script>"
			Response.End
		End if
	
		if not isnull(Trim(objRSSolic("Acl_IDAcessoLogico"))) then 
			'@@Davif - Interface APG
			If Trim(Request.Form("hdnAcao")) = "Alteracao" Then
					dblIdLogico			=	Trim(objRSSolic("Acl_IDAcessoLogico"))
			End If
			
			varRSVerFis = ""
			Set objRSVerFis = db.execute("SELECT ACF_ID FROM CLA_ACESSOLOGICOFISICO WHERE ACL_IDACESSOLOGICO = " & dblIdLogico)
			
			While Not objRSVerFis.Eof	
				varRSVerFis = objRSVerFis("ACF_ID")&","&varRSVerFis 
				objRSVerFis.movenext
			Wend				
			
			
			'Xml com os pontos
			Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
			objXmlDados.loadXml("<xDados/>")
		
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
			
			'Set doc = server.CreateObject("Microsoft.XMLDOM")
			'doc.loadXml(strXmlAcesso)
			'doc.save(Server.MapPath("Teste.xml"))
		End if
	
	End If
	
%>
<script language='javascript' src="../javascript/solicitacao.js"></script>
<script language='javascript' src="../javascript/xmlObjects.js"></script>
<script language='javascript' src="../javascript/help.js"></script>

<script language='javascript'>

    var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
    var objXmlAcessoFisComp = new ActiveXObject("Microsoft.XMLDOM")
    var objXmlApg = new ActiveXObject("Microsoft.XMLDOM")
    var objXmlFisicoAntigo = new ActiveXObject("Microsoft.XMLDOM")

    objXmlGeral.preserveWhiteSpace = true
    objXmlApg.preserveWhiteSpace = true
        <%if Trim(Request.Form("hdnAcao")) = "Alteracao" then %>
	var intIndice = <%=intIndice%>
<% Else %>
	var intIndice = 0
        <% End If%>


            function Message(objXmlRet) {

                var intRet = window.showModalDialog('Message.asp', objXmlRet, 'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	if (intRet != "")
	{
                    spnSolId.innerHTML = intRet
                    document.Form3.txtGICN.value = ""
                    document.Form3.hdntxtGICN.value = ""

                    //Qdo. for processo de alteração Volta para tela Inicial da solictação
		if	(document.Form4.hdnTipoProcesso.value == 3 || document.Form4.hdnOrigem.value == "APG" ||  document.Form4.hdnOrigem.value == "Aprov")
		{
                        VoltarOrigem()
                    }
                }
            }
    //'*********************************
    //' Good Início
    //''*********************************
    function Messagenew(objXmlRet) {

        var intRet = window.showModalDialog('Messagenew.asp', objXmlRet, 'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	if (intRet != "")
	{
            spnSolId.innerHTML = intRet
            document.Form3.txtGICN.value = ""
            document.Form3.hdntxtGICN.value = ""

            //Qdo. for processo de alteração Volta para tela Inicial da solictação
		if	(document.Form4.hdnTipoProcesso.value == 3 || document.Form4.hdnOrigem.value == "APG" ||  document.Form4.hdnOrigem.value == "Aprov")
		{
                VoltarOrigem()
            }
        }
    }





    function CarregarInfo() {

        objXmlInfoAcesso.onreadystatechange = CheckStateXmlInfo;
        objXmlInfoAcesso.resolveExternals = false;
    }

function CheckStateXmlInfo()
{
        var state = objXmlInfoAcesso.readyState;

  if (state == 4)
  {

            var err = objXmlInfoAcesso.parseError;
    if (err.errorCode != 0)
    {
                alert(err.reason)
            }
    else
    {
                CarregarDocInfo()
            }
        }
    }
function CarregarDocInfo()
{
        document.onreadystatechange = CheckStateDocInfo;
        document.resolveExternals = false;
    }

function CheckStateDocInfo()
{
        var state = document.readyState;

  if (state == "complete")
  {
            var err = objXmlInfoAcesso.parseError;
    if (err.errorCode != 0)
    {
                alert(err.reason)
            }
    else
    {
                //if (<%strXmlInfoAcesso%> != "")
                //{
                objXmlInfoAcesso.loadXML('<%=strXmlInfoAcesso%>')
                //}
            }
        }
    }

    //'*********************************
    //' Good Fim
    //''*********************************

function CarregarLista()
{

        objXmlGeral.onreadystatechange = CheckStateXml;
        objXmlGeral.resolveExternals = false;
        if ('<%=intAcesso%>' != '') {

            objXmlGeral.loadXML("<%=strXmlAcesso%>")
        } else {
            var objXmlRoot = objXmlGeral.createNode("element", "xDados", "")
            objXmlGeral.appendChild(objXmlRoot)
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


    function mostraWait(esconde) {
	with (document.forms[0])
	{
            if (esconde == '1') {
                divWait.style.display = 'none';
            } else {
                divWait.style.display = 'inline';
            }
        }
    }

    function mostraSistemaWait(esconde) {
	with (document.forms[0])
	{
            if (esconde == '1') {
                divSistemaWait.style.display = 'none';
            } else {
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

                    <%if Trim(Request.Form("hdnAcao")) = "Alteracao" or Request.Form("hdnOEOrigem") <> "APG" then %>
                        ResgatarDesigServicoGravado(<%=dblSerId %>)
                        <% End if%>
	}
        }
    }

function CarregarListaAnterior()
{

        objXmlFisicoAntigo.onreadystatechange = CheckStateXmlAnterior;
        objXmlFisicoAntigo.resolveExternals = false;
        if ('<%=intAcesso%>' != '') {
            objXmlFisicoAntigo.loadXML("<%=strXmlAcesso%>")
        } else {
            var objXmlRoot = objXmlFisicoAntigo.createNode("element", "xDados", "")
            objXmlFisicoAntigo.appendChild(objXmlRoot)
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
                    <%if Trim(Request.Form("hdnAcao")) = "Alteracao"  or Request.Form("hdnOEOrigem") <> "APG" then %>
                        ResgatarDesigServicoGravado(<%=dblSerId %>)
                        <% End if%>
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
        var objXmlRoot = objXmlApg.createNode("element", "xDados", "")
        objXmlGeral.appendChild(objXmlRoot)
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
                    <%if Trim(Request.Form("hdnAcao")) = "Alteracao" or Request.Form("hdnOEOrigem") <> "APG" then %>
                        ResgatarDesigServicoGravado(<%=dblSerId %>)
                        <% End if%>
	}
        }
    }

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
            } else {
                divOrigemSol.style.display = 'none';
            }
        }
    }
    //LP


    //-->
</script>
<form method="post" name="Form1">

    <!-- Good 22/01/2025 inicio -->

    <input type="hidden" name="hdntxtFacilidade" value="">
    <input type="hidden" name="hdnAcaoMain" value='<%=request.form("cboAcao")%>'>
    <input type="hidden" name="hdnobjChave">
    <input type="hidden" name="hdncboTecnologia" value="">
    <input type="hidden" name="hdnTecnologia1">
    <input type="hidden" name="hdnTecnologia2">
    <input type="hidden" name="hdntxtFacilidade1" value="">
    <input type="hidden" name="hdntxtFacilidade2" value="">
    <input type="hidden" name="hdnAcao" value='<%=request.form("cboAcao")%>'>
    <!-- Good 22/01/2025 fim -->

<input type=hidden name=hdnUsuario>
<input type=hidden name=hdnOrigemEst>
<input type=hidden name=hdnCNLEstUsu>
<input type=hidden name=hdnComplEstUsu>
<input type=hidden name=hdnCboServico>
<input type=hidden name=hdnNomeCbo>
<input type=hidden name=hdnNomeLocal>
<input type=hidden name=hdnUserGICL value="<%=strUserNameGICLAtual%>">
<input type=hidden name=hdnDesigServ>
<input type=hidden name=hdnOrderEntry value="<%=strOrderEntry%>">
    <%
if strOrderEntrySis = "Aprov" and objRSSolic("OriSol_ID") = 6 then strOrderEntrySis = "SGA PLUS" end if
if strOrderEntrySis = "Aprov" and objRSSolic("OriSol_ID") = 7 then strOrderEntrySis = "SGA VOZ VIP'S" end if
if objRSSolic("OriSol_ID") = 9 then strOrderEntrySis = "ASMS" end if
if strOrderEntrySis = "Aprov" and objRSSolic("OriSol_ID") = 10 then strOrderEntrySis = "CFD" end if
    %>
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
    <input type="hidden" name="hdnOriSol_ID" value="<%=objRSSolic("OriSol_ID")%>">
    <input type="hidden" name="hdnAprovisi_ID" value="<%=dblAprovisiId%>">
<input type=hidden name=hdnIdAcessoLogico value="<%=dblIdLogico%>">
<input type=hidden name=hdnId_Acesso value="<%=id_acesso%>">
<input type=hidden name=hdnSolId value="<%=dblSolId%>">
<input type=hidden name=hdnDtSolicitacao value="<%=strDtPedido%>">
<input type=hidden name=hdnPadraoDesignacao >
<input type=hidden name=hdnCtfcId value="<%=dblCtfcId%>" >
<input type=hidden name=hdnSubAcao>
<input type=hidden name=hdnXmlReturn value="<%=Request.Form("hdnXmlReturn")%>">
<input type=hidden name=hdnAcfId >
<input type=hidden name=hdnStrAcfId value="<%=strAcf_id%>" >

<input type=hidden name=hdnProjEsp>


<input type=hidden name=hdnCLINOME>
<input type=hidden name=hdnCLINOMEFANTASIA>
<input type=hidden name=hdnCliCC>
<input type=hidden name=hdnCLISUBCC>
<input type=hidden name=hdnSEGMENTO>
<input type=hidden name=hdnPORTE>
<input type=hidden name=hdnSERDESC>
<input type=hidden name=hdnVELDESC>
<input type=hidden name=hdnOBSSEV>
<input type=hidden name=hdnPABXVIRTUAL value="<%=strPabx_Virtual%>">
<input type=hidden name=hdnVarRSVerFis value="<%=varRSVerFis%>">

    <%
'<!-- Projeto ABW - Incio -->'
set rsProjABW = db.execute (" select top 1 aprovisi_id, indicador_migracao_pabx_virtual , indicador_tecnologia_ip " &_
							" from cla_aprovisionador " &_
							" where aprovisi_id = " & dblAprovisiId & " order by 1 desc")
			
if Not rsProjABW.Eof and Not rsProjABW.Bof then

	strMigraPabx_Virtual = Trim(rsProjABW("indicador_migracao_pabx_virtual"))

	IF ISNULL(strMigraPabx_Virtual) THEN
		strMigraPabx_Virtual = ""
	END IF
	
	IF strMigraPabx_Virtual = "NULL" THEN
		strMigraPabx_Virtual = ""
	END IF

	IF strMigraPabx_Virtual = "N" THEN
		strMigraPabx_Virtual = "Não"
	END IF

	IF strMigraPabx_Virtual = "S" THEN
		strMigraPabx_Virtual = "Sim"
	END IF

	strTecnologia_IP = Trim(rsProjABW("indicador_tecnologia_ip"))

	IF ISNULL(strTecnologia_IP) THEN
		strTecnologia_IP = ""
	END IF
	
	IF strTecnologia_IP = "NULL" THEN
		strTecnologia_IP = ""
	END IF

	IF strTecnologia_IP = "N" THEN
		strTecnologia_IP = "Não"
	END IF

	IF strTecnologia_IP = "S" THEN
		strTecnologia_IP = "Sim"
	END IF

end if
'<!-- Projeto ABW - Fim -->'
    %>


    <%if strOrigem="APG" then%>
	<input type=hidden name=idTarefaApg value="<%=objRSSolic("id_tarefa_Apg")%>">
    <% End if%>


    <div id="divWait" style="background-color: #dcdcdc; width: 300px; height: 100px; float: left; margin: 800px 0 0 280px; position: absolute; border: 1px solid #0f1f5f; padding: 40px 0 0 20px; display: none;">
        <p align="center" style="font-size: 12px; font-family: Arial, Helvetica; font-weight: bold; color: #003366;">Aguarde. Estamos consultando o CSL ...</p>
    </div>

    <div id="divSistemaWait" style="background-color: #dcdcdc; width: 300px; height: 100px; float: left; margin: 800px 0 0 280px; position: absolute; border: 1px solid #0f1f5f; padding: 40px 0 0 20px; display: none;">
        <p align="center" style="font-size: 12px; font-family: Arial, Helvetica; font-weight: bold; color: #003366;">Aguarde. Sistema Processando ...</p>
    </div>

    <%
'JCARTUS@ Dez/2011 - Abrir para o usuário digitar IDLógico quando o informado na interface estiver 
'                    cancelado(ou ainda em processo) ou desativado(ou ainda em processo)
if (strIDLogico <> "" and strAcao = "ALT")  then

	Vetor_Campos(1)="adWchar,10,adParamInput, " & trim(strIDLogico)
	Vetor_Campos(2)="adWChar,20,adParamInput, " & trim(strAcao)
	Vetor_Campos(3)="adInteger,10,adParamInput, " & strAprovisiId
	Vetor_Campos(4)="adInteger,4,adParamOutput,0"
	Vetor_Campos(5)="adWChar,100,adParamOutput,0"

	Call APENDA_PARAM("CLA_sp_check_IDLogico2",5,Vetor_Campos)

	ObjCmd.Execute'pega dbaction
	
	DBErro = ObjCmd.Parameters("RET").value
	DBErroDesc = ObjCmd.Parameters("RET1").value
'response.write "<script>alert('DBErro=" & DBErro & "')</script>"
'response.write "<script>alert('DBErroDesc=" & DBErroDesc & "')</script>"
 desativar_gravar = false
	if DBErro > 0 then
		desativar_gravar = true
		response.write "<script>alert('"& trim(DBErroDesc) &"')</script>"

		'<!--CH-41083NDL - inicio-->
		bbloqueiaSev = "true"
		'<!--CH-41083NDL - Fim-->

	end if
  	
	if DBErro = 9 then
	  strIDLogico = ""
	  dblIdLogico = ""
	end if

end if
    %>

<tr><td>
            <%if ( ( strIDLogico = "" or isnull(strIDLogico)) and strAcao = "ALT" )  then%>
	<table cellspacing="1" cellpadding="0" border=0 width="790">
                <th>
                    <th nowrap colspan="4" style="background-color: #A80000">&nbsp;•&nbsp;Informações de Aprovisionamento - Dados do Item de OE de Origem para facilitar busca do Acesso Lógico</th>
                </th>
            </table>

	<table cellspacing="1" cellpadding="0" border=0 width="790">
                <tr class="clsSilver">
                    <td width="465">Razão Social: <%=strPov_Razao%></td>
                    <td colspan="2">Conta Corrente: <%=strPov_CC%></td>
                    <td width="96">SubConta: <%=strPov_SubCC%></td>
                </tr>
            </table>
	<table cellspacing="1" cellpadding="0" border=0 width="790">
                <tr class="clsSilver">
                    <td width="298">Designação do Serviço: <%=strPov_DesigServ%></td>
                    <td width="426" colspan="2">Velocidade do Serviço: <%=strPov_VelServ%></td>
                    <td width="249">Nº Contrato de Serviço: <%=strPov_ContratoServ%></td>
                </tr>
            </table>
	<table cellspacing="1" cellpadding="0" border=0 width="790">
                <tr class="clsSilver">
                    <td width="228">Contato Cliente: <%=strPov_ContatoCli%></td>
                    <td width="161">Telefone: <%=strPov_Tel%></td>
                    <td width="169">CNPJ: <%=strPov_CNPJ%></td>
                    <td width="237">Prop. End.: <%=strPov_PropEnd%></td>
                </tr>
	<table cellspacing="1" cellpadding="0" border=0 width="790">
                    <tr>
                        <th nowrap colspan="4" style="background-color: #A80000">&nbsp;•&nbsp;POVOAMENTO SOB DEMANDA: Associação de Serviços enviados pelo Aprovisionador com Acesso Lógico cadastrado no CLA</th>
                    </tr>
                </table>
	<table cellspacing="1" cellpadding="0" border=0 width="790">
                    <tr class="clsSilver">
                        <td width="517">Informar o Acesso Lógico que será associado ao serviço e que sofrerá o processo de 
			<%
			select case trim(ucase(strAcao))
			  Case "ATV"
			    response.write "ativação"
			  Case "DES"
			    response.write "desativação"
			  Case "ALT"
			    response.write "alteração"
			  Case "CAN"
			    response.write "cancelamento"
			  Case "REC"
			    response.write "recadastro"
			end select
            %>.</td>
                        <td colspan="3">
                            <input type="hidden" name="hdnAcaoAprov" value="<%=strAcao%>">
                            <input type="hidden" name="hdnAprovisiId" value="<%=dblAprovisiId%>">
                            <input type="hidden" name="hdnOriSolID" value="<%=dblOriSolID%>">
                            <input type="text" class="text" name="txtNroLogico" value="<%=dblIdLogico%>" maxlength="10" size="11" onkeyup="ValidarTipo(this,0)">&nbsp;&nbsp; &nbsp;
				<input type="button" class="button" name="associarlogico" value="Associar Lógico" title="Associar Acesso Lógico" style="cursor:hand" onClick="AssociarLogico()" tabindex=0 accesskey="A" onmouseover="showtip(this,event,'Associar Lógico (Alt+A)');">
                        </td>
                    </tr>
                </table>

                <%
	'<!--CH-41083NDL - inicio-->
	bbloqueiaSev = "True"
	'<!--CH-41083NDL - Fim-->
                %>

                <%end if%>

                <%
'Correção 15 min do schedule de cancelamento automático.
if (strIDLogico <> "" and (strAcao = "ALT" or strAcao = "ATV"))  then

	strIDLogico677 = "677" & mid(strIDLogico,4,7)
	strIDLogico678 = "678" & mid(strIDLogico,4,7)

	strSQL = "select top 1 Acl_IDAcessoLogico from cla_aprovisionador tab1 WITH (NOLOCK) where acao in ('ALT','ATV') and Acl_IDAcessoLogico in (select Acl_IDAcessoLogico from cla_aprovisionador as tab2	WITH (NOLOCK) where tab1.ID_Tarefa = tab2.ID_Tarefa_Can and acao='CAN' and ID_Tarefa_Can is not null and not exists (select top 1 * from cla_aprovisionador tab3 WITH (NOLOCK) where tab3.ID_Tarefa =  CAST(tab2.ID_Tarefa_Can AS INT) and aprovisi_dtCancAuto is not null)) and Acl_IDAcessoLogico in (" & strIDLogico677 & "," & strIDLogico678 & ")"

	Set objRSCan = db.execute(strSQL)
	
	if not objRSCan.Eof then
	
	strbloqcan = true
                %>
	<table cellspacing="1" cellpadding="0" border=0 width="790">
                    <tr>
                        <th nowrap colspan="4" style="background-color: #A80000">&nbsp;•&nbsp;REGISTRO BLOQUEADO (15MIN): Aguardando cancelamento automático do registro anterior 
			&nbsp;•&nbsp;&nbsp;<input type=button class="button" name="btnrefresh" value="Atualizar página" onclick="JavaScript:location.reload(true);"></th>
                    </tr>
                </table>
                <%
	end if
end if
                %>


<table cellspacing="1" cellpadding="0" border=0 width="790">
                    <tr>
		<th width=25%>&nbsp;•&nbsp;Solicitação de Acesso</th>
		<th width=25%>&nbsp;Nº&nbsp;:&nbsp;<span id=spnSolId><%=dblSolId%></Span></th>
                        <%if Trim(Request.Form("hdnAcao")) = "Alteracao" then%>
			<th width=25%>&nbsp;Tipo&nbsp;:&nbsp;Alteração</th>
			<th width=25%>&nbsp;Acesso Lógico&nbsp;:&nbsp;<%=dblIdLogico%></th>
                        <%Elseif Trim(Request.Form("hdnOrigem")) = "Recadastro" then%>
			<th width=25%>&nbsp;Tipo&nbsp;:&nbsp;Recadastro</th>
                        <%else%>
			<th width=25%>&nbsp;Tipo&nbsp;:&nbsp;Ativação</th>
                        <%End if%>
		<th width=25%>&nbsp;Data&nbsp;:&nbsp;<%=strDtPedido%></th>
                    </tr>
                </table>

<table border=0 cellspacing="1" cellpadding="0" width="790">
                    <tr>
		<th colspan=4>&nbsp;•&nbsp;Informações do Cliente </th>

		<th width="27">
			
		</th>
                    </tr>
                    <tr class="clsSilver">
                        <td width="170">&nbsp;&nbsp;&nbsp;&nbsp;Sev para procura</td>
                        <td>
                            <input type="hidden" name="hdnOriSol_ID_N" value="<%=objRSSolic("OriSol_ID")%>">

                            <input type="text" class="text" id="txtNroSev" name="txtNroSev" readonly="readOnly" value="<%=dblNroSev%>" maxlength="8" size="10">&nbsp;

			
			<!--Alterado por Fabio Pinho em 28/04/2016 - ver 1.0 - Fim-->

                            <!--RAIO-x: BOTÃO DE REANÁLISE DE SEV HABILITADO TEMPORÁRIAMENTE -->
                            <!-- <input type="button" class="button" name="procurarsev" value=" Reanálise de Sev  " onClick="alert('O processo de reanálise de SEV sofreu alteração: O pedido deve ser realizado diretamente no Star21 através da inserção das pendências 23.X, sem necessidade de atuação no SSA e CLA.');" tabindex=-1 accesskey="P" > -->

                        </td>
                        <td class="clsSilver" colspan="3">&nbsp;&nbsp;&nbsp;Origem Solicitação 
			<select name="cboOrigemSol" <%=bbloqueia%>>
                <%if Trim(Request.Form("hdnOrigem")) = "Recadastro" then%>
                <option value="8">RECADASTRO</option>
                <%else%>
                <option value="<%=objRSSolic("OriSol_ID")%>"><%=objRSSolic("OriSol_Descricao")%></option>
                <%end if%>
            </select>
                        </td>
                    </tr>
                    <!-- ''@@ LPEREZ - 24/10/2005 -->

                    <!--	<tr>
		</td>
		<td colspan=3 class="clsSilver">Origem Solicitação <span align=right>
			<select name="cboOrigemSol" <%=bbloqueia%>>
			    <%if Trim(Request.Form("hdnOrigem")) = "Recadastro" then%>
				<option value="8">RECADASTRO</option>
				<%else%>
				<option value="<%=objRSSolic("OriSol_ID")%>"><%=objRSSolic("OriSol_Descricao")%></option>
				<%end if%>
		  	</select>
		</td>
	</tr> -->
                    <tr class="clsSilver">
                        <td width="170"><font class="clsObrig">:: </font>Razão Social</td>
                        <td colspan="4">
                            <input type="text" class="text" name="txtRazaoSocial" maxlength="80" size="80" value="<%=strRazaoSocial%>"
                                <%=bbloqueia%> onblur="ResgatarGLA()">

			
                        </td>
                    </tr>
                    <!--
	<tr class="clsSilver">
		<td width="170"><span id=spnLabelCliente></span></td>
		<td colspan="4"><span id=spnCliente></span></td>
	</tr> -->
                    <tr class="clsSilver">
                        <td width="170"><font class="clsObrig">:: </font>Nome Fantasia</td>
                        <td colspan="4">
                            <input type="text" class="text" name="txtNomeFantasia" maxlength="40"
                                <%=bbloqueia%>
                                size="40" value="<%=strNomeFantasia%>">
                        </td>
                        <!--	<td colspan=3 class="clsSilver">&nbsp;</td> -->
                        <!-- ''@@ -->
                    </tr>
                    <tr class="clsSilver" nowrap>
                        <td width="170" nowrap><font class="clsObrig">:: </font>Conta Corrente
                        </td>
		<td width=279 >
			<input type=text class="text" name=txtContaSev size=11 maxlength=11
                                <%=bbloqueia%>
                                onkeyup="ValidarTipo(this,0)" onblur="CompletarCampo(this)" tipo="N" value="<%=strContaSev%>">
                        </td>
                        <!--	<td align=right width=89 > -->
		<td><p align="left"><font class="clsObrig">:: </font>Sub Conta&nbsp;</td>
                        <td width="189" colspan="3">
			<input type=text name=txtSubContaSev class="text" size=4 maxlength=4
                                <%=bbloqueia%>
                                onkeyup="ValidarTipo(this,0)" onblur="CompletarCampo(this)" tipo="N" value="<%=strSubContaSev%>">
                            <!--	</td><td>&nbsp;</td> -->
                        </td>
                    </tr>
                    <tr class="clsSilver" nowrap>
                        <td width="170" nowrap><font class="clsObrig">:: </font>Segmento</td>
		<td width=279>
			<input type=text class="text" name=txtSegmento size=22 maxlength=22
                                <%=bbloqueia%>
                                value="<%=SEGMENTO%>">
                        </td>
                        <!--	<td align=right width=89 > -->
		<td ><p align="left"><font class="clsObrig">:: </font>Porte&nbsp; </td>
                        <td width="189" colspan="3">
			<input type=text name=txtPorte class="text" size=11 maxlength=22
                                <%=bbloqueia%>
                                value="<%=PORTE%>">
                        </td>
                    </tr>

                </table>
<table border=0 cellspacing="1" cellpadding="0" width="790">
                    <tr>
		<th colspan=4 >
			&nbsp;•&nbsp;Informações do Serviço&nbsp;
                        </th>
                    </tr>
                    <% if objRSSolic("OriSol_ID") <> 10 then %>
                    <tr class="clsSilver">
                        <td width="170px"><font class="clsObrig">:: </font>Order Entry</td>
		<td colSpan=3>
		<table border=0 border=0 cellspacing="0" cellpadding="0">
			<tr align=center class=clsSilver>
                                    <td>Sistema</td>
				<td></td>
                                    <td>Ano</td>
                                    <td></td>
                                    <td>Nro</td>
                                    <td></td>
                                    <td>Item</td>
                                    <td></td>
				<td></td>
                                </tr>
			<tr class=clsSilver>
                                    <td>
                                        <select name="cboSistemaOrderEntry" <%=bbloqueia%> onchange="SistemaOrderEntry(this);hdnOrderEntrySis.value=this.value;">
                                            <%if objRSSolic("OriSol_ID") = 7 then%>
                                            <option value="SGA VOZ VIP'S" selected>SGA VOZ VIP'S</option>
                                            <%elseif objRSSolic("OriSol_ID") = 6 then%>
                                            <option value="SGA PLUS" selected>SGA PLUS</option>
                                            <%else%>
                                            <option value="ASMS" selected>ASMS</option>
                                            <%end if%>
                                        </select>
                                    </td>
                                    <td>-</td>
				<td><input type="text" class="text" onblur="CompletarCampo(this);hdnOrderEntryAno.value=this.value;" onkeyup="ValidarTipo(this,0)" maxlength=4 size=4 <%=bbloqueia%> name=txtOrderEntry TIPO="N" value="<%=strOrderEntryAno%>" ></td>
                                    <td>-</td>
				<td><input type="text" class="text" onblur="CompletarCampo(this);hdnOrderEntryNro.value=this.value;" onkeyup="ValidarTipo(this,0)" maxlength=7 size=7 <%=bbloqueia%> name=txtOrderEntry TIPO="N" value="<%=strOrderEntryNro%>" ></td>
                                    <td>-</td>
				<td><input type="text" class="text" onblur="CompletarCampo(this);hdnOrderEntryItem.value=this.value;" onkeyup="ValidarTipo(this,0)" maxlength=3 size=3 <%=bbloqueia%> name=txtOrderEntry TIPO="N" value="<%=strOrderEntryItem%>" ></td>
				<td></td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <% else %>
                    <tr class="clsSilver">
                        <td width="170px"><font class="clsObrig">:: </font>IA</td>
		<td colSpan=3>
		<table border=0 border=0 cellspacing="0" cellpadding="0">
			<tr align=center class=clsSilver>
                                    <td>Sistema</td>
                                    <td></td>
                                    <td></td>
				
                                </tr>
			<tr class=clsSilver>
                                    <td>
                                        <select name="cboSistemaOrderEntry" <%=bbloqueia%> onchange="SistemaOrderEntry(this);hdnOrderEntrySis.value=this.value;">
                                            <%if objRSSolic("OriSol_ID") = 10 then%>
                                            <option value="CFD" selected>CFD</option>

                                            <%end if%>
                                        </select>
                                    </td>
                                    <td>-</td>
				<td><input type="text" class="text" maxlength=50 size=50 <%=bbloqueia%> name=txtIA TIPO="N" value="<%=strIA%>" ></td>
				<td></td>
                                </tr>
                            </table>
                        </td>
                    </tr>


                    <% end if  %>

                    <tr class="clsSilver">
                        <td width="170px"><font class="clsObrig">:: </font>Serviço</td>
		<td width="100px" colspan=3>
                            <%
			set objRS = db.execute("CLA_sp_sel_servico null,null,null,1")
                            %>
                            <select name="cboServicoPedido" <%=bbloqueia%> onchange="document.Form1.rdoAntAcesso[1].checked = true;ResgatarServico(this)" tyle="width:100px">
                                <option></option>
                                <%
				While Not objRS.eof
					strItemSel = ""
					if Trim(dblSerId) = Trim(objRS("Ser_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value='" & objRS("Ser_ID") & "'" & strItemSel & ">" & objRS("Ser_Desc") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
                                %>
                            </select>

                        </td>

                    </tr>



                    <tr class="clsSilver">
		<td width="150px" align=left><font class="clsObrig">:: </font>Velocidade&nbsp;</td>
		<td colspan=3 ><span id=spnVelServico>
                            <select name="cboVelServico" <%=bbloqueia%> onchange="SelVelAcesso(this)" style="width: 110px">
                                <option></option>
                                <%
						if Trim(dblSerId) <> "" then
						
							set objRS = db.execute("CLA_sp_sel_AssocServVeloc null," & dblSerId)
							While Not objRS.eof
								strItemSel = ""
								
								if Trim(dblVelServico) = Trim(objRS("Vel_ID")) then strItemSel = " Selected " End if
								'response.write "<script>alert('strItemSel=" & Trim(objRS("Vel_Desc")) & " >>" & strItemSel & "')</script>"
								Response.Write "<Option value='" & objRS("Vel_ID") & "'" & strItemSel & ">" & Trim(objRS("Vel_Desc")) & "</Option>"
								objRS.MoveNext
							Wend
							strItemSel = ""
						End if

                                %>
                        </span>
                        </td>
                    </tr>


                    <tr class="clsSilver">
                        <td width="170px">&nbsp;&nbsp;&nbsp;Designação do Serviço</td>
                        <td>
                            <input type="text" class="text" name="txtdesignacaoServico"
                                <%=bbloqueia%>
                                value="<%=strDesignacaoServico%>" maxlength="22" size="30"><br>
                        </td>
                        <%
		'EF: VEL_VOZ - esse novo campo não será obrigatório. Caso venha preenchido deverá ser exibido
		'abaixo da velocidade do serviço já existente.
		'A sua função é apoiar a decisão do PVT na escolha da velocidade do acesso físico.
		'GPON - INICIO
		IF strVel_Voz <> "" and strOriSol = 6 THEN
                        %>
		<td width="150px" align=right><font class="clsObrig">:: </font>Vel Voz&nbsp;</td>
		<td>
                            <input type="text" class="text" name="txtVelVoz"
                                <%=bbloqueia%>
                                value="<%=strVel_Voz%>" maxlength="30" size="30"><br>
                        </td>
                        <%
		ELSE
                        %>
		<td width="150px" align=right colspan=4><font class="clsObrig"> </font>&nbsp;</td>
                        <!--	<td>
		</td> -->
                        <%
		END IF
		'GPON - FIM%>
                    </tr>

                    <tr class="clsSilver">
                        <td width="170px"><font class="clsObrig">:: </font>Nº Contrato Serviço</td>
		<td colspan=3>
                            <table rules="groups" cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="70%">
				<tr><td nowrap width=200px >
					<input type=radio name=rdoNroContrato value=1
                                            <%=bbloqueia%>
					onClick="spnDescNroContr.innerHTML= 'Ex.: VEM-11 XXX000012003'" checked <%if strTipoContratoServico = "1" then Response.Write " checked " End if%>>Contrato de Serviço</td><td></td></tr>
                                <tr>
					<td nowrap><input type=radio name=rdoNroContrato onClick="spnDescNroContr.innerHTML= 'Padrão: A22'"
                                            <%=bbloqueia%>
					value=2 <%if strTipoContratoServico = "2" then Response.Write " checked " End if%>>Contrato de Referência</td>
                                    <td nowrap>
                                        <input type="text" class="text" name="txtNroContrServico" value="<%=strNroContrServico %>"
                                            <%=bbloqueia%>
                                            maxlength="22" size="30"><br>
						<span id=spnDescNroContr>Ex.: VEM-11 XXX00012003</span>
                                    </td>
                                </tr>
				<tr><td nowrap><input type=radio name=rdoNroContrato
                                            <%=bbloqueia%>
				onClick="spnDescNroContr.innerHTML= 'Padrão: A22'" value=3 <%if strTipoContratoServico = "3" then Response.Write " checked " End if%> >Carta de Compromisso</td><td></td></tr>
                            </table>
                        </td>
                    </tr>



                    <% 

	if Trim(Request.Form("hdnAcao")) = "Alteracao" then %>

		<tr><th colspan=4>&nbsp;•&nbsp;Informações de Alteração</th></tr>

                    <tr class="clsSilver">
			<td width=170px>&nbsp;&nbsp;&nbsp;&nbsp;Tipo de Alteração</td>
                        <td>
                            <select name="cboTipoAlteracao" style="width: 150px">
				<option ></option>	&nbsp;
				<option value="Endereco"
                    <% if trim(strTipoAlteracao) = "Endereco" Then %>
                    selected>Endereço </option>
                                <%Else %>
					> Endereço </option>
				<%End If %>

                                <option value="Velocidade"
                                    <% if trim(strTipoAlteracao) = "Velocidade" Then %>
                                    selected>Velocidade </option>
                                <%Else %>
					> Velocidade </option>
				<%End If %>

                                <option value="Velocidade+Facilidade"
                                    <% if trim(strTipoAlteracao) = "Velocidade+Facilidade" Then %>
                                    selected>Velocidade e Facilidade</option>
                                <%Else %>
					> Velocidade e Facilidade</option>
				<%End If %>
	
                            </select>
                        </td>
                        <td>&nbsp;&nbsp;&nbsp;Programação &nbsp;da &nbsp;&nbsp;&nbsp; Alteração&nbsp;</td>
                        <td>
                            <select name="cboProgAlteracao" style="width: 150px">
				<option ></option>	&nbsp;&nbsp;
				<option value="Programada"
                    <% if trim(strProgramacaoAlteracao) = "Programada" Then %>
                    selected>Programada  </option>
                                <%Else %>
					 > Programada  </option>
				<%End If %>


                                <option value="Nao Programada"
	
                                    <% if trim(strProgramacaoAlteracao) = "Nao Programada" Then %>
                                    selected>Não Programada  </option>
                                <%Else %>
					> Não Programada  </option>
				<%End If %>
	
                            </select>
                        </td>
                    </tr>
                    <%end if %>
</form>

</table>




<table border=0 cellspacing="1" cellpadding="0" width="790">
    <tr>
		<th colspan=4 >
			&nbsp;•&nbsp;Análise do estudo de viabilidade (SEV   SSA) &nbsp;
        </th>
    </tr>
    <tr>
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

					<table cellspacing=1 cellpadding=0 border=0 width=790> 
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
						<table cellspacing=1 cellpadding=0 border=0 width=790> 
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

								  		<table cellpadding=0 cellspacing=1 width=790 > 
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

										'------------------------------------Facilidade Array																				
										dim dataFacilidade
										dataFacilidade=""

									 While Not objRSCli.eof
		
                        %>
											<tr class=clsSilver > 
													<td width=15% nowrap >&nbsp; <%=objRSCli("FAC_DES")%>  </td>
													<td width=15% nowrap >&nbsp; <%=objRSCli("TEC_DES")%>  </td>	
                            <td>&nbsp;  <%=objRSCli("FOR_DES")%> </td>
                            <td>&nbsp;  <%=objRSCli("ESTACAO")%> </td>
                            <td nowrap>&nbsp;  <%=objRSCli("SEV_USRNAMESOLUC")%>  </td>
                            <td nowrap>&nbsp; <%=objRSCli("data_vencimento")%> </td>
													<td width=20% >&nbsp;  <%=objRSCli("SOL_OBS")%>  </td> 
                        </tr>
                        <%																		
									dataFacilidade = dataFacilidade & Trim(objRSCli("FAC_DES")) & ";"
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
        </td>
    </tr>


</table>
<table ID= tbRadio cellspacing=1 cellpadding=0 width=790 border=0>
<Form name=Form2 method=Post>
        <!--JCARTUS@ - Acessos Físicos sem CNPJ/IE/IM vinculado na base do CLA (CH-69613HSK)-->

        <!-- Good 22/01/2025 inicio -->
        <input type="hidden" name="hdnRdoAcesso" value="">
        <input type="hidden" name="hdncboTecnologia" value="">
        <input type="hidden" name="hdnfac" value="">
        <input type="hidden" name="hdnPropIdFisBkp" value="">
        <!-- Good 22/01/2025 fim -->

<input type=hidden name=hdnCNPJ value="<%=dblCNPJ%>">
<input type=hidden name=hdnIE value="<%=strIE%>">
<input type=hidden name=hdnIM value="<%=strIM%>">
<input type=hidden name=hdnUsuario>
<input type=hidden name=hdnOrigemEst>
<input type=hidden name=hdnCNLEstUsu>
<input type=hidden name=hdnComplEstUsu>
<input type=hidden name=hdnIntIndice>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnSubAcao>
<input type=hidden name=hdnProvedor>
        <!-- Good  inicio -->

        <!--<input type="hidden" name="hdnCboProvedor" value=cboProvedor.value>-->
        <input type="hidden" name="hdnCboProvedor">

        <!-- Good  fim -->

<input type=hidden name=hdnstrAcessoTipoRede value="<%=strAcessoTipoRede%>">
        <!-- Good  inicio -->

        <!--<input type="hidden" name="hdnSegmento"    value="txtSegmento.value">-->
        <input type="hidden" name="hdnSegmento">

        <!-- Good  fim -->

<input type=hidden name=hdnConta    value="<%=strContaSev%>">
        <!-- Good  inicio -->

        <!--<input type="hidden" name="hdnPorte"      value=txtPorte.value>-->
        <input type="hidden" name="hdnPorte">

        <!--<input type="hidden" name="hdnCNLSiglaCentroCliDest"      value=txtCNLSiglaCentroCliDest.value>
<input type="hidden" name="hdnComplSiglaCentroCliDest"    value=txtComplSiglaCentroCliDest.value>-->
        <input type="hidden" name="hdnCNLSiglaCentroCliDest">
        <input type="hidden" name="hdnComplSiglaCentroCliDest">

        <!-- Good  fim -->

<input type=hidden name=hdnTipoAcaoAPG>
<input type=hidden name=hdnNroSev>
<input type=hidden name=hdnOrigemSol>
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

<input type=hidden name=hdnTecnologia 		id=hdnTecnologia 		value="">
<input type=hidden name=hdnFacilidade 		id=hdnFacilidade 		value="">
<input type=hidden name=hdnNomeFacilidade 	id=hdnNomeFacilidade	value="">


<input type=hidden name=hdnVelAcessoFisSel>
<input type=hidden name=hdnAecIdFis>
<input type=hidden name=hdnEstacaoOrigem>
<input type=hidden name=hdnEstacaoDestino>
<input type=hidden name=hdnObrigaGla value="<%=strObrigaGla%>">
<input type=hidden name=hdnAcfId >
<input type=hidden name=strOrigemAPG value="<%=strOrigem%>">
<input type=hidden name=hdnUsuID value="<%=dblUsuID %>">
        <input type="hidden" name="hdnOriSol_ID" value="<%=objRSSolic("OriSol_ID")%>">
        <input type="hidden" name="hdnAprovisi_ID" value="<%=dblAprovisiId%>">
<input type=hidden name=hdnId_Acesso value="<%=id_acesso%>">
<input type=hidden name=hdnStrAcfId value="<%=strAcf_id%>" >
<input type=hidden name=hdnNroSev2 value="<%=dblNroSev%>" >
        <!-- 
	Verificar o Procedimento de cosnulta do acesso compartilhado
	ProcessoSolic.asp 11/05/2009
-->
<input type=hidden name=cboUFEndAPG value="<%=strUFEnd%>">
  <input type=hidden name=txtEndCidAPG value="<%=strEndCid%>">
  <input type=hidden name=txtEndCidDescAPG value="<%=strEndCidDesc%>">
  <input type=hidden name=cboLogrEndAPG value="<%=strLogrEnd%>">
  <input type=hidden name=txtEndAPG value="<%=strEnd%>">
  <input type=hidden name=txtNroEndAPG value="<%=strNroEnd%>">
  <input type=hidden name=txtComplEndAPG value="<%=strComplEnd%>">
  <input type=hidden name=txtBairroEndAPG value="<%=strBairroEnd%>">
  <input type=hidden name=txtCepEndAPG value="<%=strCepEnd%>">
        <!--
  <input type=hidden name=cboUFEnd value="<%=strUFEnd%>">
  <input type=hidden name=txtEndCid value="<%=strEndCid%>">
  <input type=hidden name=txtEndCidDesc value="<%=strEndCidDesc%>">
  <input type=hidden name=cboLogrEnd value="<%=strLogrEnd%>">
  <input type=hidden name=txtEnd value="<%=strEnd%>">
  <input type=hidden name=txtNroEnd value="<%=strNroEnd%>">
  <input type=hidden name=txtComplEnd value="<%=strComplEnd%>">
  <input type=hidden name=txtBairroEnd value="<%=strBairroEnd%>">
  <input type=hidden name=txtCepEnd value="<%=strCepEnd%>">

  -->

        <%if Trim(Request.Form("hdnAcao")) = "Alteracao" then%>
<input type=hidden name=hdnTipoProcesso value=3>
        <%Else%>
<input type=hidden name=hdnTipoProcesso value=1>
        <%End if%>


        <%'JKNUP: Adicionado%>
<input type=hidden name=hdnTipoTec>
<input type=hidden name=hdnCNLCliente>

        <!--	
		Verificar Procedimento de recuperação de físico utilizado para Alteração. 	
	-->

        <%'if ( (Request.Form("hdnAcao") = "Alteracao") ) then %>
        <!--			
			<tr><th colspan=4>&nbsp;•&nbsp;Ponto de Instalação do Processo Ativo Anterior</th></tr>
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
					<iframe id=IFrmAcessoFisAnterior
							name=IFrmAcessoFisAnterior
							align=left
							src="AcessosFisicosAnterior.asp"
							frameBorder=0
							width="100%"
							BORDER=0
							height=40>
					</iframe>
				</td>
			</tr>
			
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
					<iframe id=IFrmAcessoFisApg
							name=IFrmAcessoFisApg
							align=left
							src="AcessosFisicosApg.asp"
							frameBorder=0
							width="100%"
							BORDER=0
							height=40>
					</iframe>
				</td>
			</tr>
		-->

        <%'else%>



			<tr><th colspan=4>&nbsp;•&nbsp;Acessos Físicos Utilizados</th></tr>
			<tr><td colspan=4>
					<table border=0 width=790 cellspacing=1 cellpadding=0>
        <tr>
							<th  width=15>&nbsp;</th>
							<th  width=35>&nbsp;Editar</th>
							<th  width=50>&nbsp;Prop Fis</th>
							<th  width=185>&nbsp;Provedor</th>
							<th  width=200>&nbsp;Velocidade</th>
							<th	 width=307>&nbsp;Endereço</th>
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
        <% 'End if%>

        <tr>
		<th colSpan=4>&nbsp;•&nbsp;Informações do Acesso&nbsp;</th>
        </tr>
        <tr class="clsSilver">
			<td width=170px ><font class="clsObrig">:: </font>Prop do Acesso Físico</td>
            <td nowrap>
                <%

					RSsolProv= ""
					Set objRSsolProv = db.execute("select ltrim(rtrim(for1.pro_id)) as pro_id from ssa_1.dbo.solucao sol inner join ssa.dbo.fornecedor for1 on sol.for_seq = for1.for_seq where sol.sev_seq = " & dblNroSev  )
			
					While Not objRSsolProv.Eof	
						RSsolProv = objRSsolProv("pro_id")  
					   objRSsolProv.movenext
					Wend
					
					'response.write "<script>alert('"&RSsolProv&"')</script>"
                %>
                <!-- Good  inicio -->
                <!--		<input type=radio name=rdoPropAcessoFisico value="TER" <%if RSsolProv = "11" then%> disabled <%end if%>   Index=0 onclick="EsconderTecnologia(0);ResgatarTecVel();">Terceiro&nbsp;&nbsp;&nbsp;
				<input type=radio name=rdoPropAcessoFisico value="EBT" <%if RSsolProv <> "11" then%> disabled <%end if%> Index=1 onclick="EsconderTecnologia(0);ResgatarTecVel();">CLARO Brasil&nbsp;&nbsp;&nbsp;
	-->
				<input type=radio name=rdoPropAcessoFisico value="TER"  Index=0 onclick="EsconderTecnologia(0);ResgatarTecVel();" >Terceiro&nbsp;&nbsp;&nbsp;
				<input type=radio name=rdoPropAcessoFisico value="EBT"  Index=1 onclick="EsconderTecnologia(0);ResgatarTecVel();">CLARO Brasil&nbsp;&nbsp;&nbsp;	

	<!-- Good  Fim -->
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
	     dataString =  dataString & dataArray(i) & ";"    		
    End If
Next

                %>
	<td nowrap colspan=2> <span ID =spnFacilidadeTecnologia  name=spnFacilidadeTecnologia >
                <select name="txtFacilidade" onblur='VerificaFacilidade("<%=Replace(dataFacilidade, """", "\""")%>",this.options[this.selectedIndex].text)' onchange='MostraTec(this.value, "<%=Replace(dataString, """", "\""")%>")'>

                    <option value="">:: FACILIDADE </option>
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
                </select>
                <br />
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
                </select>
            </span>
            <div id=divTecnologia name=divTecnologia style="display:none;POSITION:relative"></div>			  
            </td>
        </tr>
        <!--	<tr  class="clsSilver">
		<td id = tdRadio width=170px></td>
		<td colspan = 3><span ID =spnTipoRadio></span></td>
	</tr> -->

        <%'GPON%>
        <!--	<tr  class="clsSilver">
		<td id = tdONT width=170px></td>
		<td colspan = 2><span ID=spncboFabricante></span></td>
		<td colspan = 1><span ID=spncboTipoONT></span></td>
	</tr> -->

        <tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Vel do Acesso Físico</td>
		<td ><span id=spnVelAcessoFis name=spnVelAcessoFis >
                <!-- <select name="cboVelAcesso" style="width:150px" onChange="MostrarTipoVel(this)"> -->
                <select name="cboVelAcesso" style="width: 150px">
                    <option></option>
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
            <td><font class="clsObrig">:: </font>Provedor </td>
            <td colspan="1">
                <select name="cboProvedor">
                    <option value=""></option>
                    <%	
					'response.write "<script>alert('"&strAcessoTipoRede&"')</script>"
					'if isnull(strAcessoTipoRede) then
					'	strAcessoTipoRede = "null"
						'response.write "<script>alert('"&strAcessoTipoRede&"')</script>"
					'end if 
					'set objRS = db.execute("CLA_sp_sel_provedor 0,null,1, " & strAcessoTipoRede & ", " & strOriSol )
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
                    %>
                </select>
            </td>

        </tr>

        <!--	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Provedor</td>
		<td colspan="1">
			<select name="cboProvedor">
				<option value=""></option>
				<%	
					'response.write "<script>alert('"&strAcessoTipoRede&"')</script>"
					'if isnull(strAcessoTipoRede) then
					'	strAcessoTipoRede = "null"
						'response.write "<script>alert('"&strAcessoTipoRede&"')</script>"
					'end if 
					'set objRS = db.execute("CLA_sp_sel_provedor 0,null,1, " & strAcessoTipoRede & ", " & strOriSol )
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
				%>
			</select>
		</td>
		
	</tr> -->

	<tr class=clsSilver2>
		<td width=170px >&nbsp;Endereço Origem&nbsp;</td>
		<td nowrap colspan=3>
			<font class=clsObrig>:: </font>PONTO&nbsp;
				<select name="cboTipoPonto" onchange="TipoOrigem(this.value)">
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
                <input type="text" class="text" name="txtCNLSiglaCentroCli" maxlength="4" disabled="true"
			
                    size="8" tipo="A" value="<%=strCidSigla%>">
                &nbsp;Complemento
			<input type="text" class="text" name="txtComplSiglaCentroCli" maxlength="3" disabled="true"
			
                size="10" tipo="A" value="<%=strEscSigla%>">&nbsp;
						
            </td>
        </tr>
        <tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>UF</td>
            <td>

                <input type="text" class="text" name="cboUFEnd" maxlength="2"
			    size=3" readonly="TRUE" value="<%=strUFEnd%>">
                       
                        
		
            </td>
            <td nowrap right><font class="clsObrig">:: </font>CEP&nbsp;</td>
				<td> <input type="text" class="text" name="txtCepEnd" readOnly="readOnly" value="<%=strCepEnd%>" maxlength="9" size="12">



                <!--Alterado por Fabio Pinho em 28/04/2016 - ver 1.0 - Fim-->




                <td nowrap right></td>
            <td></td>



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

                <input type="text" class="text" name="cboLogrEnd" maxlength="15" readonly="TRUE"
                    size="15" value="<%=strLogrEnd%>">
                    
                   
			
            </td>

        </tr>
        <tr class="clsSilver">
            <td><font class="clsObrig">:: </font>Nome Logr</td>
            <td nowrap>
                <input type="text" class="text" name="txtEnd" readonly="TRUE"
			
                    value="<%=strEnd%>" maxlength="60" size="40">
            </td>
		<td width=170px><font class="clsObrig">:: </font> Número</td>
            <td>
                <input type="text" class="text" name="txtNroEnd" value="<%=strNroEnd%>" readonly="TRUE"
                    maxlength="10" size="10">
            </td>
        </tr>

        <tr class="clsSilver">
            <td>&nbsp;&nbsp;&nbsp;Complemento</td>
		<td colspan= 3 >
                <input type="text" class="text" name="txtComplEnd" value="<%=strComplEnd%>" maxlength="80" size="80">
            </td>
        </tr>
        <tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Bairro</td>
		<td colspan = 3 >
                <input type="text" class="text" name="txtBairroEnd"
			 
                    value="<%=strBairroEnd%>" maxlength="120" size="40" readonly="TRUE">&nbsp;
            </td>


        </tr>
        <!--	<tr class="clsSilver">
		<td colspan=4 align=right><span id=spnCEPSInstala></span></td>
	</tr> -->
        <tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Contato</td>
            <td>
                <input type="text" class="text" name="txtContatoEnd" value="<%=strContatoEnd%>" maxlength="30" size="30">
            </td>
            <td><font class="clsObrig">:: </font>Telefone</td>
            <td>
                <input type="text" class="text" name="txtTelEndArea" value="<%=strTelEndArea%>" maxlength="3" size="3" onkeyup="ValidarTipo(this,0)">&nbsp;

			<input type="text" class="text" name="txtTelEnd" value="<%=strTelEnd%>" maxlength="9" size="11" onkeyup="ValidarTipo(this,0)">
            </td>
        </tr>
        <tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>CNPJ</td>
            <td colspan="3">
                <input type="text" class="text" name="txtCNPJ" maxlength="14"
                    <%=bbloqueia%>
                    size="16" onkeyup="ValidarTipo(this,0)" value="<%=dblCNPJ%>">&nbsp;(99999999999999)
            </td>
            <!--		<td width=170px>&nbsp;&nbsp;&nbsp;</td> -->
            <!--
		<td width=170px>&nbsp;&nbsp;&nbsp;</font>Email Contato Técnico</td>
	    <td>
	      <input type="text" class="text" name="txtEmailTec"  value="<%=strBairroEnd%>" maxlength="30" size="30">&nbsp;
	    </td>
		-->
        </tr>
        <tr class="clsSilver">
		<td width=170px>&nbsp;&nbsp;&nbsp;&nbsp;I.E.</td>
            <td>
                <input type="text" class="text" name="txtIE" maxlength="15"
                    <%if strIE <> "" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>
                    size="17" onkeyup="ValidarTipo(this,0)" value="<%=strIE%>">
            </td>
            <td>&nbsp;&nbsp;&nbsp;I.M&nbsp;</td>
            <td>
                <input type="text" class="text" name="txtIM" maxlength="15"
                    <%if strIM <> "" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <%End if%>
                    size="17" onkeyup="ValidarTipo(this,0)" value="<%=strIM%>">
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
                    maxlength="55" size="81" value="<%=strPropEnd%>">
            </td>
            <!--Alterado por Fabio Pinho em 28/04/2016 - ver 1.0 - Fim-->


        </tr>
        <tr class="clsSilver">
		<td width=170px><font class=clsObrig>:: </font><font color="#FF0000">Interface Cliente</font></td>
            <td colspan="3">
                <select name="cboInterFaceEnd">
                    <option value=""></option>
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
                </select>
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
                        <td>&nbsp;<input type="text" class="text" name="txtComplSiglaCentroCliDest" maxlength="6" size="10" onkeyup="ValidarTipo(this,7)" value="<%=strComplSiglaCli%>" onblur="CompletarCampo(this);ResgatarEstacaoDestino(document.Form2.txtCNLSiglaCentroCliDest,document.Form2.txtComplSiglaCentroCliDest)" TIPO="A">&nbsp;</td>
					<td>&nbsp;<TEXTAREA rows=2 cols=66 name="txtEndEstacaoEntrega" readonly tabIndex=-1></TEXTAREA></td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr class="clsSilver">
		<td width=170px><font class=clsObrig>:: </font><font color="#FF0000">Interface Claro Brasil</font></td>
            <td colspan="3">
                <select name="cboInterFaceEndFis">
                    <option value=""></option>
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
                </select>
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
                                <iframe id="IFrmIDFis1"
                                    name="IFrmIDFis1"
                                    width="100%"
                                    height="65px"
                                    frameborder="0"
                                    scrolling="auto"
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
			
			<input type=button name=btnIDFis1 class=button value="Compartilhar Físico"  onClick="javascript:window.open('AcessoCompartilhar.asp?FlagOrigem=CLA2&txtNroSev='+document.Form1.txtNroSev.value,'janela','toolbar=no,statusbar=no,resizable=yes,scrollbars=YES,width=900,height=800,top=100,left=100')" onmouseover="showtip(this,event,'Procurar um id físico pré existente para o endereço atual (Alt+F)');" accesskey="F" <%if strMudFisico="N" or dblOriSolID = 9 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>  <%=bbloqueiaTipoRede%> > 

                <!--<input type=button name=btnIDFis1 class=button value="Compartilhar" onClick="ProcurarIDFis(1)" onmouseover="showtip(this,event,'Procurar um id físico pré existente para o endereço atual (Alt+F)');" accesskey="F" <%if strMudFisico="N" or dblOriSolID = 9 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>  <%=bbloqueiaTipoRede%> >&nbsp; -->

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
                                <iframe id="IFrmTronco2M"
                                    name="IFrmTronco2M"
                                    width="100%"
                                    height="65px"
                                    frameborder="0"
                                    scrolling="auto"
									align		= "left">
							</iFrame>
                            </td>
                        </tr>
                    </table>
                </div>
            </td>
        </tr>
    </form>
</table>
<table border=0 cellspacing="1" cellpadding="0" width="790" >
    <tr>
        <td>
			<Form name=Form3 method=Post>
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
                <input type="hidden" name="hdnOriSol_ID" value="<%=objRSSolic("OriSol_ID")%>">
                <input type="hidden" name="hdnAprovisi_ID" value="<%=dblAprovisiId%>">
			<input type=hidden name=hdnId_Acesso value="<%=id_acesso%>">
			<input type=hidden name=hdnStrAcfId value="<%=strAcf_id%>" >

			<input type=hidden name=hdnReaproveitarFisico value="N"> <!-- Na ativação será sempre Não. Será modificado na Alteração -->
                <%if objRSSolic("OriSol_ID") <> 9 then%>
                <tr>
                    <th colspan="4">&nbsp;•&nbsp;Informações do Aprovisionador</th>
                </tr>
		  
    </tr>
    <tr class="clsSilver">
				<td nowrap width=170px class="clsSilver">&nbsp;&nbsp;&nbsp;&nbsp;Username <br>&nbsp;&nbsp;&nbsp;&nbsp;do Cadastrador</td>
				<td nowrap class="clsSilver"><input type="text" class="text" name="txtUsernamecadastrador"
                <%=bbloqueia%>
                value="<%=strUsernamecadastrador%>" maxlength="20" size="20" onkeypress="">&nbsp;</td>
				<td nowrap class="clsSilver">&nbsp;Telefone Cadastrador<br>&nbsp;</td>
				<td nowrap class="clsSilver"><input type="text" class="text" name="txtTelefoneCadastrador"
                <%=bbloqueia%>
                value="<%=strTelefoneCadastrador%>" maxlength="10" size="10" onkeypress="">&nbsp;</td>
    </tr>
    <tr class="clsSilver">
        <td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Observação</td>
				<td colspan="3" ><textarea name="txtObs" onkeydown="MaxLength(this,300);" cols="50" rows="3" readonly="true"  <%=bbloqueia%> ><%=strObservacao%></textarea></td>
    </tr>
    <% end if %>
    <%
				'response.write "<script>alert('"&typename(strAcf_id) &"')</script>"
				'--'if  strOriDesc = "ASMS" and strAcao <> "ALT" and Not IsEmpty(strAcf_id) then
				'if  strOriDesc = "ASMS" and Not IsEmpty(strAcf_id) then
				'--'	strMudFisico = "N"
				'--'	response.write "<script>ProcurarIDFis(1);</script>"
				'--'end if  	
				
				if 		strMudFisico = "N" then
					if ( ( strAcao = "ALT" and  Indicador_Alterar = "Y" ) or strAcao = "ATV" ) then
					response.write "<script>ProcurarIDFis(1);</script>"
				end if 		  
				end if 		  
    %>
		</td>
	</tr>
</table>

<table border=0 cellspacing="1" cellpadding="0" width="790" >
    <tr>
        <th colspan="4">&nbsp;•&nbsp;Informações da Embratel</th>
    </tr>
    <tr class="clsSilver">
		<td width="170px" ><font class="clsObrig">:: </font>Local de Entrega<br>&nbsp;&nbsp;&nbsp;do Acesso Lógico</td>
        <input type="Hidden" name="cboLocalEntrega">
		<td colspan="3"><input type="text" class="text" name="txtCNLLocalEntrega"  maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A">&nbsp;
		&nbsp;<input type="text" class="text" name="txtComplLocalEntrega" maxlength="3" size="6" onkeyup="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoUsu(document.Form3.txtCNLLocalEntrega,document.Form3.txtComplLocalEntrega,<%=dblUsuId%>,1);" tipo="A">
        </td>
        <!--		<td colspan="2">&nbsp;</td> -->
    </tr>
    <tr class="clsSilver">
        <td width="170px" nowrap><font class="clsObrig">:: </font>Local de Configuração</td>
        <input type="Hidden" name="cboLocalConfig">
		<td colspan="3"><input type="text" class="text" name="txtCNLLocalConfig"  maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A">&nbsp;
		&nbsp;<input type="text" class="text" name="txtComplLocalConfig" maxlength="3" size="6" onkeyup="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoUsu(document.Form3.txtCNLLocalConfig,document.Form3.txtComplLocalConfig,<%=dblUsuId%>,2);" tipo="A">
        </td>
        <!--		<td colspan="2">&nbsp;</td> -->
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
		<td align=right>Telefone&nbsp;</td>
		<td width=20%>
            <table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="80%">
				<tr><td class="lightblue">&nbsp;
					<span id=spnTelEndLocalInstala><%=strTelEscEntrega%></span>
				</td></tr>
            </table>
        </td>
    </tr>
</table>
<table  border=0 cellspacing="1" cellpadding="0" width="790" >
    <tr class="clsSilver">
		<th colspan=7 >&nbsp;•&nbsp;Coordenação Embratel</th>
    </tr>

    <tr class="clsSilver">
        <td width="170px"><font class="clsObrig">:: </font>UserName GIC-N</td>
        <td colspan="7">
            <input type="text" class="text" name="txtGICN" value="<%=strUserNameGICN%>" maxlength="20" size="20" onblur="ResgatarUserCoordenacao(this)">
        </td>
    </tr>
    <tr class="clsSilver">
        <td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Nome GIC-N</td>
		<td colspan=2 width=50%>
			<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width=355px >
				<tr><td class="lightblue">&nbsp;
					<span id=spnNomeGICN><%=strNomeGICN%></span>
				</td></tr>
            </table>
        </td>
        <td>Telefone&nbsp;</td>
		<td colspan=3 align=left width=20%>

            <table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="100px">
				<tr><td class="lightblue">&nbsp;
					<span id=spnRamalGICN><%=strRamalGICN%></span>
				</td></tr>
            </table>


        </td>

    </tr>
    <tr class="clsSilver">
        <td width="170px"><font class="clsObrig">:: </font>UserName GIC-L</td>
		<td colspan=6>
            <table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="20%">
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
            <table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="100%">
				<tr><td class="lightblue">&nbsp;
					<span id=spnNomeGICL><%=strNomeGICL%></span>
				</td></tr>
            </table>
        </td>
        <td>Telefone&nbsp;</td>
		<td colspan=3 align=left width=20%>
            <table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="100px">
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


<table border=0 cellspacing="1" cellpadding="0"width="790">
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
	<table cellspacing=1 cellpadding=1  width=790 border=0>
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
<table  border=0 cellspacing="1" cellpadding="0" width="790" >
    <tr>
		<td align=center>

            <!--CH-41083NDL - inicio-->

            <!--
				<input type="button" class="button"  style="color:darkred;;font-weight:bold;width:180px" name="btnGravar" value=".::Gravar::." onclick="Gravar()" accesskey="I" onmouseover="showtip(this,event,'Gravar uma solicitação (Alt+I)');" <%if ( (( trim(dblIdLogico) = "" or isnull(dblIdLogico) ) and stracao = "ALT") or desativar_gravar=true) then%>Disabled<%end if%> <%if strbloqcan = true then%>Disabled<%end if%>>&nbsp;
				-->

            <%if bbloqueiaSev = "false" then%>
				<input type="button" class="button"  style="color:darkred;;font-weight:bold;width:180px" name="btnGravar" value=".::Gravar::." onclick="Gravar()" accesskey="I" onmouseover="showtip(this,event,'Gravar uma solicitação (Alt+I)');" <%if ( (( trim(dblIdLogico) = "" or isnull(dblIdLogico) ) and stracao = "ALT") or desativar_gravar=true) then%>Disabled<%end if%> <%if strbloqcan = true then%>Disabled<%end if%>>&nbsp;
			<%End if%>
            <!--CH-41083NDL - fim-->

            <!--<input type="button" class="button" name="btnNovaSolic" value="Nova Solicitação" style="width:100px" onclick="javascript:window.location.replace('solicitacao.asp')" accesskey="N" onmouseover="showtip(this,event,'Nova solicitação (Alt+N)');">&nbsp;-->
            <%if Trim(Request.Form("hdnAcao")) = "Alteracao" then%>
			<input type=button	class="button" name=btnVoltar value=Voltar onclick="VoltarOrigem()" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">&nbsp;
			<%End if%>
            <input type="button" class="button" name="btnSair" value=" Sair " onclick="javascript:window.location.replace('main.asp')" style="width: 100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
        </td>
    </tr>
    <tr>
        <td>
            <font class="clsObrig">:: </font>Campos de preenchimento obrigatório.
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
<iframe id="IFrmProcesso"
    name="IFrmProcesso"
    width="0"
    height="0"
    frameborder="0"
    scrolling="no"
		align       = "left">
</iFrame>
<iframe id="IFrmProcesso2"
    name="IFrmProcesso2"
    width="0"
    height="0"
    frameborder="0"
    scrolling="no"
		align       = "left">
</iFrame>
<iframe id="IFrmProcesso3"
    name="IFrmProcesso3"
    width="0"
    height="0"
    frameborder="0"
    scrolling="no"
		align       = "left">
</iFrame>



<script language="JavaScript">
    MostrarVlanProvedor()
        <%
        Set objRS = db.execute("CLA_sp_sel_usuario 0,'" & Trim(strUserName) & "'")
    if Not 	objRS.Eof And Not objRS.Bof then
    Response.Write "spnNomeGICL.innerHTML = '" & Replace(Trim(objRS("Usu_Nome")), "'", "´") & "';"
    Response.Write "spnRamalGICL.innerHTML = '" & Replace(Trim(Cstr("" & objRS("Usu_Ramal"))), "'", "´") & "';"
    Response.Write "document.forms[2].hdntxtGICL.value = '" & Trim(strUserName) & "';"
End if
%>
//Geral
with (document.forms[0])
{
	<%if Trim(Request.Form("hdnAcao")) <> "Alteracao" or Request.Form("hdnOEOrigem") = "APG"then %>
                txtOrderEntry = document.getElementsByName("txtOrderEntry")[0];
            if (txtOrderEntry) {
                txtOrderEntry[0].readOnly = true
                txtOrderEntry[1].readOnly = true
                txtOrderEntry[2].readOnly = true
            }
	<% Else %>

                ResgatarDesigServicoGravado(<%=dblSerId %>)
                <% End if%>
}
</script>
<!--Form que envia os dados para gravação-->
<TABLE border=0>
<tr><td>
            <form method="post" name="Form4">
<input type=hidden name=hdnAcao>
<input type=hidden id=strAcessoTipoRede name=strAcessoTipoRede value="<%=strAcessoTipoRede%>">

<input type=hidden name=hdnUsuario>
<input type=hidden name=hdnOrigemEst>
<input type=hidden name=hdnCNLEstUsu>
<input type=hidden name=hdnComplEstUsu>
<input type=hidden name=hdnAntAcesso>
<input type=hidden name=hdnCboServico>
<input type=hidden name=hdnDesigServ>
<input type=hidden name=hdnTipoAcao value="<%=Request.Form("hdnAcao")%>" >
                <!-- Good  início -->
                <input type="hidden" name="hdnXml" id="hdnXml" value="">
                <!-- Good  Fim -->
<input type=hidden name=hdnApg value="<%=Request.Form("hdnAcao")%>">
<input type=hidden name=hdnOrigem value="<%=strOrigem%>">
<input type=hidden name=hdnIdAcessoLogico value="<%=dblIdLogico%>">
<input type=hidden name=hdnSolId value="<%=dblSolId%>">
<input type=hidden name=hdn678 value="<%=Request.Form("hdn678")%>">
<input type=hidden name=hdnStrAcfId value="<%=strAcf_id%>" >
                <%if Trim(Request.Form("hdnAcao")) = "Alteracao" then%>
<input type=hidden name=hdnTipoProcesso value=3>
                <%Else%>
<input type=hidden name=hdnTipoProcesso value=1>
                <%End if%>
                <input type="hidden" name="hdnOriSol_ID" value="<%=objRSSolic("OriSol_ID")%>">
                <input type="hidden" name="hdnAprovisi_ID" value="<%=dblAprovisiId%>">
<input type=hidden name=hdnId_Acesso value="<%=id_acesso%>">
<input type=hidden name=hdnVelIdServicoOld value="<%=dblVelServico%>">
                <!-- DAVIF - 16/01/2006 -->
<input type=hidden name=hdnIdInterfaceAPG value="<%=dblSolAPGId%>">
                <!-- LP -->
            </form>
        </td>
    </tr>
</table>
<%
'*********** MIGRAÇÂO HFC BSoD **********
 
If (Not IsNull(strIDLogico) and Trim(strIDLogico) <> "") and stracao = "ALT" Then 
 
			'Set objMigr = db.execute("select TOP 1 cla_provedor.Pro_Nome,cla_acessofisico.Acf_ID,Rec_Id,Pro_ID,Reg_ID,Acf_NroAcessoPtaEbt,Acf_IDAcessoFisico,Fac_SVlan,Fac_PE,Fac_Vlan,Fac_Porta from  where OriSol_Descricao = '" & Trim(objRSSolic("OriSol_Descricao")) & "'")
			
			'strIDLogico = "6787336597"
			Set objMigr = db.execute("select TOP 1 cla_provedor.Pro_Nome,cla_acessofisico.Acf_ID,Rec_Id,cla_acessofisico.Pro_ID,Reg_ID,Acf_NroAcessoPtaEbt," &_
			"Acf_IDAcessoFisico,Fac_SVlan,Fac_PE,Fac_Vlan,Fac_Porta from cla_facilidade inner join cla_acessofisico on " &_
			"cla_facilidade.acf_id=cla_acessofisico.acf_id inner join cla_acessologicofisico on cla_acessologicofisico.acf_id=cla_acessofisico.acf_id " &_
			"inner join cla_provedor on  cla_provedor.pro_id=cla_acessofisico.pro_id where cla_acessofisico.pro_id in(136,137) and " &_
			"acf_dtdesatacessofis is null and acf_dtcancacessofis is null and alf_flag is null and acl_idacessologico='" & strIDLogico & "'")

			If Not objMigr.Eof Then
			'response.write "<script>alert('strAcessoTipoRede=" & strAcessoTipoRede & "')</script>"
			  'Solicitações oriundas do SGA-Plus com marcação para BSOD NET deverão ser liberadas.
				If Not (strAcessoTipoRede = 4 and objMigr("Pro_ID") = 136) Then
					response.write "<script>alert('ATENÇÃO \nEstão bloqueadas novas alterações para provedores HFC BSoD VIA e HFC BSoD NET. \nAssim, é necessário realizar a migração do acesso físico para HFC BSoD, através da funcionalidade de Migração, que será exibida após o OK desta mensagem. \nEm caso de dúvidas, contatar a equipe do SOE.')</script>"

'			    Set objRSMigr_Compart = db.execute("select count(acfl.acl_idacessologico) As Qtd_ACL from cla_acessofisico acf inner join cla_acessologicofisico acfl on acf.acf_id=acfl.acf_id " &_
'			                                       "where alf_flag is null and acf.acf_idacessofisico='" & objMigr("Acf_IDAcessoFisico") & "'")			
'			    If CInt(objRSMigr_Compart("Qtd_ACL"))>1 Then					
'				     response.write "<script>alert('Migração não permitida para Acesso Físico com compartilhamento.\n\nPara migrar o Acesso Físico, você deverá desativar ou remanejar os Acessos Lógicos vinculados, deixando apenas um Acesso Lógico vinculado ao Acesso Físico [" & objMigr("Acf_IDAcessoFisico") & "].')</script>" 
'			    End If								
				
   				response.write "<script>"   			
   				response.write "window.location='migracaoBsod.asp?acao=consulta&txtAcf_IdAcessoFisico=" & objMigr("Acf_IDAcessoFisico") & "'"
   				response.write "</script>"
   			End If
			End If
			
'***********************lembrar de atualizar o regime de contrato do acf: cboRegimeCntr !!!
End If%>

<script type="text/javascript" language="JavaScript">
    document.getElementById("txtNroSev").focus();
</script>

</body>
<%DesconectarCla()%>
</html>
