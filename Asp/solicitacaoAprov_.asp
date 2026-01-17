<%
'	- Sistema			: CLA
'	- Arquivo			: solicitacaoAprov.ASP
'	- Descri��o			: Cadastra/Altera uma solicita��o do Aprovisionamento no sistema CLA
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%

''@@ LPEREZ 13/10/2005
Dim strVisada			'Tipo Visada
Dim strGrupo			'Grupo Cliente
Dim strOriSol			'Origem Solicitacao
Dim strProjEspecial	' Projeto Especial
Dim strNecessitaRecurso		'Se ser� necess�rio recurso (Sim/N�o)
Dim strEmiteOTS				'Se ser� necess�rio emitir OTS (Sim/N�o)
Dim strReaprFisico			'Reaproveitar Fisico (Sim/N�o)
Dim strUsernamecadastrador	'UserName do Cadastrador
Dim strtelefonecadastrador	'Telefone do Cadastrador
Dim strorigem				'OrigemSolicitacao (4=APG/x=SGAPlus, etc..)
Dim strOrderEntryNumSis		'Numero de Identifica��o da OE no Sistema Origem (APG, etc..)
Dim bbloqueia				'Variavel de controle para Bloquear campos
Dim bdesbloqueia			'Variavel de controle para para Desbloquear campos
Dim dblSolAPGId				'IDentificador da Solicita��o APG
Dim readOnly				'Variavel de controle para Travar campos
Dim strMudFisico            'Variavel para mudan�a no acesso f�sico
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
Dim dblNroSev			'N�mero da Sev do sistema SSA
Dim strRazaoSocial		'Raz�o social
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
Dim strOrderEntryNro	'N�mero da OrderEntry
Dim strOrderEntryItem	'Item da OrderEntry
Dim strDtPedido			'Data do pedido
Dim dblVelServico		'Id da Velocidade do servi�o
Dim strTipoContratoServico'Tipo do cantrato
Dim strNroContrServico	'N�mero do contrato
Dim dblIdLogico			'N�mero do acesso l�gico
Dim dblDesigAcessoPri	'Designa��o do acesso principal
Dim strDtEntrAcesServ	'Data de entrega do acesso ao servi�o
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

Dim dblSolId			'ID da Solicita��o SOL_ID
Dim strItemSel			'Controle para o item que esta selecionado nos combos (selected)
Dim dblSerId			'ID do servi�o (Ser_Id)
Dim strPropEnd			'Propriet�rio do endere�o
Dim strEndCid			'Sigla da cidade
Dim dblAcaId			'ID da A��o Aca_ID
Dim strPropAcessoFisico 'Propriet�rio do acesso f�sico
Dim dblTecId			'ID da tecnologia Tec_id
Dim dblProId			'ID do Provedor Pro_id
Dim dblRegId			'ID do Regime de contrato Reg_Id
Dim dblPrmId			'ID da promo��o Prm_Id
Dim strObsProvedor		'Obeserva��es para o provedor
Dim strEnd				'Nome do logradouro
Dim strComplEnd			'Complemento do logradouro
Dim strBairroEnd		'Bairro do logradouro
Dim strCepEnd			'CEP do logradouro
Dim strContatoEnd		'Conotao do logradouro
Dim strTelEnd			'Telefone  do logradouro
Dim strTelEndArea		'Area do Telefone  do logradouro
Dim strUFEnd			'UF do logradouro
Dim strNroEnd			'N�mero do logradouro
Dim strLogrEnd			'Sigla do logradouro
Dim strInterFaceEnd 	'Interface do logradouro
Dim strEndCidDesc		'Decri��o da cidade do logradouro
Dim dblOrgId			'ID do org�o Org_Id
Dim dblStsId			'Id do Status Sts_Id
Dim strHistoricoSol		'Hist�rico da solicita��o
Dim strPropAcessoFis	'Propriet�rio do acesso para o id f�sico gravado (Instala��o)
Dim strVelAcesso		'Velocidade do acesso para o id f�sico gravado (Instala��o)
Dim strDtIniTemp		'Data de inicio do acesso tempor�rio
Dim strDtFimTemp		'Data de fim do acesso tempor�rio
Dim strDtDevolucao		'Data de entrega do acesso tempor�rio
Dim strDtTesteFimaFim	'Data Prevista do Teste Fim a Fim
Dim dblLocalEntrega 	'ID do Local de Entrega Esc_Id
Dim dblLocalConfig		'ID do Local de Configura��o Esc_Id
Dim strInterfaceEbt		'Interface na EBT
Dim strContEscEntrega	'Contato no local de entrega
Dim strTelEscEntrega	'Telefone do contato no local de entrega
Dim objRSSolic			'Dados da solicitac�o em edi��o
Dim DBAction1			'A��o auxiliar
Dim objRSFis			'Acessos f�sicos
Dim strIdAcessoFisicoInst 'Id do Acesso f�sico de instala��o
Dim strVelDescAcessoFisicoInst'Velocidade do Acesso f�sico de instala��o
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

'<!--CH-41083NDL - inicio-->
Dim bbloqueiaSev				'Variavel de controle para Bloquear SEV quando a tela nao pode liberar a gravacao
bbloqueiaSev = "false"
'<!--CH-41083NDL - Fim-->

Set objDicProp = Server.CreateObject("Scripting.Dictionary")

'Monta o Xml de Acessos
%>
<!--#include file="../inc/xmlAcessos.asp"-->
<script>
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

'Perfis que podem acessar o essa p�gna E-GICL,GE-Ger.Usuario,GAT-GLA
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
			
			if dblOriSolID = "10" then 
				'Localiza informa��es da solu��o SSA
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
					
					
					strHigienizado = "N�O"
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
					strNroEnd		= Trim(objRSCli("PRE_NROLOGR"))			'N�mero do logradouro
					strLogrEnd		= Trim(objRSCli("TPL_SIGLA"))				'Sigla do logradouro
					strEndCid		= Trim(objRSCli("CID_SIGLA"))					'Sigla da cidade do logradouro
					strEndCidDesc	= Trim(objRSCli("LOCALIDADE"))				'Decri��o da cidade do logradouro
					
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
			else
			
					desativar_gravar = false
					
					strHigienizado = "SIM"
					
					'#Raio X#'
					strEnd			= Trim(objRSSolic("End_NomeLogr"))			'Nome do logradouro
					strComplEnd		= Trim(objRSSolic("Aec_Complemento"))		'Complemento do logradouro
					strBairroEnd	= Trim(objRSSolic("End_Bairro"))			'Bairro do logradouro
					strCepEnd		= Trim(objRSSolic("End_CEP"))				'CEP do logradouro
					
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
					strUFEnd		= Trim(objRSSolic("Est_Sigla"))	 			'UF do logradouro
					strNroEnd		= Trim(objRSSolic("End_NroLogr"))			'N�mero do logradouro
					strLogrEnd		= Trim(objRSSolic("Tpl_Sigla"))				'Sigla do logradouro
					strEndCid		= Trim(objRSSolic("CNL"))					'Sigla da cidade do logradouro
					strEndCidDesc	= Trim(objRSSolic("Cidade"))				'Decri��o da cidade do logradouro
	
			
			
			end if 
			
			
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
			
			'C�digo SAP
			strCodSap = Trim(objRSSolic("esc_cod_sap"))
			
			'Informa��es da OE
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
			''' Inicio da implementa��o redirecionamento GLA para Interface
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
				'strHtmlGla	= strHtmlGla & "<span id=spnUserNameGLA><font color=red>N�o Encontrado</font></span>"
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
			''' Fim da implementa��o redirecionamento GLA para Interface
			'''
			
			
		End If
		
	End If


	If Trim(Request.Form("hdnAcao")) = "Alteracao" then
	
		If Trim(Request.Form("hdnAcao")) = "Alteracao"  And (dblSolId = "" or dblSolId = "0") Then
			Response.Write "<script language=javascript>alert('Solicita��o indispon�vel.');window.location.replace('main.asp');</script>"
			Response.End	
		End if
	
		if not isnull(dblSolId) then
			'Adicionado PRSS 29/01/2007:
			strSQL = "select Ctf_ID from cla_solicitacaoctf WITH (NOLOCK) where sol_id = " & dblSolId
		    Set objRSCtfc = db.execute(strSQL)
			dblCtfcId = objRSCtfc("Ctf_ID")
		end if
	
		if objRSSolic.Eof then
			Response.Write "<script language=javascript>alert('Solicita��o indispon�vel.');window.location.replace('main.asp');</script>"
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
				strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))
				If trim(strLoginRede) = "EDAR" Then
					Set doc = server.CreateObject("Microsoft.XMLDOM")
					doc.loadXml(strXmlAcesso)
					doc.save(Server.MapPath("acesso_ant.xml"))
				end if
			End if
		End if
	
	End If
	
%>
<script language='javascript' src="../javascript/solicitacao.js"></script>
<script language='javascript' src="../javascript/xmlObjects.js"></script>
<script language='javascript' src="../javascript/help.js"></script>

<SCRIPT LANGUAGE=javascript>
<!--
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

		//Qdo. for processo de altera��o Volta para tela Inicial da solicta��o
		if	(document.Form4.hdnTipoProcesso.value == 3 || document.Form4.hdnOrigem.value == "APG" ||  document.Form4.hdnOrigem.value == "Aprov")
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

//Verifica se o Xml j� esta carregado
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

//Verifica se o Xml j� esta carregado
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

//Verifica se o Xml j� esta carregado
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
		strMigraPabx_Virtual = "N�o"
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
		strTecnologia_IP = "N�o"
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


<div id="divWait" style="background-color:#dcdcdc; width:300px; height:100px; float:left; margin:800px 0 0 280px; position:absolute; border:1px solid #0f1f5f; padding:40px 0 0 20px; display:none;">
	<p align="center" style="font-size: 12px; font-family:Arial, Helvetica; font-weight: bold; color:#003366;">Aguarde. Estamos consultando o CSL ...</p>
</div>

<div id="divSistemaWait" style="background-color:#dcdcdc; width:300px; height:100px; float:left; margin:800px 0 0 280px; position:absolute; border:1px solid #0f1f5f; padding:40px 0 0 20px; display:none;">
	<p align="center" style="font-size: 12px; font-family:Arial, Helvetica; font-weight: bold; color:#003366;">Aguarde. Sistema Processando ...</p>
</div>

<%
'JCARTUS@ Dez/2011 - Abrir para o usu�rio digitar IDL�gico quando o informado na interface estiver 
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
	<table cellspacing="1" cellpadding="0" border=0 width="763">
		<th>
			<th nowrap colspan="4" style=" BACKGROUND-COLOR: #A80000">&nbsp;�&nbsp;Informa��es de Aprovisionamento - Dados do Item de OE de Origem para facilitar busca do Acesso L�gico</th>
		</th>
	</table>
	
	<table cellspacing="1" cellpadding="0" border=0 width="760">
		<tr class="clsSilver">
		  <td width="465">Raz�o Social: <%=strPov_Razao%></td>
		  <td colspan="2">Conta Corrente: <%=strPov_CC%></td>
		  <td width="96">SubConta: <%=strPov_SubCC%></td>
		</tr>
	</table>
	<table cellspacing="1" cellpadding="0" border=0 width="760">
		<tr class="clsSilver">
		  <td width="298">Designa��o do Servi�o: <%=strPov_DesigServ%></td>
		  <td width="426" colspan="2">Velocidade do Servi�o: <%=strPov_VelServ%></td>
		  <td width="249">N� Contrato de Servi�o: <%=strPov_ContratoServ%></td>
		</tr>
	</table>
	<table cellspacing="1" cellpadding="0" border=0 width="760">
		<tr class="clsSilver">
		  <td width="228">Contato Cliente: <%=strPov_ContatoCli%></td>
		  <td width="161">Telefone: <%=strPov_Tel%></td>
		  <td width="169">CNPJ: <%=strPov_CNPJ%></td>
		  <td width="237">Prop. End.: <%=strPov_PropEnd%></td>
		</tr>
	<table cellspacing="1" cellpadding="0" border=0 width="760">
		<tr>
		  <th nowrap colspan="4" style=" BACKGROUND-COLOR: #A80000">&nbsp;�&nbsp;POVOAMENTO SOB DEMANDA: Associa��o de Servi�os enviados pelo Aprovisionador com Acesso L�gico cadastrado no CLA</th>
		</tr>
	</table>
	<table cellspacing="1" cellpadding="0" border=0 width="760">
		<tr class="clsSilver">
		  <td width="517">Informar o Acesso L�gico que ser� associado ao servi�o e que sofrer� o processo de 
			<%
			select case trim(ucase(strAcao))
			  Case "ATV"
			    response.write "ativa��o"
			  Case "DES"
			    response.write "desativa��o"
			  Case "ALT"
			    response.write "altera��o"
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
				<input type="button" class="button" name="associarlogico" value="Associar L�gico" title="Associar Acesso L�gico" style="cursor:hand" onClick="AssociarLogico()" tabindex=0 accesskey="A" onmouseover="showtip(this,event,'Associar L�gico (Alt+A)');">
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
'Corre��o 15 min do schedule de cancelamento autom�tico.
if (strIDLogico <> "" and (strAcao = "ALT" or strAcao = "ATV"))  then

	strIDLogico677 = "677" & mid(strIDLogico,4,7)
	strIDLogico678 = "678" & mid(strIDLogico,4,7)

	strSQL = "select top 1 Acl_IDAcessoLogico from cla_aprovisionador tab1 WITH (NOLOCK) where acao in ('ALT','ATV') and Acl_IDAcessoLogico in (select Acl_IDAcessoLogico from cla_aprovisionador as tab2	WITH (NOLOCK) where tab1.ID_Tarefa = tab2.ID_Tarefa_Can and acao='CAN' and ID_Tarefa_Can is not null and not exists (select top 1 * from cla_aprovisionador tab3 WITH (NOLOCK) where tab3.ID_Tarefa = tab2.ID_Tarefa_Can and aprovisi_dtCancAuto is not null)) and Acl_IDAcessoLogico in (" & strIDLogico677 & "," & strIDLogico678 & ")"

	Set objRSCan = db.execute(strSQL)
	
	if not objRSCan.Eof then
	
	strbloqcan = true
	%>
	<table cellspacing="1" cellpadding="0" border=0 width="760">
		<tr>
		  <th nowrap colspan="4" style=" BACKGROUND-COLOR: #A80000">&nbsp;�&nbsp;REGISTRO BLOQUEADO (15MIN): Aguardando cancelamento autom�tico do registro anterior 
			&nbsp;�&nbsp;&nbsp;<input type=button class="button" name="btnrefresh" value="Atualizar p�gina" onclick="JavaScript:location.reload(true);"></th>
		</tr>
	</table>
	<%
	end if
end if
%>


<table cellspacing="1" cellpadding="0" border=0 width="760">
	<tr >
		<th width=25%>&nbsp;�&nbsp;Solicita��o de Acesso</th>
		<th width=25%>&nbsp;N�&nbsp;:&nbsp;<span id=spnSolId><%=dblSolId%></Span></th>
		<%if Trim(Request.Form("hdnAcao")) = "Alteracao" then%>
			<th width=25%>&nbsp;Tipo&nbsp;:&nbsp;Altera��o</th>
			<th width=25%>&nbsp;Acesso L�gico&nbsp;:&nbsp;<%=dblIdLogico%></th>
		<%Elseif Trim(Request.Form("hdnOrigem")) = "Recadastro" then%>
			<th width=25%>&nbsp;Tipo&nbsp;:&nbsp;Recadastro</th>
		<%else%>
			<th width=25%>&nbsp;Tipo&nbsp;:&nbsp;Ativa��o</th>
		<%End if%>
		<th width=25%>&nbsp;Data&nbsp;:&nbsp;<%=strDtPedido%></th>
	</tr>
</table>

<table border=0 cellspacing="1" cellpadding="0" width="760">
	<tr >
		<th colspan=4>&nbsp;�&nbsp;Informa��es do Cliente </th>
		<th width="27">
			<div id="ajuda1" style="position:absolute; left:-1000px; top:-1000px; width:234; height:223; z-index:1;visibility: hidden; border: 0px none;">
		    <object name="button" classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="https://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="234" height="223">
            	<param name="movie" value="../imagens/Help/help.swf?arquivo=../imagens/Help/SolicitacaoAprov.txt">
    			<param name="quality" value="high">
    		<embed src="../imagens/Help/help.swf?arquivo=../imagens/Help/SolicitacaoAprov.txt" quality="high" pluginspage="https://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="234" height="223"></embed>
 			</object>
			</div>
						
			<div id="espaco_left" title="Ajuda">
			<a href="javascript:;" onMouseOver="showAjudaImg('duvida1','ajuda1',10,20 );" onClick="hideAjuda('ajuda1');">
			<img src="../imagens/question_b.gif" border="0" id="duvida1" style="cursor:pointer"/></a>
			</div>
		</th>
	</tr>
	<tr class="clsSilver">
		<td width="170">&nbsp;&nbsp;&nbsp;&nbsp;Sev para procura</td>
		<td colspan="4">
			<input type="hidden" name="hdnOriSol_ID_N" value="<%=objRSSolic("OriSol_ID")%>">
			
			<!--Alterado por Fabio Pinho em 28/04/2016 - ver 1.0 - Inicio-->
			<%if Trim(strOrderEntrySis) = "CFD" then%>

				<input type="text" class="text" name="txtNroSev" readOnly="readOnly" value="<%=dblNroSev%>" maxlength="8" size="10" onkeyup="ValidarTipo(this,0)" onblur="ResgatarSev()">&nbsp;

			<%else%>

				<!--CH-41083NDL - inicio-->
				<%if bbloqueiaSev = "true" then%>			

					<input type="text" class="text" name="txtNroSev" value="<%=dblNroSev%>" maxlength="8" size="10" onkeyup="ValidarTipo(this,0)" <%=bbloqueia%> onblur="ResgatarSev()">&nbsp;

				<%else%>

					<input type="text" class="text" name="txtNroSev" value="<%=dblNroSev%>" maxlength="8" size="10" onkeyup="ValidarTipo(this,0)" onblur="ResgatarSev()">&nbsp;


				<%End if%>
				<!--CH-41083NDL - fim->

			<%End if%>
			<!--Alterado por Fabio Pinho em 28/04/2016 - ver 1.0 - Fim-->
			
				<!--RAIO-x: BOT�O DE REAN�LISE DE SEV HABILITADO TEMPOR�RIAMENTE -->
				<input type="button" class="button" name="procurarsev" value=" Rean�lise de Sev  " onClick="alert('O processo de rean�lise de SEV sofreu altera��o: O pedido deve ser realizado diretamente no Star21 atrav�s da inser��o das pend�ncias 23.X, sem necessidade de atua��o no SSA e CLA.');" tabindex=-1 accesskey="P" > 
			
		</td>
	</tr>
<!-- ''@@ LPEREZ - 24/10/2005 -->
	<tr>
		<td class="clsSilver" rowspan="2">&nbsp;
			Projeto Especial
		</td>
		<td class="clsSilver" rowspan="2">&nbsp;
			<input type="radio" name="rdoProjEspecial" onClick="javascript:document.Form1.hdnProjEsp.value = 'S';" value="S" <%if strProjEspecial = "S" then%> checked <%end if%>>&nbsp; Sim
			<input type="radio" name="rdoProjEspecial" onClick="javascript:document.Form1.hdnProjEsp.value = 'N';" value="N" <%if strProjEspecial <> "S" then%> checked <%end if%>>&nbsp; N�o
		</td>
		<td colspan=3 class="clsSilver">
		    &nbsp;&nbsp;Grupo <span align=right>
		<select name="cboGrupo" onChange="CheckGrupo()" >
			<option value=""></option>
			<%
			set gr = db.execute("CLA_sp_sel_GrupoCliente 0")
			do while not gr.eof
			%>
				<option value="<%=gr("GCli_ID")%>"
			<%
				if strGrupo <> "" then
					if trim(strGrupo) = trim(gr("GCli_ID")) then
						response.write "selected"
					end if
				end if
			%>
				><%=ucase(gr("GCli_Descricao"))%></option>
			<%
				gr.movenext
				loop
			%>
		</select>
			</span>&nbsp;	
		</td>
	</tr>
	<tr>
		</td>
		<td colspan=3 class="clsSilver">Origem Solicita��o <span align=right>
			<select name="cboOrigemSol" <%=bbloqueia%>>
			    <%if Trim(Request.Form("hdnOrigem")) = "Recadastro" then%>
				<option value="8">RECADASTRO</option>
				<%else%>
				<option value="<%=objRSSolic("OriSol_ID")%>"><%=objRSSolic("OriSol_Descricao")%></option>
				<%end if%>
		  	</select>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170"><font class="clsObrig">:: </font>Raz�o Social</td>
		<td colspan="4" >
			<input type="text" class="text" name="txtRazaoSocial"  maxlength="80" size="80" value="<%=strRazaoSocial%>"
			<%=bbloqueia%> onblur="ResgatarGLA()">

			
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170"><span id=spnLabelCliente></span></td>
		<td colspan="4"><span id=spnCliente></span></td>
	</tr>
	<tr class="clsSilver">
		<td width="170"><font class="clsObrig">:: </font>Nome Fantasia</td>
		<td>
			<input type="text" class="text" name="txtNomeFantasia"  maxlength="40"
			<%=bbloqueia%>
			size="40" value="<%=strNomeFantasia%>" >
		</td>
		<td colspan=3 class="clsSilver">&nbsp;</td>
<!-- ''@@ -->
	</tr>
	<tr class="clsSilver" nowrap>
		<td width="170" nowrap><font class="clsObrig">:: </font>Conta Corrente</td>
		<td width=279>
			<input type=text class="text" name=txtContaSev size=11 maxlength=11
			<%=bbloqueia%>
			onKeyUp="ValidarTipo(this,0)" onblur="CompletarCampo(this)" TIPO="N" value="<%=strContaSev%>">
		</td>
		<td align=right width=89 >
		<p align="left"><font class="clsObrig">:: </font>Sub Conta&nbsp;</td>
		<td width="189">
			<input type=text name=txtSubContaSev class="text" size=4 maxlength=4
			<%=bbloqueia%>
			onKeyUp="ValidarTipo(this,0)" onblur="CompletarCampo(this)" TIPO="N" value="<%=strSubContaSev%>">
		</td><td>&nbsp;</td>
	</tr>
	<tr class="clsSilver" nowrap>
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
		</td><td>&nbsp;</td>
	</tr>
	<!--
	<tr class="clsSilver" nowrap>
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Observa��o para Rean�lise de SEV</td>
		<td ><textarea name="txtObsReanaliseSEV" onkeydown="MaxLength(this,300);" cols="50" rows="3"></textarea></td>
		<td colspan=3 class="clsSilver">&nbsp;</td>
	</tr>
	-->
</table>
<table border=0 cellspacing="1" cellpadding="0" width="760">
	<tr >
		<th colspan=4 >
			&nbsp;�&nbsp;Informa��es do Servi�o&nbsp;
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
					<select name="cboSistemaOrderEntry"	<%=bbloqueia%> onChange="SistemaOrderEntry(this);hdnOrderEntrySis.value=this.value;" >
						<%if objRSSolic("OriSol_ID") = 7 then%>
							<Option value="SGA VOZ VIP'S" selected>SGA VOZ VIP'S</Option>
						<%elseif objRSSolic("OriSol_ID") = 6 then%>
							<Option value="SGA PLUS" selected>SGA PLUS</Option>
						<%else%>
							<Option value="ASMS" selected>ASMS</Option>
						<%end if%>
					</Select>
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
					<select name="cboSistemaOrderEntry"	<%=bbloqueia%> onChange="SistemaOrderEntry(this);hdnOrderEntrySis.value=this.value;" >
						<%if objRSSolic("OriSol_ID") = 10 then%>
							<Option value="CFD" selected>CFD</Option>
						
						<%end if%>
					</Select>
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
		<td width="170px"><font class="clsObrig">:: </font>Servi�o</td>
		<td width="100px" colspan=3>
		<%
			set objRS = db.execute("CLA_sp_sel_servico null,null,null,1")
		%>
			<select name="cboServicoPedido" <%=bbloqueia%> onchange="document.Form1.rdoAntAcesso[1].checked = true;ResgatarServico(this)" tyle="width:100px">
				<option ></option>
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
			<%IF strPabx_Virtual <> "" then %><br><BR>PABX VIRTUAL BROADSOFT: <%= strPabx_Virtual %><br><BR><%END IF%>
		</td>
		
	</tr>

	<!-- Projeto ABW - Incio -->
	<tr class="clsSilver">
		<td nowrap width=170px>Migra��o do PABX Virtual</td>
		<td>&nbsp;<%=strMigraPabx_Virtual%></td>
		<td nowrap align=right>&nbsp;Tecnologia IP</td>
		<td >&nbsp;<%=strTecnologia_IP%></td>
	</tr>
	<!-- Projeto ABW - Fim -->
		
	<tr class="clsSilver">
		<td width="150px" align=left><font class="clsObrig">:: </font>Velocidade&nbsp;</td>
		<td colspan=3 ><span id=spnVelServico>
				<select name="cboVelServico" <%=bbloqueia%> onChange="SelVelAcesso(this)" style="width:110px">
					<option ></option>
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
		<td width="170px"><font class="clsObrig">:: </font>Antecipa��o de Acesso </td>
		<td >
			<input  type="radio"  name="rdoAntAcesso" value="S" onclick="document.Form1.rdoAntAcesso[0].checked=true;document.Form1.rdoAntAcesso[1].checked=false;DesabilitarDesignacao(1)" disabled>&nbsp; Sim
			<input  type="radio"  name="rdoAntAcesso" value="N" onclick="document.Form1.rdoAntAcesso[0].checked=false;document.Form1.rdoAntAcesso[1].checked=true;DesabilitarDesignacao(2)" checked disabled>&nbsp; N�o

		</td>		
		<% if dblOriSolID = 9 then %>
		<td width="150px" align=right><font class="clsObrig">:: </font>Velocidade Total </td>
		<td >
				<input type="text" class="text" name="txtVelocidadeTotal"
				<%=bbloqueia%>
				value="<%=strVelocidadeTotal%>" maxlength="22" size="30"><br>
		</td>
		<%else %>
			<td width="150px" align=right></td>
			<td ></td>
		<% end if %>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Designa��o do Servi�o</td>
		<td colspan="1">
				<input type="text" class="text" name="txtdesignacaoServico"
				<%=bbloqueia%>
				value="<%=strDesignacaoServico%>" maxlength="22" size="30"><br>
		</td>
		<%
		'EF: VEL_VOZ - esse novo campo n�o ser� obrigat�rio. Caso venha preenchido dever� ser exibido
		'abaixo da velocidade do servi�o j� existente.
		'A sua fun��o � apoiar a decis�o do PVT na escolha da velocidade do acesso f�sico.
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
		<td width="150px" align=right><font class="clsObrig"> </font>&nbsp;</td>
		<td>
		</td>
		<%
		END IF
		'GPON - FIM%>
	</tr>	

	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>N� Contrato Servi�o</td>
		<td colspan=3>
			<table rules="groups" cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="70%" >
				<tr><td nowrap width=200px >
					<input type=radio name=rdoNroContrato value=1
					<%=bbloqueia%>
					onClick="spnDescNroContr.innerHTML= 'Ex.: VEM-11 XXX000012003'" checked <%if strTipoContratoServico = "1" then Response.Write " checked " End if%>>Contrato de Servi�o</td><td></td></tr>
				<tr>
					<td nowrap><input type=radio name=rdoNroContrato onClick="spnDescNroContr.innerHTML= 'Padr�o: A22'"
					<%=bbloqueia%>
					value=2 <%if strTipoContratoServico = "2" then Response.Write " checked " End if%>>Contrato de Refer�ncia</td>
					<td nowrap>
						<input type="text" class="text" name="txtNroContrServico" value="<%=strNroContrServico %>"
						<%=bbloqueia%>
						maxlength="22" size="30"><br>
						<span id=spnDescNroContr>Ex.: VEM-11 XXX00012003</span>
					</td>
				</tr>
				<tr><td nowrap><input type=radio name=rdoNroContrato
				<%=bbloqueia%>
				onClick="spnDescNroContr.innerHTML= 'Padr�o: A22'" value=3 <%if strTipoContratoServico = "3" then Response.Write " checked " End if%> >Carta de Compromisso</td><td></td></tr>
			</table>
		</td>
	</tr>
	<tr class="clsSilver">
		<td nowrap width=170px><font class="clsObrig">:: </font>Data Desejada de Entrega<br>&nbsp;&nbsp;&nbsp; do Acesso ao Servi�o</td>
		<td><input type="text" class="text" name="txtDtEntrAcesServ" value="<%=strDtEntrAcesServ%>"
			<%if strDtEntrAcesServ <> "" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>
			maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaaa)
		</td>
		<td nowrap>&nbsp;Data Prevista de Entrega<br>&nbsp;do Acesso pelo Provedor</td>
		<td ><input type="text" class="text" name="txtDtPrevEntrAcesProv" value="<%=strDtPrevEntrAcesProv%>"
			maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaaa)
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Data Prevista de<br>&nbsp;&nbsp;&nbsp;&nbsp;Ativa��o do Servi�o</td>
		<td colspan="3"><input type="text" class="text" name="txtDtPrevAtvAcesServ" value="<%=strDtEntrAcesServ%>"
		<%=bbloqueia%>
		maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaaa)</td>
	</tr>
	

	<tr class="clsSilver">
		<td rowspan=2>&nbsp;&nbsp;&nbsp;&nbsp;Acesso Tempor�rio<br>&nbsp;&nbsp;&nbsp;&nbsp;(dd/mm/aaaa)</td>
		<td >&nbsp;In�cio&nbsp;</td>
		<td >&nbsp;Fim&nbsp;</td>
		<td >&nbsp;Devolu��o&nbsp;</td>
	</tr>
	<tr class="clsSilver">
		<td ><input type="text" class="text" name="txtDtIniTemp"  value="<%=strDtIniTemp%>"
		<%=bbloqueia%>  maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;</td>
		<td ><input type="text" class="text" name="txtDtFimTemp" value="<%=strDtFimTemp%>"
		<%=bbloqueia%>  maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;</td>
		<td ><input type="text" class="text" name="txtDtDevolucao" value="<%=strDtDevolucao%>"
		<%=bbloqueia%>  maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Designa��o do Acesso<br>&nbsp;&nbsp;&nbsp; Principal (678)</td>
		<td colspan=3>
			<input type="text" class="text" name="txtDesigAcessoPri0"  maxlength="3" size="3" <%=bbloqueia%> value=678 readOnly>
			<input type="text" class="text" name="txtDesigAcessoPri"  value="<%=dblDesigAcessoPri%>" maxlength="7" size="9" onKeyUp="ValidarTipo(this,0)" >(678N7)
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Observa��es p/ Provedor</td>
		<td colspan="3"><textarea name="txtObs" onkeydown="MaxLength(this,300);" cols="50" rows="3"><%=strObsProvedor%></textarea></td>
	</tr>
	

<% 

	if Trim(Request.Form("hdnAcao")) = "Alteracao" then %>
	
		<tr><th colspan=4>&nbsp;�&nbsp;Informa��es de Altera��o</th></tr>
	
		<tr class="clsSilver">
			<td width=170px>&nbsp;&nbsp;&nbsp;&nbsp;Tipo de Altera��o</td>
			<td >
				<Select name="cboTipoAlteracao" style="width:150px" >
				<option ></option>	&nbsp;
				<option value="Endereco"
				<% if trim(strTipoAlteracao) = "Endereco" Then %>
					 Selected > Endere�o </option>
				<%Else %>
					> Endere�o </option>
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
			<td >&nbsp;&nbsp;&nbsp;Programa��o &nbsp;da &nbsp;&nbsp;&nbsp; Altera��o&nbsp;</td>
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
					Selected > N�o Programada  </option>
				<%Else %>
					> N�o Programada  </option>
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
			&nbsp;�&nbsp;An�lise do estudo de viabilidade (SEV- SSA) &nbsp;
		</th>
	</tr>
	<TR>
		<TD colspan=4 >
			<!--<span id=strProvedorSelSev LANGUAGE=javascript onactivate="return strProvedorSelSev_onactivate()">&nbsp;</span> -->
			<span id=strProvedorSelSev>
				<table cellspacing=1 cellpadding=0 border=0 width=760> 
							<tr class=clsSilver>
								 <td width=321>Endere�o Higienizado:  	 <b> <%=strHigienizado%>   </b></td>
							</TR>
				</table>
				<%
				if dblNroSev <> "" then
					Vetor_Campos(1)="adInteger,4,adParamInput," & dblNroSev
					Vetor_Campos(2)="adInteger,4,adParamOutput,0"
				
					Call APENDA_PARAM("CLA_sp_check_sev",2,Vetor_Campos)
					ObjCmd.Execute'pega dbaction
					DBAction = ObjCmd.Parameters("RET").value
				
					if DBAction = "0" then
						
						'Localiza informa��es da solu��o SSA
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
								  		<td width=321>N� do pedido de Venda do SNOA:  	 <b> <%= TratarAspasJS(Trim(objRSCli("Pedido_Venda_SNOA"))) %>  </b></td> 
							    		<td width=241>Caracterist�ca do EILD :  <b> <%= TratarAspasJS(Trim(objRSCli("C_TIPO_EILD"))) %>  </b></td> 
								</tr> 
			
								<tr class=clsSilver>
										<td width=321>Ponta:  	 <b> <%= TratarAspasJS(Trim(objRSCli("Ponta"))) %>  </b></td> 
										<td width=241>N� SEV Ponta A :  <b><%= TratarAspasJS(Trim(objRSCli("sevPontaA_seq"))) %>  </b></td> 
								</tr> 
			
								<tr class=clsSilver>
										<td width=321>Operadora Solicitante:  	 <b> <%= TratarAspasJS(Trim(objRSCli("OPR_DES"))) %>  </b></td> 
										<td width=321>Pabx Virtual:  	 <b> <%= TratarAspasJS(Trim(objRSCli("pabx_des"))) %>  </b> | Migra��o do Pabx Virtual:  	 <b> <%= TratarAspasJS(Trim(objRSCli("SEV_MIGRA_PABX"))) %>  </td> 
								</tr> 
																
								  		<table cellpadding=0 cellspacing=1 width=760 > 
											<tr> 
												<th nowrap>&nbsp;Tecnologia</th> 
												<th>&nbsp;Provedor</th> 
												<th>&nbsp;Opc Indicada</th>
												<th>&nbsp;Atendimento</th> 
												<th>&nbsp;EILD</th> 
												<th>&nbsp;SIP</th>
												<th>&nbsp;Esta��o</th> 
												<th>&nbsp;Usu�rio</th> 
												<th>&nbsp;Data</th> 
												<th>&nbsp;An�lise</th> 
												<th>&nbsp;OBS</th> 
											</tr> 
									<%
									 While Not objRSCli.eof
									 %>
											<tr class=clsSilver > 
													<td width=15% nowrap >&nbsp; <%=objRSCli("FAC_DES")%>  </td>	
													<td >&nbsp;  <%=objRSCli("FOR_DES")%> </td> 
													<td >&nbsp;<%= TratarAspasJS(Trim(objRSCli("SOL_SELECIONADA"))) %></td> 
													<td >&nbsp;  <%=objRSCli("ATE_DESC")%> </td> 
													<td >&nbsp;  <%=objRSCli("TIPO_EILD")%> </td> 
													<td >&nbsp;  <%=objRSCli("SOL_SINALIZACAO_SIP")%> </td> 
													<td >&nbsp;  <%=objRSCli("ESTACAO")%> </td> 
													<td nowrap >&nbsp;  <%=objRSCli("SEV_USRNAMESOLUC")%>  </td> 
													<td nowrap >&nbsp; <%=objRSCli("SEV_DTULTRESPSOLUC")%> </td> 
													<td >&nbsp;  <%=objRSCli("ANALISE")%></td> 
													<td width=20% >&nbsp;  <%=objRSCli("SOL_OBS")%>  </td> 
												</tr> 
									<%
									objRSCli.MoveNext
									Wend
							End If
						 End If
					End If
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
<!--JCARTUS@ - Acessos F�sicos sem CNPJ/IE/IM vinculado na base do CLA (CH-69613HSK)-->
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
<input type=hidden name=hdnCboProvedor value="cboProvedor.value">

<input type=hidden name=hdnstrAcessoTipoRede value="<%=strAcessoTipoRede%>">
<input type=hidden name=hdnSegmento    value="txtSegmento.value">
<input type=hidden name=hdnPorte      value="txtPorte.value">

<input type=hidden name=hdnCNLSiglaCentroCliDest      value="txtCNLSiglaCentroCliDest.value">
<input type=hidden name=hdnComplSiglaCentroCliDest    value="txtComplSiglaCentroCliDest.value">
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
<input type=hidden name=hdnTecnologia>
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
		Verificar Procedimento de recupera��o de f�sico utilizado para Altera��o. 	
	-->

		<%'if ( (Request.Form("hdnAcao") = "Alteracao") ) then %>
		<!--			
			<tr><th colspan=4>&nbsp;�&nbsp;Ponto de Instala��o do Processo Ativo Anterior</th></tr>
			<tr><td colspan=4>
					<table border=0 width=758 cellspacing=1 cellpadding=0>
						<tr>
							<th  width=15>&nbsp;</th>
							<th  width=35>&nbsp;Editar</th>
							<th  width=50>&nbsp;Prop Fis</th>
							<th  width=185>&nbsp;Provedor</th>
							<th  width=200>&nbsp;Velocidade</th>
							<th	 width=273>&nbsp;Endere�o</th>
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
			
			<tr><th colspan=4>&nbsp;�&nbsp;Acessos F�sicos Utilizados</th></tr>
			<tr><td colspan=4>
					<table border=0 width=758 cellspacing=1 cellpadding=0>
						<tr>
							<th  width=15>&nbsp;</th>
							<th  width=35>&nbsp;Editar</th>
							<th  width=50>&nbsp;Prop Fis</th>
							<th  width=185>&nbsp;Provedor</th>
							<th  width=200>&nbsp;Velocidade</th>
							<th	 width=273>&nbsp;Endere�o</th>
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
		
			
			
			<tr><th colspan=4>&nbsp;�&nbsp;Acessos F�sicos Utilizados</th></tr>
			<tr><td colspan=4>
					<table border=0 width=758 cellspacing=1 cellpadding=0>
						<tr>
							<th  width=15>&nbsp;</th>
							<th  width=35>&nbsp;Editar</th>
							<th  width=50>&nbsp;Prop Fis</th>
							<th  width=185>&nbsp;Provedor</th>
							<th  width=200>&nbsp;Velocidade</th>
							<th	 width=273>&nbsp;Endere�o</th>
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
		<th colSpan=4>&nbsp;�&nbsp;Informa��es do Acesso&nbsp;</th>
	</tr>
	<tr class="clsSilver">
			<td width=170px ><font class="clsObrig">:: </font>Prop do Acesso F�sico</td>
			<td nowrap >			

				<input type=radio name=rdoPropAcessoFisico value="TER" <%if strAcessoTipoRede = "10" then%>disabled checked<%end if%> Index=0 onclick="javascript:EsconderTecnologia(0);ResgatarTecVel();EsconderFabONT();EsconderTipoONT();MostrarVlan();">Terceiro&nbsp;&nbsp;&nbsp;
				<input type=radio name=rdoPropAcessoFisico value="EBT" <%if strAcessoTipoRede = "10" then%>disabled<%end if%> <%if strAcessoTipoRede = "4" then%>disabled<%end if%> Index=1 onclick="EsconderTecnologia(0);ResgatarTecVel();EsconderFabONT();EsconderTipoONT();MostrarVlan();">EBT&nbsp;&nbsp;&nbsp;
				
				<!--Alterado por Fabio Pinho em 29/04/2016 - ver 1.0 - Inicio-->
				<input type=radio name=rdoPropAcessoFisico  value="CLI" disabled Index=2 onclick="EsconderTecnologia(0);ResgatarTecVel();EsconderFabONT();EsconderTipoONT();MostrarVlan();">Cliente&nbsp;&nbsp;&nbsp;
				<!--
				<input type=radio name=rdoPropAcessoFisico  value="CLI" <%if strAcessoTipoRede = "10" then%>disabled<%end if%>  <%if strOriSol = "9" then%>disabled<%end if%> <%if strAcessoTipoRede = "4" then%>disabled<%end if%> Index=2 onclick="EsconderTecnologia(0);ResgatarTecVel();EsconderFabONT();EsconderTipoONT();MostrarVlan();">Cliente&nbsp;&nbsp;&nbsp;
				-->
				<!--Alterado por Fabio Pinho em 29/04/2016 - ver 1.0 - Fim-->

<!--
				<input type=radio name=rdoPropAcessoFisico value="TER"  <%if strAcessoTipoRede = "4" then%>checked<%end if%> Index=0	<%if strPropAcessoFisico = "TER" then Response.Write " checked " End if%> onclick="EsconderTecnologia(0);ResgatarTecVel();EsconderFabONT();EsconderTipoONT();MostrarVlan();">Terceiro&nbsp;&nbsp;&nbsp;
				<input type=radio name=rdoPropAcessoFisico value="EBT" <%if strAcessoTipoRede = "4" then%>disabled<%end if%> Index=1	<%if strPropAcessoFisico = "EBT" then Response.Write " checked " End if%> onclick="EsconderTecnologia(0);ResgatarTecVel();EsconderFabONT();EsconderTipoONT();MostrarVlan();">EBT&nbsp;&nbsp;&nbsp;
				<input type=radio name=rdoPropAcessoFisico value="CLI" <%if strOriSol = "9" then%>disabled<%end if%> <%if strAcessoTipoRede = "4" then%>disabled<%end if%> Index=2	<%if strPropAcessoFisico = "CLI" then Response.Write " checked " End if%> onclick="EsconderTecnologia(0);ResgatarTecVel();EsconderFabONT();EsconderTipoONT();MostrarVlan();">Cliente&nbsp;&nbsp;&nbsp;
-->			
			<td nowrap colspan=2>
				<%if Trim(Request.Form("hdnAcao")) = "Alteracao" and TipoVel(dblTecId) <> "" then%>
					<div id=divTecnologia style="display:'';POSITION:relative">
				<%Else%>
					<div id=divTecnologia style="display:none;POSITION:relative">
				<%End if%>
				<Select name=cboTecnologia onChange="RetornaCboTipoRadio(this[this.selectedIndex].innerText,this.value,'<% = strTrdID %>', '<% = strVersao %>');ResgatarTecVel();EsconderFabONT();EsconderTipoONT();ResgatarFabONT();" >
					<Option value="">:: TECNOLOGIA EBT</Option>
					<%
					set objRS = db.execute("CLA_sp_sel_tecnologia 0 , " & strOriSol)
					While not objRS.Eof
						strItemSel = ""
						if Trim(dblTecId) = Trim(objRS("Tec_id")) then strItemSel = " Selected " End if
						Response.Write "<Option value=" & objRS("Tec_id") & strItemSel & ">" & objRS("Tec_Nome") & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
					%>
					</Select>
				</div>
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
		<td width=170px><font class="clsObrig">:: </font>Vel do Acesso F�sico</td>
		<td colspan=3><span id=spnVelAcessoFis>
			<select name="cboVelAcesso" style="width:150px" onChange="MostrarTipoVel(this)">
				<option ></option>
				<%
					if Trim(dblTecId) <> "" then
						Set objRS = db.execute("CLA_sp_sel_AssocTecVeloc null," & dblTecId)
					Else
						set objRS = db.execute("CLA_sp_sel_velocidade")
					End if
					While Not objRS.eof
						strItemSel = ""
						if Trim(strVelAcesso) = Trim(objRS("Vel_ID")) then strItemSel = " Selected " End if
						Response.Write "<Option value='" & Trim(objRS("Vel_ID")) & "'" & strItemSel & ">" & objRS("Vel_Desc") & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				%>
			</select></span>&nbsp;&nbsp;<font class="clsObrig">:: </font>Qtde de Acesso(s) Fisico(s)&nbsp;<input type="text" class="text" name="txtQtdeCircuitos" value=1  maxlength="2" size="2" onKeyUp="ValidarTipo(this,0)" value="<%=dblQtdeCircuitos%>" readonly>&nbsp;&nbsp;
			<div id=divTipoVel style="display:none;POSITION:absolute">
			<select name="cboTipoVel" style="width:170px">
				<option value="">TIPO DE VELOCIDADE</option>
				<option value="1" <%if strTipoVel=1 then Response.Write " Selected " %>>ESTRUTURADA</option>
				<option value="0" <%if strTipoVel=0 then Response.Write " Selected " %>>N�O ESTRUTURADA</option>
			</select>
			</div>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px>&nbsp;&nbsp; Senha</td>
		<td colspan="3">
						
				<input type="text" class="text" name="txtSenha" maxlength="10"	size="11"></td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Provedor</td>
		<td colspan="1">
			<select name="cboProvedor" onChange="ResgatarPromocaoRegime(this);MostrarVlan(this)" <%if ( strAcessoTipoRede = "4" or strAcessoTipoRede = "10" ) then%>disabled<%end if%>>
				<option value=""></option>
				<%	
					'response.write "<script>alert('"&strAcessoTipoRede&"')</script>"
					if isnull(strAcessoTipoRede) then
						strAcessoTipoRede = "null"
						'response.write "<script>alert('"&strAcessoTipoRede&"')</script>"
					end if 
					set objRS = db.execute("CLA_sp_sel_provedor 0,null,1, " & strAcessoTipoRede & ", " & strOriSol )
					While not objRS.Eof
						strItemSel = ""
						if Trim(dblProId) = Trim(objRS("Pro_ID")) then strItemSel = " Selected " End if
						if  ( ( strAcessoTipoRede = "4" and Trim(objRS("Pro_ID")) = 136 ) or ( strAcessoTipoRede = "10" and Trim(objRS("Pro_ID")) = 143 ) )then 
						strItemSel = " Selected " 
						End if
						Response.Write "<Option value='" & Trim(objRS("Pro_ID")) & "'" & strItemSel & ">" & objRS("Pro_Nome") & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				%>
			</select>
		</td>
		<td width=170px>
			<div id=divVLAN_1 style="display:none;POSITION:absolute">
				&nbsp;&nbsp;&nbsp;&nbsp;<font class="clsObrig">:: </font>VLAN</td>
			</div>
		<td>
			<div id=divVLAN_2 style="display:none;POSITION:absolute">
				<input type="text" class="text" name="txtVLAN"  maxlength="18" size="20" onKeyUp="ValidarTipo(this,0)" value="<%=dblVLAN%>">
			</div>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px>&nbsp;&nbsp;&nbsp;&nbsp;Prazos de Contrata��o<br>&nbsp;&nbsp;&nbsp;&nbsp;de Acesso</td>
		<td colspan=3>
			<span id=spnRegimeCntr>
				<select name="cboRegimeCntr" >
				<option value=""></option>
					<%
						if Trim(dblProId) <> "" then
							set objRS = db.execute("CLA_sp_sel_regimecontrato 0," & dblProId)
							While not objRS.Eof
								strItemSel = ""
								if Trim(dblRegId) = Trim(objRS("Reg_ID")) then strItemSel = " Selected " End if
								Response.Write "<Option value='" & Trim(objRS("Reg_ID")) & "'" & strItemSel & ">" & LimparStr(Trim(objRS("Pro_Nome"))) & " - " & LimparStr(Trim(objRS("Tct_Desc"))) & "</Option>"
								objRS.MoveNext
							Wend
							strItemSel = ""
						End if
					%>
				</select>
			</span>
		</td>
	<tr class="clsSilver">
		<td width=170px>&nbsp;&nbsp;&nbsp;Promo��o</td>
		<td colspan=3 >
			<span id=spnPromocao>
			<select name="cboPromocao" style="width:170px">
				<option value=""></option>
				<%
					if Trim(dblProId) <> "" then
						set objRS = db.execute("CLA_sp_sel_promocaoprovedor 0," & dblProId)
						While not objRS.Eof
							strItemSel = ""
							if Trim(dblPrmId) = Trim(objRS("Prm_ID")) then strItemSel = " Selected " End if
							Response.Write "<Option value='" & Trim(objRS("Prm_ID")) & "'" & strItemSel & ">" & objRS("Prm_Desc") & "</Option>"
							objRS.MoveNext
						Wend
						strItemSel = ""
					End if
				%>
			</select>
			</span>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px>&nbsp;&nbsp;&nbsp;&nbsp;C�digo SAP</td>
		<td >
			<input type="text" class="text" name="txtCodSAP"  <%if strOrigem="APG" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%> maxlength="7" size="10" onKeyUp="ValidarTipo(this,0)" value="<%=strCodSap%>" >&nbsp;(N7)
		</td>
		<td >&nbsp;&nbsp;&nbsp;N�mero PI&nbsp;</td>
		<td>
			<input type="text" class="text" name="txtNroPI"  maxlength="7" size="10" onKeyUp="ValidarTipo(this,0)" value="<%=dblNroPI%>" >&nbsp;(N7)
		</td>
	</tr>
	<tr class=clsSilver2>
		<td width=170px >&nbsp;Endere�o Origem&nbsp;</td>
		<td nowrap colspan=3>
			<font class=clsObrig>:: </font>PONTO&nbsp;
				<select name="cboTipoPonto" onChange="TipoOrigem(this.value)">
					<option value=""></option>
					<option value="I" <%if Trim(strTipoPonto) = "I" then Response.Write " selected " %>>CLIENTE</option>
					<option value="T" <%if Trim(strTipoPonto) = "T" then Response.Write " selected " %>>INTERMEDI�RIO</option>
				</select>			
		</td>
	</tr>
	
	<%'Incluir tratamento para Bloquear apenas os dados do Endere�o do Cliente %>
	<tr class="clsSilver">
		<td width=170px nowrap><span id=spnOrigem>&nbsp;&nbsp;&nbsp;Sigla Esta��o Origem(CNL)</span></td>
		<td colspan=3>
			<input type="text" class="text" name="txtCNLSiglaCentroCli"  maxlength="4"
			<%if strCidSigla <> "" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>
			size="8" TIPO="A" value="<%=strCidSigla%>">

			&nbsp;Complemento
			<input type="text" class="text" name="txtComplSiglaCentroCli"  maxlength="3"
			<%if strEscSigla <> "" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>
			size="10" TIPO="A" value="<%=strEscSigla%>">&nbsp;
						
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>UF</td>
		<td>
                        <%if strOriSol = "10" then%>
			     <input type="text" class="text" name="cboUFEnd"  maxlength="2"
			    size=3" readonly="TRUE" value="<%=strUFEnd%>">
                       
                        <% else %>
                               <select name="cboUFEnd">
			            <Option value=""></Option>
			            <%
			                 set objRS = db.execute("CLA_sp_sel_estado ''")
			                 While not objRS.Eof
				             strItemSel = ""
				              if Trim(strUFEnd) = Trim(objRS("Est_Sigla")) then strItemSel = " Selected " End if
				              Response.Write "<Option value=" & objRS("Est_Sigla") & strItemSel & ">" & objRS("Est_Sigla") & "</Option>"
				              objRS.MoveNext
			                  Wend
			                  strItemSel = ""
		                   %>
			       </select>

                        <% end if %>
		
		</td>


			<!--Alterado por Fabio Pinho em 28/04/2016 - ver 1.0 - Inicio-->
			<%if Trim(strOrderEntrySis) = "CFD" then%>

				<td nowrap right><font class="clsObrig">:: </font>CEP&nbsp;</td>
				<td> <input type="text" class="text" name="txtCepEnd" readOnly="readOnly" value="<%=strCepEnd%>" maxlength="9" size="12">

			<%else%>

			<td nowrap right><font class="clsObrig">:: </font>CEP&nbsp;</td>
				<td> <input type="text" class="text" name="txtCepEnd" value="<%=replace(strCepEnd,"-","")%>" maxlength="8" size="12" onkeypress=OnlyNumbers()>

			<%End if%>
			<!--Alterado por Fabio Pinho em 28/04/2016 - ver 1.0 - Fim-->

		
		
		<%if strOriSol <> "10" then%>
			<input type=button name=btnProcurarCepInstala value="Procurar CEP"
				
				class="button" onclick="ProcurarCEP(1,1)" tabindex=-1 onmouseover="showtip(this,event,'Procurar por CEP exato ou pelos 5 primeiros d�gitos (Alt+D)');" accesskey="D">
			</td>
		 <% else %>
		 	<td nowrap right></td>
			<td> </td>
		 
		<%end if%>
			
	</tr>

	<tr class="clsSilver">
		<td nowrap><font class="clsObrig">:: </font>Munic�pio (CNL - ANATEL)</td>
		<td nowrap colspan=3 >
                     <%if strOriSol = "10" then%>
			<input type=text size=5 maxlength=4 class=text name="txtEndCid" readonly="TRUE" value="<%=strEndCid%>">&nbsp;

			<input type=text size=70 readonly class=text readonly="TRUE" name="txtEndCidDesc" value="<%=strEndCidDesc%>">
		     
                     <% else %>

                         <input type=text size=5 maxlength=4 class=text name="txtEndCid" value="<%=strEndCid%>"

			onBlur="if (ValidarTipo(this,1)){ResgatarCidade(document.forms[1].cboUFEnd,1,this)}">&nbsp;

			<input type=text size=27 readonly style="BACKGROUND-COLOR:#eeeeee" class=text
			<%=bbloqueia%>
			name="txtEndCidDesc" value="<%=strEndCidDesc%>" tabIndex=-1>
		<% end if %>

                </td>
	</tr>
	
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Tipo do Logradouro</td>
		<td colspan=3>
                    <%if strOriSol = "10" then%>
			<input type="text" class="text" name="cboLogrEnd"  maxlength="15" readonly="TRUE"
			size="15"  value="<%=strLogrEnd%>">
                    
                    <% else %>
                         <select name="cboLogrEnd">
				<option value=""></option>
				<% set objRS = db.execute("CLA_sp_sel_tplogradouro")
					While not objRS.Eof
						strItemSel = ""
						if Trim(strLogrEnd) = Trim(objRS("Tpl_Sigla")) then strItemSel = " Selected " End if
						Response.Write "<Option value=""" & Trim(objRS("Tpl_Sigla")) &""" " & strItemSel & ">" & Trim(objRS("Tpl_Sigla")) & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				%>
			</select>

                    <% end if %>
			
		</td>
		
	</tr>
	<tr class="clsSilver">
		<td><font class="clsObrig">:: </font>Nome Logr</td>
		<td nowrap>
			<input type="text" class="text" name="txtEnd" <%if strOriSol = "10" then%> readonly="TRUE" <% end if%>
			
			value="<%=strEnd%>" maxlength="60" size="40">
		</td>
		<td width=170px><font class="clsObrig">:: </font> N�mero</td>
		<td>
			<input type="text" class="text" name="txtNroEnd" value="<%=strNroEnd%>" <%if strOriSol = "10" then%> readonly="TRUE" <% end if%>
			 maxlength="10" size="10">
		</td>
	</tr>
	
	<tr class="clsSilver">
		<td>&nbsp;&nbsp;&nbsp;Complemento</td>
		<td colspan= 3 >
			<input type="text" class="text" name="txtComplEnd"  value="<%=strComplEnd%>" <%if strOriSol = "10" then%> readonly="TRUE" <% end if%> maxlength="80" size="80" >
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Bairro</td>
		<td colspan = 3 >
			<input type="text" class="text" name="txtBairroEnd"
			 
			value="<%=strBairroEnd%>" maxlength="120" size="40" <%if strOriSol = "10" then%> readonly="TRUE" <% end if%> >&nbsp;
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
			<input type="text" class="text" name="txtTelEndArea" value="<%=strTelEndArea%>"	maxlength="2" size="2" onkeyUp="ValidarTipo(this,0)">&nbsp;

			<input type="text" class="text" name="txtTelEnd" value="<%=strTelEnd%>"	maxlength="9" size="11" onkeyUp="ValidarTipo(this,0)">
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>CNPJ</td>
		<td >
			<input type="text" class="text" name="txtCNPJ"  maxlength="14"
			<%=bbloqueia%>
			size="16" onKeyUp="ValidarTipo(this,0)" value="<%=dblCNPJ%>" >&nbsp;(99999999999999)
		</td>
		<td width=170px>&nbsp;&nbsp;&nbsp;</font>Email Contato T�cnico</td>
	    <td>
	      <input type="text" class="text" name="txtEmailTec"  value="<%=strBairroEnd%>" maxlength="30" size="30">&nbsp;
	    </td>
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
		<td width=170px nowrap>&nbsp;&nbsp;&nbsp;Propriet�rio do Endere�o</td>
		
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
		<td width=170px><font class=clsObrig>:: </font><font color="#FF0000">Interface</font></td>
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
		<td width=170px><span id=spnDestino>&nbsp;&nbsp;&nbsp;Sigla Esta��o Entrega<br>&nbsp;&nbsp;&nbsp;do Acesso F�sico</span></td>
		<td colspan=3 nowrap>
			<table border=0 cellspacing=0 cellpadding=0>
				<tr>
					<td>&nbsp;CNL</td>
					<td>&nbsp;Complemento</td>
					<td>&nbsp;Endere�o de Entrega do Acesso F�sico</td>
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
		<td width=170px><font class=clsObrig>:: </font><font color="#FF0000">Interface</font></td>
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
			<div id=divIDFis1 style="DISPLAY: 'none'">
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
			<% if strOriSol = "10" then %> 
				<input type=button name=btnIDFis1 class=button value="Acesso F�sico"  onClick="javascript:window.open('ProcurarAcessoFisico.asp?FlagOrigem=CLA&txtNroSev='+document.Form1.txtNroSev.value,'janela','toolbar=no,statusbar=no,resizable=yes,scrollbars=YES,width=900,height=800,top=100,left=100')" onmouseover="showtip(this,event,'Procurar um id f�sico pr� existente para o endere�o atual (Alt+F)');" accesskey="F" <%if strMudFisico="N" or dblOriSolID = 9 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>  <%=bbloqueiaTipoRede%> >&nbsp; 
			<% else %>
				<input type=button name=btnIDFis1 class=button value="Procurar Id F�sico" onClick="ProcurarIDFis(1)" onmouseover="showtip(this,event,'Procurar um id f�sico pr� existente para o endere�o atual (Alt+F)');" accesskey="F" <%if strMudFisico="N" or dblOriSolID = 9 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>  <%=bbloqueiaTipoRede%> >&nbsp;
			<% end if %>			
			<span id="spnBtnLimparIdFis1"></span>&nbsp;
			<input type=button name=btnLimparAcesso class=button value="Limpar" onClick="LimparInfoAcesso()" accesskey="L" onmouseover="showtip(this,event,'Limpar dados do Acesso (Alt+L)');" <%if strMudFisico="N" then%> <%=readOnlyASMS%> <%else%> <%=bdesbloqueia%> <% End if%>  >&nbsp;
		</td>
	</tr>
	<tr class="clsSilver">
		<td colspan="4">
			<div id=divTronco2M style="DISPLAY: 'none'">
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
<table border=0 cellspacing="1" cellpadding="0" width="760" >
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
			<input type=hidden name=hdnNecessitaRecurso value="S"> <!-- Na ativa��o ser� sempre SIM -->
			<input type=hidden name=hdnEmiteOTS>
			<input type="hidden" name="hdnOriSol_ID" value="<%=objRSSolic("OriSol_ID")%>">
			<input type="hidden" name="hdnAprovisi_ID" value="<%=dblAprovisiId%>">
			<input type=hidden name=hdnId_Acesso value="<%=id_acesso%>">
			<input type=hidden name=hdnStrAcfId value="<%=strAcf_id%>" >
			<input type=hidden name=hdnReaproveitarFisico value="N"> <!-- Na ativa��o ser� sempre N�o. Ser� modificado na Altera��o -->
			<%if objRSSolic("OriSol_ID") <> 9 then%>
			<tr>
					<th colspan="4" >&nbsp;�&nbsp;Informa��es do Aprovisionador</th>
			</tr>
		  
			</tr>
				<tr class="clsSilver">
				<td nowrap width=170px class="clsSilver">&nbsp;&nbsp;&nbsp;&nbsp;Username <br>&nbsp;&nbsp;&nbsp;&nbsp;do Cadastrador</td>
				<td nowrap class="clsSilver"><input type="text" class="text" name="txtUsernamecadastrador"
				<%=bbloqueia%>
				value="<%=strUsernamecadastrador%>" maxlength="20" size="20" onKeyPress="">&nbsp;</td>
				<td nowrap class="clsSilver">&nbsp;Telefone Cadastrador<br>&nbsp;</td>
				<td nowrap class="clsSilver"><input type="text" class="text" name="txtTelefoneCadastrador"
				<%=bbloqueia%>
				value="<%=strTelefoneCadastrador%>" maxlength="10" size="10" onKeyPress="">&nbsp;</td>
			</tr>
			<tr class="clsSilver">
				<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Observa��o</td>
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

<table border=0 cellspacing="1" cellpadding="0" width="760" >
	<tr>
		<th colspan="4" >&nbsp;�&nbsp;Informa��es da Embratel</th>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Local de Entrega<br>&nbsp;&nbsp;&nbsp;do Acesso L�gico</td>
			<input type="Hidden" name="cboLocalEntrega">
		<td><input type="text" class="text" name="txtCNLLocalEntrega"  maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A">&nbsp;
		&nbsp;<input type="text" class="text" name="txtComplLocalEntrega"  maxlength="3" size="6" onKeyUp="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoUsu(document.Form3.txtCNLLocalEntrega,document.Form3.txtComplLocalEntrega,<%=dblUsuId%>,1);" TIPO="A">
		</td>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px" nowrap><font class="clsObrig">:: </font>Local de Configura��o</td>
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
		<th colspan=7 >&nbsp;�&nbsp;Coordena��o Embratel</th>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>�rg�o de Venda</td>
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
	<tr>
		<th colspan="2">&nbsp;�&nbsp;Hist�rico</th>
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
			
			<!--CH-41083NDL - inicio-->

				<!--
				<input type="button" class="button"  style="color:darkred;;font-weight:bold;width:180px" name="btnGravar" value=".::Gravar::." onclick="Gravar()" accesskey="I" onmouseover="showtip(this,event,'Gravar uma solicita��o (Alt+I)');" <%if ( (( trim(dblIdLogico) = "" or isnull(dblIdLogico) ) and stracao = "ALT") or desativar_gravar=true) then%>Disabled<%end if%> <%if strbloqcan = true then%>Disabled<%end if%>>&nbsp;
				-->

			<%if bbloqueiaSev = "false" then%>
				<input type="button" class="button"  style="color:darkred;;font-weight:bold;width:180px" name="btnGravar" value=".::Gravar::." onclick="Gravar()" accesskey="I" onmouseover="showtip(this,event,'Gravar uma solicita��o (Alt+I)');" <%if ( (( trim(dblIdLogico) = "" or isnull(dblIdLogico) ) and stracao = "ALT") or desativar_gravar=true) then%>Disabled<%end if%> <%if strbloqcan = true then%>Disabled<%end if%>>&nbsp;
			<%End if%>
			<!--CH-41083NDL - fim-->

			<!--<input type="button" class="button" name="btnNovaSolic" value="Nova Solicita��o" style="width:100px" onclick="javascript:window.location.replace('solicitacao.asp')" accesskey="N" onmouseover="showtip(this,event,'Nova solicita��o (Alt+N)');">&nbsp;-->
			<%if Trim(Request.Form("hdnAcao")) = "Alteracao" then%>
			<input type=button	class="button" name=btnVoltar value=Voltar onclick="VoltarOrigem()" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">&nbsp;
			<%End if%>
			<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
		</td>
	</tr>
	<tr>
		<td>
			<font class="clsObrig">:: </font> Campos de preenchimento obrigat�rio.
		</td>
	</tr>
	<tr>
		<td>
			<font class="clsObrig">:: </font>Legenda: A - Alfanum�rico;  N - Num�rico;  L - Letra
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
MostrarVlanProvedor()
<%
Set objRS = db.execute("CLA_sp_sel_usuario 0,'" & Trim(strUserName) & "'")
if Not 	objRS.Eof And Not objRS.Bof then
	Response.Write "spnNomeGICL.innerHTML = '" & Replace(Trim(objRS("Usu_Nome")),"'","�")  & "';"
	Response.Write "spnRamalGICL.innerHTML = '" & Replace(Trim(Cstr("" & objRS("Usu_Ramal"))),"'","�")  & "';"
	Response.Write "document.forms[2].hdntxtGICL.value = '" & Trim(strUserName) & "';"
End if
%>
//Geral
with (document.forms[0])
{
	<%if Trim(Request.Form("hdnAcao")) <> "Alteracao" or Request.Form("hdnOEOrigem")="APG"then %>
		txtOrderEntry[0].readOnly = true
		txtOrderEntry[1].readOnly = true
		txtOrderEntry[2].readOnly = true
	<%Else%>
		
		ResgatarDesigServicoGravado(<%=dblSerId%>)
	<%End if%>
}
</script>
<!--Form que envia os dados para grava��o-->
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
<input type=hidden name=hdnXml>
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
'*********** MIGRA��O HFC BSoD **********
 
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
			  'Solicita��es oriundas do SGA-Plus com marca��o para BSOD NET dever�o ser liberadas.
				If Not (strAcessoTipoRede = 4 and objMigr("Pro_ID") = 136) Then
					response.write "<script>alert('ATEN��O \nEst�o bloqueadas novas altera��es para provedores HFC BSoD VIA e HFC BSoD NET. \nAssim, � necess�rio realizar a migra��o do acesso f�sico para HFC BSoD, atrav�s da funcionalidade de Migra��o, que ser� exibida ap�s o OK desta mensagem. \nEm caso de d�vidas, contatar a equipe do SOE.')</script>"

'			    Set objRSMigr_Compart = db.execute("select count(acfl.acl_idacessologico) As Qtd_ACL from cla_acessofisico acf inner join cla_acessologicofisico acfl on acf.acf_id=acfl.acf_id " &_
'			                                       "where alf_flag is null and acf.acf_idacessofisico='" & objMigr("Acf_IDAcessoFisico") & "'")			
'			    If CInt(objRSMigr_Compart("Qtd_ACL"))>1 Then					
'				     response.write "<script>alert('Migra��o n�o permitida para Acesso F�sico com compartilhamento.\n\nPara migrar o Acesso F�sico, voc� dever� desativar ou remanejar os Acessos L�gicos vinculados, deixando apenas um Acesso L�gico vinculado ao Acesso F�sico [" & objMigr("Acf_IDAcessoFisico") & "].')</script>" 
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
