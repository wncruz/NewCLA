<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: AlteracaoCadAprov.ASP
'	- Descrição			: Cadastra/Altera uma solicitação no sistema CLA
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<!--#include file="../inc/xmlAcessos.asp"-->
<%
Dim strVisada			'Tipo Visada
Dim strGrupo			'Grupo Cliente
Dim strOriSol			'Origem Solicitacao
Dim strProjEspecial	' Projeto Especial
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
dim Acao
Dim bdesbloqueia			'Veriavel de controle para para Desbloquear campos
Dim dblSolAPGId				'IDentificador da Solicitação APG
Dim strcomp_troncoInicio
Dim	strcomp_tronco2m
Dim	strcomp_rota
Dim	strcomp_troncoInterface
Dim	strcomp_troncoFim
Dim	strcomp_rotaCNG
Dim	strcomp_contrato
Dim	strcomp_codServico
Dim	strcomp_designacao
Dim	strcomp_obs
Dim bbloqueia_GICN
'Inicializacao das variáveis:
bbloqueia = "disabled=true"

strOrigem = Server.HTMLEncode(Request.Form("hdnOEOrigem"))
dblAprovisiId = Trim(Server.HTMLEncode(Request.Form("hdnAprovisiId")))

'response.write "<b>Aprovisionador Origem: </b>"&strOrigem& "<br>"
'response.write "<b>AprovisiId: </b>"&dblAprovisiId& "<br>"

intAno = Year(Now)
strUserNameGICLAtual = Trim(strUserName)

dblSolId = Trim(Server.HTMLEncode(Request.Form("hdnSolId")))
if dblSolId = "" then 
	dblSolId = Trim(Server.HTMLEncode(Request.QueryString("SolId"))) 
end if

if dblSolId = "" then
	Response.Write "<script language=javascript>alert('Solicitação indisponível.');window.location.replace('main.asp');</script>"
	Response.End
End if

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

acao = Server.HTMLEncode(Request.Form("acao"))
if acao = "" then
  acao = Server.HTMLEncode(request.querystring("acao"))
  select case acao
    case 1
	  acao = "ATV"
	  intTipoProcesso = 1
	Case 2
	  acao = "DES"
	  intTipoProcesso = 2
	Case 3
	  acao = "ALT"
	  intTipoProcesso = 3
	Case 4
	  acao = "CAN"
	  intTipoProcesso = 4
  end select
end if

Set objRSAprov = db.execute("CLA_sp_sel_Aprovisionador " & dblAprovisiId)
If Not objRSAprov.eof or Not objRSAprov.bof Then

	if not isnull(objRSAprov("Acl_IDAcessoLogico")) then
		strIDLogico 			= Trim(Cstr(objRSAprov("Acl_IDAcessoLogico")))
	else
		strIDLogico 			= ""
	end if
	strIDSol 				= Trim(objRSAprov("Sol_ID"))
	strOriSol 				= Trim(objRSAprov("Orisol_ID"))
	strOriDesc 				= Trim(objRSAprov("Orisol_Descricao"))
	
	strOrderEntryAno		= Trim(objRSAprov("OE_Ano"))
	strOrderEntryNro		= Trim(objRSAprov("OE_Numero"))
	strOrderEntryItem		= Trim(objRSAprov("OE_Item"))
	
	strDesignacaoServico    = Trim(objRSAprov("Acl_DesignacaoServico"))
	strTipoContratoServico	= Trim(objRSAprov("acl_tipoContratoServico"))
	strNroContrServico		= Trim(objRSAprov("Acl_NContratoServico"))
	
  '  strProcesso 			= objRSAprov("Processo")
    strAcao 				= objRSAprov("Acao")
	IdAcesso				= objRSAprov("id_acesso")
	
 	strPov_Razao = Ucase(Trim(objRSAprov("Cli_Nome")))
 	strPov_CC = Trim(objRSAprov("Cli_CC"))
	strPov_SubCC = Trim(objRSAprov("Cli_SubCC"))
	strPov_ContratoServ = Trim(objRSAprov("Acl_TipoContratoServico"))
	strPov_ContatoCli = Ucase(Trim(objRSAprov("Aec_Contato")))
	strPov_Tel = Trim(objRSAprov("Aec_Telefone"))
	strPov_CNPJ = Trim(objRSAprov("Aec_CNPJ"))
	strPov_PropEnd = Ucase(Trim(objRSAprov("Aec_PropEnd")))
	strPov_End = Ucase(Trim(objRSAprov("End_NomeLogr")))
	strPov_VelServ	= Ucase(Trim(objRSAprov("Vel_Desc")))
	strPov_DesigServ = Trim(objRSAprov("Acl_DesignacaoServico"))
	strPov_SerDesc = Ucase(Trim(objRSAprov("Ser_Desc")))
	
	Set objPov2 = db.execute("select Ser_ID from cla_servico where ser_desc = '" & Trim(strPov_SerDesc) & "'")
	If Not objPov2.eof or Not objPov2.bof Then
		strPov_SerID	= Trim(objPov2("Ser_ID"))
		dblSerId = strPov_SerID
	End If
	
	Set objPov3 = db.execute("select vel_id from cla_velocidade where vel_desc = '" & strPov_VelServ & "'")
	If Not objPov3.eof or Not objPov3.bof Then
		strPov_VelID	= Trim(objPov3("Vel_ID"))
	End If
	
	dblSolId = strIDSol
	strOrderEntrySis = strOriDesc
End If

if dblSolId <> "" and dblSolId <> "0" then
	Set objRSSolic		= db.execute("CLA_sp_view_solicitacaomin " & dblSolId)
	If Not objRSSolic.eof or Not objRSSolic.bof Then
		dblIdLogico			= Trim(objRSSolic("Acl_IDAcessoLogico"))
		intTipoProcesso		= objRSSolic("Tprc_ID")
		strAntAcesso			= objRSSolic("Acl_AntAcesso")
		
		'Xml com os pontos
		Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
		objXmlDados.loadXml("<xDados/>")
		
		Vetor_Campos(1)="adInteger,4,adParamInput,"
		Vetor_Campos(2)="adInteger,4,adParamInput,"
		Vetor_Campos(3)="adInteger,4,adParamInput," & dblIdLogico
		Vetor_Campos(4)="adInteger,4,adParamInput," & dblSolId
		
		'''IF request("libera") = 1 THEN ' LIBERACAO DE ESTOQUE
			Vetor_Campos(5)="adVarchar,1,adParamInput,NULL"  
			Vetor_Campos(6)="adInteger,1,adParamInput,0"
		'''ELSE
		''	Vetor_Campos(5)="adVarchar,1,adParamInput,A"  
		''	Vetor_Campos(6)="adInteger,1,adParamInput,NULL"
		'''END IF
		
	  	strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_VIEW_PONTO",6,Vetor_Campos)

		Set objRSFis = db.Execute(strSqlRet)
		Set objDicProp = Server.CreateObject("Scripting.Dictionary")
		
		if Not objRSFis.EOF and not objRSFis.BOF then
			Set objXmlDados = MontarXmlAcesso(objXmlDados,objRSFis,"")
			strXmlAcesso = FormatarXml(objXmlDados)
			intAcesso = 1
		End if

		'Cliente
		dblNroSev		= Trim(objRSSolic("Sol_SevSeq"))
		strRazaoSocial	= Trim(objRSSolic("Cli_Nome"))
		strNomeFantasia = Trim(objRSSolic("Cli_NomeFantasia"))
		strContaSev		= Trim(objRSSolic("Cli_CC"))
		strSubContaSev	= Trim(objRSSolic("Cli_SubCC"))
		strGrupo		= Trim(objRSSolic("GCli_ID"))
		
		strconta15 = strContaSev + strSubContaSev
		
		'Alterado por Fabio Pinho em 03/05/2016 - ver 1.0 - Inicio
		'Set Tronco = CreateObject("EOL.CLESP22O")	
		'Tronco.CLESP220 strconta15, strRazaoSocial , strNomeFantasia , SEGMENTO	, PORTE , COD-RETORNO , TXT-MSG

       ''' Set Tronco = CreateObject("EOL.CLESP22O")
        'Homologação
        'Tronco.ServerAddress = "ETBHMGBA.NT.EMBRATEL.COM.BR:1971@RPC/SRVHCLE/CALLNAT"
        'Produção
       ''' Tronco.ServerAddress = "ETBPRDBA.NT.EMBRATEL.COM.BR:1971@RPC/SRVPCLE/CALLNAT"
       ''' Tronco.Logon
       ''' Tronco.CLESP220 strconta15, strRazaoSocial , strNomeFantasia , SEGMENTO , PORTE , COD-RETORNO , TXT-MSG
        'Alterado por Fabio Pinho em 03/05/2016 - ver 1.0 - Inicio

		
		'response.write "<script>alert('"&strRazaoSocial&"')</script>"
		'response.write "<script>alert('"&strNomeFantasia&"')</script>"
		'response.write "<script>alert('"&SEGMENTO&"')</script>"
		'response.write "<script>alert('"&PORTE&"')</script>"
		'response.write "<script>alert('"&COD-RETORNO&"')</script>"
		'response.write "<script>alert('"&TXT-MSG&"')</script>"
		
		'Set objRSSolic = db.execute("CLA_sp_sel_tarefas_APG null, null, null, " & dblSolAPGId)
				
		
		Set objServ = db.execute("CLA_sp_sel_Servico null,'" & Trim(objRSSolic("Ser_Desc")) & "'")
		
		If Not objServ.Eof Then
			dblSerId	= Trim(objServ("Ser_ID"))
		End If
		
		
	End If
	
	'strOriSol					= Trim(objRSSolic("OriSol_ID"))
	strProjEspecial		        = Trim(objRSSolic("Sol_IndProjEspecial"))
	
	strDtPedido				= Formatar_Data(Trim(objRSSolic("Sol_Data")))
	dblVelServico			= Trim(objRSSolic("IDVelAcessoLog"))
	'strTipoContratoServico	= Trim(objRSSolic("Acl_TipoContratoServico"))
	'strNroContrServico		= Trim(objRSSolic("Acl_NContratoServico"))
	dblDesigAcessoPriFull	= Trim(objRSSolic("Acl_IDAcessoLogicoPrincipal"))
	
	if dblDesigAcessoPriFull <> "" then
		dblDesigAcessoPri		= Right(dblDesigAcessoPriFull,len(dblDesigAcessoPriFull)-3)
	End if
	strDtEntrAcesServ		= Formatar_Data(Trim(objRSSolic("Acl_DtDesejadaEntregaAcessoServico")))
	
	strDtIniTemp		= Formatar_Data(Trim(objRSSolic("Acl_DtIniAcessoTemp")))
	strDtFimTemp		= Formatar_Data(Trim(objRSSolic("Acl_DtFimAcessoTemp"))) '@@JKNUP: Correção. BO 50886
	strDtDevolucao		= Formatar_Data(Trim(objRSSolic("Acl_DtDevolAcessoTemp")))
	
	If trim(strOrigem) <> "APG" Then
		dblSerId	= Trim(objRSSolic("Ser_ID"))
	end if 
	strObsProvedor		= Trim(objRSSolic("Sol_Obs"))
	
	dblLocalEntrega = Trim(objRSSolic("Esc_IDEntrega"))
	
	'strDesignacaoServico    = Trim(objRSSolic("Acl_DesignacaoServico"))
	
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
	
	'Usuario de coordenação embratel
	Set objRS = db.execute("CLA_sp_view_agentesolicitacao " & dblSolId)
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
			End Select
			objRS.MoveNext
		Wend
	End if
	
	dblOrgId = Trim(objRSSolic("Org_id"))
	dblStsId = Trim(objRSSolic("Sts_id"))
	strHistoricoSol = Trim(objRSSolic("StsSol_Historico"))
Else
	strbtndes = "disabled"
End if

strHtmlGla	= "<table cellspacing=1 cellpadding=0 width=760px border=0 ><tr class=clsSilver >"
strHtmlGla	= strHtmlGla & "<td width=170px ><font class=clsObrig>:: </font>UserName GLA</td>"
strHtmlGla	= strHtmlGla & "<td colspan=5 >"
strHtmlGla	= strHtmlGla & "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=2 bordercolorlight=#003388 bordercolordark=#ffffff width=355px >"
strHtmlGla	= strHtmlGla & "<tr><td class=lightblue>&nbsp;"
strHtmlGla	= strHtmlGla & "<span id=spnUserNameGLA>" & strUserNameGLA &  "</span>"
strHtmlGla	= strHtmlGla & "</td></tr>"
strHtmlGla	= strHtmlGla & "</table>"
strHtmlGla	= strHtmlGla & "</td>"
strHtmlGla	= strHtmlGla & "</tr>"
strHtmlGla	= strHtmlGla & "<tr class=clsSilver>"
strHtmlGla	= strHtmlGla & "<td width=170px>&nbsp;&nbsp;&nbsp;&nbsp;GLA</td>"
strHtmlGla	= strHtmlGla & "<td width=355px>"
strHtmlGla	= strHtmlGla & "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=2 bordercolorlight=#003388 bordercolordark=#ffffff width=100% >"
strHtmlGla	= strHtmlGla & "<tr><td class=lightblue >&nbsp;"
strHtmlGla	= strHtmlGla & "<span id=spnNomeGLA>" & strNomeGLA &  "</span>"
strHtmlGla	= strHtmlGla & "</td></tr>"
strHtmlGla	= strHtmlGla & "</table>"
strHtmlGla	= strHtmlGla & "</td>"
strHtmlGla	= strHtmlGla & "<td align=right >Ramal&nbsp;</td>"
strHtmlGla	= strHtmlGla & "<td colspan=3 align=left >"
strHtmlGla	= strHtmlGla & "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=2 bordercolorlight=#003388 bordercolordark=#ffffff width=100px >"
strHtmlGla	= strHtmlGla & "<tr><td class=lightblue>&nbsp;"
strHtmlGla	= strHtmlGla & "<span id=spnRamalGLA>" & strRamalGLA & "</span>"
strHtmlGla	= strHtmlGla & "</td></tr>"
strHtmlGla	= strHtmlGla & "</table>"
strHtmlGla	= strHtmlGla & "</td>"
strHtmlGla	= strHtmlGla & "</tr></table>"

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
if not objRS.Eof and not objRS.Bof then
	strObrigaGla = objRS("Cfg_RedirecionamentoCarteira")
Else
	strObrigaGla = 0
End if
Set objRS = db.execute("CLA_sp_sel_usuario 0,'" & Trim(strUserName) & "'")
if Not 	objRS.Eof And Not objRS.Bof then
	strNomeGICL = Replace(Trim(objRS("Usu_Nome")),"'","´")
	strRamalGICL = Replace(Trim(Cstr("" & objRS("Usu_Ramal"))),"'","´")
	strUserNameGICLAtual = Trim(strUserName)
End if
%>

<script>window.name = "Desativacao"</script>
<script language='javascript' src="../javascript/solicitacao.js"></script>
<script language='javascript' src="../javascript/xmlObjects.js"></script>
<script language='javascript' src="../javascript/xmlFacObjects.js"></script>
<script language='javascript' src="../javascript/AlteracaoCad.js"></script>
<SCRIPT LANGUAGE=javascript>
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
var intIndice = <%=intIndice%>

function Message(objXmlRet){
	var intRet = window.showModalDialog('Message.asp',objXmlRet,'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	
	if (intRet != "")
	{
		//Qdo. for processo origem APG 
		if	(document.Form4.hdnOrigem.value == "APG")
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

function VoltarPrincipal()
{
	with (document.forms[0])
	{
		target = self.name
		action = "main.asp"
		submit()
	}
}

function desativa()
{
	var Mensagem
	var acaoUsu
	with(document.forms[0])
	{
		var span = Form3.spnNameGICL
		
			
		if(document.forms[2].txtGICN.value ==""){
			alert('Informe o GIGN!')
			document.forms[2].txtGICN.focus();
			return false
		}
		
		switch(document.forms[0].hdnOrigem.value)
		{
			
			case "des":
				Mensagem = 'Confirma a desativação do serviço?'
				if (confirm(Mensagem))
				{
					
					hdnTipoProcesso.value = 2
					var ie = document.all;
					var url = 'ProcessoDesativacao.asp?hdnAcao=Desativacao&hdnIdLog='+hdn678.value+'&hdnAcfId='+hdnAcfId.value+'&hdnTipoProcesso='+hdnTipoProcesso.value+'&hdnGicN='+document.forms[2].txtGICN.value+'&hdnOEOrigem='+document.forms[0].hdnOEOrigem.value+'&hdnSolId='+document.forms[0].hdnSolId.value+'&hdnAprovisiID='+document.forms[0].hdnAprovisiId.value+'&hdnIdAcessoLogico='+document.forms[0].hdnIdAcessoLogico.value+'&hdnOriSol_ID='+document.forms[0].hdnOriSol_ID.value
			 
					
					//if(ie)
					  //intRet = window.showModalDialog(url, 'window', 'status:no; help:no; dialogWidth:700px; dialogHeight:300px; dialogTop: px; dialogLeft: px; center: Yes; resizable: No');
					//else {
					  intRet = window.open(url, null, 'status:no, help:no, dialogWidth:200px, dialogHeight:300px, dialogTop: px, dialogLeft: px, center: Yes, resizable: No');
						
						//theWin.focus();
					//}
					
					//hdnTipoProcesso.value = 2
					//intRet = window.showModalDialog('ProcessoDesativacao.asp?hdnAcao=Desativacao&hdnIdLog='+hdn678.value+'&hdnAcfId='+hdnAcfId.value+'&hdnTipoProcesso='+hdnTipoProcesso.value+'&hdnGicN='+document.forms[2].txtGICN.value+'&hdnOEOrigem='+document.forms[0].hdnOEOrigem.value+'&hdnSolId='+document.forms[0].hdnSolId.value+'&hdnAprovisiID='+document.forms[0].hdnAprovisiId.value+'&hdnIdAcessoLogico='+document.forms[0].hdnIdAcessoLogico.value+'&hdnOriSol_ID='+document.forms[0].hdnOriSol_ID.value,"","dialogHeight: 300px; dialogWidth: 700px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
					
					try
						{
							objAryRet = intRet.split(",")
							intRet = objAryRet[0]
						}
					catch(e){intRet = 0}
					
					if (parseInt(intRet) == 146 || parseInt(intRet) == 145 || parseInt(intRet) == 117 || parseInt(intRet) == 2 || parseInt(intRet) == 124)
						{
							//alert(parseInt(intRet));
							//AtualizaDados(objAryRet[1],0)
							
							//spnSolId.innerHTML = objAryRet[1]
							
							var strXml = "<?xml version='1.0' encoding='ISO-8859-1'?>"
							strXml = "<root><CLA_RetornoTmp Msg_Id='155' intOrdem='0' Valor='" +  objAryRet[1] + "' Status='0'><CLA_Mensagem Msg_Titulo='Número da Solicitação gerada'/></CLA_RetornoTmp></root>"
							
							var objXmlRet = new ActiveXObject("Microsoft.XMLDOM");
							objXmlRet.loadXML(strXml);
							Message(objXmlRet);
							
						}
				var data = new Date()
				spnpeddt.innerHTML = data.getDate() + '/' + data.getMonth() + '/' + data.getFullYear()
				document.Form3.btnAlterar.disabled = true
				VoltarOrigem()
				}
				break	
			case "can":
				Mensagem = 'Confirma o cancelamento do serviço?'
				if (confirm(Mensagem))
				{
					
					hdnTipoProcesso.value = 4
					document.forms[0].hdnUsugicN.value = document.forms[2].txtGICN.value
					intRet = window.showModalDialog('ProcessoDesativacao.asp?hdnAcao=Desativacao&hdnIdLog='+hdn678.value+'&hdnAcfId='+hdnAcfId.value+'&hdnTipoProcesso='+hdnTipoProcesso.value+'&hdnGicN='+document.forms[2].txtGICN.value+'&hdnOEOrigem='+document.forms[0].hdnOEOrigem.value+'&hdnSolId='+document.forms[0].hdnSolId.value+'&hdnAprovisiID='+document.forms[0].hdnAprovisiId.value+'&hdnIdAcessoLogico='+document.forms[0].hdnIdAcessoLogico.value,"","dialogHeight: 300px; dialogWidth: 700px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
					try
						{
							objAryRet = intRet.split(",")
							intRet = objAryRet[0]
							
							var strXml = "<?xml version='1.0' encoding='ISO-8859-1'?>"
							strXml = "<root><CLA_RetornoTmp Msg_Id='155' intOrdem='0' Valor='" +  objAryRet[1] + "' Status='0'><CLA_Mensagem Msg_Titulo='Número da Solicitação gerada'/></CLA_RetornoTmp></root>"
			
							var objXmlRet = new ActiveXObject("Microsoft.XMLDOM");
							objXmlRet.loadXML(strXml);
							Message(objXmlRet);
							
							
						}
					catch(e){intRet = 0}
					if (parseInt(intRet) == 146 || parseInt(intRet) == 145 || parseInt(intRet) == 117 || parseInt(intRet) == 2 || parseInt(intRet) == 124)
						{
							spnSolId.innerHTML = objAryRet[1]
							AtualizaDados(objAryRet[1],0)
						}
					if(parseInt(intRet) == 170)
					{
						AtualizaDados(objAryRet[1],1)
					}
				var data = new Date()
				spnpeddt.innerHTML = data.getDate() + '/' + data.getMonth() + '/' + data.getFullYear()
				document.Form3.btnAlterar.disabled = true
				}
				break
		}
	}
	
}

function AtualizaDados(SolId,Informa){
	
	with(document.Form4)
	{
		
		hdnObsSol.value= document.Form1.txtObs.value
		hdnGicN.value=document.Form3.txtGICN.value
		hdnSolId.value= SolId
		intRet = window.showModalDialog('ProcessoAtualizaSol.asp?ObjSol='+hdnObsSol.value+'&GicN='+hdnGicN.value+'&SolId='+SolId+' &Informa='+Informa,"","dialogHeight: 300px; dialogWidth: 700px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
		
		
	
	}
}

function ResgatarAcessoLogico()
{
	with (document.forms[0])
	{
		switch (hdnOrigem.value)
		{
			case "des":
				hdnTipoProcesso.value = 2
				break
			case "alt":
				hdnTipoProcesso.value = 3
				break
			case "can":
				hdnTipoProcesso.value = 4
				break
		}
			hdnAcao.value = "ResgatarAcessoLogico"

			hdn678.value = ""
			hdnSolId.value = ""
			hdnAcfId.value = ""
			method = "post"
			target = "IFrmProcesso2"
			action = "ProcessoAcessoLog.asp"
			submit()

	}
}
</script>
<%
	if (strOriSol = 9)  then
	
		
		'strSQL = "select top 1 Acl_IDAcessoLogico from cla_aprovisionador tab1 where acao in ('ALT','ATV') and Acl_IDAcessoLogico in (select Acl_IDAcessoLogico from cla_aprovisionador as tab2	where tab1.ID_Tarefa = tab2.ID_Tarefa_Can and acao='CAN' and ID_Tarefa_Can is not null and not exists (select top 1 * from cla_aprovisionador tab3 where tab3.ID_Tarefa = tab2.ID_Tarefa_Can and aprovisi_dtCancAuto is not null)) and Acl_IDAcessoLogico in (" & strIDLogico677 & "," & strIDLogico678 & ")"
		strSQL = "select distinct oe_numero from cla_aprovisionador WITH (NOLOCK) where Aprovisi_dtRetornoEntregar is null and aprovisi_dtCancAuto is null and acao = 'ATV' and id_acesso = '" & IdAcesso & "' and aprovisi_id <> " & dblAprovisiId
		Set objRSCan = db.execute(strSQL)
		
		if not objRSCan.Eof then
			
		strbloqcan = true
		%>
		<table cellspacing="1" cellpadding="0" border=0 width="760">
			<tr>
			  <th nowrap colspan="4" style=" BACKGROUND-COLOR: #A80000">&nbsp;•&nbsp;Desativação BLOQUEADO: Aguardando ativação da Ordem numero <%= objRSCan("oe_numero") %></th>
				<!--input type=button class="button" name="btnrefresh" value="Atualizar página" onclick="JavaScript:location.reload(true);" -->
			</tr>
		</table>
		<%
		end if
	end if
%>


<form method="post" name="Form1">
	<input type=hidden name=hdnAcao>
	<input type=hidden name=hdnCboServico>
	<input type=hidden name=hdnNomeCbo>
	<input type=hidden name=hdnNomeLocal>
	<input type=hidden name=hdnUserGICL value="<%=strUserNameGICLAtual%>">
	<input type=hidden name=hdntxtGICL value="<%=strUserNameGICLAtual%>">
	<input type=hidden name=hdnDesigServ>
	<input type=hidden name=hdnOrderEntry>
	
	<input type=hidden name=hdnOrderEntrySis value="<%=strOrderEntrySis%>">
	<input type=hidden name=hdnOrderEntryAno value="<%=strOrderEntryAno%>">
	<input type=hidden name=hdnOrderEntryNro value="<%=strOrderEntryNro%>">
	<input type=hidden name=hdnOrderEntryItem value="<%=strOrderEntryItem%>">
	<input type=hidden name=hdnIdEnd>
	<input type=hidden name=hdnIdEndInterme>
	<input type=hidden name=hdnCNLAtual2>
	<input type=hidden name=hdnDesigAcessoPri>
	<input type=hidden name=hdnDesigAcessoPriDB value="<%=dblDesigAcessoPriFull%>">
	<input type=hidden name=hdnProjEsp>
	<input type=hidden name=hdnIdAcessoLogico value="<%=dblIdLogico%>">
	<input type=hidden name=hdnSolId value="<%=dblSolId%>">
	<input type=hidden name=hdnPedId >
	<input type=hidden name=hdnDtSolicitacao value="<%=strDtPedido%>">
	<input type=hidden name=hdnPadraoDesignacao>
	<input type=hidden name=hdnCtfcId value="<%=dblCtfcId%>" >
	<input type=hidden name=hdnSubAcao>
	<input type=hidden name=hdnXmlReturn value="<%=Request.Form("hdnXmlReturn")%>">
	<%
	if Request.Form("hdn678") = "" then
		hdn678Aprov = dblIdLogico
	else
		hdn678Aprov = Request.Form("hdn678")
	end if
	%>
	
	<input type=hidden name="hdn678" value =<%=hdn678Aprov%>>
	<input type=hidden name="hdnOrigem" value = <%=Request.Form("hdnOrigem")%><%if Request.Form("hdnOrigem") = "" then response.write Lcase(acao) end if%>>
	<input type=hidden name="hdnOEOrigem" value = <%=Request.Form("hdnOEOrigem")%>>
	<input type="hidden" name="hdnOriSol_ID" value="<%=strOriSol%>">
	<input type=hidden name=hdnTipoProcesso>
	<input type=hidden name=hdnAcfId value=<%=Request.Form("hdnAcfId")%>>
	<input type=hidden name=hdnUsugicN value="">
	<input type="hidden" name="hdnAprovisiId" value="<%=dblAprovisiId%>">

<%
If acao = "DES" And (not isnull(strIDLogico)) And Trim(strIDLogico) <> "" then
 
	Set objRSAprov_des = db.execute("select acl_dtdesaticavaoservico from cla_acessologico where acl_dtdesaticavaoservico is null and acl_idacessologico=" & strIDLogico)
	If objRSAprov_des.eof Then
	  Response.Write "<script language=javascript>alert('O Acesso Lógico " & strIDLogico & " foi desativado.');</script>"
		strIDLogico = "" 
		dblIdLogico = NULL
		strbtndes = "disabled"		
	End If
 
	If Left(strIDLogico,3)="677" Then
	  Response.Write "<script language=javascript>alert('Acesso Lógico " & strIDLogico & " inválido.');</script>"
		strIDLogico = "" 
		dblIdLogico = NULL
		strbtndes = "disabled"	
	End If
	end if

if strIDLogico = "" and acao = "DES" then%>
	<table cellspacing="1" cellpadding="0" border=0 width="763">
		<th>
			<th nowrap colspan="4" style=" BACKGROUND-COLOR: #A80000">&nbsp;•&nbsp;Informações de Aprovisionamento - Dados do Item de OE de Origem para facilitar busca do Acesso Lógico</th>
		</th>
	</table>
	
	<table cellspacing="1" cellpadding="0" border=0 width="760">
		<tr class="clsSilver">
		  <td width="465">Razão Social: <%=strPov_Razao%></td>
		  <td colspan="2">Conta Corrente: <%=strPov_CC%></td>
		  <td width="96">SubConta: <%=strPov_SubCC%></td>
		</tr>
	</table>
	<table cellspacing="1" cellpadding="0" border=0 width="760">
		<tr class="clsSilver">
		  <td width="298">Designação do Serviço: <%=strPov_DesigServ%></td>
		  <td width="426" colspan="2">Velocidade do Serviço: <%=strPov_VelServ%></td>
		  <td width="249">Nº Contrato de Serviço: <%=strPov_ContratoServ%></td>
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
		  <th nowrap colspan="4" style=" BACKGROUND-COLOR: #A80000">&nbsp;•&nbsp;POVOAMENTO SOB DEMANDA: Associação de Serviços enviados pelo Aprovisionador com Acesso Lógico cadastrado no CLA</th>
		</tr>
	</table>
	<table cellspacing="1" cellpadding="0" border=0 width="760">
		<tr class="clsSilver">
		  <td width="517">Informar o Acesso Lógico que será associado ao serviço que sofrerá o processo de 
			<%
			select case trim(ucase(acao))
			  Case "ATV"
			    response.write "ativação"
			  Case "DES"
			    response.write "desativação"
			  Case "ALT"
			    response.write "alteração"
			  Case "CAN"
			    response.write "cancelamento"
			end select
			%>.</td>
			<td colspan="3">
			    <input type="hidden" name="hdnAcaoAprov" value="<%=strAcao%>">
				<input type="hidden" name="hdnOriSolID" value="<%=strOriSol%>">			
				<input type="text" class="text" name="txtNroLogico" value="<%=dblIdLogico%>" maxlength="10" size="11" onkeyup="ValidarTipo(this,0)">&nbsp;&nbsp; &nbsp;
				<input type="button" class="button" name="associarlogico" value="Associar Lógico" title="Associar Acesso Lógico" style="cursor:hand" onClick="AssociarLogico()" tabindex=0 accesskey="A" onmouseover="showtip(this,event,'Associar Lógico (Alt+A)');">
			</td>
		</tr>
	</table>
<%end if%>	
<table cellspacing="1" cellpadding="0" border=0 width="760">
	<tr >
		<th nowrap>&nbsp;•&nbsp;Solicitação de Acesso</th>
		<th >&nbsp;Nº&nbsp;:&nbsp;<span id=spnSolId><%=dblSolId%></Span></th>
		<%
		if trim(ucase(acao)) = "CAN" then
		%>
			<th nowrap>&nbsp;Cancelamento de Solicitação</th>
		<%
		elseif trim(ucase(acao)) = "DES" then
		%>
			<th nowrap>&nbsp;Desativação</th>
		<%
		end if
		%>
		<th nowrap>&nbsp;Acesso Lógico&nbsp;:&nbsp;<%=dblIdLogico%></th>
		<th >&nbsp;Data&nbsp;:&nbsp;<span id=spnpeddt><%=strDtPedido%></span></th>
	</tr>
</table>
<table border=0 cellspacing="1" cellpadding="0" width="760">
	<tr >
		<th colspan=4>&nbsp;•&nbsp;Informações do Cliente</th>
	</tr>
	<tr class="clsSilver">
		<td width="170">&nbsp;&nbsp;&nbsp;&nbsp;Sev para procura</td>
		<td colspan="3">
			<input type="text" class="text" name="txtNroSev" value="<%=dblNroSev%>" <%=bbloqueia%> maxlength="8" size="10" onkeyup="ValidarTipo(this,0)">&nbsp;
			<input type="button" class="button" name="procurarsev" value=" Procurar Sev  " onClick="ResgatarSev()" tabindex=-1 accesskey="P" onmouseover="showtip(this,event,'Procurar uma SEV no sistema SSA (Alt+P)');" <%=bbloqueia%>>
		</td>
	</tr>
	<tr>
		<td class="clsSilver" rowspan="2">&nbsp;
			Projeto Especial
		</td>
		<td class="clsSilver" rowspan="2">&nbsp;
			<input type="radio" name="rdoProjEspecial" onClick="javascript:document.Form1.hdnProjEsp.value = 'S';" value="S" <%if strProjEspecial = "S" then%> checked <%end if%> <%=bbloqueia%>>&nbsp; Sim
			<input type="radio" name="rdoProjEspecial" onClick="javascript:document.Form1.hdnProjEsp.value = 'N';" value="N" <%if strProjEspecial <> "S" then%> checked <%end if%> <%=bbloqueia%>>&nbsp; Não
		</td>
		<td colspan=2 class="clsSilver">
		    &nbsp;&nbsp;Grupo <span align=right>
		<select name="cboGrupo" onChange="CheckGrupo()" <%=bbloqueia%>>
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
		<td colspan=2 class="clsSilver">
			&nbsp;Origem Solicitação 
			<select name="cboOrigemSol" disabled>
				<option value="<%=strOriSol%>"><%=strOriDesc%></option>
			</select>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170"><font class="clsObrig">:: </font>Razão Social</td>
		<td colspan="3" >
			<input type="text" class="text" name="txtRazaoSocial"  maxlength="80" size="80" value="<%=strRazaoSocial%>" <%=bbloqueia%> onblur="ResgatarGLA()">
			
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170"><span id=spnLabelCliente></span></td>
		<td colspan="3"><span id=spnCliente></span></td>
	</tr>
	<tr class="clsSilver">
		<td width="170"><font class="clsObrig">:: </font>Nome Fantasia</td>
		<td colspan="3" >
			<input type="text" class="text" name="txtNomeFantasia"  maxlength="80" size="80" value="<%=strNomeFantasia%>" <%=bbloqueia%>>
		</td>
		
	</tr>
	<tr class="clsSilver">
		<td width="170" ><font class="clsObrig">:: </font>Conta Corrente</td>
		<td width=183>
			<input type=text name=txtContaSev class="text" size=11 maxlength=11 onKeyUp="ValidarTipo(this,0)" onblur="CompletarCampo(this)" TIPO="N" value="<%=strContaSev%>" <%=bbloqueia%>>
		</td>
		<td align=right width=198 ><font class="clsObrig">:: </font>Sub Conta&nbsp;</td>
		<td width="204" >
			<input type=text name=txtSubContaSev class="text" size=4 maxlength=4 onKeyUp="ValidarTipo(this,0)" onblur="CompletarCampo(this)" TIPO="N" value="<%=strSubContaSev%>" <%=bbloqueia%>>
		</td>
	</tr>
	<tr class="clsSilver" nowrap>
		<td width="170" nowrap><font class="clsObrig">:: </font>Segmento</td>
		<td width=279>
			<input type=text class="text" name=txtSegmento size=22 maxlength=22
			<%=bbloqueia%>
			 value="<%=SEGMENTO%>">
		</td>
		<td align=right width=198 >
		<font class="clsObrig">:: </font>Porte&nbsp;</td>
		<td width="189">
			<input type=text name=txtPorte class="text" size=22 maxlength=22
			<%=bbloqueia%>
			value="<%=PORTE%>">
		</td>
		<!--Alterado por Fabio Pinho em 2/05/2016 - ver 1.0 - Inicio-->
		<!--<td>&nbsp;</td>-->
		<!--Alterado por Fabio Pinho em 2/05/2016 - ver 1.0 - Fim-->
	</tr>
</table>
<table border=0 cellspacing="1" cellpadding="0" width="760">
	<tr >
		<th colspan=4 >
			&nbsp;•&nbsp;Informações do Serviço&nbsp;
		</th>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Order Entry</td>
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
		</tr>
		<tr class=clsSilver>
			<td>
				<select name="cboSistemaOrderEntry" onChange="SistemaOrderEntry(this);hdnOrderEntrySis.value=this.value;" <%=bbloqueia%>>
					<Option ></Option>
					<Option value="APG"			<%if strOrderEntrySis = "APG" then Response.Write " selected " End If%>>APG</Option>
					<Option value="ASMS"			<%if strOrderEntrySis = "ASMS" then Response.Write " selected " End If%>>ASMS</Option>
					<Option value="CFD"			<%if strOrderEntrySis = "CFD" then Response.Write " selected " End If%>>CFD</Option>
					<Option value="SGA VOZ 0300"			<%if strOrderEntrySis = "SGA VOZ 0300" then Response.Write " selected " End If%>>SGA VOZ 0300</Option>
					<Option value="SGA VOZ 0800 FASE 1"		<%if strOrderEntrySis = "SGA VOZ 0800 FASE 1" then Response.Write " selected " End If%>>SGA VOZ 0800 FASE 1</Option>
					<Option value="SGA VOZ VIP'S"			<%if strOrderEntrySis = "SGA VOZ VIP'S" then Response.Write " selected " End If%>>SGA VOZ VIP'S</Option>
					<Option value="SGA VOZ"		<%if strOrderEntrySis = "SGAV" then Response.Write " selected " End If%>>SGA Voz</Option>
					<Option value="SGA PLUS"	<%if strOrderEntrySis = "SGAP" then Response.Write " selected " End If%>>SGA PLUS</Option>
					<Option value="ADFAC"		<%if strOrderEntrySis = "ADFAC" then Response.Write " selected " End If%>>ADFAC</Option>
					<Option value="CFM"			<%if strOrderEntrySis = "CFM" then Response.Write " selected " End If%>>CFM</Option>
					<Option value="CFT"			<%if strOrderEntrySis = "CFT" then Response.Write " selected " End If%>>CFT</Option>
				</Select>
			</td>
			<td>-</td>
			<td><input type="text" class="text" onblur="CompletarCampo(this);hdnOrderEntryAno.value=this.value;" onkeyup="ValidarTipo(this,0)" maxlength=4 size=4 name=txtOrderEntry TIPO="N" value="<%=strOrderEntryAno%>" <%=bbloqueia%>></td>
			<td>-</td>
			<td><input type="text" class="text" onblur="CompletarCampo(this);hdnOrderEntryNro.value=this.value;" onkeyup="ValidarTipo(this,0)" maxlength=7 size=7 name=txtOrderEntry TIPO="N" value="<%=strOrderEntryNro%>" <%=bbloqueia%>></td>
			<td>-</td>
			<td><input type="text" class="text" onblur="CompletarCampo(this);hdnOrderEntryItem.value=this.value;" onkeyup="ValidarTipo(this,0)" maxlength=3 size=3 name=txtOrderEntry TIPO="N" value="<%=strOrderEntryItem%>" <%=bbloqueia%>></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Serviço</td>
		<td>
			<%if strIDLogico = "" and acao = "DES" then%>
			<select name="cboServicoPedido" onchange="ResgatarServico(this)" <%=bbloqueia%>>
				<%'Seleciona servico
				set objRS = db.execute("CLA_sp_sel_servico null,null,null,1")
				
				While Not objRS.eof
					strItemSel = ""
					if Trim(strPov_SerID) = Trim(objRS("Ser_ID")) then 
						strItemSel = " Selected " 
					end if
					Response.Write "<Option value='" & objRS("Ser_ID") & "'" & strItemSel & ">" & objRS("Ser_Desc") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
				%>
				</select>
			<%else%>
				<select name="cboServicoPedido" onchange="ResgatarServico(this)" <%=bbloqueia%>>
				<%
				'Seleciona servico
				set objRS = db.execute("CLA_sp_sel_servico null,null,null,1")
				
				While Not objRS.eof
					strItemSel = ""
					if Trim(dblSerId) = Trim(objRS("Ser_ID")) then 
						strItemSel = " Selected " 
					end if
					Response.Write "<Option value='" & objRS("Ser_ID") & "'" & strItemSel & ">" & objRS("Ser_Desc") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
				%>
				</select>
			<%end if%>
		</td>
		<td width="150px" align=right><font class="clsObrig">:: </font>Velocidade&nbsp;</td>
		<td width="200px"><span id=spnVelServico>
				<!--Alterado por Fabio Pinho em 2/05/2016 - ver 1.0 - Inicio-->
				<!-- <select name="cboVelServico" onChange="SelVelAcesso(this)" style="width:200px" <%=bbloqueia%>> -->
				<select name="cboVelServico" onChange="SelVelAcesso(this)" style="width:150px" <%=bbloqueia%>>
				<!--Alterado por Fabio Pinho em 2/05/2016 - ver 1.0 - Fim-->
				
				<%if strIDLogico = "" and acao = "DES" then%>				
					<%
					if Trim(dblSerId) <> "" then
						set objRS = db.execute("CLA_sp_sel_AssocServVeloc null," & dblSerId)
						While Not objRS.eof
							strItemSel = ""
							if Trim(strPov_VelID) = Trim(objRS("Vel_ID")) then 
								strItemSel = " Selected " 
							End if
							Response.Write "<Option value='" & objRS("Vel_ID") & "'" & strItemSel & ">" & Trim(objRS("Vel_Desc")) & "</Option>"
							objRS.MoveNext
						Wend
						strItemSel = ""
					End if
				else					
					if Trim(dblSerId) <> "" then
						set objRS = db.execute("CLA_sp_sel_AssocServVeloc null," & dblSerId)
						While Not objRS.eof
							strItemSel = ""
							if Trim(dblVelServico) = Trim(objRS("Vel_ID")) then 
								strItemSel = " Selected " 
							End if
							Response.Write "<Option value='" & objRS("Vel_ID") & "'" & strItemSel & ">" & Trim(objRS("Vel_Desc")) & "</Option>"
							objRS.MoveNext
						Wend
						strItemSel = ""
					End if
				end if%>
			</span>
		</td>
	</tr>
	
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Antecipação de Acesso </td>
		<td colspan="3">
			<input  type="radio" name="rdoAntAcesso" value="S" onclick="document.Form1.rdoAntAcesso[0].checked=true;document.Form1.rdoAntAcesso[1].checked=false;DesabilitarDesignacao(1)" <%if strAntAcesso = "S" then%>checked<%end if%> disabled>&nbsp; Sim
			<input  type="radio" name="rdoAntAcesso" value="N" onclick="document.Form1.rdoAntAcesso[0].checked=false;document.Form1.rdoAntAcesso[1].checked=true;DesabilitarDesignacao(2)" <%if strAntAcesso = "N" then%>checked<%end if%> disabled>&nbsp; Não
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Designação do Serviço</td>
		<td colspan="3">
				<input type="text" class="text" name="txtdesignacaoServico"
				<%=bbloqueia%>
				value="<%=strDesignacaoServico%>" maxlength="22" size="30"><br>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Nº Contrato Serviço</td>
		<td colspan=3>
			<table rules="groups" cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="70%" >
				<tr><td nowrap width=200px >
					<input type=radio name=rdoNroContrato value=1 onClick="spnDescNroContr.innerHTML= 'Ex.: VEM-11 XXX000012003'" checked <%if strTipoContratoServico = "1" then Response.Write " checked " End if%> <%=bbloqueia%>>Contrato de Serviço</td><td></td></tr>
				<tr>
					<td nowrap><input type=radio name=rdoNroContrato onClick="spnDescNroContr.innerHTML= 'Padrão: A22'" value=2 <%if strTipoContratoServico = "2" then Response.Write " checked " End if%> <%=bbloqueia%>>Contrato de Referência</td>
					<td nowrap>
						<input type="text" class="text" name="txtNroContrServico" value="<%=strNroContrServico%>" maxlength="22" size="30" <%=bbloqueia%>><br>
						<span id=spnDescNroContr>Ex.: VEM-11 XXX00012003</span>
					</td>
				</tr>
				<tr><td nowrap><input type=radio name=rdoNroContrato onClick="spnDescNroContr.innerHTML= 'Padrão: A22'" value=3 <%if strTipoContratoServico = "3" then Response.Write " checked " End if%> <%=bbloqueia%>>Carta de Compromisso</td><td></td></tr>
			</table>
		</td>
	</tr>
	<tr class="clsSilver">
		<td nowrap width=170px><font class="clsObrig">:: </font>Data Desejada de Entrega<br>&nbsp;&nbsp;&nbsp; do Acesso ao Serviço</td>
		<td><input type="text" class="text" name="txtDtEntrAcesServ" value="<%=strDtEntrAcesServ%>" <%=bbloqueia%> maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaaa)</td>
		<td nowrap>&nbsp;Data Prevista de Entrega<br>&nbsp;do Acesso pelo Provedor</td>
		<td ><input type="text" class="text" name="txtDtPrevEntrAcesProv" value="<%=strDtPrevEntrAcesProv%>" <%=bbloqueia%> maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaaa)</td>
	</tr>
	<tr class="clsSilver">
		<td rowspan=2>&nbsp;&nbsp;&nbsp;&nbsp;Acesso Temporário<br>&nbsp;&nbsp;&nbsp;&nbsp;(dd/mm/aaaa)</td>
		<td >&nbsp;Início&nbsp;</td>
		<td >&nbsp;Fim&nbsp;</td>
		<td >&nbsp;Devolução&nbsp;</td>
	</tr>
	<tr class="clsSilver">
		<td ><input type="text" class="text" name="txtDtIniTemp"  value="<%=strDtIniTemp%>" <%=bbloqueia%>  maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;</td>
		<td ><input type="text" class="text" name="txtDtFimTemp" value="<%=strDtFimTemp%>"  <%=bbloqueia%>  maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;</td>
		<td ><input type="text" class="text" name="txtDtDevolucao" value="<%=strDtDevolucao%>"  <%=bbloqueia%>  maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Designação do Acesso<br>&nbsp;&nbsp;&nbsp; Principal (678)</td>
		<td colspan=3>
			<input type="text" class="text" name="txtDesigAcessoPri0" <%=bbloqueia%> maxlength="3" size="3" value=678 readOnly>
			<input type="text" class="text" name="txtDesigAcessoPri"  value="<%=dblDesigAcessoPri%>" <%=bbloqueia%> maxlength="7" size="9" onKeyUp="ValidarTipo(this,0)" >(678N7)
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Observações p/ Provedor</td>
		<td colspan="3"><textarea name="txtObs" onkeydown="MaxLength(this,300);" onKeyUp="ValidarTipo(this,2)" cols="50" rows="3"><%=strObsProvedor%></textarea></td>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Resposta <br>&nbsp;&nbsp;&nbsp;&nbsp;Indicada no SSA</td>
		<td colspan="3">
			<span id=strProvedorSelSev>&nbsp;</span>
		</td>
	</tr>
</Form>
</table>

<table cellspacing=1 cellpadding=0 width=760 border=0>
<Form name=Form2 method=Post>
<input type="hidden" name="hdnOriSol_ID" value="<%=strOriSol%>">
<input type=hidden name=hdnIntIndice>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnSubAcao>
<input type=hidden name=hdnProvedor>
<input type=hidden name=hdnTipoCEP>
<input type=hidden name=hdnCEP>
<input type=hidden name=hdnCNLNome>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnCNLAtual1>
<input type=hidden name=hdnNomeTxtCidDesc>
<input type=hidden name=hdnNomeCboCid>
<input type=hidden name=hdnCNLAtual>
<input type=hidden name=hdnUserGICL value="<%=strUserNameGICLAtual%>">
<input type=hidden name=hdntxtGLA value="<%=strUserNameGLA%>">
<input type=hidden name=hdntxtGLAE value="">
<input type=hidden name=hdnCtfcId value="<%=dblCtfcId%>" >
<input type=hidden name=hdnRazaoSocial>
<input type=hidden name=hdnChaveAcessoFis>
<input type=hidden name=hdnIdAcessoFisico	value="<%'=strIdAcessoFisicoInst%>">
<input type=hidden name=hdnIdAcessoFisico1	value="<%'=strIdAcessoFisicoPtoI%>">
<input type=hidden name=hdnPropIdFisico	value="<%=strPropAcessoFis%>">
<input type=hidden name=hdnPropIdFisico1>
<input type=hidden name=hdnCompartilhamento		value="0">
<input type=hidden name=hdnNodeCompartilhado	value="0">
<input type=hidden name=hdnCompartilhamento1	value="0">
<input type=hidden name=hdnNovoPedido>
<input type=hidden name=hdnTecnologia>
<input type=hidden name=hdnstrAcessoTipoRede >
<input type=hidden name=hdnVelAcessoFisSel>
<input type=hidden name=hdnSolId value=<%=dblSolId%>>
<input type=hidden name=hdnPedId>
<input type=hidden name=hdnProId>
<input type=hidden name=hdnTipoProcesso value=<%=intTipoProcesso%>>
<input type=hidden name=hdnEstacaoOrigem>
<input type=hidden name=hdnEstacaoDestino>
<input type=hidden name=hdnObrigaGla value="<%=strObrigaGla%>">
<input type=hidden name=hdnUsuID value="<%=dblUsuID %>">
<input type=hidden name=hdnTipoTec>
<input type=hidden name=hdnCNLCliente>
<input type=hidden name=hdnAprovisiId value="<%=dblAprovisiId%>">

	<tr><th colspan=6>&nbsp;•&nbsp;Acessos Físicos Utilizados</th></tr>
	<tr><td colspan=6>
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
		<td colSpan=6>
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
	<tr>
		<th colSpan=4>&nbsp;•&nbsp;Informações do Acesso&nbsp;</th>
	</tr>
	<tr class="clsSilver">
			<td width=170px ><font class="clsObrig">:: </font>Prop do Acesso Físico</td>
			<td nowrap >
				<input type=radio name=rdoPropAcessoFisico value="TER"	Index=0	<%if strPropAcessoFisico = "TER" then Response.Write " checked " End if%> onclick="EsconderTecnologia(0);ResgatarTecVel()" <%=bbloqueia%>>Terceiro&nbsp;&nbsp;&nbsp;
				<input type=radio name=rdoPropAcessoFisico value="EBT"	Index=1	<%if strPropAcessoFisico = "EBT" then Response.Write " checked " End if%> onclick="EsconderTecnologia(0);ResgatarTecVel()" <%=bbloqueia%>>EBT&nbsp;&nbsp;&nbsp;
				<input type=radio name=rdoPropAcessoFisico value="CLI"	Index=2	<%if strPropAcessoFisico = "CLI" then Response.Write " checked " End if%> onclick="EsconderTecnologia(0);ResgatarTecVel()" <%=bbloqueia%>>Cliente&nbsp;&nbsp;&nbsp;
			<td nowrap colspan=2>
				<%if TipoVel(dblTecId) <> "" then%>
					<div id=divTecnologia style="display:'';POSITION:relative">
				<%Else%>
					<div id=divTecnologia style="display:none;POSITION:relative">
				<%End if%>
				<Select name=cboTecnologia onChange="RetornaCboTipoRadio(this[this.selectedIndex].innerText,this.value,'<% = strTrdID %>', '<% = strVersao %>');ResgatarTecVel()" <%=bbloqueia%>>
					<Option value="">:: TECNOLOGIA EBT</Option>
					<%
					set objRS = db.execute("CLA_sp_sel_tecnologia 0")
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
	<tr class="clsSilver">
		<td id = tdRadio width=170px></td>
		<td colspan = 3><span ID =spnTipoRadio></span></td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Vel do Acesso Físico</td>
		<td colspan=3><span id=spnVelAcessoFis>
			<select name="cboVelAcesso" style="width:150px" onChange="MostrarTipoVel(this)" <%=bbloqueia%>>
				<option ></option>
				<%
				SET objRS = db.execute("CLA_sp_sel_velocidade")
				While Not objRS.eof
					strItemSel = ""
					if Trim(strVelAcesso) = Trim(objRS("Vel_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value='" & Trim(objRS("Vel_ID")) & "'" & strItemSel & ">" & objRS("Vel_Desc") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
				%>
			</select></span>&nbsp;&nbsp;<font class="clsObrig">:: </font>Qtde de Acesso(s) Fisico(s)&nbsp;<input type="text" class="text" name="txtQtdeCircuitos" value=1  maxlength="2" size="2" onKeyUp="ValidarTipo(this,0)" value="<%=dblQtdeCircuitos%>" <%=bbloqueia%>>&nbsp;&nbsp;
			<%if TipoVel(strTipoVel) <> "" then%>
				<div id=divTipoVel style="display:'';POSITION:absolute">
			<%Else%>
				<div id=divTipoVel style="display:none;POSITION:absolute">
			<%End if%>
			<select name="cboTipoVel" style="width:170px" <%=bbloqueia%>>
				<option value="">TIPO DE VELOCIDADE</option>
				<option value="1" <%if strTipoVel=1 then Response.Write " Selected " %>>ESTRUTURADA</option>
				<option value="0" <%if strTipoVel=0 then Response.Write " Selected " %>>NÃO ESTRUTURADA</option>
			</select>
			</div>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Provedor</td>
		<td colspan="3">
			<select name="cboProvedor" onChange="ResgatarPromocaoRegime(this)" <%=bbloqueia%>>
				<option value=""></option>
				<%
				set objRS = db.execute("CLA_sp_sel_provedor 0")
				While not objRS.Eof
					strItemSel = ""
					if Trim(dblProId) = Trim(objRS("Pro_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value='" & Trim(objRS("Pro_ID")) & "'" & strItemSel & ">" & objRS("Pro_Nome") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
				%>
			</select>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px>&nbsp;&nbsp;&nbsp;&nbsp;Prazos de Contratação<br>&nbsp;&nbsp;&nbsp;&nbsp;de Acesso</td>
		<td colspan=3>
			<span id=spnRegimeCntr <%=bbloqueia%>>
				<select name="cboRegimeCntr">
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
		<td width=170px>&nbsp;&nbsp;&nbsp;Promoção</td>
		<td colspan=3 >
			<span id=spnPromocao>
			<select name="cboPromocao" style="width:170px" <%=bbloqueia%>>
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
		<td width=170px>&nbsp;&nbsp;&nbsp;&nbsp;Código SAP</td>
		<td >
			<input type="text" class="text" name="txtCodSAP"  maxlength="7" size="10" onKeyUp="ValidarTipo(this,0)" value="<%=strCodSap%>" <%=bbloqueia%>>&nbsp;(N7)
		</td>
		<td >&nbsp;&nbsp;&nbsp;Número PI&nbsp;</td>
		<td>
			<input type="text" class="text" name="txtNroPI"  maxlength="7" size="10" onKeyUp="ValidarTipo(this,0)" value="<%=dblNroPI%>" <%=bbloqueia%>>&nbsp;(N7)
		</td>
	</tr>
	<tr class=clsSilver2>
		<td width=170px >&nbsp;Endereço Origem&nbsp;</td>
		<td nowrap colspan=3>
			<font class=clsObrig>:: </font>PONTO&nbsp;
			<select name="cboTipoPonto" onChange="TipoOrigem(this.value)" <%=bbloqueia%>>
				<option value=""></option>
				<option value="I" <%if Trim(strTipoPonto) = "I" then Response.Write " selected " end if%>>CLIENTE</option>
				<option value="T" <%if Trim(strTipoPonto) = "T" then Response.Write " selected " end if%>>INTERMEDIÁRIO</option>
			</select>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px nowrap><span id=spnOrigem>&nbsp;&nbsp;&nbsp;Sigla Estação Origem(CNL)</span></td>
		<td colspan=3>
			<input type="text" class="text" name="txtCNLSiglaCentroCli"  maxlength="4" size="8" onKeyUp="ValidarTipo(this,1)"	 onblur="CompletarCampo(this)" TIPO="A" <%=bbloqueia%>>&nbsp;Complemento
			<input type="text" class="text" name="txtComplSiglaCentroCli"  maxlength="6" size="10" onKeyUp="ValidarTipo(this,2)" onblur="CompletarCampo(this);ResgatarEstacaoOrigem(document.Form2.txtCNLSiglaCentroCli,document.Form2.txtComplSiglaCentroCli)" TIPO="A" <%=bbloqueia%>>&nbsp;
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>UF</td>
		<td>
			<select name="cboUFEnd" <%=bbloqueia%>>
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
		</td>
		<td nowrap><font class="clsObrig">:: </font>Cidade (CNL)</td>
		<td nowrap>
			<input type=text size=5 maxlength=4 class=text name="txtEndCid" value="<%=strEndCid%>" onBlur="if (ValidarTipo(this,1)){ResgatarCidade(document.forms[1].cboUFEnd,1,this)}" <%=bbloqueia%>>&nbsp;
			<input type=text size=27 readonly style="BACKGROUND-COLOR:#eeeeee" class=text name="txtEndCidDesc" value="<%=strEndCidDesc%>" tabIndex=-1 <%=bbloqueia%>>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Tipo do Logradouro</td>
		<td colspan="0">
			<select name="cboLogrEnd" <%=bbloqueia%>>
				<option value=""></option>
				<%
				set objRS = db.execute("CLA_sp_sel_tplogradouro")
				While not objRS.Eof
					strItemSel = ""
					if Trim(strLogrEnd) = Trim(objRS("Tpl_Sigla")) then strItemSel = " Selected " End if
					Response.Write "<Option value=" & Trim(objRS("Tpl_Sigla")) & strItemSel & ">" & Trim(objRS("Tpl_Sigla")) & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
				%>
			</select>
		</td>
		<td><font class="clsObrig">:: </font>Nome Logr</td>
		<td nowrap>
			<input type="text" class="text" name="txtEnd"  value="<%=strEnd%>" maxlength="60" size="35" <%=bbloqueia%>>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font> Número</td>
		<td>
			<input type="text" class="text" name="txtNroEnd" value="<%=strNroEnd%>" maxlength="5" size="5" <%=bbloqueia%>>
		</td>
		<td>&nbsp;&nbsp;&nbsp;Complemento</td>
		<td >
			<input type="text" class="text" name="txtComplEnd"  value="<%=strComplEnd%>" maxlength="25" size="25" <%=bbloqueia%>>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Bairro</td>
		<td>
			<input type="text" class="text" name="txtBairroEnd"  value="<%=strBairroEnd%>" maxlength="30" size="30" <%=bbloqueia%>>&nbsp;
		</td>
		<td nowrap><font class="clsObrig">:: </font>CEP&nbsp;(99999-999)</td>
		<td>
			<input type="text" class="text" name="txtCepEnd"  value="<%=strCepEnd%>" maxlength="9" size="10" onKeyPress="OnlyNumbers();AdicionaBarraCep(this)" <%=bbloqueia%>>&nbsp;
			<input type=button name=btnProcurarCepInstala value="Procurar CEP" class="button" onclick="ProcurarCEP(1,1)" tabindex=-1 onmouseover="showtip(this,event,'Procurar por CEP exato ou pelos 5 primeiros dígitos (Alt+D)');" accesskey="D" <%=bbloqueia%>>
		</td>
	</tr>
	<tr class="clsSilver">
		<td colspan=4 align=right><span id=spnCEPSInstala></span></td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Contato</td>
		<td>
			<input type="text" class="text" name="txtContatoEnd" value="<%=strContatoEnd%>" maxlength="30" size="30" <%=bbloqueia%>>
		</td>
		<td><font class="clsObrig">:: </font>Telefone</td>
		<td >
			<input type="text" class="text" name="txtTelEndArea" maxlength="2" size="2" onkeyUp="ValidarTipo(this,0)" <%=bbloqueia%>>&nbsp;
			<input type="text" class="text" name="txtTelEnd" value="<%=strTelEnd%>" maxlength="9" size="10" onkeyUp="ValidarTipo(this,0)" <%=bbloqueia%>>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>CNPJ</td>
		<td colspan="3">
			<input type="text" class="text" name="txtCNPJ"  maxlength="14" size="16" onKeyUp="ValidarTipo(this,0)" value="<%=dblCNPJ%>" <%=bbloqueia%>>&nbsp;(99999999999999)
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px>&nbsp;&nbsp;&nbsp;&nbsp;I.E.</td>
		<td >
			<input type="text" class="text" name="txtIE"  maxlength="15" size="17" onKeyUp="ValidarTipo(this,0)" value="<%=strIE%>" <%=bbloqueia%>>
		</td>
		<td >&nbsp;&nbsp;&nbsp;I.M&nbsp;</td>
		<td>
			<input type="text" class="text" name="txtIM"  maxlength="15" size="17" onKeyUp="ValidarTipo(this,0)" value="<%=strIM%>" <%=bbloqueia%>>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px nowrap>&nbsp;&nbsp;&nbsp;&nbsp;Proprietário do Endereço</td>
		<td colspan="3">
			<input type="text" class="text" name="txtPropEnd"  maxlength="55" size="50" value="<%=strPropEnd%>" <%=bbloqueia%>>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class=clsObrig>:: </font><font color="#FF0000">Interface</font></td>
		<td colspan="3">
			<Select name="cboInterFaceEnd" <%=bbloqueia%>>
				<Option value=""></Option>
				<%
				set objRS = db.execute("CLA_sp_sel_interface")
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
					<td><input type="text" class="text" name="txtCNLSiglaCentroCliDest"  maxlength="4" size="8" onKeyUp="ValidarTipo(this,1)" value="<%=strCNLSiglaCli%>" onblur="CompletarCampo(this)" TIPO="A" <%=bbloqueia%>>&nbsp;</td>
					<td>&nbsp;<input type="text" class="text" name="txtComplSiglaCentroCliDest"  maxlength="6" size="10" onKeyUp="ValidarTipo(this,2)" value="<%=strCNLSiglaCli%>" onblur="CompletarCampo(this);ResgatarEstacaoDestino(document.Form2.txtCNLSiglaCentroCliDest,document.Form2.txtComplSiglaCentroCliDest)" TIPO="A" <%=bbloqueia%>>&nbsp;</td>
					<td>&nbsp;<TEXTAREA rows=2 cols=66 name="txtEndEstacaoEntrega" readonly tabIndex=-1></TEXTAREA></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class=clsObrig>:: </font><font color="#FF0000">Interface</font></td>
		<td colspan="3">
			<Select name="cboInterFaceEndFis" <%=bbloqueia%>>
				<Option value=""></Option>
				<%
				set objRS = db.execute("CLA_sp_sel_interface")
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
									height		= "45px"
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
			<span id="spnBtnLimparIdFis1"></span>
			<input type=button		class=button	name=btnAddAcesso 		style="width:150px" value="Alterar" onmouseover="showtip(this,event,'Atualizar um acesso da lista (Alt+A)');" onClick="AdicionarAcessoListaAlteracao()" accesskey="A" <%=bbloqueia%>>&nbsp;
			<input type=button		class=button	name=btnLimparAcesso	style="width:150px" value="Limpar" onClick="LimparInfoAcesso()" accesskey="L" onmouseover="showtip(this,event,'Limpar dados do Acesso (Alt+L)');" <%=bbloqueia%>>&nbsp;
		</td>
	</tr>
</Form>
</table>

		<table border=0 cellspacing="1" cellpadding="0" width="760" >
			<Form name=Form3 method=Post>
				<input type="hidden" name="hdnOriSol_ID" value="<%=strOriSol%>">
				<input type=hidden name=hdnAcao>
				<input type=hidden name=hdnEstacaoAtual>
				<input type=hidden name=hdnCtfcId value="<%=dblCtfcId%>" >
				<input type=hidden name=hdntxtGICN value="<%=strUserNameGICN%>">
				<input type=hidden name=hdnCoordenacaoAtual>
				<input type="hidden" name="hdnAprovisiId" value="<%=dblAprovisiId%>">
			<tr>
				<th colspan="4" >&nbsp;•&nbsp;Informações da Embratel</th>
			</tr>
			<tr class="clsSilver">
				<td width="170px"><font class="clsObrig">:: </font>Local de Entrega<br>&nbsp;&nbsp;&nbsp;do Acesso Lógico</td>
				<td colspan="3">
					<select name="cboLocalEntrega" onChange="ResgatarEnderecoEstacao(this);SelecionarLocalConfig(this)"  <%=bbloqueia%>>
						<option value=""></option>
						<%
						set objRS = db.execute("CLA_sp_sel_estacao " & Trim(dblLocalEntrega) )
						While not objRS.Eof
							strItemSel = ""
							if Trim(dblLocalEntrega) = Trim(objRS("Esc_ID")) then strItemSel = " Selected " End if
							Response.Write "<Option value=" & objRS("Esc_ID") & strItemSel & ">" & objRS("Cid_Sigla") & "  " & objRS("Esc_Sigla") & "</Option>"
							objRS.MoveNext
						Wend
						strItemSel = ""
						%>
					</select>
				</td>
			</tr>
			
			<tr class="clsSilver">
				<td width="170px" nowrap><font class="clsObrig">:: </font>Local de Configuração</td>
					<td colspan="3">
						<select name="cboLocalConfig" <%=bbloqueia%>>
							<option value=""></option>
							<%
							set objRS = db.execute("CLA_sp_sel_estacao " & Trim(dblLocalConfig) )
							While not objRS.Eof
								strItemSel = ""
								if Trim(dblLocalConfig) = Trim(objRS("Esc_ID")) then strItemSel = " Selected " End if
								Response.Write "<Option value=" & objRS("Esc_ID") & strItemSel & ">" & objRS("Cid_Sigla") & "  " & objRS("Esc_Sigla") & "</Option>"
								objRS.MoveNext
							Wend
							strItemSel = ""
							%>
						</select>
					</td>
				</tr>
				<tr class="clsSilver">
					<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Contato</td>
					<td width=50% >
						<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width=355px>
							<tr>
								<td class="lightblue">&nbsp;<span id=spnContEndLocalInstala><%=strContEscEntrega%></span></td>
							</tr>
						</table>
					</td>
					<td align=right>Telefone</td>
					<td width=20%>
						<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="80%" >
							<tr><td class="lightblue">&nbsp;<span id=spnTelEndLocalInstala><%=strTelEscEntrega%></span></td></tr>
						</table>
					</td>
				</tr>
			</table>
			<table  border=0 cellspacing="1" cellpadding="0" width="760" >
				<tr class="clsSilver">
					<th colspan=6 >&nbsp;•&nbsp;Coordenação Embratel</th>
				</tr>
				<!--
				<tr class="clsSilver">
					<td width="170px"><font class="clsObrig">:: </font>Órgão de Venda</td>
					<td colspan="5" >
						<select name="cboOrgao" <%=bbloqueia%>>
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
				-->
				<tr class="clsSilver">
					<td width="170px"><font class="clsObrig">:: </font>UserName GIC-N</td>
					<td colspan=5>
<%'JCARTUS@ CH-71284VWE - Solicitações vinculadas com usuários (GIC-N) excluídos da base do CLA					
					if isnull(strUserNameGICN) or trim(strUserNameGICN)="" then 
						bbloqueia_GICN=""
					else
						bbloqueia_GICN="disabled=true"
					end if
%>
					<input type="text" class="text" name="txtGICN"  value="<%=strUserNameGICN%>" <%=bbloqueia_GICN%> maxlength="20" size="20" onblur="ResgatarUserCoordenacao(this)">
					</td>
				</tr>
				<tr class="clsSilver">
					<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Nome GIC-N</td>
					<td>
						<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width=355px >
							<tr>
								<td class="lightblue">&nbsp;<span id=spnNomeGICN><%=strNomeGICN%></span></td>
							</tr>
						</table>
					</td>
					<td align=right >Ramal&nbsp;</td>
					<td colspan=3 align=left>
						<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="100px" >
							<tr>
								<td class="lightblue">&nbsp;
									<span id=spnRamalGICN><%=strRamalGICN%></span>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr class="clsSilver">
					<td width="170px"><font class="clsObrig">:: </font>UserName GIC-L</td>
					<td colspan=5>
						<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="100%" >
							<tr>
								<td class="lightblue">&nbsp;<span id=spnNameGICL><%=strUserNameGICLAtual%></span></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr class="clsSilver">
					<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Nome GIC-L</td>
					<td width=355px>
						<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="100%" >
							<tr>
								<td class="lightblue">&nbsp;<span id=spnNomeGICL><%=strNomeGICL%></span></td>
							</tr>
						</table>
					</td>
					<td align=right >Ramal&nbsp;</td>
					<td colspan=3 align=left >
						<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="100px" >
							<tr>
								<td class="lightblue">&nbsp;<span id=spnRamalGICL><%=strRamalGICL%></span></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<span id=spnGLA>
			<%if (strPropAcessoFisico = "TER" or strPropAcessoFisico = "CLI") then
				Response.Write strHtmlGla
			  End if
			%>
			</span>
			<%if dblSolId <> "" then%>
				<table border=0 cellspacing="1" cellpadding="0" width="760">
					<tr>
						<td>
							<iframe	id		= "IFrmMotivoPend"
					    		name        = "IFrmMotivoPend"
					    		width       = "100%"
					    		height      = "240px"
					    		src			= "../inc/MotivoPendencia.asp?dblSolId=<%=dblSolId%>&dblPedId=<%=dblPedId%>&dblLibera=<%Response.Write Server.HTMLEncode(request("libera"))%>"
					    		frameborder = "0"
					    		scrolling   = "auto"
					    		align       = "left">
							</iFrame>
						</td>
					</tr>
				</table>
			<%else%>
				<br>
			<%end if%>
			<table  border=0 cellspacing="1" cellpadding="0" width="760" >
				<tr valign=middle>
					<td align=center >
						<%if trim(ucase(acao)) = "DES" then ' desativação%>
							<input type="button" class="button" name="btnAlterar" <%=strbtndes%> value=".::Desativar::."  onclick="desativa()" accesskey="I" onmouseover="showtip(this,event,'Alterar a solicitação (Alt+I)');" style="color:darkred;;font-weight:bold;width:180px" <%if strbloqcan = true then%>Disabled<%end if%> >&nbsp;
						<%elseif trim(ucase(acao)) = "CAN" then' cancelamento%>
							<input type="button" class="button" name="btnAlterar" value=".::Cancelar::."  onclick="desativa()" accesskey="I" onmouseover="showtip(this,event,'Alterar a solicitação (Alt+I)');" style="color:darkred;font-weight:bold;width:180px">&nbsp;
						<%else%>
							<input type="button" class="button" name="btnAlterar" value="Alterar"  onclick="AprovarAlteracao()" accesskey="I" onmouseover="showtip(this,event,'Alterar a solicitação (Alt+I)');"<%=bbloqueia%>>&nbsp;
						<%end if %>
			
						<%if Request.Form("hdnPaginaOrig") <> "" then %>
							<input type=button	class="button" name=btnVoltar value=Voltar onclick="VoltarOrigem()" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">&nbsp;
							<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
						<%end if %>
						<input type="button" class="button" name="btnFechar" value=" Fechar " onClick="javascript:window.close()" style="width:100px" accesskey="F" onmouseover="showtip(this,event,'Fechar (Alt+F)');">
					</td>
				</tr>
				<tr>
					<td><font class="clsObrig">:: </font> Campos de preenchimento obrigatório.</td>
				</tr>
				<tr>
					<td><font class="clsObrig">:: </font>Legenda: A - Alfanumérico;  N - Numérico;  L - Letra</td>
				</tr>
			</table>
			<input type=hidden name="hdnStatus" value="<%=dblStsId%>">
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
	<!--Form que envia os dados para gravação-->
	<TABLE border=0>
	<tr>
		<td>
			<form method="post" name="Form4">
				<input type="hidden" name="hdnOriSol_ID" value="<%=strOriSol%>">
				<input type=hidden name=hdnAntAcesso value="<%=strAntAcesso%>">
				<input type=hidden name=hdnCboServico>
				<input type=hidden name=hdnDesigServ>
				<input type=hidden name=hdnAcao>
				<input type=hidden name=hdnTipoAcao value="<%=Request.Form("hdnAcao")%>" >
				<input type=hidden name=hdnXml>
				<input type=hidden name=hdnIdAcessoLogico value="<%=dblIdLogico%>">
				<input type=hidden name=hdnSolId value="<%=dblSolId%>">
				<input type=hidden name=hdn678 value="<%=dblIdLogico%>">
				<input type=hidden name=hdnTipoProcesso value=3>
				<input type=hidden name=hdnOrigem value="<%=strOrigem%>">
				<input type=hidden name=hdnObsSol value="">
				<input type=hidden name=hdnGicL value="<%=strUserNameGICLAtual%>">
				<input type=hidden name=hdnGicN value="<%=strUserNameGICLAtual%>">
				<input type="hidden" name="hdnAprovisiId" value="<%=dblAprovisiId%>">
			</form>
		</td>
	</tr>
</table>
</body>
<%DesconectarCla()%>
</html>