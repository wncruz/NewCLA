<%
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
dim Acao			' acao

'Adicionado PRSS - 09/09/2007
Dim bdesbloqueia			'Veriavel de controle para para Desbloquear campos
Dim dblSolAPGId				'IDentificador da Solicitação APG
bbloqueia = "disabled=true"
bdesbloqueia =" "

strOrigem = Server.HTMLEncode(request.form("hdnOEOrigem"))
dblSolAPGId = Trim(Server.HTMLEncode(Request.Form("hdnSolAPGId")))

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

'response.write "<script>alert('"&strOrigem&"')</script>"


'Monta o Xml de Acessos

'for each item in Request.Form
'Response.Write item & " = " & Request.Form(item) & "<BR>"
'next
'Response.End
%>
<script>
window.name = "Edicao"
</script>
<!--#include file="../inc/xmlAcessos.asp"-->
<%
acao = Request.Form("acao")
if acao = "" then
  acao = request.querystring("acao")
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

dblSolId = Trim(Server.HTMLEncode(Request.Form("hdnSolId")))
if dblSolId = "" then dblSolId = Trim(Server.HTMLEncode(Request.QueryString("SolId")))

if dblSolId = "" then
	Response.Write "<script language=javascript>alert('Solicitação indisponível.');window.location.replace('main.asp');</script>"
	Response.End
End if

If trim(strOrigem) = "APG" Then
  Set objRSSolic = db.execute("CLA_sp_sel_tarefas_APG null, null, null, " & dblSolAPGId)
  strOrderEntrySis = StrOrigem
  
  If Not objRSSolic.eof or Not objRSSolic.bof Then
	dblSolId = Trim(objRSSolic("Solicitacao"))

	if isnull(dblSolId) or trim(dblSolId) = "" then
	  dblSolId = "0"
	end if
  End if
End if

if dblSolId <> "" and dblSolId <> "0" then
  Set objRSSolic		= db.execute("CLA_sp_view_solicitacaomin " & dblSolId)
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
  '--> PSOUTO 19/05/2006 PEGAR APENAS ATIVOS
  IF Server.HTMLEncode(request("libera")) = 1 THEN ' LIBERACAO DE ESTOQUE
    Vetor_Campos(5)="adVarchar,1,adParamInput,NULL"  
    Vetor_Campos(6)="adInteger,1,adParamInput,0"
  ELSE
    Vetor_Campos(5)="adVarchar,1,adParamInput,A"  
    Vetor_Campos(6)="adInteger,1,adParamInput,NULL"
  END IF 

  strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_VIEW_PONTO",6,Vetor_Campos)

  Set objRSFis = db.Execute(strSqlRet)
  Set objDicProp = Server.CreateObject("Scripting.Dictionary")

  if Not objRSFis.EOF and not objRSFis.BOF then
	Set objXmlDados = MontarXmlAcesso(objXmlDados,objRSFis,"")

	strXmlAcesso = FormatarXml(objXmlDados)
	intAcesso = 1
  End if
  '**********
  'Dados do Acesso lógico
end if
  if Not objRSSolic.Eof then

	'Cliente
	dblNroSev		= Trim(objRSSolic("Sol_SevSeq"))
	strRazaoSocial	= Trim(objRSSolic("Cli_Nome"))
	strNomeFantasia = Trim(objRSSolic("Cli_NomeFantasia"))
	strContaSev		= Trim(objRSSolic("Cli_CC"))
	strSubContaSev	= Trim(objRSSolic("Cli_SubCC"))

  ''@@ LPEREZ - 17/10/2005
		strGrupo				= Trim(objRSSolic("GCli_ID"))
  ''@@ LP
	'response.write "<script>alert('"&objRSSolic("Sol_OrderEntry")&"')</script>"
	If trim(strOrigem) = "APG" Then
		Set objRSSolic = db.execute("CLA_sp_sel_tarefas_APG null, null, null, " & dblSolAPGId)
		strOrderEntrySis = StrOrigem
		If Not objRSSolic.eof or Not objRSSolic.bof Then
			 strOrderEntryAno	= Trim(objRSSolic("OE_Ano"))
			 strOrderEntryNro	= Trim(objRSSolic("OE_Numero"))
			 strOrderEntryItem	= Trim(objRSSolic("OE_Item"))
			 strOrderEntryNumSis= Trim(objRSSolic("OE_Solicitacao_OEU"))
			 strDesignacaoServico    = Trim(objRSSolic("Designacao_Servico"))
			 strTipoContratoServico	= Trim(objRSSolic("Tipo_Contrato_Cliente"))
			 strNroContrServico		= Trim(objRSSolic("Numero_Contrato_Cliente"))

		     strPovID = objRSSolic("Pov_ID")
		     strProcesso = objRSSolic("Processo")
		     strAcao = objRSSolic("Acao")
		     strSolAcessoID = objRSSolic("Sol_Acesso_ID")
			 
			Set objServ = db.execute("CLA_sp_sel_Servico null,'" & Trim(objRSSolic("Servico")) & "'")

			If Not objServ.Eof Then
				dblSerId	= Trim(objServ("Ser_ID"))
			End If
		End If
	else
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
	end if 

	'Solicitação
	'' @@ LPEREZ - 17/10/2005
	strOriSol					= Trim(objRSSolic("OriSol_ID"))
	strProjEspecial		= Trim(objRSSolic("Sol_IndProjEspecial"))
	''@@ LP

	strDtPedido				= Formatar_Data(Trim(objRSSolic("Sol_Data")))
	dblVelServico			= Trim(objRSSolic("IDVelAcessoLog"))
	strTipoContratoServico	= Trim(objRSSolic("Acl_TipoContratoServico"))
	strNroContrServico		= Trim(objRSSolic("Acl_NContratoServico"))
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
	
	strDesignacaoServico    = Trim(objRSSolic("Acl_DesignacaoServico"))
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
	If trim(strOrigem) <> "APG" and  dblSolId = "" Then
       dblSolId = Trim(Server.HTMLEncode(Request.Form("hdnSolId")))
      if dblSolId = "" then 
	     dblSolId = Trim(Server.HTMLEncode(Request.QueryString("SolId")))
	  end if
	end if	 

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
	Response.Write "<script language=javascript>alert('Solicitação indisponível.');window.location.replace('facilidade_main.asp');</script>"
	Response.End
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
'Response.Write strUserName
'Response.End
if Not 	objRS.Eof And Not objRS.Bof then
	'strUserNameGICLAtual = Replace(Trim(objRS("Usu_Nome")),"'","´")
	'strRamalGICL = Replace(Trim(Cstr("" & objRS("Usu_Ramal"))),"'","´")
	'strNomeGICL = Trim(strUserName)
	
	strNomeGICL = Replace(Trim(objRS("Usu_Nome")),"'","´")
	strRamalGICL = Replace(Trim(Cstr("" & objRS("Usu_Ramal"))),"'","´")
	strUserNameGICLAtual = Trim(strUserName)
End if
%>
<script language='javascript' src="../javascript/solicitacao.js"></script>
<script language='javascript' src="../javascript/xmlObjects.js"></script>
<script language='javascript' src="../javascript/xmlFacObjects.js"></script>
<script language='javascript' src="../javascript/AlteracaoCad.js"></script>
<SCRIPT LANGUAGE=javascript>
<!--
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
var intIndice = <%=intIndice%>

function Message(objXmlRet){
	var intRet = window.showModalDialog('Message.asp',objXmlRet,'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	
	if (intRet != "")
	{
		//Qdo. for processo origem APG 
		if	(document.Form4.hdnOrigem.value == "APG" || document.Form4.hdnOrigem.value == "Aprov")
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
		<%if strAntAcesso <> "S" then%>
		Alteracao()
		<%end if%>
		AtualizarLista()
	}
  }
}
function VoltarOrigem()
{
	with (document.forms[0])
	{
		target = self.name
		action = "<%=Server.HTMLEncode(Request.Form("hdnPaginaOrig"))%>"
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


function EnviarEmailProvedor()
{
	with(document.Form2)
	{
		//alert(hdnIntIndice.value)
		//return(false)
		if (hdnIntIndice.value != "")
		{
		
			var intChave = hdnIntIndice.value
			var intChaveFis = RequestNodeAcesso(objXmlGeral,"Aec_Id",intChave)
			if (intChaveFis == "") intChaveFis = 0
			var objNodeFis = objXmlGeral.selectNodes("//xDados/Acesso/IdFisico[Aec_Id="+intChaveFis+"]")
			if (objNodeFis.length > 0)
			{
				var intAcfId = objNodeFis[0].childNodes[0].text
				var objNodePed = objXmlGeral.selectNodes("//xDados/Acesso/Pedido[Acf_Id="+intAcfId+"]")
				if (objNodePed.length > 0)
				{
					hdnPedId.value = objNodePed[0].childNodes[0].text
				}else{
					alert("Pedido não encotrado.")
					return
				}
			}else{
				alert("Pedido não encotrado.")
				return
			}

			hdnProId.value = cboProvedor.value
		}
		else{
			alert("Selecione um acesso da lista \"Acessos Adicionados\".")
			return
		}
		hdnAcao.value = "EnviarEmailProvedor"
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
}

// LPEREZ - 24/10/2005
/*
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
*/
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
					intRet = window.showModalDialog('ProcessoDesativacao.asp?hdnAcao=Desativacao&hdnIdLog='+hdn678.value+'&hdnAcfId='+hdnAcfId.value+'&hdnTipoProcesso='+hdnTipoProcesso.value+'&hdnGicN='+document.forms[2].txtGICN.value+'&hdnOEOrigem='+document.forms[0].hdnOEOrigem.value+'&hdnSolId='+document.forms[0].hdnSolId.value+'&hdnIdAcessoLogico='+document.forms[0].hdnIdAcessoLogico.value,"","dialogHeight: 300px; dialogWidth: 700px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
					
					try
						{
							objAryRet = intRet.split(",")
							intRet = objAryRet[0]
						}
					catch(e){intRet = 0}
					
					if (parseInt(intRet) == 146 || parseInt(intRet) == 145 || parseInt(intRet) == 117 || parseInt(intRet) == 2 || parseInt(intRet) == 124)
						{
							
							AtualizaDados(objAryRet[1],0)
							
							spnSolId.innerHTML = objAryRet[1]
							
							var strXml = "<?xml version='1.0' encoding='ISO-8859-1'?>"
							strXml = "<root><CLA_RetornoTmp Msg_Id='155' intOrdem='0' Valor='" +  objAryRet[1] + "' Status='0'><CLA_Mensagem Msg_Titulo='Número da Solicitação gerada'/></CLA_RetornoTmp></root>"
							
							var objXmlRet = new ActiveXObject("Microsoft.XMLDOM");
							objXmlRet.loadXML(strXml);
							Message(objXmlRet);
							
						}
				var data = new Date()
				spnpeddt.innerHTML = data.getDate() + '/' + data.getMonth() + '/' + data.getFullYear()
				document.Form3.btnAlterar.disabled = true
				//document.Form3.btnAlterar.style.visibility = 'hidden'
				VoltarOrigem()
				}
				break	
			case "can":
				Mensagem = 'Confirma o cancelamento do serviço?'
				if (confirm(Mensagem))
				{
					
					hdnTipoProcesso.value = 4
					document.forms[0].hdnUsugicN.value = document.forms[2].txtGICN.value
					intRet = window.showModalDialog('ProcessoDesativacao.asp?hdnAcao=Desativacao&hdnIdLog='+hdn678.value+'&hdnAcfId='+hdnAcfId.value+'&hdnTipoProcesso='+hdnTipoProcesso.value+'&hdnGicN='+document.forms[2].txtGICN.value+'&hdnOEOrigem='+document.forms[0].hdnOEOrigem.value+'&hdnSolId='+document.forms[0].hdnSolId.value+'&hdnIdAcessoLogico='+document.forms[0].hdnIdAcessoLogico.value,"","dialogHeight: 300px; dialogWidth: 700px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
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
				//VoltarOrigem()
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

//-->
</SCRIPT>
<form method="post" name="Form1">
<input type="hidden" name="hdnOriSol_ID" value="<%=strOriSol%>">
<input type=hidden name=hdnUsuario>
<input type=hidden name=hdnOrigemEst>
<input type=hidden name=hdnCNLEstUsu>
<input type=hidden name=hdnComplEstUsu>
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

<!-- LPEREZ - 21/10/2005 -->
<input type=hidden name=hdnProjEsp>
<!-- LP -->

<input type=hidden name=hdnIdAcessoLogico value="<%=dblIdLogico%>">
<input type=hidden name=hdnSolId value="<%=dblSolId%>">
<input type=hidden name=hdnPedId >
<input type=hidden name=hdnDtSolicitacao value="<%=strDtPedido%>">
<input type=hidden name=hdnPadraoDesignacao>
<input type=hidden name=hdnCtfcId value="<%=dblCtfcId%>" >
<input type=hidden name=hdnSubAcao>
<input type=hidden name=hdnXmlReturn value="<%=Server.HTMLEncode(Request.Form("hdnXmlReturn"))%>">
<input type=hidden name="hdn678" value =<%=Server.HTMLEncode(Request.Form("hdn678"))%>>
<input type=hidden name="hdnOrigem" value = <%=Server.HTMLEncode(Request.Form("hdnOrigem"))%><%if Server.HTMLEncode(Request.Form("hdnOrigem")) = "" then response.write Lcase(acao) end if%>> <!-- Alterado PRSS: 19/04/2007 - APG-->
<input type=hidden name="hdnOEOrigem" value = <%=Server.HTMLEncode(Request.Form("hdnOEOrigem")%>> <!-- Adicionado PRSS: 19/04/2007 - APG-->
<input type=hidden name=hdnTipoProcesso>
<input type=hidden name=hdnAcfId value=<%=Server.HTMLEncode(Request.Form("hdnAcfId"))%>>
<input type=hidden name=hdnUsugicN value="">

<tr><td>
<%

if strOrigem="APG" and (trim(dblIdLogico) = "" or isnull(dblIdLogico)) then
if (strprocesso = "DES" and stracao = "DES") then

Set objRSPov = db.execute("CLA_sp_Sel_PovoamentoAPG " & strPovID)
If Not objRSPov.eof or Not objRSPov.bof Then
  strPov_Razao = objRSPov("Pov_Razao")
  strPov_CC = objRSPov("Pov_CC")
  strPov_SubCC = objRSPov("Pov_SubCC")
  strPov_DesigServ = objRSPov("Pov_DesigServ")
  strPov_VelServ = objRSPov("Pov_VelServ")
  strPov_ContratoServ = objRSPov("Pov_ContratoServ")
  'strPov_VelFis = objRSPov("Pov_VelFis")
  strPov_ContatoCli = objRSPov("Pov_ContatoCli")
  strPov_Tel = objRSPov("Pov_Tel")
  strPov_CNPJ = objRSPov("Pov_CNPJ")
  strPov_PropEnd = objRSPov("Aec_PropEnd")
  strPov_End = objRSPov("Pov_End")
end if
%>

<%if 1=0 then%> <%'Desabilitado%>
<table cellspacing="1" cellpadding="0" border=0 width="763">
	<th>
		<th nowrap colspan="4" style=" BACKGROUND-COLOR: #A80000">&nbsp;•&nbsp;Informações de Aprovisionamento APG - Dados do Item de OE de Origem para facilitar busca do Acesso Lógico</th>
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
	  <td colspan="4"><%=strPov_End%></td>
	</tr>
	<tr class="clsSilver">
	  <td width="407" colspan="4">Velocidade do Acesso Físico: <%=strPov_VelFis%></td>
	</tr>
	<tr class="clsSilver">
	  <td width="228">Contato Cliente: <%=strPov_ContatoCli%></td>
	  <td width="161">Telefone: <%=strPov_Tel%></td>
	  <td width="169">CNPJ: <%=strPov_CNPJ%></td>
	  <td width="237">Prop. End.: <%=strPov_PropEnd%></td>
	</tr>
<%end if%>
<table cellspacing="1" cellpadding="0" border=0 width="760">
	<tr>
	  <th nowrap colspan="4" style=" BACKGROUND-COLOR: #A80000">&nbsp;•&nbsp;POVOAMENTO SOB DEMANDA: Associação de Serviços 0800 enviado pelo APG com Acesso Lógico cadastrado no sistema CLA</th>
	</tr>
</table>
<table cellspacing="1" cellpadding="0" border=0 width="760">
	<tr class="clsSilver">
	  <td width="517">Informar o Acesso Lógico que será 
		associado ao serviço 0800 e que sofrerá o processo de 
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
		    <input type="hidden" name="hdnAcaoAPG" value="<%=strAcao%>">
			<input type="hidden" name="hdnSolAcessoID" value="<%=strSolAcessoID%>">
			<input type="text" class="text" name="txtNroLogico" value="<%=dblIdLogico%>" maxlength="10" size="11" onkeyup="ValidarTipo(this,0)">&nbsp;&nbsp; &nbsp;
			<input type="button" class="button" name="associarlogico" value="Associar Lógico" title="Associar Acesso Lógico" onClick="AssociarLogico()" tabindex=0 accesskey="A" onmouseover="showtip(this,event,'Associar Lógico (Alt+A)');">
		</td>
	</tr>
</table>
<%end if%>
<%end if%>

<table cellspacing="1" cellpadding="0" border=0 width="760">
		<!-- ALTERADO POR PSOUTO -->
	<tr >
		<th nowrap>&nbsp;•&nbsp;Solicitação de Acesso</th>
		<th >&nbsp;Nº&nbsp;:&nbsp;<span id=spnSolId><%=dblSolId%></Span></th>
		<%
		Set objRS = db.execute("select sol_Id from cla_pedido where Ped_DtEnvioEmail is not null and Sol_ID = " & dblSolId )
		If Not objRS.eof or Not objRS.bof Then
		  var_bloqueia = 1
		else
		  var_bloqueia = 0
		end if
		
		if trim(ucase(acao)) = "CAN" then
		%>
			<th nowrap>&nbsp;Cancelamento de Solicitação</th>
		<%
		elseif trim(ucase(acao)) = "DES" then
		%>
			<th nowrap>&nbsp;Desativação</th>
		<%
		else
		  If trim(strOrigem) <> "APG" Then
		    var_ModifAcesso = 1
		  end if
		%>
			<th nowrap>&nbsp;Modificações de Informações do Acesso</th>
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
			<input type="text" class="text" name="txtNroSev" value="<%=dblNroSev%>" <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%> maxlength="8" size="10" onkeyup="ValidarTipo(this,0)">&nbsp;
			<input type="button" class="button" name="procurarsev" value=" Procurar Sev  " onClick="ResgatarSev()" tabindex=-1 accesskey="P" onmouseover="showtip(this,event,'Procurar uma SEV no sistema SSA (Alt+P)');" <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%end if%>>
		</td>
	</tr>
<!-- ''@@ LPEREZ - 24/10/2005 -->
	<tr>
		<td class="clsSilver" rowspan="2">&nbsp;
			Projeto Especial
		</td>
		<td class="clsSilver" rowspan="2">&nbsp;
			<input type="radio" name="rdoProjEspecial" onClick="javascript:document.Form1.hdnProjEsp.value = 'S';" value="S" <%if strProjEspecial = "S" then%> checked <%end if%> <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%end if%>>&nbsp; Sim
			<input type="radio" name="rdoProjEspecial" onClick="javascript:document.Form1.hdnProjEsp.value = 'N';" value="N" <%if strProjEspecial <> "S" then%> checked <%end if%> <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%end if%>>&nbsp; Não
		</td>
		<td colspan=2 class="clsSilver">
		    &nbsp;&nbsp;Grupo <span align=right>
		<select name="cboGrupo" onChange="CheckGrupo()" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%> >
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
			
<!--	<div id=divOrigemSol style="display:'none';position:absolute;"> -->
		&nbsp;
		Origem Solicitação <span align=right>
<%
if strOrigem = "APG" then
%>
<select name="cboOrigemSol" disabled>
  <option value="4">APG</option>
</select>

<%else%>
				<select name="cboOrigemSol" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%end if%>>
					<option value="" ></option>
					<%
					set os = db.execute("CLA_sp_sel_OrigSolicitacao")
					do while not os.eof
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
						os.movenext
					loop
					%>
				</select>
<%end if%>
<!--	</div>		-->
		</td>
	</tr>
<!-- ''@@ -->
	<tr class="clsSilver">
		<td width="170"><font class="clsObrig">:: </font>Razão Social</td>
		<td colspan="3" >
			<input type="text" class="text" name="txtRazaoSocial"  maxlength="55" size="55" value="<%=strRazaoSocial%>" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%> onblur="ResgatarGLA()">
			<input type="button" class="button" name="btnProcuraCli" value="Procurar" onClick="ProcurarCliente()" tabindex=-1 accesskey="C" onmouseover="showtip(this,event,'Procurar um cliente no CLA (Alt+C)');" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
			<input type="button" class="button" name="btnNovoCli" value="Limpar Cliente" onClick="NovoCliente()" tabindex=-1 accesskey="Q" onmouseover="showtip(this,event,'Limpar dados do cliente (Alt+Q)');" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170"><span id=spnLabelCliente></span></td>
		<td colspan="3"><span id=spnCliente></span></td>
	</tr>
	<tr class="clsSilver">
		<td width="170"><font class="clsObrig">:: </font>Nome Fantasia</td>
		<td>
			<input type="text" class="text" name="txtNomeFantasia"  maxlength="20" size="25" value="<%=strNomeFantasia%>" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
		</td>
		<td colspan=2 class="clsSilver">&nbsp;</td>
	</tr>
	<tr class="clsSilver">
		<td width="170" ><font class="clsObrig">:: </font>Conta Corrente</td>
		<td width=183>
			<input type=text name=txtContaSev class="text" size=11 maxlength=11 onKeyUp="ValidarTipo(this,0)" onblur="CompletarCampo(this)" TIPO="N" value="<%=strContaSev%>" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
		</td>
		<td align=right width=198 ><font class="clsObrig">:: </font>Sub Conta&nbsp;</td>
		<td width="204" >
			<input type=text name=txtSubContaSev class="text" size=4 maxlength=4 onKeyUp="ValidarTipo(this,0)" onblur="CompletarCampo(this)" TIPO="N" value="<%=strSubContaSev%>" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
		</td>
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
				<select name="cboSistemaOrderEntry" onChange="SistemaOrderEntry(this);hdnOrderEntrySis.value=this.value;" <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
					<Option ></Option>
					<Option value="APG"			<%if strOrderEntrySis = "APG" then Response.Write " selected " End If%>>APG</Option>
					<Option value="CFD"			<%if strOrderEntrySis = "CFD" then Response.Write " selected " End If%>>CFD</Option>
					<Option value="SGA VOZ 0300"			<%if strOrderEntrySis = "SGA VOZ 0300" then Response.Write " selected " End If%>>SGA VOZ 0300</Option>
					<Option value="SGA VOZ 0800 FASE 1"		<%if strOrderEntrySis = "SGA VOZ 0800 FASE 1" then Response.Write " selected " End If%>>SGA VOZ 0800 FASE 1</Option>
					<Option value="SGA VOZ VIP'S"			<%if strOrderEntrySis = "SGA VOZ VIP'S" then Response.Write " selected " End If%>>SGA VOZ VIP'S</Option>
					<Option value="SGA DADOS"	<%if strOrderEntrySis = "SGA DADOS" then Response.Write " selected " End If%>>SGA DADOS</Option>
					<Option value="SGA PLUS"	<%if strOrderEntrySis = "SGA PLUS" then Response.Write " selected " End If%>>SGA PLUS</Option>
					<Option value="ADFAC"		<%if strOrderEntrySis = "ADFAC" then Response.Write " selected " End If%>>ADFAC</Option>
					<Option value="CFM"			<%if strOrderEntrySis = "CFM" then Response.Write " selected " End If%>>CFM</Option>
					<Option value="CFT"			<%if strOrderEntrySis = "CFT" then Response.Write " selected " End If%>>CFT</Option>
				</Select>
			</td>
			<td>-</td>
			<td><input type="text" class="text" onblur="CompletarCampo(this);hdnOrderEntryAno.value=this.value;" onkeyup="ValidarTipo(this,0)" maxlength=4 size=4 name=txtOrderEntry TIPO="N" value="<%=strOrderEntryAno%>" <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>></td>
			<td>-</td>
			<td><input type="text" class="text" onblur="CompletarCampo(this);hdnOrderEntryNro.value=this.value;" onkeyup="ValidarTipo(this,0)" maxlength=7 size=7 name=txtOrderEntry TIPO="N" value="<%=strOrderEntryNro%>" <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>></td>
			<td>-</td>
			<td><input type="text" class="text" onblur="CompletarCampo(this);hdnOrderEntryItem.value=this.value;" onkeyup="ValidarTipo(this,0)" maxlength=3 size=3 name=txtOrderEntry TIPO="N" value="<%=strOrderEntryItem%>" <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Serviço</td>
		<td >
		<%
		'seleciona servico
		set objRS = db.execute("CLA_sp_sel_servico")
		%>
			<select name="cboServicoPedido" onchange="ResgatarServico(this)" <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
			<%
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
		</td>
		<td width="150px" align=right><font class="clsObrig">:: </font>Velocidade&nbsp;</td>
		<td width="200px"><span id=spnVelServico>
				<select name="cboVelServico" onChange="SelVelAcesso(this)" style="width:200px" <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
					<%if Trim(dblSerId) <> "" then
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
					End if%>
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
	
	<%if strOrigem="APG" then%>

	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Designação do Serviço</td>
		<td colspan="3">
				<input type="text" class="text" name="txtdesignacaoServico"
				<%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>
				value="<%=strDesignacaoServico%>" maxlength="22" size="30"><br>
				
		
			

		</td>
	</tr>
    
	<%else%>
	
	   <%if strAntAcesso <> "S" then%>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Designação do Serviço</td>
		<td colspan="3">
			<span id=spnServico <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%end if%>></span>
			<!--Table serviço-->
		</td>
	</tr>
	 <% End if%>
	 <%End if%>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Nº Contrato Serviço</td>
		<td colspan=3>
			<table rules="groups" cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="70%" >
				<tr><td nowrap width=200px >
					<input type=radio name=rdoNroContrato value=1 onClick="spnDescNroContr.innerHTML= 'Ex.: VEM-11 XXX000012003'" checked <%if strTipoContratoServico = "1" then Response.Write " checked " End if%> <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>Contrato de Serviço</td><td></td></tr>
				<tr>
					<td nowrap><input type=radio name=rdoNroContrato onClick="spnDescNroContr.innerHTML= 'Padrão: A22'" value=2 <%if strTipoContratoServico = "2" then Response.Write " checked " End if%> <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>Contrato de Referência</td>
					<td nowrap>
						<input type="text" class="text" name="txtNroContrServico" value="<%=strNroContrServico%>" maxlength="22" size="30" <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>><br>
						<span id=spnDescNroContr>Ex.: VEM-11 XXX00012003</span>
					</td>
				</tr>
				<tr><td nowrap><input type=radio name=rdoNroContrato onClick="spnDescNroContr.innerHTML= 'Padrão: A22'" value=3 <%if strTipoContratoServico = "3" then Response.Write " checked " End if%> <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>Carta de Compromisso</td><td></td></tr>
			</table>
		</td>
	</tr>
	<tr class="clsSilver">
		<td nowrap width=170px><font class="clsObrig">:: </font>Data Desejada de Entrega<br>&nbsp;&nbsp;&nbsp; do Acesso ao Serviço</td>
		<td><input type="text" class="text" name="txtDtEntrAcesServ" value="<%=strDtEntrAcesServ%>" <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%> maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaaa)</td>
		<td nowrap>&nbsp;Data Prevista de Entrega<br>&nbsp;do Acesso pelo Provedor</td>
		<td ><input type="text" class="text" name="txtDtPrevEntrAcesProv" value="<%=strDtPrevEntrAcesProv%>" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%> maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaaa)</td>
	</tr>
	<tr class="clsSilver">
		<td rowspan=2>&nbsp;&nbsp;&nbsp;&nbsp;Acesso Temporário<br>&nbsp;&nbsp;&nbsp;&nbsp;(dd/mm/aaaa)</td>
		<td >&nbsp;Início&nbsp;</td>
		<td >&nbsp;Fim&nbsp;</td>
		<td >&nbsp;Devolução&nbsp;</td>
	</tr>
	<tr class="clsSilver">
		<td ><input type="text" class="text" name="txtDtIniTemp"  value="<%=strDtIniTemp%>" <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%> maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;</td>
		<td ><input type="text" class="text" name="txtDtFimTemp" value="<%=strDtFimTemp%>" <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%> maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;</td>
		<td ><input type="text" class="text" name="txtDtDevolucao" value="<%=strDtDevolucao%>" <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%> maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Designação do Acesso<br>&nbsp;&nbsp;&nbsp; Principal (678)</td>
		<td colspan=3>
			<input type="text" class="text" name="txtDesigAcessoPri0" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>  maxlength="3" size="3" value=678 readOnly>
			<input type="text" class="text" name="txtDesigAcessoPri"  value="<%=dblDesigAcessoPri%>" <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%> maxlength="7" size="9" onKeyUp="ValidarTipo(this,0)" >(678N7)
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
<input type=hidden name=hdnUsuario>
<input type=hidden name=hdnOrigemEst>
<input type=hidden name=hdnCNLEstUsu>
<input type=hidden name=hdnComplEstUsu>
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
<input type=hidden name=hdnstrAcessoTipoRede >
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
<input type=hidden name=hdnVelAcessoFisSel>
<input type=hidden name=hdnSolId value=<%=dblSolId%>>
<input type=hidden name=hdnPedId>
<input type=hidden name=hdnProId>
<input type=hidden name=hdnTipoProcesso value=<%=intTipoProcesso%>>
<input type=hidden name=hdnEstacaoOrigem>
<input type=hidden name=hdnEstacaoDestino>
<input type=hidden name=hdnObrigaGla value="<%=strObrigaGla%>">
<input type=hidden name=hdnUsuID value="<%=dblUsuID %>">


<%'JKNUP: Adicionado%>
<input type=hidden name=hdnTipoTec>
<input type=hidden name=hdnCNLCliente>


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
	<tr>
		<th colSpan=4>&nbsp;•&nbsp;Informações do Acesso&nbsp;</th>
	</tr>
	<tr class="clsSilver">
			<td width=170px ><font class="clsObrig">:: </font>Prop do Acesso Físico</td>
			<td nowrap >
				<input type=radio name=rdoPropAcessoFisico value="TER"	Index=0	<%if strPropAcessoFisico = "TER" then Response.Write " checked " End if%> onclick="EsconderTecnologia(0);ResgatarTecVel()" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>Terceiro&nbsp;&nbsp;&nbsp;
				<input type=radio name=rdoPropAcessoFisico value="EBT"	Index=1	<%if strPropAcessoFisico = "EBT" then Response.Write " checked " End if%> onclick="EsconderTecnologia(0);ResgatarTecVel()" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>EBT&nbsp;&nbsp;&nbsp;
				<input type=radio name=rdoPropAcessoFisico value="CLI"	Index=2	<%if strPropAcessoFisico = "CLI" then Response.Write " checked " End if%> onclick="EsconderTecnologia(0);ResgatarTecVel()" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>Cliente&nbsp;&nbsp;&nbsp;
			<td nowrap colspan=2>
				<%if TipoVel(dblTecId) <> "" then%>
					<div id=divTecnologia style="display:'';POSITION:relative">
				<%Else%>
					<div id=divTecnologia style="display:none;POSITION:relative">
				<%End if%>
				<Select name=cboTecnologia onChange="RetornaCboTipoRadio(this[this.selectedIndex].innerText,this.value,'<% = strTrdID %>', '<% = strVersao %>');ResgatarTecVel()" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
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
			<select name="cboVelAcesso" style="width:150px" onChange="MostrarTipoVel(this)" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
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
			</select></span>&nbsp;&nbsp;<font class="clsObrig">:: </font>Qtde de Acesso(s) Fisico(s)&nbsp;<input type="text" class="text" name="txtQtdeCircuitos" value=1  maxlength="2" size="2" onKeyUp="ValidarTipo(this,0)" value="<%=dblQtdeCircuitos%>" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>&nbsp;&nbsp;
			<%if TipoVel(strTipoVel) <> "" then%>
				<div id=divTipoVel style="display:'';POSITION:absolute">
			<%Else%>
				<div id=divTipoVel style="display:none;POSITION:absolute">
			<%End if%>
			<select name="cboTipoVel" style="width:170px" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
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
			<select name="cboProvedor" onChange="ResgatarPromocaoRegime(this)" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
				<option value=""></option>
				<%	set objRS = db.execute("CLA_sp_sel_provedor 0")
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
			<span id=spnRegimeCntr <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
				<select name="cboRegimeCntr" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
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
			<select name="cboPromocao" style="width:170px" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
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
			<input type="text" class="text" name="txtCodSAP"  maxlength="7" size="10" onKeyUp="ValidarTipo(this,0)" value="<%=strCodSap%>" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>&nbsp;(N7)
		</td>
		<td >&nbsp;&nbsp;&nbsp;Número PI&nbsp;</td>
		<td>
			<input type="text" class="text" name="txtNroPI"  maxlength="7" size="10" onKeyUp="ValidarTipo(this,0)" value="<%=dblNroPI%>" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>&nbsp;(N7)
		</td>
	</tr>
	<tr class=clsSilver2>
		<td width=170px >&nbsp;Endereço Origem&nbsp;</td>
		<td nowrap colspan=3>
			<font class=clsObrig>:: </font>PONTO&nbsp;
			<select name="cboTipoPonto" onChange="TipoOrigem(this.value)" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
				<option value=""></option>
				<option value="I" <%if Trim(strTipoPonto) = "I" then Response.Write " selected " %>>CLIENTE</option>
				<option value="T" <%if Trim(strTipoPonto) = "T" then Response.Write " selected " %>>INTERMEDIÁRIO</option>
			</select>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px nowrap><span id=spnOrigem>&nbsp;&nbsp;&nbsp;Sigla Estação Origem(CNL)</span></td>
		<td colspan=3>
			<input type="text" class="text" name="txtCNLSiglaCentroCli"  maxlength="4" size="8" onKeyUp="ValidarTipo(this,1)"	 onblur="CompletarCampo(this)" TIPO="A" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>&nbsp;Complemento
			<input type="text" class="text" name="txtComplSiglaCentroCli"  maxlength="6" size="10" onKeyUp="ValidarTipo(this,7)" onblur="CompletarCampo(this);ResgatarEstacaoOrigem(document.Form2.txtCNLSiglaCentroCli,document.Form2.txtComplSiglaCentroCli)" TIPO="A" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>&nbsp;
			<!--<input type=button name=btnProcurarEstacaoOrigem class=button value="Procurar" onmouseover="showtip(this,event,'Procurar uma estação de origem(Alt+Y)');" onClick="ResgatarEstacaoOrigem(document.Form2.txtCNLSiglaCentroCli,document.Form2.txtComplSiglaCentroCli)" accesskey="Y" >-->
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>UF</td>
		<td>
			<select name="cboUFEnd" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
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
			<input type=text size=5 maxlength=4 class=text name="txtEndCid" value="<%=strEndCid%>" onBlur="if (ValidarTipo(this,1)){ResgatarCidade(document.forms[1].cboUFEnd,1,this)}" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>&nbsp;
			<input type=text size=27 readonly style="BACKGROUND-COLOR:#eeeeee" class=text name="txtEndCidDesc" value="<%=strEndCidDesc%>" tabIndex=-1 <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Tipo do Logradouro</td>
		<td colspan="0">
			<select name="cboLogrEnd" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
				<option value=""></option>
				<% set objRS = db.execute("CLA_sp_sel_tplogradouro")
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
			<input type="text" class="text" name="txtEnd"  value="<%=strEnd%>" maxlength="60" size="35" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font> Número</td>
		<td>
			<input type="text" class="text" name="txtNroEnd" value="<%=strNroEnd%>" maxlength="5" size="5" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
		</td>
		<td>&nbsp;&nbsp;&nbsp;Complemento</td>
		<td >
			<input type="text" class="text" name="txtComplEnd"  value="<%=strComplEnd%>" maxlength="25" size="25" <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Bairro</td>
		<td>
			<input type="text" class="text" name="txtBairroEnd"  value="<%=strBairroEnd%>" maxlength="30" size="30" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>&nbsp;
		</td>
		<td nowrap><font class="clsObrig">:: </font>CEP&nbsp;(99999-999)</td>
		<td>
			<input type="text" class="text" name="txtCepEnd"  value="<%=strCepEnd%>" maxlength="9" size="10" onKeyPress="OnlyNumbers();AdicionaBarraCep(this)" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>&nbsp;
			<input type=button name=btnProcurarCepInstala value="Procurar CEP" class="button" onclick="ProcurarCEP(1,1)" tabindex=-1 onmouseover="showtip(this,event,'Procurar por CEP exato ou pelos 5 primeiros dígitos (Alt+D)');" accesskey="D" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
		</td>
	</tr>
	<tr class="clsSilver">
		<td colspan=4 align=right><span id=spnCEPSInstala></span></td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Contato</td>
		<td>
			<input type="text" class="text" name="txtContatoEnd" value="<%=strContatoEnd%>" maxlength="30" size="30" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
		</td>
		<td><font class="clsObrig">:: </font>Telefone</td>
		<td >
			<input type="text" class="text" name="txtTelEndArea" maxlength="2" size="2" onkeyUp="ValidarTipo(this,0)" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>&nbsp;
			<input type="text" class="text" name="txtTelEnd" value="<%=strTelEnd%>" maxlength="9" size="10" onkeyUp="ValidarTipo(this,0)" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>CNPJ</td>
		<td colspan="3">
			<input type="text" class="text" name="txtCNPJ"  maxlength="14" size="16" onKeyUp="ValidarTipo(this,0)" value="<%=dblCNPJ%>" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>&nbsp;(99999999999999)
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px>&nbsp;&nbsp;&nbsp;&nbsp;I.E.</td>
		<td >
			<input type="text" class="text" name="txtIE"  maxlength="15" size="17" onKeyUp="ValidarTipo(this,0)" value="<%=strIE%>" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
		</td>
		<td >&nbsp;&nbsp;&nbsp;I.M&nbsp;</td>
		<td>
			<input type="text" class="text" name="txtIM"  maxlength="15" size="17" onKeyUp="ValidarTipo(this,0)" value="<%=strIM%>" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px nowrap>&nbsp;&nbsp;&nbsp;&nbsp;Proprietário do Endereço</td>
		<td colspan="3">
			<input type="text" class="text" name="txtPropEnd"  maxlength="55" size="50" value="<%=strPropEnd%>" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class=clsObrig>:: </font><font color="#FF0000">Interface</font></td>
		<td colspan="3">
			<Select name="cboInterFaceEnd" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%end if%>>
				<Option value=""></Option>
				<%'PRSS - 07/09/2005 
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
					<td><input type="text" class="text" name="txtCNLSiglaCentroCliDest"  maxlength="4" size="8" onKeyUp="ValidarTipo(this,1)"	value="<%=strCNLSiglaCli%>" onblur="CompletarCampo(this)" TIPO="A" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>&nbsp;</td>
					<td>&nbsp;<input type="text" class="text" name="txtComplSiglaCentroCliDest"  maxlength="6" size="10" onKeyUp="ValidarTipo(this,7)" value="<%=strCNLSiglaCli%>" onblur="CompletarCampo(this);ResgatarEstacaoDestino(document.Form2.txtCNLSiglaCentroCliDest,document.Form2.txtComplSiglaCentroCliDest)" TIPO="A" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>&nbsp;</td>
					<!--<td><input type=button name=btnProcurarEstacao class=button value="Procurar" onmouseover="showtip(this,event,'Procurar uma estação(Alt+T)');" onClick="ResgatarEstacaoDestino(document.Form2.txtCNLSiglaCentroCliDest,document.Form2.txtComplSiglaCentroCliDest)" accesskey="T" >&nbsp;</td>-->
					<td>&nbsp;<TEXTAREA rows=2 cols=66 name="txtEndEstacaoEntrega" readonly tabIndex=-1></TEXTAREA></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class=clsObrig>:: </font><font color="#FF0000">Interface</font></td>
		<td colspan="3">
			<Select name="cboInterFaceEndFis" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%end if%> >
				<Option value=""></Option>
				<%'PRSS - 07/09/2005
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
			<input type=button		class=button	name=btnAddAcesso 		style="width:150px" value="Alterar" onmouseover="showtip(this,event,'Atualizar um acesso da lista (Alt+A)');" onClick="AdicionarAcessoListaAlteracao()" accesskey="A" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>&nbsp;
			<!--<input type="button"	class="button"	name="btnEmailPro"		style="width:150px" value="Enviar e-mail para provedor" onclick="EnviarEmailProvedor()" accesskey="M" onmouseover="showtip(this,event,'Enviar email para provedor(Alt+M)');">-->
			<input type=button		class=button	name=btnLimparAcesso	style="width:150px" value="Limpar" onClick="LimparInfoAcesso()" accesskey="L" onmouseover="showtip(this,event,'Limpar dados do Acesso (Alt+L)');" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>&nbsp;
		</td>
	</tr>
</Form>
</table>


<table border=0 cellspacing="1" cellpadding="0" width="760" >
<Form name=Form3 method=Post>
<input type="hidden" name="hdnOriSol_ID" value="<%=strOriSol%>">
<input type=hidden name=hdnUsuario>
<input type=hidden name=hdnOrigemEst>
<input type=hidden name=hdnCNLEstUsu>
<input type=hidden name=hdnComplEstUsu>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnEstacaoAtual>
<input type=hidden name=hdnCtfcId value="<%=dblCtfcId%>" >
<!--<input type=hidden name=hdntxtGICN value="<%=strUserNameGICN%>">-->
<input type=hidden name=hdntxtGICL value="<%=strUserNameGICLAtual%>">
<input type=hidden name=hdnCoordenacaoAtual>
<input type=hidden name=hdnNecessitaRecurso value="S"> <!-- Na ativação será sempre SIM -->
<input type=hidden name=hdnEmiteOTS>
<input type=hidden name=hdnReaproveitarFisico value="N"> <!-- Na ativação será sempre Não. Será modificado na Alteração -->

<%if strOrigem = "APG" then


	Set objRS2 = db.execute("CLA_sp_sel_APG_compartilhamento " & objRSSolic("id_tarefa_Apg"))
	
	if not objRS2.Eof then
	
	strcomp_troncoInterface = objRS2("comp_troncoInterface")
	strcomp_troncoInicio = objRS2("comp_troncoInicio")
	strcomp_troncoFim = objRS2("comp_troncoFim")
	strcomp_rotaCNG = objRS2("comp_rotaCNG")
	strcomp_codServico = objRS2("comp_codServico")
	strcomp_obs = objRS2("comp_obs")
	strcomp_rota = objRS2("comp_rota")
	strcomp_tronco2m = objRS2("comp_tronco2m")
	
	End if

%>
<tr>
		<th colspan="4" >&nbsp;•&nbsp;Informações APG</th>
</tr>
<!--	<tr class="clsSilver">
		<td nowrap width=170px class="clsSilver">&nbsp;&nbsp;&nbsp;&nbsp;Data Pré-Agendamento<br>&nbsp;&nbsp;&nbsp;&nbsp;do Pré-Teste</td>
		<td nowrap class="clsSilver"><input type="text" class="text" name="txtDtEntrAcesServ" value="<%=strDtEntrAcesServ%>" maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaaa)</td>
		<td nowrap class="clsSilver">&nbsp;Data Pré-Agendamento<br>&nbsp;Teste Fim a Fim</td>
		<td nowrap class="clsSilver"><input type="text" class="text" name="txtDtTesteFimaFim" value="<%=strDtTesteFimaFim%>" maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaaa)</td>

	</tr>

-->

	<tr class="clsSilver">
			<td class="clsSilver" title="Quando for necessário a ativação de um novo link de 2Mbps.">&nbsp;&nbsp;&nbsp;
				Emite OTS?
			</td>
			<td <%if strOrigem="APG" and ( strcomp_rota = "S" or strcomp_tronco2m = "S") then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%> class="clsSilver" colspan="3">&nbsp;
				<input  type="radio" name="rdoEmiteOTS" onClick="javascript:document.Form3.hdnEmiteOTS.value = 'S';" value="S" <%if strcomp_rota = "N" or strcomp_tronco2m = "N"then%> checked <%end if%>>&nbsp; Sim
				<input  type="radio" name="rdoEmiteOTS" onClick="javascript:document.Form3.hdnEmiteOTS.value = 'N';" value="N" <%if strcomp_rota = "S" or strcomp_tronco2m = "N" then%> checked <%end if%>>&nbsp; Não
				<%if strOrigem="APG" and strcomp_rota = "S" then%> 
					<% response.write "<script>document.Form3.hdnEmiteOTS.value = 'N';</script>" %>
				<% End if%>
			</td>
	</tr>
	</tr>
		<tr class="clsSilver">
		<td nowrap width=170px class="clsSilver">&nbsp;&nbsp;&nbsp;&nbsp;Username <br>&nbsp;&nbsp;&nbsp;&nbsp;do Cadastrador</td>
		<td nowrap class="clsSilver"><input type="text" class="text" name="txtUsernamecadastrador"
		<%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>
		value="<%=strUsernamecadastrador%>" maxlength="20" size="20" onKeyPress="">&nbsp;</td>
		<td nowrap class="clsSilver">&nbsp;Telefone Cadastrador<br>&nbsp;</td>
		<td nowrap class="clsSilver"><input type="text" class="text" name="txtTelefoneCadastrador"
		<%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>
		value="<%=strTelefoneCadastrador%>" maxlength="10" size="10" onKeyPress="">&nbsp;</td>

	</tr>
<% 'End if
%>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Observações</td>
		<td colspan="3" <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%> ><textarea name="txtObs" onkeydown="MaxLength(this,300);" cols="50" rows="3"><%=strcomp_obs%></textarea></td>
	</tr>
<table border=0 cellspacing="1" cellpadding="0" width="760" >
<tr class=clsSilver2>
		<td rowspan="2">&nbsp;Informações Centro Cliente</td>
</tr>

</table>
<table border=0 cellspacing="1" cellpadding="0" width="760" >	
<tr class="clsSilver">
		<td rowspan="2"><font class="clsObrig">:: </font> Servico</td>
		<td colspan="3" align="center"><font class="clsObrig">:: </font>Tronco</td>
		<td rowspan="2"><font class="clsObrig">:: </font> Compartilha rota com outro CNG?</td>
</tr>
	<tr class="clsSilver">
		<td><font class="clsObrig">:: </font>Interface</td>
		<td><font class="clsObrig">:: </font>Inicio</td>
		<td><font class="clsObrig">:: </font>Fim</td>
	</tr>
	<tr class="clsSilver">
		<td <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>><input type="text" class="text" value = "<%=strcomp_codServico%>"		 name="txtcodServico"		size="20"></td>
		<td <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>><input type="text" class="text" value = "<%=strcomp_rotaCNG%>"		 name="txtrotaCNG"			size="10"></td>
		<td <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>><input type="text" class="text" value = "<%=strcomp_troncoInterface%>" name="txttroncoInterface"  size="10"></td>
		<td <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>><input type="text" class="text" value = "<%=strcomp_troncoInicio%>"	 name="txttroncoInicio"		size="10"></td>
		<td <%if var_bloqueia = 1 then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>><input type="text" class="text" value = "<%=strcomp_troncoFim%>"		 name="txttroncoFim"		size="05"></td>
	</tr>
</table>
<% End if%>
	
		

</table>

<table border=0 cellspacing="1" cellpadding="0" width="760" >
<Form name=Form3 method=Post>
<input type="hidden" name="hdnOriSol_ID" value="<%=strOriSol%>">
<input type=hidden name=hdnUsuario>
<input type=hidden name=hdnOrigemEst>
<input type=hidden name=hdnCNLEstUsu>
<input type=hidden name=hdnComplEstUsu>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnEstacaoAtual>
<input type=hidden name=hdnCtfcId value="<%=dblCtfcId%>" >
<input type=hidden name=hdntxtGICN value="<%=strUserNameGICN%>">
<input type=hidden name=hdnCoordenacaoAtual>
	<tr>
		<th colspan="4" >&nbsp;•&nbsp;Informações da Embratel</th>
	</tr>
	<tr class="clsSilver">
	    <%
		set objRS = db.execute("CLA_sp_sel_estacao " & Trim(dblLocalEntrega))
		%>
		<td width="170px"><font class="clsObrig">:: </font>Local de Entrega<br>&nbsp;&nbsp;&nbsp;do Acesso Lógico</td>
			<input type="Hidden" name="cboLocalEntrega" value="<%=Trim(dblLocalEntrega)%>">
		<td><input type="text" value="<%=Trim(objRS("Cid_Sigla"))%>" class="text" name="txtCNLLocalEntrega"  maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A">&nbsp;
		&nbsp;<input type="text" value="<%=Trim(objRS("Esc_Sigla"))%>" class="text" name="txtComplLocalEntrega"  maxlength="3" size="6" onKeyUp="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoUsuDes(document.Form3.txtCNLLocalEntrega,document.Form3.txtComplLocalEntrega,<%=dblUsuId%>,1);" TIPO="A">
		</td>
		<td colspan="2">&nbsp;</td>
	</tr>
	<%
	set objRS = db.execute("CLA_sp_sel_estacao " & Trim(dblLocalConfig))
	%>
	<tr class="clsSilver">
		<td width="170px" nowrap><font class="clsObrig">:: </font>Local de Configuração</td>
			<input type="Hidden" name="cboLocalConfig" value="<%=Trim(dblLocalConfig)%>">
		<td><input type="text" value="<%=Trim(objRS("Cid_Sigla"))%>" class="text" name="txtCNLLocalConfig"  maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A">&nbsp;
		&nbsp;<input type="text" value="<%=Trim(objRS("Esc_Sigla"))%>" class="text" name="txtComplLocalConfig"  maxlength="3" size="6" onKeyUp="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoUsuDes(document.Form3.txtCNLLocalConfig,document.Form3.txtComplLocalConfig,<%=dblUsuId%>,2);" TIPO="A">
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
<% if not Server.HTMLEncode(request("libera")) = 1 then %>
<table  border=0 cellspacing="1" cellpadding="0" width="760" >
	<tr class="clsSilver">
		<th colspan=6 >&nbsp;•&nbsp;Coordenação Embratel</th>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Órgão de Venda</td>
		<td colspan="5" >
			<select name="cboOrgao" <%if var_bloqueia = 1 and var_ModifAcesso <> 1 then%> <%=bbloqueia%> <%end if%>>
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
		<td colspan=5>
			<%if strUserNameGICN = "" then strUserNameGICN = Ucase(strUserNameGICLAtual) end if%>
			<%if strUserNameGICN = "" then strUserNameGICN = Ucase(strLoginRede) end if%>
			<input type="text" class="text" name="txtGICN"  value="<%=Ucase(strUserNameGICN)%>" maxlength="20" size="20" onblur="ResgatarUserCoordenacao(this)" <%if var_bloqueia = 1 and strUserNameGICN <> "" then%> <%=bbloqueia%> <%end if%>>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Nome GIC-N</td>
		<td >
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
		<td colspan=5>
			<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="100%" >
				<tr><td class="lightblue">&nbsp;
					<span id=spnNameGICL><%=strUserNameGICLAtual%></span>
				</td></tr>
			</table>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Nome GIC-L</td>
		<td width=355px>
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
			<%if (strPropAcessoFisico = "TER" or strPropAcessoFisico = "CLI") then
				Response.Write strHtmlGla
			  End if
			%>
</span>
<% end if %>
<%'Response.Write "libera =" &  request("libera")%>
<table border=0 cellspacing="1" cellpadding="0" width="760">
	<tr>
		<td>
			<iframe	id			= "IFrmMotivoPend"
				    name        = "IFrmMotivoPend"
				    width       = "100%"
				    height      = "120px"
				    src			= "../inc/MotivoPendencia.asp?dblSolId=<%=dblSolId%>&dblPedId=<%=dblPedId%>&dblLibera=<%Response.Write Server.HTMLEncode(request("libera"))%>"
				    frameborder = "0"
				    scrolling   = "no"
				    align       = "left">
			</iFrame>
		</td>
	</tr>
</table>
<table  border=0 cellspacing="1" cellpadding="0" width="760" >
	<tr valign=middle>
		<td align=center >
		
		<% if Server.HTMLEncode(request("libera")) = 1 then %>
																							<!-- xmlFacObjects.js -->
			<input type="button" class="button" name="btnAlterar" value="Enviar"  onclick="EnviarEmailLiberacao('<%=Server.HTMLEncode(request("provedor"))%>','<%=dblSolId%>','<%=strLoginRede%>')" accesskey="I" onmouseover="showtip(this,event,'Alterar a solicitação (Alt+I)');" style="color:darkred;;font-weight:bold;width:180px">&nbsp;
			<input type="button" class="button" name="btnFechar" value=" Fechar " onClick="javascript:window.open('DesativacaoLote.asp','_self')" style="width:100px" accesskey="F" onmouseover="showtip(this,event,'Fechar (Alt+F)');">
		<%else%>
			<% if trim(ucase(acao)) = "DES" then ' desativação%>
			<input type="button" class="button" name="btnAlterar" value=".::Desativar::."  onclick="desativa()" accesskey="I" onmouseover="showtip(this,event,'Alterar a solicitação (Alt+I)');" style="color:darkred;;font-weight:bold;width:180px">&nbsp;
			<%elseif trim(ucase(acao)) = "CAN" then' cancelamento%>
			<input type="button" class="button" name="btnAlterar" value=".::Cancelar::."  onclick="desativa()" accesskey="I" onmouseover="showtip(this,event,'Alterar a solicitação (Alt+I)');" style="color:darkred;font-weight:bold;width:180px">&nbsp;
			<%else%>
			<!--<input type="button" class="button" name="btnAlterar" value="Alterar"  onclick="GravarAlteracao()" accesskey="I" onmouseover="showtip(this,event,'Alterar a solicitação (Alt+I)');">&nbsp;-->
			<input type="button" class="button" name="btnAlterar" value="Alterar"  onclick="AprovarAlteracao()" accesskey="I" onmouseover="showtip(this,event,'Alterar a solicitação (Alt+I)');" <%if strOrigem="APG" and (acao = "DES" or acao = "CAN") and (trim(dblIdLogico) = "" or isnull(dblIdLogico)) then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>&nbsp;
			<%end if %>
			
			<% if Request.Form("hdnPaginaOrig") <> "" then %>
				<input type=button	class="button" name=btnVoltar value=Voltar onclick="VoltarOrigem()" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">&nbsp;
				<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
			<% end if %>
			<input type="button" class="button" name="btnFechar" value=" Fechar " onClick="javascript:window.close()" style="width:100px" accesskey="F" onmouseover="showtip(this,event,'Fechar (Alt+F)');">
		<%end if %>
			
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
<SCRIPT LANGUAGE="JavaScript">
function Alteracao()
{
	with (document.forms[0])
	{
		ResgatarDesigServicoGravado(<%=dblSerId%>)
	}
}
//Geral
</script>
<!--Form que envia os dados para gravação-->
<TABLE border=0>
<tr><td>
<form method="post" name="Form4">
<input type="hidden" name="hdnOriSol_ID" value="<%=strOriSol%>">
<input type=hidden name=hdnUsuario>
<input type=hidden name=hdnOrigemEst>
<input type=hidden name=hdnCNLEstUsu>
<input type=hidden name=hdnComplEstUsu>
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

</form>
</td>
</tr>
</table>
</body>
<%DesconectarCla()%>
</html>