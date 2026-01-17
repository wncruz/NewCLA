<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: smbdAltSolicitacao.ASP
'	- Descrição			: Altera uma solicitação no sistema CLA
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/smbdheader.asp"-->
<%
Dim intAno				'Ano
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
Dim objDicProp


Set objDicProp = Server.CreateObject("Scripting.Dictionary") 

'Monta o Xml de Acessos
%>
<!--#include file="../inc/smbdxmlAcessos.asp"-->
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

	dblSolId = Request.QueryString("SolID")
	if dblSolId = "" then 
		Response.Write "<script language=javascript>alert('Solicitação indisponível.');window.location.replace('main.asp');</script>"
		Response.End 
	End if	

	Set objRSSolic = db.execute("CLA_sp_view_solicitacaomin " & dblSolId)

	if objRSSolic.Eof then
		Response.Write "<script language=javascript>alert('Solicitação indisponível.');window.location.replace('main.asp');</script>"
		Response.End 
	End if

	dblIdLogico			= Trim(objRSSolic("Acl_IDAcessoLogico")) 

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

	'Dados do Acesso lógico
	if Not objRSSolic.Eof then

		'Cliente
		dblNroSev		= Trim(objRSSolic("Sol_SevSeq"))
		strRazaoSocial	= Trim(objRSSolic("Cli_Nome"))
		strNomeFantasia = Trim(objRSSolic("Cli_NomeFantasia"))
		strContaSev		= Trim(objRSSolic("Cli_CC"))
		strSubContaSev	= Trim(objRSSolic("Cli_SubCC"))

		if Trim(objRSSolic("Sol_OrderEntry")) <> "" then
			strOrder			= Trim(objRSSolic("Sol_OrderEntry"))
			intTamSis			= len(strOrder)-12
			strOrderEntrySis	= Ucase(Trim(Left(strOrder,intTamSis)))
			intTamSis			= intTamSis + 1
			strOrderEntryAno	= Mid(strOrder,intTamSis,4)
			intTamSis			= intTamSis + 4
			strOrderEntryNro	= Mid(strOrder,intTamSis,5) 
			intTamSis			= intTamSis + 5
			strOrderEntryItem	= Right(strOrder,3)
		End if
		
		'Solicitação
		strDtPedido				= Formatar_Data(Trim(objRSSolic("Sol_Data")))
		dblVelServico			= Trim(objRSSolic("IDVelAcessoLog"))
		strdDesigServico		= Trim(objRSSolic("Acl_DesignacaoServico")) 
		strTipoContratoServico	= Trim(objRSSolic("Acl_TipoContratoServico"))
		strNroContrServico		= Trim(objRSSolic("Acl_NContratoServico")) 
		dblDesigAcessoPriFull	= Trim(objRSSolic("Acl_IDAcessoLogicoPrincipal")) 
		if dblDesigAcessoPriFull <> "" then
			dblDesigAcessoPri		= Right(dblDesigAcessoPriFull,len(dblDesigAcessoPriFull)-3)
		End if	
		strDtEntrAcesServ		= Formatar_Data(Trim(objRSSolic("Acl_DtDesejadaEntregaAcessoServico")))

		strDtIniTemp		= Formatar_Data(Trim(objRSSolic("Acl_DtIniAcessoTemp")))
		strDtFimTemp		= Formatar_Data(Trim(objRSSolic("Acl_DtDevolAcessoTemp")))
		strDtDevolucao		= Formatar_Data(Trim(objRSSolic("Acl_DtDevolAcessoTemp")))

		dblSerId	= Trim(objRSSolic("Ser_ID")) 
		strObsProvedor		= Trim(objRSSolic("Sol_Obs")) 

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
		
		'Usuario de coordenação embratel
		Set objRS = db.execute("CLA_sp_view_agentesolicitacao " & dblSolId)
		
		if Not objRS.Eof then
			While Not objRS.Eof
				Select Case Trim(Ucase(objRS("Age_Desc")))
					Case "GLA"
						strUserNameGLA = Trim(objRS("Usu_Username")) 
					Case "GICN"
						strUserNameGICN = Trim(objRS("Usu_Username")) 
					Case "GICL"
						strUserNameGICL = Trim(objRS("Usu_Username")) 
					Case "GLAE"
						strUserNameGLAE = Trim(objRS("Usu_Username"))
				End Select
				objRS.MoveNext
			Wend	
		End if
		
		dblOrgId = Trim(objRSSolic("Org_id")) 
		dblStsId = Trim(objRSSolic("Sts_id")) 
		strHistoricoSol = Trim(objRSSolic("StsSol_Historico")) 

	End if	
'Else
'	strDtPedido = right("0" & day(now),2) & "/" & right("0" & month(now),2) & "/" & year(now)
'	strDtPrevEntrAcesProv = now() +  30
'	strDtPrevEntrAcesProv = right("0" & day(strDtPrevEntrAcesProv),2) & "/" & right("0" & month(strDtPrevEntrAcesProv),2) & "/" & year(strDtPrevEntrAcesProv)
'End if	
%>
<script language='javascript' src="../javascript/smbdsolicitacao.js"></script>
<script language='javascript' src="../javascript/smbdxmlObjects.js"></script>
<SCRIPT LANGUAGE=javascript>
<!--
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
var objXmlAcessoFisComp = new ActiveXObject("Microsoft.XMLDOM")

objXmlGeral.preserveWhiteSpace = true
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
		if	(document.Form4.hdnTipoProcesso.value == 3)
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
		<%if Trim(Request.Form("hdnAcao")) = "Alteracao" then%>
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
function EditarFac(dblSolId,dblPedId)
{
	with (document.forms[0])
	{
			hdnSolId.value = dblSolId
			hdnPedId.value = dblPedId
			target = self.name
			action = "smbdFacilidade.asp"
			submit()
	}
}
//-->
</SCRIPT>
<form method="post" name="Form1">
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnCboServico>
<input type=hidden name=hdnNomeCbo>
<input type=hidden name=hdnNomeLocal>
<input type=hidden name=hdnUserGICL value="<%=strUserNameGICL%>">
<input type=hidden name=hdnDesigServ>
<input type=hidden name=hdnOrderEntry>
<input type=hidden name=hdnIdEnd>
<input type=hidden name=hdnIdEndInterme>
<input type=hidden name=hdnCNLAtual2>
<input type=hidden name=hdnDesigAcessoPri>
<input type=hidden name=hdnDesigAcessoPriDB value="<%=dblDesigAcessoPriFull%>">

<input type=hidden name=hdnIdAcessoLogico value="<%=dblIdLogico%>">
<input type=hidden name=hdnSolId value="<%=dblSolId%>">
<input type=hidden name=hdnPedId value="<%=dblPedId%>">
<input type=hidden name=hdnDtSolicitacao value="<%=strDtPedido%>">
<input type=hidden name=hdnPadraoDesignacao >
<input type=hidden name=hdnCtfcId value="<%=dblCtfcId%>" >
<input type=hidden name=hdnSubAcao>
<input type=hidden name=hdnXmlReturn value="<%=Request.Form("hdnXmlReturn")%>">

<tr><td>
<table cellspacing="1" cellpadding="0" border=0 width="760">
	<tr >
		<th width=25%>&nbsp;•&nbsp;Solicitação de Acesso</th>
		<th width=25%>&nbsp;Nº&nbsp;:&nbsp;<span id=spnSolId><%=dblSolId%></Span></th>
		<%if Trim(Request.Form("hdnAcao")) = "Alteracao" then%>
			<th width=25%>&nbsp;Tipo&nbsp;:&nbsp;Alteração</th>
			<th width=25%>&nbsp;Acesso Lógico&nbsp;:&nbsp;<%=Request.Form("hdn678")%></th>
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
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Sev </td>
		<td colspan="3">
			<input type="text" class="text" name="txtNroSev" value="<%=dblNroSev%>" maxlength="8" size="10" onkeyup="ValidarTipo(this,0)">&nbsp;
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Razão Social</td>
		<td colspan="3" >
			<input type="text" class="text" name="txtRazaoSocial"  maxlength="55" size="55" value="<%=strRazaoSocial%>" onblur="ResgatarGLA()">
			<input type="button" class="button" name="btnProcuraCli" value="Procurar" onClick="ProcurarCliente()" tabindex=-1 accesskey="C" onmouseover="showtip(this,event,'Procurar um cliente no CLA (Alt+C)');">
			<input type="button" class="button" name="btnNovoCli" value="Limpar Cliente" onClick="NovoCliente()" tabindex=-1 accesskey="Q" onmouseover="showtip(this,event,'Limpar dados do cliente (Alt+Q)');">
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><span id=spnLabelCliente></span></td>
		<td colspan="3"><span id=spnCliente></span></td>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Nome Fantasia</td>
		<td colspan=3>
			<input type="text" class="text" name="txtNomeFantasia"  maxlength="20" size="25" value="<%=strNomeFantasia%>" >
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px" ><font class="clsObrig">:: </font>Conta Corrente</td>
		<td width=25%>
			<input type=text class="text" name=txtContaSev size=11 maxlength=11 onKeyUp="ValidarTipo(this,0)" onblur="CompletarCampo(this)" TIPO="N" value="<%=strContaSev%>">
		</td>
		<td align=right width=10% ><font class="clsObrig">:: </font>Sub Conta&nbsp;</td>
		<td >
			<input type=text name=txtSubContaSev class="text" size=4 maxlength=4 onKeyUp="ValidarTipo(this,0)" onblur="CompletarCampo(this)" TIPO="N" value="<%=strSubContaSev%>">
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
				<select name="cboSistemaOrderEntry" onChange="SistemaOrderEntry(this)" >
					<Option ></Option>
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
			<td><input type="text" class="text" onblur="CompletarCampo(this)" onkeyup="ValidarTipo(this,0)" maxlength=4 size=4 name=txtOrderEntry TIPO="N" value="<%=strOrderEntryAno%>" ></td>
			<td>-</td>
			<td><input type="text" class="text" onblur="CompletarCampo(this)" onkeyup="ValidarTipo(this,0)" maxlength=5 size=5 name=txtOrderEntry TIPO="N" value="<%=strOrderEntryNro%>" ></td>
			<td>-</td>
			<td><input type="text" class="text" onblur="CompletarCampo(this)" onkeyup="ValidarTipo(this,0)" maxlength=3 size=3 name=txtOrderEntry TIPO="N" value="<%=strOrderEntryItem%>" ></td>
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
			<select name="cboServicoPedido" onchange="ResgatarServico(this)">
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
		</td>
		<td width="150px" align=right><font class="clsObrig">:: </font>Velocidade&nbsp;</td>
		<td width="200px"><span id=spnVelServico>
				<select name="cboVelServico" onChange="SelVelAcesso(this)" style="width:200px">
					<option ></option>
					<%if Trim(dblSerId) <> "" then 
						set objRS = db.execute("CLA_sp_sel_AssocServVeloc null," & dblSerId)
						While Not objRS.eof
							strItemSel = ""
							if Trim(dblVelServico) = Trim(objRS("Vel_ID")) then strItemSel = " Selected " End if
							Response.Write "<Option value='" & objRS("Vel_ID") & "'" & strItemSel & ">" & Trim(objRS("Vel_Desc")) & "</Option>"
							objRS.MoveNext
						Wend
						strItemSel = ""
					End if%>
			</span>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Designação do Serviço</td>
		<td colspan="3">
			<span id=spnServico></span>
			<!--Table serviço-->
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp; </td>
		<td colspan="3">
			<input type="text" class="text" name="txtDesigServicoAux" value="<%=strdDesigServico%>" maxlength="100" size="60" Disabled >&nbsp;
		</td>
	</tr>
	
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Nº Contrato Serviço</td>
		<td colspan=3>
			<table rules="groups" cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="70%" >
				<tr><td nowrap width=200px >
					<input type=radio name=rdoNroContrato value=1 onClick="spnDescNroContr.innerHTML= 'Ex.: VEM-11 SSS000012003'" checked <%if strTipoContratoServico = "1" then Response.Write " checked " End if%>>Contrato de Serviço</td><td></td></tr>
				<tr>
					<td nowrap><input type=radio name=rdoNroContrato onClick="spnDescNroContr.innerHTML= 'Padrão: A22'" value=2 <%if strTipoContratoServico = "2" then Response.Write " checked " End if%>>Contrato de Referência</td>
					<td nowrap>
						<input type="text" class="text" name="txtNroContrServico" value="<%=strNroContrServico%>" maxlength="22" size="30"><br>
						<span id=spnDescNroContr>Ex.: VEM-11 SSS000012003</span>
					</td>
				</tr>
				<tr><td nowrap><input type=radio name=rdoNroContrato onClick="spnDescNroContr.innerHTML= 'Padrão: A22'" value=3 <%if strTipoContratoServico = "3" then Response.Write " checked " End if%> >Carta de Compromisso</td><td></td></tr>
			</table>
		</td>
	</tr>
	<tr class="clsSilver">
		<td nowrap width=170px><font class="clsObrig"></font>Data Desejada de Entrega<br>&nbsp;&nbsp;&nbsp; do Acesso ao Serviço</td>
		<td><input type="text" class="text" name="txtDtEntrAcesServ" value="<%=strDtEntrAcesServ%>" maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaaa)</td>
		<td nowrap>&nbsp;Data Prevista de Entrega<br>&nbsp;do Acesso pelo Provedor</td>
		<td ><input type="text" class="text" name="txtDtPrevEntrAcesProv" disabled value="<%=strDtPrevEntrAcesProv%>" maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaaa)</td>
	</tr>
	<tr class="clsSilver">
		<td rowspan=2>&nbsp;&nbsp;&nbsp;&nbsp;Acesso Temporário<br>&nbsp;&nbsp;&nbsp;&nbsp;(dd/mm/aaaa)</td>
		<td >&nbsp;Início&nbsp;</td>
		<td >&nbsp;Fim&nbsp;</td>
		<td >&nbsp;Devolução&nbsp;</td>
	</tr>
	<tr class="clsSilver">
		<td ><input type="text" class="text" name="txtDtIniTemp"  value="<%=strDtIniTemp%>" maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;</td>
		<td ><input type="text" class="text" name="txtDtFimTemp" value="<%=strDtFimTemp%>" maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;</td>
		<td ><input type="text" class="text" name="txtDtDevolucao" value="<%=strDtDevolucao%>" maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Designação do Acesso<br>&nbsp;&nbsp;&nbsp; Principal (678)</td>
		<td colspan=3>
			<input type="text" class="text" name="txtDesigAcessoPri0"  maxlength="3" size="3" value=678 readOnly>
			<input type="text" class="text" name="txtDesigAcessoPri"  value="<%=dblDesigAcessoPri%>" maxlength="7" size="9" onKeyUp="ValidarTipo(this,0)" >(678N7)
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Observações p/ Provedor</td>
		<td colspan="3"><textarea name="txtObs" onkeydown="MaxLength(this,300);" cols="50" rows="3"><%=strObsProvedor%></textarea></td>
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
<input type=hidden name=hdnIntIndice>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnSubAcao>
<input type=hidden name=hdnProvedor>
<input type=hidden name=hdnTipoCEP>
<input type=hidden name=hdnCEP>
<input type=hidden name=hdnCNLNome>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnCNLAtual>
<input type=hidden name=hdnCNLAtual1>
<input type=hidden name=hdnNomeTxtCidDesc>
<input type=hidden name=hdnNomeCboCid>
<input type=hidden name=hdnUserGICL value="<%=strUserNameGICLAtual%>">
<input type=hidden name=hdntxtGICL value="<%=strUserNameGICLAtual%>">
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

<%if Trim(Request.Form("hdnAcao")) = "Alteracao" then%>
<input type=hidden name=hdnTipoProcesso value=3>
<%Else%>
<input type=hidden name=hdnTipoProcesso value=1>
<%End if%>
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
				<input type=radio name=rdoPropAcessoFisico value="TER"	Index=0	<%if strPropAcessoFisico = "TER" then Response.Write " checked " End if%> onclick="EsconderTecnologia(0);ResgatarTecVel()">Terceiro&nbsp;&nbsp;&nbsp;
				<input type=radio name=rdoPropAcessoFisico value="EBT"	Index=1	<%if strPropAcessoFisico = "EBT" then Response.Write " checked " End if%> onclick="EsconderTecnologia(0);ResgatarTecVel()">EBT&nbsp;&nbsp;&nbsp;
				<input type=radio name=rdoPropAcessoFisico value="CLI"	Index=2	<%if strPropAcessoFisico = "CLI" then Response.Write " checked " End if%> onclick="EsconderTecnologia(0);ResgatarTecVel()">Cliente&nbsp;&nbsp;&nbsp;
			<td nowrap colspan=2>
				<%if Trim(Request.Form("hdnAcao")) = "Alteracao" and TipoVel(dblTecId) <> "" then%>
					<div id=divTecnologia style="display:'';POSITION:relative">
				<%Else%>
					<div id=divTecnologia style="display:none;POSITION:relative">
				<%End if%>	
				<Select name=cboTecnologia onChange="ResgatarTecVel()">
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
		<td width=170px><font class="clsObrig">:: </font>Vel do Acesso Físico</td>
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
			</select></span>&nbsp;&nbsp;<font class="clsObrig">:: </font>Qtde de Acesso(s) Fisico(s)&nbsp;<input type="text" class="text" name="txtQtdeCircuitos" value=1  maxlength="2" size="2" onKeyUp="ValidarTipo(this,0)" value="<%=dblQtdeCircuitos%>">&nbsp;&nbsp;
			<div id=divTipoVel style="display:none;POSITION:absolute">
			<select name="cboTipoVel" style="width:170px">
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
			<select name="cboProvedor" onChange="ResgatarPromocaoRegime(this)">
				<option value=""></option>
				<%	set objRS = db.execute("CLA_sp_sel_provedor 0,null,1")
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
		<td width=170px>&nbsp;&nbsp;&nbsp;Promoção</td>
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
		<td width=170px>&nbsp;&nbsp;&nbsp;&nbsp;Código SAP</td>
		<td >
			<input type="text" class="text" name="txtCodSAP"  maxlength="7" size="10" onKeyUp="ValidarTipo(this,0)" value="<%=strCodSap%>" >&nbsp;(N7)
		</td>
		<td >&nbsp;&nbsp;&nbsp;Número PI&nbsp;</td>
		<td>
			<input type="text" class="text" name="txtNroPI"  maxlength="7" size="10" onKeyUp="ValidarTipo(this,0)" value="<%=dblNroPI%>" >&nbsp;(N7)
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
	<tr class="clsSilver">
		<td width=170px nowrap><span id=spnOrigem>&nbsp;&nbsp;&nbsp;Sigla Estação Origem(CNL)</span></td>
		<td colspan=3>
			<input type="text" class="text" name="txtCNLSiglaCentroCli"  maxlength="4" size="8" onKeyUp="ValidarTipo(this,1)"	 onblur="CompletarCampo(this)" TIPO="A">&nbsp;Complemento
			<input type="text" class="text" name="txtComplSiglaCentroCli"  maxlength="6" size="10" onKeyUp="ValidarTipo(this,2)" onblur="CompletarCampo(this);ResgatarEstacaoOrigem(document.Form2.txtCNLSiglaCentroCli,document.Form2.txtComplSiglaCentroCli)" TIPO="A">&nbsp;
			<!--<input type=button name=btnProcurarEstacaoOrigem class=button value="Procurar" onmouseover="showtip(this,event,'Procurar uma estação de origem(Alt+Y)');" onClick="ResgatarEstacaoOrigem(document.Form2.txtCNLSiglaCentroCli,document.Form2.txtComplSiglaCentroCli)" accesskey="Y" >-->
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>UF</td>
		<td>
			<select name="cboUFEnd" >
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
			<input type=text size=5 maxlength=4 class=text name="txtEndCid" value="<%=strEndCid%>" onBlur="if (ValidarTipo(this,1)){ResgatarCidade(document.forms[1].cboUFEnd,1,this)}">&nbsp;
			<input type=text size=27 readonly style="BACKGROUND-COLOR:#eeeeee" class=text name="txtEndCidDesc" value="<%=strEndCidDesc%>" tabIndex=-1>
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Tipo do Logradouro</td>
		<td colspan="0">
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
		</td>
		<td><font class="clsObrig">:: </font>Nome Logr</td>
		<td nowrap>
			<input type="text" class="text" name="txtEnd"  value="<%=strEnd%>" maxlength="60" size="35">
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font> Número</td>
		<td>
			<input type="text" class="text" name="txtNroEnd" value="<%=strNroEnd%>" maxlength="5" size="5">
		</td>
		<td>&nbsp;&nbsp;&nbsp;Complemento</td>
		<td >
			<input type="text" class="text" name="txtComplEnd"  value="<%=strComplEnd%>" maxlength="25" size="25">
		</td>
	</tr>
	<tr class="clsSilver">	
		<td width=170px><font class="clsObrig">:: </font>Bairro</td>
		<td>
			<input type="text" class="text" name="txtBairroEnd"  value="<%=strBairroEnd%>" maxlength="30" size="30">&nbsp;
		</td>
		<td nowrap><font class="clsObrig">:: </font>CEP&nbsp;(99999-999)</td>
		<td>
			<input type="text" class="text" name="txtCepEnd"  value="<%=strCepEnd%>" maxlength="9" size="10" onKeyPress="OnlyNumbers();AdicionaBarraCep(this)" >&nbsp;
			<input type=button name=btnProcurarCepInstala value="Procurar CEP" class="button" onclick="ProcurarCEP(1,1)" tabindex=-1 onmouseover="showtip(this,event,'Procurar por CEP exato ou pelos 5 primeiros dígitos (Alt+D)');" accesskey="D">
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
			<input type="text" class="text" name="txtTelEndArea" maxlength="2" size="2" onkeyUp="ValidarTipo(this,0)">&nbsp;
			<input type="text" class="text" name="txtTelEnd" value="<%=strTelEnd%>" maxlength="8" size="10" onkeyUp="ValidarTipo(this,0)">
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>CNPJ</td>
		<td colspan="3">
			<input type="text" class="text" name="txtCNPJ"  maxlength="14" size="16" onKeyUp="ValidarTipo(this,0)" value="<%=dblCNPJ%>" >&nbsp;(99999999999999)
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px>&nbsp;&nbsp;&nbsp;&nbsp;I.E.</td>
		<td >
			<input type="text" class="text" name="txtIE"  maxlength="15" size="17" onKeyUp="ValidarTipo(this,0)" value="<%=strIE%>" >
		</td>
		<td >&nbsp;&nbsp;&nbsp;I.M&nbsp;</td>
		<td>
			<input type="text" class="text" name="txtIM"  maxlength="15" size="17" onKeyUp="ValidarTipo(this,0)" value="<%=strIM%>" >
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px nowrap>&nbsp;&nbsp;&nbsp;Proprietário do Endereço</td>
		<td colspan="3">
			<input type="text" class="text" name="txtPropEnd"  maxlength="55" size="50" value="<%=strPropEnd%>" >
		</td>
	</tr>
	<tr class="clsSilver">
		<td width=170px><font class=clsObrig>:: </font><font color="#FF0000">Interface</font></td>
		<td colspan="3">
			<Select name="cboInterFaceEnd">
				<Option value=""></Option>
				<Option value="G703" <%if strInterFaceEnd = "G703" then Response.Write " Selected " End if%>>G703</Option>
				<Option value="V24"  <%if strInterFaceEnd = "V24" then Response.Write " Selected " End if%>>V24</Option>
				<Option value="V35"  <%if strInterFaceEnd = "V35" then Response.Write " Selected " End if%>>V35</Option>
				<Option value="V36"  <%if strInterFaceEnd = "V36" then Response.Write " Selected " End if%>>V36</Option>
				<Option value="RS422"<%if strInterFaceEnd = "RS422" then Response.Write " Selected " End if%>>RS422</Option>
				<Option value="RS499"<%if strInterFaceEnd = "RS499" then Response.Write " Selected " End if%>>RS499</Option>
				<Option value="X21"  <%if strInterFaceEnd = "X21" then Response.Write " Selected " End if%>>X21</Option>
				<Option value="LGE"   <%if strInterFaceEnd = "LGE" then Response.Write " Selected " End if%>>LGE</Option>
				<Option value="LGS"   <%if strInterFaceEnd = "LGS" then Response.Write " Selected " End if%>>LGS</Option>
				<Option value="E&M"    <%if strInterFaceEnd = "E&M" then Response.Write " Selected " End if%>>E&M</Option>
				<Option value="RS232"    <%if strInterFaceEnd = "RS232" then Response.Write " Selected " End if%>>RS232</Option>
				<Option value="DSP-3"    <%if strInterFaceEnd = "DSP-3" then Response.Write " Selected " End if%>>DSP-3</Option>
				<Option value="DSP-4"    <%if strInterFaceEnd = "DSP-4" then Response.Write " Selected " End if%>>DSP-4</Option>
				<Option value="DUAE1"    <%if strInterFaceEnd = "DUAE1" then Response.Write " Selected " End if%>>DUAE1</Option>
				<Option value="GFC"    <%if strInterFaceEnd = "GFC" then Response.Write " Selected " End if%>>GFC</Option>				
				<Option value="V11"    <%if strInterFaceEnd = "V11" then Response.Write " Selected " End if%>>V11</Option>				
				<Option value="INDEF"    <%if strInterFaceEnd = "INDEF" then Response.Write " Selected " End if%>>INDEF</Option>				
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
					<td>&nbsp;<input type="text" class="text" name="txtComplSiglaCentroCliDest"  maxlength="6" size="10" onKeyUp="ValidarTipo(this,2)" value="<%=strCNLSiglaCli%>" onblur="CompletarCampo(this);ResgatarEstacaoDestino(document.Form2.txtCNLSiglaCentroCliDest,document.Form2.txtComplSiglaCentroCliDest)" TIPO="A">&nbsp;</td>
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
				<Option value="G703" <%if strInterFaceEnd = "G703" then Response.Write " Selected " End if%>>G703</Option>
				<Option value="V24"  <%if strInterFaceEnd = "V24" then Response.Write " Selected " End if%>>V24</Option>
				<Option value="V35"  <%if strInterFaceEnd = "V35" then Response.Write " Selected " End if%>>V35</Option>
				<Option value="V36"  <%if strInterFaceEnd = "V36" then Response.Write " Selected " End if%>>V36</Option>
				<Option value="RS422"<%if strInterFaceEnd = "RS422" then Response.Write " Selected " End if%>>RS422</Option>
				<Option value="RS499"<%if strInterFaceEnd = "RS499" then Response.Write " Selected " End if%>>RS499</Option>
				<Option value="X21"  <%if strInterFaceEnd = "X21" then Response.Write " Selected " End if%>>X21</Option>
				<Option value="LGE"   <%if strInterFaceEnd = "LGE" then Response.Write " Selected " End if%>>LGE</Option>
				<Option value="LGS"   <%if strInterFaceEnd = "LGS" then Response.Write " Selected " End if%>>LGS</Option>
				<Option value="E&M"    <%if strInterFaceEnd = "E&M" then Response.Write " Selected " End if%>>E&M</Option>
				<Option value="RS232"    <%if strInterFaceEnd = "RS232" then Response.Write " Selected " End if%>>RS232</Option>
				<Option value="DSP-3"    <%if strInterFaceEnd = "DSP-3" then Response.Write " Selected " End if%>>DSP-3</Option>
				<Option value="DSP-4"    <%if strInterFaceEnd = "DSP-4" then Response.Write " Selected " End if%>>DSP-4</Option>
				<Option value="DUAE1"    <%if strInterFaceEnd = "DUAE1" then Response.Write " Selected " End if%>>DUAE1</Option>
				<Option value="GFC"    <%if strInterFaceEnd = "GFC" then Response.Write " Selected " End if%>>GFC</Option>				
				<Option value="V11"    <%if strInterFaceEnd = "V11" then Response.Write " Selected " End if%>>V11</Option>				
				<Option value="INDEF"    <%if strInterFaceEnd = "INDEF" then Response.Write " Selected " End if%>>INDEF</Option>				
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
			<input type=button name=btnAddAcesso class=button value="Alterar" onmouseover="showtip(this,event,'Adicionar/Atualizar um acesso da lista (Alt+A)');" onClick="AtualizarAcessoListaSMBD()" accesskey="A" >&nbsp;
			<span id="spnBtnLimparIdFis1"></span>&nbsp;
			<input type=button name=btnLimparAcesso class=button value="Limpar" onClick="LimparInfoAcesso()" accesskey="L" onmouseover="showtip(this,event,'Limpar dados do Acesso (Alt+L)');">&nbsp;
		</td>
	</tr>
</Form>
</table>
<table border=0 cellspacing="1" cellpadding="0" width="760" >
<Form name=Form3 method=Post>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnEstacaoAtual>
<input type=hidden name=hdnCtfcId value="<%=dblCtfcId%>" >
<input type=hidden name=hdntxtGICN value="<%=strUserNameGICN%>">
<input type=hidden name=hdnCoordenacaoAtual>
	<tr>
		<th colspan="4" >&nbsp;•&nbsp;Informações da Embratel</th>
	</tr>	
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Local de Entrega<br>&nbsp;&nbsp;&nbsp;do Acesso Lógico</td>
		<td colspan="3">
			<select name="cboLocalEntrega" onChange="ResgatarEnderecoEstacao(this);SelecionarLocalConfig(this)">
				<option value=""></option>
				<%set objRS = db.execute("CLA_sp_sel_usuarioesc " & dblUsuId)
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
			<select name="cboLocalConfig" >
				<option value=""></option>
				<%'set objRS = db.execute("CLA_sp_sel_estacao null")
					set objRS = db.execute("CLA_sp_sel_usuarioesc " & dblUsuId)

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
		<th colspan=6 >&nbsp;•&nbsp;Coordenação Embratel</th>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>Órgão de Venda</td>
		<td colspan="5" >
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
		<td colspan=5>
			<input type="text" class="text" name="txtGICN"  value="<%=strUserNameGICN%>" maxlength="20" size="20">
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>UserName GIC-L</td>
		<td colspan=5>
			<input type="text" class="text" name="txtGICL"  value="<%=strUserNameGICL%>" maxlength="20" size="20">
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>UserName GLA</td>
		<td colspan=5>
			<input type="text" class="text" name="txtGLA"  value="<%=strUserNameGLA%>" maxlength="20" size="20">
		</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px"><font class="clsObrig">:: </font>UserName GLA-E</td>
		<td colspan=5>
			<input type="text" class="text" name="txtGLAE"  value="<%=strUserNameGLAE%>" maxlength="20" size="20">
		</td>
	</tr>
</table>
</span>
<table  border=0 cellspacing="1" cellpadding="0" width="760" >
	<tr class="clsSilver">
		<th colspan=6 >&nbsp;•&nbsp;Facilidades (Para trocar facilidade clique sobre o Nro. do Pedido)</th>
	</tr>
		<%
			set objRS = db.execute("[dbo].[CLA_Sp_Sel_PedidoSolicitacao] " & dblSolId)
			if not objRS.Eof then
				While not objRS.Eof 
					if objRS("Prop") <> "EBT" then
		%>
						<tr class="clsSilver">
							<td width="170px"><a href="javascript:EditarFac(<%=dblSolId%>,<%=objRS("Ped_ID")%>)">&nbsp;•&nbsp;<%=objRS("Pedido")%></a></td>
						</tr>
		<%
					End if
					objRS.MoveNext
				Wend
			End if
		%>
		</td>
	</tr>
</table>
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
				<option value="<%=objRS("Sts_id")%>" ><%=ucase(objRS("Sts_Desc"))%>
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
	'Set objRS = db.Execute("CLA_sp_sel_Status null," & dblSolId) 
	Vetor_Campos(1)="adInteger,2,adParamInput,"
	Vetor_Campos(2)="adInteger,2,adParamInput," & dblSolId
	Vetor_Campos(3)="adInteger,2,adParamInput,1"
	Vetor_Campos(4)="adInteger,2,adParamInput,"

	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_StatusSolicitacao",4,Vetor_Campos)
	Set objRS = db.Execute(strSqlRet)

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
			<input type="button" class="button" name="btnGravar" value="Atualizar"  onclick="Gravar()" accesskey="I" onmouseover="showtip(this,event,'Gravar uma solicitação (Alt+I)');" >&nbsp;
			<input type="button" class="button" name="btnFechar" value="  Fechar " onClick="javascript:window.close()" style="width:100px" accesskey="F" onmouseover="showtip(this,event,'Fechar (Alt+F)');">
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
//Geral
with (document.forms[0])
{
	ResgatarDesigServicoGravado(<%=dblSerId%>)
}
</script>
<!--Form que envia os dados para gravação-->
<TABLE border=0> 
<tr><td>
<form method="post" name="Form4">
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnTipoAcao value="<%=Request.Form("hdnAcao")%>" >
<input type=hidden name=hdnXml>
<input type=hidden name=hdnIdAcessoLogico value="<%=dblIdLogico%>">
<input type=hidden name=hdnSolId value="<%=dblSolId%>">
<input type=hidden name=hdn678 value="<%=Request.Form("hdn678")%>">
<%if Trim(Request.Form("hdnAcao")) = "Alteracao" then%>
<input type=hidden name=hdnTipoProcesso value=3>
<%Else%>
<input type=hidden name=hdnTipoProcesso value=1>
<%End if%>
<input type=hidden name=hdnVelIdServicoOld value="<%=dblVelServico%>">
</form>
</td>
</tr>
</table>
</body>
<%DesconectarCla()%>
</html>