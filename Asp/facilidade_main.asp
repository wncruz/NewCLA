<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Facilidade_Main.asp
'	- Descrição			: Consulta de Pedidos Pendentes

'Response.Write "<br><br>"
'Response.Write "<th colspan=2 ><p align=center><font size=5>Estamos em manutenção previsão 30 Minutos.</font></p></th>"
'Response.end
Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
if Request.Form("hdnXmlReturn") <> "" then
	objXmlDados.loadXml(Request.Form("hdnXmlReturn"))
Else
	objXmlDados.loadXml("<xDados/>") 
End if
%>
<SCRIPT LANGUAGE="JavaScript">
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
function checa(f) {
	if (f.cboUsuario.value == "" && f.cboEstacao.value == "") {
			alert("Seleção inválida!");
			f.cboUsuario.focus();
	    	return false;
		}
	return true;
}

function ConsultarPedidosPend()
{
	with (document.forms[0])
	{
	
	var unico = false
	
	if(txtPedido.value.length > 4){unico = true}
	if(txtSolId.value != "") {unico = true}
	if(txtIdFac.value != "") {unico = true}
	
	if((cboUsuario.value == "")&&(unico == false)){
		alert('Obrigatório o preenchimento do Usuário');
		return
	}
// LPEREZ - 09/03/2006
		if 	(cboUsuario.value == "" && cboEstacao.value == ""   &&
				 txtPedido.value == "DM-"  && txtSolId.value == "" && txtIdFac.value == "" &&
				 txtCliente.value == "" && txtEndereco.value == "" )
		{
			alert('Obrigatório o preenchimento de pelo menos um campo')
			return
		}
//LP
		hdnAcao.value = "SEL"
		target = self.name
		action = "facilidade_main.asp?Consulta=1"
		//action = "facilidade_main.asp"
		submit()
	}
}

function DetalharItem(dblSolId)
{
	with (document.forms[0])
	{
		PopularXml()
		hdnSolId.value = dblSolId
		DetalharFac()
	}
}

function CheckEstacaoFac(objCNL,objCompl,usu,origemEst)
{
	with (document.forms[0])
	{
	
		if (objCNL.value != "" && objCompl.value != "")
		{
			hdnCNLEstUsu.value = objCNL.value
			hdnComplEstUsu.value = objCompl.value
			hdnOrigemEst.value = origemEst
			hdnUsuario.value = usu
			hdnAcao.value = "CheckEstacaoFac"
			target = "IFrmProcesso"
			action = "ProcessoSolic.asp"
			submit()
		}
	}
}
</script>
<tr>
<td >

<form action="facilidade_main.asp" name="Form1" method="post" onsubmit="return checa(this)">
<input type=hidden name=hdnUsuario>
<input type=hidden name=hdnOrigemEst>
<input type=hidden name=hdnCNLEstUsu>
<input type=hidden name=hdnComplEstUsu>
<input type=hidden name=hdnPedId>
<input type=hidden name=hdnSolId>
<input type=hidden name=hdnXmlReturn>
<table border=0 cellspacing="1" cellpadding="0" width="760" >
<tr >
	<th colspan=2 ><p align=center>Alocação de Facilidades</p></th>
</tr>
<tr class=clsSilver>
	<td>
		Pendências de
	</td>
	<td >
		<select name="cboUsuario">
			<option value="" <%if request("cboUsuario")&"" = "" then%>Selected<%end if%>></option>
			<option value="999999999" <%if request("cboUsuario")&"" = "999999999" then%>Selected<%end if%>>Pendente de Alocação</option>
			<%
			Dim dblUsuIdFac

			Vetor_Campos(1)="adInteger,4,adParamInput," & dblUsuId
			Vetor_Campos(2)="adWChar,3,adParamInput, "
			Vetor_Campos(3)="adInteger,4,adParamOutput,0"

			Call APENDA_PARAM("CLA_sp_sel_usuarioCtfcAge",3,Vetor_Campos)
			ObjCmd.Execute'pega dbaction
			DBAction = ObjCmd.Parameters("RET").value
			Set objRS = ObjCmd.Execute()

			dblUsuIdFac = Request("cboUsuario")
			if Request.ServerVariables("CONTENT_LENGTH") = 0  then
				dblUsuIdFac = dblUsuId
			End If
			if dblUsuIdFac = "" then
				set objNode = objXmlDados.getElementsByTagName("cboUsuario")
				if objNode.length > 0 then
					dblUsuIdFac = objNode(0).childNodes(0).text
				End if
			End if

			if DBAction = 0 then
				While not objRS.Eof
					strSel = ""
					'Response.Write "<script language=javascript>alert('" & Trim(objRS("usu_inativo")) & "');</script>"
					if Trim(objRS("usu_inativo")) = "S" then										
					   strSel = ""
					else
						if Trim(dblUsuIdFac) = Trim(objRS("Usu_ID")) then strSel = " Selected " End if
						Response.Write "<Option value=" & objRS("Usu_ID") & strSel & ">" & objRS("Usu_Nome") & "</Option>"
					end if	
					objRS.MoveNext
				Wend
				strSel = ""
			End if
			%>
		</select>
	</td>
</tr>
<!-- Eduardo Araujo 29/08/2007 Otimização -->
<tr class=clsSilver>
	<td>
		Provedor
	</td>
	<td>
		<select name="cboProvedor">
			<option value=""></option>
			<%
			dblProId = Request.Form("cboProvedor")
			set objRS = db.execute("CLA_sp_sel_provedor 0")
			do while not objRS.eof 
			%>
				<option value="<%=objRS("Pro_ID")%>"
			<%
				if Trim(dblProId) <> "" then
					if Trim(dblProId) = Trim(objRS("Pro_ID")) then
						response.write "selected"
					end if
				end if
			%>
				><%=objRS("Pro_Nome")%></option>
			<%
				objRS.movenext
			loop
			%>
		</select>
	</td>
</tr>

<!-- Eduardo Araujo -->
<tr class=clsSilver>
	<td>
		Tecnologia
	</td>
	<td>
		<select name="cboTecnologia">
			<%
			dblTecnologia = Request.Form("cboTecnologia")
			if trim(dblTecnologia) <> "" then 
				if trim(dblTecnologia) = "3" then 
				%>
					<option value="3" selected >ADE</option>
					<option value="8">ADE DSLAM</option>
					<option value="7" >FO EDD</option>
					<option value="6">GPON</option>
					<option value="999">Terceiro</option>
				<%
				end if 
				if trim(dblTecnologia) = "6" then 
				%>
					<option value="3">ADE</option>
					<option value="8">ADE DSLAM</option>
					<option value="7" >FO EDD</option>
					<option value="6" selected >GPON</option>
					<option value="999">Terceiro</option>
				<%
				end if 
				if trim(dblTecnologia) = "7" then 
				%>
					<option value="3">ADE</option>
					<option value="8">ADE DSLAM</option>
					<option value="7" selected >FO EDD</option>
					<option value="6"  >GPON</option>
					<option value="999">Terceiro</option>
				<%
				end if 
				if trim(dblTecnologia) = "8" then 
				%>
					<option value="3">ADE</option>
					<option value="8" selected >ADE DSLAM</option>
					<option value="7"  >FO EDD</option>
					<option value="6"  >GPON</option>
					<option value="999">Terceiro</option>
				<%
				end if
				if trim(dblTecnologia) = "999" then 
				%>
					<option value="3">ADE</option>
					<option value="8">ADE DSLAM</option>
					<option value="7" >FO EDD</option>
					<option value="6">GPON</option>
					<option value="999" selected >Terceiro</option>
				<%
				end if 
			
			else
			%>
			<option value="3">ADE </option>
			<option value="8">ADE DSLAM</option>
			<option value="7" >FO EDD</option>
			<option value="6">GPON</option>
			<option value="999">Terceiro</option>
			<%
			end if
			%>
		</select>
	</td>
</tr>
<!-- Eduardo Araujo -->
<tr class=clsSilver>
	<td>
		UF
	</td>
	<td>
		<select name="cboUF">
			<option value=""></option>
			<%
			dblSigla = Request.Form("cboUF")
			'response.write "<script>alert('"&dblUsuId&"')</script>"
			set objRS = db.execute("CLA_sp_sel_UsuarioCtfcEstado " & dblUsuId )
			do while not objRS.eof 
			%>
				
				<option value="<%=objRS("Est_Sigla")%>"
			<%
				if Trim(dblSigla) <> "" then
					if Trim(dblSigla) = Trim(objRS("Est_Sigla")) then
						response.write "selected"
					end if
				end if
			%>
				><%=objRS("Est_Sigla")%></option>
			<%
				objRS.movenext
			loop
			%>
		</select>
	</td>
</tr>

<!-- Fim da Otimização 29/08/2007 -->
<!-- LPEREZ 13/12/2005 -->
<tr class=clsSilver>
	<td>
		Estação
	</td>
	<td>
		<input type="Hidden" name="cboEstacao">
		<input type="text" class="text" name="txtCNLEstacaoFac"  maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A">&nbsp;
		&nbsp;<input type="text" class="text" name="txtComplEstacaoFac"  maxlength="3" size="6" onKeyUp="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoFac(document.Form1.txtCNLEstacaoFac,document.Form1.txtComplEstacaoFac,<%=dblUsuId%>,1);" TIPO="A">
	Ex.: SPO IG</td>
</tr>
<!--
<tr class=clsSilver>
	<td>
		CNL
	</td>
	<td>
		<select name="cboCNL">
			<option value="">- TODAS -</option>
			<%
				dblCnlId = Request.Form("cboCNL")
				if dblCnlId = "" then
					set objNode = objXmlDados.getElementsByTagName("cboCNL")
					if objNode.length > 0 then
						dblCnlId = objNode(0).childNodes(0).text
					End if
				End if

				Set objRS = db.execute("CLA_sp_sel_usuarioesc '" & dblUsuId &"',5")
				While not objRS.Eof
					strSel = ""
					if Trim(dblCnlId) = Trim(objRS("Cid_Sigla")) then strSel = " Selected " End if
					Response.Write "<Option value='" & objRS("Cid_Sigla") & "'" & strSel & ">" & objRS("Cid_Sigla") & "</Option>"
					objRS.MoveNext
				Wend
			%>
		</select>
	</td>
</tr>
-->
<tr class=clsSilver>
	<td width=200px >Pedido de Acesso</td>
	<td>
	<input type="text" class="text" name="txtPedido" value="<%if request("txtPedido") <> "" then response.write ucase(request("txtPedido")) else response.write "DM-" end if%>" maxlength="25" size="20">
	</td>
</tr>
<tr class=clsSilver>
	<td width=25% >Nº Solicitação</td>
	<td><input type="text" name="txtSolId" size=10 class=text value="<%=request("txtSolId")%>" onKeyUp="ValidarTipo(this,0)" maxlength=9> <-- Para ADE DSLAM / GPON / FO EDD utilizar somente a solicitação.</td>
</tr>
<tr class=clsSilver>
	<td width=25% >Numero do Acesso</td>
	<td><input type="text" name="txtIdFac" size=10  value="<%=request("txtIdFac")%>" class=text onKeyUp="ValidarTipo(this,0)" maxlength=10></td>
</tr>


<!-- @@Davif - Retirar 

<tr class=clsSilver>
	<td >Cliente</td>
	<td ><input type="text" class="text" name="txtCliente" value="<%=request("txtCliente")%>"  maxlength="60" size="50"></td>
</tr>
<tr class=clsSilver>
	<td nowrap>Endereço</td>
	<td nowrap>
		<input type="text" class="text" name="txtEndereco" value="<%=request("txtEndereco")%>" maxlength="60" size="50">&nbsp;Nº&nbsp;
		<input type="text" class="text" name="txtNroEnd" value="<%=request("txtNroEnd")%>" maxlength="10" size="10">&nbsp;
		Compl&nbsp;<input type="text" class="text" name="txtComplemento" value="<%=request("txtComplemento")%>" maxlength="30" size="20">
	</td>
</tr>
<!-- LP -->

<tr>
	<td colspan=2 align=center height=35px>
		<input type="button" name="btconsulta" value="Consultar" class="button" onClick="ConsultarPedidosPend()">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px">
	</td>
</tr>
<tr>
	<td colspan=2 align=center height=35px>
	<iframe	id			= "IFrmProcesso"
			name        = "IFrmProcesso"
			width       = "0"
			height      = "0"
			frameborder = "0"
			scrolling   = "no"
			align       = "left">
	</iFrame>
	</td>
</tr>
</table>
<span id=spnLinks></span>
<%
Dim strClass
Dim dblEstId
Dim intIndex
Dim strSql
Dim intCount
Dim strSel
Dim strXls
Dim strLink
Dim strHtml

dblUsuIdFac = Request.Form("cboUsuario")
if Request.ServerVariables("CONTENT_LENGTH") = 0  then
	dblUsuIdFac = dblUsuId
End If

dblEstId = Request.Form("cboEstacao")

Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
if Request.Form("hdnXMLReturn") <> "" then
	objXmlDados.loadXml(Request.Form("hdnXMLReturn"))
	set objNodeAux = objXmlDados.getElementsByTagName("cboUsuario")
	if objNodeAux.length > 0 then dblUsuIdFac = objNodeAux(0).childNodes(0).text
	set objNodeAux = objXmlDados.getElementsByTagName("cboEstacao")
	if objNodeAux.length > 0 then dblEstId = objNodeAux(0).childNodes(0).text
End if

if dblUsuIdFac = "" then dblUsuIdFac = "null" End if
if dblEstId = "" then dblEstId = "null" End if

'inseri teste para verificar se deve ou não realizar a consulta
if Request.QueryString ("Consulta") = "1" or Request.QueryString ("btn") <> "" then
	strResponse = RetornaTabela
	Response.Write strResponse
	%>
	<!--#include file="../inc/ControlesPaginacao.asp"-->
	<%
end if
%>
<input type=hidden name=hdnXls value="<%=strXls%>">
<input type="hidden" name="hdnPaginaOrig"	value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnNomeCons value="PedidosPend">
</form>
<%

%>
<SCRIPT LANGUAGE=javascript>
<!--
spnLinks.innerHTML = '<%=TratarAspasJS(strLink)%>'
setarFocus('cboUsuario')
//-->
function EditarFac(dblSolId,dblPedId)
{
	with (document.forms[0])
	{

		<%if (dblUsuIdFac <> "999999999" and not isnull(dblUsuIdFac)) then%> // - 12/04/06 - PSOUTO
			PopularXml()
			hdnSolId.value = dblSolId
			hdnPedId.value = dblPedId
			var strNome = "Facilidade" + dblSolId + dblPedId
			var objJanela = window.open()
			objJanela.name = strNome
			target = strNome
			action = "Facilidade.asp"
			submit()
		<%else%>
			hdnSolId.value = dblSolId
			hdnPedId.value = dblPedId
			hdnAcao.value = "AlocacaoGLA"
			target = "IFrmProcesso"
			action = "ProcessoFac.asp"
			submit()
		<%End if%>
	}
}
function ContinuaAlocacao(dblSolId,dblPedId)
{
	with (document.forms[0])
	{
		PopularXml()
		hdnSolId.value = dblSolId
		hdnPedId.value = dblPedId
		target = self.name
		action = "Facilidade.asp"
		submit()
	}
}
</SCRIPT>
</body>
</html>
<%
Set objRSPag = Nothing
DesconectarCla()
%>
<%

function RetornaTabela()
Dim strPedido
strPedido = Request.Form("txtPedido")


	Vetor_Campos(1)="adInteger,2,adParamInput,"	& Request.form("txtSolId")	'@sol_id
	Vetor_Campos(2)="adInteger,2,adParamInput," & Request.form("txtIdLog")	'@Acl_IDAcessoLogico
	Vetor_Campos(3)="adInteger,2,adParamInput," & dblUsuIdFac	'@Usu_ID
	Vetor_Campos(4)="adInteger,2,adParamInput," & dblEstId		'@Esc_ID
	Vetor_Campos(5)="adInteger,2,adParamInput,"	& Request.Form("cboProvedor")'@Pro_ID
	Vetor_Campos(6)="adInteger,2,adParamInput,"					'@Sts_ID
	Vetor_Campos(7)="adInteger,2,adParamInput,"					'@Ped_ID
	Vetor_Campos(8)="adWChar,3,adParamInput,FAC"				'@OrigemChamada
	Vetor_Campos(9)="adWChar,1,adParamInput,P"					'@Agp_Origem
	Vetor_Campos(10)="adWChar,1,adParamInput,A"					'@Situacao
	Vetor_Campos(11)="adInteger,2,adParamInput,"				'@Acf_Id
	Vetor_Campos(12)="adInteger,2,adParamInput," & dblUsuId		'@UsuID_Logado
	Vetor_Campos(13)="adWChar,13,adParamInput," & mid(strPedido,1,2)
	Vetor_Campos(14)="adWChar,13,adParamInput," & mid(strPedido,4,5)
	Vetor_Campos(15)="adWChar,13,adParamInput," & mid(strPedido,10,4)
	Vetor_Campos(16)="adWChar,60,adParamInput,"	& Request.Form("txtCliente")
	Vetor_Campos(17)="adWChar,60,adParamInput,"	& Request.Form("txtEndereco")
	Vetor_Campos(18)="adWChar,10,adParamInput,"	& Request.Form("txtNroEnd")
	Vetor_Campos(19)="adWChar,30,adParamInput," & Request.Form("txtComplemento")
	Vetor_Campos(20)="adInteger,2,adParamInput," 
	' --> psouto 12/05/2006
	'Vetor_Campos(21)="adWChar,30,adParamInput," & Request.Form("cboCNL")'
	Vetor_Campos(21)="adInteger,2,adParamInput," 
	Vetor_Campos(22)="adInteger,2,adParamInput," & Request.Form("txtIdFac")
	
	Vetor_Campos(23)="adInteger,10,adParamInput," & Request.Form("cboTecnologia")
	Vetor_Campos(24)="adWChar,5,adParamInput," & Request.Form("cboUF")
	
	strSql = APENDA_PARAMSTR("CLA_sp_view_pedido_GPON",24,Vetor_Campos)
'if strloginrede="edar" THEN
	'response.write strsql
	'response.end
'END IF
''LP
	call paginarRS(1,strSql)
	intCount=1

	on error resume next
	
	if not objRSPag.Eof and not objRSPag.Bof then
	
		'i = i + 1
		'Response.Write i  
		'Link Xls/Impressão
		strLink =	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
						"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
						"<a href='javascript:TelaImpressao(800,600,""Pedidos Pendentes - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _
						"</table>"

		strHtml = "<table border=0 cellspacing=1 width=760>" & _
						"<tr>" & _
							"<th align=center>+</th>" & _
							"<th width=120>&nbsp;Pedido de Acesso</th>" & _
							"<th>&nbsp;Sol</th>" & _
							"<th>&nbsp;Cliente</th>" & _
							"<th>&nbsp;Ação</th>" & _
							"<th nowrap>&nbsp;Nº do Contrato</th>" & _
							"<th>&nbsp;Provedor</th>" & _
							"<th>&nbsp;Status Atual</th>" & _
						"</tr>"
		strXls = strHtml

		'response.write "<script>alert('"&objRSPag.PageSize&"')</script>"
		For intIndex = 1 to objRSPag.PageSize
        ' Verificar_MErge
	'	if intPedId <> objRSPag("Ped_Id") then 'Usado para controle de exibição de pedidos duplicados (Devido a qtde. de acso. fisicos)
			intPedId = objRSPag("Ped_Id")
			if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

			'Set objRSSts = db.Execute("CLA_sp_sel_StatusSolicitacao null,null,3," & objRSPag("Ped_id"))
			'if Not objRSSts.Eof and Not objRSSts.Bof then strStatusDet = objRSSts("Sts_Desc") else strStatusDet = "" End if
			'@@davif
			
			
			if objRSPag("Sol_PossuiAvaliador") = "0" or   isNull(objRSPag("Sol_PossuiAvaliador")) or (objRSPag("Sol_PossuiAvaliador") = "1" and not isNull(objRSPag("Sol_DtFimAvaliacao"))) then
				strStatusDet = objRSPag("Sts_Desc")
				if (objRSPag("Acf_Proprietario") = "TER" or objRSPag("Acf_Proprietario") = "CLI") and objDicCef.Exists("GAT") then
					strHtml = strHtml & "<tr class='" & strClass & "'>" & _
									"<td ><a href='javascript:DetalharItem(" & objRSPag("Sol_id") & ")'>...&nbsp;</a></td>" & _
									"<td ><a href='javascript:EditarFac(" & objRSPag("Sol_id") & "," & objRSPag("Ped_ID")  &  ")'>" & ucase(objRSPag("Ped_Prefixo") & "-" & right("00000" & objRSPag("Ped_Numero"),5) & "/" & objRSPag("Ped_Ano")) & "</a></td>" & _
									"<td >" & objRSPag("Sol_id") & "</td>" & _
									"<td >" & objRSPag("Cli_Nome") & "</td>" & _
									"<td >" & AcaoPedido(objRSPag("Tprc_ID")) & "</td>" & _
									"<td >" & objRSPag("Acl_NContratoServico") & "</td>" & _
									"<td >" & objRSPag("Pro_Nome") & "</td>" & _
									"<td >" & strStatusDet & "</td>" & _
								"</tr>"

					strXls = strXls & "<tr class='" & strClass & "'>" & _
									"<td ></td>" & _
									"<td >" & ucase(objRSPag("Ped_Prefixo") & "-" & right("00000" & objRSPag("Ped_Numero"),5) & "/" & objRSPag("Ped_Ano")) & "</td>" & _
									"<td >" & objRSPag("Sol_id") & "</td>" & _
									"<td >" & objRSPag("Cli_Nome") & "</td>" & _
									"<td >" & AcaoPedido(objRSPag("Tprc_ID")) & "</td>" & _
									"<td >" & objRSPag("Acl_NContratoServico") & "</td>" & _
									"<td >" & objRSPag("Pro_Nome") & "</td>" & _
									"<td >" & strStatusDet & "</td>" & _
								"</tr>"
				End if

				if objRSPag("Acf_Proprietario") = "EBT" and objDicCef.Exists("GAE") then
					strHtml = strHtml &   "<tr class='" & strClass & "'>" & _
									"<td ><a href='javascript:DetalharItem(" & objRSPag("Sol_id") & ")' >...&nbsp;</a></td>" & _
									"<td ><a href='javascript:EditarFac(" & objRSPag("Sol_id") & "," & objRSPag("Ped_ID")  &  ")'>" & ucase(objRSPag("Ped_Prefixo") & "-" & right("00000" & objRSPag("Ped_Numero"),5) & "/" & objRSPag("Ped_Ano")) & "</a></td>" & _
									"<td >" & objRSPag("Sol_id") & "</td>" & _
									"<td >" & objRSPag("Cli_Nome") & "</td>" & _
									"<td >" & AcaoPedido(objRSPag("Tprc_ID")) & "</td>" & _
									"<td >" & objRSPag("Acl_NContratoServico") & "</td>" & _
									"<td >" & objRSPag("Pro_Nome") & "</td>" & _
									"<td >" & strStatusDet & "</td>" & _
								"</tr>"
					strXls = strXls & "<tr class='" & strClass & "'>" & _
									"<td ></td>" & _
									"<td >" & ucase(objRSPag("Ped_Prefixo") & "-" & right("00000" & objRSPag("Ped_Numero"),5) & "/" & objRSPag("Ped_Ano")) & "</td>" & _
									"<td >" & objRSPag("Sol_id") & "</td>" & _
									"<td >" & objRSPag("Cli_Nome") & "</td>" & _
									"<td >" & AcaoPedido(objRSPag("Tprc_ID")) & "</td>" & _
									"<td >" & objRSPag("Acl_NContratoServico") &"</td>" & _
									"<td >" & objRSPag("Pro_Nome") & "</td>" & _
									"<td >" & strStatusDet & "</td>" & _
								"</tr>"
				End if
			End IF
		'End if

			intCount = intCount+1
			objRSPag.MoveNext
			if objRSPag.EOF then Exit For
		Next
		strHtml = strHtml & "</table>"
		strXls = strXls & "</table>"
		RetornaTabela = strHtml
	Else
		strHtml ="<table width= 760 border= 0 cellspacing= 0 cellpadding= 0 valign=top>"
		strHtml = strHtml + "<tr>"
		strHtml = strHtml + "<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>"
		strHtml = strHtml + "</tr>"
		strHtml = strHtml + "</table>"
		RetornaTabela = strHtml
	End if
End function
%>