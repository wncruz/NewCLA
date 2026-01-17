<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%

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

	function checa(f) 
	{
		if (f.cboEstacao.value == "") 
		{
			alert("Seleção inválida!");
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
			if(cboUF.value != "") {unico = true}

			if(numeroSNOA.value != "") {unico = true}

			if(txtCNLEstacaoFac.value != "") {unico = true}
			if(txtComplEstacaoFac.value != "") {unico = true}

			//<!-- CH-36688SYI  - Inicio -->
			if(cboTpAcao.value != "") {unico = true}
			//<!-- CH-36688SYI  - Fim -->
	

			if 	(txtPedido.value == "DM-"  && txtSolId.value == "" && txtIdFac.value == "" && cboUF.value == "" && numeroSNOA.value == "" && txtCNLEstacaoFac.value == "" && txtComplEstacaoFac.value == "" && cboTpAcao.value == "")
			{
				alert('Obrigatório o preenchimento de pelo menos um campo')
				return
			}

			hdnAcao.value = "SEL"
			target = self.name
			action = "SolicPedSNOA_main.asp?Consulta=1"
			submit()
		}
	}

	function DetalharItem(dblSolId, dblPedId, dblStsID)
	{
		with (document.forms[0])
		{
		
			if (dblStsID == "282" || dblStsID == "283" || dblStsID == "284")
			{
			
				//PopularXml()
				hdnSolId.value = dblSolId
				hdnPedId.value = dblPedId
				hdnStsId.value = dblStsID
				var strNome = "Facilidade" + dblSolId + dblPedId
				var objJanela = window.open()
				objJanela.name = strNome
				target = strNome

				//target = window.top.name
				action = "facilidade_new_cns.asp"
				submit()
			}
			else
			{
				PopularXml()
				hdnSolId.value = dblSolId
				DetalharFac()
			}

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
	
	<td>

		<form action="facilidade_new_main.asp" name="Form1" method="post" onsubmit="return checa(this)">
			<input type=hidden name=hdnUsuario>
			<input type=hidden name=hdnOrigemEst>
			<input type=hidden name=hdnCNLEstUsu>
			<input type=hidden name=hdnComplEstUsu>
			<input type=hidden name=hdnPedId>
			<input type=hidden name=hdnStsId>
			<input type=hidden name=hdnSolId>
			<input type=hidden name=hdnXmlReturn>
			
			<table border=0 cellspacing="1" cellpadding="0" width="760" >
				<tr >
					<th colspan=2 ><p align=center>Solicitação de Pedido SNOA</p></th>
				</tr>

				<tr class=clsSilver>
					<td>
						UF
					</td>
					<td>
						<select name="cboUF">
							<option value=""></option>
							<%
							dblSigla = Request.Form("cboUF")
							'response.write "<script>alert('"&dblSigla&"')</script>"
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

				<tr class=clsSilver>
					<td>
						Estação
					</td>
					<td>
						<input type="Hidden" name="cboEstacao">
						<input type="text" class="text" name="txtCNLEstacaoFac"  maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	
						onblur="CompletarCampo(this)" TIPO="A">
						&nbsp;
						&nbsp;
						<input type="text" class="text" name="txtComplEstacaoFac"  maxlength="3" size="6" onKeyUp="ValidarTipo(this,7)" 
						onblur="CompletarCampo(this);CheckEstacaoFac(document.Form1.txtCNLEstacaoFac,document.Form1.txtComplEstacaoFac,<%=dblUsuId%>,1);"TIPO="A"> Ex.: SPO IG
					</td>
				</tr>

				<tr class=clsSilver>
					<td width=200px >Pedido de Acesso</td>
					<td>
					<input type="text" class="text" name="txtPedido" value="<%if request("txtPedido") <> "" then response.write ucase(request("txtPedido")) else response.write "DM-" end if%>" maxlength="25" size="20">
					</td>
				</tr>

				<tr class=clsSilver>
					<td width=25% >Nº Solicitação</td>
					<td><input type="text" name="txtSolId" size=10 class=text value="<%=request("txtSolId")%>" onKeyUp="ValidarTipo(this,0)" maxlength=9> <-- Para GPON utilizar somente a solicitação.</td>
				</tr>

				<tr class=clsSilver>
					<td width=25% >Numero do Acesso</td>
					<td><input type="text" name="txtIdFac" size=10  style="WIDTH: 140px" value="<%=request("txtIdFac")%>" class=text onKeyUp="ValidarTipo(this,0)" maxlength=10></td>
				</tr>

				<!-- CH-36688SYI  - Inicio -->
				<tr class=clsSilver>
					<td>
						Tipo de Ação
					</td>
					<td>
						<select name="cboTpAcao" style="WIDTH: 140px">
							<option value=""></option>
							<%
							dblTpAcao = Request.Form("cboTpAcao")
							'response.write "<script>alert('"&dblTpAcao&"')</script>"
							'response.write "<script>alert('"&dblUsuId&"')</script>"
							set objRS = db.execute("CLA_sp_sel_TipoAcaoSNOA ")
							do while not objRS.eof 
							%>
								<option value="<%=objRS("Tprc_Id")%>"
							<%
								if Trim(dblTpAcao) <> "" then
									if Trim(dblTpAcao) = Trim(objRS("Tprc_Id")) then
										response.write "selected"
									end if
								end if
							%>
								><%=objRS("TpAcaoDesc")%></option>
							<%
								objRS.movenext
							loop
							%>
						</select>
					</td>
				</tr>
				<!-- CH-36688SYI  - Fim -->

				<tr class=clsSilver>
					<td width=25% >Número SNOA</td>
					<td>
						<input type="text" class="text" name="numeroSNOA" maxlength="25" size="20">
					</td>
				</tr>


				<tr>
					<td colspan=2 align=center height=35px>
						<input type="button" name="btconsulta" value="Consultar" class="button" onClick="ConsultarPedidosPend()">&nbsp;
						<input type="button" name="btnLimpar" value="Limpar" class="button" onClick="javascript:window.location.replace('SolicPedSNOA_main.asp')">&nbsp;
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

'dblUsuIdFac = Request.Form("cboUsuario")
if Request.ServerVariables("CONTENT_LENGTH") = 0  then
	dblUsuIdFac = dblUsuId
End If

dblEstId = Request.Form("cboEstacao")

Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
if Request.Form("hdnXMLReturn") <> "" then
	objXmlDados.loadXml(Request.Form("hdnXMLReturn"))
	'set objNodeAux = objXmlDados.getElementsByTagName("cboUsuario")
	'if objNodeAux.length > 0 then dblUsuIdFac = objNodeAux(0).childNodes(0).text
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
//setarFocus('cboUsuario')
//-->

function EditarFac(dblSolId,dblPedId, dblStsID)
{
	with (document.forms[0])
	{
		

		<%if (dblUsuIdFac <> "999999999" and not isnull(dblUsuIdFac)) then%> // - 12/04/06 - PSOUTO
			PopularXml()
			hdnSolId.value = dblSolId
			hdnPedId.value = dblPedId
			hdnStsId.value = dblStsID
			var strNome = "Facilidade" + dblSolId + dblPedId
			var objJanela = window.open()
			objJanela.name = strNome
			target = strNome
			//action = "Facilidade_new.asp"
			if (dblStsID == "282" || dblStsID == "283" || dblStsID == "284")
			{
				action = "facilidade_new.asp"
			}else{
				action = "facilidade.asp"
			}
			
			submit()
		<%else%>
			hdnSolId.value = dblSolId
			hdnPedId.value = dblPedId
			hdnStsId.value = dblStsID
			hdnAcao.value = "AlocacaoGLA"
			target = "IFrmProcesso"
			action = "ProcessoFac.asp"
			submit()
		<%End if%>
	}
}
function ContinuaAlocacao(dblSolId,dblPedId , dblStsID )
{
	with (document.forms[0])
	{
		PopularXml()
		hdnSolId.value = dblSolId
		hdnPedId.value = dblPedId
		hdnStsId.value = dblStsID
		target = self.name
		//action = "Facilidade.asp"
		if (dblStsID == "282" || dblStsID == "283" || dblStsID == "284"  )
		{
			action = "facilidade_new.asp"
		}else{
			action = "facilidade.asp"
		}
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


	'response.write "<script>alert('"&Request.Form("cboUF")&"')</script>"
	'response.write "<script>alert('"&Request.Form("cboTpAcao")&"')</script>"

	Vetor_Campos(1)="adWChar,5,adParamInput," & Request.Form("cboUF")			'@UF'
	Vetor_Campos(2)="adInteger,2,adParamInput," & dblEstId						'@Esc_ID
	Vetor_Campos(3)="adWChar,13,adParamInput," & mid(strPedido,1,2) 			'@Ped_Prefixo
	Vetor_Campos(4)="adWChar,13,adParamInput," & mid(strPedido,4,5) 			'@Ped_Numero
	Vetor_Campos(5)="adWChar,13,adParamInput," & mid(strPedido,10,4) 			'@Ped_Ano
	Vetor_Campos(6)="adInteger,2,adParamInput,"	& Request.form("txtSolId")		'@sol_id
	Vetor_Campos(7)="adInteger,2,adParamInput," & Request.Form("txtIdFac") 		'@Acf_NroAcesso
	Vetor_Campos(8)="adInteger,2,adParamInput," & Request.Form("numeroSNOA") 	'@Snoa_num
	Vetor_Campos(9)="adInteger,2,adParamInput," & dblUsuId						'@UsuID_Logado

	'!-- CH-36688SYI  - Inicio -->
	Vetor_Campos(10)="adInteger,2,adParamInput," & Request.Form("cboTpAcao")	'@tprc_id
	'!-- CH-36688SYI  - Fim -->

	strSql = APENDA_PARAMSTR("CLA_sp_view_pedido_snoa_main",10,Vetor_Campos)

	'response.write "<script>alert('"& strSql  &"')</script>"

	call paginarRS(1,strSql)
	intCount=1

	'on error resume next

	'response.write "<script>alert('"& objRSPag.recordcount &"')</script>"
	
	if not objRSPag.Eof and not objRSPag.Bof then

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

		For intIndex = 1 to objRSPag.PageSize

			intPedId = objRSPag("Ped_Id")

			if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

			if objRSPag("Sol_PossuiAvaliador") = "0" or   isNull(objRSPag("Sol_PossuiAvaliador")) or (objRSPag("Sol_PossuiAvaliador") = "1" and not isNull(objRSPag("Sol_DtFimAvaliacao"))) then

				strStatusDet = objRSPag("Sts_Desc")

				if  objDicCef.Exists("GAT") then
					
					strHtml = strHtml & "<tr class='" & strClass & "'>" & _
									"<td ><a href='javascript:DetalharItem(" & objRSPag("Sol_id") & "," & objRSPag("Ped_ID")  &  "," & objRSPag("STS_ID")  & ")'>...&nbsp;</a></td>" & _
									"<td ><a href='javascript:EditarFac(" & objRSPag("Sol_id") & "," & objRSPag("Ped_ID")  &  "," & objRSPag("STS_ID")  & ")'>" & ucase(objRSPag("Ped_Prefixo") & "-" & right("00000" & objRSPag("Ped_Numero"),5) & "/" & objRSPag("Ped_Ano")) & "</a></td>" & _
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

			End IF

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