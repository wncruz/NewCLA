<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
'•  @@ Davif - 14/06/05 - RN de Gerenciamento
'	- Sistema			: CLA
'	- Arquivo			: Gerenciamento_Main.asp
'	- Descrição			: Consulta Pedidos Pendentes de Gerenciamento pelo Avaliador

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
	if (f.cboUsuario.value == "" && f.cboEstacao.value == "")
	{
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
		hdnAcao.value = "SEL"
		target = self.name
		action = "Gerenciamento_main.asp?Consulta=1"
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
</script>
<form action="Gerenciamento_main.asp" name="Form1" method="post" onsubmit="return checa(this)">
<input type=hidden name=hdnPedId>
<input type=hidden name=hdnSolId>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnXmlReturn>
<table border=0 cellspacing="1" cellpadding="0" width="760" >
<tr >
	<th colspan=2 ><p align=center>Avaliação de Acesso </p></th>
</tr>
<tr class=clsSilver >
	<td>
		Pendências de
	</td>
	<td >
		<select name="cboUsuario">
			<option value="null">Pendente de Gerenciamento</option>
			<%
			Dim dblUsuIdFac

			Vetor_Campos(1)="adInteger,4,adParamInput," & dblUsuId
			Vetor_Campos(2)="adWChar,3,adParamInput,"
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
					if Trim(dblUsuIdFac) = Trim(objRS("Usu_ID")) then strSel = " Selected " End if
					Response.Write "<Option value=" & objRS("Usu_ID") & strSel & ">" & objRS("Usu_Nome") & "</Option>"
					objRS.MoveNext
				Wend
				strSel = ""
			End if
			%>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td width=25% >Nº Solicitação</td>
	<% dim strSol_Id
	  strSol_Id = Request.Form("txtSolId")
	 %>
	<td><input type="text" name="txtSolId" size=10 class=text value="<%=request("txtSolId")%>" onKeyUp="ValidarTipo(this,0)" maxlength=9></td>
</tr>
<tr class=clsSilver>
	<td>
		&nbsp;
	</td>
	<td>
	   <input type="Checkbox" name="cboStatus" value="DV" <%if request("cboStatus") = "DV" then%>checked<%end if%>> Devolvida para o GIC
	</td>
</tr>
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
Dim strStatus

dblUsuIdFac = Request.Form("cboUsuario")
if Request.ServerVariables("CONTENT_LENGTH") = 0  then
	dblUsuIdFac = dblUsuId
End If

dblEstId = Request.Form("cboEstacao")
strStatus = Request.Form("cboStatus")

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
if strStatus = "" then strStatus = "null" End if

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
Set objRSPag = Nothing
DesconectarCla()
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
		<%if dblUsuIdFac <> "null" then%>
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
Dim intPedId
dim lint

	intQtdRegistro = 0
	
	'SET objRS = Server.CreateObject("ADODB.Recordset")
    Vetor_Campos(1)="adInteger,8,adParamInput,"	& 	dblUsuIdFac
    Vetor_Campos(2)="adWChar,20,adParamInput,"	& 	strLoginRede
    Vetor_Campos(3)="adWChar,20,adParamInput,"	& 	request("cboStatus")
    Vetor_Campos(4)="adInteger,8,adParamInput,"	& 	strSol_Id
	
    strSql = APENDA_PARAMSTRSQL("cla_sp_view_avaliacao ",4,Vetor_Campos)

	call paginarRS(0,strSql)
	intCount=1

	if not objRSPag.Eof and not objRSPag.Bof then
		'Link Xls/Impressão
		strLink =	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
						"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
						"<a href='javascript:TelaImpressao(800,600,""Pedidos Pendentes - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _
						"</table>"

		strHtml = "<table border=0 cellspacing=1 width=760>" & _
						"<tr>" & _
							"<th align=center>+</th>" & _
							"<th>&nbsp;Solicitação ID</th>" & _
							"<th>&nbsp;Cliente</th>" & _
							"<th>&nbsp;Ação</th>" & _
							"<th nowrap>&nbsp;Nº do Contrato</th>" & _
							"<th>&nbsp;Provedor</th>" & _
							"<th>&nbsp;Status Atual</th>" & _
						"</tr>"
		'"<th width=120>&nbsp;Pedido de Acesso</th>" & _
		strXls = strHtml

		intQtdRegistro = 1

		For intIndex = 1 to objRSPag.PageSize
			if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

			strStatusDet = objRSPag("sts_desc")
			
			StrHtml = strHtml & "<tr class='" & strClass & "'>" & _

									"<td ><a href='javascript:DetalharItem(" & objRSPag("Sol_id") & ")'>...&nbsp;</a></td>" & _
									"<td ><a href=""javascript:parent.AvaliarSolicitacao(" & objRSPag("Sol_Id") & ")"">"  & objRSPag("Sol_Id") & "</a></td>" & _
									"<td >" & objRSPag("Cli_Nome") & "</td>" & _
									"<td >" & AcaoPedido(objRSPag("Tprc_ID")) & "</td>" & _
									"<td >" & objRSPag("Acl_NContratoServico") & "</td>" & _
									"<td >" & objRSPag("Pro_Nome") & "</td>" & _
									"<td >" & strStatusDet & "</td>" & _
								"</tr>"
			

				strXls = strXls & "<tr class='" & strClass & "'>" & _
									"<td ></td>" & _
									"<td >" & objRSPag("Sol_id") & "</td>" & _
									"<td >" & objRSPag("Cli_Nome") & "</td>" & _
									"<td >" & AcaoPedido(objRSPag("Tprc_ID")) & "</td>" & _
									"<td >" & objRSPag("Acl_NContratoServico") & "</td>" & _
									"<td >" & objRSPag("Pro_Nome") & "</td>" & _
									"<td >" & strStatusDet & "</td>" & _
								"</tr>"
			intCount = intCount+1
			objRSPag.MoveNext
			if objRSPag.EOF then Exit For
		Next
	End If

	If intQtdRegistro <> "0" Then
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
