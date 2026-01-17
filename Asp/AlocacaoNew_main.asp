<%@ CodePage=65001 %>   
<%
Response.ContentType = "text/html; charset=utf-8"
Response.Charset = "UTF-8"
%>
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
	
		if(txtSolId.value != "") {unico = true}
	

		hdnAcao.value = "SEL"
		target = self.name
		action = "AlocacaoNew_main.asp?Consulta=1"
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

<form action="AlocacaoNew_main.asp" name="Form1" method="post" >
<input type=hidden name=hdnUsuario>
<input type=hidden name=hdnOrigemEst>
<input type=hidden name=hdnCNLEstUsu>
<input type=hidden name=hdnComplEstUsu>
<input type=hidden name=hdnPedId>
<input type=hidden name=hdnSolId>
<input type=hidden name=hdnAcfId>
<input type=hidden name=hdnXmlReturn>
<table border=0 cellspacing="1" cellpadding="0" width="760" >
<tr >
	<th colspan=2 ><p align=center>Alocação \ Aceite</p></th>
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


<tr class=clsSilver>
	<td width=25% >Nº Solicitação</td>
	<td><input type="text" name="txtSolId" size=10 class=text value="<%=Server.HTMLEncode(request("txtSolId"))%>" onKeyUp="ValidarTipo(this,0)" maxlength=9> </td>
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
		action = "AlocacaoNew.asp"
		submit()
	}
}

function ContinuaAlocacao2(dblSolId)
{
	with (document.forms[0])
	{
		PopularXml()
		hdnSolId.value = dblSolId
		target = self.name
		//target = "IFrmProcesso"
		action = "AlocacaoNew.asp"
		submit()
	}
}
function ContinuaAlocacao2Good(dblSolId, dblAcfId)
{
	
	with (document.forms[0])
	{
		PopularXml()
		hdnSolId.value = dblSolId
		hdnAcfId.value = dblAcfId
	
		target = self.name
		//target = "IFrmProcesso"
		action = "AlocacaoNew.asp"
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

	'response.write "<script>alert('"&Request.form("txtSolId")&"')</script>"

	Vetor_Campos(1)="adInteger,2,adParamInput,"	& Request.form("txtSolId")	'@sol_id
	Vetor_Campos(2)="adWChar,5,adParamInput," & Request.Form("cboUF")
	
	strSql = APENDA_PARAMSTR("CLA_sp_view_AlocacaoAceite",2,Vetor_Campos)
	
	'response.write "<script>alert('"&strSql &"')</script>"
'if strloginrede="EDAR" THEN
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
							"<th>&nbsp;Sol</th>" & _
							"<th>&nbsp;Cliente</th>" & _
							"<th>&nbsp;ID Físico</th>" & _
							"<th>&nbsp;Facilidade</th>" & _
							"<th>&nbsp;Tecnologia</th>" & _
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
			'intPedId = objRSPag("Ped_Id")
'******************
' Good início 
'******************		

		  if   Len(cstr(objRSPag("Acf_DtConstrAcessoFis"))) = 0  then
		  
'******************
' Good fim
'******************				
			intPedId = ""
			if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

			
			
			'response.write "<script>alert('"&objRSPag("Sol_id") &"')</script>"
			
			strStatusDet = objRSPag("Sts_Desc")
				

				
					strHtml = strHtml &   "<tr class='" & strClass & "'>" & _
									"<td ><a href='javascript:DetalharItem(" & objRSPag("Sol_id") & ")' >...&nbsp;</a></td>" & _
									"<td ><a href='javascript:ContinuaAlocacao2Good(" & objRSPag("Sol_id") & "," & objRSPag("acf_id") &")'>" & objRSPag("Sol_id") & "</a></td>" & _
									"<td >" & objRSPag("Cli_Nome") & "</td>" & _
									"<td >" & objRSPag("Acf_IDAcessoFisico") & "</td>" & _
									"<td >" & objRSPag("newfac_nome") & "</td>" & _
									"<td >" & objRSPag("newtec_nome") & "</td>" & _
									"<td >" & AcaoPedido(objRSPag("Tprc_ID")) & "</td>" & _
									"<td >" & objRSPag("Acl_NContratoServico") & "</td>" & _
									"<td >" & objRSPag("Pro_Nome") & "</td>" & _
									"<td >" & strStatusDet & "</td>" & _
								"</tr>"
					strXls = strXls & "<tr class='" & strClass & "'>" & _
									"<td ></td>" & _
									"<td >" & objRSPag("Sol_id") & "</td>" & _
									"<td >" & objRSPag("Cli_Nome") & "</td>" & _
									"<td >" & objRSPag("Acf_IDAcessoFisico") & "</td>" & _
									"<td >" & objRSPag("newfac_nome") & "</td>" & _
									"<td >" & objRSPag("newtec_nome") & "</td>" & _
									"<td >" & AcaoPedido(objRSPag("Tprc_ID")) & "</td>" & _
									"<td >" & objRSPag("Acl_NContratoServico") &"</td>" & _
									"<td >" & objRSPag("Pro_Nome") & "</td>" & _
									"<td >" & strStatusDet & "</td>" & _
								"</tr>"
				
		

			intCount = intCount+1
		  end if	
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