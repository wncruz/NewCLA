<!--#include file="../inc/data.asp"-->
<html>
	<head>
		<link rel=stylesheet type="text/css" href="../css/cla.css">
		<script language='javascript' src="../javascript/cla.js"></script>
		<script language='javascript' src="../javascript/claMsg.js"></script>
		<script LANGUAGE=javascript>
			//Envia para tela que detalha o pedido
			function DetalharItem(dblSolId)
			{
				with (document.forms[0])
				{
					hdnSolId.value = dblSolId
					hdnAcao.value = "DetalheSolicitacao"
					target = "DetalheSolic"
					action = "ConsultaGeralDet.asp"
					submit()
				}	
			}
		</script>
	</head>
	<body leftmargin="0" rightmargin="0" topmargin="0" marginwidth="0">
		<form name=Form1 method=Post onSubmit="return false">
			<input type=hidden	name=cboTpAprovacao		  value="1">
			<input type=hidden  name=txtNome   	  	value="<%=Request.Form("txtNome")%>">
			<input type=hidden  name=txtConta   	  value="<%=Request.Form("txtConta")%>">
			<input type=hidden  name=txtQTD   	  value="<%=Request.Form("txtQTD")%>">

			<table border="0" width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td align="right">
						<a name="AbrirXLS" href='javascript:AbrirXls()' style="visibility:hidden" onmouseover="showtip(this,event,'Consulta em formato Excel...')"><img src='../imagens/excel.gif' border=0></a>&nbsp;
						<!--<a name="Imprimir" href='javascript:TelaImpressao(800,600,"Lista de OS Provedor")' style="visibility:hidden" onmouseover="showtip(this,event,'Tela de Impressão...')"><img src='../imagens/impressora.gif' border=0></a>-->
					</td>
				</tr>
			</table>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
Dim strHtmlRet
Dim strXlsRet
Dim objXml
Dim tpProc

'Preenche o vetor de parametros da procedure
Vetor_Campos(1)="adWChar,9,adParamInput, " & Request.Form("txtQTD")
Vetor_Campos(2)="adWChar,60,adParamInput," & Request.Form("txtNome") 
Vetor_Campos(3)="adWChar,15,adParamInput," & Request.Form("txtConta") 

strSql = APENDA_PARAMSTR("CLA_sp_sel_MetaDados",3,Vetor_Campos)

'Chama a procedure
Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then

	'Completa o Header do Html para ser mandando para o IFrame
	strHtmlRet = strHtmlRet & "<table border=0 width=100% cellspacing=1 cellpadding=0>"
	'strHtmlRet = strHtmlRet & "<th nowrap class=clsSilver >"
	strHtmlRet = strHtmlRet & "<th nowrap >&nbsp;Sol</th>"
	strHtmlRet = strHtmlRet & "<th nowrap >&nbsp;Número do Pedido</th>"
	strHtmlRet = strHtmlRet & "<th nowrap >&nbsp;Acesso Lógico</th>"
	strHtmlRet = strHtmlRet & "<th nowrap >&nbsp;Acesso Físico</th>"
	strHtmlRet = strHtmlRet & "<th nowrap >&nbsp;Número do Acesso</th>"
	strHtmlRet = strHtmlRet & "<th nowrap >&nbsp;Data do Pedido</th>"
	strHtmlRet = strHtmlRet & "<th nowrap >&nbsp;Serviço</th>"
	strHtmlRet = strHtmlRet & "<th nowrap >&nbsp;UF</th>"
	
	strHtmlRet = strHtmlRet & "<th nowrap >&nbsp;Velocidade do Serviço</th>"
	strHtmlRet = strHtmlRet & "<th nowrap >&nbsp;Vigência</th>"
	strHtmlRet = strHtmlRet & "<th nowrap >&nbsp;Cliente</th>"
	strHtmlRet = strHtmlRet & "<th nowrap >&nbsp;Conta Corrente</th>"
	strHtmlRet = strHtmlRet & "<th nowrap >&nbsp;Data Solicitacao</th>"
	strHtmlRet = strHtmlRet & "<th nowrap >&nbsp;Tipo do Processo</th>"
	strHtmlRet = strHtmlRet & "<th nowrap >&nbsp;Status Macro</th>"
		
	strHtmlRet = strHtmlRet & "	</tr>"
	
	
	'Início inclusão do cabeçalho da consulta na varável que irá gerar o relatório.
	strXlsRet = strXlsRet & "<table border=0 width=100% cellspacing=1 cellpadding=0>"
	strXlsRet = strXlsRet & "<th nowrap >&nbsp;Sol</th>"
	strXlsRet = strXlsRet & "<th nowrap >&nbsp;Número do Pedido</th>"
	strXlsRet = strXlsRet & "<th nowrap >&nbsp;Acesso Lógico</th>"
	strXlsRet = strXlsRet & "<th nowrap >&nbsp;Acesso Físico</th>"
	strXlsRet = strXlsRet & "<th nowrap >&nbsp;Número do Acesso</th>"
	strXlsRet = strXlsRet & "<th nowrap >&nbsp;Data do Pedido</th>"
	strXlsRet = strXlsRet & "<th nowrap >&nbsp;Serviço</th>"
	strXlsRet = strXlsRet & "<th nowrap >&nbsp;UF</th>"
	
	strXlsRet = strXlsRet & "<th nowrap >&nbsp;Velocidade do Serviço</th>"
	strXlsRet = strXlsRet & "<th nowrap >&nbsp;Vigência</th>"
	strXlsRet = strXlsRet & "<th nowrap >&nbsp;Cliente</th>"
	strXlsRet = strXlsRet & "<th nowrap >&nbsp;Conta Corrente</th>"
	strXlsRet = strXlsRet & "<th nowrap >&nbsp;Data Solicitacao</th>"
	strXlsRet = strXlsRet & "<th nowrap >&nbsp;Tipo do Processo</th>"
	strXlsRet = strXlsRet & "<th nowrap >&nbsp;Status Macro</th>"
	strXlsRet = strXlsRet & "	</tr>"
	'Fim inclusão do cabeçalho da consulta na varável que irá gerar o relatório.
	
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		'Completa o Detail do Html para ser mandando para o IFrame
		strHtmlRet = strHtmlRet & "<tr class="&strClass&">"
		'strHtmlRet = strHtmlRet & "<td nowrap><span id=spnCliente onClick=parent.DetalharItem("&objRSPag("sol_id")&")>" & objRSPag("sol_id") & "</span></td>"
		
		strHtmlRet = strHtmlRet & "<td nowrap><a href='javascript:parent.DetalharItem(" & objRSPag("Sol_ID") & ")' >&nbsp;" & objRSPag("Sol_ID") & "</a></td>"
		strHtmlRet = strHtmlRet & "<td nowrap><span id=spnCliente>" & objRSPag("NumeroDoPedido") & "</span></td>"
		strHtmlRet = strHtmlRet & "<td nowrap><span id=spnCliente>" & objRSPag("Acl_IDAcessoLogico") & "</span></td>"
		strHtmlRet = strHtmlRet & "<td nowrap><span id=spnCliente>" & objRSPag("Acf_IDAcessoFisico") & "</span></td>"
		strHtmlRet = strHtmlRet & "<td nowrap><span id=spnCliente>"& objRSPag("Acf_NroAcessoPtaEbt") & "</span></td>"
		strHtmlRet = strHtmlRet & "<td nowrap><span id=spnCliente>"& objRSPag("Ped_Data") & "</span></td>"
		strHtmlRet = strHtmlRet & "<td nowrap><span id=spnCliente>" & objRSPag("Ser_Desc") & "</span></td>"
		strHtmlRet = strHtmlRet & "<td nowrap><span id=spnCliente>" & objRSPag("Est_Sigla") & "</span></td>"
		strHtmlRet = strHtmlRet & "<td nowrap><span id=spnCliente>" & objRSPag("VEL_DESCLOG") & "</span></td>"
		
		strHtmlRet = strHtmlRet & "<td nowrap><span id=spnCliente>" & objRSPag("Tct_Desc") & "</span></td>"
		strHtmlRet = strHtmlRet & "<td nowrap><span id=spnCliente>" & objRSPag("Cli_Nome") & "</span></td>"
		strHtmlRet = strHtmlRet & "<td nowrap><span id=spnCliente>" & objRSPag("Cli_CC") & "</span></td>"
		strHtmlRet = strHtmlRet & "<td nowrap><span id=spnCliente>" & objRSPag("Sol_Data") & "</span></td>"
		strHtmlRet = strHtmlRet & "<td nowrap><span id=spnCliente>"& objRSPag("Tprc_Des") & "</span></td>"
		strHtmlRet = strHtmlRet & "<td nowrap><span id=spnCliente>"& objRSPag("STS_DESCMAC") & "</span></td>"
				
		strHtmlRet = strHtmlRet & "</tr>"
		
		'strHtmlRet = strHtmlRet & "<input type=hidden name=hdnDesig value='"& trim(objRSPag("Acl_DesignacaoServico")) & "'>"
		'strHtmlRet = strHtmlRet & "<input type=hidden name=hdnAcessoLog value='"& trim(objRSPag("Acl_IDacessoLogico")) & "'>"
		
		'Início inclusão do resultado da consulta na varável que irá gerar o relatório.
		strXlsRet = strXlsRet & "<tr class=clsSilver>"
		strXlsRet = strXlsRet & "<td nowrap>" & objRSPag("sol_id") & "</td>"
		strXlsRet = strXlsRet & "<td nowrap >" & objRSPag("NumeroDoPedido") & "</td>"
		strXlsRet = strXlsRet & "<td nowrap >" & objRSPag("Acl_IDAcessoLogico") & "</td>"
		strXlsRet = strXlsRet & "<td nowrap >" & objRSPag("Acf_IDAcessoFisico") & "</td>"
		strXlsRet = strXlsRet & "<td nowrap >"& objRSPag("Acf_NroAcessoPtaEbt") & "</td>"
		strXlsRet = strXlsRet & "<td nowrap >"& objRSPag("Ped_Data") & "</td>"
		strXlsRet = strXlsRet & "<td nowrap >" & objRSPag("Ser_Desc") & "</td>"
		strXlsRet = strXlsRet & "<td nowrap >" & objRSPag("Est_Sigla") & "</td>"
		strXlsRet = strXlsRet & "<td nowrap >" & objRSPag("VEL_DESCLOG") & "</td>"
		
		strXlsRet = strXlsRet & "<td nowrap>" & objRSPag("Tct_Desc") & "</td>"
		strXlsRet = strXlsRet & "<td nowrap >" & objRSPag("Cli_Nome") & "</td>"
		strXlsRet = strXlsRet & "<td nowrap >" & objRSPag("Cli_CC") & "</td>"
		strXlsRet = strXlsRet & "<td nowrap >" & objRSPag("Sol_Data") & "</td>"
		strXlsRet = strXlsRet & "<td nowrap >"& objRSPag("Tprc_Des") & "</td>"
		strXlsRet = strXlsRet & "<td nowrap >"& objRSPag("STS_DESCMAC") & "</td>"
		
		strXlsRet = strXlsRet & "</tr>"	
		'Fim inclusão do resultado da consulta na varável que irá gerar o relatório.
		
		intCount = intCount+1
		objRSPag.MoveNext
		if objRSPag.EOF then Exit For
	Next
	
	strHtmlRet = strHtmlRet & "</table>"
Else
	strHtmlRet = ""
End if	
'Se nao teve resultado na busca informa na tela, se não mostra o Iframe e os butões de imprimir e excel
if Trim(strHtmlRet) = "" then
	strHtmlRet = strHtmlRet & "<table width=730 border=0 cellspacing=0 cellpadding=0 valign=top>"
	strHtmlRet = strHtmlRet & "<tr>"
	strHtmlRet = strHtmlRet & "<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>"
	strHtmlRet = strHtmlRet & "</tr>"
	strHtmlRet = strHtmlRet & "</table>"
Else
	Response.Write "<script language=javascript>document.getElementById('AbrirXLS').style.visibility = ""visible"";</script>"
	'Response.Write "<script language=javascript>document.getElementById('Imprimir').style.visibility = ""visible"";</script>"
End IF

'Response.Write strHtmlRet
%>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" valign="top">
			<tr>
				<td>
					<input type=hidden name=hdnXls value="<%=strXlsRet%>">
					<input type=hidden name=hdnXls>
					<input type=hidden name=hdnXls>
					<input type=hidden name=hdnXls>
					<input type=hidden name=hdnXls>
					<input type=hidden name=hdnSolId> 
					<input type=hidden name=hdnAcao> 
					<input type="hidden" name="hdstrHtmlRet"  value="<%=strHtmlRet%>">  
					
					<iframe	id			= "IFrmPagina"
						    name        = "IFrmPagina" 
						    width       = "100%"
							height      = "270"
							border		= "0"
							marginwidth = "0"
							marginheight= "0" 
							hspace		= "0"
							vspace		= "0"
							frameborder = "0"
							scrolling   = "overflow" 
							align       = "left">
					</iFrame>
				</td>
			</tr>
			<tr>
				<td><!--#include file="../inc/ControlesPaginacao.asp"--></td>
			</tr>
		</table>
	</form>
</body>
<iframe	id			= "IFrmPagina"
	    name        = "IFrmPagina" 
	    width       = "100%"
	    height      = "0"
		marginwidth ='0' 
		marginheight='0' 
		hspace		='0' 
		vspace		='0' 
		frameborder	='0'
	    scrolling   = "no" 
	    align       = "left">
</iFrame>

<script type="text/javascript">

   	with (document.forms[0])
	{
		target = "IFrmPagina"
		action = "ProcessoListarDesignacao.asp"
		submit()
	}

</script>

<%
Set objRSPag = Nothing
DesconectarCla()
%>
</html>
