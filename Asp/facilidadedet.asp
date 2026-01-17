<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: FacilidadeDet.ASP
'	- Descrição			: Detalha a solicitação
%>
<!--#include file="../inc/data.asp"-->
<HTML>
<HEAD>

<% if instr(1,Request.Form("hdnPaginaOrig"),"ConsRedeDet.asp") = 0 and Request.QueryString("RedeDet") <> 1  then  %>
	<!--#include file="../inc/header.asp"-->
<% else%>
	<link rel=stylesheet type="text/css" href="../css/cla.css">		
	<script language='JavaScript' src='../javascript/formatamenu.js'></script>
	<script language='JavaScript' src='../javascript/montamenu.js'></script>
	<script language='javascript' src="../javascript/cla.js"></script>
	<script language='javascript' src="../javascript/claMsg.js"></script>
	<title>CLA - Controle Local de Acesso</title>
<% end if %>


<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY topmargin=0 leftmargin=0>
<!--#include file="ConsultaProcesso.asp"-->
<SCRIPT LANGUAGE=javascript>
<!--
function VoltarOrigem()
{
	with (document.forms[0])
	{
		target = self.name 
		action = "<%=Request.Form("hdnPaginaOrig")%>"
		submit()
	}
}
window.name = "Facilidadedet.asp"
//-->
</SCRIPT>

<Form name=Form1 method=Post action="facilidade.asp">
	<input type=hidden name=hdnSolId	value="<%=dblSolId%>">
	<input type=hidden name=hdnPedId	value="<%=dblPedId%>">
	<input type=hidden name=dblRecId	value="<%=Request.Form("dblRecId")%>">
	<input type=hidden name=id			value="<%=Request("id")%>">
	<input type=hidden name=cboUsuario	value="<%=Request.Form("cboUsuario")%>">
	<input type=hidden name=cboStatus	value="<%=Request.Form("cboStatus")%>">
	<input type=hidden name=cboProvedor value="<%=Request.Form("cboProvedor")%>">
	<input type=hidden name=provedor	value="<%=Request.Form("provedor")%>">
	<input type=hidden name=cboEstacao	value="<%=Request.Form("cboEstacao")%>">
	<input type=hidden name=acao		value="<%=Request.Form("acao")%>">
	<input type=hidden name=datainicio	value="<%=Request.Form("datainicio")%>">
	<input type=hidden name=datafim		value="<%=Request.Form("datafim")%>">
	<input type=hidden name=hdnXmlReturn value="<%=Request.Form("hdnXmlReturn")%>">
	<!--Consulta de Rede Det-->
	<input type=hidden name=cboLocalInstala value="<%=Request.Form("cboLocalInstala")%>">
	<input type=hidden name=cboDominioNO	value="<%=Request.Form("cboDominioNO")%>">
	<table cellspacing=1 cellpadding=1  width=760 border=0>
		<tr>
			<td align=center>
				<% if instr(1,Request.Form("hdnPaginaOrig"),"ConsRedeDet.asp") = 0 and Request.QueryString("RedeDet") <> 1  then  %>
					<input type=button	class="button" name=btnVoltar value=Voltar onclick="VoltarOrigem()" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">&nbsp;
					<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
				<% else %>
					<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.close()" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
				<% end if %>
			</td>
		</tr>		
	</table>
</form>								 
<P>&nbsp;</P>

</BODY>
</HTML>
