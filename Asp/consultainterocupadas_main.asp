<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ConsultaInterOcupada.asp
'	- Responsável		: Vital
'	- Descrição			: Consulta Interligações Ocupadas
%>

<!--#include file="../inc/data.asp"-->
<html>
<head>
	<title>CLA - Controle Local de Acesso</title>
	<link rel=stylesheet type="text/css" href="../css/cla.css">
	<script language='javascript' src="../javascript/cla.js"></script>
</head>
<body leftmargin="0" topmargin="10" class=TA>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td align="center">
	<form action="consultainterocupadas_main.asp" method="post" onSubmit="return false">
	<input type="hidden" name="hdnRecId" value="<%=request("rec_id")%>">
	<table rules="groups" border=0 cellspacing="1" cellpadding="0" width="100%">
	<tr>
		<th colspan=2><p align=center>Consulta de Posições Ocupadas</p></th>
	</tr>
	<tr class=clsSilver>
		<%if Request("hdnRede") = "3" then %>
			<td width=130><font class="clsObrig">::</font>PADE/PAC</td>
		<%Else%>
			<td width=130><font class="clsObrig">::</font>Coordenada</td>
		<%End if%>	
		<td><input type="text" class="text" name="txtCoordenada" size="15" maxlength="20"></td>
	</tr>
</table>
</td>
</tr>
<tr>
<td align="center">
<input type="button" class="button" name="btnProcurar" value="Procurar" onClick="Procurar()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');">
&nbsp;&nbsp;
<input type="button" class="button" name="btnSair" value="Sair" onClick="javascript:window.close()" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
</td>
</tr>
</table>
<table width="100%">
<tr>
<td>
<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
</td>
</tr>
</table>
<span id=spnPosicoes></span>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnSistema value="<%=Request("hdnRede")%>">
<input type=hidden name=hdnNomeCons value="ConsInter">
</form>
<SCRIPT LANGUAGE="JavaScript">
function Procurar()
{
	with (document.forms[0])
	{
		<%if Request("hdnRede") = "3" then %>
			if (!ValidarCampos(txtCoordenada,"PADE/PAC")) return
		<%Else%>
			if (!ValidarCampos(txtCoordenada,"Coordenada")) return
		<%End if%>	

		target = "IFrmProcesso"
		hdnAcao.value = "ConsultarCoordenadasOcupadasAlocacao"
		action = "ProcessoConsFac.asp"
		submit()
	}	
		
}
</script>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "0" 
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
</body>
</html>