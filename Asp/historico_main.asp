<!--#include file="../inc/data.asp"-->

<!--#include file="../inc/header.asp"-->
<SCRIPT LANGUAGE=javascript>
<!--
function Procurar()
{
	with (document.forms[0])
	{
		if (checa(document.forms[0])) 
		{
			target = "IFrmProcesso"
			action = "ProcessoHisPedido.asp"
			submit()
		}	
	}
}
//-->
</SCRIPT>

<tr><td width=100%>
<form method="post" id=form1 name=form1>
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
<td>
<tr>
	<th colspan=2><p align=center>Histórico de Pedido</p></th>
</tr>
<tr class=clsSilver>
	<td width=200px >
		<font class="clsObrig">:: </font>Pedido de Acesso
	</td>
	<td>
		<input type="text" class="text" name="id" value="DM-" maxlength="25" size="25">
	</td>
</tr>
<tr class=clsSilver>
	<td width="130">
		<font class="clsObrig">:: </font>Número de Acesso:
	</td>
	<td>
		<input type="text" class="text" name="numero" maxlength="25" size="25">
	</td>
</tr>
<tr class=clsSilver>
	<td width="130">
		<font class="clsObrig">:: </font>Id físico:
	</td>
	<td>
		<input type="text" class="text" name="idFisico" maxlength="25" size = "25">
	</td>
</tr>
<tr class=clsSilver>
	<td width="130">
		<font class="clsObrig">:: </font>Solicitação:
	</td>
	<td>
		<input type="text" class="text" name="idSolicitacao" maxlength="8" size="25">
	</td>
</tr>
</table>
</td>
</tr>
<tr>
	<td align="center">
		<input type="button" class="button" name="ok" value="Procurar" onClick="Procurar()">&nbsp;
		<input type="button" class="button" name="sair" value="Sair" onClick="javascript:window.location.replace('main.asp')">
</td>
</tr>
</table>
<table width="760">
<tr>
<td>
<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
</td>
</tr>
</table>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso"
	    width       = "100%"
	    height      = "275px"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
</td>
</tr>
</table>
</center>
</form>
<SCRIPT LANGUAGE="JavaScript">
function checa(f) {
	if (f.id.value.length > 3) 
	{
		if (f.id.value.length != 13) {
			alert("O campo Pedido de Acesso não foi preenchido corretamente!");
			f.id.focus();
	    	return false;
		}
		if (f.id.value.substr(2,1) != "-") {
			alert("O campo Pedido de Acesso não foi preenchido corretamente!");
			f.id.focus();
	    	return false;
		}
			if (f.id.value.substr(8,1) != "/") {
			alert("O campo Pedido de Acesso não foi preenchido corretamente!");
			f.id.focus();
	    	return false;
		}
	}
	else{
		if ((f.numero.value == "" )&&(f.idFisico.value == "" )&&(f.idSolicitacao.value == "" )){
			alert("Preencha um campo para efetuar a consulta.");
			f.id.focus();
   			return false;
		}
	}
	return true;
}
</script>

</body>
</html>
