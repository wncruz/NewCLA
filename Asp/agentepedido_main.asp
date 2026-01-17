<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: AgentePedido_main.asp
'	- Responsável		: Vital
'	- Descrição			: Adicionar/Remover Agente para acompanhamento do pedido
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<SCRIPT LANGUAGE=javascript>
<!--
function ProcurarAgente()
{
	with (document.forms[0])
	{
		if (ValidarPedido())
		{
			hdnAcao.value = "ListarAgentes"
			target = "IFrmProcesso"
			action = "ProcessoAgentepedido.asp"
			submit()
		}
	}
}
function IncluirAgente()
{
	with (document.forms[0])
	{
		if (ValidarPedido())
		{
			if (hdnSolId.value == "")
			{
				alert("Solicitação não encontrada.");
				return
			}
			hdnAcao.value = "IncluirAgente"
			target = "IFrmProcesso"
			action = "ProcessoAgentepedido.asp"
			submit()
		}
		else
		{return}
	}
}

function ExcluirAgente()
{
	with (document.forms[0])
	{
		hdnAcao.value = "RemoverAgente"
		target = "IFrmProcesso"
		action = "ProcessoAgentepedido.asp"
		submit()
	}
}

function ValidarPedido()
{
	with(document.forms[0])
	{
		if (txtPedido.value.length > 3) 
		{
			if (txtPedido.value.length != 13) {
				alert("O campo Pedido de Acesso não foi preenchido corretamente!");
				txtPedido.focus();
		    	return false;
			}
			if (txtPedido.value.substr(2,1) != "-") {
				alert("O campo Pedido de Acesso não foi preenchido corretamente!");
				txtPedido.focus();
		    	return false;
			}
				if (txtPedido.value.substr(8,1) != "/") {
				alert("O campo Pedido de Acesso não foi preenchido corretamente!");
				txtPedido.focus();
		    	return false;
			}
		}else
		{
			if(txtNroAcesso.value == "" && txtSolId.value=="") {
				alert("Preencha um campo para efetuar a consulta.");
				txtSolId.focus();
   				return false;
			}
		}
		return true;
	}
}

-->
</SCRIPT>
<form method="post" name=form1>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnSolId>
<input type=hidden name=hdnPedido>
<input type=hidden name=hdnNroAcesso>

<%
Dim itemCtfc
for Each itemCtfc in objDicCef
	Response.Write "<input type=hidden name=hdnCtfcId value='" &  objDicCef(itemCtfc) & "' >"
	Exit For
Next	
%>
<tr>
	<td >
		<table border=0 cellspacing="1" cellpadding="0" width="760">	
			<tr>
				<th colspan=2><p align="center">Cadastro de Agentes do Pedido</p></td>
			</tr>
			<tr class=clsSilver>
				<td width=150>&nbsp;Solicitação de Acesso</td>
				<td>
					<input type="text" class="text" name="txtSolId" maxlength="13" size="15" onKeyUp="ValidarTipo(this,0)" >
				</td>
			</tr>
			<tr class=clsSilver>
				<td >&nbsp;Pedido de Acesso
			</td>
			<td>
				<input type="text" class="text" name="txtPedido" value="DM-" maxlength="25" size="20"></td>
			</tr>
			<tr class=clsSilver>
				<td >&nbsp;Número de Acesso</td>
				<td><input type="text" class="text" name="txtNroAcesso" maxlength="25" size="25"></td>
			</tr>
		</table>
	</td>
<tr>
	<td >
		<span id="spnAgentes"></span>
	</td>
</tr>
</table>
<table cellspacing="1" cellpadding="0" width="760">
<tr>
	<td align=center>
		<span id="spnBtnInc"></span>
		<input type="button" class="button" name="btnProcurar" value="Procurar" onclick="ProcurarAgente()">
		<input type="button" class="button" name="sair" value="   Sair   " onClick="javascript:window.location.replace('main.asp')">
	</td>	
</tr>
<tr>
	<td>
		<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
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
</form>
</td>
</tr>
</body>
</html>
<%DesconectarCla()%>