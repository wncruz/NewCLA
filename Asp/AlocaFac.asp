<!--#include file="../inc/data.asp"-->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<Title>CLA - Controle Local de Acesso</Title>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<SCRIPT LANGUAGE="JavaScript">
var intTimeSlot = window.dialogArguments
function checa(f) {
	if (!ValidarDM(f.id)) return false
	return true;
}

function ValidarPedido()
{
	if (checa(document.forms[0]))
	{
		with (document.forms[0])
		{
			if (id.value.length != 13)
			{
				alert("Pedido de Acesso é obrigatório.");
				id.focus();
				return
			}
			if (!ValidarCampos(txttimeslot,"Time-slot")) return

			hdnAcao.value = "AlocarFacConsRedeDet"
			target = "IFrmProcesso"
			action = "ProcessoFac.asp"
			submit()
		}	
	}
}

function AlocarFacildade()
{
	with (document.forms[0])
	{
		window.returnValue = hdnSolId.value + "," + hdnPedId.value + "," + txttimeslot.value;
		window.close();
	}	
}
</script>
</HEAD>
<BODY topmargin=0 leftmargin=0 >
<Form name=Form1 Method=Post>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnSolId>
<input type=hidden name=hdnPedId>

<table border=0 cellspacing="1" cellpadding="0" width="100%">  
	<tr class=clsSilver>
		<th colspan=2><p align=center>Alocar Facilidade</p></th>
	</tr>
	<tr class=clsSilver>
		<td>
			<font class="clsObrig">:: </font>Pedido de Acesso
		</td>
		<td>
			<input type="text" class="text" name="id" value="DM-" maxlength="13" size="15">
		</td>
	</tr>
	<tr class=clsSilver>
		<td>
			<font class="clsObrig">:: </font>Time-slot
		</td>
	    <td>
			<input class=text onkeyup=ValidarNTipo(this,0,4,4,1,0,4) maxlength=9 size=10 name="txttimeslot" >(N4-N4)
		</td>
	</tr>
	<tr>
		<td align="center" colspan=2><br>
			<input type="button" class="button" name="btnAlocarFac" value="Alocar Facilidade" onClick="ValidarPedido()">&nbsp;
			<input type="button" class="button" name="btnSair" value="Sair" onClick="javascript:window.returnValue = '';window.close()">
		</td>
	</tr>	
	<tr>
		<td colspan=2>
			<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
		</td>
	</tr>
</table>
</Form>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso"
	    width       = "0"
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
</BODY>
<SCRIPT LANGUAGE=javascript>
<!--
document.forms[0].txttimeslot.value = intTimeSlot
//-->
</SCRIPT>
</HTML>
