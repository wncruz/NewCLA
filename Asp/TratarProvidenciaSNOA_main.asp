
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->

<html>

	<body>

		<SCRIPT LANGUAGE=javascript>
		<!--
		function Procurar()
		{
			with (document.forms[0])
			{
				if (checa(document.forms[0])) 
				{
					target = "IFrmProcesso"
					action = "ProcessoHisTratarProvSNOA.asp"
					submit()
				}	
			}
		}
		//-->
		</SCRIPT>

		<form method="post" id=form1 name=form1>

			<table border=0 cellspacing="1" cellpadding="0" width="760">

				<tr>
					<th colspan=2><p align=center>Tratar Providencia SNOA</p></th>
				</tr>

				<tr class=clsSilver>
					<td width="200"><font class="clsObrig">:: </font>Pedido de Acesso</td>
					<td>
						<input type="text" class="text" name="id" value="DM-" maxlength="25" size="25">
					</td>
				</tr>
				
				<tr class=clsSilver>
					<td width="200">
						<font class="clsObrig">:: </font>Número de Acesso:
					</td>
					<td>
						<input type="text" class="text" name="numero" maxlength="25" size="25">
					</td>
				</tr>

				<tr class=clsSilver>
					<td width="200">
						<font class="clsObrig">:: </font>Id físico:
					</td>
					<td>
						<input type="text" class="text" name="idFisico" maxlength="25" size = "25">
					</td>
				</tr>

				<tr class=clsSilver>
					<td width="200">
						<font class="clsObrig">:: </font>Solicitação:
					</td>
					<td>
						<input type="text" class="text" name="idSolicitacao" maxlength="8" size="25">
					</td>
				</tr>

				<tr class=clsSilver>
					<td width="200">
						<font class="clsObrig">:: </font>Número SNOA:
					</td>
					<td>
						<input type="text" class="text" name="numeroSNOA" maxlength="25" size="25">
					</td>
				</tr>

			</table>

			<table border=0 cellspacing="1" cellpadding="0" width="760">
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
		    		height      = "375px"
		    		frameborder = "0"
		    		scrolling   = "auto" 
		    		align       = "left">
			</iFrame>

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
		if ((f.numero.value == "" )&&(f.idFisico.value == "" )&&(f.idSolicitacao.value == "" )&&(f.numeroSNOA.value == "" ))
		{
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
