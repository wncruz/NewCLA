
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->

<script language="Javascript">

function Procurar()
{
	with (document.forms[0])
	{	 		 
		if (document.getElementById("txtbusca").value==""){
			alert("Informar CNL Sigla.")
			document.getElementById("txtbusca").focus()
			return
		}
		
		hdnAcaoEST.value = "consultar"
		action = "Assoc_estacaoCtf.asp"
		submit()
	}
}	
</script>

<form name=Form1 action="Assoc_estacaoCtf.asp" method="post" >
<input type=hidden name="hdnAcaoEST" id="hdnAcaoEST">

<table>
<tr>
	<td width=100%>
		<table border="0" cellspacing="1" cellpadding=0 width=760 >
			<tr>
				<th colspan=2><p align="center">Consulta Estação/Associação Estação CF</p></th>
			</tr>
			<tr class=clsSilver>
				<td >Busca (CNL Sigla. Ex.: SPO ou SPO IG)&nbsp;&nbsp;</td>
				<td >
				<input type="text" name="txtBusca" id="txtBusca" maxlength=8 class="text">&nbsp;
				<input type="button" class="button" name=btnProcurar value="Procurar" onclick="Procurar()"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
		
		<% if objDicCef.Exists("E") OR  objDicCef.Exists("C")   then '- Perfil GICL e GICN %>
			<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('manter_estacao.asp?TpProc=INC')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
		<%end if%>

		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
</table>
</form>
</body>
</html>
