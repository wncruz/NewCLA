<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<html>
	<head>
		<script language='javascript' src="../javascript/xmlFacObjects.js"></script>
		<SCRIPT LANGUAGE=javascript>
		<!--
		var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")

		function HabilitaProcurar(obj)
		{
			if(document.getElementById('btnProcurar').disabled == true)
			{
				if(obj.value == "")
				{
					document.getElementById('btnProcurar').disabled = true;
				}
				else
				{
					document.getElementById('btnProcurar').disabled = false;
				}
			}
			else if (Vazio() == true)
			{
				document.getElementById('btnProcurar').disabled = true;
			}
		}

		function Vazio()
		{
			if(document.getElementById('txtAcessoLogico').value  == "")
			{
				return true;
			}
			
			return false;
		}


		//Executa a busca de senha
		function Procurar(){
			
			with (document.forms[0])
			{
				document.getElementById("IFrmLista").height = 370
				document.getElementById("BarraAzul").style.height = 0
				document.getElementById("BarraAzul").style.visibility = "hidden";
				//if (!ValidarDM(txtPedido)) return;
				target = "IFrmLista"
				action = "ListarDesignacao.asp"
				submit()
			}
		}

		//-->
		</SCRIPT>
	</head>
	<body>
		<form method="post" name=Form1 >
			<input type=hidden name="hdnAcao">
			<table border=0 cellspacing="1" cellpadding = 0 width="760" >
				<tr><th colspan=2 align=center><p align=center>Consulta Expressa</p></th></tr>
					<tr align=center class=clsSilver>
					<td><b>&nbsp;<b></td>
				</tr>
				<tr><td><br>
				<table align=center border=1 cellspacing="4" cellpadding="2" >
					<TR>
						<TD width="200" bgcolor="#31659c"><a href="ExpressaEnderecoCliente.asp"><font color=#ffffff size=2><b>Endereço + Cliente<b></FONT></A></TD>
					</TR>
					<TR>
						<TD width="200" bgcolor="#31659c"><a href="ExpressaOrderEntry.asp"><font color=#ffffff size=2><b>Order Entry<b></FONT></A></TD>
					</TR>
					<TR>
						<TD width="200" bgcolor="#31659c"><a href="ConsultaClaAprovisionador.asp"><font color=#ffffff size=2><b>Order Entry (Interface)<b></FONT></A></TD>
					</TR>					
				</table>
				</td></tr></table>
			<br>
			<table border="0" cellspacing="1" cellpadding="0" width="760">
				<tr>
					<td>
						<iframe	id			= "IFrmLista"
								name        = "IFrmLista" 
								width       = "100%"
								height      = "0"
								frameborder = "0"
								border		= "0"
								scrolling   = "no">
						</iFrame>
					</td>
				</tr>
			</table>
			<div id=divXls style="display:none;POSITION:relative">
				<table border=0 width=760><tr><td colspan=2 align=right></table>
			</div>
			<input type=hidden name=hdnCheck> 
			<input type=hidden name=hdnNomeCons value="ConsultaOSProvedor">
		</form>
		<iframe	id			= "IFrmLista"
			    name        = "IFrmLista" 
			    width       = "800"
			    height      = "0"
			    frameborder = "0"
			    scrolling   = "no" 
			    align       = "left">
		</iFrame>
	</body>
</html>
