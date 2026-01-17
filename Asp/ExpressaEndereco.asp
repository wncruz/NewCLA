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
			if(document.getElementById('txtNome').value  == "" && document.getElementById('txtNro').value  == "" && document.getElementById('txtBairro').value  == "" && document.getElementById('txtCep').value  == "" && document.getElementById('txtEstado').value  == "")
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
				action = "ExpressaListarEndereco.asp"
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
				<tr class=clsSilver>
					<th colspan=2><p align=center>Consulta Expressa - Endereço</p></th>
				</tr>
				<tr>
					<td class=clsSilver colspan=2>&nbsp;</td>
				</tr>
				<tr class=clsSilver>
					<td>&nbsp;&nbsp;&nbsp;Nome do Logradouro:</td>
					<td>
						<input type="text" class="text" name="txtNome" onclick="document.Form1.btnProcurar.disabled=false;" onKeyUp="HabilitaProcurar(this)" maxlength="60" size="65">
						&nbsp;&nbsp;&nbsp;Nº
						<input type="text" class="text" name="txtNro" onclick="document.Form1.btnProcurar.disabled=false;" onKeyUp="HabilitaProcurar(this)" maxlength="10" size="10">
					</td>
				</tr>
				<tr class=clsSilver>
					<td>&nbsp;&nbsp;&nbsp;Bairro:</td>
					<td>
						<input type="text" class="text" name="txtBairro" onclick="document.Form1.btnProcurar.disabled=false;" onKeyUp="HabilitaProcurar(this)" maxlength="30" size="32">
						&nbsp;&nbsp;&nbsp;CEP						
						<input type="text" class="text" name="txtCep" onclick="document.Form1.btnProcurar.disabled=false;" onKeyUp="HabilitaProcurar(this)" maxlength="9" size="10" onKeyPress="OnlyNumbers();AdicionaBarraCep(this)" maxlength="9" size="10">
					</td>
				</tr>
				<tr class=clsSilver>
					<td>&nbsp;&nbsp;&nbsp;Estado:</td>
					<td>
						<input type="text" class="text" name="txtEstado" onclick="document.Form1.btnProcurar.disabled=false;" onKeyUp="ValidarTipo(this,1);HabilitaProcurar(this)" maxlength="2" size="4">
						&nbsp;&nbsp;&nbsp;Quantidade:
						<input type="text" class="text" name="txtQTD" onclick="document.Form1.btnProcurar.disabled=false;" onKeyUp="ValidarTipo(this,0);HabilitaProcurar(this)" maxlength="2" size="4" value="10">					
						&nbsp;&nbsp;&nbsp;
						<input type="button" class="button" name="btnProcurar" value="Procurar" disabled="disabled" accesskey="P" onMouseOver="showtip(this,event,'Procurar (Alt+P)');" style="width:100px" onClick="Procurar()">
					</td>
				</tr>
				<tr class=clsSilver>
					
					<th id="BarraAzul" height=20 colspan=2></th>
				</tr>
			</table>
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
