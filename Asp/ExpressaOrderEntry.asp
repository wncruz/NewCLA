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
			if(document.getElementById('txt_oe_numero').value  == "" && document.getElementById('txt_oe_ano').value  == "" && document.getElementById('txt_oe_item').value  == "")
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
				action = "ExpressaListarOrderEntry.asp"
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
					<th colspan=2><p align=center>Consulta Expressa - OrderEntry</p></th>
				</tr>
				<tr>
					<td class=clsSilver colspan=2>&nbsp;</td>
				</tr>
				
				<tr class=clsSilver>
					<td>&nbsp;&nbsp;&nbsp;OE</td>
					<td><select name="cboSistema" title="Sistema">
							<option value=""></option>
							<option value="ADFAC">ADFAC</option>
							<option value="APG">APG</option>
							<option value="ASMS">ASMS</option>
							<option value="CFD">CFD</option>
							<option value="CFM">CFM</option>
							<option value="CFT">CFT</option>
							<option value="SGA DADOS">SGA DADOS</option>
							<option value="SGA PLUS">SGA PLUS</option>
							<option value="SGA VOZ VIP''S">SGA VOZ VIP'S</option>
						</select>
						<input id="txt_oe_numero" type="text" title="Número" maxlength="7" size="8" class=text onKeyUp="ValidarTipo(this,0);HabilitaProcurar(this);" name="txt_oe_numero" value=''>&nbsp;/
        				<input id="txt_oe_ano" type="text" title="Ano" maxlength="4" size="5" class=text onKeyUp="ValidarTipo(this,0);HabilitaProcurar(this);" name="txt_oe_ano" value=''>&nbsp;item&nbsp;
        				<input id="txt_oe_item" type="text" title="Item" maxlength="3" size="4" class=text onKeyUp="ValidarTipo(this,0);HabilitaProcurar(this);" name="txt_oe_item" value=''>
					</td>
				</tr>
				<tr class=clsSilver>
					<td>&nbsp;&nbsp;&nbsp;Quantidade:</td>
					<td>
						<input type="text" class="text" name="txtQTD" onKeyUp="ValidarTipo(this,0);" maxlength="2" size="4" value="10">
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
